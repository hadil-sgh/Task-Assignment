"""One run: notes in, a checked .pptx out, with at most two model calls.

1. pick the template and load its cached layout catalog;
2. one structured model call;
3. validate the answer; if it is invalid, one repair call with the error list (never a loop);
4. if only a table or chart is still wrong, show it as bullets and keep the model's deck;
   if anything else is still wrong, or the API cannot be reached: a plain deck from the notes;
5. build the deck, then re-open the saved file and check it.

A missing or rejected API key is not covered by the fallback: only the user can fix it.
"""

import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

from pydantic import ValidationError

from slidegen.deck import build_deck, check_deck_file
from slidegen.llm import LLMError, MissingKeyError, RejectedKeyError, Usage
from slidegen.prompts import first_messages, repair_messages
from slidegen.schema import (
    BULLET_MARKS,
    MAX_SLIDES,
    Chart,
    Deck,
    DeckError,
    Slide,
    Table,
    check_deck,
    extract_json,
    parse_deck,
    wire_schema,
)
from slidegen.template import Catalog, Layout, load_catalog, resolve_template


class EmptyNotesError(ValueError):
    pass


@dataclass
class Result:
    path: Path
    template: Path
    outcome: str = "first try"  # or "repaired" or "fallback"
    calls: list[Usage] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)  # from the check of the saved file

    @property
    def tokens(self) -> int:
        return sum(c.prompt_tokens + c.completion_tokens for c in self.calls)

    @property
    def cost_usd(self) -> float:
        return sum(c.cost_usd for c in self.calls)

    @property
    def latency_s(self) -> float:
        return sum(c.latency_s for c in self.calls)


def generate(notes: str, client, out: Path, team: str | None = None,
             template: Path | None = None) -> Result:  # fmt: skip
    """Turn rough notes into a checked deck at `out`. `client` is a model client or a mock."""
    if not notes.strip():
        raise EmptyNotesError("The notes are empty.")
    template_path = resolve_template(notes, team, template)
    catalog = load_catalog(template_path)
    result = Result(out, template_path)
    deck = ask_model(notes, catalog, client, result)
    out.parent.mkdir(parents=True, exist_ok=True)
    result.warnings += build_deck(deck, catalog, template_path, out)
    result.problems = check_deck_file(out, catalog)
    return result


def ask_model(notes: str, catalog: Catalog, client, result: Result) -> Deck:
    """First call, at most one repair call, then the plain fallback deck."""
    schema = wire_schema(catalog)
    try:
        answer = call(client, first_messages(notes, catalog), schema, result)
        try:
            deck, warnings = parse_deck(answer, catalog, notes)
        except DeckError as error:
            result.outcome = "repaired"
            result.warnings += [f"first answer rejected: {problem}" for problem in error.errors]
            messages = repair_messages(notes, catalog, answer, error.errors)
            answer = call(client, messages, schema, result)
            try:
                deck, warnings = parse_deck(answer, catalog, notes)
            except DeckError as still:
                deck, warnings = visuals_as_bullets(answer, catalog, notes, still)
    except (MissingKeyError, RejectedKeyError):
        raise
    except (LLMError, DeckError) as error:
        result.outcome = "fallback"
        result.warnings.append(f"plain fallback deck used: {str(error)[:200]}")
        return fallback_deck(notes, catalog)
    result.warnings += warnings
    return deck


VISUAL_ERROR = re.compile(r"^slides\[(\d+)\]\.(table|chart)")


def visuals_as_bullets(
    answer: str, catalog: Catalog, notes: str, error: DeckError
) -> tuple[Deck, list[str]]:
    """Keep the model's deck when only a table or chart is still wrong after the repair.

    Those slides move to the plain title-and-bullets layout, and the table rows or chart
    values become bullets, so no content is lost. Invented numbers are never kept: if any
    error is about a number missing from the notes, the error is raised again.
    """
    matches = [VISUAL_ERROR.match(message) for message in error.errors]
    if not all(matches) or any("does not appear in the notes" in m for m in error.errors):
        raise error
    try:
        deck = Deck.model_validate(extract_json(answer))
    except ValidationError:
        raise error from None
    simplified = sorted({int(match.group(1)) for match in matches})
    plain = plain_layout(catalog)
    body = next(slot.name for slot in plain.slots if slot.kind == "list")
    texts = [slot.name for slot in plain.slots if slot.kind == "text"]  # title, subheadline...
    for i in simplified:
        slide = deck.slides[i]
        layout = catalog.layout(slide.layout)
        bullets = [
            b for s in layout.slots if s.kind == "list" for b in slide.slots.get(s.name) or []
        ]
        bullets += _table_bullets(slide.table) + _chart_bullets(slide.chart)
        kept = {name: slide.slots[name] for name in texts if slide.slots.get(name)}
        deck.slides[i] = Slide.model_validate({"layout": plain.id, **kept, body: bullets})
    errors, warnings = check_deck(deck, catalog, notes)
    if errors:
        raise DeckError(errors)
    notes_on_change = [f"slides[{i}]: table or chart shown as bullets" for i in simplified]
    return deck, warnings + notes_on_change


def _table_bullets(table: Table | None) -> list[str]:
    """One bullet per row: "Next.js: Load (s) 1.1, Weeks 18"."""
    if table is None:
        return []
    return [
        f"{row[0]}: "
        + ", ".join(
            f"{name} {cell}" for name, cell in zip(table.columns[1:], row[1:], strict=False)
        )
        for row in table.rows
        if row
    ]


def _chart_bullets(chart: Chart | None) -> list[str]:
    """One bullet per category: "Next.js: Lighthouse 96"."""
    if chart is None:
        return []
    return [
        f"{category}: " + ", ".join(f"{s.name} {s.values[i]:g}" for s in chart.series
                                     if i < len(s.values))
        for i, category in enumerate(chart.categories)
    ]  # fmt: skip


def call(client, messages: list[dict], schema: dict, result: Result) -> str:
    text, usage = client.complete(messages, schema)
    result.calls.append(usage)
    return text


def fallback_deck(notes: str, catalog: Catalog) -> Deck:
    """A plain but valid deck without the model: each line of the notes becomes text.

    The lines are spread over at most 4 slides of the template's title-and-bullets layout;
    the first sentence of each slide's first line is its title, the rest becomes a bullet.
    Nothing is dropped: the fitter shrinks the font when there is a lot of text.
    """
    layout = plain_layout(catalog)
    body = next(slot for slot in layout.slots if slot.kind == "list")
    lines = [BULLET_MARKS.sub("", line).strip() for line in notes.splitlines()]
    lines = [line for line in lines if line]
    count = min(MAX_SLIDES, max(1, math.ceil(len(lines) / (body.lines + 1))))
    size = math.ceil(len(lines) / count)
    slides = []
    for start in range(0, len(lines), size):
        title, *bullets = lines[start : start + size]
        first_sentence = re.match(r"(.{8,}?[?.!])\s+(.+)", title)  # "Website relaunch: which...?"
        if first_sentence:
            title, rest = first_sentence.groups()
            bullets.insert(0, rest)
        slides.append({"layout": layout.id, "title": title, body.name: bullets or [title]})
    return Deck.model_validate({"slides": slides})


def plain_layout(catalog: Catalog) -> Layout:
    """The layout whose only required slots are a title and one bullet box."""

    def required(layout: Layout) -> list[str]:
        return [slot.kind for slot in layout.slots if not slot.optional]

    for layout in catalog.layouts:
        if required(layout) == ["text", "list"]:
            return layout
    return next(layout for layout in catalog.layouts if "list" in required(layout))


def to_json(deck: Deck) -> str:
    """The deck in the flat wire format the model uses (handy for mocks and logs)."""
    return json.dumps({"slides": [{"layout": s.layout, **s.slots} for s in deck.slides]})
