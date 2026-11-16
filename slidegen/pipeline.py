"""One run: notes in, a checked .pptx out, with at most two model calls.

1. pick the template and load its cached layout catalog;
2. one structured model call;
3. validate the answer; if it is invalid, one repair call with the error list (never a loop);
4. if the repair fails too, or the API cannot be reached: a plain deck built from the notes;
5. build the deck, then re-open the saved file and check it.

A missing or rejected API key is not covered by the fallback: only the user can fix it.
"""

import json
import math
from dataclasses import dataclass, field
from pathlib import Path

from slidegen.deck import build_deck, check_deck_file
from slidegen.llm import LLMError, MissingKeyError, RejectedKeyError, Usage
from slidegen.prompts import first_messages, repair_messages
from slidegen.schema import BULLET_MARKS, MAX_SLIDES, Deck, DeckError, parse_deck, wire_schema
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
            messages = repair_messages(notes, catalog, answer, error.errors)
            deck, warnings = parse_deck(call(client, messages, schema, result), catalog, notes)
    except (MissingKeyError, RejectedKeyError):
        raise
    except (LLMError, DeckError) as error:
        result.outcome = "fallback"
        result.warnings.append(f"plain fallback deck used: {str(error)[:200]}")
        return fallback_deck(notes, catalog)
    result.warnings += warnings
    return deck


def call(client, messages: list[dict], schema: dict, result: Result) -> str:
    text, usage = client.complete(messages, schema)
    result.calls.append(usage)
    return text


def fallback_deck(notes: str, catalog: Catalog) -> Deck:
    """A plain but valid deck without the model: each line of the notes becomes text.

    The lines are spread over at most 4 slides of the template's title-and-bullets layout;
    the first line of each slide is its title. Nothing is dropped: the fitter shrinks the font
    when there is a lot of text.
    """
    layout = plain_layout(catalog)
    body = next(slot for slot in layout.slots if slot.kind == "list")
    lines = [BULLET_MARKS.sub("", line).strip() for line in notes.splitlines()]
    lines = [line for line in lines if line]
    count = min(MAX_SLIDES, max(1, math.ceil(len(lines) / (body.lines + 1))))
    size = math.ceil(len(lines) / count)
    slides = []
    for start in range(0, len(lines), size):
        chunk = lines[start : start + size]
        slides.append({"layout": layout.id, "title": chunk[0], body.name: chunk[1:] or chunk})
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
