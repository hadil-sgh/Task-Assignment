"""The deck the model must return: the JSON schema we send, and the checks we run on the answer.

Two layers, on purpose:
- The wire schema sent to the provider is flat and simple (strict mode, every field required,
  no length keywords), because providers support only a subset of JSON Schema.
- After the call, Pydantic checks types and slide count, and check_deck() checks the answer
  against the layout catalog. All errors are collected at once, as "path: message" lines,
  so a single repair call can fix them all.

Every text has two limits: the budget (what fits at the template font size, which the model is
asked for) and a hard limit at twice the budget. Between the two, the text is accepted and the
fitter shrinks the font later; above the hard limit it is an error and triggers the repair call.

A slide may also carry a native table (in place of its first bullet box) or a bar chart (in the
layout's chart box). Every number in them must appear in the notes: the model never invents
or calculates figures.
"""

import json
import math
import re

from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

from slidegen.template import Catalog, Slot

MAX_SLIDES = 4
HARD_LIMIT = 2  # hard limit = 2 x budget
BULLET_MARKS = re.compile(r"^\s*(?:[-*•▪■●►–]|\d+[.)])\s+")
NUMBER = re.compile(r"\d+(?:[.,]\d+)?")
MAX_ROWS, MAX_COLUMNS, MAX_CELL = 6, 5, 30  # table limits (rows without the header)
MAX_CATEGORIES, MAX_SERIES = 8, 3  # chart limits


class DeckError(Exception):
    """The model's answer is not a valid deck. Carries every problem found."""

    def __init__(self, errors: list[str]):
        super().__init__("; ".join(errors))
        self.errors = errors


class Series(BaseModel):
    name: str
    values: list[float]


class Chart(BaseModel):
    title: str = ""
    categories: list[str] = []
    series: list[Series] = []


class Table(BaseModel):
    columns: list[str] = []
    rows: list[list[str]] = []


class Slide(BaseModel):
    layout: str
    slots: dict[str, str | list[str]]  # slot name -> text or bullets, e.g. "title", "body2"
    chart: Chart | None = None
    table: Table | None = None

    @model_validator(mode="before")
    @classmethod
    def collect_slots(cls, data):
        """On the wire a slide is flat ({"layout", "title", "body", ...}); keep slots together."""
        if isinstance(data, dict) and "slots" not in data:
            visuals = {"chart": data.get("chart"), "table": data.get("table")}
            slots = {k: v for k, v in data.items() if k not in ("layout", "chart", "table")}
            return {"layout": data.get("layout"), "slots": slots, **visuals}
        return data

    @field_validator("chart", "table")
    @classmethod
    def drop_empty(cls, value):
        """An empty chart or table on the wire means "not used"."""
        if isinstance(value, Chart) and not (value.categories or value.series):
            return None
        if isinstance(value, Table) and not (value.columns or value.rows):
            return None
        return value

    @field_validator("slots")
    @classmethod
    def clean(cls, slots: dict) -> dict:
        """Trim spaces, drop empty bullets, and remove bullet marks the model may have typed."""
        return {
            name: value.strip()
            if isinstance(value, str)
            else [BULLET_MARKS.sub("", item).strip() for item in value if item.strip()]
            for name, value in slots.items()
        }


class Deck(BaseModel):
    slides: list[Slide] = Field(min_length=1, max_length=MAX_SLIDES)


# ---------------------------------------------------------------- the schema we send


def wire_schema(catalog: Catalog) -> dict:
    """Strict JSON schema for the response_format: one flat object per slide."""
    slots = [slot for layout in catalog.layouts for slot in layout.slots]
    names = list(dict.fromkeys(slot.name for slot in slots if slot.kind != "chart"))
    props: dict[str, dict] = {
        "layout": {
            "type": "string",
            "enum": [layout.id for layout in catalog.layouts],
            "description": "Layout id from the catalog.",
        }
    }
    for name in names:
        if name.startswith("body"):
            props[name] = {
                "type": "array",
                "items": {"type": "string"},
                "description": f"Bullets for the '{name}' slot: short fragments. 'NxM' in "
                "the catalog means at most N bullets of M characters. [] if the layout has no "
                "such slot.",
            }
        else:
            props[name] = {
                "type": "string",
                "description": f"Text for the '{name}' slot, at most the characters the catalog "
                "allows for this layout. Empty string if the layout has no such slot.",
            }
    if any(slot.kind == "list" for slot in slots):
        props["table"] = TABLE_SCHEMA
    if any(slot.kind == "chart" for slot in slots):
        props["chart"] = CHART_SCHEMA
    slide = {
        "type": "object",
        "properties": props,
        "required": list(props),
        "additionalProperties": False,
    }
    deck = {
        "type": "object",
        "properties": {
            "slides": {
                "type": "array",
                "description": f"1 to {MAX_SLIDES} slides, as few as the content needs.",
                "items": slide,
            }
        },
        "required": ["slides"],
        "additionalProperties": False,
    }
    return {"name": "deck", "strict": True, "schema": deck}


def _object(description: str, **properties) -> dict:
    return {"type": "object", "description": description, "properties": properties,
            "required": list(properties), "additionalProperties": False}  # fmt: skip


STRINGS = {"type": "array", "items": {"type": "string"}}
TABLE_SCHEMA = _object(
    "A table to compare several options on several figures. It replaces the layout's first "
    f"bullet box ('body'), which must then be []. At most {MAX_ROWS} rows and {MAX_COLUMNS} "
    "columns. Only numbers that appear in the notes. Empty ([] and []) when not used.",
    columns=STRINGS,
    rows={"type": "array", "items": STRINGS},
)
CHART_SCHEMA = _object(
    "A bar chart for one figure compared across items, or for percentages. Only on a layout "
    f"with a chart box. At most {MAX_CATEGORIES} categories and {MAX_SERIES} series. Only "
    "numbers that appear in the notes. Empty (no categories, no series) when not used.",
    title={"type": "string"},
    categories=STRINGS,
    series={"type": "array", "items": _object(
        "One row of numbers, one per category.", name={"type": "string"},
        values={"type": "array", "items": {"type": "number"}})},
)  # fmt: skip


# ---------------------------------------------------------------- parsing and checking the answer


def extract_json(text: str) -> dict:
    """Find one JSON object in the reply, even inside code fences or surrounding prose."""
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    if fenced:
        text = fenced.group(1)
    decoder = json.JSONDecoder()
    for start in (i for i, char in enumerate(text) if char == "{"):
        try:
            value, _ = decoder.raw_decode(text, start)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise DeckError(["response: no JSON object found"])


def parse_deck(text: str, catalog: Catalog, notes: str | None = None) -> tuple[Deck, list[str]]:
    """Turn the model's reply into a valid Deck. Returns (deck, warnings) or raises DeckError.

    With the notes, every number in a table or chart is checked against them.
    """
    data = extract_json(text)
    try:
        deck = Deck.model_validate(data)
    except ValidationError as exc:
        messages = [_format_pydantic(error) for error in exc.errors()]
        raise DeckError(list(dict.fromkeys(messages))) from None  # drop duplicates, keep order
    errors, warnings = check_deck(deck, catalog, notes)
    if errors:
        raise DeckError(errors)
    return deck, warnings


def check_deck(
    deck: Deck, catalog: Catalog, notes: str | None = None
) -> tuple[list[str], list[str]]:
    """Check every slide against its layout. Returns (errors, warnings)."""
    errors: list[str] = []
    warnings: list[str] = []
    ids = [layout.id for layout in catalog.layouts]
    for i, slide in enumerate(deck.slides):
        at = f"slides[{i}]"
        if slide.layout not in ids:
            errors.append(
                f"{at}.layout: unknown layout {slide.layout!r}; use one of {', '.join(ids)}"
            )
            continue
        layout = catalog.layout(slide.layout)
        allowed = {slot.name for slot in layout.slots}
        for name, value in slide.slots.items():
            if name not in allowed and value:
                errors.append(
                    f"{at}.{name}: layout {layout.id} has no '{name}' slot; leave it empty"
                )
        table_box = next((slot for slot in layout.slots if slot.kind == "list"), None)
        for slot in layout.slots:
            value = slide.slots.get(slot.name)
            if slot.kind == "chart":
                continue  # checked with the chart below
            if slide.table and slot is table_box:
                if value:
                    errors.append(f"{at}.{slot.name}: leave it empty, the table takes its place")
                continue
            problem, warning = _check_slot(slot, value)
            if problem:
                errors.append(f"{at}.{slot.name}: {problem}")
            if warning:
                warnings.append(f"{at}.{slot.name}: {warning}")
        errors += [f"{at}.{error}" for error in _check_visuals(slide, layout, notes)]
    return errors, warnings


def _check_visuals(slide: Slide, layout, notes: str | None) -> list[str]:
    """Chart and table: the right layout, the limits, and only numbers from the notes."""
    errors = []
    has_chart_box = any(slot.kind == "chart" for slot in layout.slots)
    if has_chart_box and slide.chart is None:
        errors.append("chart: this layout has a chart box; give chart data or pick another layout")
    if slide.chart and not has_chart_box:
        errors.append("chart: this layout has no chart box; pick a chart layout or leave it empty")
    if chart := slide.chart:
        if not 1 <= len(chart.categories) <= MAX_CATEGORIES:
            errors.append(f"chart: use 1 to {MAX_CATEGORIES} categories")
        if not 1 <= len(chart.series) <= MAX_SERIES:
            errors.append(f"chart: use 1 to {MAX_SERIES} series")
        if any(len(series.values) != len(chart.categories) for series in chart.series):
            errors.append("chart: every series needs one value per category")
    if table := slide.table:
        if not any(slot.kind == "list" for slot in layout.slots):
            errors.append("table: this layout has no bullet box for a table")
        if not 2 <= len(table.columns) <= MAX_COLUMNS or not 1 <= len(table.rows) <= MAX_ROWS:
            errors.append(f"table: use 2 to {MAX_COLUMNS} columns and 1 to {MAX_ROWS} rows")
        if any(len(row) != len(table.columns) for row in table.rows):
            errors.append("table: every row needs one cell per column")
        if any(len(cell) > MAX_CELL for row in table.rows for cell in [*row, *table.columns]):
            errors.append(f"table: keep every cell to {MAX_CELL} characters")
    if notes is not None:
        errors += _invented_numbers(slide, notes)
    return errors


def _invented_numbers(slide: Slide, notes: str) -> list[str]:
    """Numbers in the chart or the table rows that do not appear anywhere in the notes."""
    known = {_number(text) for text in NUMBER.findall(notes)}
    used = {"chart": [v for s in slide.chart.series for v in s.values] if slide.chart else []}
    rows = slide.table.rows if slide.table else []
    used["table"] = [_number(text) for row in rows for cell in row for text in NUMBER.findall(cell)]
    return [
        f"{name}: {value:g} does not appear in the notes; use only figures from the notes, "
        "never calculated ones"
        for name, values in used.items()
        for value in dict.fromkeys(values)
        if value not in known
    ]


def _number(text: str) -> float:
    return float(text.replace(",", "."))


def is_required(slot: Slot) -> bool:
    """Every slot must be filled unless the catalog marks it optional ("?").

    Optional are subtitles and boxes whose template prompt says "Optional ...", such as the
    Master's subheadline and source line.
    """
    return not slot.optional


def _check_slot(slot: Slot, value) -> tuple[str | None, str | None]:
    """Return (error, warning) for one slot."""
    expected = list if slot.kind == "list" else str
    if value is None:
        value = expected()
    if not isinstance(value, expected):
        return f"expected {'a list of bullets' if expected is list else 'text'}", None
    if not value:
        return ("must not be empty" if is_required(slot) else None), None
    if slot.kind == "list":
        used = sum(max(1, math.ceil(len(item) / slot.chars)) for item in value)
        budget, unit = slot.lines, "lines"
    else:
        used, budget, unit = len(value), slot.budget, "characters"
    if used > HARD_LIMIT * budget:
        return (
            f"too long: {used} {unit}, limit {budget}; shorten, merge or move to another slide",
            None,
        )
    if used > budget:
        return None, f"over budget: {used} of {budget} {unit}; font will be reduced"
    return None, None


def _format_pydantic(error: dict) -> str:
    """Pydantic error -> "slides[1].title: message" (hiding our internal 'slots' grouping)."""
    hidden = ("slots", "str", "list[str]")
    parts = [p for p in error["loc"] if p not in hidden]
    path = "".join(f"[{part}]" if isinstance(part, int) else f".{part}" for part in parts)
    message = error["msg"]
    if error["type"] == "too_long" and error["loc"] == ("slides",):
        message += f"; merge related slides into at most {MAX_SLIDES}, without dropping content"
    return f"{path.lstrip('.') or 'response'}: {message}"
