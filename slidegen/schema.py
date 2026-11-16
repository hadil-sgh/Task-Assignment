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
"""

import json
import math
import re

from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

from slidegen.template import Catalog, Slot

MAX_SLIDES = 4
HARD_LIMIT = 2  # hard limit = 2 x budget
BULLET_MARKS = re.compile(r"^\s*(?:[-*•▪■●►–]|\d+[.)])\s+")


class DeckError(Exception):
    """The model's answer is not a valid deck. Carries every problem found."""

    def __init__(self, errors: list[str]):
        super().__init__("; ".join(errors))
        self.errors = errors


class Slide(BaseModel):
    layout: str
    slots: dict[str, str | list[str]]  # slot name -> text or bullets, e.g. "title", "body2"

    @model_validator(mode="before")
    @classmethod
    def collect_slots(cls, data):
        """On the wire a slide is flat ({"layout", "title", "body", ...}); keep slots together."""
        if isinstance(data, dict) and "slots" not in data:
            slots = {key: value for key, value in data.items() if key != "layout"}
            return {"layout": data.get("layout"), "slots": slots}
        return data

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
    names = list(dict.fromkeys(slot.name for layout in catalog.layouts for slot in layout.slots))
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


def parse_deck(text: str, catalog: Catalog) -> tuple[Deck, list[str]]:
    """Turn the model's reply into a valid Deck. Returns (deck, warnings) or raises DeckError."""
    data = extract_json(text)
    try:
        deck = Deck.model_validate(data)
    except ValidationError as exc:
        messages = [_format_pydantic(error) for error in exc.errors()]
        raise DeckError(list(dict.fromkeys(messages))) from None  # drop duplicates, keep order
    errors, warnings = check_deck(deck, catalog)
    if errors:
        raise DeckError(errors)
    return deck, warnings


def check_deck(deck: Deck, catalog: Catalog) -> tuple[list[str], list[str]]:
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
        for slot in layout.slots:
            problem, warning = _check_slot(slot, slide.slots.get(slot.name))
            if problem:
                errors.append(f"{at}.{slot.name}: {problem}")
            if warning:
                warnings.append(f"{at}.{slot.name}: {warning}")
    return errors, warnings


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
    parts = [p for p in error["loc"] if p not in ("slots", "str", "list[str]")]
    path = "".join(f"[{part}]" if isinstance(part, int) else f".{part}" for part in parts)
    message = error["msg"]
    if error["type"] == "too_long" and error["loc"] == ("slides",):
        message += f"; merge related slides into at most {MAX_SLIDES}, without dropping content"
    return f"{path.lstrip('.') or 'response'}: {message}"
