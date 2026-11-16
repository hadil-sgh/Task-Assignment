"""Parsing and validating the model's answer against the base template catalog.

Base catalog, for reference:
  L0 Title Slide: title, subtitle          L3 Two Content: title, body, body2
  L1 Title and Content: title, body        L4 Comparison: title, head, body, head2, body2
  L2 Key Message: title, body (2 lines)    L5 Title Only: title
"""

import json

import pytest

from slidegen.schema import DeckError, check_deck, extract_json, parse_deck, wire_schema
from slidegen.template import Catalog

GOOD = {
    "slides": [
        {
            "layout": "L1",
            "title": "Phase 1 finished on time",
            "body": ["312 luminaires", "No downtime"],
        },
        {"layout": "L2", "title": "Approve phase 2 budget by 1 December", "body": []},
    ]
}


def answer(*slides: dict) -> str:
    return json.dumps({"slides": list(slides)})


def errors_of(text: str, catalog: Catalog) -> list[str]:
    with pytest.raises(DeckError) as caught:
        parse_deck(text, catalog)
    return caught.value.errors


# ---------------------------------------------------------------- extracting JSON


def test_plain_json_is_read() -> None:
    assert extract_json(json.dumps(GOOD)) == GOOD


def test_json_inside_code_fences_and_prose_is_read() -> None:
    reply = f"Sure! Here is the deck:\n```json\n{json.dumps(GOOD)}\n```\nHope this helps."
    assert extract_json(reply) == GOOD


def test_json_after_prose_without_fences_is_read() -> None:
    assert extract_json(f"Here you go: {json.dumps(GOOD)} Enjoy") == GOOD


def test_non_json_is_an_error(catalog: Catalog) -> None:
    assert errors_of("I cannot help with that.", catalog) == ["response: no JSON object found"]


# ---------------------------------------------------------------- valid answers


def test_valid_deck_passes(catalog: Catalog) -> None:
    deck, warnings = parse_deck(json.dumps(GOOD), catalog)
    assert [slide.layout for slide in deck.slides] == ["L1", "L2"]
    assert deck.slides[0].slots["title"] == "Phase 1 finished on time"
    assert warnings == []


def test_unused_slots_may_be_sent_empty(catalog: Catalog) -> None:
    slide = {"layout": "L5", "title": "Questions", "subtitle": "", "body": [], "head2": ""}
    deck, _ = parse_deck(answer(slide), catalog)
    assert deck.slides[0].layout == "L5"


def test_text_is_cleaned(catalog: Catalog) -> None:
    slide = {
        "layout": "L1",
        "title": "  Results  ",
        "body": ["- First", "• Second", "  ", "3) Third"],
    }
    deck, _ = parse_deck(answer(slide), catalog)
    assert deck.slides[0].slots["title"] == "Results"
    assert deck.slides[0].slots["body"] == ["First", "Second", "Third"]


def test_optional_slots_may_be_empty(catalog: Catalog) -> None:
    title_slide = {"layout": "L0", "title": "Ex-Lighting Upgrade", "subtitle": ""}
    key_message = {"layout": "L2", "title": "Approve the budget", "body": []}
    parse_deck(answer(title_slide, key_message), catalog)  # no error


# ---------------------------------------------------------------- invalid answers


def test_unknown_layout_lists_the_valid_ids(catalog: Catalog) -> None:
    errors = errors_of(answer({"layout": "L9", "title": "x"}), catalog)
    assert errors == ["slides[0].layout: unknown layout 'L9'; use one of L0, L1, L2, L3, L4, L5"]


def test_too_many_slides(catalog: Catalog) -> None:
    slides = [{"layout": "L5", "title": f"Slide {i}"} for i in range(6)]
    [error] = errors_of(answer(*slides), catalog)
    assert error.startswith("slides: List should have at most 4 items")
    assert "merge related slides" in error


def test_no_slides(catalog: Catalog) -> None:
    [error] = errors_of(answer(), catalog)
    assert error.startswith("slides: List should have at least 1 item")


def test_empty_required_text(catalog: Catalog) -> None:
    slide = {
        "layout": "L4",
        "title": "",
        "head": "Retrofit",
        "body": ["Cheap"],
        "head2": "",
        "body2": [],
    }
    assert errors_of(answer(slide), catalog) == [
        "slides[0].title: must not be empty",
        "slides[0].head2: must not be empty",
        "slides[0].body2: must not be empty",
    ]


def test_slot_the_layout_does_not_have(catalog: Catalog) -> None:
    slide = {"layout": "L1", "title": "Results", "body": ["One"], "body2": ["Stray column"]}
    assert errors_of(answer(slide), catalog) == [
        "slides[0].body2: layout L1 has no 'body2' slot; leave it empty"
    ]


def test_wrong_types(catalog: Catalog) -> None:
    errors = errors_of(answer({"layout": "L1", "title": 5, "body": ["One"]}), catalog)
    assert errors == ["slides[0].title: Input should be a valid string",
                      "slides[0].title: Input should be a valid list"]  # fmt: skip
    errors = errors_of(answer({"layout": "L1", "title": ["A", "B"], "body": ["One"]}), catalog)
    assert errors == ["slides[0].title: expected text"]


def test_all_errors_are_collected_at_once(catalog: Catalog) -> None:
    slides = [{"layout": "L9", "title": "x"}, {"layout": "L1", "title": "", "body": []}]
    assert len(errors_of(answer(*slides), catalog)) == 3


# ---------------------------------------------------------------- budget and hard limit


def test_text_over_budget_is_accepted_with_a_warning(catalog: Catalog) -> None:
    budget = catalog.layout("L1").slots[0].budget
    slide = {"layout": "L1", "title": "x" * (budget + 10), "body": ["One"]}
    _, warnings = parse_deck(answer(slide), catalog)
    assert warnings == [f"slides[0].title: over budget: {budget + 10} of {budget} characters; "
                        "font will be reduced"]  # fmt: skip


def test_text_over_the_hard_limit_is_an_error(catalog: Catalog) -> None:
    budget = catalog.layout("L1").slots[0].budget
    slide = {"layout": "L1", "title": "x" * (2 * budget + 1), "body": ["One"]}
    [error] = errors_of(answer(slide), catalog)
    assert error.startswith(
        f"slides[0].title: too long: {2 * budget + 1} characters, limit {budget}"
    )


def test_bullets_count_wrapped_lines(catalog: Catalog) -> None:
    body = catalog.layout("L1").slots[1]  # 10 lines of 83 characters
    two_lines_each = ["y" * (body.chars + 1)] * body.lines  # each bullet wraps to 2 lines
    _, warnings = parse_deck(
        answer({"layout": "L1", "title": "T", "body": two_lines_each}), catalog
    )
    assert "over budget: 20 of 10 lines" in warnings[0]
    too_many = ["short"] * (2 * body.lines + 1)
    [error] = errors_of(answer({"layout": "L1", "title": "T", "body": too_many}), catalog)
    assert "too long: 21 lines, limit 10" in error


def test_check_deck_returns_errors_and_warnings_separately(catalog: Catalog) -> None:
    from slidegen.schema import Deck

    deck = Deck.model_validate(GOOD)
    assert check_deck(deck, catalog) == ([], [])


# ---------------------------------------------------------------- wire schema


def test_wire_schema_is_strict_and_flat(catalog: Catalog) -> None:
    schema = wire_schema(catalog)
    assert schema["strict"] is True
    slide = schema["schema"]["properties"]["slides"]["items"]
    assert slide["properties"]["layout"]["enum"] == ["L0", "L1", "L2", "L3", "L4", "L5"]
    assert list(slide["properties"]) == [
        "layout",
        "title",
        "subtitle",
        "body",
        "body2",
        "head",
        "head2",
    ]
    assert slide["required"] == list(slide["properties"])
    assert slide["additionalProperties"] is False
    assert slide["properties"]["body"]["type"] == "array"
    assert slide["properties"]["head"]["type"] == "string"


def test_wire_schema_has_no_length_keywords(catalog: Catalog) -> None:
    """Limits live in descriptions; many providers reject minItems/maxLength in strict mode."""
    text = json.dumps(wire_schema(catalog))
    for keyword in ("minItems", "maxItems", "minLength", "maxLength"):
        assert keyword not in text
