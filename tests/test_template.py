"""Layout catalog, its cache, and the content-based team choice."""

import shutil
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

from slidegen import template
from slidegen.template import (
    TEMPLATE_DIR,
    Layout,
    Slot,
    build_catalog,
    dedupe,
    load_catalog,
    pick_team,
    resolve_template,
)

TEAMS = ["base", "marketing", "engineering", "finance"]


@pytest.fixture
def base_copy(tmp_path: Path) -> Path:
    """A private copy of the base template, so cache files land in tmp_path."""
    return Path(shutil.copy(TEMPLATE_DIR / "base.pptx", tmp_path / "base.pptx"))


# ---------------------------------------------------------------- catalog


def test_catalog_keeps_only_text_layouts(base_copy: Path) -> None:
    names = [layout.name for layout in build_catalog(base_copy).layouts]
    assert names == [
        "Title Slide",
        "Title and Content",
        "Key Message",
        "Two Content",
        "Comparison",
        "Title Only",
    ]
    # Chart and Table layouts are skipped: the model only writes text.


def test_layout_ids_point_at_the_template_layout_index(base_copy: Path) -> None:
    prs = Presentation(base_copy)
    for layout in build_catalog(base_copy).layouts:
        assert prs.slide_layouts[layout.index].name == layout.name


def test_comparison_slots_are_named_left_to_right(base_copy: Path) -> None:
    comparison = next(lay for lay in build_catalog(base_copy).layouts if lay.name == "Comparison")
    assert [slot.name for slot in comparison.slots] == ["title", "head", "body", "head2", "body2"]
    assert [slot.kind for slot in comparison.slots] == ["text", "text", "list", "text", "list"]


def test_budgets_follow_box_size(base_copy: Path) -> None:
    layouts = {layout.name: layout for layout in build_catalog(base_copy).layouts}
    full = layouts["Title and Content"].slots[1]
    half = layouts["Two Content"].slots[1]
    assert full.font_pt == 20
    assert 60 <= full.chars <= 120 and 6 <= full.lines <= 12
    assert half.chars < full.chars / 2 + 5  # a half-width column holds about half as much


def test_prompt_is_short(base_copy: Path) -> None:
    prompt = build_catalog(base_copy).prompt()
    assert prompt.splitlines()[4].startswith("L4 Comparison: title<=")
    assert len(prompt) < 600  # roughly 150 tokens at most


def test_all_teams_expose_the_same_layouts_and_slots() -> None:
    def shape(team: str) -> list:
        catalog = build_catalog(TEMPLATE_DIR / f"{team}.pptx")
        return [(lay.id, lay.name, [s.name for s in lay.slots]) for lay in catalog.layouts]

    assert all(shape(team) == shape("base") for team in TEAMS)


def test_works_with_a_foreign_template(tmp_path: Path) -> None:
    """Any template works: here python-pptx's own default, which we never designed for."""
    path = tmp_path / "default.pptx"
    Presentation().save(path)
    names = [layout.name for layout in build_catalog(path).layouts]
    assert "Comparison" in names
    assert "Picture with Caption" not in names  # has a picture placeholder
    assert "Blank" not in names  # has no text placeholder
    assert "Title and Vertical Text" not in names  # vertical text


def set_prompt(path: Path, layout_index: int, idx: int, text: str) -> None:
    """Write prompt text into a layout placeholder, like a template designer would."""
    prs = Presentation(path)
    layout = prs.slide_layouts[layout_index]
    next(p for p in layout.placeholders if p.placeholder_format.idx == idx).text_frame.text = text
    prs.save(path)


def test_optional_prompt_text_makes_an_optional_slot_with_a_hint(base_copy: Path) -> None:
    set_prompt(base_copy, 4, 3, "Optional: Sources and/or captions")  # Comparison, head2
    head2 = next(s for s in build_catalog(base_copy).layout("L4").slots if s.name == "head2")
    assert head2.optional
    assert head2.describe().startswith("head2?<=")
    assert head2.describe().endswith(' "Optional: Sources and/or captions"')


def test_boilerplate_prompt_text_is_not_a_hint(base_copy: Path) -> None:
    set_prompt(base_copy, 4, 3, "Click to edit Master text styles")
    head2 = next(s for s in build_catalog(base_copy).layout("L4").slots if s.name == "head2")
    assert head2.hint == "" and not head2.optional


def test_one_line_box_without_bullets_is_an_optional_subtitle(base_copy: Path) -> None:
    prs = Presentation(base_copy)
    body = next(p for p in prs.slide_layouts[2].placeholders if p.placeholder_format.idx == 1)
    body.height = Inches(0.45)  # Key Message: room for one line only
    prs.save(base_copy)
    names = [(s.name, s.optional) for s in build_catalog(base_copy).layout("L2").slots]
    assert names == [("title", False), ("subtitle", True)]


def test_identical_layouts_are_listed_once() -> None:
    """A template with white and blue copies of a layout shows the model only the first."""
    title = Slot("title", 0, "text", 1, 40, 30.0)
    layouts = [
        Layout("L1", "Title", [title]),
        Layout("L7", "Blue_Title", [title]),
        Layout("L8", "Big title", [Slot("title", 0, "text", 2, 40, 30.0)]),
    ]
    assert [layout.id for layout in dedupe(layouts)] == ["L1", "L8"]


def test_narrow_font_fits_more_characters(base_copy: Path, tmp_path: Path) -> None:
    """Our templates use Arial Narrow; the same box holds about 20% more than with Arial."""
    narrow = build_catalog(base_copy).layout("L1").slots[1].chars
    prs = Presentation(base_copy)
    for latin in prs.slide_master.element.iter(qn("a:latin")):  # "+mn-lt" -> explicit Arial
        latin.set("typeface", "Arial")
    regular_path = tmp_path / "regular.pptx"
    prs.save(regular_path)
    regular = build_catalog(regular_path).layout("L1").slots[1].chars
    assert narrow > regular * 1.15


# ---------------------------------------------------------------- cache


def test_cache_is_written_and_reused(base_copy: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    first = load_catalog(base_copy)
    assert base_copy.with_suffix(".catalog.json").exists()

    def fail(*args, **kwargs):
        raise AssertionError("catalog was rebuilt instead of read from the cache")

    monkeypatch.setattr(template, "build_catalog", fail)
    assert load_catalog(base_copy) == first


def test_cache_rebuilds_when_the_template_changes(base_copy: Path) -> None:
    first = load_catalog(base_copy)
    prs = Presentation(base_copy)
    prs.slide_layouts.remove(prs.slide_layouts[5])  # drop "Title Only"
    prs.save(base_copy)
    second = load_catalog(base_copy)
    assert second.key != first.key
    assert "Title Only" not in [layout.name for layout in second.layouts]


# ---------------------------------------------------------------- team choice


def test_figures_heavy_notes_go_to_finance() -> None:
    notes = "Q3 revenue EUR 4.2m, margin 18%, costs up 6%, forecast EUR 5m for Q4"
    assert pick_team(notes) == "finance"


def test_technical_notes_go_to_engineering() -> None:
    notes = "Prototype passed ATEX test. Firmware release next week, sensor module redesign."
    assert pick_team(notes) == "engineering"


def test_campaign_notes_go_to_marketing() -> None:
    notes = "Launch campaign for the new brand, trade fair in March, LinkedIn leads up"
    assert pick_team(notes) == "marketing"


def test_french_notes_are_recognised() -> None:
    assert pick_team("Lancement de la campagne au salon, nouvelle marque") == "marketing"


def test_neutral_or_mixed_notes_use_base() -> None:
    assert pick_team("Team offsite on Friday, agenda and lunch") == "base"
    assert pick_team("Campaign launch; prototype test") == "base"  # tie stays neutral


FINANCE_NOTES = "revenue margin costs 12%"


def test_company_template_is_the_default_when_installed(
    base_copy: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(template, "COMPANY_TEMPLATE", base_copy)
    assert resolve_template(FINANCE_NOTES) == base_copy
    assert resolve_template(FINANCE_NOTES, team="auto").stem == "finance"
    assert resolve_template(FINANCE_NOTES, team="marketing").stem == "marketing"


def test_without_company_template_the_team_is_picked_from_the_notes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(template, "COMPANY_TEMPLATE", tmp_path / "missing.pptx")
    assert resolve_template(FINANCE_NOTES).stem == "finance"


def test_explicit_template_file_wins() -> None:
    custom = Path("company.pptx")
    assert resolve_template(FINANCE_NOTES, team="marketing", template=custom) == custom


def test_unknown_team_lists_the_choices() -> None:
    with pytest.raises(ValueError, match="base, engineering, finance, marketing"):
        resolve_template("notes", team="sales")
