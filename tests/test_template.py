"""Layout catalog, its cache, and the choice of template."""

import shutil
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn

from slidegen import template
from slidegen.template import (
    MASTER,
    TEMPLATE_DIR,
    Layout,
    Slot,
    build_catalog,
    dedupe,
    load_catalog,
    pick_team,
    resolve_template,
)

TEAMS = ["marketing", "engineering", "finance"]
CONTENT, COLUMNS_3, TITLE_ONLY, SEPARATION = 1, 3, 4, 15  # layout positions in the Master


@pytest.fixture
def master_copy(tmp_path: Path) -> Path:
    """A private copy of the Master, so cache files and edits stay in tmp_path."""
    return Path(shutil.copy(MASTER, tmp_path / "master.pptx"))


def slot(catalog, layout_name: str, slot_name: str) -> Slot:
    layout = next(lay for lay in catalog.layouts if lay.name == layout_name)
    return next(s for s in layout.slots if s.name == slot_name)


# ---------------------------------------------------------------- catalog


def test_catalog_keeps_distinct_text_layouts(master_copy: Path) -> None:
    """35 layouts in the Master: picture, multi-chart and duplicate blue layouts are left out."""
    names = [layout.name for layout in build_catalog(master_copy).layouts]
    assert names == [
        "9_Title slide refinery",
        "Titel und Inhalt",
        "Title and 2 Columns",
        "Title and 3 Columns",
        "Nur Titel",
        "Title and Text w/ Diagram 1",
        "Title and Text w/ Diagram 2",
        "Separation Slide",
        "1_End Slide",
    ]


def test_layout_ids_point_at_the_template_layout_index(master_copy: Path) -> None:
    prs = Presentation(master_copy)
    for layout in build_catalog(master_copy).layouts:
        assert prs.slide_layouts[layout.index].name == layout.name


def test_slots_are_named_by_role_and_position(master_copy: Path) -> None:
    """Title, the subtitle under it, then the rest left to right; one-line boxes are heads."""
    catalog = build_catalog(master_copy)
    columns = next(lay for lay in catalog.layouts if lay.name == "Title and 3 Columns")
    assert [s.name for s in columns.slots] == [
        "title",
        "subtitle",
        "body",
        "head",
        "body2",
        "body3",
    ]
    assert [s.kind for s in columns.slots] == ["text", "text", "list", "text", "list", "list"]


def test_subheadline_and_source_are_optional_with_hints(master_copy: Path) -> None:
    catalog = build_catalog(master_copy)
    subheadline = slot(catalog, "Titel und Inhalt", "subtitle")
    source = slot(catalog, "Titel und Inhalt", "head")
    assert subheadline.optional and subheadline.hint == "Optional Subheadline"
    assert source.optional and source.hint == "Optional: Sources and/or captions"
    assert not slot(catalog, "Titel und Inhalt", "body").optional


def test_title_slide_presenter_line_is_an_optional_subtitle(master_copy: Path) -> None:
    subtitle = slot(build_catalog(master_copy), "9_Title slide refinery", "subtitle")
    assert subtitle.optional and subtitle.hint == "Presenter, Place, Date"


def test_budgets_follow_box_size(master_copy: Path) -> None:
    catalog = build_catalog(master_copy)
    full = slot(catalog, "Titel und Inhalt", "body")
    half = slot(catalog, "Title and 2 Columns", "body")
    assert full.font_pt == 20
    assert 60 <= full.chars <= 120 and 6 <= full.lines <= 12
    # a column is 5.71 in wide, the full body 9.96 in: it holds proportionally fewer characters
    assert abs(half.chars / full.chars - 5.71 / 9.96) < 0.05


def test_prompt_is_short(master_copy: Path) -> None:
    prompt = build_catalog(master_copy).prompt()
    assert prompt.splitlines()[1].startswith("L1 Titel und Inhalt: title<=")
    assert len(prompt) < 1100  # roughly 275 tokens at most, hints included


def test_team_templates_expose_the_same_layouts_and_slots() -> None:
    def shape(path: Path) -> list:
        catalog = build_catalog(path)
        return [(lay.id, lay.name, [s.name for s in lay.slots]) for lay in catalog.layouts]

    assert all(shape(TEMPLATE_DIR / f"{team}.pptx") == shape(MASTER) for team in TEAMS)


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


def test_optional_prompt_text_makes_an_optional_slot_with_a_hint(master_copy: Path) -> None:
    set_prompt(master_copy, COLUMNS_3, 26, "Optional: a middle column")
    body2 = slot(build_catalog(master_copy), "Title and 3 Columns", "body2")
    assert body2.optional
    assert body2.describe().startswith("body2?<=")
    assert body2.describe().endswith(' "Optional: a middle column"')


def test_boilerplate_prompt_text_is_not_a_hint(master_copy: Path) -> None:
    set_prompt(master_copy, COLUMNS_3, 26, "Click to edit Master text styles")
    body2 = slot(build_catalog(master_copy), "Title and 3 Columns", "body2")
    assert body2.hint == "" and not body2.optional


def test_one_line_box_without_bullets_is_an_optional_subtitle(master_copy: Path) -> None:
    """Even without an "Optional" prompt, a one-line box next to a title is a subtitle."""
    set_prompt(master_copy, SEPARATION, 12, "")
    slots = build_catalog(master_copy).layout(f"L{SEPARATION}").slots
    assert [(s.name, s.optional) for s in slots] == [("title", False), ("subtitle", True)]


def test_identical_layouts_are_listed_once() -> None:
    """A template with white and blue copies of a layout shows the model only the first."""
    title = Slot("title", 0, "text", 1, 40, 30.0)
    layouts = [
        Layout("L1", "Title", [title]),
        Layout("L17", "Blue_Title", [title]),
        Layout("L18", "Big title", [Slot("title", 0, "text", 2, 40, 30.0)]),
    ]
    assert [layout.id for layout in dedupe(layouts)] == ["L1", "L18"]


def test_narrow_font_fits_more_characters(master_copy: Path, tmp_path: Path) -> None:
    """The Master uses Arial Narrow; the same box holds about 20% more than with Arial."""
    narrow = slot(build_catalog(master_copy), "Titel und Inhalt", "body").chars
    prs = Presentation(master_copy)
    for latin in prs.slide_master.element.iter(qn("a:latin")):  # "+mn-lt" -> explicit Arial
        latin.set("typeface", "Arial")
    regular_path = tmp_path / "regular.pptx"
    prs.save(regular_path)
    regular = slot(build_catalog(regular_path), "Titel und Inhalt", "body").chars
    assert narrow > regular * 1.15


# ---------------------------------------------------------------- cache


def test_cache_is_written_and_reused(master_copy: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    first = load_catalog(master_copy)
    assert master_copy.with_suffix(".catalog.json").exists()

    def fail(*args, **kwargs):
        raise AssertionError("catalog was rebuilt instead of read from the cache")

    monkeypatch.setattr(template, "build_catalog", fail)
    assert load_catalog(master_copy) == first


def test_cache_rebuilds_when_the_template_changes(master_copy: Path) -> None:
    first = load_catalog(master_copy)
    prs = Presentation(master_copy)
    prs.slide_layouts.remove(prs.slide_layouts[TITLE_ONLY])  # drop "Nur Titel"
    prs.save(master_copy)
    second = load_catalog(master_copy)
    assert second.key != first.key
    assert "Nur Titel" not in [layout.name for layout in second.layouts]


# ---------------------------------------------------------------- choosing the template


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


def test_the_master_is_the_default() -> None:
    assert resolve_template(FINANCE_NOTES) == MASTER
    assert resolve_template(FINANCE_NOTES, team="base") == MASTER


def test_teams_are_chosen_by_name_or_from_the_notes() -> None:
    assert resolve_template(FINANCE_NOTES, team="marketing").stem == "marketing"
    assert resolve_template(FINANCE_NOTES, team="auto").stem == "finance"
    assert resolve_template("Team offsite on Friday", team="auto") == MASTER


def test_explicit_template_file_wins() -> None:
    custom = Path("company.pptx")
    assert resolve_template(FINANCE_NOTES, team="marketing", template=custom) == custom


def test_unknown_team_lists_the_choices() -> None:
    with pytest.raises(ValueError, match="auto, base, engineering, finance, marketing"):
        resolve_template("notes", team="sales")
