"""The team templates must be interchangeable: same layouts, same placeholders, different look."""

from pathlib import Path

import pytest
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
TEAMS = ["base", "marketing", "engineering", "finance"]
EXPECTED_LAYOUTS = [
    "Title Slide",
    "Title and Content",
    "Separation Slide",
    "Title and 2 Columns",
    "Title and 3 Columns",
    "Title Only",
    "Title, Text and Chart",
    "Title and Table",
    "End Slide",
]

# Text box positions of the company Master (x, y, width, height in inches), by layout and index.
MASTER_POSITIONS = {
    ("Title and Content", 0): (0.72, 0.78, 11.89, 0.47),  # title
    ("Title and Content", 13): (0.72, 1.27, 11.89, 0.29),  # subheadline
    ("Title and Content", 1): (0.72, 1.98, 9.96, 4.62),  # body
    ("Title and Content", 22): (0.72, 6.75, 9.96, 0.15),  # source
    ("Title and 2 Columns", 2): (6.90, 1.98, 5.71, 4.62),
    ("Title and 3 Columns", 4): (8.96, 1.98, 3.65, 4.62),
    ("Title Slide", 0): (0.72, 5.86, 9.96, 0.47),
    ("Title Slide", 1): (0.72, 6.33, 9.96, 0.57),
    ("Separation Slide", 0): (0.72, 1.92, 9.94, 1.82),
    ("End Slide", 0): (0.72, 1.92, 9.94, 1.07),
}


def load(team: str) -> Presentation:
    return Presentation(TEMPLATE_DIR / f"{team}.pptx")


def layout_shape(prs: Presentation) -> dict[str, list[int]]:
    """Layout name -> sorted placeholder indexes: what the generator relies on."""
    return {
        layout.name: sorted(ph.placeholder_format.idx for ph in layout.placeholders)
        for layout in prs.slide_layouts
    }


def accent1(prs: Presentation) -> str:
    theme = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    return theme.find(f".//{qn('a:accent1')}/{qn('a:srgbClr')}").get("val")


@pytest.mark.parametrize("team", TEAMS)
def test_template_has_the_master_layouts_and_no_slides(team: str) -> None:
    prs = load(team)
    assert [layout.name for layout in prs.slide_layouts] == EXPECTED_LAYOUTS
    assert len(prs.slides) == 0


@pytest.mark.parametrize("team", TEAMS)
def test_template_is_16_by_9(team: str) -> None:
    prs = load(team)
    assert round(Emu(prs.slide_width).inches / Emu(prs.slide_height).inches, 2) == 1.78


@pytest.mark.parametrize("team", TEAMS)
def test_no_footer_placeholders_left(team: str) -> None:
    footer = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}
    for layout in load(team).slide_layouts:
        assert not [ph for ph in layout.placeholders if ph.placeholder_format.type in footer]


@pytest.mark.parametrize("team", TEAMS)
def test_chart_and_table_layouts_have_native_placeholders(team: str) -> None:
    layouts = {layout.name: layout for layout in load(team).slide_layouts}
    types = {
        name: {ph.placeholder_format.type for ph in layouts[name].placeholders}
        for name in ("Title, Text and Chart", "Title and Table")
    }
    assert PP_PLACEHOLDER.CHART in types["Title, Text and Chart"]
    assert PP_PLACEHOLDER.TABLE in types["Title and Table"]


@pytest.mark.parametrize("team", TEAMS)
def test_text_boxes_sit_where_the_master_puts_them(team: str) -> None:
    layouts = {layout.name: layout for layout in load(team).slide_layouts}
    for (name, idx), expected in MASTER_POSITIONS.items():
        ph = next(p for p in layouts[name].placeholders if p.placeholder_format.idx == idx)
        box = tuple(round(Emu(v).inches, 2) for v in (ph.left, ph.top, ph.width, ph.height))
        assert box == expected, (team, name, idx)


def test_all_teams_share_the_same_layouts_and_placeholders() -> None:
    reference = layout_shape(load("base"))
    for team in TEAMS[1:]:
        assert layout_shape(load(team)) == reference


def test_each_team_has_its_own_accent_colour() -> None:
    accents = [accent1(load(team)) for team in TEAMS]
    assert len(set(accents)) == len(TEAMS)
