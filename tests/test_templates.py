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
    "Key Message",
    "Two Content",
    "Comparison",
    "Title Only",
    "Chart and Text",
    "Table",
]


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
def test_template_has_the_eight_layouts_and_no_slides(team: str) -> None:
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
        for name in ("Chart and Text", "Table")
    }
    assert PP_PLACEHOLDER.CHART in types["Chart and Text"]
    assert PP_PLACEHOLDER.TABLE in types["Table"]


def test_all_teams_share_the_same_layouts_and_placeholders() -> None:
    reference = layout_shape(load("base"))
    for team in TEAMS[1:]:
        assert layout_shape(load(team)) == reference


def test_each_team_has_its_own_accent_colour() -> None:
    accents = [accent1(load(team)) for team in TEAMS]
    assert len(set(accents)) == len(TEAMS)
