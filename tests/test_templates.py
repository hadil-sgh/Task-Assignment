"""The team templates are the company Master plus a team colour, bullets and decorations."""

from pathlib import Path

import pytest
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
TEAMS = ["marketing", "engineering", "finance"]
TITLE_SLIDE, END_SLIDE = 0, 34  # layout positions in the Master


def load(name: str) -> Presentation:
    return Presentation(TEMPLATE_DIR / f"{name}.pptx")


def layout_shape(prs: Presentation) -> list[tuple[str, list[int]]]:
    """Layout names and their placeholder indexes: what the generator relies on."""
    return [
        (layout.name, sorted(ph.placeholder_format.idx for ph in layout.placeholders))
        for layout in prs.slide_layouts
    ]


def accent(prs: Presentation, n: int) -> str:
    theme = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    return theme.find(f".//{qn(f'a:accent{n}')}/{qn('a:srgbClr')}").get("val")


def has_bullet(ph) -> bool:
    level = ph.element.txBody.find(f"{qn('a:lstStyle')}/{qn('a:lvl1pPr')}")
    return level is not None and level.find(qn("a:buChar")) is not None


@pytest.mark.parametrize("team", TEAMS)
def test_team_keeps_every_master_layout_and_placeholder(team: str) -> None:
    assert layout_shape(load(team)) == layout_shape(load("master"))


@pytest.mark.parametrize("team", TEAMS)
def test_team_has_no_sample_slides(team: str) -> None:
    assert len(load(team).slides) == 0


@pytest.mark.parametrize("team", TEAMS)
def test_first_and_last_slides_are_the_masters(team: str) -> None:
    master, copy = load("master"), load(team)
    for index in (TITLE_SLIDE, END_SLIDE):
        names = [s.name for s in master.slide_layouts[index].shapes]
        assert [s.name for s in copy.slide_layouts[index].shapes] == names


@pytest.mark.parametrize("team", TEAMS)
def test_mosaic_is_recoloured_in_the_team_colour(team: str) -> None:
    """The Master's mosaic picture gets a duotone effect; its other pictures stay untouched."""
    for index in (TITLE_SLIDE, END_SLIDE):
        layout = load(team).slide_layouts[index].element
        assert len(list(layout.iter(qn("a:duotone")))) == 1
    assert not list(load("master").slide_layouts[TITLE_SLIDE].element.iter(qn("a:duotone")))


@pytest.mark.parametrize("team", TEAMS)
def test_master_footer_date_and_logo_are_kept(team: str) -> None:
    names = {s.name for s in load("master").slide_master.shapes}
    assert names <= {s.name for s in load(team).slide_master.shapes}


@pytest.mark.parametrize("team", TEAMS)
def test_only_content_boxes_get_bullets(team: str) -> None:
    for layout in load(team).slide_layouts:
        for ph in layout.placeholders:
            kind = ph.placeholder_format.type
            if kind == PP_PLACEHOLDER.OBJECT:
                assert has_bullet(ph), (layout.name, ph.placeholder_format.idx)
            elif kind == PP_PLACEHOLDER.BODY:  # subtitles, subheadlines, source lines
                assert not has_bullet(ph), (layout.name, ph.placeholder_format.idx)


def test_each_team_has_its_own_main_colour() -> None:
    """Accent 1 colours the Master's titles, subheadlines, charts and tables."""
    colours = {accent(load(team), 1) for team in TEAMS}
    assert len(colours) == len(TEAMS)
    assert accent(load("master"), 1) not in colours


@pytest.mark.parametrize("team", TEAMS)
def test_first_and_last_slides_keep_the_masters_colours(team: str) -> None:
    """Their accent colours are pinned, so the team colours never reach them."""
    for index in (TITLE_SLIDE, END_SLIDE):
        layout = load(team).slide_layouts[index].element
        assert not [r for r in layout.iter(qn("a:schemeClr")) if r.get("val").startswith("accent")]
    end_slide = load(team).slide_layouts[END_SLIDE].element  # its address is the Master navy
    assert accent(load("master"), 1) in {c.get("val") for c in end_slide.iter(qn("a:srgbClr"))}
