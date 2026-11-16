"""Building the .pptx on the company Master, and the independent check of the saved file."""

from datetime import date
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Pt

from slidegen import fit
from slidegen.deck import FOOTER_TYPES, build_deck, check_deck_file
from slidegen.schema import Deck
from slidegen.template import MASTER, Catalog

DECK = {
    "slides": [
        {"layout": "L0", "title": "Ex-Lighting Upgrade 2026", "subtitle": ""},
        {
            "layout": "L1",
            "title": "Phase 1 finished on time and 4% under budget",
            "head": "Zone 1 rollout at Weil and Tunis",
            "body": ["312 luminaires replaced in Zone 1 areas", "No unplanned downtime"],
            "head2": "",
        },
        {"layout": "L15", "title": "Approve phase 2 by 1 December", "subtitle": ""},
    ]
}
DAY = date(2026, 11, 17)


@pytest.fixture(autouse=True)
def estimated_widths(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(fit, "load_font", lambda family, bold: None)


def build(tmp_path: Path, catalog: Catalog, data: dict = DECK) -> tuple[Path, list[str]]:
    out = tmp_path / "deck.pptx"
    warnings = build_deck(Deck.model_validate(data), catalog, MASTER, out, day=DAY)
    return out, warnings


def kinds(slide) -> set:
    return {ph.placeholder_format.type for ph in slide.placeholders}


# ---------------------------------------------------------------- building


def test_one_slide_per_deck_slide_and_no_sample_slides(tmp_path: Path, catalog: Catalog) -> None:
    out, warnings = build(tmp_path, catalog)
    prs = Presentation(out)
    names = [slide.slide_layout.name for slide in prs.slides]
    assert names == ["9_Title slide refinery", "Titel und Inhalt", "Separation Slide"]
    assert warnings == []


def test_text_lands_in_the_right_boxes(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    content = Presentation(out).slides[1]
    texts = {ph.placeholder_format.idx: ph.text_frame.text for ph in content.placeholders}
    assert texts[0] == "Phase 1 finished on time and 4% under budget"
    assert texts[13] == "Zone 1 rollout at Weil and Tunis"
    assert texts[18] == "312 luminaires replaced in Zone 1 areas\nNo unplanned downtime"


def test_empty_optional_boxes_are_removed(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    title_slide, content, _ = Presentation(out).slides
    assert [ph.placeholder_format.idx for ph in title_slide.placeholders] == [0]  # no subtitle
    assert 22 not in {ph.placeholder_format.idx for ph in content.placeholders}  # no source


def test_every_run_has_an_explicit_font_size(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    for slide in Presentation(out).slides:
        for ph in slide.placeholders:
            if ph.placeholder_format.type in FOOTER_TYPES:  # filled by the layout's style
                continue
            runs = [r for p in ph.text_frame.paragraphs for r in p.runs]
            assert runs and all(r.font.size is not None for r in runs)


def test_content_slides_get_footer_date_and_slide_number(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    title_slide, content, separation = Presentation(out).slides
    footer = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER}
    assert footer <= kinds(content) and footer <= kinds(separation)
    assert not footer & kinds(title_slide)  # the Master's title slide has no footer boxes
    texts = {ph.placeholder_format.type: ph.text_frame.text for ph in content.placeholders}
    assert texts[PP_PLACEHOLDER.DATE] == "17.11.2026"
    assert texts[PP_PLACEHOLDER.FOOTER] == "Ex-Lighting Upgrade 2026"  # the first title


def test_text_that_cannot_fit_is_cut_and_reported(tmp_path: Path, catalog: Catalog) -> None:
    data = {"slides": [{"layout": "L1", "title": "Long", "body": ["word " * 60] * 20}]}
    out, warnings = build(tmp_path, catalog, data)
    assert warnings == ["slides[0].body: text cut to fit its box"]
    assert check_deck_file(out, catalog) == []  # what was written still fits


# ---------------------------------------------------------------- the independent check


def test_a_good_deck_passes_the_check(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    assert check_deck_file(out, catalog) == []


def test_check_finds_overflowing_text(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    prs = Presentation(out)
    body = next(ph for ph in prs.slides[1].placeholders if ph.placeholder_format.idx == 18)
    body.text_frame.text = "much longer text than the box can hold " * 30
    body.text_frame.paragraphs[0].runs[0].font.size = Pt(40)
    prs.save(out)
    assert check_deck_file(out, catalog) == ["slide 2: 'body' overflows its box"]


def test_check_finds_empty_boxes(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    prs = Presentation(out)
    next(
        ph for ph in prs.slides[1].placeholders if ph.placeholder_format.idx == 13
    ).text_frame.text = ""
    prs.save(out)
    assert check_deck_file(out, catalog) == ["slide 2: 'head' is empty"]


def test_check_finds_too_many_slides(tmp_path: Path, catalog: Catalog) -> None:
    out, _ = build(tmp_path, catalog)
    prs = Presentation(out)
    for _ in range(2):
        slide = prs.slides.add_slide(prs.slide_layouts[15])
        slide.shapes.title.text = "Extra"
        for ph in list(slide.placeholders)[1:]:
            ph.element.getparent().remove(ph.element)
    prs.save(out)
    assert check_deck_file(out, catalog) == ["deck has 5 slides, expected 1 to 4"]
