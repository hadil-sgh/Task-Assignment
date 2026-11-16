"""Build the .pptx from a validated deck, then check the saved file independently.

Building: open the template, delete its sample slides, add one slide per deck slide from the
chosen layout, fill the placeholders by index with fitted text, remove the placeholders that
stay empty, and add the layout's footer, date and slide number.

Checking: re-open the saved file and verify, from what is really in it, that it has 1 to 4
slides, no empty placeholders and no text that overflows its box.
"""

from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Pt

from slidegen.fit import Fit, fit_text, fits
from slidegen.schema import MAX_SLIDES, Deck
from slidegen.template import Catalog

FOOTER_TYPES = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER}


def build_deck(
    deck: Deck,
    catalog: Catalog,
    template: Path,
    out: Path,
    footer: str | None = None,
    day: date | None = None,
) -> list[str]:
    """Write the deck to `out`. Returns warnings (text that had to be cut to fit)."""
    prs = Presentation(template)
    remove_sample_slides(prs)
    footer = footer if footer is not None else str(deck.slides[0].slots.get("title", ""))
    day_text = (day or date.today()).strftime("%d.%m.%Y")
    warnings = []
    for number, spec in enumerate(deck.slides):
        layout = catalog.layout(spec.layout)
        slide = prs.slides.add_slide(prs.slide_layouts[layout.index])
        slots = {slot.idx: slot for slot in layout.slots}
        for ph in list(slide.placeholders):
            slot = slots.get(ph.placeholder_format.idx)
            value = spec.slots.get(slot.name) if slot else None
            if not value:
                ph.element.getparent().remove(ph.element)  # nothing to show: no empty box
                continue
            fitted = fit_text(slot, value)
            write_text(ph, fitted)
            if fitted.cut:
                warnings.append(f"slides[{number}].{slot.name}: text cut to fit its box")
        add_footer(slide, footer, day_text)
    prs.save(out)
    return warnings


def write_text(ph, fitted: Fit) -> None:
    """One paragraph per bullet, with the fitted size written on every run."""
    frame = ph.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    for i, text in enumerate(fitted.paragraphs):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(fitted.size)


def remove_sample_slides(prs) -> None:
    """A template's own slides are examples, not part of the deck."""
    for slide_id in list(prs.slides._sldIdLst):
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)


def add_footer(slide, footer: str, day: str) -> None:
    """Put the layout's footer, date and slide number boxes on the slide.

    python-pptx leaves these three out when it creates a slide, so we copy them from the
    layout (at the layout's own position and style) and fill them. Layouts without them, such
    as a title slide, get none.
    """
    texts = {PP_PLACEHOLDER.FOOTER: footer, PP_PLACEHOLDER.DATE: day}
    for layout_ph in slide.slide_layout.placeholders:
        kind = layout_ph.placeholder_format.type
        if kind not in FOOTER_TYPES:
            continue
        slide.shapes.clone_placeholder(layout_ph)
        frame = slide.placeholders[layout_ph.placeholder_format.idx].text_frame
        if kind in texts:
            frame.text = texts[kind]
        else:  # a slide-number field: PowerPoint shows the current number
            old = frame.paragraphs[0]._p
            old.getparent().replace(old, parse_xml(
                f'<a:p {nsdecls("a")}><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" '
                'type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld></a:p>'
            ))  # fmt: skip


def check_deck_file(path: Path, catalog: Catalog) -> list[str]:
    """Re-open the saved deck and list its problems. An empty list means the deck is fine."""
    prs = Presentation(path)
    problems = []
    if not 1 <= len(prs.slides) <= MAX_SLIDES:
        problems.append(f"deck has {len(prs.slides)} slides, expected 1 to {MAX_SLIDES}")
    layouts = {layout.index: layout for layout in catalog.layouts}
    for number, slide in enumerate(prs.slides, start=1):
        layout = layouts.get(prs.slide_layouts.index(slide.slide_layout))
        if layout is None:
            problems.append(
                f"slide {number}: layout '{slide.slide_layout.name}' is not in the catalog"
            )
            continue
        slots = {slot.idx: slot for slot in layout.slots}
        for ph in slide.placeholders:
            slot = slots.get(ph.placeholder_format.idx)
            if ph.placeholder_format.type in FOOTER_TYPES or slot is None:
                continue
            paragraphs = [p.text for p in ph.text_frame.paragraphs]
            if not "".join(paragraphs).strip():
                problems.append(f"slide {number}: '{slot.name}' is empty")
            elif not fits(slot, paragraphs, _font_size(ph, slot)):
                problems.append(f"slide {number}: '{slot.name}' overflows its box")
    return problems


def _font_size(ph, slot) -> float:
    """The size written on the runs (the largest one), or the template size if none is set."""
    sizes = [r.font.size.pt for p in ph.text_frame.paragraphs for r in p.runs if r.font.size]
    return max(sizes, default=slot.font_pt)
