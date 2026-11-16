"""Build the .pptx from a validated deck, then check the saved file independently.

Building: open the template, delete its sample slides, add one slide per deck slide from the
chosen layout, fill the placeholders by index with fitted text, remove the placeholders that
stay empty, and add the layout's footer, date and slide number.

Checking: re-open the saved file and verify, from what is really in it, that it has 1 to 4
slides, no empty placeholders and no text that overflows its box.
"""

from datetime import date
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

from slidegen.fit import Fit, fit_text, fits
from slidegen.schema import MAX_SLIDES, Chart, Deck, Table
from slidegen.template import Catalog

FOOTER_TYPES = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.SLIDE_NUMBER}
TABLE_MAX_PT, TABLE_MIN_PT = 16, 12


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
        table_box = next((slot for slot in layout.slots if slot.kind == "list"), None)
        for ph in list(slide.placeholders):
            slot = slots.get(ph.placeholder_format.idx)
            if slot and slot.kind == "chart" and spec.chart:
                add_chart(ph, spec.chart)
                continue
            if slot and slot is table_box and spec.table:
                add_table(slide, ph, spec.table, header_text_color(prs))
                continue
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


def add_chart(ph, chart: Chart) -> None:
    """A native, editable bar chart in the layout's chart box, in the theme's colours."""
    data = CategoryChartData()
    data.categories = chart.categories
    for series in chart.series:
        data.add_series(series.name, series.values)
    graph = ph.insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, data).chart
    graph.font.size = Pt(12)
    graph.has_title = bool(chart.title)
    if chart.title:
        graph.chart_title.text_frame.text = chart.title
        graph.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
    graph.has_legend = len(chart.series) > 1
    if graph.has_legend:
        graph.legend.position = XL_LEGEND_POSITION.BOTTOM
        graph.legend.include_in_layout = False
    plot = graph.plots[0]
    plot.gap_width = 60
    plot.vary_by_categories = False  # one colour per series, not per bar
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(11)
    graph.value_axis.minimum_scale = 0  # bars start at zero, so differences are not exaggerated


def add_table(slide, ph, table: Table, header_color: RGBColor) -> None:
    """A native table where the bullet box was: same position, header row in bold.

    The font goes from 16 pt down so that all rows fit the box, never below 12 pt.
    """
    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    ph.element.getparent().remove(ph.element)
    rows = [table.columns, *table.rows]
    size = max(TABLE_MIN_PT, min(TABLE_MAX_PT, int(Emu(height).pt / (len(rows) * 2.2))))
    row_height = Pt(size * 2.2)
    shape = slide.shapes.add_table(len(rows), len(table.columns), left, top, width,
                                   row_height * len(rows))  # fmt: skip
    for r, values in enumerate(rows):
        shape.table.rows[r].height = row_height
        for c, value in enumerate(values):
            cell = shape.table.cell(r, c)
            cell.text = value
            run = cell.text_frame.paragraphs[0].runs[0] if value else None
            if run:
                run.font.size = Pt(size)
                run.font.bold = r == 0
                if r == 0:
                    run.font.color.rgb = header_color


def header_text_color(prs) -> RGBColor:
    """The table header is filled with the theme's accent 1 (the team colour): dark text on a
    light colour, white text on a dark one."""
    theme = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    color = theme.find(f".//{qn('a:accent1')}")[0]
    value = color.get("lastClr") or color.get("val")
    red, green, blue = (int(value[i : i + 2], 16) for i in (0, 2, 4))
    light = 0.299 * red + 0.587 * green + 0.114 * blue > 150  # perceived brightness
    return RGBColor(0x1E, 0x1E, 0x1E) if light else RGBColor(0xFF, 0xFF, 0xFF)


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
            if ph.placeholder_format.type in FOOTER_TYPES or slot is None or not ph.has_text_frame:
                continue  # footer boxes, and chart boxes now holding a chart
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
