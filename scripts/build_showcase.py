"""Fill every template with demo slides, including a native chart and table.

Only used to look at the templates (render with scripts/render.py); not part of the tool.

Run:  uv run python scripts/build_showcase.py
Writes out/showcase/<template>.pptx
"""

from pathlib import Path

from build_templates import remove_sample_slides
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT / "slidegen" / "templates"
SHOWCASE_DIR = ROOT / "out" / "showcase"

CHART, TABLE = "<chart>", "<table>"
FOOTER, DATE = "Ex-Lighting Upgrade 2026 | Plant management update", "17.11.2026"

# (layout index in the Master, {placeholder index: text}). In the Master's content layouts,
# 0 is the title, 13 the subheadline, 18/26/27 the body columns and 22 the source line.
SLIDES = [
    (0, {0: "Ex-Lighting Upgrade 2026", 10: "Plant management update\nTunis, 17 November 2026"}),
    (1, {
        0: "Phase 1 finished on time and 4% under budget",
        13: "Zone 1 rollout at Weil and Tunis",
        18: "312 luminaires replaced in Zone 1 areas\nEnergy use down 33% versus legacy fittings\n"
        "No unplanned downtime during installation\nCertification audit passed on 3 November",
        22: "Source: plant energy reports, October 2026",
    }),
    (2, {
        0: "Retrofit pays back faster than full replacement",
        13: "Two options for phase 2",
        18: "Retrofit\nEUR 410k total cost\nPayback in 2.1 years\nInstalled in 6 weeks",
        26: "Full replacement\nEUR 690k total cost\nPayback in 3.4 years\nInstalled in 11 weeks",
    }),
    (3, {
        0: "Three sites, three different bottlenecks",
        18: "Weil\nSupplier lead time 14 weeks\nSpare parts for 6 months",
        26: "Tunis\nInstallers booked until March\nPermits pending for Hall 3",
        27: "Cologne\nShutdown window in July\nTraining in December",
    }),
    (8, {
        0: "Every site cut energy use by at least a quarter",
        13: "Energy use per site, MWh",
        18: "Weil: −38%\nTunis: −32%\nCologne: −31%\nHouston: −27%",
        28: CHART,
    }),
    (4, {0: "Three of four sites are finished", 13: "Status per site", TABLE: TABLE}),
    (15, {0: "Approve phase 2 by 1 December", 12: "Installation can then start in January"}),
    (34, {0: "Thank you", 10: "Questions?"}),
]  # fmt: skip

TABLE_ROWS = [
    ["Site", "Luminaires", "Energy saved", "Status"],
    ["Weil", "124", "38%", "Done"],
    ["Tunis", "96", "32%", "Done"],
    ["Cologne", "92", "31%", "Done"],
    ["Houston", "64", "27%", "In progress"],
]


def add_chart(placeholder) -> None:
    data = CategoryChartData()
    data.categories = ["Weil", "Tunis", "Cologne", "Houston"]
    data.add_series("2025", (410, 280, 350, 220))
    data.add_series("2026", (255, 190, 240, 160))
    chart = placeholder.insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, data).chart
    chart.font.size = Pt(12)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(11)


def add_table(slide) -> None:
    """The Master has no table layout, so the table goes where the body text would be."""
    shape = slide.shapes.add_table(
        len(TABLE_ROWS), len(TABLE_ROWS[0]), Inches(0.72), Inches(1.98), Inches(9.96), Inches(2)
    )
    for r, row in enumerate(TABLE_ROWS):
        for c, value in enumerate(row):
            cell = shape.table.cell(r, c)
            cell.text = value
            cell.text_frame.paragraphs[0].runs[0].font.size = Pt(16)


def add_footer(slide, footer: str, date: str) -> None:
    """Put the layout's footer, date and slide number boxes on the slide.

    python-pptx leaves these three out when it creates a slide, so we copy them from the
    layout (the Master's own boxes, at the Master's positions) and fill them.
    """
    texts = {PP_PLACEHOLDER.FOOTER: footer, PP_PLACEHOLDER.DATE: date}
    for layout_ph in slide.slide_layout.placeholders:
        kind = layout_ph.placeholder_format.type
        if kind not in (*texts, PP_PLACEHOLDER.SLIDE_NUMBER):
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


def build_showcase(template: Path) -> Path:
    prs = Presentation(template)
    remove_sample_slides(prs)  # the Master ships with three
    for layout_index, texts in SLIDES:
        slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
        for placeholder in list(slide.placeholders):
            value = texts.get(placeholder.placeholder_format.idx)
            if value == CHART:
                add_chart(placeholder)
            elif value:
                placeholder.text_frame.text = value
            else:
                placeholder.element.getparent().remove(placeholder.element)  # unused box
        if TABLE in texts:
            add_table(slide)
        add_footer(slide, FOOTER, DATE)
    SHOWCASE_DIR.mkdir(parents=True, exist_ok=True)
    path = SHOWCASE_DIR / template.name
    prs.save(path)
    return path


def main() -> None:
    for template in sorted(TEMPLATE_DIR.glob("*.pptx")):
        print(build_showcase(template))


if __name__ == "__main__":
    main()
