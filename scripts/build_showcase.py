"""Fill every template with one demo slide per layout, including a native chart and table.

Only used to look at the templates (render with scripts/render.py); not part of the tool.

Run:  uv run python scripts/build_showcase.py
Writes out/showcase/<team>.pptx
"""

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT / "slidegen" / "templates"
SHOWCASE_DIR = ROOT / "out" / "showcase"

CHART, TABLE = "<chart>", "<table>"

# Layout name -> {placeholder index: text}. Indexes 13 and 22 are subheadline and source.
SLIDES = {
    "Title Slide": {
        0: "Ex-Lighting Upgrade 2026",
        1: "Plant management update\nTunis, 17 November 2026",
    },
    "Title and Content": {
        0: "Phase 1 finished on time and 4% under budget",
        13: "Zone 1 rollout at Weil and Tunis",
        1: "312 luminaires replaced in Zone 1 areas\nEnergy use down 33% versus legacy fittings\n"
        "No unplanned downtime during installation\nCertification audit passed on 3 November",
        22: "Source: plant energy reports, October 2026",
    },
    "Title and 2 Columns": {
        0: "Retrofit pays back faster than full replacement",
        13: "Two options for phase 2",
        1: "Retrofit\nEUR 410k total cost\nPayback in 2.1 years\nInstalled in 6 weeks",
        2: "Full replacement\nEUR 690k total cost\nPayback in 3.4 years\nInstalled in 11 weeks",
    },
    "Title and 3 Columns": {
        0: "Three sites, three different bottlenecks",
        2: "Weil\nSupplier lead time 14 weeks\nSpare parts for 6 months",
        3: "Tunis\nInstallers booked until March\nPermits pending for Hall 3",
        4: "Cologne\nShutdown window in July\nTraining in December",
    },
    "Title, Text and Chart": {
        0: "Every site cut energy use by at least a quarter",
        13: "Energy use per site, MWh",
        2: "Weil: −38%\nTunis: −32%\nCologne: −31%\nHouston: −27%",
        1: CHART,
    },
    "Title and Table": {0: "Three of four sites are finished", 1: TABLE},
    "Separation Slide": {
        0: "Approve phase 2 by 1 December",
        1: "Installation can then start in January",
    },
    "End Slide": {0: "Thank you", 1: "Questions?"},
}  # fmt: skip

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


def add_table(placeholder) -> None:
    table = placeholder.insert_table(len(TABLE_ROWS), len(TABLE_ROWS[0])).table
    for r, row in enumerate(TABLE_ROWS):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            cell.text_frame.paragraphs[0].runs[0].font.size = Pt(16)


def build_showcase(template: Path) -> Path:
    prs = Presentation(template)
    layouts = {layout.name: layout for layout in prs.slide_layouts}
    for name, texts in SLIDES.items():
        slide = prs.slides.add_slide(layouts[name])
        for placeholder in list(slide.placeholders):
            value = texts.get(placeholder.placeholder_format.idx)
            if value == CHART:
                add_chart(placeholder)
            elif value == TABLE:
                add_table(placeholder)
            elif value:
                placeholder.text_frame.text = value
            else:
                placeholder.element.getparent().remove(placeholder.element)  # unused box
    SHOWCASE_DIR.mkdir(parents=True, exist_ok=True)
    path = SHOWCASE_DIR / template.name
    prs.save(path)
    return path


def main() -> None:
    for template in sorted(TEMPLATE_DIR.glob("*.pptx")):
        print(build_showcase(template))


if __name__ == "__main__":
    main()
