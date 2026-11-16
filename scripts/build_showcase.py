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

# Layout name -> text for each placeholder, in placeholder index order.
SLIDES = {
    "Title Slide": ["Ex-Lighting Upgrade 2026", "Status update for plant managers · 17 Nov 2026"],
    "Title and Content": [
        "Phase 1 finished on time and 4% under budget",
        "312 luminaires replaced in Zone 1 areas\nEnergy use down 33% versus legacy fittings\n"
        "No unplanned downtime during installation\nCertification audit passed on 3 November",
    ],
    "Chart and Text": [
        "Every site cut energy use by at least a quarter",
        None,  # chart
        "Weil: −38%\nTunis: −32%\nCologne: −31%\nHouston: −27%",
    ],
    "Comparison": [
        "LED retrofit pays back faster than full replacement",
        "Retrofit", "EUR 410k total cost\nPayback in 2.1 years\nInstalled in 6 weeks",
        "Full replacement", "EUR 690k total cost\nPayback in 3.4 years\nInstalled in 11 weeks",
    ],
    "Two Content": [
        "Two sites, two different bottlenecks",
        "Weil: supplier lead time 14 weeks\nSpare parts stocked for 6 months\nTeam fully trained",
        "Tunis: installers booked until March\nPermits pending for Hall 3\nTraining in December",
    ],
    "Table": ["Three of four sites are finished", None],  # table
    "Key Message": [
        "Phase 2 can start in January if the budget is approved by 1 December",
        "Decision needed from the steering committee",
    ],
    "Title Only": ["Questions and next steps"],
}  # fmt: skip

TABLE = [
    ["Site", "Luminaires", "Energy saved", "Status"],
    ["Weil", "124", "38%", "Done"],
    ["Tunis", "96", "32%", "Done"],
    ["Cologne", "92", "31%", "Done"],
    ["Houston", "64", "27%", "In progress"],
]


def add_chart(placeholder) -> None:
    data = CategoryChartData()
    data.categories = ["Weil", "Tunis", "Cologne", "Houston"]
    data.add_series("2025 (MWh)", (410, 280, 350, 220))
    data.add_series("2026 (MWh)", (255, 190, 240, 160))
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
    table = placeholder.insert_table(len(TABLE), len(TABLE[0])).table
    for r, row in enumerate(TABLE):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            cell.text_frame.paragraphs[0].runs[0].font.size = Pt(16)


def build_showcase(template: Path) -> Path:
    prs = Presentation(template)
    layouts = {layout.name: layout for layout in prs.slide_layouts}
    for name, texts in SLIDES.items():
        slide = prs.slides.add_slide(layouts[name])
        placeholders = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)
        for placeholder, value in zip(placeholders, texts, strict=True):
            if value is not None:
                placeholder.text_frame.text = value
            elif name == "Chart and Text":
                add_chart(placeholder)
            else:
                add_table(placeholder)
    SHOWCASE_DIR.mkdir(parents=True, exist_ok=True)
    path = SHOWCASE_DIR / template.name
    prs.save(path)
    return path


def main() -> None:
    for template in sorted(TEMPLATE_DIR.glob("*.pptx")):
        print(build_showcase(template))


if __name__ == "__main__":
    main()
