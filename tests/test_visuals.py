"""Native tables and charts: the catalog, the checks (only numbers from the notes), the build."""

import json
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu

from slidegen import fit
from slidegen.deck import build_deck, check_deck_file
from slidegen.llm import MockClient
from slidegen.pipeline import generate
from slidegen.schema import DeckError, parse_deck, wire_schema
from slidegen.template import MASTER, Catalog

NOTES = (
    "Website relaunch: which framework? Next.js test page 1.1 s, build EUR 150k, 18 weeks. "
    "Nuxt 1.3 s, EUR 165k, 21 weeks. Astro 0.8 s, EUR 140k, 19 weeks."
)
CHART_SLIDE = {
    "layout": "L8",
    "title": "Next.js is fast enough and fits the team",
    "head": "",
    "body": ["Fits the React skills of the team", "Astro is faster but needs React too"],
    "head2": "",
    "head3": "Mobile load time in seconds",
    "chart": {
        "title": "Load time (s)",
        "categories": ["Next.js", "Nuxt", "Astro"],
        "series": [{"name": "Load time", "values": [1.1, 1.3, 0.8]}],
    },
}
TABLE_SLIDE = {
    "layout": "L1",
    "title": "Three options compared",
    "head": "",
    "body": [],
    "head2": "",
    "table": {
        "columns": ["Option", "Load (s)", "Cost (EUR k)", "Weeks"],
        "rows": [["Next.js", "1.1", "150", "18"], ["Nuxt", "1.3", "165", "21"],
                 ["Astro", "0.8", "140", "19"]],
    },
}  # fmt: skip


@pytest.fixture(autouse=True)
def estimated_widths(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(fit, "load_font", lambda family, bold: None)


def answer(*slides: dict) -> str:
    return json.dumps({"slides": list(slides)})


def changed(slide: dict, **changes) -> dict:
    return json.loads(json.dumps(slide)) | changes


def errors_of(text: str, catalog: Catalog) -> list[str]:
    with pytest.raises(DeckError) as caught:
        parse_deck(text, catalog, NOTES)
    return caught.value.errors


# ---------------------------------------------------------------- catalog and wire schema


def test_catalog_offers_the_masters_chart_layout(catalog: Catalog) -> None:
    diagram = catalog.layout("L8")
    assert [s.name for s in diagram.slots][-2:] == ["head3", "chart"]
    assert diagram.slots[-1].describe() == "chart"
    assert "L10" not in [layout.id for layout in catalog.layouts]  # two charts: left out


def test_wire_schema_has_strict_table_and_chart_objects(catalog: Catalog) -> None:
    props = wire_schema(catalog)["schema"]["properties"]["slides"]["items"]["properties"]
    assert props["table"]["required"] == ["columns", "rows"]
    assert props["chart"]["required"] == ["title", "categories", "series"]
    assert props["chart"]["additionalProperties"] is False
    assert "only numbers that appear in the notes" in props["chart"]["description"].lower()


# ---------------------------------------------------------------- checks


def test_valid_chart_and_table_pass(catalog: Catalog) -> None:
    deck, _ = parse_deck(answer(CHART_SLIDE, TABLE_SLIDE), catalog, NOTES)
    assert deck.slides[0].chart.categories == ["Next.js", "Nuxt", "Astro"]
    assert deck.slides[1].table.rows[2] == ["Astro", "0.8", "140", "19"]


def test_empty_chart_and_table_mean_not_used(catalog: Catalog) -> None:
    empty = {"chart": {"title": "", "categories": [], "series": []},
             "table": {"columns": [], "rows": []}}  # fmt: skip
    slide = {"layout": "L1", "title": "Plain slide", "body": ["One point"], **empty}
    deck, _ = parse_deck(answer(slide), catalog, NOTES)
    assert deck.slides[0].chart is None and deck.slides[0].table is None


def test_invented_numbers_are_rejected(catalog: Catalog) -> None:
    chart = changed(CHART_SLIDE)
    chart["chart"]["series"][0]["values"] = [1.1, 1.3, 2.5]
    table = changed(TABLE_SLIDE)
    table["table"]["rows"][0][2] = "200"
    errors = errors_of(answer(chart, table), catalog)
    assert errors == [
        "slides[0].chart: 2.5 does not appear in the notes; use only figures from the notes, "
        "never calculated ones",
        "slides[1].table: 200 does not appear in the notes; use only figures from the notes, "
        "never calculated ones",
    ]


def test_chart_needs_a_chart_layout_and_a_chart_layout_needs_a_chart(catalog: Catalog) -> None:
    on_text_layout = changed(TABLE_SLIDE, table={"columns": [], "rows": []}, body=["Point"],
                             chart=CHART_SLIDE["chart"])  # fmt: skip
    without_chart = changed(CHART_SLIDE, chart={"title": "", "categories": [], "series": []})
    errors = errors_of(answer(on_text_layout, without_chart), catalog)
    assert "slides[0].chart: this layout has no chart box" in errors[0]
    assert "slides[1].chart: this layout has a chart box" in errors[1]


def test_table_replaces_the_bullet_box(catalog: Catalog) -> None:
    errors = errors_of(answer(changed(TABLE_SLIDE, body=["Also bullets"])), catalog)
    assert errors == ["slides[0].body: leave it empty, the table takes its place"]


def test_limit_errors_say_exactly_what_is_wrong(catalog: Catalog) -> None:
    """Precise messages, so the single repair call can fix them."""
    table = changed(TABLE_SLIDE)
    table["table"]["rows"] = [["Next.js", "1.1", "150", "18"]] * 8 + [["Nuxt", "1.3"]]
    table["table"]["columns"][0] = "Framework option under consideration here"  # 41 characters
    chart = changed(CHART_SLIDE)
    chart["chart"]["series"][0]["values"] = [1.1, 1.3]
    errors = errors_of(answer(table, chart), catalog)
    assert errors == [
        "slides[0].table: 4 columns and 9 rows; use 2 to 6 columns and 1 to 8 rows "
        "(merge or drop columns)",
        "slides[0].table: row 9 has 2 cells for 4 columns",
        'slides[0].table: cell "Framework option under consideration here" has 41 characters; '
        "shorten it to 40",
        "slides[1].chart: series 'Load time' has 2 values for 3 categories",
    ]


def test_tables_up_to_8_rows_and_6_columns_fit_the_box(tmp_path: Path, catalog: Catalog) -> None:
    columns = ["Option", "Load (s)", "Cost (EUR k)", "Weeks", "Team fit", "Risk"]
    rows = [["Next.js", "1.1", "150", "18", "React skills match", "Vendor lock-in"]] * 8
    big = changed(TABLE_SLIDE, table={"columns": columns, "rows": rows})
    deck, _ = parse_deck(answer(big), catalog, NOTES)
    out = tmp_path / "deck.pptx"
    build_deck(deck, catalog, MASTER, out)
    [frame] = [s for s in Presentation(out).slides[0].shapes if s.has_table]
    assert Emu(frame.top + frame.height).inches <= 1.98 + 4.62 + 0.01  # inside the body box


def test_a_table_still_wrong_after_repair_becomes_bullets(tmp_path: Path) -> None:
    too_wide = changed(TABLE_SLIDE)
    too_wide["table"]["columns"] = ["Option", "a", "b", "c", "d", "e", "f"]
    too_wide["table"]["rows"] = [["Next.js", "1.1", "150", "18", "x", "y", "z"]]
    mock = MockClient([answer(too_wide, CHART_SLIDE)])  # the repair repeats the same answer
    result = generate(NOTES, mock, tmp_path / "deck.pptx")
    assert result.outcome == "repaired"  # the model's deck is kept, not the plain fallback
    assert "slides[0]: table or chart shown as bullets" in result.warnings
    first, second = Presentation(result.path).slides
    assert "Next.js: a 1.1, b 150, c 18" in first.placeholders[18].text_frame.text
    assert any(getattr(s, "has_chart", False) and s.has_chart for s in second.shapes)
    assert result.problems == []


def test_invented_numbers_are_never_kept_as_bullets(tmp_path: Path) -> None:
    invented = changed(TABLE_SLIDE)
    invented["table"]["rows"][0][2] = "999"
    result = generate(NOTES, MockClient([answer(invented)]), tmp_path / "deck.pptx")
    assert result.outcome == "fallback"  # the plain deck from the notes, without the 999


# ---------------------------------------------------------------- building


def test_chart_and_table_are_drawn_natively_in_the_layout(tmp_path: Path, catalog: Catalog) -> None:
    deck, _ = parse_deck(answer(CHART_SLIDE, TABLE_SLIDE), catalog, NOTES)
    out = tmp_path / "deck.pptx"
    build_deck(deck, catalog, MASTER, out)
    chart_slide, table_slide = Presentation(out).slides

    [chart_frame] = [
        s for s in chart_slide.shapes if getattr(s, "has_chart", False) and s.has_chart
    ]
    assert chart_frame.placeholder_format.idx == 28  # the Master's chart box
    assert list(chart_frame.chart.plots[0].categories) == ["Next.js", "Nuxt", "Astro"]

    [table_frame] = [s for s in table_slide.shapes if s.has_table]
    position = [round(Emu(v).inches, 2) for v in (table_frame.left, table_frame.top)]
    assert position == [0.72, 1.98]  # where the bullet box was
    cells = table_frame.table
    assert len(cells.rows) == 4 and cells.cell(0, 0).text == "Option"
    assert cells.cell(0, 0).text_frame.paragraphs[0].runs[0].font.bold
    assert check_deck_file(out, catalog) == []


def test_invented_number_is_repaired(tmp_path: Path) -> None:
    bad = changed(CHART_SLIDE)
    bad["chart"]["series"][0]["values"] = [1.1, 1.3, 9.9]
    mock = MockClient([answer(bad), answer(CHART_SLIDE)])
    result = generate(NOTES, mock, tmp_path / "deck.pptx")
    assert result.outcome == "repaired"
    assert "9.9 does not appear in the notes" in mock.calls[1][-1]["content"]
    assert result.problems == []


@pytest.mark.parametrize(("team", "dark"), [("master", False), ("engineering", True)])
def test_table_header_text_contrasts_with_the_team_colour(team: str, dark: bool) -> None:
    from slidegen.deck import header_text_color
    from slidegen.template import TEMPLATE_DIR

    color = header_text_color(Presentation(TEMPLATE_DIR / f"{team}.pptx"))
    assert (str(color) == "1E1E1E") is dark  # dark text on engineering's light cyan
