# slidegen

Turn rough notes into a short, well-structured PowerPoint deck (at most 4 slides) with one
structured LLM call, validated output and a template-driven layout.

> Work in progress. Sections below are filled in as each part lands.

## Install

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```bash
uv sync
```

## API key

The tool reads `OPENROUTER_API_KEY` from the environment or from a local `.env` file:

```
OPENROUTER_API_KEY=your-key-here
```

`.env` is gitignored, and CI scans every push for committed secrets. The key is never printed
or logged.

## Templates

The company template was not provided, so the repository ships its own, styled after the
R. STAHL website: navy `#072A5B`, light blue `#D3E5F7`, orange highlight `#EB7D00`, and Arial
as a stand-in for the licensed Univers typeface (the website itself falls back to Arial).

`slidegen/templates/` holds a base template and three team variants, each with its own look:

| Template | Look |
|---|---|
| base | corporate: navy header band, brand-colour mosaic title slide, navy quote-style key message |
| marketing | bold: raspberry and orange, overlapping circles |
| engineering | technical, dark: blueprint grid, teal corner brackets |
| finance | light, data-driven: navy side rail, bar-chart motifs, green accents |

All four share the same eight layouts and placeholder indexes (Title Slide, Title and Content,
Key Message, Two Content, Comparison, Title Only, Chart and Text, Table), so the code never
changes per team. The team colours sit in the theme accents, so native charts and tables pick
them up automatically.

python-pptx cannot create layouts from scratch, so `scripts/build_templates.py` starts from the
python-pptx default template and reshapes its theme, master and layouts. The four looks live in
`scripts/template_designs.py`.

```bash
uv run python scripts/build_templates.py   # writes slidegen/templates/*.pptx
uv run python scripts/build_showcase.py    # one demo slide per layout, with a chart and a table
uv run python scripts/render.py out/showcase/*.pptx
```

`scripts/render.py` exports every slide to PNG through PowerPoint, so the check shows exactly
what a PowerPoint user will see (Windows only, development use; the tool itself does not need it).

## Development

```bash
uv run ruff check .
uv run ruff format --check .
uv run pytest
```

Tests run fully offline: a shared fixture removes the API key and blocks every network
connection, so the model is always mocked.
