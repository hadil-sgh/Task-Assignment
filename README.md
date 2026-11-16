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

**The company Master is the default** (`slidegen/templates/master.pptx`). The code reads its
layouts directly; nothing in the code is specific to it, and `--template` accepts any other
.pptx.

The three team templates are **the Master itself, in a different colour**.
`scripts/build_templates.py` copies `master.pptx` and gives each team its own hue and style,
while every element of the Master stays where it is:

1. the Master's triangle mosaic on the first and last slide is recoloured in the team hue
   (PowerPoint's duotone effect: same picture, same pattern, new colour); the slogan, logo,
   address and all text keep the Master's colours, which are pinned first;
2. the team colour becomes the main theme colour (accent 1), which the Master uses for titles,
   subheadlines, chart series and table headers;
3. bullet marks on the content boxes, a background for the content and separation slides
   (engineering is dark), and decorative shapes in the free margins, drawn behind the
   Master's footer, date, slide number and logo.

All layouts, text boxes, fonts, the optional subheadline and source line, the footer, date
and slide number are the Master's in every team.

| Template | Look |
|---|---|
| base (the Master) | the company template, unchanged: green-to-teal mosaic, navy titles |
| marketing | raspberry mosaic, raspberry titles, orange accents, raspberry band and circles |
| engineering | blue mosaic, dark slides with a blueprint grid, cyan titles and arrow bullets |
| finance | amber mosaic, amber titles, navy side rail, amber rule and bar-chart motif |

Choose with `--team base|marketing|engineering|finance`, or `--team auto` to pick one from the
content of the notes. Because the layouts and placeholders are the Master's in every team,
the generator works the same with all of them, and a wrong pick only changes the look.

```bash
uv run python scripts/build_templates.py   # writes slidegen/templates/<team>.pptx
uv run python scripts/build_showcase.py    # demo slides per template, with a chart and a table
uv run python scripts/render.py out/showcase/*.pptx
```

`scripts/render.py` exports every slide to PNG through PowerPoint, so the check shows exactly
what a PowerPoint user will see (Windows only, development use; the tool itself does not need it).

## How the model learns the template

`slidegen/template.py` reads the template once with python-pptx and keeps only the text
layouts (no picture, chart or table placeholders). For each placeholder it computes a
character budget from the box size, inner margins, indent, line spacing and the font the
template sets (size, bold, narrow, uppercase), and turns the result into a short catalog.
For the company template (35 layouts, of which 7 distinct text layouts are kept):

```
L0 9_Title slide refinery: title<=42, subtitle?<=100 "Presenter, Place, Date"
L1 Titel und Inhalt: title<=51, head?<=120 "Optional Subheadline", body<=11x85, head2?<=171 "Optional: Sources and/or captions"
L2 Title and 2 Columns: title<=51, head?<=120 "Optional Subheadline", body<=11x48, head2?<=171 "Optional: Sources and/or captions", body2<=11x48
...
L15 Separation Slide: title<=44, subtitle?<=85 "Optional Subtext"
```

- `body<=11x85` means up to 11 bullets of 85 characters.
- `?` marks an optional slot. A slot is optional when the template's own prompt text says
  "Optional ...", or when it is a one-line subtitle. The model never has to invent a source
  line or a presenter name.
- The quoted hint is the template's prompt text, so the model knows what each box is for.
- Layouts with identical slots (the template's white and blue variants) are listed once.

The model never sees template XML. The catalog is cached as `<template>.catalog.json`, keyed
by a hash of the template file, so it rebuilds automatically when the template changes.

**Which team template (`--team auto`)?** Chosen from the content of the notes: a short word
list per team (English, French, German), plus money and percentage figures counting towards
finance. The base template is used when the signal is weak or tied. A keyword score costs no
tokens, gives the same answer every time and is easy to test. A wrong pick is harmless: every
team template has the same layouts, so only the look changes.

## Structured output and validation

The model answers with JSON that follows a strict schema built from the catalog: per slide a
`layout` (one of the catalog ids) and one field per slot name (`title`, `subtitle`, `head`,
`head2` as text; `body`, `body2` as bullet lists), with unused slots left empty. The wire schema
stays flat and has no length keywords, because providers support only part of JSON Schema in
strict mode; limits are given in descriptions and the catalog, and enforced after the call.

`slidegen/schema.py` then:

1. pulls one JSON object out of the reply, even inside code fences or prose;
2. checks types and 1 to 4 slides with Pydantic;
3. checks every slide against its layout: known layout, no text in slots the layout lacks,
   required slots filled, and length limits;
4. returns every problem at once as `path: message`, for example
   `slides[1].head2: must not be empty`, ready for one repair call.

Each text has a **budget** (what fits at the template font size) and a **hard limit** at twice
the budget. Between the two, the text is accepted and the font is reduced when the deck is
built; above the hard limit it triggers the repair call.

**Prompt injection.** The notes are wrapped in `<notes>` tags, the system prompt says their
content is data and never instructions, and any `<notes>`/`</notes>` tag inside the notes is
neutralised so it cannot close the block. Validation is the second layer: whatever the model
returns must still be a valid deck.

## Development

```bash
uv run ruff check .
uv run ruff format --check .
uv run pytest
```

Tests run fully offline: a shared fixture removes the API key and blocks every network
connection, so the model is always mocked.
