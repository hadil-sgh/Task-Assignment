# slidegen

Turn rough notes into a short, well-structured PowerPoint deck (at most 4 slides) with one
structured LLM call, validated output and a template-driven layout.

> Work in progress. Sections below are filled in as each part lands.

## Install

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```bash
uv sync
```

## Usage

```bash
uv run python -m slidegen "rough notes..." --out deck.pptx
cat notes.txt | uv run python -m slidegen --verbose        # notes from stdin
uv run python -m slidegen "notes" --team auto              # team template picked from the notes
uv run python -m slidegen "notes" --mock repair --verbose  # no key, no network
```

| Option | Meaning |
|---|---|
| `--out` | output file (default `deck.pptx`) |
| `--model` | OpenRouter model id |
| `--team` | `base` (the company Master), `marketing`, `engineering`, `finance`, or `auto` |
| `--template` | any other .pptx template |
| `--verbose` | outcome, tokens, cost and latency per call, warnings |
| `--mock` | canned answers: `ok`, `repair` (bad first answer) or `fallback` (only bad answers) |

Exit codes: 0 deck written and checked, 1 deck written but the check found problems,
2 empty notes or bad arguments, 3 API key missing or rejected.

**One run** (`slidegen/pipeline.py`): pick the template and load its cached catalog, make one
structured model call, validate it; if it is invalid, make one repair call with the error
list (never a loop); if that fails too or the API cannot be reached, build a plain deck from
the notes (each line becomes text, at most 4 slides, nothing dropped). Then build the deck
and check the saved file. At most two model calls per run, no agent loop, no code generation.
A missing or rejected key is not hidden by the fallback: it stops with a clear message.

## API key

The tool reads `OPENROUTER_API_KEY` from the environment or from a local `.env` file:

```
OPENROUTER_API_KEY=your-key-here
```

`.env` is gitignored, and CI scans every push for committed secrets. The key is never printed
or logged.

## The model call

`slidegen/llm.py` makes one call to OpenRouter's OpenAI-compatible endpoint with a strict JSON
schema response format, and records tokens, cost (reported by OpenRouter) and latency.

- Timeouts, rate limits (429) and server errors (5xx) are retried up to 3 times, waiting
  1, 2 and 4 seconds, or what the server's `Retry-After` header asks for.
- A rejected key (401/403) stops at once with a clear message; retrying cannot help.
- Other client errors (4xx) fail at once.
- `MockClient` returns canned answers with no key and no network, for tests and `--mock`.

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
3. checks every slide against its layout: known layout, required slots filled, and length
   limits (text in a slot the layout does not have is ignored with a warning, since it would
   never be placed on the slide);
4. returns every problem at once as `path: message`, for example
   `slides[1].head2: must not be empty`, ready for one repair call.

Each text has a **budget** (what fits at the template font size) and a **hard limit** at twice
the budget. Between the two, the text is accepted and the font is reduced when the deck is
built; above the hard limit it triggers the repair call.

**Prompt injection.** The notes are wrapped in `<notes>` tags, the system prompt says their
content is data and never instructions, and any `<notes>`/`</notes>` tag inside the notes is
neutralised so it cannot close the block. Validation is the second layer: whatever the model
returns must still be a valid deck.

## Fitting text and building the deck

**No overflow** (`slidegen/fit.py`). Each text is measured with Pillow and the real font file
(Arial Narrow for the Master), wrapped to the box width with a 5% safety margin (PowerPoint
lays text out slightly wider than Pillow), and checked against the box height. If it does not
fit, the font shrinks one point at a time down to a floor of 60% of the template size, never
below 12 pt. Only at the floor are words cut, ending with "…", and a warning is reported.
Without the font file (for example on Linux CI) the catalog's measured average character
width is used instead. The chosen size is written on every run; PowerPoint autofit is not
used, because it only recalculates when someone edits the slide.

**Building** (`slidegen/deck.py`). The template's sample slides are removed, one slide is added
per deck slide from the chosen layout, placeholders are filled by index, empty placeholders
are removed, and the layout's footer, date and slide number are added (python-pptx leaves
these out of new slides). The footer defaults to the first slide's title, the date to today.

**Independent check.** After saving, the file is opened again and checked from what is really
in it: 1 to 4 slides, no empty placeholders, and no text that overflows its box.

## Tables and charts

The model may add a **table** when the notes compare several options on several figures, or a
**bar chart** when one figure is compared across items or for percentages and shares.

- The catalog offers the Master's diagram layouts (text on the left, a chart box on the
  right); layouts with several charts are left out.
- A table takes the place of the layout's first bullet box, at its exact position, because the
  Master has no table layout. A chart goes into the layout's chart box.
- **Every number in a table or chart must appear in the notes.** An invented or calculated
  number is a validation error and goes to the repair call. Limits: 8 rows by 6 columns and 40 characters per cell, 8 categories, 3 series.
  Errors say exactly what is wrong ("7 columns and 9 rows; use 2 to 6 columns and 1 to 8
  rows"), so the single repair call can fix them.
- Both are native PowerPoint objects, editable after generation and coloured by the theme, so
  they follow the team template. Bars start at zero, one colour per series. Table rows grow
  when a cell wraps, the font goes from 16 down to 10 pt to fit the box, and the header text
  switches to dark on a light team colour.

## Development

```bash
uv run ruff check .
uv run ruff format --check .
uv run pytest
```

Tests run fully offline: a shared fixture removes the API key and blocks every network
connection, so the model is always mocked.
