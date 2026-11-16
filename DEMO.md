# Demo script (about 15 minutes)

Before the call: open a terminal in the repository, run `uv sync`, check that `.env` holds the
key, and open `samples/previews/framework/overview.png` and the notebook in tabs.

## 1. The problem and the idea (1 minute)

Today an LLM writes and runs a Python script for every deck: about a minute per deck, many
tokens, and generated code running on our machines. slidegen turns that around: **the model
only writes the content**, as JSON that follows a schema built from the template, and plain
Python builds the deck. One call, about 4 seconds, under 1 cent.

## 2. Everything is tested (1 minute)

```bash
uv run pytest -q
```

150 tests, no key and no network needed. The same tests, ruff and a secret scan run in CI on
every pull request.

## 3. A live deck (3 minutes)

```bash
uv run python -m slidegen "$(cat samples/framework_notes.txt)" --out out/demo.pptx --verbose
```

Point at the output: **first try**, tokens, cost (about $0.008), time (about 4 s). Open
`out/demo.pptx` in PowerPoint and show:

- the Master's own layouts, footer, date and slide number
- the comparison as a **native table** (editable, not an image), with only numbers from the
  notes; when the notes compare one figure across items, the model picks a native bar chart
- the titles state the point ("Choose Next.js for delivery confidence"), not a topic label

Each run can word the deck a little differently; `samples/framework.pptx` is one such run.

## 4. Team templates (1 minute)

```bash
uv run python -m slidegen "$(cat samples/framework_notes.txt)" --team engineering --out out/demo_eng.pptx
```

Same notes, same layouts, the engineering colours. `--team auto` picks the team from the notes.
Any other company .pptx works with `--template`.

## 5. When the model answers badly (2 minutes)

No key or network needed here:

```bash
uv run python -m slidegen "$(cat samples/framework_notes.txt)" --mock repair --verbose
```

The first answer breaks the rules; the exact errors go back in **one** repair call, and the
second answer is used. Then:

```bash
uv run python -m slidegen "$(cat samples/framework_notes.txt)" --mock fallback --verbose
```

Both answers are bad, so the tool builds a plain deck straight from the notes. It never
crashes and never loops. A missing or wrong key is the one case that stops with a clear message
(exit code 3), because a silent fallback would hide it.

## 6. How the model was chosen (2 minutes)

Open `notebooks/evaluate.ipynb`, scroll to the results: 8 notes, 3 models, 2 runs each, scored
blind by a stronger judge, twice. GPT-5.6 Terra: 100% first try, best quality, cheapest,
fastest. Then the honest finding: the decks are too compressed (one idea per slide is the next
prompt fix, see `NOTE.md`).

## 7. Code walkthrough (4 minutes)

Follow one run through the files, one sentence each:

| File | One sentence |
|---|---|
| `__main__.py` | reads the options and the notes, maps the result to an exit code |
| `pipeline.py` | one call, at most one repair, else the fallback; then build and check |
| `template.py` | reads the template once into a short catalog of layouts and box budgets, cached |
| `prompts.py` | the rules, the catalog, and the notes inside `<notes>` tags against injection |
| `llm.py` | OpenRouter with strict JSON schema, retries, real token and cost numbers |
| `schema.py` | validates the answer: schema, budgets, numbers that must come from the notes |
| `fit.py` | measures text with the real font and shrinks it a point at a time to fit |
| `deck.py` | fills the placeholders, draws tables and charts, checks the saved file again |

## Backup plan

If the network or the key fails during the call: use `--mock ok` for a deck without the
network, and show the committed samples in `samples/` (the decks, their PNG previews and the
notes they came from). The notebook is committed with its results, so it needs no rerun.
