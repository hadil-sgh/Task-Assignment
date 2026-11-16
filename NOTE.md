# Note: measuring slide quality, and what to improve next

## How to measure quality

Quality has two parts: is the deck **correct** (it follows the notes, nothing breaks), and is it
**good** (a reader gets the point fast). The first can be checked by code on every run; the
second needs a judge, and then people.

### 1. Automatic checks on every deck (already in the tool)

These are free and run in production on every deck:

| Check | Where |
|---|---|
| The answer follows the JSON schema; layouts and slots exist | `schema.py` |
| Every text is within its box's budget (hard limit at 2x) | `schema.py` |
| Every number in a table or chart appears in the notes | `schema.py` |
| Text fits its box with the real font, at a readable size | `fit.py` |
| The saved file opens, has 1 to 4 slides, no empty placeholders, no overflow | `deck.py` |
| How the deck was made: first try, repaired, or fallback | `pipeline.py` |

The share of decks made on the first try, repaired or falling back is the first health number
to watch: a rising repair or fallback rate means the prompt, the model or the template changed
for the worse.

### 2. A blind LLM judge on a fixed set of notes (done in `notebooks/evaluate.ipynb`)

A fixed set of notes is run through the tool, and a stronger model scores each deck from 1 to 5
on five criteria:

- **faithfulness:** no facts or numbers that are not in the notes, nothing distorted
- **coverage:** every key point of the notes is on a slide
- **conciseness:** short titles that state the point, bullets as fragments, no filler
- **layout fit:** the layout matches the content (a comparison side by side, figures in a chart)
- **flow:** the slides tell the story in a sensible order

The judge never sees the model name, the decks are shuffled, and every deck is scored twice to
check the judge is stable. This set doubles as a **regression test**: run it before changing
the prompt, the model or the catalog, and compare.

### 3. People

The judge only reads text. Two human measures matter more in the end:

- **Review sessions:** a few colleagues rate a sample of decks (rendered as images) with the
  same five criteria. This also checks that the judge agrees with people.
- **Edit rate in production:** how much users change a deck before they use it (text edited,
  slides deleted, layouts changed). Little editing is the best sign of a good deck. A thumbs
  up or down next to each generated deck is a cheap start.

## Results so far

8 notes x 3 models x 2 runs, judged twice by Claude Opus 5.5 (total spent: $1.46):

| Model | First try | Fallback | Quality (1-5) | Cost per deck | Time |
|---|---|---|---|---|---|
| GPT-5.6 Terra (default) | 100% | 0% | 3.83 | $0.003 | 3.0 s |
| Claude Sonnet 5.5 | 94% | 0% | 3.58 | $0.011 | 4.6 s |
| Qwen 3.8 Max | 0% | 100% | (fallback only) | - | - |

What the numbers say:

- **Correctness is solved:** with both working models, every deck was valid, and no deck needed
  the fallback. The repair call handled the few rejected answers.
- **Faithfulness and coverage are high** (about 4.2 and 4.7): the number guard works, and the
  models keep the content.
- **Conciseness, layout fit and flow are the weak points** (about 3 to 3.6). The decks are
  compressed: 1.25 slides on average with Sonnet and 1.9 with GPT, for notes that often deserve
  three. The models squeeze everything onto as few slides as possible.
- **The fallback deck is faithful but not a good deck:** Qwen's plain decks scored high on
  faithfulness and coverage (they copy the notes) and low on everything else. That is the right
  trade for a safety net.
- **The judge is stable:** its two scorings differed by 0.11 points on average.

Cost note: a live deck from the command line costs about $0.008 (the samples), more than the
notebook's $0.003 average, because the sample notes are longer and use tables and charts. Both
are **under 1 cent per deck**.

## What to improve next

In order of value for effort:

1. **One idea per slide.** The main quality finding. Change the rule "use as few slides as the
   content needs" to "one idea per slide, up to 4", and add the slide count to the judge's
   view. Then rerun the evaluation to confirm conciseness and flow go up. A small prompt change.
2. **A larger, harder evaluation set.** 8 notes is a start. Add 30 to 50 real notes from the
   teams: long notes, very short notes, German and French notes, notes with many figures. Add
   people's ratings on a sample to check the judge.
3. **Qwen and other models.** Find why Qwen's requests fail (most likely its provider does not
   accept strict JSON schema), and fall back to plain JSON mode for such models. More models
   means cheaper options and less dependence on one vendor.
4. **Judge the rendered slides, not only the text.** Send the PNG previews to a vision model to
   catch what text cannot show: a crowded slide, a chart that is hard to read.
5. **Speaker notes.** The schema can carry a short speaker note per slide, written into the
   notes page, so details left out of the bullets are not lost.
6. **Prompt caching.** The system prompt and catalog are the same for every deck of a template.
   With provider caching, most input tokens become cheap and the first token arrives sooner.
7. **A dedicated table layout in the templates.** Today a table takes the place of the bullet
   box. A layout with a full-width table box would give tables more room.
