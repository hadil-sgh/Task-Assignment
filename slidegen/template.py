"""Read a .pptx template into a short layout catalog, and pick the team template for some notes.

The model never sees template XML. It sees a catalog of about 250 tokens, one line per
distinct text layout, for example (from the company Master):

    L2 Title and 2 Columns: title<=51, head?<=120 "Optional Subheadline", body<=11x48, ...

Each budget comes from the placeholder's box size and its font in the template:
"11x48" means 11 bullets of up to 48 characters; "51" means up to 51 characters in total.
A "?" marks an optional slot, and a quoted hint repeats the template's own prompt text.
The catalog is cached as JSON next to the template and rebuilt when the file changes.
"""

import hashlib
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu

TEMPLATE_DIR = Path(__file__).parent / "templates"
MASTER = TEMPLATE_DIR / "master.pptx"  # the company template: the default, and team "base"
TEAMS = ["base", "engineering", "finance", "marketing"]  # the others are built from MASTER
DEFAULT_TEAM = "base"
CATALOG_VERSION = 5  # bump when the budget rules change, so old caches are rebuilt

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
TEXT_TYPES = TITLE_TYPES | {PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT}
IGNORED_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}

# Average character width as a fraction of the font size, measured on sample sentences in
# Arial and Arial Narrow, plus about 5% margin. An estimate for the budget only; the fitter
# later measures the real text with the real font.
CHAR_EM = {"regular": 0.5, "regular bold": 0.53, "narrow": 0.41, "narrow bold": 0.44}
CAPS_FACTOR = 1.2  # ALL CAPS text is about 20% wider than mixed case
LINE_HEIGHT = 1.12  # Arial family: ascent + descent = 1.12 x font size at 100% spacing
GENERIC_PROMPTS = ("click to edit", "master")  # template boilerplate, not a useful hint


@dataclass
class Slot:
    name: str  # what the model fills: title, subtitle, head, body, head2, body2 ...
    idx: int  # placeholder index on the layout
    kind: str  # "text" (one string), "list" (bullets) or "chart"
    lines: int
    chars: int  # characters per line
    font_pt: float  # font size in the template
    optional: bool = False
    hint: str = ""  # the template's prompt text, e.g. "Optional: Sources and/or captions"
    # Box and font details, kept so the fitter uses exactly the catalog's numbers.
    width_pt: float = 0.0  # usable text width: box minus inner margins and indent
    height_pt: float = 0.0  # usable text height
    line_spacing: float = 1.0
    gap_pt: float = 0.0  # space before each paragraph
    font: str = ""
    bold: bool = False
    caps: bool = False

    def char_width(self, size: float) -> float:
        """Estimated average character width at a font size (see CHAR_EM)."""
        family = "narrow" if re.search(r"narrow|condensed", self.font, re.IGNORECASE) else "regular"
        style_key = f"{family} bold" if self.bold else family
        return size * CHAR_EM[style_key] * (CAPS_FACTOR if self.caps else 1)

    def line_height(self, size: float) -> float:
        return size * LINE_HEIGHT * self.line_spacing

    @property
    def budget(self) -> int:
        return self.lines * self.chars

    def describe(self) -> str:
        if self.kind == "chart":
            return "chart"
        limit = f"{self.lines}x{self.chars}" if self.kind == "list" else str(self.budget)
        hint = f' "{self.hint}"' if self.hint else ""
        return f"{self.name}{'?' if self.optional else ''}<={limit}{hint}"

    def shape(self) -> tuple:
        return (self.name, self.kind, self.lines, self.chars, self.optional)


@dataclass
class Layout:
    id: str  # "L4" = slide layout number 4 in the template
    name: str
    slots: list[Slot]

    @property
    def index(self) -> int:
        return int(self.id[1:])


@dataclass
class Catalog:
    template: str
    key: str  # template hash + catalog version
    layouts: list[Layout]

    def prompt(self) -> str:
        """The short text the model sees."""
        return "\n".join(
            f"{layout.id} {layout.name}: " + ", ".join(slot.describe() for slot in layout.slots)
            for layout in self.layouts
        )

    def layout(self, layout_id: str) -> Layout:
        return next(layout for layout in self.layouts if layout.id == layout_id)


# ---------------------------------------------------------------- catalog and cache


def load_catalog(template: Path) -> Catalog:
    """Return the catalog for a template, from the cache when the template has not changed."""
    key = f"{hashlib.sha256(template.read_bytes()).hexdigest()[:16]}-v{CATALOG_VERSION}"
    cache = template.with_suffix(".catalog.json")
    if cache.exists():
        data = json.loads(cache.read_text(encoding="utf-8"))
        if data.get("key") == key:
            return _from_dict(data)
    catalog = build_catalog(template, key)
    try:
        cache.write_text(json.dumps(asdict(catalog), indent=1), encoding="utf-8")
    except OSError:
        pass  # a read-only install still works, just without the cache
    return catalog


def build_catalog(template: Path, key: str = "") -> Catalog:
    prs = Presentation(template)
    fonts = _theme_fonts(prs.slide_master)
    layouts = []
    for index, layout in enumerate(prs.slide_layouts):
        placeholders = [
            p for p in layout.placeholders if p.placeholder_format.type not in IGNORED_TYPES
        ]
        if not _usable(placeholders):
            continue
        slots = [_measure(p, prs.slide_master, fonts) for p in placeholders]
        layouts.append(Layout(f"L{index}", layout.name, _name_slots(slots, placeholders)))
    return Catalog(template.name, key, dedupe(layouts))


def dedupe(layouts: list[Layout]) -> list[Layout]:
    """Keep the first of layouts with identical slots, e.g. a white and a blue colour variant.

    The model only needs each structure once; it saves tokens and avoids random style switches.
    """
    seen: set[tuple] = set()
    kept = []
    for layout in layouts:
        shape = tuple(slot.shape() for slot in layout.slots)
        if shape not in seen:
            seen.add(shape)
            kept.append(layout)
    return kept


def _theme_fonts(master) -> dict[str, str]:
    """Theme font names: '+mj-lt' (headings) and '+mn-lt' (body) point to these."""
    theme = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
    major = theme.find(f".//{qn('a:majorFont')}/{qn('a:latin')}")
    minor = theme.find(f".//{qn('a:minorFont')}/{qn('a:latin')}")
    return {"+mj-lt": major.get("typeface", ""), "+mn-lt": minor.get("typeface", "")}


def _usable(placeholders: list) -> bool:
    """Text layouts, plus layouts with exactly one chart box next to text.

    Layouts with pictures, tables, several charts or no text are skipped.
    """
    charts = [p for p in placeholders if p.placeholder_format.type == PP_PLACEHOLDER.CHART]
    texts = [p for p in placeholders if _is_text(p)]
    return bool(texts) and len(charts) <= 1 and len(texts) + len(charts) == len(placeholders)


def _is_text(ph) -> bool:
    vertical = ph.element.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}").get("orient")
    return ph.placeholder_format.type in TEXT_TYPES and vertical != "vert"


def _from_dict(data: dict) -> Catalog:
    layouts = [
        Layout(item["id"], item["name"], [Slot(**slot) for slot in item["slots"]])
        for item in data["layouts"]
    ]
    return Catalog(data["template"], data["key"], layouts)


# ---------------------------------------------------------------- measuring a placeholder


def _measure(ph, master, fonts: dict[str, str]) -> Slot:
    """Work out how many lines and characters fit in a placeholder at its template font."""
    if ph.placeholder_format.type == PP_PLACEHOLDER.CHART:  # filled with a chart, not text
        width, height = round(Emu(ph.width).pt, 1), round(Emu(ph.height).pt, 1)
        return Slot("chart", ph.placeholder_format.idx, "chart", 0, 0, 0.0,
                    width_pt=width, height_pt=height)  # fmt: skip
    style = _style_sources(ph, master)
    size = int(_find(style, "a:defRPr", "sz") or 1800) / 100
    bold = _find(style, "a:defRPr", "b") == "1"
    caps = _find(style, "a:defRPr", "cap") == "all"
    is_title = ph.placeholder_format.type in TITLE_TYPES
    font = _find(style, "a:defRPr/a:latin", "typeface") or ("+mj-lt" if is_title else "+mn-lt")
    font = fonts.get(font, font)
    indent = Emu(int(_find(style, None, "marL") or 0)).pt
    spacing = int(_find(style, "a:lnSpc/a:spcPct", "val") or 100000) / 100000
    gap = int(_find(style, "a:spcBef/a:spcPts", "val") or 0) / 100  # space before each paragraph
    left, right, top, bottom = _insets(ph)
    width = round(Emu(ph.width).pt - left - right - indent, 1)
    height = round(Emu(ph.height).pt - top - bottom, 1)
    hint = _hint(ph)
    slot = Slot(
        "", ph.placeholder_format.idx, "text", 0, 0, size, hint.lower().startswith("optional"),
        hint, width, height, spacing, gap, font, bold, caps,
    )  # fmt: skip
    slot.chars = int(width / slot.char_width(size))
    slot.lines = max(1, int((height + gap) / (slot.line_height(size) + gap)))
    return slot


def _hint(ph) -> str:
    """The layout's own prompt text ("Optional Subheadline"), unless it is boilerplate."""
    text = (
        re.sub(r"\s*[\n\v\x0b]+\s*", ", ", ph.text_frame.text).strip(" ,")
        if ph.has_text_frame
        else ""
    )
    return "" if any(word in text.lower() for word in GENERIC_PROMPTS) else text[:60]


def _style_sources(ph, master) -> list:
    """Level-1 paragraph styles that apply to a placeholder, most specific first.

    PowerPoint inherits text style: layout placeholder, then the matching master placeholder,
    then the master's title or body text style.
    """
    is_title = ph.placeholder_format.type in TITLE_TYPES
    wanted = PP_PLACEHOLDER.TITLE if is_title else PP_PLACEHOLDER.BODY
    master_ph = next((p for p in master.placeholders if p.placeholder_format.type == wanted), None)
    text_style = master.element.find(
        f"{qn('p:txStyles')}/{qn('p:titleStyle' if is_title else 'p:bodyStyle')}"
    )
    sources = [_level1(ph), _level1(master_ph) if master_ph is not None else None]
    sources.append(text_style.find(qn("a:lvl1pPr")) if text_style is not None else None)
    return [s for s in sources if s is not None]


def _level1(ph):
    body = ph.element.txBody
    return None if body is None else body.find(f"{qn('a:lstStyle')}/{qn('a:lvl1pPr')}")


def _find(sources: list, path: str | None, attr: str) -> str | None:
    """First value of an attribute along the inheritance chain."""
    for source in sources:
        element = source if path is None else source.find("/".join(qn(p) for p in path.split("/")))
        if element is not None and element.get(attr) is not None:
            return element.get(attr)
    return None


def _insets(ph) -> tuple[float, float, float, float]:
    """Inner text margins in points (PowerPoint default: 0.1 in sides, 0.05 in top/bottom)."""
    body_pr = ph.element.txBody.find(qn("a:bodyPr")) if ph.element.txBody is not None else None
    defaults = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
    values = {k: int(body_pr.get(k, v)) if body_pr is not None else v for k, v in defaults.items()}
    return tuple(Emu(values[k]).pt for k in ("lIns", "rIns", "tIns", "bIns"))


def _name_slots(slots: list[Slot], placeholders: list) -> list[Slot]:
    """Give slots stable, readable names: title, subtitle, then heads and bodies left to right.

    A body box with room for one line only is a heading ("head"); larger ones hold bullets.
    On a layout without bullet boxes, a one-line box is the subtitle. Subtitles are optional.
    """
    order = sorted(zip(slots, placeholders, strict=True), key=lambda sp: (sp[1].left, sp[1].top))
    has_bullets = any(
        slot.lines > 1 and ph.placeholder_format.type not in TITLE_TYPES for slot, ph in order
    )
    counts: dict[str, int] = {}
    for slot, ph in order:
        kind = ph.placeholder_format.type
        if kind in TITLE_TYPES:
            base = "title"
        elif kind == PP_PLACEHOLDER.CHART:
            base = "chart"
        elif kind == PP_PLACEHOLDER.SUBTITLE or (slot.lines == 1 and not has_bullets):
            base, slot.optional = "subtitle", True
        elif slot.lines == 1:
            base = "head"
        else:
            base, slot.kind = "body", "list"
        counts[base] = counts.get(base, 0) + 1
        slot.name = base if counts[base] == 1 else f"{base}{counts[base]}"
    rank = {"title": 0, "subtitle": 1}
    return [s for s, _ in sorted(order, key=lambda sp: rank.get(sp[0].name, 2))]


# ---------------------------------------------------------------- choosing the team template

# Words that signal each team. Matched as whole words, case-insensitive. English, French, German.
TEAM_WORDS = {
    "finance": [
        "budget", "cost", "costs", "revenue", "margin", "profit", "forecast", "invoice", "price",
        "roi", "payback", "ebit", "cash", "capex", "opex", "savings", "eur", "usd", "tnd",
        "coût", "coûts", "chiffre", "marge", "bénéfice", "kosten", "umsatz", "gewinn",
    ],
    "engineering": [
        "design", "prototype", "test", "tests", "spec", "specification", "firmware", "sensor",
        "atex", "iecex", "voltage", "module", "architecture", "api", "bug", "release",
        "hardware", "software", "integration", "certification", "enclosure", "pcb",
        "conception", "essai", "logiciel", "prüfung", "entwicklung",
    ],
    "marketing": [
        "campaign", "brand", "launch", "customer", "customers", "audience", "social", "event",
        "fair", "leads", "newsletter", "website", "press", "seo", "linkedin", "webinar",
        "campagne", "marque", "lancement", "salon", "kampagne", "marke", "messe",
    ],
}  # fmt: skip
MONEY_OR_PERCENT = re.compile(r"[€$£%]|\d+(?:[.,]\d+)?\s?(?:k|m|bn)\b", re.IGNORECASE)
MIN_SCORE = 2  # fewer signals than this: stay on the neutral base template


def pick_team(notes: str) -> str:
    """Pick the team whose vocabulary the notes use most; the base template when unclear.

    Simple and deterministic on purpose: no tokens, testable, explainable. A wrong pick is
    harmless because every team template has the same layouts; only the colours change.
    """
    words = re.findall(r"[\w']+", notes.lower())
    scores = {team: sum(words.count(w) for w in vocab) for team, vocab in TEAM_WORDS.items()}
    scores["finance"] += len(MONEY_OR_PERCENT.findall(notes))  # figures-heavy notes -> finance
    best = max(scores, key=scores.get)
    runner_up = max(score for team, score in scores.items() if team != best)
    if scores[best] < MIN_SCORE or scores[best] == runner_up:
        return DEFAULT_TEAM
    return best


def resolve_template(notes: str, team: str | None = None, template: Path | None = None) -> Path:
    """Which template to use, in order of priority:

    1. an explicit template file (--template);
    2. a team (--team finance), or "auto" to pick the team from the notes;
    3. otherwise the company Master.
    """
    if template is not None:
        return template
    if team is None:
        return MASTER
    if team == "auto":
        team = pick_team(notes)
    if team not in TEAMS:
        raise ValueError(f"Unknown team {team!r}. Choose one of: auto, {', '.join(TEAMS)}")
    return MASTER if team == "base" else TEMPLATE_DIR / f"{team}.pptx"
