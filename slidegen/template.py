"""Read a .pptx template into a short layout catalog, and pick the team template for some notes.

The model never sees template XML. It sees a catalog of about 100 tokens, for example:

    L4 Comparison: title<=56, head<=38, body<=8x38, head2<=38, body2<=8x38

Each budget comes from the placeholder's box size and its font size in the template:
"8x38" means 8 bullets of up to 38 characters; "56" means up to 56 characters in total.
The catalog is cached as JSON next to the template and rebuilt when the file changes.
"""

import hashlib
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Emu

TEMPLATE_DIR = Path(__file__).parent / "templates"
DEFAULT_TEAM = "base"
CATALOG_VERSION = 1  # bump when the budget rules change, so old caches are rebuilt

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
TEXT_TYPES = TITLE_TYPES | {PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT}
IGNORED_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}

AVG_CHAR_EM = 0.5  # average Arial-like character width, as a fraction of the font size
BOLD_CHAR_EM = 0.55
LINE_HEIGHT = 1.2  # line height as a multiple of the font size at 100% line spacing


@dataclass
class Slot:
    name: str  # what the model fills: title, subtitle, head, body, head2, body2 ...
    idx: int  # placeholder index on the layout
    kind: str  # "text" (one string) or "list" (bullets)
    lines: int
    chars: int  # characters per line
    font_pt: float  # font size in the template

    @property
    def budget(self) -> int:
        return self.lines * self.chars

    def describe(self) -> str:
        limit = f"{self.lines}x{self.chars}" if self.kind == "list" else str(self.budget)
        return f"{self.name}<={limit}"


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
    layouts = []
    for index, layout in enumerate(prs.slide_layouts):
        placeholders = [
            p for p in layout.placeholders if p.placeholder_format.type not in IGNORED_TYPES
        ]
        if not placeholders or not all(_is_text(p) for p in placeholders):
            continue  # skip layouts with pictures, charts, tables, or no text at all
        slots = _name_slots([_measure(p, prs.slide_master) for p in placeholders], placeholders)
        layouts.append(Layout(f"L{index}", layout.name, slots))
    return Catalog(template.name, key, layouts)


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


def _measure(ph, master) -> Slot:
    """Work out how many lines and characters fit in a placeholder at its template font size."""
    style = _style_sources(ph, master)
    size = int(_find(style, "a:defRPr", "sz") or 1800) / 100
    bold = _find(style, "a:defRPr", "b") == "1"
    indent = Emu(int(_find(style, None, "marL") or 0)).pt
    spacing = int(_find(style, "a:lnSpc/a:spcPct", "val") or 100000) / 100000
    gap = int(_find(style, "a:spcBef/a:spcPts", "val") or 0) / 100  # space before each paragraph
    left, right, top, bottom = _insets(ph, master)
    width = Emu(ph.width).pt - left - right - indent
    height = Emu(ph.height).pt - top - bottom
    chars = int(width / (size * (BOLD_CHAR_EM if bold else AVG_CHAR_EM)))
    lines = max(1, int((height + gap) / (size * LINE_HEIGHT * spacing + gap)))
    return Slot("", ph.placeholder_format.idx, "text", lines, chars, size)


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


def _insets(ph, master) -> tuple[float, float, float, float]:
    """Inner text margins in points (PowerPoint default: 0.1 in sides, 0.05 in top/bottom)."""
    body_pr = ph.element.txBody.find(qn("a:bodyPr")) if ph.element.txBody is not None else None
    defaults = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
    values = {k: int(body_pr.get(k, v)) if body_pr is not None else v for k, v in defaults.items()}
    return tuple(Emu(values[k]).pt for k in ("lIns", "rIns", "tIns", "bIns"))


def _name_slots(slots: list[Slot], placeholders: list) -> list[Slot]:
    """Give slots stable, readable names: title, subtitle, then heads and bodies left to right.

    A body box with room for one line only is a heading ("head"); larger ones hold bullets.
    """
    order = sorted(zip(slots, placeholders, strict=True), key=lambda sp: (sp[1].left, sp[1].top))
    counts: dict[str, int] = {}
    for slot, ph in order:
        kind = ph.placeholder_format.type
        if kind in TITLE_TYPES:
            base = "title"
        elif kind == PP_PLACEHOLDER.SUBTITLE:
            base = "subtitle"
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
    """An explicit template file wins, then an explicit team, then the content-based pick."""
    if template is not None:
        return template
    path = TEMPLATE_DIR / f"{team or pick_team(notes)}.pptx"
    if not path.exists():
        teams = sorted(p.stem for p in TEMPLATE_DIR.glob("*.pptx"))
        raise ValueError(f"Unknown team {team!r}. Choose one of: {', '.join(teams)}")
    return path
