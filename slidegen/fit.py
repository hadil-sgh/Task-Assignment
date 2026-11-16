"""Make text fit its box: measure it, shrink the font one point at a time, cut only at the end.

The size we choose is written explicitly on every run of the slide. We never rely on
PowerPoint's autofit, which only recalculates when someone edits the slide.

Text is measured with Pillow and the real font file (for the company Master: Arial Narrow).
When the font is not installed (for example on a Linux CI machine) we fall back to the
catalog's average character width, which is a little generous, so text still fits.
"""

import os
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

from slidegen.template import Slot

MIN_PT = 12  # never shrink below this
FLOOR = 0.6  # ... or below 60% of the template size
ELLIPSIS = "…"
SAFETY = 0.95  # wrap at 95% of the box width: PowerPoint lays text out slightly wider
FONT_DIRS = [
    Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
    Path.home() / "Library" / "Fonts",
    Path("/Library/Fonts"),
    Path("/usr/share/fonts"),
]
FONT_FILES = {  # font family -> (regular, bold) file names
    "arial narrow": ("arialn.ttf", "arialnb.ttf"),
    "arial": ("arial.ttf", "arialbd.ttf"),
}


@dataclass
class Fit:
    size: int  # font size in points
    paragraphs: list[str]  # the text, one string per paragraph (bullet)
    cut: bool = False  # True when words had to be removed at the smallest size


def fit_text(slot: Slot, value: str | list[str]) -> Fit:
    """The largest font size (template size or smaller) at which the text fits the box."""
    paragraphs = value if isinstance(value, list) else value.splitlines() or [value]
    size = round(slot.font_pt)
    floor = min(size, max(MIN_PT, round(slot.font_pt * FLOOR)))
    while not fits(slot, paragraphs, size):
        if size == floor:
            return Fit(size, cut_to_fit(slot, paragraphs, size), cut=True)
        size -= 1
    return Fit(size, paragraphs)


def fits(slot: Slot, paragraphs: list[str], size: float) -> bool:
    """Do the paragraphs, wrapped to the box width, fit the box height at this size?

    A box always holds one line: templates often draw one-line boxes (titles, subheadlines)
    a little lower than the line itself, and PowerPoint shows the line in full. The catalog
    uses the same rule.
    """
    lines = sum(len(wrap(slot, text, size)) for text in paragraphs)
    gaps = slot.gap_pt * (len(paragraphs) - 1)
    return lines == 1 or lines * slot.line_height(size) + gaps <= slot.height_pt


def wrap(slot: Slot, text: str, size: float) -> list[str]:
    """Greedy word wrap at the box width. A word longer than a line is split."""
    limit = slot.width_pt * SAFETY
    lines, line = [], ""
    for word in text.split():
        while width(slot, word, size) > limit:  # an overlong word: split it
            cut = max(1, int(len(word) * limit / width(slot, word, size)))
            if line:
                lines.append(line)
                line = ""
            lines.append(word[:cut])
            word = word[cut:]
        candidate = f"{line} {word}".strip()
        if line and width(slot, candidate, size) > limit:
            lines.append(line)
            line = word
        else:
            line = candidate
    return lines + [line] if line else lines or [""]


def width(slot: Slot, text: str, size: float) -> float:
    """Width of the text in points, with the real font when we have it."""
    text = text.upper() if slot.caps else text
    font = load_font(slot.font, slot.bold)
    if font is None:
        return len(text) * slot.char_width(size)
    return font.getlength(text) * size / 100  # the font is loaded at 100 pt


def cut_to_fit(slot: Slot, paragraphs: list[str], size: float) -> list[str]:
    """Last resort at the smallest size: drop words from the end and finish with an ellipsis."""
    words = [text.split() for text in paragraphs if text.split()]
    while words:
        candidate = [" ".join(w) for w in words]
        candidate[-1] += ELLIPSIS
        if fits(slot, candidate, size):
            return candidate
        words[-1].pop()  # remove the last word ...
        if not words[-1]:
            words.pop()  # ... and the paragraph once it is empty
    return [ELLIPSIS]


@lru_cache
def load_font(family: str, bold: bool) -> ImageFont.FreeTypeFont | None:
    names = FONT_FILES.get(family.lower())
    if not names:
        return None
    wanted = names[1 if bold else 0]
    for folder in (d for d in FONT_DIRS if d.is_dir()):
        for path in folder.rglob("*"):
            if path.name.lower() == wanted:
                return ImageFont.truetype(str(path), 100)
    return None
