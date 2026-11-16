"""Render .pptx files to PNG images with PowerPoint, to check how slides really look.

This is a development check only (Windows + PowerPoint); the slidegen package never needs it.
For each deck it writes one PNG per slide and an overview sheet with all slides.

Run:  uv run python scripts/render.py out/showcase/*.pptx
Writes out/render/<deck name>/slide1.png ... and out/render/<deck name>/overview.png
"""

import sys
from pathlib import Path

import win32com.client
from PIL import Image

OUT_DIR = Path(__file__).resolve().parent.parent / "out" / "render"
WIDTH, HEIGHT = 1600, 900  # 16:9


def render(pptx: Path, app) -> Path:
    target = OUT_DIR / pptx.stem
    target.mkdir(parents=True, exist_ok=True)
    deck = app.Presentations.Open(str(pptx.resolve()), ReadOnly=True, WithWindow=False)
    try:
        for i, slide in enumerate(deck.Slides, start=1):
            slide.Export(str(target / f"slide{i}.png"), "PNG", WIDTH, HEIGHT)
    finally:
        deck.Close()
    make_overview(sorted(target.glob("slide*.png")), target / "overview.png")
    return target


def make_overview(images: list[Path], path: Path, columns: int = 2) -> None:
    """Tile the slide images into one picture, half size, with a thin grey gap."""
    w, h, gap = WIDTH // 2, HEIGHT // 2, 12
    rows = (len(images) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * (w + gap) + gap, rows * (h + gap) + gap), "#9a9a9a")
    for i, image in enumerate(images):
        tile = Image.open(image).convert("RGB").resize((w, h))
        sheet.paste(tile, (gap + (i % columns) * (w + gap), gap + (i // columns) * (h + gap)))
    sheet.save(path)


def main(paths: list[str]) -> None:
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for path in paths:
            print(render(Path(path), app))
    finally:
        if app.Presentations.Count == 0:  # never close a PowerPoint the user has open
            app.Quit()


if __name__ == "__main__":
    main(sys.argv[1:])
