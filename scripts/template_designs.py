"""The four team looks, on exactly the text grid of the company Master template.

Every text box sits where the Master puts it (GRID below, read from Master.pptx), with the
Master's font sizes. "base" looks like the Master itself. The other teams change only
colours and add decorations in the free areas around the text, never on top of it:
the strip above the title, the side margins and the footer zone.

- base         like the Master: white slides, navy uppercase titles, green-to-teal cover
- marketing    teal, green and periwinkle circles in the corners
- engineering  dark slides with a blueprint grid and corner brackets
- finance      navy side rail and green bar-chart motifs
"""

from collections.abc import Callable
from dataclasses import dataclass

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.shapes.autoshape import Shape
from pptx.util import Inches, Pt

W, H = 13.333, 7.5  # slide size in inches (16:9), as in the Master
TOP, MIDDLE, BOTTOM = MSO_ANCHOR.TOP, MSO_ANCHOR.MIDDLE, MSO_ANCHOR.BOTTOM

# Text boxes of the Master (x, y, width, height in inches) and their font sizes in points.
GRID = {
    "title": ((0.72, 0.78, 11.89, 0.47), 31),
    "subheadline": ((0.72, 1.27, 11.89, 0.29), 17),
    "body": ((0.72, 1.98, 9.96, 4.62), 20),
    "source": ((0.72, 6.75, 9.96, 0.15), 10),
    "cover_title": ((0.72, 5.86, 9.96, 0.47), 31),
    "cover_subtitle": ((0.72, 6.33, 9.96, 0.57), 17),
    "separation_title": ((0.72, 1.92, 9.94, 1.82), 60),
    "separation_subtitle": ((0.72, 3.74, 9.94, 0.34), 20),
    "end_title": ((0.72, 1.92, 9.94, 1.07), 60),
    "end_subtitle": ((0.72, 3.01, 9.94, 0.52), 31),
}
COLUMNS_2 = [0.72, 6.90]  # left edges of the columns; each 5.71 in wide
COLUMNS_3 = [0.72, 4.84, 8.96]  # each 3.65 in wide
CHART_BOX = (6.90, 2.47, 5.71, 4.12)

# Colours of the company Master: theme navy and blues, plus the logo green and the cover
# gradient (sampled from the rendered title slide).
NAVY, PERIWINKLE, PALE = "003282", "6673B8", "BAC6E7"
DEEP_GREEN, GREEN, TEAL, SKY, MIST = "00825F", "00966D", "2EA19C", "55A2B6", "BDE0E2"
GREY_DARK, GREY, GREY_LIGHT = "505050", "9A9A9A", "D5D5D5"
TEXT, WHITE = "000000", "FFFFFF"


@dataclass
class Design:
    accents: list[str]  # theme accent1..6: accent1 colours bullets, charts use them in order
    bullet: str | None  # None: plain paragraphs without bullet marks, as in the Master
    title_color: str  # content slide titles
    sub_color: str  # subheadlines
    master: Callable  # draws decoration on the slide master (all content slides)
    cover: Callable  # decoration of the title slide
    separation: Callable  # decoration of the separation (key message) slide
    end: Callable  # decoration of the end slide
    cover_text: tuple[str, str] = (NAVY, NAVY)  # title and subtitle colours
    separation_text: tuple[str, str] = (NAVY, NAVY)
    end_text: tuple[str, str] = (WHITE, WHITE)
    dark: bool = False  # dark slides: theme text colours are swapped (see build_templates.py)


# ---------------------------------------------------------------- drawing helpers


def place(ph, box, anchor=TOP, size=None, bold=False, color=None, bullets=True, caps=False):
    """Set a placeholder's box and text behaviour; optionally override its first-level style."""
    ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in box)
    if ph.element.txBody is None:
        return  # chart and table placeholders hold no text
    frame = ph.text_frame
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE  # our code fits text itself; no PowerPoint autofit
    frame.vertical_anchor = anchor
    lst_style = ph.element.txBody.find(qn("a:lstStyle"))
    for child in list(lst_style):
        lst_style.remove(child)  # drop the default template's per-layout overrides
    if size or bold or color or not bullets:
        indent = "" if bullets else ' marL="0" indent="0"'
        bullet = "" if bullets else "<a:buNone/>"
        run = (f' sz="{size * 100}"' if size else "") + (' b="1"' if bold else "")
        run += ' cap="all"' if caps else ""
        fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ""
        level = f"<a:lvl1pPr {nsdecls('a')}{indent}>{bullet}<a:defRPr{run}>{fill}</a:defRPr>"
        lst_style.append(parse_xml(level + "</a:lvl1pPr>"))


# python-pptx can only add shapes to slides, so for masters and layouts we create the
# shape element with python-pptx's own constructor and append it to the shape tree.
def _add(owner, sp) -> Shape:
    owner.element.cSld.spTree.append(sp)
    return Shape(sp, owner.shapes)


def _next_id(owner) -> int:
    return max(int(e.get("id")) for e in owner.element.iter(qn("p:cNvPr"))) + 1


def shape(owner, box, color: str, kind: str = "rect", rotation: float = 0) -> None:
    """A filled rectangle, ellipse or triangle with no outline and no shadow."""
    x, y, w, h = (Inches(v) for v in box)
    sp = CT_Shape.new_autoshape_sp(_next_id(owner), "Decoration", kind, x, y, w, h)
    sp.remove(sp.find(qn("p:style")))  # the default shape style adds a drop shadow
    deco = _add(owner, sp)
    deco.fill.solid()
    deco.fill.fore_color.rgb = RGBColor.from_string(color)
    deco.line.fill.background()
    deco.rotation = rotation


def text(owner, box, value: str, color: str, size: int, bold=False, align=PP_ALIGN.LEFT) -> None:
    """A fixed text box that is part of the design (never filled by the generator)."""
    x, y, w, h = (Inches(v) for v in box)
    frame = _add(owner, CT_Shape.new_textbox_sp(_next_id(owner), "Label", x, y, w, h)).text_frame
    frame.paragraphs[0].alignment = align
    run = frame.paragraphs[0].add_run()
    run.text = value
    run.font.size, run.font.bold = Pt(size), bold
    run.font.color.rgb = RGBColor.from_string(color)


def footer(owner, color: str, mark: str) -> None:
    """As in the Master: "© R. STAHL" bottom left, a mark bottom right, then the slide number.

    The mark is a bold text wordmark standing in for the logo, which we never copy.
    """
    text(owner, (0.72, 6.97, 3, 0.3), "© R. STAHL", color, 9)
    text(owner, (10.4, 6.9, 2.1, 0.4), "R. STAHL", mark, 14, bold=True, align=PP_ALIGN.RIGHT)
    box = (Inches(v) for v in (12.55, 6.97, 0.5, 0.3))
    frame = _add(owner, CT_Shape.new_textbox_sp(_next_id(owner), "Slide Number", *box)).text_frame
    old = frame.paragraphs[0]._p
    old.getparent().replace(old, parse_xml(  # a slide-number field: PowerPoint fills in the number
        f'<a:p {nsdecls("a")}><a:pPr algn="r"/>'
        '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
        f'<a:rPr lang="en-US" sz="900"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        "</a:rPr><a:t>‹#›</a:t></a:fld></a:p>"
    ))  # fmt: skip


def background(owner, color: str) -> None:
    owner.background.fill.solid()
    owner.background.fill.fore_color.rgb = RGBColor.from_string(color)


def hide_master_shapes(layout) -> None:
    layout.element.set("showMasterSp", "0")


# ---------------------------------------------------------------- base: like the Master

GRADIENT = [DEEP_GREEN, "018977", GREEN, TEAL, "2E9BA8", SKY, "8FC7D6", MIST, "DCEFF2"]


def triangle_mosaic(owner, columns: int, rows: int) -> None:
    """Rows of up and down triangles, green on the left fading to white: the Master's cover."""
    size = 0.9
    for row in range(rows):
        for col in range(columns):
            fade = col / (columns - 1)
            if fade > 0.55 and (row + col) % 3 == 0:
                continue  # gaps on the right so the mosaic dissolves into white
            color = GRADIENT[min(len(GRADIENT) - 1, int(fade * len(GRADIENT)) + row % 2)]
            flipped = (row + col) % 2 == 1
            box = (col * size / 2 - size / 2, row * size * 0.87, size, size * 0.87)
            shape(owner, box, color, "triangle", 180 if flipped else 0)


def base_master(m) -> None:
    footer(m, GREY, NAVY)


def base_cover(layout) -> None:
    hide_master_shapes(layout)
    triangle_mosaic(layout, columns=17, rows=6)  # stops above the title area, as in the Master
    text(layout, (8.8, 2.4, 4.0, 1.0), "R. STAHL", NAVY, 44, bold=True, align=PP_ALIGN.RIGHT)


def base_separation(layout) -> None:
    shape(layout, (0.72, 1.6, 1.2, 0.08), GREEN)  # small accent above the title


def base_end(layout) -> None:
    hide_master_shapes(layout)
    triangle_mosaic(layout, columns=23, rows=9)
    text(layout, (8.8, 5.9, 4.0, 1.0), "R. STAHL", NAVY, 40, bold=True, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- marketing: circles

DARK_TEAL, TEAL_TINT = "17807C", "E3F3F2"


def marketing_master(m) -> None:
    shape(m, (11.6, -1.75, 2.4, 2.4), MIST, "ellipse")  # top right, above the title
    shape(m, (12.75, 0.75, 0.4, 0.4), GREEN, "ellipse")  # right margin, beside the title
    shape(m, (12.85, 1.35, 0.2, 0.2), NAVY, "ellipse")
    footer(m, GREY, NAVY)


def marketing_cover(layout) -> None:
    hide_master_shapes(layout)
    background(layout, NAVY)
    shape(layout, (6.6, -2.2, 7.6, 7.6), TEAL, "ellipse")
    shape(layout, (5.4, 2.6, 2.6, 2.6), GREEN, "ellipse")
    shape(layout, (11.0, 3.9, 1.5, 1.5), PERIWINKLE, "ellipse")
    shape(layout, (0.72, 5.5, 1.2, 0.09), TEAL)


def marketing_separation(layout) -> None:
    background(layout, TEAL_TINT)
    shape(layout, (11.0, 3.9, 3.0, 3.0), TEAL, "ellipse")  # clear of text and footer
    shape(layout, (10.6, 4.2, 0.7, 0.7), GREEN, "ellipse")


def marketing_end(layout) -> None:
    hide_master_shapes(layout)
    background(layout, NAVY)
    shape(layout, (9.4, 3.8, 5.2, 5.2), TEAL, "ellipse")
    shape(layout, (8.6, 3.9, 1.6, 1.6), GREEN, "ellipse")
    shape(layout, (11.6, -0.9, 2.2, 2.2), PERIWINKLE, "ellipse")


# ---------------------------------------------------------------- engineering: dark, technical

DEEP, PANEL, GRIDLINE, LILAC = "001A4A", "0A2C66", "0C2A5E", "A3AEEA"


def blueprint(owner) -> None:
    for i in range(1, 27):  # blueprint grid, one line every half inch
        shape(owner, (i * 0.5, 0, 0.012, H), GRIDLINE)
    for i in range(1, 15):
        shape(owner, (0, i * 0.5, W, 0.012), GRIDLINE)


def engineering_master(m) -> None:
    background(m, DEEP)
    blueprint(m)
    shape(m, (0.3, 0.3, 0.5, 0.04), LILAC)  # corner bracket in the top-left margin
    shape(m, (0.3, 0.3, 0.04, 0.5), LILAC)
    text(m, (9.6, 0.25, 3.0, 0.35), "ENGINEERING", LILAC, 10, bold=True, align=PP_ALIGN.RIGHT)
    footer(m, PALE, WHITE)


def engineering_cover(layout) -> None:
    hide_master_shapes(layout)  # the footer would collide with the presenter line
    blueprint(layout)
    shape(layout, (0.3, 0.3, 0.5, 0.04), LILAC)
    shape(layout, (0.3, 0.3, 0.04, 0.5), LILAC)
    shape(layout, (0.72, 5.55, 1.2, 0.08), LILAC)
    for i in range(6):  # progress-style squares, green like the Master's logo
        shape(layout, (0.72 + i * 0.32, 4.9, 0.2, 0.2), GREEN if i < 4 else PANEL)


def engineering_separation(layout) -> None:
    hide_master_shapes(layout)
    background(layout, PANEL)
    shape(layout, (0, 0, 0.25, H), GREEN)
    shape(layout, (0.72, 1.6, 1.2, 0.08), LILAC)


def engineering_end(layout) -> None:
    shape(layout, (0.72, 1.6, 1.2, 0.08), GREEN)


# ---------------------------------------------------------------- finance: rail and bars

LIGHT_GREEN, PALE_GREEN, GREEN_TINT = "4DB88E", "A8DCC6", "EEF6F2"


def bars(owner, x, baseline, width, gap, heights, colors) -> None:
    for i, (height, color) in enumerate(zip(heights, colors, strict=True)):
        shape(owner, (x + i * (width + gap), baseline - height, width, height), color)


def finance_master(m) -> None:
    shape(m, (0, 0, 0.3, H), NAVY)  # side rail in the left margin
    bars(
        m, 11.6, 0.68, 0.16, 0.09, [0.2, 0.32, 0.44, 0.56], [PALE_GREEN, LIGHT_GREEN, GREEN, NAVY]
    )  # fmt: skip  (above the title)
    footer(m, GREY, NAVY)


def finance_cover(layout) -> None:
    hide_master_shapes(layout)
    shape(layout, (0, 0, 0.3, H), NAVY)
    bars(layout, 6.6, 5.4, 0.8, 0.28, [1.3, 2.0, 1.7, 2.9, 3.8, 4.7],
         [PALE_GREEN, LIGHT_GREEN, PALE_GREEN, LIGHT_GREEN, GREEN, NAVY])  # fmt: skip
    shape(layout, (6.4, 5.4, 6.5, 0.03), GREY_LIGHT)  # chart baseline
    shape(layout, (0.72, 5.55, 1.2, 0.08), GREEN)


def finance_separation(layout) -> None:
    background(layout, GREEN_TINT)
    bars(layout, 11.1, 6.6, 0.3, 0.12, [0.8, 1.3, 1.9, 2.6],
         [PALE_GREEN, LIGHT_GREEN, GREEN, NAVY])  # fmt: skip


def finance_end(layout) -> None:
    hide_master_shapes(layout)
    background(layout, NAVY)
    bars(
        layout,
        10.85,
        6.6,
        0.45,
        0.15,
        [1.2, 1.9, 2.7, 3.4],
        [PALE_GREEN, LIGHT_GREEN, GREEN, WHITE],
    )  # fmt: skip  (right of the text)


DESIGNS = {
    "base": Design(
        accents=[NAVY, PERIWINKLE, PALE, GREY_DARK, GREY, GREY_LIGHT], bullet=None,
        title_color=NAVY, sub_color=NAVY,
        master=base_master, cover=base_cover, separation=base_separation, end=base_end,
    ),
    "marketing": Design(
        accents=[TEAL, NAVY, GREEN, PERIWINKLE, SKY, PALE], bullet="●",
        title_color=NAVY, sub_color=DARK_TEAL,
        master=marketing_master, cover=marketing_cover, separation=marketing_separation,
        end=marketing_end, cover_text=(WHITE, WHITE), separation_text=(NAVY, DARK_TEAL),
    ),
    "engineering": Design(
        accents=[LILAC, GREEN, SKY, PERIWINKLE, PALE, WHITE], bullet="►",
        title_color=WHITE, sub_color=LILAC, dark=True,
        master=engineering_master, cover=engineering_cover, separation=engineering_separation,
        end=engineering_end, cover_text=(WHITE, PALE), separation_text=(WHITE, PALE),
        end_text=(WHITE, PALE),
    ),
    "finance": Design(
        accents=[DEEP_GREEN, NAVY, LIGHT_GREEN, SKY, PALE_GREEN, GREY], bullet="▪",
        title_color=NAVY, sub_color=DEEP_GREEN,
        master=finance_master, cover=finance_cover, separation=finance_separation,
        end=finance_end, separation_text=(NAVY, DEEP_GREEN),
    ),
}  # fmt: skip
