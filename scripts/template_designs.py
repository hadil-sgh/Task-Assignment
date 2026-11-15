"""The four team looks, plus the small drawing helpers they use.

Every team keeps the same eight layouts and placeholder indexes (see build_templates.py).
A design only decides colours, where the boxes sit, and the decorative shapes:

- base         corporate: navy header band, mosaic title slide, navy key message
- marketing    bold: raspberry, overlapping circles
- engineering  technical and dark: blueprint grid, teal brackets
- finance      light and data-driven: navy side rail, bar-chart motifs, green accents
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

W, H, M = 13.333, 7.5, 0.6  # slide width, height and side margin in inches (16:9)
TOP, BOTTOM = MSO_ANCHOR.TOP, MSO_ANCHOR.BOTTOM

# R. STAHL brand colours, read from the computed CSS of r-stahl.com.
NAVY, BLUE, MID_BLUE, LIGHT_BLUE = "072A5B", "13387E", "31549E", "D3E5F7"
ORANGE, GREY, TEXT, WHITE = "EB7D00", "D0D0D0", "2B2B2B", "FFFFFF"


@dataclass
class Design:
    accents: list[str]  # theme accent1..6: accent1 colours bullets, charts use them in order
    bullet: str
    title_color: str
    head_color: str  # comparison headings
    title_box: tuple[float, float, float, float]
    body_box: tuple[float, float, float, float]
    master: Callable  # draws content-slide decoration on the slide master
    title_slide: Callable  # positions placeholders and draws the title slide
    key_message: Callable  # same for the key message slide
    title_anchor: MSO_ANCHOR = BOTTOM
    dark: bool = False  # dark slides: theme text colours are swapped (see build_templates.py)


# ---------------------------------------------------------------- drawing helpers


def place(ph, box, anchor=TOP, size=None, bold=False, color=None, bullets=True) -> None:
    """Set a placeholder's box and text behaviour; optionally override its first-level style."""
    ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in box)
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


def shape(owner, box, color: str, kind: str = "rect") -> None:
    """A filled rectangle or ellipse with no outline and no shadow."""
    x, y, w, h = (Inches(v) for v in box)
    sp = CT_Shape.new_autoshape_sp(_next_id(owner), "Decoration", kind, x, y, w, h)
    sp.remove(sp.find(qn("p:style")))  # the default shape style adds a drop shadow
    deco = _add(owner, sp)
    deco.fill.solid()
    deco.fill.fore_color.rgb = RGBColor.from_string(color)
    deco.line.fill.background()


def text(owner, box, value: str, color: str, size: int, bold=False, align=PP_ALIGN.LEFT) -> None:
    """A fixed text box that is part of the design (never filled by the generator)."""
    x, y, w, h = (Inches(v) for v in box)
    frame = _add(owner, CT_Shape.new_textbox_sp(_next_id(owner), "Label", x, y, w, h)).text_frame
    frame.paragraphs[0].alignment = align
    run = frame.paragraphs[0].add_run()
    run.text = value
    run.font.size, run.font.bold = Pt(size), bold
    run.font.color.rgb = RGBColor.from_string(color)


def footer(owner, color: str, x: float = M) -> None:
    """Text wordmark (no logo) bottom left, slide number bottom right."""
    text(owner, (x, 6.9, 3, 0.4), "R. STAHL", color, 12, bold=True)
    box = (Inches(v) for v in (W - M - 1.5, 6.9, 1.5, 0.4))
    frame = _add(owner, CT_Shape.new_textbox_sp(_next_id(owner), "Slide Number", *box)).text_frame
    old = frame.paragraphs[0]._p
    old.getparent().replace(old, parse_xml(  # a slide-number field: PowerPoint fills in the number
        f'<a:p {nsdecls("a")}><a:pPr algn="r"/>'
        '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
        f'<a:rPr lang="en-US" sz="1100"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        "</a:rPr><a:t>‹#›</a:t></a:fld></a:p>"
    ))  # fmt: skip


def background(owner, color: str) -> None:
    owner.background.fill.solid()
    owner.background.fill.fore_color.rgb = RGBColor.from_string(color)


def hide_master_shapes(layout) -> None:
    layout.element.set("showMasterSp", "0")


# ---------------------------------------------------------------- base: corporate


def base_master(m) -> None:
    shape(m, (0, 0, W, 1.45), NAVY)  # header band; the title sits inside it in white
    shape(m, (0, 1.45, W, 0.06), ORANGE)
    footer(m, NAVY)


def base_title(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, NAVY)
    shape(layout, (7.9, 0, W - 7.9, H), LIGHT_BLUE)
    tiles = [BLUE, LIGHT_BLUE, MID_BLUE, MID_BLUE, ORANGE, NAVY,
             "7FA7D9", NAVY, LIGHT_BLUE, NAVY, MID_BLUE, "7FA7D9"]  # fmt: skip
    for i, color in enumerate(tiles):  # a 3 x 4 mosaic of brand colours
        shape(layout, (8.35 + (i % 3) * 1.6, 0.55 + (i // 3) * 1.65, 1.45, 1.5), color)
    shape(layout, (0.8, 4.45, 1.0, 0.08), ORANGE)
    place(ph[0], (0.8, 1.5, 6.6, 2.8), BOTTOM, size=40, bold=True, color=WHITE)
    place(ph[1], (0.8, 4.75, 6.6, 1.3), size=20, color=LIGHT_BLUE, bullets=False)
    text(layout, (0.8, 6.75, 3, 0.4), "R. STAHL", WHITE, 12, bold=True)


def base_key(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, NAVY)
    text(layout, (0.85, 0.55, 1.6, 2.0), "“", ORANGE, 140, bold=True)  # big quote mark
    place(ph[0], (1.4, 1.9, 10.6, 2.75), BOTTOM, size=36, bold=True, color=WHITE)
    shape(layout, (1.4, 4.8, 1.0, 0.06), ORANGE)
    place(ph[1], (1.4, 5.0, 10.6, 1.2), size=20, color=LIGHT_BLUE, bullets=False)
    footer(layout, WHITE)


# ---------------------------------------------------------------- marketing: bold

RASPBERRY, PINK = "C2185B", "F8D7E3"


def marketing_master(m) -> None:
    shape(m, (10.9, -1.9, 4.0, 4.0), PINK, "ellipse")
    shape(m, (12.3, 1.2, 0.9, 0.9), ORANGE, "ellipse")
    shape(m, (10.5, 1.45, 0.4, 0.4), RASPBERRY, "ellipse")
    shape(m, (M, 1.8, 1.1, 0.09), RASPBERRY)  # short underline below the title
    footer(m, NAVY)


def marketing_title(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, RASPBERRY)
    shape(layout, (8.2, 1.3, 6.6, 6.6), NAVY, "ellipse")
    shape(layout, (7.3, 4.9, 2.2, 2.2), ORANGE, "ellipse")
    shape(layout, (11.2, -0.9, 2.6, 2.6), PINK, "ellipse")
    place(ph[0], (0.8, 1.4, 7.0, 3.0), BOTTOM, size=44, bold=True, color=WHITE)
    place(ph[1], (0.8, 4.65, 6.3, 1.3), size=22, color=WHITE, bullets=False)
    text(layout, (0.8, 6.75, 3, 0.4), "R. STAHL", WHITE, 12, bold=True)


def marketing_key(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, PINK)
    shape(layout, (9.9, -1.6, 5.0, 5.0), RASPBERRY, "ellipse")
    shape(layout, (10.2, 3.6, 1.3, 1.3), ORANGE, "ellipse")
    place(ph[0], (0.9, 1.5, 8.7, 3.0), BOTTOM, size=38, bold=True, color=NAVY)
    place(ph[1], (0.9, 4.75, 8.7, 1.2), size=20, color=RASPBERRY, bullets=False)
    footer(layout, NAVY)


# ---------------------------------------------------------------- engineering: technical, dark

DEEP, PANEL, GRID, TEAL, ICE = "0A1F3D", "12345F", "15335C", "00B3C6", "7FDBE6"


def engineering_master(m) -> None:
    background(m, DEEP)
    for i in range(1, 27):  # blueprint grid, one line every half inch
        shape(m, (i * 0.5, 0, 0.012, H), GRID)
    for i in range(1, 15):
        shape(m, (0, i * 0.5, W, 0.012), GRID)
    shape(m, (0.45, 0.4, 0.6, 0.05), TEAL)  # corner bracket, top left
    shape(m, (0.45, 0.4, 0.05, 0.6), TEAL)
    text(m, (W - M - 3, 0.3, 3, 0.35), "ENGINEERING", TEAL, 10, bold=True, align=PP_ALIGN.RIGHT)
    footer(m, WHITE, x=0.9)


def engineering_title(layout, ph) -> None:
    shape(layout, (0.9, 1.6, 0.08, 3.8), TEAL)  # keeps the master grid and adds a frame
    for i in range(6):  # progress-style squares
        shape(layout, (1.3 + i * 0.32, 5.75, 0.2, 0.2), TEAL if i < 4 else PANEL)
    place(ph[0], (1.3, 1.5, 10.5, 2.6), BOTTOM, size=44, bold=True, color=WHITE)
    place(ph[1], (1.3, 4.3, 10.5, 1.1), size=22, color=ICE, bullets=False)


def engineering_key(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, PANEL)
    shape(layout, (0, 0, 0.25, H), TEAL)
    shape(layout, (1.2, 1.55, 1.2, 0.08), TEAL)
    place(ph[0], (1.2, 1.8, 10.9, 2.8), BOTTOM, size=38, bold=True, color=WHITE)
    place(ph[1], (1.2, 4.85, 10.9, 1.2), size=20, color=ICE, bullets=False)
    footer(layout, WHITE, x=1.2)


# ---------------------------------------------------------------- finance: light, data-driven

GREEN, LIGHT_GREEN, PALE_GREEN, MIST = "2E7D32", "66BB6A", "A5D6A7", "F3F6F4"


def bars(owner, x, baseline, width, gap, heights, colors) -> None:
    for i, (height, color) in enumerate(zip(heights, colors, strict=True)):
        shape(owner, (x + i * (width + gap), baseline - height, width, height), color)


def finance_master(m) -> None:
    shape(m, (0, 0, 0.3, H), NAVY)  # side rail
    bars(m, W - M - 0.92, 1.25, 0.16, 0.09, [0.25, 0.4, 0.55, 0.75],
         [PALE_GREEN, LIGHT_GREEN, GREEN, NAVY])  # fmt: skip
    shape(m, (0.9, 1.55, W - 0.9 - M, 0.015), GREY)
    shape(m, (0.9, 1.52, 1.6, 0.07), GREEN)
    footer(m, NAVY, x=0.9)


def finance_title(layout, ph) -> None:
    hide_master_shapes(layout)
    shape(layout, (0, 0, 0.3, H), NAVY)
    bars(layout, 6.9, 6.6, 0.75, 0.25, [1.4, 2.1, 1.8, 3.0, 3.9, 4.9],
         [PALE_GREEN, LIGHT_GREEN, PALE_GREEN, LIGHT_GREEN, GREEN, NAVY])  # fmt: skip
    shape(layout, (6.7, 6.6, 6.1, 0.03), GREY)  # chart baseline
    place(ph[0], (0.9, 1.5, 5.6, 2.9), BOTTOM, size=40, bold=True, color=NAVY)
    shape(layout, (0.9, 4.42, 1.0, 0.08), GREEN)
    place(ph[1], (0.9, 4.65, 5.4, 1.3), size=20, color="4A4A4A", bullets=False)
    text(layout, (0.9, 6.75, 3, 0.4), "R. STAHL", NAVY, 12, bold=True)


def finance_key(layout, ph) -> None:
    hide_master_shapes(layout)
    background(layout, MIST)
    shape(layout, (0, 0, 0.3, H), NAVY)
    shape(layout, (0.9, 1.7, 0.14, 3.9), GREEN)
    place(ph[0], (1.4, 1.6, 10.8, 2.9), BOTTOM, size=36, bold=True, color=NAVY)
    place(ph[1], (1.4, 4.75, 10.8, 1.2), size=20, color=GREEN, bullets=False)
    footer(layout, NAVY, x=0.9)


DESIGNS = {
    "base": Design(
        accents=[ORANGE, BLUE, MID_BLUE, "7FA7D9", GREY, NAVY], bullet="■",
        title_color=WHITE, head_color=NAVY, title_anchor=MSO_ANCHOR.MIDDLE,
        title_box=(M, 0.2, W - 2 * M, 1.05), body_box=(M, 1.9, W - 2 * M, 4.7),
        master=base_master, title_slide=base_title, key_message=base_key,
    ),
    "marketing": Design(
        accents=[RASPBERRY, ORANGE, NAVY, "E87AA0", MID_BLUE, "F2B366"], bullet="●",
        title_color=NAVY, head_color=RASPBERRY,
        title_box=(M, 0.45, 9.6, 1.25), body_box=(M, 2.2, W - 2 * M, 4.4),
        master=marketing_master, title_slide=marketing_title, key_message=marketing_key,
    ),
    "engineering": Design(
        accents=[TEAL, ICE, ORANGE, "6F8FCF", LIGHT_BLUE, WHITE], bullet="►",
        title_color=WHITE, head_color=ICE, dark=True,
        title_box=(0.9, 0.55, 11.0, 1.1), body_box=(0.9, 1.95, W - 0.9 - M, 4.6),
        master=engineering_master, title_slide=engineering_title, key_message=engineering_key,
    ),
    "finance": Design(
        accents=[GREEN, NAVY, LIGHT_GREEN, MID_BLUE, PALE_GREEN, GREY], bullet="▪",
        title_color=NAVY, head_color=GREEN,
        title_box=(0.9, 0.35, 10.3, 1.1), body_box=(0.9, 1.9, W - 0.9 - M, 4.7),
        master=finance_master, title_slide=finance_title, key_message=finance_key,
    ),
}  # fmt: skip
