"""Build the team templates from the company Master template.

Each team template is a copy of slidegen/templates/master.pptx with its own colours and style
on the content slides:

1. a team colour as the main theme colour (accent 1). The Master uses accent 1 for its titles,
   subheadlines, chart series and table headers, so all of them follow the team;
2. bullet marks on the content boxes (the Master's body text has none);
3. a background tint for the content slides and the separation slide;
4. a few decorative shapes in the free margins, drawn behind the Master's own footer, date,
   slide number and logo.

The Master's first and last slides, its layouts, every text box and font, the optional
subheadline and source line, the footer, date and slide number stay exactly as they are.
Because the layouts and placeholders are unchanged, the generator works the same with every
team. The base template is the Master itself and is not built here.

Run:  uv run python scripts/build_templates.py
Writes slidegen/templates/<team>.pptx
"""

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.shapes.autoshape import Shape
from pptx.util import Inches, Pt

from slidegen.deck import remove_sample_slides

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
MASTER = TEMPLATE_DIR / "master.pptx"
W, H = 13.333, 7.5  # slide size in inches

# Colours of the Master: navy from its theme, green from its logo.
NAVY, GREEN = "003282", "00966D"


@dataclass
class Team:
    primary: str  # theme accent 1: titles, subheadlines, first chart series, table headers
    secondary: str  # theme accent 2: second chart series
    bullet: str
    bullet_color: str
    background: str  # content slides
    separation: str  # background of the separation slide
    cover: str  # colour of the Master's mosaic on the first and last slide
    decorate: Callable  # draws on the slide master
    dark: bool = False  # light text on dark content slides


def build_template(name: str) -> Path:
    team = TEAMS[name]
    prs = Presentation(MASTER)
    remove_sample_slides(prs)
    for layout in (prs.slide_layouts[0], prs.slide_layouts[-1]):  # first and last slide
        pin_theme_colors(layout, prs)
        recolor_mosaic(layout, team.cover)
    if team.dark:
        swap_text_colors(prs.slide_master)
    set_accents(prs, team.primary, team.secondary)
    set_bullets(prs, team.bullet, team.bullet_color)
    background(prs.slide_master, team.background)
    for layout in prs.slide_layouts:
        if layout.name == "Separation Slide":
            background(layout, team.separation)
    team.decorate(prs.slide_master)
    path = TEMPLATE_DIR / f"{name}.pptx"
    prs.save(path)
    return path


def theme_colors(prs: Presentation) -> dict[str, str]:
    """Theme colour name -> hex value, e.g. {"accent1": "003282", "tx1": "000000", ...}."""
    theme = etree.fromstring(prs.slide_master.part.part_related_by(RT.THEME).blob)
    scheme = theme.find(f".//{qn('a:clrScheme')}")
    colors = {  # system colours (e.g. "window") carry their hex value in lastClr
        etree.QName(slot).localname: slot[0].get("lastClr") or slot[0].get("val") for slot in scheme
    }
    color_map = prs.slide_master.element.find(qn("p:clrMap")).attrib  # tx1 -> dk1, bg1 -> lt1 ...
    return colors | {alias: colors[slot] for alias, slot in color_map.items() if slot in colors}


def pin_theme_colors(layout, prs: Presentation) -> None:
    """Replace theme colour references on a layout with the Master's actual colours.

    The Master's title and end slides use theme colours (for example accent 1 for the address
    on the end slide). Pinning them first means the team colours never change these slides.
    """
    colors = theme_colors(prs)
    for ref in list(layout.element.iter(qn("a:schemeClr"))):
        value = colors.get(ref.get("val"))
        if value:
            pinned = etree.SubElement(ref.getparent(), qn("a:srgbClr"), val=value)
            for child in ref:  # keep tints and shades
                pinned.append(child)
            ref.addprevious(pinned)
            ref.getparent().remove(ref)


def recolor_mosaic(layout, color: str) -> None:
    """Show the Master's full-slide mosaic picture in the team colour, fading to white.

    PowerPoint's duotone effect maps the picture's dark-to-light tones onto two colours, so
    the mosaic keeps its exact pattern. The logo and other pictures are left as they are.
    """
    for pic in layout.element.iter(qn("p:pic")):
        width = int(pic.find(f".//{qn('a:xfrm')}/{qn('a:ext')}").get("cx"))
        if width >= Inches(13):
            blip = pic.find(f".//{qn('a:blip')}")
            blip.insert(0, parse_xml(
                f'<a:duotone {nsdecls("a")}><a:srgbClr val="{color}"/>'
                '<a:srgbClr val="FFFFFF"/></a:duotone>'
            ))  # fmt: skip


def swap_text_colors(master) -> None:
    """Dark slides: theme text colours become light and backgrounds dark."""
    master.element.find(qn("p:clrMap")).attrib.update(
        {"bg1": "dk1", "tx1": "lt1", "bg2": "dk2", "tx2": "lt2"}
    )


def set_accents(prs: Presentation, primary: str, secondary: str) -> None:
    """Write the team colours into theme accents 1 and 2.

    The title and end slides use other theme colours (dark 2 and white), so they keep the
    Master's look.
    """
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    for slot_name, color in (("a:accent1", primary), ("a:accent2", secondary)):
        slot = theme.find(f".//{qn('a:clrScheme')}/{qn(slot_name)}")
        slot.clear()
        etree.SubElement(slot, qn("a:srgbClr"), val=color)
    # The theme is a plain binary part in python-pptx, so we replace its bytes.
    part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)


def background(owner, color: str) -> None:
    owner.background.fill.solid()
    owner.background.fill.fore_color.rgb = RGBColor.from_string(color)


def set_bullets(prs: Presentation, char: str, color: str) -> None:
    """Give the content boxes a team bullet; the Master's body text has none.

    Only the content boxes ("object" placeholders) get it. Subtitles, subheadlines and source
    lines are body placeholders and stay as they are.
    """
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.OBJECT:
                add_bullet(ph.element.txBody, char, color)


def add_bullet(tx_body, char: str, color: str) -> None:
    lst_style = tx_body.find(qn("a:lstStyle"))
    level = lst_style.find(qn("a:lvl1pPr"))
    if level is None:
        level = parse_xml(f"<a:lvl1pPr {nsdecls('a')}><a:defRPr/></a:lvl1pPr>")
        lst_style.insert(0, level)
    level.set("marL", "228600")  # 0.25 in: room for the bullet
    level.set("indent", "-228600")
    for tag in ("a:buClr", "a:buFont", "a:buNone", "a:buChar"):
        for old in level.findall(qn(tag)):
            level.remove(old)
    bullet = parse_xml(
        f'<a:x {nsdecls("a")}><a:buClr><a:srgbClr val="{color}"/></a:buClr>'
        f'<a:buFont typeface="Arial"/><a:buChar char="{char}"/></a:x>'
    )
    anchor = level.find(qn("a:defRPr"))  # bullet elements must come before defRPr
    for element in list(bullet):
        anchor.addprevious(element)


# ---------------------------------------------------------------- drawing helpers

# python-pptx can only add shapes to slides, so for the slide master we create the shape
# element with python-pptx's own constructor. We insert it at the back of the shape tree,
# so the Master's footer, date, slide number and logo stay on top.


def _add(master, sp) -> Shape:
    master.element.cSld.spTree.insert(2, sp)  # after the tree's two property elements
    return Shape(sp, master.shapes)


def _next_id(master) -> int:
    return max(int(e.get("id")) for e in master.element.iter(qn("p:cNvPr"))) + 1


def shape(master, box, color: str, kind: str = "rect") -> None:
    """A filled rectangle or ellipse with no outline and no shadow."""
    x, y, w, h = (Inches(v) for v in box)
    sp = CT_Shape.new_autoshape_sp(_next_id(master), "Decoration", kind, x, y, w, h)
    sp.remove(sp.find(qn("p:style")))  # the default shape style adds a drop shadow
    deco = _add(master, sp)
    deco.fill.solid()
    deco.fill.fore_color.rgb = RGBColor.from_string(color)
    deco.line.fill.background()


def label(master, box, value: str, color: str) -> None:
    x, y, w, h = (Inches(v) for v in box)
    frame = _add(master, CT_Shape.new_textbox_sp(_next_id(master), "Label", x, y, w, h)).text_frame
    frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    run = frame.paragraphs[0].add_run()
    run.text = value
    run.font.size, run.font.bold = Pt(10), True
    run.font.color.rgb = RGBColor.from_string(color)


# ---------------------------------------------------------------- the three teams
# Free areas of the Master's content slides: above the title (y < 0.7 in), between the
# subheadline and the body (y 1.6 to 1.9 in), the left margin (x < 0.65 in) and the right
# margin (x > 12.65 in). Text boxes never overlap these.

RASPBERRY, PINK, BLUSH, ORANGE = "C2185B", "F8D7E3", "FCE8EF", "EB7D00"
DEEP, PANEL, GRIDLINE, CYAN = "0A1F3D", "12345F", "16365F", "5CC8E0"
AMBER, LIGHT_AMBER, PALE_AMBER, PEACH = "B35900", "F5A54A", "FFD8A8", "FFF1E0"


def marketing(master) -> None:
    """Bold: a raspberry band across the top and circles in the corner."""
    shape(master, (12.78, 0.8, 0.36, 0.36), ORANGE, "ellipse")  # right margin
    shape(master, (12.88, 1.35, 0.16, 0.16), RASPBERRY, "ellipse")
    shape(master, (0, 0, W, 0.16), RASPBERRY)  # band across the top
    shape(master, (11.3, -1.9, 2.6, 2.6), PINK, "ellipse")  # top right, above the title


def engineering(master) -> None:
    """Technical: dark slides with a blueprint grid and a cyan corner bracket."""
    label(master, (9.6, 0.22, 3.0, 0.35), "ENGINEERING", CYAN)
    shape(master, (0.3, 0.3, 0.5, 0.05), CYAN)  # corner bracket, top-left margin
    shape(master, (0.3, 0.3, 0.05, 0.5), CYAN)
    for i in range(1, 15):  # blueprint grid behind everything
        shape(master, (0, i * 0.5, W, 0.012), GRIDLINE)
    for i in range(1, 27):
        shape(master, (i * 0.5, 0, 0.012, H), GRIDLINE)


def finance(master) -> None:
    """Report style: a navy side rail, an amber rule under the heading, a bar-chart motif."""
    shape(master, (0.72, 1.68, 9.96, 0.03), ORANGE)  # rule between subheadline and body
    heights, colors = [0.2, 0.32, 0.44, 0.56], [PALE_AMBER, LIGHT_AMBER, ORANGE, NAVY]
    for i, (height, color) in enumerate(zip(heights, colors, strict=True)):  # above the title
        shape(master, (11.6 + i * 0.25, 0.68 - height, 0.16, height), color)
    shape(master, (0, 0, 0.3, H), NAVY)  # side rail in the left margin


TEAMS = {
    "marketing": Team(RASPBERRY, ORANGE, "●", RASPBERRY, "FFFFFF", BLUSH, "9B0F4E", marketing),
    "engineering": Team(CYAN, GREEN, "►", CYAN, DEEP, PANEL, "0B3C8F", engineering, dark=True),
    "finance": Team(AMBER, NAVY, "▪", ORANGE, "FFFFFF", PEACH, "A34700", finance),
}


def main() -> None:
    for name in TEAMS:
        print(build_template(name))


if __name__ == "__main__":
    main()
