"""Build the team templates from the company Master template.

Each team template is a copy of slidegen/templates/master.pptx with three small changes:

1. a team colour (theme accent 2, so it is also the second colour of native charts);
2. bullet marks in that colour (the Master's body text has none);
3. a few decorative shapes in the free margins of the content slides, drawn behind the
   Master's own footer, date, slide number and logo.

Everything else is the Master's: its first and last slides, its layouts, every text box and
font, the optional subheadline and source line, the footer, date and slide number. Because
the layouts and placeholders are unchanged, the generator works the same with every team.

Run:  uv run python scripts/build_templates.py
Writes slidegen/templates/<team>.pptx
"""

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

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
MASTER = TEMPLATE_DIR / "master.pptx"
W, H = 13.333, 7.5  # slide size in inches

# Colours of the Master: navy and blues from its theme, greens and teals from its cover.
NAVY, PERIWINKLE, PALE = "003282", "6673B8", "BAC6E7"
GREEN, TEAL, MIST = "00966D", "2EA19C", "BDE0E2"
LIGHT_GREEN, PALE_GREEN, GRID = "4DB88E", "A8DCC6", "E9EEF7"


def build_template(team: str) -> Path:
    accent, bullet, decorate = TEAMS[team]
    prs = Presentation(MASTER)
    remove_sample_slides(prs)
    set_accent(prs, accent)
    set_bullets(prs, bullet, accent)
    decorate(prs.slide_master)
    path = TEMPLATE_DIR / f"{team}.pptx"
    prs.save(path)
    return path


def remove_sample_slides(prs: Presentation) -> None:
    """A template holds layouts, not slides."""
    for slide_id in list(prs.slides._sldIdLst):
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)


def set_accent(prs: Presentation, color: str) -> None:
    """Theme accent 2 = team colour. Accent 1 stays the Master's navy, used by its titles."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    slot = theme.find(f".//{qn('a:clrScheme')}/{qn('a:accent2')}")
    slot.clear()
    etree.SubElement(slot, qn("a:srgbClr"), val=color)
    # The theme is a plain binary part in python-pptx, so we replace its bytes.
    part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)


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
# Free areas of the Master's content slides: above the title (y < 0.7 in), the left margin
# (x < 0.65 in) and the right margin (x > 12.65 in). Text boxes never overlap these.


def marketing(master) -> None:
    shape(master, (11.6, -1.75, 2.4, 2.4), MIST, "ellipse")  # top right, above the title
    shape(master, (12.78, 0.8, 0.36, 0.36), TEAL, "ellipse")  # right margin
    shape(master, (12.88, 1.35, 0.16, 0.16), NAVY, "ellipse")


def engineering(master) -> None:
    for i in range(1, 27):  # faint blueprint grid behind the slide
        shape(master, (i * 0.5, 0, 0.01, H), GRID)
    for i in range(1, 15):
        shape(master, (0, i * 0.5, W, 0.01), GRID)
    shape(master, (0.3, 0.3, 0.5, 0.04), PERIWINKLE)  # corner bracket in the top-left margin
    shape(master, (0.3, 0.3, 0.04, 0.5), PERIWINKLE)
    label(master, (9.6, 0.22, 3.0, 0.35), "ENGINEERING", PERIWINKLE)


def finance(master) -> None:
    shape(master, (0, 0, 0.3, H), NAVY)  # side rail in the left margin
    heights, colors = [0.2, 0.32, 0.44, 0.56], [PALE_GREEN, LIGHT_GREEN, GREEN, NAVY]
    for i, (height, color) in enumerate(zip(heights, colors, strict=True)):  # above the title
        shape(master, (11.6 + i * 0.25, 0.68 - height, 0.16, height), color)


# team -> (team colour, bullet character, decoration)
TEAMS = {
    "marketing": (TEAL, "●", marketing),
    "engineering": (PERIWINKLE, "►", engineering),
    "finance": (GREEN, "▪", finance),
}


def main() -> None:
    for team in TEAMS:
        print(build_template(team))


if __name__ == "__main__":
    main()
