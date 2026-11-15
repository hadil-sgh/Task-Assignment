"""Build the R. STAHL base template and its team variants.

python-pptx cannot create slide layouts from scratch, so this script starts from
python-pptx's built-in default template and reshapes it in three steps:

1. Theme: R. STAHL colours and font. Each team's colours go into accent1..6, so bullets,
   charts and tables pick them up automatically.
2. Master: text styles plus the team's decoration (template_designs.py).
3. Layouts: keep eight layouts, give them clear names, turn two of them into chart and
   table layouts, and position every placeholder for a 16:9 slide.

All templates share the same layout names and placeholder indexes, so the generator never
changes per team; only the look does.

Run:  uv run python scripts/build_templates.py
Writes slidegen/templates/<team>.pptx
"""

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches
from template_designs import (
    BOTTOM,
    DESIGNS,
    LIGHT_BLUE,
    NAVY,
    TEXT,
    WHITE,
    Design,
    H,
    W,
    place,
)

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
FONT = "Arial"  # R. STAHL uses the licensed Univers; its own website falls back to Arial
GAP = 0.4  # inches between columns

# Default layout name -> our name. Every other default layout is removed.
LAYOUTS = {
    "Title Slide": "Title Slide",
    "Title and Content": "Title and Content",
    "Section Header": "Key Message",
    "Two Content": "Two Content",
    "Comparison": "Comparison",
    "Title Only": "Title Only",
    "Content with Caption": "Chart and Text",
    "Title and Vertical Text": "Table",
}
FOOTER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}


def build_template(team: str) -> Path:
    design = DESIGNS[team]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    set_theme(prs, design)
    style_master(prs, design)
    for layout in list(prs.slide_layouts):
        if layout.name not in LAYOUTS:
            prs.slide_layouts.remove(layout)
            continue
        layout.element.cSld.set("name", LAYOUTS[layout.name])
        remove_footer_placeholders(layout)
        style_layout(layout, design)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    path = TEMPLATE_DIR / f"{team}.pptx"
    prs.save(path)
    return path


def set_theme(prs: Presentation, design: Design) -> None:
    """Write the palette and font into the theme, so PowerPoint, charts and tables use them."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    if design.dark:  # dk1 is the background, lt1 the text (see the colour map below)
        base = {"dk1": "0A1F3D", "lt1": WHITE, "dk2": "12345F", "lt2": "BFE9EE"}
    else:
        base = {"dk1": TEXT, "lt1": WHITE, "dk2": NAVY, "lt2": LIGHT_BLUE}
    accents = {f"accent{i}": color for i, color in enumerate(design.accents, start=1)}
    colors = base | accents | {"hlink": design.accents[0], "folHlink": design.accents[2]}
    scheme = theme.find(".//" + qn("a:clrScheme"))
    scheme.set("name", "R. STAHL")
    for slot, value in colors.items():
        element = scheme.find(qn(f"a:{slot}"))
        element.clear()
        etree.SubElement(element, qn("a:srgbClr"), val=value)
    for kind in ("a:majorFont", "a:minorFont"):
        theme.find(f".//{qn(kind)}/{qn('a:latin')}").set("typeface", FONT)
    # The theme is a plain binary part in python-pptx, so we replace its bytes.
    part._blob = etree.tostring(theme, xml_declaration=True, encoding="UTF-8", standalone=True)
    if design.dark:  # text = light theme colour, background = dark theme colour
        prs.slide_master.element.find(qn("p:clrMap")).attrib.update(
            {"bg1": "dk1", "tx1": "lt1", "bg2": "dk2", "tx2": "lt2"}
        )


def style_master(prs: Presentation, design: Design) -> None:
    master = prs.slide_master
    old = master.element.find(qn("p:txStyles"))
    old.getparent().replace(old, parse_xml(text_styles(design)))
    remove_footer_placeholders(master)
    for ph in master.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            place(ph, design.title_box, design.title_anchor)
        else:
            place(ph, design.body_box)
    design.master(master)


def style_layout(layout, design: Design) -> None:
    """Position each placeholder by layout name and index; title and key slides are custom."""
    ph = {p.placeholder_format.idx: p for p in layout.placeholders}
    if layout.name == "Title Slide":
        return design.title_slide(layout, ph)
    if layout.name == "Key Message":
        return design.key_message(layout, ph)
    place(ph[0], design.title_box, design.title_anchor)
    x, y, w, h = design.body_box
    half = (w - GAP) / 2
    if layout.name in ("Title and Content", "Table"):
        place(ph[1], design.body_box)
    elif layout.name == "Two Content":
        place(ph[1], (x, y, half, h))
        place(ph[2], (x + half + GAP, y, half, h))
    elif layout.name == "Comparison":
        for head, body, left in ((1, 2, x), (3, 4, x + half + GAP)):
            place(ph[head], (left, y, half, 0.6), BOTTOM, size=22, bold=True,
                  color=design.head_color, bullets=False)  # fmt: skip
            place(ph[body], (left, y + 0.75, half, h - 0.75))
    elif layout.name == "Chart and Text":
        chart_w = round(w * 0.62, 2)
        place(ph[1], (x, y, chart_w, h))
        place(ph[2], (x + chart_w + GAP, y, w - chart_w - GAP, h))
    if layout.name == "Chart and Text":
        set_placeholder_type(ph[1], "chart")
    elif layout.name == "Table":
        set_placeholder_type(ph[1], "tbl")


def set_placeholder_type(ph, kind: str) -> None:
    """Turn a content placeholder into a chart or table placeholder."""
    ph_element = ph.element.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
    ph_element.set("type", kind)
    ph_element.attrib.pop("orient", None)  # the default table source was vertical text
    ph.element.txBody.find(qn("a:bodyPr")).attrib.pop("vert", None)


def text_styles(design: Design) -> str:
    """Master text styles: bold titles in the team title colour, body 20 pt with team bullet."""

    def level(n: int, size: int, char: str, margin: int) -> str:
        return (
            f'<a:lvl{n}pPr marL="{margin}" indent="-285750">'
            '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>'
            '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/>'
            f'<a:buChar char="{char}"/><a:defRPr sz="{size}">'
            '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/>'
            f"</a:defRPr></a:lvl{n}pPr>"
        )

    return (
        f"<p:txStyles {nsdecls('a', 'p')}>"
        '<p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc>'
        f'<a:defRPr sz="3000" b="1"><a:solidFill><a:srgbClr val="{design.title_color}"/>'
        '</a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
        f"<p:bodyStyle>{level(1, 2000, design.bullet, 285750)}"
        f"{level(2, 1800, '–', 742950)}</p:bodyStyle>"
        '<p:otherStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
        "</p:txStyles>"
    )


def remove_footer_placeholders(owner) -> None:
    """Drop date/footer/slide-number placeholders; each design draws its own footer."""
    for ph in list(owner.placeholders):
        if ph.placeholder_format.type in FOOTER_TYPES:
            ph.element.getparent().remove(ph.element)


def main() -> None:
    for team in DESIGNS:
        print(build_template(team))


if __name__ == "__main__":
    main()
