"""Build the base template and its team variants on the text grid of the company Master.

python-pptx cannot create slide layouts from scratch, so this script starts from
python-pptx's built-in default template and reshapes it in three steps:

1. Theme: the Master's colours and font (Arial Narrow). Each team's colours go into
   accent1..6, so bullets, charts and tables pick them up automatically.
2. Master: text styles (bold uppercase titles, as in the Master) and the team decoration.
3. Layouts: the Master's set of text layouts, with every box at the Master's position and
   size (GRID in template_designs.py), plus a chart and a table layout.

All templates share the same layout names and placeholder indexes, so the generator never
changes per team; only the look does.

Run:  uv run python scripts/build_templates.py
Writes slidegen/templates/<team>.pptx
"""

from copy import deepcopy
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
    CHART_BOX,
    COLUMNS_2,
    COLUMNS_3,
    DEEP,
    DESIGNS,
    GRID,
    NAVY,
    PALE,
    PANEL,
    TEXT,
    TOP,
    WHITE,
    Design,
    H,
    W,
    place,
)

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "slidegen" / "templates"
FONT = "Arial Narrow"  # the font of the company Master template
SUBHEADLINE, SOURCE = 13, 22  # placeholder indexes we add, the same numbers as in the Master

# Our layout name -> (default layout it is made from, {placeholder index: role}).
# Placeholders not listed are removed; listed indexes that do not exist yet are added.
LAYOUTS = {
    "Title Slide": ("Title Slide", {0: "cover_title", 1: "cover_subtitle"}),
    "Title and Content": (
        "Title and Content", {0: "title", SUBHEADLINE: "subheadline", 1: "body", SOURCE: "source"},
    ),
    "Title and 2 Columns": (
        "Two Content",
        {0: "title", SUBHEADLINE: "subheadline", 1: "col2_0", 2: "col2_1", SOURCE: "source"},
    ),
    "Title and 3 Columns": (
        "Comparison",
        {0: "title", SUBHEADLINE: "subheadline", 2: "col3_0", 3: "col3_1", 4: "col3_2",
         SOURCE: "source"},
    ),
    "Title Only": ("Title Only", {0: "title", SUBHEADLINE: "subheadline"}),
    "Separation Slide": ("Section Header", {0: "separation_title", 1: "separation_subtitle"}),
    "End Slide": ("Title and Vertical Text", {0: "end_title", 1: "end_subtitle"}),
    "Title, Text and Chart": (
        "Content with Caption",
        {0: "title", SUBHEADLINE: "subheadline", 2: "col2_0", 1: "chart", SOURCE: "source"},
    ),
    "Title and Table": (
        "Picture with Caption", {0: "title", SUBHEADLINE: "subheadline", 1: "table", 2: "source"},
    ),
}  # fmt: skip

# Prompt text written into the layouts, as in the Master. The catalog reads "Optional ..."
# as an optional slot and shows the text to the model as a hint.
PROMPTS = {
    "subheadline": "Optional Subheadline",
    "source": "Optional: Sources and/or captions",
    "cover_subtitle": "Presenter, Place, Date",
    "separation_subtitle": "Optional Subtext",
    "end_title": "Thank you",
    "end_subtitle": "Optional text",
}


def build_template(team: str) -> Path:
    design = DESIGNS[team]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    set_theme(prs, design)
    style_master(prs, design)
    by_source = {source: (name, roles) for name, (source, roles) in LAYOUTS.items()}
    for layout in list(prs.slide_layouts):
        if layout.name not in by_source:
            prs.slide_layouts.remove(layout)
            continue
        name, roles = by_source[layout.name]
        layout.element.cSld.set("name", name)
        build_layout(layout, roles, design)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    path = TEMPLATE_DIR / f"{team}.pptx"
    prs.save(path)
    return path


def set_theme(prs: Presentation, design: Design) -> None:
    """Write the palette and font into the theme, so PowerPoint, charts and tables use them."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    if design.dark:  # dk1 is the background, lt1 the text (see the colour map below)
        base = {"dk1": DEEP, "lt1": WHITE, "dk2": PANEL, "lt2": PALE}
    else:
        base = {"dk1": TEXT, "lt1": WHITE, "dk2": NAVY, "lt2": PALE}
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
    remove_placeholders(master, keep=lambda ph: ph.placeholder_format.type not in FOOTER_TYPES)
    for ph in master.placeholders:
        is_title = ph.placeholder_format.type == PP_PLACEHOLDER.TITLE
        place(ph, GRID["title" if is_title else "body"][0])
    design.master(master)


def build_layout(layout, roles: dict[int, str], design: Design) -> None:
    """Keep, add and position the placeholders of one layout, then draw its decoration."""
    remove_placeholders(layout, keep=lambda ph: ph.placeholder_format.idx in roles)
    existing = {ph.placeholder_format.idx for ph in layout.placeholders}
    for idx in roles:
        if idx not in existing:
            add_placeholder(layout, idx)
    for ph in layout.placeholders:
        role = roles[ph.placeholder_format.idx]
        place_role(ph, role, design)
        if role in PROMPTS:
            ph.text_frame.text = PROMPTS[role]
    decoration = {"Title Slide": design.cover, "Separation Slide": design.separation,
                  "End Slide": design.end}.get(layout.name)  # fmt: skip
    if decoration:
        decoration(layout)


def place_role(ph, role: str, design: Design) -> None:
    """Put a placeholder at the Master's position for its role, with the Master's font size."""
    if role in ("title", "body"):
        return place(ph, GRID[role][0])  # styled by the master text styles
    if role.startswith("col"):
        lefts, width = (COLUMNS_2, 5.71) if role.startswith("col2") else (COLUMNS_3, 3.65)
        _, y, _, h = GRID["body"][0]
        return place(ph, (lefts[int(role[-1])], y, width, h))
    if role == "chart":
        place(ph, CHART_BOX)
        return retype(ph, "chart")
    if role == "table":
        _, y, _, h = GRID["body"][0]
        place(ph, (0.72, y, 11.89, h))
        return retype(ph, "tbl")
    box, size = GRID[role]
    colors = {
        "subheadline": design.sub_color, "source": None,
        "cover_title": design.cover_text[0], "cover_subtitle": design.cover_text[1],
        "separation_title": design.separation_text[0],
        "separation_subtitle": design.separation_text[1],
        "end_title": design.end_text[0], "end_subtitle": design.end_text[1],
    }  # fmt: skip
    if role == "end_subtitle":
        retype(ph, "body")  # its default source holds vertical text
    bold = role in ("cover_title", "separation_title", "end_title", "end_subtitle")
    anchor = BOTTOM if role == "source" else TOP
    place(ph, box, anchor, size=size, bold=bold, color=colors[role], bullets=False,
          caps=role == "end_subtitle")  # fmt: skip


def add_placeholder(layout, idx: int) -> None:
    """python-pptx cannot add placeholders, so we copy the master's body placeholder."""
    master_body = next(
        ph for ph in layout.slide_master.placeholders
        if ph.placeholder_format.type == PP_PLACEHOLDER.BODY
    )  # fmt: skip
    sp = deepcopy(master_body.element)
    sp.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}").set("idx", str(idx))
    c_nv_pr = sp.find(f"{qn('p:nvSpPr')}/{qn('p:cNvPr')}")
    c_nv_pr.set("id", str(max(int(e.get("id")) for e in layout.element.iter(qn("p:cNvPr"))) + 1))
    c_nv_pr.set("name", f"Text Placeholder {idx}")
    layout.element.cSld.spTree.append(sp)


def retype(ph, kind: str) -> None:
    """Change a placeholder's type, e.g. a content box into a chart or table placeholder."""
    ph_element = ph.element.find(f"{qn('p:nvSpPr')}/{qn('p:nvPr')}/{qn('p:ph')}")
    ph_element.set("type", kind)
    ph_element.attrib.pop("orient", None)  # some default sources hold vertical text
    if ph.element.txBody is not None:
        ph.element.txBody.find(qn("a:bodyPr")).attrib.pop("vert", None)


FOOTER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}


def remove_placeholders(owner, keep) -> None:
    for ph in list(owner.placeholders):
        if not keep(ph):
            ph.element.getparent().remove(ph.element)


def text_styles(design: Design) -> str:
    """Master text styles as in the company Master: bold uppercase titles, body 20 pt."""
    title_size, body_size = GRID["title"][1] * 100, GRID["body"][1] * 100
    if design.bullet:
        bullet = (
            '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/>'
            f'<a:buChar char="{design.bullet}"/>'
        )
        indent = 'marL="285750" indent="-285750"'
    else:
        bullet, indent = "<a:buNone/>", 'marL="0" indent="0"'
    return (
        f"<p:txStyles {nsdecls('a', 'p')}>"
        '<p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc>'
        f'<a:defRPr sz="{title_size}" b="1" cap="all"><a:solidFill>'
        f'<a:srgbClr val="{design.title_color}"/></a:solidFill><a:latin typeface="+mj-lt"/>'
        "</a:defRPr></a:lvl1pPr></p:titleStyle>"
        f"<p:bodyStyle><a:lvl1pPr {indent}>"
        '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="800"/></a:spcBef>'
        f'{bullet}<a:defRPr sz="{body_size}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
        '<p:otherStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
        "</p:txStyles>"
    )


def main() -> None:
    for team in DESIGNS:
        print(build_template(team))


if __name__ == "__main__":
    main()
