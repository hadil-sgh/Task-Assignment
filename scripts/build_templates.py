"""Build the R. STAHL base template and its team variants.

python-pptx cannot create slide layouts from scratch, so this script starts from
python-pptx's built-in default template and restyles it in three places:

1. Theme: R. STAHL colours and font. The team accent colour is theme colour "accent1",
   so bullets and decorations follow the team without any per-team code.
2. Master: text styles, a footer with a text wordmark and slide number, and the team's
   one small decoration.
3. Layouts: keep the six text layouts, rename "Section Header" to "Key Message",
   and position every placeholder for a 16:9 slide.

All templates share the same layouts and placeholder indexes, so the generator never
changes per team. Each team differs only in accent colour, bullet character and decoration.

Run:  uv run python scripts/build_templates.py
Writes slidegen/templates/<team>.pptx and a filled showcase deck per team to out/showcase/.
"""

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.shapes.autoshape import CT_Shape
from pptx.shapes.autoshape import Shape
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = ROOT / "slidegen" / "templates"
SHOWCASE_DIR = ROOT / "out" / "showcase"

# Brand colours, read from the computed CSS of r-stahl.com.
NAVY = "072A5B"  # headings, primary brand colour
BLUE = "13387E"
MID_BLUE = "31549E"
LIGHT_BLUE = "D3E5F7"
ORANGE = "EB7D00"  # highlight colour used sparingly on the website
GREY = "D0D0D0"
TEXT = "2B2B2B"  # body text: near-black reads softer than pure black

# R. STAHL uses Univers, a licensed font. Its own website falls back to Arial,
# which is installed with every copy of Office, so the deck looks the same everywhere.
FONT = "Arial"

SLIDE_W, SLIDE_H = 13.333, 7.5  # inches, 16:9
MARGIN = 0.6
CONTENT_W = SLIDE_W - 2 * MARGIN
COL_W = (CONTENT_W - 0.4) / 2  # two columns with a 0.4 in gap
COL2_X = MARGIN + COL_W + 0.4

# Team look: accent colour, level-1 bullet, and one decoration rectangle (x, y, w, h in inches).
TEAMS = {
    "base": {"accent": ORANGE, "bullet": "■", "decoration": (MARGIN, 0, 1.2, 0.1)},
    "marketing": {"accent": "E0457B", "bullet": "●", "decoration": (0, 0, SLIDE_W, 0.1)},
    "engineering": {"accent": "0097A7", "bullet": "►", "decoration": (0, 0, 0.14, SLIDE_H)},
    "finance": {"accent": "43A047", "bullet": "–", "decoration": (MARGIN, 6.8, CONTENT_W, 0.05)},
}

# Layouts we keep (default name -> our name). Everything else is removed.
KEEP_LAYOUTS = {
    "Title Slide": "Title Slide",
    "Title and Content": "Title and Content",
    "Section Header": "Key Message",
    "Two Content": "Two Content",
    "Comparison": "Comparison",
    "Title Only": "Title Only",
}
FOOTER_TYPES = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}

TITLE_BOX = (MARGIN, 0.35, CONTENT_W, 1.15)
BODY_BOX = (MARGIN, 1.8, CONTENT_W, 4.8)


def build_template(team: str) -> Path:
    look = TEAMS[team]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    set_theme(prs, look["accent"])
    style_master(prs, look)
    style_layouts(prs)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    path = TEMPLATE_DIR / f"{team}.pptx"
    prs.save(path)
    return path


def set_theme(prs: Presentation, accent: str) -> None:
    """Write the brand palette and font into the theme, so PowerPoint offers them too."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    theme = etree.fromstring(part.blob)
    colors = {
        "dk1": TEXT, "lt1": "FFFFFF", "dk2": NAVY, "lt2": LIGHT_BLUE,
        "accent1": accent, "accent2": ORANGE, "accent3": BLUE, "accent4": MID_BLUE,
        "accent5": GREY, "accent6": NAVY, "hlink": BLUE, "folHlink": MID_BLUE,
    }  # fmt: skip
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


def style_master(prs: Presentation, look: dict) -> None:
    master = prs.slide_master
    old = master.element.find(qn("p:txStyles"))
    old.getparent().replace(old, parse_xml(text_styles(look["bullet"])))
    remove_footer_placeholders(master)
    for ph in master.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            place(ph, TITLE_BOX, MSO_ANCHOR.BOTTOM)
        else:
            place(ph, BODY_BOX)
    # Footer: thin grey rule, text wordmark (no logo), slide number, then the team decoration.
    add_rect(master, (MARGIN, 6.82, CONTENT_W, 0.012), MSO_THEME_COLOR.ACCENT_5)
    add_wordmark(master, MARGIN, 6.9, MSO_THEME_COLOR.TEXT_2)
    add_slide_number(master)
    add_rect(master, look["decoration"], MSO_THEME_COLOR.ACCENT_1)


def style_layouts(prs: Presentation) -> None:
    for layout in list(prs.slide_layouts):
        if layout.name not in KEEP_LAYOUTS:
            prs.slide_layouts.remove(layout)
            continue
        layout.element.cSld.set("name", KEEP_LAYOUTS[layout.name])
        remove_footer_placeholders(layout)
        style_layout(layout)


def style_layout(layout) -> None:
    """Position (and where needed restyle) each placeholder, by layout name and index."""
    ph = {p.placeholder_format.idx: p for p in layout.placeholders}
    name = layout.name
    if name == "Title Slide":
        layout.element.set("showMasterSp", "0")  # own look: no master footer or decoration
        set_background(layout, MSO_THEME_COLOR.TEXT_2)  # navy
        place(ph[0], (0.9, 1.9, 11.5, 2.0), MSO_ANCHOR.BOTTOM, size=40, bold=True, color="bg1")
        place(ph[1], (0.9, 4.25, 11.5, 1.3), size=22, color="bg2", bullets=False)
        add_rect(layout, (0.9, 4.02, 1.2, 0.08), MSO_THEME_COLOR.ACCENT_1)
        add_wordmark(layout, 0.9, 6.7, MSO_THEME_COLOR.BACKGROUND_1)
    elif name == "Key Message":
        set_background(layout, MSO_THEME_COLOR.BACKGROUND_2)  # light blue
        place(ph[0], (1.0, 1.4, 11.3, 2.9), MSO_ANCHOR.BOTTOM, size=36, bold=True, color="tx2")
        place(ph[1], (1.0, 4.55, 11.3, 1.5), size=20, color="tx1", bullets=False)
        add_rect(layout, (MARGIN, 1.6, 0.1, 4.3), MSO_THEME_COLOR.ACCENT_1)
    else:
        place(ph[0], TITLE_BOX, MSO_ANCHOR.BOTTOM)
    if name == "Title and Content":
        place(ph[1], BODY_BOX)
    elif name == "Two Content":
        place(ph[1], (MARGIN, 1.8, COL_W, 4.8))
        place(ph[2], (COL2_X, 1.8, COL_W, 4.8))
    elif name == "Comparison":
        for head, body, x in ((1, 2, MARGIN), (3, 4, COL2_X)):
            place(ph[head], (x, 1.8, COL_W, 0.6), MSO_ANCHOR.BOTTOM, size=22, bold=True,
                  color="tx2", bullets=False)  # fmt: skip
            place(ph[body], (x, 2.55, COL_W, 4.05))


def place(ph, box, anchor=MSO_ANCHOR.TOP, size=None, bold=False, color=None, bullets=True):
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
        lst_style.append(parse_xml(level_style(size, bold, color, bullets)))


def level_style(size, bold, color, bullets) -> str:
    indent = "" if bullets else ' marL="0" indent="0"'
    bullet = "" if bullets else "<a:buNone/>"
    run = (f' sz="{size * 100}"' if size else "") + (' b="1"' if bold else "")
    fill = f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>' if color else ""
    return f"<a:lvl1pPr {nsdecls('a')}{indent}>{bullet}<a:defRPr{run}>{fill}</a:defRPr></a:lvl1pPr>"


def text_styles(bullet: str) -> str:
    """Master text styles: titles bold navy 32 pt, body 20 pt with the team bullet."""

    def level(n, size, char, mar_l):
        return (
            f'<a:lvl{n}pPr marL="{mar_l}" indent="-285750">'
            '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>'
            '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/>'
            f'<a:buChar char="{char}"/><a:defRPr sz="{size}">'
            '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/>'
            f"</a:defRPr></a:lvl{n}pPr>"
        )

    return (
        f"<p:txStyles {nsdecls('a', 'p')}>"
        '<p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc>'
        '<a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill>'
        '<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
        f"<p:bodyStyle>{level(1, 2000, bullet, 285750)}{level(2, 1800, '–', 742950)}</p:bodyStyle>"
        '<p:otherStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
        "</p:txStyles>"
    )


def remove_footer_placeholders(owner) -> None:
    """Drop date/footer/slide-number placeholders; the master draws its own footer."""
    for ph in list(owner.placeholders):
        if ph.placeholder_format.type in FOOTER_TYPES:
            ph.element.getparent().remove(ph.element)


def set_background(layout, theme_color) -> None:
    fill = layout.background.fill
    fill.solid()
    fill.fore_color.theme_color = theme_color


# python-pptx can only add shapes to slides, so for masters and layouts we create the
# shape element with python-pptx's own constructor and append it to the shape tree.
def new_shape(owner, sp) -> Shape:
    owner.element.cSld.spTree.append(sp)
    return Shape(sp, owner.shapes)


def next_id(owner) -> int:
    return max(int(e.get("id")) for e in owner.element.iter(qn("p:cNvPr"))) + 1


def add_rect(owner, box, theme_color) -> None:
    x, y, w, h = (Inches(v) for v in box)
    sp = CT_Shape.new_autoshape_sp(next_id(owner), "Decoration", "rect", x, y, w, h)
    sp.remove(sp.find(qn("p:style")))  # the default shape style adds a drop shadow
    shape = new_shape(owner, sp)
    shape.fill.solid()
    shape.fill.fore_color.theme_color = theme_color
    shape.line.fill.background()  # no outline


def add_wordmark(owner, x, y, theme_color) -> None:
    sp = CT_Shape.new_textbox_sp(next_id(owner), "Wordmark", Inches(x), Inches(y),
                                 Inches(3), Inches(0.4))  # fmt: skip
    run = new_shape(owner, sp).text_frame.paragraphs[0].add_run()
    run.text = "R. STAHL"
    run.font.bold, run.font.size = True, Pt(12)
    run.font.color.theme_color = theme_color


def add_slide_number(owner) -> None:
    """A text box with a slide-number field: PowerPoint fills in the number on every slide."""
    sp = CT_Shape.new_textbox_sp(next_id(owner), "Slide Number", Inches(SLIDE_W - MARGIN - 1.5),
                                 Inches(6.9), Inches(1.5), Inches(0.4))  # fmt: skip
    shape = new_shape(owner, sp)
    paragraph = shape.text_frame.paragraphs[0]._p
    paragraph.getparent().replace(paragraph, parse_xml(
        f'<a:p {nsdecls("a")}><a:pPr algn="r"/>'
        '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
        '<a:rPr lang="en-US" sz="1100"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill>'
        "</a:rPr><a:t>‹#›</a:t></a:fld></a:p>"
    ))  # fmt: skip


# Showcase: one filled slide per layout, used only to render and eyeball the templates.
SHOWCASE = {
    "Title Slide": [
        "Ex-Lighting Upgrade 2026",
        "Status update for plant managers · 17 November 2026",
    ],
    "Title and Content": [
        "Phase 1 finished on time and 4% under budget",
        "312 luminaires replaced in Zone 1 areas\nEnergy use down 38% versus legacy fittings\n"
        "No unplanned downtime during installation\nCertification audit passed on 3 November",
    ],
    "Key Message": [
        "Phase 2 can start in January if the budget is approved by 1 December",
        "Decision needed from the steering committee",
    ],
    "Two Content": [
        "Two sites, two different bottlenecks",
        "Weil: supplier lead time 14 weeks\nSpare parts stocked for 6 months\nTeam fully trained",
        "Tunis: installers booked until March\nPermits still pending for Hall 3\n"
        "Training starts in December",
    ],
    "Comparison": [
        "LED retrofit pays back faster than full replacement",
        "Retrofit", "EUR 410k total cost\nPayback in 2.1 years\nInstalled in 6 weeks",
        "Full replacement", "EUR 690k total cost\nPayback in 3.4 years\nInstalled in 11 weeks",
    ],
    "Title Only": ["Questions and next steps"],
}  # fmt: skip


def build_showcase(template: Path) -> Path:
    prs = Presentation(template)
    for layout in prs.slide_layouts:
        slide = prs.slides.add_slide(layout)
        texts = SHOWCASE[layout.name]
        for i, ph in enumerate(sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)):
            ph.text_frame.text = texts[i]
    SHOWCASE_DIR.mkdir(parents=True, exist_ok=True)
    path = SHOWCASE_DIR / template.name
    prs.save(path)
    return path


def main() -> None:
    for team in TEAMS:
        template = build_template(team)
        showcase = build_showcase(template)
        print(f"{team:12} {template.relative_to(ROOT)}  showcase: {showcase.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
