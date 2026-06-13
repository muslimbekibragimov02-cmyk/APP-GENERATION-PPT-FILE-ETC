"""
ELT Grammar Course Designer — PPTX Generator
Generates publication-ready .pptx files using pure Python (zipfile + OOXML).
13 slides per unit, 16:9 widescreen, Canva-quality design.
"""

import zipfile, os
from io import BytesIO

# ── EMU constants (English Metric Units: 1 inch = 914400 EMU) ─────────────────
W  = 12192000   # slide width  33.87 cm  (16:9)
H  =  6858000   # slide height 19.05 cm

CM = 360000     # 1 cm in EMU
IN = 914400     # 1 inch in EMU
PT = 12700      # 1 point in EMU

# ── Color palette ─────────────────────────────────────────────────────────────
NAVY   = "1F3864"
BLUE   = "2E75B6"
GOLD   = "BF8F00"
GREEN  = "548235"
LBLUE  = "DCE6F1"
LGREY  = "F2F2F2"
WHITE  = "FFFFFF"
BLACK  = "1A1A1A"
DGREY  = "404040"
GOLD_L = "FFF2CC"   # light gold tint

# ── Fonts ─────────────────────────────────────────────────────────────────────
TITLE_FONT = "Cambria"
BODY_FONT  = "Calibri"

# ─────────────────────────────────────────────────────────────────────────────
# XML escape
# ─────────────────────────────────────────────────────────────────────────────
def x(s): 
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

# ─────────────────────────────────────────────────────────────────────────────
# Shape / Text helpers (DrawingML)
# ─────────────────────────────────────────────────────────────────────────────

def solid_fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

def grad_fill(c1, c2):
    """Two-stop horizontal gradient."""
    return (f'<a:gradFill rot="1"><a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs>'
            f'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')

def no_fill():
    return '<a:noFill/>'

def ln(w_pt=0, color=None):
    if color:
        return f'<a:ln w="{w_pt*PT}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>'
    return f'<a:ln w="{w_pt*PT}"><a:noFill/></a:ln>'

def sp_pr(x_emu, y_emu, cx_emu, cy_emu, fill_xml, line_xml="", rounding=0):
    rnd = f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {rounding}"/></a:avLst></a:prstGeom>' if rounding else '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    return (f'<p:spPr>'
            f'<a:xfrm><a:off x="{x_emu}" y="{y_emu}"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
            f'{rnd}'
            f'{fill_xml}'
            f'{line_xml}'
            f'</p:spPr>')

def txBody(paragraphs_xml, anchor="ctr", wrap="square", lIns=0, rIns=0, tIns=0, bIns=0):
    return (f'<p:txBody>'
            f'<a:bodyPr anchor="{anchor}" wrap="{wrap}" lIns="{lIns}" rIns="{rIns}" tIns="{tIns}" bIns="{bIns}"/>'
            f'<a:lstStyle/>'
            f'{paragraphs_xml}'
            f'</p:txBody>')

def pPr(align="l", spcBef=0, spcAft=0, indent=0, marL=0):
    ind = f' indent="{indent}"' if indent else ""
    mar = f' marL="{marL}"' if marL else ""
    return f'<a:pPr algn="{align}" spcBef="{spcBef}" spcAft="{spcAft}"{ind}{mar}><a:spcPct val="100000"/></a:pPr>'

def rPr(bold=False, italic=False, color=WHITE, size_pt=18, font=BODY_FONT, underline=False):
    b  = ' b="1"' if bold else ' b="0"'
    i  = ' i="1"' if italic else ''
    u  = ' u="sng"' if underline else ''
    sz = size_pt * 100  # hundredths of a point
    return (f'<a:rPr lang="en-GB"{b}{i}{u} sz="{sz}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/>'
            f'</a:rPr>')

def run(text, bold=False, italic=False, color=WHITE, size_pt=18, font=BODY_FONT):
    return f'<a:r>{rPr(bold,italic,color,size_pt,font)}<a:t>{x(text)}</a:t></a:r>'

def br(size_pt=8, font=BODY_FONT, color=WHITE):
    return f'<a:br><a:rPr sz="{size_pt*100}" lang="en-GB"><a:latin typeface="{font}"/><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr></a:br>'

def para(*runs, align="l", spcBef=0, spcAft=0, marL=0):
    return f'<a:p>{pPr(align, spcBef, spcAft, marL=marL)}{"".join(runs)}</a:p>'

def bullet_para(text, color=WHITE, size_pt=16, bold=False, indent=342900, marL=342900):
    r = run(text, bold=bold, color=color, size_pt=size_pt)
    bullet = f'<a:pPr algn="l" indent="-{indent}" marL="{marL}" spcBef="60000" spcAft="20000"><a:buChar char="•"/></a:pPr>'
    return f'<a:p>{bullet}{r}</a:p>'

def icon_bullet_para(icon, text, icon_color=GOLD, text_color=WHITE, size_pt=16):
    ri = run(icon + "  ", bold=True, color=icon_color, size_pt=size_pt+2)
    rt = run(text, color=text_color, size_pt=size_pt)
    pp = f'<a:pPr algn="l" spcBef="80000" spcAft="30000"/>'
    return f'<a:p>{pp}{ri}{rt}</a:p>'

# ─────────────────────────────────────────────────────────────────────────────
# Shape builders
# ─────────────────────────────────────────────────────────────────────────────

_sp_id = [10]
def next_id():
    _sp_id[0] += 1
    return _sp_id[0]

def rect_shape(sid, x_emu, y_emu, cx_emu, cy_emu, fill_xml, line_xml="",
               text_xml="", rounding=0, shadow=False):
    shad = ""
    if shadow:
        shad = ('<p:spPr/>')  # placeholder; full shadow omitted for compat
    sp_pr_xml = sp_pr(x_emu, y_emu, cx_emu, cy_emu, fill_xml, line_xml, rounding)
    nv = (f'<p:nvSpPr>'
          f'<p:cNvPr id="{sid}" name="shape{sid}"/>'
          f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
          f'<p:nvPr/>'
          f'</p:nvSpPr>')
    body = text_xml if text_xml else '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
    return f'<p:sp>{nv}{sp_pr_xml}{body}</p:sp>'

def text_box(sid, x_emu, y_emu, cx_emu, cy_emu, paragraphs_xml,
             anchor="t", lIns=int(0.1*CM), rIns=int(0.1*CM),
             tIns=int(0.1*CM), bIns=int(0.1*CM)):
    sp_p = sp_pr(x_emu, y_emu, cx_emu, cy_emu, no_fill(), "")
    nv = (f'<p:nvSpPr>'
          f'<p:cNvPr id="{sid}" name="tb{sid}"/>'
          f'<p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>'
          f'<p:nvPr/>'
          f'</p:nvSpPr>')
    body = txBody(paragraphs_xml, anchor=anchor,
                  lIns=lIns, rIns=rIns, tIns=tIns, bIns=bIns)
    return f'<p:sp>{nv}{sp_p}{body}</p:sp>'

def colored_box_with_text(sid, x_emu, y_emu, cx_emu, cy_emu, fill,
                           paragraphs_xml="", line_color=None, rounding=0,
                           anchor="ctr", lIns=int(0.2*CM), rIns=int(0.2*CM),
                           tIns=int(0.15*CM), bIns=int(0.15*CM)):
    lx = ln(1, line_color) if line_color else ln(0)
    sp_p = sp_pr(x_emu, y_emu, cx_emu, cy_emu, solid_fill(fill), lx, rounding)
    nv = (f'<p:nvSpPr>'
          f'<p:cNvPr id="{sid}" name="box{sid}"/>'
          f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
          f'<p:nvPr/>'
          f'</p:nvSpPr>')
    body = txBody(paragraphs_xml, anchor=anchor,
                  lIns=lIns, rIns=rIns, tIns=tIns, bIns=bIns)
    return f'<p:sp>{nv}{sp_p}{body}</p:sp>'

# ─────────────────────────────────────────────────────────────────────────────
# Table builder
# ─────────────────────────────────────────────────────────────────────────────

def pptx_table(sid, x_emu, y_emu, cx_emu, cy_emu,
               headers, rows, col_widths_emu,
               hdr_fill=NAVY, even_fill=LBLUE, odd_fill=WHITE,
               font_size=14):
    """Build a DrawingML table shape."""
    row_h = max(int((cy_emu - int(0.6*CM)) // (len(rows) + 1)), int(0.5*CM))

    def tc(text, fill, bold=False, color=WHITE, italic=False):
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        rp = rPr(bold=bold, italic=italic, color=color, size_pt=font_size, font=BODY_FONT)
        return (f'<a:tc>'
                f'<a:txBody><a:bodyPr/><a:lstStyle/>'
                f'<a:p><a:pPr algn="l"/><a:r>{rp}<a:t>{x(text)}</a:t></a:r></a:p>'
                f'</a:txBody>'
                f'<a:tcPr>{fill_xml}'
                f'<a:lnL w="6350"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnL>'
                f'<a:lnR w="6350"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnR>'
                f'<a:lnT w="6350"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnT>'
                f'<a:lnB w="6350"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:lnB>'
                f'</a:tcPr></a:tc>')

    def tr(cells, fill):
        return f'<a:tr h="{row_h}">{"".join(cells)}</a:tr>'

    # header row
    hdr_cells = [tc(h, hdr_fill, bold=True, color=WHITE) for h in headers]
    table_rows = [tr(hdr_cells, hdr_fill)]
    for ri, row in enumerate(rows):
        fill = even_fill if ri % 2 == 0 else odd_fill
        txt_color = NAVY if fill == LBLUE else BLACK
        cells = [tc(str(cell), fill, bold=(ci==0), color=txt_color if ci>0 else NAVY)
                 for ci, cell in enumerate(row)]
        table_rows.append(tr(cells, fill))

    col_defs = "".join(f'<a:gridCol w="{w}"/>' for w in col_widths_emu)
    tbl_xml = (f'<a:tbl>'
               f'<a:tblPr firstRow="1" bandRow="1">'
               f'<a:tableStyleId>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>'
               f'</a:tblPr>'
               f'<a:tblGrid>{col_defs}</a:tblGrid>'
               f'{"".join(table_rows)}'
               f'</a:tbl>')

    nv = (f'<p:nvGraphicFramePr>'
          f'<p:cNvPr id="{sid}" name="tbl{sid}"/>'
          f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
          f'<p:nvPr/>'
          f'</p:nvGraphicFramePr>')
    xfrm = (f'<p:xfrm><a:off x="{x_emu}" y="{y_emu}"/>'
            f'<a:ext cx="{cx_emu}" cy="{cy_emu}"/></p:xfrm>')
    graphic = (f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
               f'{tbl_xml}</a:graphicData></a:graphic>')
    return f'<p:graphicFrame>{nv}{xfrm}{graphic}</p:graphicFrame>'

# ─────────────────────────────────────────────────────────────────────────────
# Slide assembly
# ─────────────────────────────────────────────────────────────────────────────

def slide(shapes_xml, bg_color=None):
    """Wrap shapes in a full slide XML."""
    bg = ""
    if bg_color:
        bg = (f'<p:bg><p:bgPr>'
              f'<a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>'
              f'<a:effectLst/></p:bgPr></p:bg>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:dc="http://purl.org/dc/elements/1.1/">
  <p:cSld>{bg}<p:spTree>
    <p:nvGrpSpPr>
      <p:cNvPr id="1" name="content"/><p:cNvGrpSpPr/><p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/>
        <a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm>
    </p:grpSpPr>
    {shapes_xml}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClr/></p:clrMapOvr>
</p:sld>"""

# ─────────────────────────────────────────────────────────────────────────────
# Helper: title bar for content slides
# ─────────────────────────────────────────────────────────────────────────────

def content_title_bar(title_text, accent_color=BLUE, icon=""):
    """Full-width navy top bar with title text + colored accent strip."""
    sid1, sid2, sid3 = next_id(), next_id(), next_id()
    # Dark bar
    bar = colored_box_with_text(
        sid1, 0, 0, W, int(1.1*CM),
        NAVY, "",
    )
    # Accent strip
    strip = colored_box_with_text(
        sid2, 0, int(1.1*CM), W, int(0.18*CM),
        accent_color, "",
    )
    # Title text
    title = text_box(
        sid3, int(0.5*CM), int(0.06*CM), W - int(1*CM), int(1.0*CM),
        para(
            run(icon + "  " if icon else "", bold=True, color=GOLD, size_pt=22, font=TITLE_FONT),
            run(title_text, bold=True, color=WHITE, size_pt=22, font=TITLE_FONT),
            align="l"
        ),
        anchor="ctr"
    )
    return bar + strip + title

# ─────────────────────────────────────────────────────────────────────────────
# The 13 slide builders
# ─────────────────────────────────────────────────────────────────────────────

def slide_01_title(u):
    """Slide 1: Dark navy title slide"""
    _sp_id[0] = 10
    shapes = []

    # Full background gradient
    shapes.append(rect_shape(next_id(), 0, 0, W, H,
                              grad_fill(NAVY, "0D2245")))

    # Gold decorative accent bar (left edge)
    shapes.append(rect_shape(next_id(), 0, 0, int(0.4*CM), H,
                              solid_fill(GOLD)))

    # Gold top decorative bar
    shapes.append(rect_shape(next_id(), 0, 0, W, int(0.3*CM),
                              solid_fill(GOLD)))

    # Unit label pill (gold background)
    shapes.append(colored_box_with_text(
        next_id(), int(1.2*CM), int(1.3*CM), int(5*CM), int(0.9*CM),
        GOLD,
        para(
            run(f"UNIT {u['unit_num']}  ·  {u['cefr']}",
                bold=True, color=NAVY, size_pt=16, font=TITLE_FONT),
            align="ctr"
        ),
        rounding=20000, anchor="ctr"
    ))

    # Course name (small, top)
    shapes.append(text_box(
        next_id(), int(1.2*CM), int(2.5*CM), W - int(2*CM), int(0.8*CM),
        para(run("English Grammar in Action — Level 1",
                 color="BFC9D6", size_pt=15, font=BODY_FONT), align="l"),
        anchor="ctr"
    ))

    # Main title (large white)
    shapes.append(text_box(
        next_id(), int(1.2*CM), int(3.3*CM), W - int(2*CM), int(3.0*CM),
        para(run(u['topic'], bold=True, color=WHITE, size_pt=40, font=TITLE_FONT),
             align="l"),
        anchor="t"
    ))

    # Sub-focus
    shapes.append(text_box(
        next_id(), int(1.2*CM), int(5.8*CM), W - int(2*CM), int(1.0*CM),
        para(run(u['sub_focus'], italic=True, color=LBLUE, size_pt=17, font=BODY_FONT),
             align="l"),
        anchor="ctr"
    ))

    # Bottom gold line
    shapes.append(rect_shape(next_id(), 0, H - int(0.25*CM), W, int(0.25*CM),
                              solid_fill(GOLD)))

    # CEFR badge (bottom right)
    shapes.append(colored_box_with_text(
        next_id(), W - int(3*CM), H - int(1.4*CM), int(2.6*CM), int(1.0*CM),
        BLUE,
        para(
            run(f"CEFR  {u['cefr']}", bold=True, color=WHITE, size_pt=14),
            align="ctr"
        ),
        rounding=20000, anchor="ctr"
    ))

    return slide("".join(shapes), bg_color=NAVY)


def slide_02_objectives(u):
    """Slide 2: Learning Objectives — icon + text rows"""
    _sp_id[0] = 50
    shapes = []
    shapes.append(content_title_bar("LEARNING OBJECTIVES", BLUE, "🎯"))

    icons = ["📌","✏️","🔍","💡","🗣️"]
    colors_cycle = [GOLD, LBLUE, GREEN+"CC", BLUE, GOLD]

    y_start = int(1.5*CM)
    box_h   = int(0.85*CM)
    gap     = int(0.22*CM)
    obj_w   = W - int(1.6*CM)

    for i, obj in enumerate(u['objectives']):
        yy = y_start + i * (box_h + gap)
        # Icon circle
        shapes.append(colored_box_with_text(
            next_id(), int(0.5*CM), yy, int(0.85*CM), box_h,
            BLUE,
            para(run(icons[i % len(icons)], size_pt=18, color=WHITE), align="ctr"),
            rounding=50000, anchor="ctr"
        ))
        # Objective text box
        shapes.append(colored_box_with_text(
            next_id(), int(1.55*CM), yy, obj_w, box_h,
            LGREY,
            para(run(obj, color=NAVY, size_pt=15, font=BODY_FONT), align="l"),
            line_color="CCCCCC", anchor="ctr",
            lIns=int(0.3*CM), rIns=int(0.2*CM)
        ))

    return slide("".join(shapes))


def slide_03_warmup(u):
    """Slide 3: Warm-Up Activity"""
    _sp_id[0] = 100
    shapes = []
    shapes.append(content_title_bar("WARM-UP ACTIVITY", GOLD, "⚡"))

    # Big activity name box
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(1.5*CM), W - int(1*CM), int(1.1*CM),
        GOLD,
        para(run(u['warmup']['name'], bold=True, color=NAVY,
                                size_pt=24, font=TITLE_FONT), align="l"),
        rounding=15000, anchor="ctr",
        lIns=int(0.5*CM)
    ))

    # Time badge
    shapes.append(colored_box_with_text(
        next_id(), W - int(3.2*CM), int(1.5*CM), int(2.7*CM), int(1.1*CM),
        NAVY,
        para(run(f"⏱  {u['warmup']['time']}", bold=True,
                                color=GOLD, size_pt=15), align="ctr"),
        rounding=15000, anchor="ctr"
    ))

    # Steps
    y0 = int(2.85*CM)
    for i, step in enumerate(u['warmup']['steps']):
        yy = y0 + i * int(1.2*CM)
        shapes.append(colored_box_with_text(
            next_id(), int(0.5*CM), yy, int(0.7*CM), int(0.9*CM),
            BLUE,
            para(run(str(i+1), bold=True, color=WHITE, size_pt=16), align="ctr"),
            rounding=50000, anchor="ctr"
        ))
        shapes.append(text_box(
            next_id(), int(1.4*CM), yy, W - int(1.9*CM), int(0.9*CM),
            para(run(step, color=DGREY, size_pt=15), align="l"),
            anchor="ctr"
        ))

    return slide("".join(shapes))


def slide_04_vocab(u):
    """Slide 4: Vocabulary Preview — 10 key words in card layout"""
    _sp_id[0] = 150
    shapes = []
    shapes.append(content_title_bar("VOCABULARY PREVIEW", BLUE, "📖"))

    # Select 10 most teachable words (first 10 that are verbs or core adj)
    ten = u['vocab'][:10]

    card_w = int((W - int(1.5*CM)) / 2) - int(0.3*CM)
    card_h = int(0.78*CM)
    col1_x = int(0.5*CM)
    col2_x = col1_x + card_w + int(0.4*CM)
    y_start = int(1.5*CM)
    gap     = int(0.18*CM)

    for i, (word, pos, cefr, defn, example) in enumerate(ten):
        col = i % 2
        row = i // 2
        xp  = col1_x if col == 0 else col2_x
        yp  = y_start + row * (card_h + gap)

        # Word chip
        shapes.append(colored_box_with_text(
            next_id(), xp, yp, int(2.2*CM), card_h,
            NAVY,
            para(run(word, bold=True, color=WHITE, size_pt=15), align="ctr"),
            rounding=12000, anchor="ctr"
        ))
        # POS tag
        shapes.append(colored_box_with_text(
            next_id(), xp + int(2.3*CM), yp, int(1.3*CM), card_h,
            GOLD,
            para(run(pos[:3], bold=True, color=NAVY, size_pt=12), align="ctr"),
            rounding=12000, anchor="ctr"
        ))
        # Definition
        shapes.append(colored_box_with_text(
            next_id(), xp + int(3.7*CM), yp, card_w - int(3.8*CM), card_h,
            LGREY,
            para(run(defn, color=DGREY, size_pt=13), align="l"),
            line_color="CCCCCC", anchor="ctr",
            lIns=int(0.2*CM)
        ))

    # Caption
    shapes.append(text_box(
        next_id(), int(0.5*CM), H - int(0.7*CM), W - int(1*CM), int(0.55*CM),
        para(run("📚  Full 30-word vocabulary bank in your workbook — Section 4.",
                 italic=True, color=BLUE, size_pt=13), align="l"),
        anchor="ctr"
    ))

    return slide("".join(shapes))


def slide_05_grammar_rule(u):
    """Slide 5: Grammar Rule + Form Table"""
    _sp_id[0] = 200
    shapes = []
    shapes.append(content_title_bar("GRAMMAR RULE", GOLD, "📐"))

    # Rule box (highlighted)
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(1.45*CM), W - int(1*CM), int(1.1*CM),
        GOLD_L,
        para(run("📌  " + u['grammar']['rule'],
                                bold=True, color=NAVY, size_pt=17), align="l"),
        line_color=GOLD, rounding=10000, anchor="ctr",
        lIns=int(0.4*CM)
    ))

    # Form table
    hdrs = u['grammar']['form_headers']
    rows = u['grammar']['form_rows']
    n    = len(hdrs)
    tbl_w = W - int(1*CM)
    col_w_each = tbl_w // n
    col_ws = [col_w_each] * n

    shapes.append(pptx_table(
        next_id(),
        int(0.5*CM), int(2.75*CM),
        tbl_w, H - int(3.0*CM),
        hdrs, rows, col_ws,
        font_size=14
    ))

    return slide("".join(shapes))


def slide_06_spelling_notes(u):
    """Slide 6: Spelling / Usage Notes + Watch Out box"""
    _sp_id[0] = 250
    shapes = []
    shapes.append(content_title_bar("SPELLING & USAGE NOTES", GOLD, "✏️"))

    if 'spelling_headers' in u['grammar']:
        hdrs = u['grammar']['spelling_headers']
        rows = u['grammar']['spelling_rows']
        n    = len(hdrs)
        tbl_w = W - int(1*CM)
        col_w_each = tbl_w // n
        col_ws = [col_w_each] * n
        tbl_h = int(0.6*CM) * (len(rows) + 1) + int(0.4*CM)
        shapes.append(pptx_table(
            next_id(),
            int(0.5*CM), int(1.5*CM),
            tbl_w, min(tbl_h, int(3.5*CM)),
            hdrs, rows, col_ws,
            font_size=13
        ))
        note_y = int(1.5*CM) + min(tbl_h, int(3.5*CM)) + int(0.3*CM)
    else:
        note_y = int(1.5*CM)

    # Watch Out box
    wo_h = int(1.4*CM)
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), note_y, W - int(1*CM), wo_h,
        LBLUE,
        para(
            run("⚠️  Watch Out!  ", bold=True, color=NAVY, size_pt=15),
            run(u['grammar']['note'], color=NAVY, size_pt=14),
            align="l"
        ),
        line_color=NAVY, anchor="ctr",
        lIns=int(0.4*CM)
    ))

    # Usage bullets
    usage_y = note_y + wo_h + int(0.35*CM)
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), usage_y, int(3.5*CM), int(0.6*CM),
        GOLD,
        para(run("Usage Examples", bold=True, color=NAVY, size_pt=14), align="ctr"),
        rounding=10000, anchor="ctr"
    ))
    for i, ex in enumerate(u['grammar']['usage'][:4]):
        yy = usage_y + int(0.75*CM) + i * int(0.6*CM)
        shapes.append(text_box(
            next_id(), int(0.5*CM), yy, W - int(1*CM), int(0.55*CM),
            para(run("▶  " + ex, color=DGREY, size_pt=14), align="l"),
            anchor="ctr"
        ))

    return slide("".join(shapes))


def slide_07_guided(u):
    """Slide 7: Guided Practice"""
    _sp_id[0] = 300
    shapes = []
    shapes.append(content_title_bar("GUIDED PRACTICE", GOLD, "✍️"))

    # Instruction
    shapes.append(text_box(
        next_id(), int(0.5*CM), int(1.5*CM), W - int(1*CM), int(0.65*CM),
        para(run(u['guided']['instruction'], italic=True, color=GOLD, size_pt=14), align="l"),
        anchor="ctr"
    ))

    # Show first 4 items (items 1-2 completed, 3-4 blank)
    items_to_show = u['guided']['items'][:4]
    hdrs = u['guided']['headers']
    tbl_w = W - int(1*CM)
    n = len(hdrs)
    col_ws = [int(tbl_w * r) for r in [0.06, 0.52, 0.42]]

    shapes.append(pptx_table(
        next_id(),
        int(0.5*CM), int(2.35*CM),
        tbl_w, H - int(2.6*CM),
        hdrs, items_to_show, col_ws,
        font_size=14
    ))

    return slide("".join(shapes))


def slide_08_controlled(u):
    """Slide 8: Controlled Practice"""
    _sp_id[0] = 350
    shapes = []
    shapes.append(content_title_bar("CONTROLLED PRACTICE", GOLD, "📝"))

    shapes.append(text_box(
        next_id(), int(0.5*CM), int(1.5*CM), W - int(1*CM), int(0.65*CM),
        para(run(u['controlled']['instruction'], italic=True, color=GOLD, size_pt=14), align="l"),
        anchor="ctr"
    ))

    items = u['controlled']['items'][:5]
    hdrs  = u['controlled']['headers']
    tbl_w = W - int(1*CM)
    col_ws = [int(tbl_w * r) for r in [0.06, 0.47, 0.47]]

    shapes.append(pptx_table(
        next_id(),
        int(0.5*CM), int(2.35*CM),
        tbl_w, H - int(2.6*CM),
        hdrs, items, col_ws,
        font_size=14
    ))

    return slide("".join(shapes))


def slide_09_error_correction(u):
    """Slide 9: Error Correction (semi-controlled)"""
    _sp_id[0] = 400
    shapes = []
    shapes.append(content_title_bar("ERROR CORRECTION", GOLD, "🔎"))

    shapes.append(text_box(
        next_id(), int(0.5*CM), int(1.5*CM), W - int(1*CM), int(0.65*CM),
        para(run("Find and correct the ONE mistake in each sentence.", italic=True,
                 color=GOLD, size_pt=14), align="l"),
        anchor="ctr"
    ))

    items = u['semi_controlled']['error_items'][:5]
    hdrs  = u['semi_controlled']['error_headers']
    tbl_w = W - int(1*CM)
    col_ws = [int(tbl_w * r) for r in [0.05, 0.50, 0.45]]

    shapes.append(pptx_table(
        next_id(),
        int(0.5*CM), int(2.35*CM),
        tbl_w, H - int(2.6*CM),
        hdrs, items, col_ws,
        font_size=14
    ))

    return slide("".join(shapes))


def slide_10_speaking(u):
    """Slide 10: Speaking Activity"""
    _sp_id[0] = 450
    shapes = []
    shapes.append(content_title_bar("SPEAKING ACTIVITY", GREEN, "🗣️"))

    # Activity name
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(1.45*CM), W - int(1*CM), int(1.0*CM),
        GREEN,
        para(run(u['speaking']['name'], bold=True, color=WHITE,
                                size_pt=22, font=TITLE_FONT), align="l"),
        rounding=12000, anchor="ctr",
        lIns=int(0.5*CM)
    ))

    # Speech bubble icon area
    shapes.append(colored_box_with_text(
        next_id(), W - int(3*CM), int(2.65*CM), int(2.5*CM), int(3.5*CM),
        LBLUE,
        para(run("💬", size_pt=60, color=BLUE), align="ctr"),
        rounding=20000, anchor="ctr"
    ))

    # Instructions
    for i, instr in enumerate(u['speaking']['instructions']):
        yy = int(2.65*CM) + i * int(0.88*CM)
        shapes.append(colored_box_with_text(
            next_id(), int(0.5*CM), yy, int(0.65*CM), int(0.72*CM),
            BLUE,
            para(run(str(i+1), bold=True, color=WHITE, size_pt=14), align="ctr"),
            rounding=50000, anchor="ctr"
        ))
        shapes.append(text_box(
            next_id(), int(1.3*CM), yy, W - int(4.8*CM), int(0.72*CM),
            para(run(instr, color=DGREY, size_pt=14), align="l"),
            anchor="ctr"
        ))

    # Pair work tag
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), H - int(0.8*CM), int(3.5*CM), int(0.6*CM),
        NAVY,
        para(run("👥  Pair Work", bold=True, color=GOLD, size_pt=13), align="ctr"),
        rounding=10000, anchor="ctr"
    ))

    return slide("".join(shapes))


def slide_11_reading_writing(u):
    """Slide 11: Reading & Writing"""
    _sp_id[0] = 500
    shapes = []
    shapes.append(content_title_bar("READING & WRITING", GREEN, "📄"))

    # Reading label
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(1.5*CM), int(2.5*CM), int(0.55*CM),
        GREEN,
        para(run("Reading", bold=True, color=WHITE, size_pt=13), align="ctr"),
        rounding=10000, anchor="ctr"
    ))

    # Reading text
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(2.2*CM), W - int(1*CM), int(1.7*CM),
        LGREY,
        para(run(u['reading']['text'], color=DGREY, size_pt=13,
                                italic=True), align="l"),
        line_color="CCCCCC", anchor="t",
        lIns=int(0.3*CM), tIns=int(0.2*CM)
    ))

    # Writing label
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(4.1*CM), int(2.5*CM), int(0.55*CM),
        GREEN,
        para(run("Writing", bold=True, color=WHITE, size_pt=13), align="ctr"),
        rounding=10000, anchor="ctr"
    ))

    # Writing instruction
    shapes.append(text_box(
        next_id(), int(0.5*CM), int(4.8*CM), W - int(1*CM), int(0.65*CM),
        para(run(u['writing']['instruction'], italic=True, color=GREEN, size_pt=14), align="l"),
        anchor="ctr"
    ))

    # First 3 writing prompts
    for i, pr in enumerate(u['writing']['prompts'][:3]):
        yy = int(5.6*CM) + i * int(0.38*CM)
        shapes.append(text_box(
            next_id(), int(0.5*CM), yy, W - int(1*CM), int(0.35*CM),
            para(run(f"• {pr}", color=DGREY, size_pt=13), align="l"),
            anchor="ctr"
        ))

    return slide("".join(shapes))


def slide_12_homework(u):
    """Slide 12: Homework"""
    _sp_id[0] = 550
    shapes = []
    shapes.append(content_title_bar("HOMEWORK", BLUE, "🏠"))

    # Item 1
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(1.5*CM), int(1.2*CM), int(1.2*CM),
        GOLD,
        para(run("1", bold=True, color=NAVY, size_pt=28, font=TITLE_FONT), align="ctr"),
        rounding=50000, anchor="ctr"
    ))
    shapes.append(colored_box_with_text(
        next_id(), int(1.9*CM), int(1.5*CM), W - int(2.4*CM), int(1.2*CM),
        LBLUE,
        para(
            run("Production Task  ", bold=True, color=NAVY, size_pt=15) +
            run(u['homework']['item1'], color=DGREY, size_pt=14),
            align="l"
        ),
        line_color=BLUE, anchor="ctr",
        lIns=int(0.3*CM)
    ))

    # Item 2
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), int(3.0*CM), int(1.2*CM), int(1.2*CM),
        BLUE,
        para(run("2", bold=True, color=WHITE, size_pt=28, font=TITLE_FONT), align="ctr"),
        rounding=50000, anchor="ctr"
    ))
    shapes.append(colored_box_with_text(
        next_id(), int(1.9*CM), int(3.0*CM), W - int(2.4*CM), int(1.2*CM),
        LGREY,
        para(
            run("Vocabulary Consolidation  ", bold=True, color=NAVY, size_pt=15) +
            run(u['homework']['item2'], color=DGREY, size_pt=14),
            align="l"
        ),
        line_color="CCCCCC", anchor="ctr",
        lIns=int(0.3*CM)
    ))

    # Due date note
    shapes.append(colored_box_with_text(
        next_id(), int(0.5*CM), H - int(1.0*CM), int(4*CM), int(0.6*CM),
        NAVY,
        para(run("📅  Due: next class", bold=True, color=GOLD, size_pt=13), align="ctr"),
        rounding=10000, anchor="ctr"
    ))

    return slide("".join(shapes))


def slide_13_closing(u):
    """Slide 13: Dark navy closing slide"""
    _sp_id[0] = 600
    shapes = []

    # Full background
    shapes.append(rect_shape(next_id(), 0, 0, W, H, grad_fill(NAVY, "0D2245")))

    # Gold accent bars
    shapes.append(rect_shape(next_id(), 0, 0, W, int(0.3*CM), solid_fill(GOLD)))
    shapes.append(rect_shape(next_id(), 0, H - int(0.3*CM), W, int(0.3*CM), solid_fill(GOLD)))
    shapes.append(rect_shape(next_id(), 0, 0, int(0.4*CM), H, solid_fill(GOLD)))

    # Big unit label
    shapes.append(colored_box_with_text(
        next_id(), int(1.2*CM), int(1.2*CM), int(5*CM), int(0.9*CM),
        GOLD,
        para(run(f"UNIT {u['unit_num']}  ·  {u['cefr']}",
                                bold=True, color=NAVY, size_pt=16, font=TITLE_FONT), align="ctr"),
        rounding=20000, anchor="ctr"
    ))

    # Topic text
    shapes.append(text_box(
        next_id(), int(1.2*CM), int(2.3*CM), W - int(2*CM), int(2.5*CM),
        para(run(u['topic'], bold=True, color=WHITE, size_pt=36, font=TITLE_FONT), align="l"),
        anchor="t"
    ))

    # Motivational one-liner in gold italic
    mottos = [
        "Every great speaker was once a student. Keep going!",
        "Questions are the engines of learning. Ask more!",
        "Grammar is not a rule — it is a tool. Use it!",
        "The more you practise, the more natural it feels.",
    ]
    motto = mottos[int(u['unit_num']) % len(mottos)]
    shapes.append(text_box(
        next_id(), int(1.2*CM), int(5.0*CM), W - int(2*CM), int(0.9*CM),
        para(run(f"✦  {motto}", italic=True, color=GOLD, size_pt=16), align="l"),
        anchor="ctr"
    ))

    # See you next class tag
    shapes.append(colored_box_with_text(
        next_id(), W - int(4.0*CM), H - int(1.1*CM), int(3.5*CM), int(0.65*CM),
        BLUE,
        para(run("See you next class! 👋", bold=True, color=WHITE, size_pt=13), align="ctr"),
        rounding=10000, anchor="ctr"
    ))

    return slide("".join(shapes), bg_color=NAVY)


# ─────────────────────────────────────────────────────────────────────────────
# PPTX file assembly
# ─────────────────────────────────────────────────────────────────────────────

def build_pptx(unit_data):
    u = unit_data
    slides = [
        slide_01_title(u),
        slide_02_objectives(u),
        slide_03_warmup(u),
        slide_04_vocab(u),
        slide_05_grammar_rule(u),
        slide_06_spelling_notes(u),
        slide_07_guided(u),
        slide_08_controlled(u),
        slide_09_error_correction(u),
        slide_10_speaking(u),
        slide_11_reading_writing(u),
        slide_12_homework(u),
        slide_13_closing(u),
    ]

    CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"  ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"   ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>""" + \
    "".join(
        f'\n  <Override PartName="/ppt/slides/slide{i+1}.xml"\n'
        f'    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(len(slides))
    ) + """
</Types>"""

    ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>"""

    PPT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="slideMasters/slideMaster1.xml"/>""" + \
    "".join(
        f'\n  <Relationship Id="rId{i+1}"\n'
        f'    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"\n'
        f'    Target="slides/slide{i+1}.xml"/>'
        for i in range(len(slides))
    ) + """
</Relationships>"""

    SLIDE_MASTER = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/>
      <a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
            accent2="accent2" accent3="accent3" accent4="accent4"
            accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lstStyle/></p:titleStyle>
    <p:bodyStyle><a:lstStyle/></p:bodyStyle>
    <p:otherStyle><a:lstStyle/></p:otherStyle>
  </p:txStyles>
</p:sldMaster>"""

    SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""

    SLIDE_LAYOUT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/>
      <a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClr/></p:clrMapOvr>
</p:sldLayout>"""

    SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

    slide_ids = "".join(
        f'<p:sldId id="{256+i}" r:id="rId{i+1}"/>'
        for i in range(len(slides))
    )

    PRESENTATION = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId0"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="{W}" cy="{H}" type="screen16x9"/>
  <p:notesSz cx="{H}" cy="{W}"/>
</p:presentation>"""

    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("ppt/presentation.xml", PRESENTATION)
        z.writestr("ppt/_rels/presentation.xml.rels", PPT_RELS)
        z.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        for i, sld_xml in enumerate(slides):
            z.writestr(f"ppt/slides/slide{i+1}.xml", sld_xml)
            # Each slide needs a .rels pointing to the layout
            sld_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""
            z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", sld_rels)

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Import unit data from generate_docx and run
# ─────────────────────────────────────────────────────────────────────────────
import sys
sys.path.insert(0, '/projects/sandbox/APP-GENERATION-PPT-FILE-ETC')
from generate_docx import UNITS

OUT_DIR = "/projects/sandbox/APP-GENERATION-PPT-FILE-ETC/generated"
os.makedirs(OUT_DIR, exist_ok=True)

topic_slugs = [
    "VerbToBe_Affirmative_Negative",
    "VerbToBe_Questions_ShortAnswers",
    "PresentContinuous_Affirmative_Negative",
    "PresentContinuous_Questions_ShortAnswers",
]

for i, (unit, slug) in enumerate(zip(UNITS, topic_slugs)):
    fname = f"Level1_Unit{i+1}_{slug}.pptx"
    path  = os.path.join(OUT_DIR, fname)
    data  = build_pptx(unit)
    with open(path, "wb") as f:
        f.write(data)
    print(f"✅ Generated: {fname}  ({len(data):,} bytes)")

print("\nAll PPTX files generated successfully.")
