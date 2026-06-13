"""
ELT Grammar Course Designer — DOCX Generator
Generates publication-ready .docx files using pure Python (zipfile + XML/OOXML).
No external libraries required.
"""

import zipfile
import os
from io import BytesIO

# ── Color palette ──────────────────────────────────────────────────────────────
NAVY   = "1F3864"
BLUE   = "2E75B6"
GOLD   = "BF8F00"
GREEN  = "548235"
LBLUE  = "DCE6F1"
LGREY  = "F2F2F2"
WHITE  = "FFFFFF"
BLACK  = "000000"

# ─────────────────────────────────────────────────────────────────────────────
# Low-level XML helpers
# ─────────────────────────────────────────────────────────────────────────────

def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def rpr(bold=False, italic=False, color=BLACK, size=22, font="Calibri", underline=False):
    """Run properties (size in half-points, so 11pt = 22)"""
    b  = "<w:b/>" if bold else ""
    i  = "<w:i/>" if italic else ""
    u  = '<w:u w:val="single"/>' if underline else ""
    return (f'<w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            f'<w:color w:val="{color}"/>'
            f'{b}{i}{u}'
            f'</w:rPr>')

def run(text, bold=False, italic=False, color=BLACK, size=22, font="Calibri", underline=False):
    r_pr = rpr(bold=bold, italic=italic, color=color, size=size, font=font, underline=underline)
    return f'<w:r>{r_pr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def para(content="", align="left", before=80, after=80, border_bottom=None, shading=None, indent_left=0):
    shade_xml = ""
    if shading:
        shade_xml = f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
    border_xml = ""
    if border_bottom:
        border_xml = f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="{border_bottom}"/></w:pBdr>'
    ind_xml = f'<w:ind w:left="{indent_left}"/>' if indent_left else ""
    return (f'<w:p><w:pPr>'
            f'<w:jc w:val="{align}"/>'
            f'<w:spacing w:before="{before}" w:after="{after}"/>'
            f'{shade_xml}{border_xml}{ind_xml}'
            f'</w:pPr>{content}</w:p>')

def heading(text, color, size=24, bold=True, border_color=None, before=160, after=60):
    r = run(text, bold=bold, color=WHITE, size=size, font="Calibri")
    shade_xml = f'<w:shd w:val="clear" w:color="auto" w:fill="{color}"/>'
    b_xml = ""
    if border_color:
        b_xml = f'<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="{border_color}"/></w:pBdr>'
    return (f'<w:p><w:pPr>'
            f'<w:spacing w:before="{before}" w:after="{after}"/>'
            f'{shade_xml}{b_xml}'
            f'<w:ind w:left="120" w:right="120"/>'
            f'</w:pPr>{r}</w:p>')

def sub_heading(text, color, size=22):
    r = run(text, bold=True, color=color, size=size)
    return (f'<w:p><w:pPr>'
            f'<w:spacing w:before="120" w:after="60"/>'
            f'<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="{color}"/></w:pBdr>'
            f'</w:pPr>{r}</w:p>')

def bullet(text, color=BLACK, bold_prefix=None):
    content = ""
    if bold_prefix:
        content += run(bold_prefix + " ", bold=True, color=color)
    content += run(text, color=color)
    return (f'<w:p><w:pPr>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            f'<w:spacing w:before="40" w:after="40"/>'
            f'<w:ind w:left="720" w:hanging="360"/>'
            f'</w:pPr>{content}</w:p>')

def note_box(text, title="📌 Note"):
    """Light blue background, navy left border note box"""
    xml  = f'<w:p><w:pPr>'
    xml += f'<w:shd w:val="clear" w:color="auto" w:fill="{LBLUE}"/>'
    xml += f'<w:pBdr>'
    xml += f'<w:left w:val="single" w:sz="24" w:space="4" w:color="{NAVY}"/>'
    xml += f'</w:pBdr>'
    xml += f'<w:spacing w:before="60" w:after="40"/>'
    xml += f'<w:ind w:left="240" w:right="120"/>'
    xml += f'</w:pPr>'
    xml += run(title + "  ", bold=True, color=NAVY, size=20)
    xml += run(text, color=NAVY, size=20)
    xml += f'</w:p>'
    return xml

def watch_out(text):
    return note_box(text, title="⚠️ Watch Out!")

def empty_para(before=60, after=60):
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr></w:p>'

# ─────────────────────────────────────────────────────────────────────────────
# Table helpers
# ─────────────────────────────────────────────────────────────────────────────

def tbl_cell(text, bg=None, bold=False, italic=False, color=BLACK, width=1500, size=20, align="left", colspan=1):
    shade = f'<w:shd w:val="clear" w:color="auto" w:fill="{bg}"/>' if bg else ""
    span  = f'<w:gridSpan w:val="{colspan}"/>' if colspan > 1 else ""
    r = run(text, bold=bold, italic=italic, color=color, size=size)
    return (f'<w:tc>'
            f'<w:tcPr>{span}<w:tcW w:w="{width}" w:type="dxa"/>{shade}'
            f'<w:tcBorders>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'</w:tcBorders>'
            f'</w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="{align}"/><w:spacing w:before="40" w:after="40"/>'
            f'<w:ind w:left="80" w:right="80"/></w:pPr>{r}</w:p>'
            f'</w:tc>')

def tbl_row(cells_xml):
    return f'<w:tr>{"".join(cells_xml)}</w:tr>'

def tbl_wrap(rows_xml, total_width=8200):
    return (f'<w:tbl>'
            f'<w:tblPr>'
            f'<w:tblW w:w="{total_width}" w:type="dxa"/>'
            f'<w:tblLayout w:type="fixed"/>'
            f'<w:tblBorders>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
            f'</w:tblBorders>'
            f'<w:tblCellMar>'
            f'<w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
            f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>'
            f'</w:tblCellMar>'
            f'</w:tblPr>'
            f'{"".join(rows_xml)}</w:tbl>')

def vocab_table(words):
    """30-word vocabulary table: Word | POS | CEFR | Definition | Example"""
    col_w = [1400, 900, 600, 2500, 2800]
    headers = ["Word", "Part of Speech", "CEFR", "Definition", "Example Sentence"]
    rows = []
    # Header row
    rows.append(tbl_row([
        tbl_cell(h, bg=NAVY, bold=True, color=WHITE, width=col_w[i], size=18)
        for i, h in enumerate(headers)
    ]))
    for idx, (word, pos, cefr, defn, example) in enumerate(words):
        bg = WHITE if idx % 2 == 0 else LGREY
        rows.append(tbl_row([
            tbl_cell(word,    bg=bg, bold=True,   color=NAVY,  width=col_w[0], size=18),
            tbl_cell(pos,     bg=bg,               color=BLACK, width=col_w[1], size=18),
            tbl_cell(cefr,    bg=bg,               color=BLUE,  width=col_w[2], size=18),
            tbl_cell(defn,    bg=bg,               color=BLACK, width=col_w[3], size=18),
            tbl_cell(example, bg=bg, italic=True,  color=BLACK, width=col_w[4], size=18),
        ]))
    return tbl_wrap(rows, total_width=sum(col_w))

def grammar_form_table(headers, rows_data, col_widths):
    rows = []
    rows.append(tbl_row([
        tbl_cell(h, bg=NAVY, bold=True, color=WHITE, width=col_widths[i], size=20)
        for i, h in enumerate(headers)
    ]))
    for ridx, row in enumerate(rows_data):
        bg = LBLUE if ridx % 2 == 0 else WHITE
        rows.append(tbl_row([
            tbl_cell(cell, bg=bg, color=NAVY if cidx == 0 else BLACK,
                     bold=(cidx == 0), width=col_widths[cidx], size=20)
            for cidx, cell in enumerate(row)
        ]))
    return tbl_wrap(rows, total_width=sum(col_widths))

def practice_table(headers, rows_data, col_widths):
    rows = []
    rows.append(tbl_row([
        tbl_cell(h, bg=BLUE, bold=True, color=WHITE, width=col_widths[i], size=20)
        for i, h in enumerate(headers)
    ]))
    for ridx, row in enumerate(rows_data):
        bg = WHITE if ridx % 2 == 0 else LGREY
        rows.append(tbl_row([
            tbl_cell(cell, bg=bg, color=BLACK, width=col_widths[cidx], size=20)
            for cidx, cell in enumerate(row)
        ]))
    return tbl_wrap(rows, total_width=sum(col_widths))

# ─────────────────────────────────────────────────────────────────────────────
# Unit banner (two-column header table)
# ─────────────────────────────────────────────────────────────────────────────

def unit_banner(unit_num, cefr, topic, sub_focus):
    left  = tbl_cell(f"Unit {unit_num}\n{cefr}", bg=NAVY, bold=True, color=WHITE, width=2000, size=28)
    right_content = run(topic, bold=True, color=NAVY, size=26, font="Calibri") + run(f"\n{sub_focus}", color=BLUE, size=20)
    right = (f'<w:tc><w:tcPr><w:tcW w:w="6200" w:type="dxa"/>'
             f'<w:shd w:val="clear" w:color="auto" w:fill="{LBLUE}"/>'
             f'<w:tcBorders>'
             f'<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
             f'<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
             f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
             f'<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
             f'</w:tcBorders>'
             f'</w:tcPr>'
             f'<w:p><w:pPr><w:spacing w:before="80" w:after="80"/>'
             f'<w:ind w:left="120" w:right="120"/></w:pPr>'
             f'{right_content}</w:p></w:tc>')
    return tbl_wrap([tbl_row([left, right])], total_width=8200)

# ─────────────────────────────────────────────────────────────────────────────
# DOCX file assembly
# ─────────────────────────────────────────────────────────────────────────────

def build_docx(unit_data):
    """Build a complete .docx from unit_data dict and return bytes."""

    body_xml = _build_body(unit_data)

    CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

    RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Target="numbering.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>"""

    STYLES = _build_styles()
    NUMBERING = _build_numbering()
    SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>"""

    DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
{body_xml}
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"
           w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>"""

    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", WORD_RELS)
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/settings.xml", SETTINGS)
    return buf.getvalue()

def _build_styles():
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
      <w:color w:val="{BLACK}"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>"""

def _build_numbering():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"""

# ─────────────────────────────────────────────────────────────────────────────
# Body builder — assembles all 11 sections
# ─────────────────────────────────────────────────────────────────────────────

def _build_body(u):
    xml = []

    # ── Title header (single line) ──────────────────────────────────────────
    xml.append(para(
        run("English Grammar in Action", bold=True, color=NAVY, size=20) +
        run(f"  ·  Level {u['level']}  ·  {u['course']}", color=BLUE, size=20),
        before=40, after=40
    ))

    # ── Unit banner ─────────────────────────────────────────────────────────
    xml.append(unit_banner(u['unit_num'], u['cefr'], u['topic'], u['sub_focus']))
    xml.append(empty_para(40, 40))

    # ── SECTION 1: Learning Objectives ──────────────────────────────────────
    xml.append(heading("1. LEARNING OBJECTIVES", NAVY))
    for obj in u['objectives']:
        xml.append(bullet(obj, color=BLACK))
    xml.append(empty_para())

    # ── SECTION 2: Warm-Up Activity ─────────────────────────────────────────
    xml.append(heading("2. WARM-UP ACTIVITY", BLUE))
    xml.append(para(
        run(u['warmup']['name'], bold=True, color=BLUE, size=22) +
        run(f"  ({u['warmup']['time']})", color=BLACK, size=20),
        before=60, after=40
    ))
    for step in u['warmup']['steps']:
        xml.append(bullet(step))
    xml.append(empty_para())

    # ── SECTION 3: Lead-In Discussion ───────────────────────────────────────
    xml.append(heading("3. LEAD-IN DISCUSSION", BLUE))
    for q in u['lead_in']:
        xml.append(bullet(q, color=BLACK))
    xml.append(empty_para())

    # ── SECTION 4: Vocabulary Preview ───────────────────────────────────────
    xml.append(heading("4. VOCABULARY PREVIEW", BLUE))
    xml.append(para(
        run("Study the 30 words below. They will appear throughout this unit.", italic=True, color=BLUE, size=20),
        before=40, after=60
    ))
    xml.append(vocab_table(u['vocab']))
    xml.append(empty_para())

    # ── SECTION 5: Grammar Presentation ─────────────────────────────────────
    xml.append(heading("5. GRAMMAR PRESENTATION", GOLD))
    xml.append(para(run(u['grammar']['rule'], bold=True, color=NAVY, size=22), before=60, after=40))
    xml.append(grammar_form_table(
        u['grammar']['form_headers'],
        u['grammar']['form_rows'],
        u['grammar']['form_col_widths']
    ))
    xml.append(empty_para(40, 40))
    if 'spelling_headers' in u['grammar']:
        xml.append(sub_heading("Spelling Rules", GOLD))
        xml.append(grammar_form_table(
            u['grammar']['spelling_headers'],
            u['grammar']['spelling_rows'],
            u['grammar']['spelling_col_widths']
        ))
        xml.append(empty_para(40, 40))
    xml.append(note_box(u['grammar']['note'], title="📌 Key Rule"))
    xml.append(empty_para(40, 40))
    xml.append(sub_heading("Usage Examples", GOLD))
    for ex in u['grammar']['usage']:
        xml.append(bullet(ex, color=BLACK))
    xml.append(empty_para())

    # ── SECTION 6: Guided Practice ──────────────────────────────────────────
    xml.append(heading("6. GUIDED PRACTICE", GOLD))
    xml.append(para(run(u['guided']['instruction'], italic=True, color=GOLD, size=20), before=40, after=40))
    xml.append(note_box("Items 1–2 are completed as models. Complete items 3–8 on your own.", "📝 How to use this section:"))
    xml.append(empty_para(40, 40))
    xml.append(practice_table(
        u['guided']['headers'],
        u['guided']['items'],
        u['guided']['col_widths']
    ))
    xml.append(empty_para())

    # ── SECTION 7: Controlled Practice ──────────────────────────────────────
    xml.append(heading("7. CONTROLLED PRACTICE", GOLD))
    xml.append(para(run(u['controlled']['instruction'], italic=True, color=GOLD, size=20), before=40, after=40))
    xml.append(practice_table(
        u['controlled']['headers'],
        u['controlled']['items'],
        u['controlled']['col_widths']
    ))
    xml.append(empty_para())

    # ── SECTION 8: Semi-Controlled Practice ─────────────────────────────────
    xml.append(heading("8. SEMI-CONTROLLED PRACTICE", GOLD))
    xml.append(sub_heading("Exercise D — Error Correction", GOLD))
    xml.append(para(run("Find and correct the ONE mistake in each sentence.", italic=True, color=GOLD, size=20), before=40, after=40))
    xml.append(practice_table(
        u['semi_controlled']['error_headers'],
        u['semi_controlled']['error_items'],
        u['semi_controlled']['error_col_widths']
    ))
    xml.append(empty_para(40, 40))
    xml.append(sub_heading("Exercise E — Sentence Transformation", GOLD))
    xml.append(para(run("Rewrite each sentence as instructed. Keep the same meaning.", italic=True, color=GOLD, size=20), before=40, after=40))
    xml.append(practice_table(
        u['semi_controlled']['transform_headers'],
        u['semi_controlled']['transform_items'],
        u['semi_controlled']['transform_col_widths']
    ))
    xml.append(empty_para())

    # ── SECTION 9: Speaking ──────────────────────────────────────────────────
    xml.append(heading("9. SPEAKING", GREEN))
    xml.append(para(
        run(u['speaking']['name'], bold=True, color=GREEN, size=22),
        before=60, after=40
    ))
    for instr in u['speaking']['instructions']:
        xml.append(bullet(instr, color=BLACK))
    xml.append(empty_para())

    # ── SECTION 10: Reading & Writing ───────────────────────────────────────
    xml.append(heading("10. READING & WRITING", GREEN))
    xml.append(sub_heading("Reading", GREEN))
    xml.append(para(
        run(u['reading']['text'], color=BLACK, size=20),
        shading=LGREY, before=60, after=40, indent_left=240
    ))
    xml.append(empty_para(40, 40))
    xml.append(para(run("True or False? Write T or F next to each statement.", italic=True, color=GREEN, size=20), before=40, after=40))
    xml.append(practice_table(
        u['reading']['tf_headers'],
        u['reading']['tf_items'],
        u['reading']['tf_col_widths']
    ))
    xml.append(empty_para(40, 40))
    xml.append(sub_heading("Writing", GREEN))
    xml.append(para(run(u['writing']['instruction'], italic=True, color=GREEN, size=20), before=40, after=40))
    for pr in u['writing']['prompts']:
        xml.append(bullet(pr, color=BLACK))
    xml.append(empty_para())

    # ── SECTION 11: Homework ────────────────────────────────────────────────
    xml.append(heading("11. HOMEWORK", BLUE))
    xml.append(para(run("Item 1 — Production Task", bold=True, color=BLUE, size=22), before=60, after=40))
    xml.append(para(run(u['homework']['item1'], color=BLACK, size=20), before=40, after=40, indent_left=360))
    xml.append(para(run("Item 2 — Vocabulary Consolidation", bold=True, color=BLUE, size=22), before=60, after=40))
    xml.append(para(run(u['homework']['item2'], color=BLACK, size=20), before=40, after=60, indent_left=360))
    xml.append(empty_para())

    return "\n".join(xml)


# ─────────────────────────────────────────────────────────────────────────────
# UNIT DATA — all 4 units
# ─────────────────────────────────────────────────────────────────────────────

UNITS = []

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 1 — Verb To Be: Affirmative & Negative (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "1",
    "cefr": "A1",
    "topic": "Verb To Be — Affirmative & Negative",
    "sub_focus": "am / is / are + positive & negative forms; contractions",
    "objectives": [
        "Form affirmative and negative sentences using am, is, and are.",
        "Use contractions correctly: I'm, she's, they're, isn't, aren't.",
        "Identify the correct form of to be for all subject pronouns.",
        "Distinguish between long and short (contracted) forms in writing.",
        "Produce simple descriptive sentences about people and things.",
    ],
    "warmup": {
        "name": "✋ True About Me",
        "time": "5 minutes",
        "steps": [
            "Write three sentences about yourself on a card using 'I am' or 'I'm'.",
            "Share with a partner — they decide: True or False?",
            "Example: 'I'm 20 years old.  I'm from London.  I'm not a teacher.'",
        ],
    },
    "lead_in": [
        "Look at the picture on the board. What can you see? Describe the people.",
        "Are you tired? Are you hungry? Are you a student? Answer naturally.",
    ],
    "vocab": [
        ("tall",       "adjective", "A1", "having great height",              "He is very tall."),
        ("tired",      "adjective", "A1", "feeling the need to rest",         "I'm tired today."),
        ("hungry",     "adjective", "A1", "feeling the need to eat",          "She is hungry now."),
        ("cold",       "adjective", "A1", "having a low temperature",         "My hands are cold."),
        ("hot",        "adjective", "A1", "having a high temperature",        "The soup is hot."),
        ("angry",      "adjective", "A1", "feeling strong displeasure",       "He is angry today."),
        ("afraid",     "adjective", "A1", "feeling fear",                     "She's afraid of cats."),
        ("sunny",      "adjective", "A1", "bright with sunlight",             "It's sunny outside."),
        ("warm",       "adjective", "A1", "fairly hot; comfortable",          "The room is warm."),
        ("heavy",      "adjective", "A1", "weighing a lot",                   "This bag is heavy."),
        ("dirty",      "adjective", "A1", "not clean",                        "The car is dirty."),
        ("open",       "adjective", "A1", "not closed",                       "The shop is open."),
        ("late",       "adjective", "A1", "arriving after the right time",    "You are late again."),
        ("cheap",      "adjective", "A1", "low in price",                     "These shoes are cheap."),
        ("big",        "adjective", "A1", "large in size",                    "The house is big."),
        ("new",        "adjective", "A1", "recently made or bought",          "My bag is new."),
        ("well",       "adjective", "A1", "in good health",                   "Are your parents well?"),
        ("policeman",  "noun",      "A1", "male police officer",              "He's a policeman."),
        ("taxi driver","noun",      "A1", "person who drives a taxi",         "He is a taxi driver."),
        ("friend",     "noun",      "A1", "person you like and trust",        "Ann is my friend."),
        ("table",      "noun",      "A1", "piece of furniture with a flat top","Your keys are on the table."),
        ("politics",   "noun",      "A1", "activities related to government", "She's interested in politics."),
        ("music",      "noun",      "A1", "sounds arranged in a pattern",     "He loves music."),
        ("Australian", "adjective", "A1", "from or related to Australia",     "They're Australian."),
        ("diamond",    "noun",      "A1", "a precious, very hard gemstone",   "Diamonds are not cheap."),
        ("key",        "noun",      "A1", "object used to lock or unlock",     "The keys are here."),
        ("window",     "noun",      "A1", "opening in a wall covered by glass","Close the window, please."),
        ("nationality","noun",      "A2", "the country a person belongs to",  "What's her nationality?"),
        ("interesting","adjective", "A1", "holding one's attention",          "His job is interesting."),
        ("comfortable","adjective", "A1", "giving physical ease",             "The chair is comfortable."),
    ],
    "grammar": {
        "rule": "Use am (I), is (he/she/it), or are (you/we/they) + adjective/noun/place to describe people and things.",
        "form_headers": ["Subject", "Affirmative", "Negative (full)", "Negative (contraction)"],
        "form_rows": [
            ["I",           "I am",       "I am not",       "I'm not"],
            ["He / She / It","He is",     "He is not",      "He isn't / He's not"],
            ["You / We / They","They are","They are not",   "They aren't / They're not"],
        ],
        "form_col_widths": [1600, 1900, 2200, 2500],
        "spelling_headers": ["Contraction", "Long form", "Example"],
        "spelling_rows": [
            ["I'm",      "I am",      "I'm a student."],
            ["She's",    "She is",    "She's my teacher."],
            ["They're",  "They are",  "They're not here."],
            ["isn't",    "is not",    "He isn't tired."],
            ["aren't",   "are not",   "We aren't late."],
        ],
        "spelling_col_widths": [1800, 2200, 4200],
        "note": "I'm not is the only correct negative contraction for I. NEVER say 'I amn't'.",
        "usage": [
            "Descriptions: She's tall. He's a doctor.",
            "Feelings: I'm tired. They're hungry.",
            "Location: The keys are on the table.",
            "Time: It's ten o'clock.",
            "Negatives: Tom isn't interested in politics.",
        ],
    },
    "guided": {
        "instruction": "Write the correct form of to be. Items 1–2 are completed for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "The weather ___ nice today.", "The weather IS nice today."],
            ["2 ✅", "I ___ not tired.", "I AM not tired."],
            ["3",   "This bag ___ heavy.", "________________________"],
            ["4",   "These bags ___ heavy.", "________________________"],
            ["5",   "My brother and I ___ good players.", "________________________"],
            ["6",   "Ann ___ at home. Her children ___ at school.", "________________________"],
            ["7",   "I ___ a taxi driver. My sister ___ a nurse.", "________________________"],
            ["8",   "It ___ sunny today but it ___ not warm.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write the full negative sentence using the correct form of to be.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(this house / not / very big)",     "________________________"],
            ["2", "(the shops / not / open today)",    "________________________"],
            ["3", "(my keys / in / my bag)",           "________________________"],
            ["4", "(Jenny / 18 years old)",            "________________________"],
            ["5", "(you / not / very tall)",           "________________________"],
            ["6", "(Canada / a very big country)",     "________________________"],
            ["7", "(diamonds / cheap)",                "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She are a nurse.", "________________________"],
            ["2", "They is at school.", "________________________"],
            ["3", "I amn't hungry.", "________________________"],
            ["4", "He not is my friend.", "________________________"],
            ["5", "Are the bag heavy?", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "He is a teacher. (make negative)",          "________________________"],
            ["2", "They are not tired. (use contraction)",     "________________________"],
            ["3", "I am hungry. (make negative using aren't)", "________________________"],
            ["4", "She is not from Canada. (write in full)",   "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Find Someone Who…",
        "instructions": [
            "Walk around the classroom. Ask your classmates questions with 'Are you…?'",
            "Use: 'Are you tired / hungry / cold / from [city] / interested in sport?'",
            "When a classmate answers 'Yes, I am,' write their name next to the question.",
            "Report back: 'Ana is tired. Tom isn't hungry.'",
        ],
    },
    "reading": {
        "text": (
            "My name is David. I am 25 years old. I am a nurse at a big hospital. "
            "My sister is a designer. She is not married. Our parents are retired. "
            "They are not young, but they are very healthy and happy."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "David is 25 years old.",               "___"],
            ["2", "David is a doctor.",                   "___"],
            ["3", "His sister is married.",               "___"],
            ["4", "His parents are retired.",             "___"],
            ["5", "His parents are unhealthy.",           "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 5 true sentences about yourself. Use am, is, are, isn't, or aren't.",
        "prompts": [
            "Write your name and age: My name is ___ . I am ___ years old.",
            "Write your job or role: I am a ___ .",
            "Write where you are from: I am from ___ .",
            "Write something you are NOT: I am not ___ .",
            "Write about a friend or family member: My ___ is / isn't ___ .",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about a person you know (friend or family member). "
            "Use am / is / are in at least 4 sentences, and isn't / aren't in at least 4 sentences. "
            "Underline every form of to be."
        ),
        "item2": (
            "Match each of the 30 vocabulary words from Section 4 to its definition. "
            "Write the matching word next to each definition in your workbook."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 2 — Verb To Be: Questions & Short Answers (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "2",
    "cefr": "A1",
    "topic": "Verb To Be — Questions & Short Answers",
    "sub_focus": "Yes/No questions; Wh- questions (What/Where/How/Who/Why); short answers",
    "objectives": [
        "Form Yes/No questions using Am I / Is he / Are they correctly.",
        "Produce Wh- questions using What, Where, Who, How, and Why with to be.",
        "Give grammatically correct short answers: Yes, I am. / No, he isn't.",
        "Apply correct word order in questions (verb before subject).",
        "Use question words + to be in natural conversational exchanges.",
    ],
    "warmup": {
        "name": "🎯 Question Ping-Pong",
        "time": "5 minutes",
        "steps": [
            "Teacher asks a student: 'Are you cold?' Student replies: 'Yes, I am.' or 'No, I'm not.'",
            "That student then asks the next person a different question using Is/Are.",
            "Continue around the class — no question word can be repeated twice in a row.",
        ],
    },
    "lead_in": [
        "Look at the two photos. Ask: 'Where are they? Who are they? Are they happy?'",
        "How do you ask someone's name, age, and job in English? Try together.",
    ],
    "vocab": [
        ("question",    "noun",      "A1", "a sentence asking for information",  "Ask a question now."),
        ("answer",      "noun",      "A1", "a response to a question",           "Your answer is right."),
        ("address",     "noun",      "A1", "where a person lives",               "What's your address?"),
        ("colour",      "noun",      "A1", "a visual property (red, blue…)",     "What colour is it?"),
        ("favourite",   "adjective", "A1", "most liked",                         "What's your favourite sport?"),
        ("sport",       "noun",      "A1", "physical activity or game",          "His favourite sport is tennis."),
        ("street",      "noun",      "A1", "a road in a town or city",           "She lives on a long street."),
        ("pence",       "noun",      "A1", "British currency (small unit)",      "It costs fifty pence."),
        ("postcard",    "noun",      "A1", "a card sent by post with a picture", "How much is this postcard?"),
        ("stamp",       "noun",      "A1", "small paper stuck on a letter",      "I need a stamp for this letter."),
        ("colleague",   "noun",      "A2", "a person you work with",             "She's my colleague."),
        ("appointment", "noun",      "A2", "a fixed time to meet someone",       "Is your appointment at two?"),
        ("Italian",     "adjective", "A1", "from or related to Italy",           "His wife is Italian."),
        ("German",      "adjective", "A1", "from or related to Germany",         "Is he German?"),
        ("Australian",  "adjective", "A1", "from or related to Australia",       "She's Australian, not British."),
        ("single",      "adjective", "A1", "not married",                        "Are you single or married?"),
        ("nervous",     "adjective", "A1", "feeling worry or anxiety",           "He's nervous about the test."),
        ("stuffy",      "adjective", "A1", "lacking fresh air; stiff",           "Is it stuffy in the room?"),
        ("asleep",      "adjective", "A1", "in a state of sleep",                "Is the baby asleep?"),
        ("expensive",   "adjective", "A1", "costing a lot of money",             "Is this coat expensive?"),
        ("near",        "preposition","A1","at a short distance",                "Is the bank near here?"),
        ("post office", "noun",      "A1", "place to send letters and parcels",  "Where's the post office?"),
        ("bus stop",    "noun",      "A1", "place where buses stop",             "The bus stop is near here."),
        ("seat",        "noun",      "A1", "a place to sit",                     "Is this seat free?"),
        ("factory",     "noun",      "A1", "a building where goods are made",    "The factory is open today."),
        ("camera",      "noun",      "A1", "device used to take photographs",    "Where's the camera?"),
        ("lawyer",      "noun",      "A1", "a person who practises law",         "Is your father a lawyer?"),
        ("designer",    "noun",      "A1", "a person who creates designs",       "She's a fashion designer."),
        ("passenger",   "noun",      "A1", "a traveller in a vehicle",           "Are the passengers ready?"),
        ("dictionary",  "noun",      "A1", "book that explains word meanings",   "Where is my dictionary?"),
    ],
    "grammar": {
        "rule": "To form a Yes/No question, put Am / Is / Are BEFORE the subject. For Wh- questions, put the question word first.",
        "form_headers": ["Question Type", "Structure", "Example"],
        "form_rows": [
            ["Yes/No question",  "Am/Is/Are + subject + …?",         "Is she at home?"],
            ["Short answer (✓)", "Yes, + subject + am/is/are.",       "Yes, she is."],
            ["Short answer (✗)", "No, + subject + isn't/aren't.",     "No, she isn't."],
            ["Wh- question",     "Wh-word + am/is/are + subject + …?","Where are you from?"],
        ],
        "form_col_widths": [2000, 3000, 3200],
        "spelling_headers": ["Question Word", "Use it to ask about…", "Example"],
        "spelling_rows": [
            ["What",       "things, ideas, names",     "What's your job?"],
            ["Where",      "places, locations",         "Where's the post office?"],
            ["Who",        "people",                    "Who's that woman?"],
            ["How",        "condition, manner, health", "How are your parents?"],
            ["How old",    "age",                       "How old is Joe?"],
            ["How much",   "price, uncountable",        "How much is this stamp?"],
            ["How many",   "countable quantity",        "How many books are there?"],
            ["What colour","colour",                    "What colour is your bag?"],
        ],
        "spelling_col_widths": [1600, 2600, 4000],
        "note": "Word order: Is your mother at home? ✅  Is at home your mother? ❌  — The verb ALWAYS comes before the subject in questions.",
        "usage": [
            "Yes/No: Are you tired?  —  Yes, I am. / No, I'm not.",
            "Wh-: Where's Jill?  —  She's at the library.",
            "How: How are you?  —  I'm very well, thanks.",
            "What colour: What colour is your car?  —  It's red.",
            "How much: How much are these postcards?  —  Fifty pence.",
        ],
    },
    "guided": {
        "instruction": "Write the question for each answer. Items 1–2 are completed for you.",
        "headers": ["#", "Answer", "Question"],
        "items": [
            ["1 ✅", "Yes, I am. (tired?)",             "Are you tired?"],
            ["2 ✅", "It's red. (colour of the car?)",   "What colour is your car?"],
            ["3",   "She's from London.",                "________________________"],
            ["4",   "He's 24 years old.",                "________________________"],
            ["5",   "No, she isn't. She's American.",    "________________________"],
            ["6",   "Very well, thank you.",             "________________________"],
            ["7",   "It's in your bag.",                 "________________________"],
            ["8",   "Fifty pence.",                      "________________________"],
        ],
        "col_widths": [500, 3700, 4000],
    },
    "controlled": {
        "instruction": "Write a short answer (Yes, I am. / No, he isn't. etc.) for each question.",
        "headers": ["#", "Question", "Short Answer"],
        "items": [
            ["1", "Are you married?",         "No, ___________________"],
            ["2", "Are you thirsty?",          "___________________"],
            ["3", "Is it cold today?",         "___________________"],
            ["4", "Are your hands cold?",      "___________________"],
            ["5", "Is it dark now?",           "___________________"],
            ["6", "Are you a teacher?",        "___________________"],
            ["7", "Is your friend English?",   "___________________"],
        ],
        "col_widths": [500, 3500, 4200],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "Is at home your mother?",          "________________________"],
            ["2", "Are new your shoes?",              "________________________"],
            ["3", "How old are have you?",            "________________________"],
            ["4", "What colour are it?",              "________________________"],
            ["5", "Are you from where?",              "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original statement", "Write as a Yes/No question"],
        "transform_items": [
            ["1", "She is from Canada.",               "________________________"],
            ["2", "They are at the post office.",      "________________________"],
            ["3", "He is a lawyer.",                   "________________________"],
            ["4", "The bag is expensive.",             "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "🗣️ Interview Your Partner",
        "instructions": [
            "Work in pairs. Student A is the interviewer; Student B is being interviewed.",
            "Use the prompts: name? / from? / age? / job? / married? / favourite sport?",
            "Ask questions like: 'What's your name?' / 'Where are you from?' / 'Are you married?'",
            "Then swap roles and repeat.",
        ],
    },
    "reading": {
        "text": (
            "Hi, my name is Paul. I'm from Sydney, so I'm Australian, not British. "
            "I'm 30 years old. I'm not a teacher — I'm a lawyer. My wife's name is Anna. "
            "She's Italian. She's 27 and she's a designer. She's not from Sydney. She's from Rome."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Paul is British.",                   "___"],
            ["2", "Paul is 30 years old.",              "___"],
            ["3", "Paul is a teacher.",                 "___"],
            ["4", "Anna is from Italy.",                "___"],
            ["5", "Anna is older than Paul.",           "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 5 questions you would ask a new classmate. Use Am/Is/Are and Wh- question words.",
        "prompts": [
            "Write a What question: What _______________?",
            "Write a Where question: Where _______________?",
            "Write a How old question: How old _______________?",
            "Write a Yes/No question using Is: Is _______________?",
            "Write a Yes/No question using Are: Are _______________?",
        ],
    },
    "homework": {
        "item1": (
            "Write a short paragraph (6–8 sentences) introducing a famous person. "
            "Use at least 3 Yes/No questions and 3 Wh- questions as if interviewing them. "
            "Include model short answers."
        ),
        "item2": (
            "Match each of the 30 vocabulary words from Section 4 to its definition. "
            "Write the matching word next to each definition in your workbook."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 3 — Present Continuous: Affirmative & Negative (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "3",
    "cefr": "A1",
    "topic": "Present Continuous — Affirmative & Negative",
    "sub_focus": "am/is/are + verb-ing; spelling rules; actions happening now",
    "objectives": [
        "Form affirmative Present Continuous sentences: subject + am/is/are + verb-ing.",
        "Form negative Present Continuous sentences using isn't / aren't.",
        "Apply correct -ing spelling rules (doubling, silent-e drop, -ie → -y).",
        "Use Present Continuous to describe actions happening at the moment of speaking.",
        "Distinguish Present Continuous from Present Simple in context.",
    ],
    "warmup": {
        "name": "🎭 Mime the Action",
        "time": "5 minutes",
        "steps": [
            "Teacher mimes an action (e.g., swimming, reading, cooking).",
            "Students call out: 'You're swimming!' or 'You're not eating!'",
            "A student then takes a turn miming — classmates make sentences.",
        ],
    },
    "lead_in": [
        "Look at the picture. What are the people doing right now? Name as many actions as you can.",
        "Think about right now: Are you sitting? Are you wearing shoes? Are you listening?",
    ],
    "vocab": [
        ("rain",       "verb",      "A1", "water falling from clouds",         "It's raining outside."),
        ("shine",      "verb",      "A1", "to give out bright light",          "The sun is shining."),
        ("run",        "verb",      "A1", "to move quickly on foot",           "They are running fast."),
        ("read",       "verb",      "A1", "to look at and understand text",    "She is reading a book."),
        ("walk",       "verb",      "A1", "to move at a normal pace",          "We are walking slowly."),
        ("write",      "verb",      "A1", "to put words on paper",             "He is writing a letter."),
        ("eat",        "verb",      "A1", "to put food in the mouth",          "She is not eating now."),
        ("ring",       "verb",      "A1", "to make a bell sound",              "The phone is ringing."),
        ("listen",     "verb",      "A1", "to pay attention to sound",         "Are you listening to me?"),
        ("sleep",      "verb",      "A1", "to rest with eyes closed",          "The baby is sleeping."),
        ("swim",       "verb",      "A1", "to move through water",             "He is swimming in the river."),
        ("cook",       "verb",      "A1", "to prepare food using heat",        "She's cooking dinner."),
        ("build",      "verb",      "A1", "to construct a structure",          "They're building a theatre."),
        ("stand",      "verb",      "A1", "to be in an upright position",      "You are standing on my foot."),
        ("stay",       "verb",      "A1", "to remain in a place",              "We're staying at a hotel."),
        ("sit",        "verb",      "A1", "to rest on a surface",              "She's sitting on a chair."),
        ("laugh",      "verb",      "A1", "to make sounds showing amusement",  "They are laughing loudly."),
        ("watch",      "verb",      "A1", "to look at carefully",              "He's watching television."),
        ("wear",       "verb",      "A1", "to have clothing on the body",      "She's wearing a new hat."),
        ("play",       "verb",      "A1", "to take part in a game",            "The children are playing."),
        ("have",       "verb",      "A1", "to possess; to experience",         "We're having dinner now."),
        ("work",       "verb",      "A1", "to do a job or task",               "Please be quiet. I'm working."),
        ("go",         "verb",      "A1", "to move from one place to another", "She's going to the park."),
        ("make",       "verb",      "A1", "to produce or create",              "She's making a cake."),
        ("sit down",   "phrasal v.","A1", "to lower oneself to a seat",        "He's sitting down slowly."),
        ("get up",     "phrasal v.","A1", "to rise from bed or a seat",        "She's getting up early."),
        ("come",       "verb",      "A1", "to move toward the speaker",        "The bus is coming."),
        ("try",        "verb",      "A1", "to attempt to do something",        "He's trying to fix it."),
        ("carry",      "verb",      "A1", "to hold and move something",        "She's carrying a heavy bag."),
        ("look",       "verb",      "A1", "to direct one's eyes at something", "She's looking at the board."),
    ],
    "grammar": {
        "rule": "Present Continuous = am / is / are + verb-ing. Use it for actions happening RIGHT NOW.",
        "form_headers": ["Subject", "Affirmative", "Negative"],
        "form_rows": [
            ["I",              "I am working.",          "I'm not working."],
            ["He / She / It",  "She is eating.",         "She isn't eating."],
            ["You / We / They","They are playing.",      "They aren't playing."],
        ],
        "form_col_widths": [1800, 3100, 3300],
        "spelling_headers": ["Rule", "Base verb", "-ing form"],
        "spelling_rows": [
            ["Most verbs: add -ing",           "work / read / eat",     "working / reading / eating"],
            ["End in -e: drop -e, add -ing",   "write / have / make",   "writing / having / making"],
            ["Short vowel + consonant: double","run / sit / swim / get","running / sitting / swimming / getting"],
            ["End in -ie: change to -y + ing", "die / lie",             "dying / lying"],
        ],
        "spelling_col_widths": [2800, 2200, 3200],
        "note": "Use Present Continuous for NOW: 'Please be quiet. I'm working.' NOT for habits — use Present Simple for habits.",
        "usage": [
            "Action now: The phone is ringing. Don't answer!",
            "Temporary action: We're staying at the Grand Hotel this week.",
            "Negatives: She isn't reading — she's watching TV.",
            "With 'look!': Look! It's snowing!",
        ],
    },
    "guided": {
        "instruction": "Write affirmative or negative Present Continuous sentences. Items 1–2 are done for you.",
        "headers": ["#", "Prompt", "Sentence"],
        "items": [
            ["1 ✅", "(have dinner) — They",               "They are having dinner."],
            ["2 ✅", "(not listen) — You",                  "You aren't listening."],
            ["3",   "(swim) — Somebody ___ in the river.",  "________________________"],
            ["4",   "(stay) — We ___ at the Central Hotel.","________________________"],
            ["5",   "(have) — She ___ a shower.",           "________________________"],
            ["6",   "(build) — They ___ a new theatre.",    "________________________"],
            ["7",   "(not watch TV) — Jane",                "________________________"],
            ["8",   "(sit on the floor) — She",             "________________________"],
        ],
        "col_widths": [500, 4000, 3700],
    },
    "controlled": {
        "instruction": "Write the -ing form of the verb. Then complete the sentence.",
        "headers": ["#", "Base Verb", "-ing form", "Complete sentence"],
        "items": [
            ["1", "write",  "___________", "Chris ___ a letter."],
            ["2", "have",   "___________", "We ___ dinner now."],
            ["3", "run",    "___________", "They ___ in the park."],
            ["4", "make",   "___________", "She ___ a cake."],
            ["5", "sit",    "___________", "He ___ on the floor."],
            ["6", "get",    "___________", "It ___ dark outside."],
        ],
        "col_widths": [500, 1500, 1800, 4400],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She is eat her lunch.",              "________________________"],
            ["2", "They are not played football.",      "________________________"],
            ["3", "The sun are shining.",               "________________________"],
            ["4", "I am not watching the TV now.",      "________________________"],
            ["5", "He writeing a letter.",              "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "She reads a book. (make Present Continuous)",        "________________________"],
            ["2", "They are playing. (make negative)",                  "________________________"],
            ["3", "He is working. (add 'right now')",                   "________________________"],
            ["4", "I run in the park. (make Present Continuous)",       "________________________"],
        ],
        "transform_col_widths": [500, 3500, 4200],
    },
    "speaking": {
        "name": "📸 Describe the Scene",
        "instructions": [
            "Work in pairs. Student A looks at the picture on the board for 10 seconds, then closes their eyes.",
            "Student B asks: 'What is the woman doing?', 'Are the children sitting?', etc.",
            "Student A answers from memory using Present Continuous.",
            "Then swap. The student with the most correct sentences wins.",
        ],
    },
    "reading": {
        "text": (
            "It is Saturday morning. Tom is at home. He isn't sleeping — he's cooking breakfast. "
            "His sister Anna is sitting in the garden. She's reading a magazine and listening to music. "
            "Their dog is running around the garden. It isn't raining. The sun is shining."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Tom is sleeping on Saturday morning.",  "___"],
            ["2", "Tom is cooking breakfast.",             "___"],
            ["3", "Anna is in the garden.",                "___"],
            ["4", "Anna is watching television.",          "___"],
            ["5", "It is raining.",                        "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 5 sentences about what people are doing RIGHT NOW (in class or at home). Use Present Continuous.",
        "prompts": [
            "Write about yourself: I am _______________.",
            "Write about a classmate (affirmative): ___ is _______________.",
            "Write about two people (affirmative): ___ and ___ are _______________.",
            "Write a negative sentence: I am not _______________.",
            "Write about what the teacher is doing: The teacher is _______________.",
        ],
    },
    "homework": {
        "item1": (
            "Look out of your window for 2 minutes. Write 8 sentences about what you see. "
            "Use Present Continuous (affirmative and negative). "
            "Underline every -ing verb."
        ),
        "item2": (
            "Match each of the 30 vocabulary words from Section 4 to its definition. "
            "Write the matching word next to each definition in your workbook."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 4 — Present Continuous: Questions & Short Answers (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "4",
    "cefr": "A1",
    "topic": "Present Continuous — Questions & Short Answers",
    "sub_focus": "Am/Is/Are + subject + verb-ing? | Short answers | Wh- questions with -ing",
    "objectives": [
        "Form Yes/No questions in the Present Continuous: Is/Are + subject + verb-ing?",
        "Form Wh- questions: What / Where / Why / Who + is/are + subject + verb-ing?",
        "Give correct short answers: Yes, he is. / No, they aren't.",
        "Apply correct word order in Present Continuous questions.",
        "Ask and answer questions about ongoing actions in natural conversation.",
    ],
    "warmup": {
        "name": "🔍 What Am I Doing?",
        "time": "5 minutes",
        "steps": [
            "Teacher mimes a secret action silently (e.g., cooking, reading, running).",
            "Students ask Yes/No questions: 'Are you cooking?' / 'Are you reading?'",
            "Teacher answers only with 'Yes, I am.' or 'No, I'm not.' until someone guesses.",
            "The student who guesses correctly takes the next turn.",
        ],
    },
    "lead_in": [
        "Look at the picture. What questions can you ask about what the people are doing?",
        "How do you form a question from: 'Paul is working today.' → Try it before looking at the grammar.",
    ],
    "vocab": [
        ("repair",      "verb",      "A1", "to fix something broken",          "He's repairing the roof."),
        ("paint",       "verb",      "A1", "to cover a surface with paint",     "She's painting the wall."),
        ("cry",         "verb",      "A1", "to shed tears",                     "Why is she crying?"),
        ("wait",        "verb",      "A1", "to stay until something happens",   "Who are you waiting for?"),
        ("arrive",      "verb",      "A1", "to reach a destination",            "Is the bus arriving now?"),
        ("leave",       "verb",      "A1", "to go away from a place",           "Are they leaving today?"),
        ("talk",        "verb",      "A1", "to speak to someone",               "Who is she talking to?"),
        ("meet",        "verb",      "A1", "to come together with someone",     "Who are you meeting?"),
        ("smile",       "verb",      "A1", "to turn the corners of the mouth up","Why is he smiling?"),
        ("point",       "verb",      "A1", "to aim a finger towards something", "What is she pointing at?"),
        ("hurry",       "verb",      "A1", "to move or act quickly",            "Why are you hurrying?"),
        ("fall",        "verb",      "A1", "to drop down from a higher place",  "Is the rain falling?"),
        ("turn on",     "phrasal v.","A1", "to start a machine or device",      "Why are you turning on the light?"),
        ("turn off",    "phrasal v.","A1", "to stop a machine or device",       "Is she turning off the radio?"),
        ("look at",     "phrasal v.","A1", "to direct your eyes at something",  "What are you looking at?"),
        ("hold",        "verb",      "A1", "to carry in one's hands",           "What is he holding?"),
        ("open",        "verb",      "A1", "to cause to be no longer closed",   "Who is opening the door?"),
        ("close",       "verb",      "A1", "to shut",                           "Is she closing the window?"),
        ("move",        "verb",      "A1", "to go from one position to another","Why are they moving?"),
        ("help",        "verb",      "A1", "to make it easier for someone",     "Who is helping her?"),
        ("ask",         "verb",      "A1", "to put a question to someone",      "What are you asking?"),
        ("answer",      "verb",      "A1", "to reply to a question",            "Is he answering correctly?"),
        ("drive",       "verb",      "A1", "to operate a vehicle",              "Who is driving the car?"),
        ("ride",        "verb",      "A1", "to sit on and control a vehicle",   "Is she riding a bike?"),
        ("climb",       "verb",      "A1", "to go up something",                "Why is he climbing?"),
        ("cross",       "verb",      "A1", "to go from one side to the other",  "Are they crossing the road?"),
        ("study",       "verb",      "A1", "to learn something",                "What is she studying?"),
        ("learn",       "verb",      "A1", "to gain knowledge or a skill",      "Are you learning English?"),
        ("practise",    "verb",      "A2", "to do repeatedly to improve",       "Is he practising his speech?"),
        ("explain",     "verb",      "A2", "to make something clear",           "What is the teacher explaining?"),
    ],
    "grammar": {
        "rule": "For questions: put Am / Is / Are BEFORE the subject. For Wh- questions, put the question word first.",
        "form_headers": ["Type", "Structure", "Example"],
        "form_rows": [
            ["Yes/No",        "Is/Are + subject + verb-ing?",            "Is Paul working today?"],
            ["Short answer ✓","Yes, + subject + is/are.",                "Yes, he is."],
            ["Short answer ✗","No, + subject + isn't/aren't.",           "No, they aren't."],
            ["What",          "What + is/are + subject + doing?",        "What are the children doing?"],
            ["Where",         "Where + is/are + subject + going?",       "Where are your friends going?"],
            ["Why",           "Why + is/are + subject + verb-ing?",      "Why are you wearing a coat?"],
            ["Who",           "Who + is/are + subject + verb-ing?",      "Who are you waiting for?"],
        ],
        "form_col_widths": [1500, 3200, 3500],
        "spelling_headers": ["Correct Word Order ✅", "Incorrect Word Order ❌"],
        "spelling_rows": [
            ["Is Paul working today?",              "Is working Paul today?"],
            ["Where are those people going?",       "Where are going those people?"],
            ["What are the children doing?",        "What are doing the children?"],
        ],
        "spelling_col_widths": [4100, 4100],
        "note": "The verb ALWAYS comes directly after the question word. The subject comes AFTER the verb: Wh-word + am/is/are + subject + -ing",
        "usage": [
            "Are you going now?  —  Yes, I am.",
            "Is it raining?  —  No, it isn't.",
            "What's Paul doing?  —  He's reading the newspaper.",
            "Where's Sally going?  —  She's going to the shops.",
            "Why are you wearing a coat? — Because it's cold.",
        ],
    },
    "guided": {
        "instruction": "Write the question correctly. Items 1–2 are completed for you.",
        "headers": ["#", "Words given", "Correct question"],
        "items": [
            ["1 ✅", "working / Paul / today?",            "Is Paul working today?"],
            ["2 ✅", "what / doing / the children?",       "What are the children doing?"],
            ["3",   "you / listening / to me?",            "________________________"],
            ["4",   "where / going / your friends?",       "________________________"],
            ["5",   "your parents / television / watching?","________________________"],
            ["6",   "what / cooking / Ann?",               "________________________"],
            ["7",   "why / you / looking / at me?",        "________________________"],
            ["8",   "coming / the bus?",                   "________________________"],
        ],
        "col_widths": [500, 3700, 4000],
    },
    "controlled": {
        "instruction": "Write a short answer for each question.",
        "headers": ["#", "Question", "Short Answer"],
        "items": [
            ["1", "Are you watching TV?",                   "No, _______________"],
            ["2", "Are you wearing a watch?",               "___________________"],
            ["3", "Are you eating something?",              "___________________"],
            ["4", "Is it raining?",                         "___________________"],
            ["5", "Are you sitting on the floor?",          "___________________"],
            ["6", "Are you feeling well?",                  "___________________"],
            ["7", "Is the teacher explaining something?",   "___________________"],
        ],
        "col_widths": [500, 3700, 4000],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "Is working Paul today?",                   "________________________"],
            ["2", "Where are going your friends?",            "________________________"],
            ["3", "What the children are doing?",             "________________________"],
            ["4", "Are you feel OK?",                         "________________________"],
            ["5", "Why is she crying? — No, she is.",         "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Statement", "Write as a question"],
        "transform_items": [
            ["1", "Paul is working today.",                   "________________________"],
            ["2", "Your friends are going to the stadium.",   "________________________"],
            ["3", "She is cooking dinner.",                   "________________________"],
            ["4", "They are not watching TV.",                "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "🎲 What's Happening in the Picture?",
        "instructions": [
            "Work in pairs. Each pair receives a scene picture (given by teacher).",
            "Take turns asking questions: 'What is the man doing?', 'Where are the children going?'",
            "Your partner answers using Present Continuous: 'He's repairing his bicycle.'",
            "Each student must ask at least 4 questions and answer 4 questions.",
        ],
    },
    "reading": {
        "text": (
            "It is 3 pm on a Monday. Sara is at the park with her brother. "
            "She is talking on her phone. Her brother is riding his bicycle. "
            "Two children are playing near the fountain. A man is walking his dog. "
            "Nobody is sitting on the benches — it is cold today."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Sara and her brother are at the park.",   "___"],
            ["2", "Sara is riding a bicycle.",               "___"],
            ["3", "Two children are playing.",               "___"],
            ["4", "The man is sitting on a bench.",          "___"],
            ["5", "It is warm today.",                       "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 5 questions and answers about actions happening right now. Use Present Continuous.",
        "prompts": [
            "Write a Yes/No question and answer: Are you ___? Yes/No, ___.",
            "Write a What question: What is ___ doing? He/She is ___.",
            "Write a Where question: Where are ___ going? They are ___.",
            "Write a Why question: Why are you ___? Because I am ___.",
            "Write a Who question: Who is ___? It's ___ — he/she is ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 questions in the Present Continuous about someone in your family. "
            "Include at least 2 Yes/No questions and 2 Wh- questions. "
            "Then write the answers. Underline every verb-ing form."
        ),
        "item2": (
            "Match each of the 30 vocabulary words from Section 4 to its definition. "
            "Write the matching word next to each definition in your workbook."
        ),
    },
})

# ─────────────────────────────────────────────────────────────────────────────
# Generate all docx files
# ─────────────────────────────────────────────────────────────────────────────
OUT_DIR = "/projects/sandbox/APP-GENERATION-PPT-FILE-ETC/generated"
os.makedirs(OUT_DIR, exist_ok=True)

topic_slugs = [
    "VerbToBe_Affirmative_Negative",
    "VerbToBe_Questions_ShortAnswers",
    "PresentContinuous_Affirmative_Negative",
    "PresentContinuous_Questions_ShortAnswers",
]

for i, (unit, slug) in enumerate(zip(UNITS, topic_slugs)):
    fname = f"Level1_Unit{i+1}_{slug}.docx"
    path  = os.path.join(OUT_DIR, fname)
    data  = build_docx(unit)
    with open(path, "wb") as f:
        f.write(data)
    print(f"✅ Generated: {fname}  ({len(data):,} bytes)")

print("\nAll DOCX files generated successfully.")
