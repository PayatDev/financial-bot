"""
living_will_builder.py
สร้างหนังสือแสดงเจตนาเกี่ยวกับการรักษาพยาบาล
- Claude อ่าน field Living Will → เลือก template แบบ ก หรือ ข
- build Word 1 หน้า + embed font
"""
import os
import re
import json
import tempfile
import zipfile
import shutil

import anthropic
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT     = "TH Sarabun New"
BLACK    = RGBColor(0x00, 0x00, 0x00)
GRAY     = RGBColor(0x88, 0x88, 0x88)
RED      = RGBColor(0xAA, 0x00, 0x00)
NAVY     = RGBColor(0x1A, 0x3A, 0x5C)
LS       = 1.3
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")

# ── ข้อความหลัก 2 แบบ ────────────────────────────────────────────────
INTENT_A = """\
ข้าพเจ้าไม่ประสงค์จะรับการรักษาที่มีเป้าหมายเพื่อยืดการตายเท่านั้น โดยเฉพาะ

- การใส่เครื่องช่วยหายใจในกรณีที่ไม่มีโอกาสฟื้นตัว
- การกระตุ้นหัวใจในกรณีที่แพทย์วินิจฉัยว่าไม่มีโอกาสฟื้นตัว
- การให้อาหารทางสายยางในกรณีที่อยู่ในภาวะหมดสติถาวร
- การรักษาที่ก่อให้เกิดความทรมานโดยไม่มีประโยชน์ต่อการฟื้นตัว

ข้าพเจ้าประสงค์จะได้รับการดูแลแบบประคับประคองเพื่อบรรเทาความเจ็บปวด \
และให้สามารถจากไปได้อย่างสงบ รวมถึงยาบรรเทาปวดในปริมาณที่เพียงพอ"""

INTENT_B = """\
ข้าพเจ้าประสงค์จะรับการรักษาพยาบาลอย่างเต็มที่ทุกวิธีที่แพทย์เห็นสมควร \
รวมถึงการใช้เครื่องช่วยหายใจ การกระตุ้นหัวใจ และการรักษาอื่นๆ \
เพื่อยืดชีวิตให้นานที่สุดเท่าที่จะเป็นไปได้ \
ข้าพเจ้าต้องการโอกาสในการฟื้นตัวทุกโอกาสที่มี"""


# ── Claude: เลือก template ───────────────────────────────────────────
def choose_intent(client_data: dict) -> str:
    """คืน 'A' หรือ 'B'"""
    living_will = client_data.get("Living Will", "")
    if not living_will or living_will.strip() in ("ไม่ระบุ", "-", ""):
        return "B"

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=10,
        messages=[{"role": "user", "content":
            f"""อ่านข้อความนี้แล้วตอบแค่ตัวอักษร A หรือ B เท่านั้น

ข้อความ: "{living_will}"

A = ไม่ต้องการให้ยื้อชีวิต ปล่อยตามธรรมชาติ การดูแลแบบประคับประคอง
B = ต้องการรักษาเต็มที่ ยืดชีวิต หรือไม่ชัดเจน

ตอบ:"""}]
    )
    result = response.content[0].text.strip().upper()
    return "A" if result.startswith("A") else "B"


# ── Text helpers ──────────────────────────────────────────────────────
def _run(para, text, bold=False, sz=13, color=None, underline=False):
    run = para.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(sz)
    run.bold = bold
    run.font.color.rgb = color or BLACK
    if underline:
        run.underline = True
    return run

def _add_runs(para, text, sz=13):
    for part in re.split(r'(\[.*?\])', text):
        if not part:
            continue
        if part.startswith('['):
            _run(para, part, bold=True, sz=sz, color=RED, underline=True)
        else:
            _run(para, part, sz=sz)

def _p(doc, text, sz=13, bold=False, color=None,
       align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=4):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(sb)
    para.paragraph_format.space_after  = Pt(sa)
    para.paragraph_format.line_spacing = Pt(sz * LS)
    para.alignment = align
    if bold or color:
        _run(para, text, bold=bold, sz=sz, color=color)
    else:
        _add_runs(para, text, sz=sz)

def _head(doc, text, sz=14, sb=8, sa=2):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(sb)
    para.paragraph_format.space_after  = Pt(sa)
    para.paragraph_format.line_spacing = Pt(sz * LS)
    _run(para, text, bold=True, sz=sz)

def _bullet(doc, text, sz=13):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    para.paragraph_format.line_spacing = Pt(sz * LS)
    para.paragraph_format.left_indent  = Pt(12)
    _run(para, "- ", sz=sz)
    _add_runs(para, text, sz=sz)


# ── Signature table ───────────────────────────────────────────────────
def _no_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBdr = OxmlElement('w:tcBdr')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:color'), 'auto')
        tcBdr.append(el)
    tcPr.append(tcBdr)

def _no_table_border(tbl):
    tbl_elem = tbl._tbl
    tblPr = tbl_elem.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tblPr)
    tblBdr = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:color"), "auto")
        tblBdr.append(el)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(tblBdr)

def _sig_table(doc, pairs, sz=12, sb=8):
    n_cols    = len(pairs)
    TOTAL     = 9000
    GAP       = 800
    col_w     = (TOTAL - GAP * (n_cols - 1)) // n_cols
    col_widths = []
    for i in range(n_cols):
        col_widths.append(col_w)
        if i < n_cols - 1:
            col_widths.append(GAP)

    n_tc = len(col_widths)
    tbl  = doc.add_table(rows=3, cols=n_tc)
    tbl.style = "Table Grid"
    _no_table_border(tbl)

    for i, w in enumerate(col_widths):
        for row in tbl.rows:
            row.cells[i].width = w * 635

    for row in tbl.rows:
        for cell in row.cells:
            _no_border(cell)

    tbl.rows[0].height = Pt(30)
    tbl.rows[1].height = Pt(2)
    tbl.rows[2].height = Pt(sz * LS + 2)

    content_cols = [i * 2 for i in range(n_cols)]

    for idx, (role, _) in enumerate(pairs):
        col  = content_cols[idx]
        cell = tbl.rows[0].cells[col]
        cp   = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cp.paragraph_format.space_before = Pt(sb if idx == 0 else 0)
        cp.paragraph_format.space_after  = Pt(0)
        cp.paragraph_format.line_spacing = Pt(sz * LS)
        _run(cp, role, sz=sz, color=GRAY)

    for idx in range(n_cols):
        col  = content_cols[idx]
        cell = tbl.rows[1].cells[col]
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBdr = OxmlElement('w:tcBdr')
        for side in ['top', 'left', 'bottom', 'right']:
            el = OxmlElement(f'w:{side}')
            if side == 'top':
                el.set(qn('w:val'), 'single')
                el.set(qn('w:sz'), '6')
                el.set(qn('w:color'), 'AAAAAA')
            else:
                el.set(qn('w:val'), 'none')
                el.set(qn('w:sz'), '0')
                el.set(qn('w:color'), 'auto')
            tcBdr.append(el)
        old = tcPr.find(qn('w:tcBdr'))
        if old is not None:
            tcPr.remove(old)
        tcPr.append(tcBdr)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)

    for idx, (_, name_ph) in enumerate(pairs):
        col  = content_cols[idx]
        cell = tbl.rows[2].cells[col]
        cp   = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(2)
        cp.paragraph_format.space_after  = Pt(6)
        cp.paragraph_format.line_spacing = Pt(sz * LS)
        _run(cp, "(", sz=sz - 1, color=GRAY)
        _add_runs(cp, name_ph, sz=sz - 1)
        _run(cp, ")", sz=sz - 1, color=GRAY)


# ── Build docx ────────────────────────────────────────────────────────
def build_lw_docx(client_data: dict, intent: str) -> str:
    d  = client_data
    n  = d.get("ชื่อเล่น", "ผู้ทำหนังสือ")
    sp = d.get("คู่สมรส",  "คู่สมรส")
    fin_poa = d.get("Financial POA", sp)

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(0.787)   # 2.0 cm
        section.bottom_margin = Inches(0.787)   # 2.0 cm
        section.left_margin   = Inches(1.181)   # 3.0 cm
        section.right_margin  = Inches(0.787)   # 2.0 cm

    # ── Title ─────────────────────────────────────────────────────────
    _p(doc, "หนังสือแสดงเจตนาเกี่ยวกับการรักษาพยาบาล",
       sz=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, sa=3)
    _p(doc, "ตามมาตรา ๑๒ แห่งพระราชบัญญัติสุขภาพแห่งชาติ พ.ศ. ๒๕๕๐",
       sz=12, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, sa=8)

    _p(doc, "ทำที่ [ที่อยู่จัดทำเอกสาร]", sa=1)
    _p(doc, "วันที่ [วันที่ทำหนังสือ] เดือน [เดือน] พ.ศ. [ปี]", sa=6)

    # ── ข้อมูลผู้ทำหนังสือ ────────────────────────────────────────────
    _p(doc,
       f"ข้าพเจ้า [ชื่อ-นามสกุลจริงคุณ{n}] อายุ [อายุคุณ{n}] ปี สัญชาติไทย "
       f"ถือบัตรประจำตัวประชาชนเลขที่ [เลขบัตรประชาชนคุณ{n}] "
       f"อยู่บ้านเลขที่ [ที่อยู่ปัจจุบันคุณ{n}]", sa=4)
    _p(doc,
       "ขณะที่ข้าพเจ้ามีสติสัมปชัญญะสมบูรณ์ดี ข้าพเจ้าขอแสดงเจตนาไว้ล่วงหน้าว่า "
       "หากข้าพเจ้าอยู่ในสภาวะใดสภาวะหนึ่งดังต่อไปนี้", sa=4)

    _bullet(doc, "อยู่ในระยะสุดท้ายของชีวิต ซึ่งแพทย์วินิจฉัยว่าไม่สามารถรักษาให้หายได้")
    _bullet(doc, "อยู่ในภาวะหมดสติถาวรซึ่งไม่มีโอกาสฟื้นคืนสติ")
    _bullet(doc, "อยู่ในภาวะที่การรักษาไม่มีโอกาสทำให้กลับมาใช้ชีวิตได้อย่างมีคุณภาพ")

    # ── เจตนา ─────────────────────────────────────────────────────────
    _head(doc, "เจตนาของข้าพเจ้า", sb=6, sa=3)

    if intent == "A":
        # แบบ ก — ไม่ยื้อ
        _p(doc,
           "ข้าพเจ้าไม่ประสงค์จะรับการรักษาที่มีเป้าหมายเพื่อยืดการตายเท่านั้น "
           "โดยเฉพาะ", sa=4)
        _bullet(doc, "การใส่เครื่องช่วยหายใจในกรณีที่ไม่มีโอกาสฟื้นตัว")
        _bullet(doc, "การกระตุ้นหัวใจในกรณีที่แพทย์วินิจฉัยว่าไม่มีโอกาสฟื้นตัว")
        _bullet(doc, "การให้อาหารทางสายยางในกรณีที่อยู่ในภาวะหมดสติถาวร")
        _bullet(doc, "การรักษาที่ก่อให้เกิดความทรมานโดยไม่มีประโยชน์ต่อการฟื้นตัว")
        _p(doc,
           "ข้าพเจ้าประสงค์จะได้รับการดูแลแบบประคับประคองเพื่อบรรเทาความเจ็บปวด "
           "และให้สามารถจากไปได้อย่างสงบ รวมถึงยาบรรเทาปวดในปริมาณที่เพียงพอ",
           sb=4, sa=4)
    else:
        # แบบ ข — ยื้อ
        _p(doc,
           "ข้าพเจ้าประสงค์จะรับการรักษาพยาบาลอย่างเต็มที่ทุกวิธีที่แพทย์เห็นสมควร "
           "รวมถึงการใช้เครื่องช่วยหายใจ การกระตุ้นหัวใจ และการรักษาอื่นๆ "
           "เพื่อยืดชีวิตให้นานที่สุดเท่าที่จะเป็นไปได้ "
           "ข้าพเจ้าต้องการโอกาสในการฟื้นตัวทุกโอกาสที่มี",
           sa=4)

    # ── ผู้ตัดสินใจแทน ────────────────────────────────────────────────
    _head(doc, "ผู้ตัดสินใจแทน", sb=6, sa=2)
    _p(doc,
       "หากข้าพเจ้าไม่สามารถแสดงเจตนาด้วยตนเองได้ ข้าพเจ้าขอมอบหมายให้ "
       "บุคคลต่อไปนี้เป็นผู้ตัดสินใจแทนตามเจตนาที่ระบุไว้", sa=4)
    _p(doc,
       f"หลัก: [ชื่อ-นามสกุลจริงคุณ{sp}]  "
       f"ความสัมพันธ์: [ความสัมพันธ์กับคุณ{n}]  "
       f"โทร: [เบอร์โทรคุณ{sp}]", sa=2)
    _p(doc,
       "สำรอง: [ชื่อ-นามสกุลผู้ตัดสินใจแทนสำรอง]  "
       "ความสัมพันธ์: [ความสัมพันธ์]  "
       "โทร: [เบอร์โทร]", sa=6)

    # ── ปิดท้าย ───────────────────────────────────────────────────────
    _p(doc,
       "ข้าพเจ้าได้ทำหนังสือฉบับนี้ด้วยความสมัครใจ มีสติสัมปชัญญะสมบูรณ์ "
       "และขอให้แพทย์และบุคลากรทางการแพทย์ทุกท่านปฏิบัติตามเจตนาที่ระบุไว้",
       sa=4)

    # ── ลายเซ็น ───────────────────────────────────────────────────────
    _sig_table(doc, [
        ("ผู้แสดงเจตนา", f"ชื่อ-นามสกุลจริงคุณ{n}"),
    ], sz=12, sb=4)

    _p(doc, "พยานรับรองว่าผู้แสดงเจตนามีสติสัมปชัญญะสมบูรณ์และลงนามด้วยความสมัครใจ",
       sb=8, sa=2)

    _sig_table(doc, [
        ("พยานที่ ๑", "ชื่อ-นามสกุลพยานที่ ๑"),
        ("พยานที่ ๒", "ชื่อ-นามสกุลพยานที่ ๒"),
    ], sz=12, sb=4)

    # ── Note ──────────────────────────────────────────────────────────
    _p(doc,
       "แนะนำ: ทำหนังสือแสดงเจตนาออนไลน์เพิ่มเติมได้ที่ "
       "e-livingwill.nationalhealth.or.th "
       "ระบบของสำนักงานคณะกรรมการสุขภาพแห่งชาติ ฟรี "
       "และโรงพยาบาลทั่วประเทศรู้จักระบบนี้",
       sz=11, color=GRAY, sb=6, sa=0)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    tmp.close()
    return tmp.name


# ── Embed fonts ───────────────────────────────────────────────────────
def embed_fonts(docx_path: str) -> str:
    out = docx_path.replace(".docx", "_emb.docx")
    shutil.copy(docx_path, out)

    font_map = {
        "THSarabunNew":            os.path.join(FONT_DIR, "THSarabunNew.ttf"),
        "THSarabunNew-Bold":       os.path.join(FONT_DIR, "THSarabunNew Bold.ttf"),
        "THSarabunNew-Italic":     os.path.join(FONT_DIR, "THSarabunNew Italic.ttf"),
        "THSarabunNew-BoldItalic": os.path.join(FONT_DIR, "THSarabunNew BoldItalic.ttf"),
    }
    with zipfile.ZipFile(out, "a") as z:
        for name, path in font_map.items():
            arcname = f"word/fonts/{name}.ttf"
            try:
                z.getinfo(arcname)
            except KeyError:
                if os.path.exists(path):
                    z.write(path, arcname)

    with zipfile.ZipFile(out, "r") as z:
        names = z.namelist()
        ft = z.read("word/fontTable.xml").decode()
        ct = z.read("[Content_Types].xml").decode()
        fr = z.read("word/_rels/fontTable.xml.rels").decode() \
             if "word/_rels/fontTable.xml.rels" in names else \
             '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        all_files = {n: z.read(n) for n in names}

    old = '<w:font w:name="TH Sarabun New"/>'
    new = (
        '<w:font w:name="TH Sarabun New">'
        '<w:embedRegular r:id="rIdF1"/>'
        '<w:embedBold r:id="rIdF2"/>'
        '<w:embedItalic r:id="rIdF3"/>'
        '<w:embedBoldItalic r:id="rIdF4"/>'
        "</w:font>"
    )
    ft = ft.replace(old, new) if old in ft else ft.replace("</w:fonts>", new + "</w:fonts>")
    rels = "".join([
        '<Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/THSarabunNew.ttf"/>',
        '<Relationship Id="rIdF2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/THSarabunNew-Bold.ttf"/>',
        '<Relationship Id="rIdF3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/THSarabunNew-Italic.ttf"/>',
        '<Relationship Id="rIdF4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/THSarabunNew-BoldItalic.ttf"/>',
    ])
    fr = fr.replace("</Relationships>", rels + "</Relationships>")
    if "ttf" not in ct:
        ct = ct.replace("</Types>", '<Default Extension="ttf" ContentType="application/x-font-ttf"/></Types>')

    all_files["word/fontTable.xml"] = ft.encode()
    all_files["word/_rels/fontTable.xml.rels"] = fr.encode()
    all_files["[Content_Types].xml"] = ct.encode()

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in all_files.items():
            z.writestr(name, data)

    os.unlink(docx_path)
    return out


# ── Entry point ───────────────────────────────────────────────────────
def build_living_will(client_data: dict) -> str:
    """คืน path ไฟล์ .docx พร้อมส่งลูกค้า"""
    print("กำลังเลือก template Living Will...")
    intent = choose_intent(client_data)
    print(f"✅ เลือกแบบ {'ก (ไม่ยื้อ)' if intent == 'A' else 'ข (ยื้อ)'}")
    raw_path   = build_lw_docx(client_data, intent)
    final_path = embed_fonts(raw_path)
    print("✅ สร้างหนังสือแสดงเจตนาเสร็จ")
    return final_path
