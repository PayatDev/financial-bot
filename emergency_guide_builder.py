"""
emergency_guide_builder.py
สร้างคู่มือฉุกเฉิน Word จากข้อมูลลูกค้า
- fill ข้อมูลที่รู้จาก client_data
- placeholder สีแดงสำหรับข้อมูลที่ต้องกรอกเพิ่ม
- note เตือนลูกค้าว่าเป็นร่าง ต้องปรับแก้เอง
"""
import os
import re
import tempfile
import zipfile
import shutil

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
AMBER    = RGBColor(0x7A, 0x58, 0x00)
AMBER_BG = "FFFBEA"
LS       = 1.3
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")


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

def _section(doc, n, title, sz=14, sb=10, sa=2):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(sb)
    para.paragraph_format.space_after  = Pt(sa)
    para.paragraph_format.line_spacing = Pt(sz * LS)
    _run(para, f"หมวด {n} — ", bold=True, sz=sz, color=NAVY)
    _run(para, title, bold=True, sz=sz, color=NAVY)

def _row(doc, label, value_text, sz=12):
    """label: value"""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    para.paragraph_format.line_spacing = Pt(sz * LS)
    para.paragraph_format.left_indent  = Pt(12)
    _run(para, f"{label}: ", bold=True, sz=sz)
    _add_runs(para, value_text, sz=sz)

def _note_box(doc, text, sz=11):
    """กล่อง warning สีเหลือง"""
    from docx.oxml import OxmlElement
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)

    # bg สีเหลือง
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), AMBER_BG)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)

    # border สีเหลืองเข้ม
    tcBdr = OxmlElement('w:tcBdr')
    for side in ['top', 'left', 'bottom', 'right']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), 'D4A000')
        tcBdr.append(el)
    tcPr.append(tcBdr)

    cp = cell.paragraphs[0]
    cp.paragraph_format.space_before = Pt(4)
    cp.paragraph_format.space_after  = Pt(4)
    cp.paragraph_format.line_spacing = Pt(sz * LS)
    _run(cp, text, sz=sz, color=AMBER)


# ── Build docx ────────────────────────────────────────────────────────
def build_guide_docx(client_data: dict) -> str:
    d  = client_data
    n  = d.get("ชื่อเล่น",        "ผู้ทำเอกสาร")
    sp = d.get("คู่สมรส",         "[คู่สมรส]")
    executor   = d.get("ผู้จัดการมรดก",  "[ผู้จัดการมรดก]")
    urgent_mgr = d.get("ผู้จัดการฉุกเฉิน", "[ผู้จัดการฉุกเฉิน]")
    doc_loc    = d.get("ที่อยู่เอกสาร",  "[ที่เก็บเอกสาร]")
    occupation = d.get("อาชีพ",          "[อาชีพ]")
    insurance  = d.get("ประกันชีวิต",    "")
    welfare    = d.get("สวัสดิการ",      "")
    crypto     = d.get("assets_crypto_wallet", "")

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(0.787)
        section.bottom_margin = Inches(0.787)
        section.left_margin   = Inches(1.181)
        section.right_margin  = Inches(0.787)

    # ── Title ─────────────────────────────────────────────────────────
    _p(doc, f"คู่มือฉุกเฉินของคุณ{n}",
       sz=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, sa=2)
    _p(doc, "สำหรับผู้จัดการเรื่องเมื่อเกิดเหตุไม่คาดฝัน",
       sz=12, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, sa=2)

    _note_box(doc,
        "⚠️  เอกสารฉบับนี้เป็นร่างเบื้องต้น กรุณากรอกข้อมูลจริงให้ครบก่อนพิมพ์ "
        "และเก็บไว้ในที่ปลอดภัย อย่าเก็บไฟล์นี้ไว้ในโทรศัพท์หรืออีเมลที่ไม่มีรหัสผ่าน")

    # ── หมวด 1 ───────────────────────────────────────────────────────
    _section(doc, "๑", "สิ่งที่ต้องทำใน 24 ชั่วโมงแรก")
    _row(doc, "แจ้งตาย",
         "นำใบรับรองแพทย์ไปแจ้งที่สำนักงานเขตหรืออำเภอ ภายใน 24 ชั่วโมง")
    _row(doc, "ขอใบมรณบัตร",
         "ติดต่อโรงพยาบาลที่รักษา นำไปใช้ดำเนินเรื่องต่างๆ")
    _row(doc, "แจ้งนายจ้าง",
         f"{occupation} — ติดต่อ [ชื่อฝ่าย HR หรือหัวหน้างาน] โทร [เบอร์]")
    _row(doc, "เปิดพินัยกรรม",
         f"เอกสารเก็บอยู่ที่ {doc_loc}")

    # ── หมวด 2 ───────────────────────────────────────────────────────
    _section(doc, "๒", "เอกสารสำคัญและที่เก็บ")
    _row(doc, "พินัยกรรม / หนังสือมอบอำนาจ / หนังสือแสดงเจตนา",
         doc_loc)
    _row(doc, "โฉนดที่ดิน / ทะเบียนรถ",
         "[ที่เก็บโฉนดและทะเบียน]")
    _row(doc, "กรมธรรม์ประกัน",
         "[ที่เก็บกรมธรรม์]")
    _row(doc, "สมุดบัญชีธนาคาร",
         "[ที่เก็บสมุดบัญชี]")

    # ── หมวด 3 ───────────────────────────────────────────────────────
    _section(doc, "๓", "บัญชีธนาคารและการลงทุน")
    _row(doc, "ธนาคาร",
         "[ชื่อธนาคาร] เลขบัญชี [เลขบัญชี] สาขา [สาขา]")
    _row(doc, "บัญชีหุ้น",
         "[ชื่อโบรกเกอร์] หมายเลขบัญชี [เลขบัญชีหุ้น] โทร [เบอร์]")
    _row(doc, "กองทุน RMF / กองทุนเกษียณ",
         "[ชื่อบริษัทจัดการ] โทร [เบอร์]")

    # ── หมวด 4 ───────────────────────────────────────────────────────
    _section(doc, "๔", "ประกันชีวิต — เรียกร้องได้ทันที")
    if insurance:
        _row(doc, "ข้อมูลประกัน", insurance)
    _row(doc, "วิธีเรียกร้อง",
         "ติดต่อบริษัทประกันโดยตรง โทร [เบอร์บริษัทประกัน] กรมธรรม์เลข [เลขกรมธรรม์]")
    _row(doc, "หมายเหตุ",
         "ประกันชีวิตจ่ายตรงให้ผู้รับประโยชน์ ไม่ผ่านกองมรดก ติดต่อได้ทันทีหลังได้ใบมรณบัตร")

    # ── หมวด 5 ───────────────────────────────────────────────────────
    _section(doc, "๕", "สวัสดิการที่ต้องเรียกร้อง")
    if welfare:
        _row(doc, "สวัสดิการที่มี", welfare)
    _row(doc, "สวัสดิการนายจ้าง",
         "ติดต่อ [ฝ่าย HR] ภายใน [จำนวน] วัน")
    _row(doc, "ประกันสังคม",
         "ติดต่อสำนักงานประกันสังคมใกล้บ้าน ภายใน 90 วัน")

    # ── หมวด 6 ───────────────────────────────────────────────────────
    _section(doc, "๖", "บุคคลที่ต้องติดต่อ")
    _row(doc, "ผู้จัดการฉุกเฉิน",
         f"{urgent_mgr}  โทร [เบอร์{urgent_mgr}]")
    _row(doc, "ผู้จัดการมรดก",
         f"{executor}  โทร [เบอร์{executor}]")
    _row(doc, "ทนาย",
         "คุณพยัต จิรสุวรรณพงศ์  payat.jira@gmail.com")

    # ── หมวด 7 ───────────────────────────────────────────────────────
    _section(doc, "๗", "หนังสือแสดงเจตนาการรักษาพยาบาล (Living Will)")
    _row(doc, "ที่เก็บ",
         "อยู่ในชุดเอกสารนี้ และ/หรือที่ [ที่เก็บเอกสาร]")
    _row(doc, "เมื่อไหร่ควรใช้",
         "เมื่อแพทย์ถามว่าต้องการให้ยื้อชีวิตหรือไม่ หรือเมื่อผู้ป่วยอยู่ในภาวะหมดสติถาวร")
    _row(doc, "วิธีใช้",
         "นำหนังสือยื่นให้แพทย์หรือพยาบาลที่ดูแลได้ทันที ไม่ต้องรอขออนุญาตใคร")
    _row(doc, "สำคัญ",
         "หนังสือฉบับนี้มีผลตามกฎหมาย ม.๑๒ พ.ร.บ.สุขภาพแห่งชาติ พ.ศ.๒๕๕๐ แพทย์ต้องปฏิบัติตาม")

    _section(doc, "๘", "รหัสและความลับสำคัญ")
    _row(doc, "Password Manager",
         "[ชื่อ app เช่น Bitwarden] — Master password เก็บในซองปิดผนึกที่ [ที่เก็บ]")
    if crypto:
        _row(doc, "สินทรัพย์ดิจิทัล",
             "Seed phrase เขียนบนกระดาษ เก็บในเซฟที่ [ที่เก็บ] — "
             "ห้ามเปิดเผยกับใครเด็ดขาด ไม่ว่าจะอ้างตัวเป็นใคร ทางโทรศัพท์ อีเมล หรือออนไลน์ "
             "ถ้าไม่รู้วิธีจัดการ ให้ปรึกษาที่ claude.ai ก่อนดำเนินการใดๆ")
    else:
        _row(doc, "สินทรัพย์ดิจิทัล",
             "[ระบุหากมี] Seed phrase เก็บที่ [ที่เก็บ] — "
             "ห้ามเปิดเผยกับใครเด็ดขาด ถ้าไม่รู้วิธีจัดการ ให้ปรึกษาที่ claude.ai ก่อนดำเนินการใดๆ")
    _row(doc, "รหัสโทรศัพท์มือถือ",
         "เก็บในซองปิดผนึกที่ [ที่เก็บ]")
    _row(doc, "ความลับอื่นๆ",
         "ดูซองสีน้ำตาลที่ [ที่เก็บ]")

    _note_box(doc,
        "⚠️  ห้ามกรอกรหัสจริงลงในเอกสารนี้ ให้ระบุเฉพาะที่เก็บเท่านั้น "
        "หากต้องการเก็บรหัสจริง ให้ใช้ซองปิดผนึกและเก็บในเซฟแยกต่างหาก")

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
def build_emergency_guide(client_data: dict) -> str:
    """คืน path ไฟล์ .docx พร้อมส่งลูกค้า"""
    raw_path   = build_guide_docx(client_data)
    final_path = embed_fonts(raw_path)
    print("✅ สร้างคู่มือฉุกเฉินเสร็จ")
    return final_path
