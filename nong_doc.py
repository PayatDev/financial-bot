import os
import json
import tempfile
import io
from datetime import datetime

import anthropic
from openpyxl import load_workbook
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

from drive_service import get_drive_service, upload_file_to_folder
from email_service import send_line_message

# ── หลักการทั่วไป 12 ประเด็น (hardcode) ──────────────────────────────
GENERAL_PRINCIPLES = {
    1: "ทันทีที่เราจากไป โลกไม่ได้หยุดรอ มีค่าใช้จ่ายเกิดขึ้นภายในไม่กี่ชั่วโมง ทั้งรักษาพยาบาลครั้งสุดท้าย ค่าจัดการงานศพ ค่าธรรมเนียมแจ้งตาย ไปจนถึงค่าจ้างทนายดำเนินเรื่องต่างๆ สิ่งเหล่านี้ต้องใช้เงินสดทันที ก่อนที่ใครจะได้เงินจากมรดกหรือประกันสักบาท ถ้าเราไม่เตรียมไว้การ คนที่เรารักต้องควักเงินตัวเองจ่ายไปก่อน ในช่วงเวลาที่เจ็บปวดที่สุดในชีวิต",
    2: "การสูญเสียคนในครอบครัวไม่ได้กระทบแค่ความรู้สึก แต่เขย่าวิถีชีวิตทุกด้านพร้อมกัน คนที่มีชีวิตอยู่ต่ออาจต้องลาออกหรือลดเวลาทำงาน เพื่อดูแลลูกๆ บ้านก็ยังต้องผ่อน ไหนจะค่าใช้จ่ายประจำวัน แต่รายได้หายไปครึ่งหนึ่ง ช่วง 6–12 เดือนแรก คือช่วงที่หนักที่สุดและต้องการเงินก้อนนึงเพื่อปรับตัว",
    3: "เมื่อเราจากไป ไม่ใช่แค่ตัวเราที่หายไป แต่รายได้ที่เราจะสร้างได้อีกหลายสิบปีก็หายไปด้วย ประกันชีวิตจะมีบทบาทเข้ามาเติมเต็มส่วนนี้ให้ครอบครัว เสมือนว่าเรายังคงหาเงินให้พวกเขาได้อยู่ แม้จะไม่ได้อยู่แล้ว นี่คือการแสดงความรักผ่านการวางแผน",
    4: "ถ้ามีกิจการส่วนตัว การจากไปกะทันหันอาจทำให้กิจการหยุดชะงัก เพราะไม่มีใครมีอำนาจลงนามหรืออนุมัติงานต่อได้ทันที ยิ่งถ้าครอบครัวยังต้องพึ่งรายได้จากกิจการนั้น ความล่าช้าในการบริหารอาจกลายเป็นภาระซ้ำซ้อนในช่วงเวลาที่ยากที่สุดอยู่แล้ว",
    5: "สิ่งที่น่ากลัวเหนือความตาย คือ การพิการหรือป่วยหนักจนทำงานไม่ได้ ประกันชีวิตก็ยังไม่จ่าย นั่นอาจสร้างภาระมากกว่าเพราะยังมีชีวิตอยู่แต่หาเงินไม่ได้ แถมยังมีค่ารักษาพยาบาลพิเศษเพิ่มเข้ามาอีก ",
    6: "ถ้าพินัยกรรมยกทุกอย่างให้คู่สมรสโดยไม่มีเงื่อนไข และหากวันหนึ่งคู่สมรสแต่งงานใหม่ ทรัพย์สินที่เราสร้างมาทั้งชีวิตอาจไหลไปสู่ครอบครัวใหม่ แทนที่จะถึงมือลูกแท้ๆ ของเรา เหตุการณ์แบบนี้เกิดขึ้นบ่อย เพราะพินัยกรรมที่ไม่ได้วางแผนไว้ล่วงหน้า",
    7: "พินัยกรรมส่วนใหญ่เขียนขึ้นโดยสมมติฐานว่าคู่สมรสยังมีชีวิตอยู่ แต่ถ้าทั้งสองจากไปพร้อมกันหรือคู่สมรสจากไปก่อน โดยที่ไม่มีแผนสำรองกรณีฉุกเฉิน มรดกอาจจะกระจายไปตามกฎหมายมรดก",
    8: "อุบัติเหตุที่พรากพ่อแม่ทั้งสองในคราวเดียวเป็นเหตุการณ์ที่ไม่มีใครอยากคิดถึง แต่ถ้าเกิดขึ้นจริงและไม่ได้วางแผนไว้ ศาลจะเป็นผู้ตัดสินว่าลูกจะไปอยู่กับใคร ระหว่างรอคำสั่ง อาจเกิดความวุ่นวายในครอบครัว ทั้งการแย่งตัวเด็กหรือการเกี่ยงกันรับผิดชอบ",
    9: "ถ้าผู้ปกครองและผู้ดูแลเงินของลูกเป็นคนเดียวกัน ไม่มีกลไกใดตรวจสอบว่าเงินถูกใช้เพื่อประโยชน์ของลูกจริงๆ การแยกสองบทบาทออกจากกันไม่ได้หมายความว่าไม่ไว้ใจ แต่เป็นระบบที่ทำให้ทั้งสองฝ่ายช่วยดูแลซึ่งกันและกันได้",
    10: "เด็กที่ได้รับมรดกก้อนใหญ่ทันทีที่บรรลุนิติภาวะ โดยที่ยังไม่มีประสบการณ์บริหารเงิน มักไม่ได้ผลลัพธ์อย่างที่พ่อแม่หวังไว้ การกำหนดเงื่อนไขในพินัยกรรม เช่น ให้ประกันชีวิตจ่ายเงินเมื่อครบอายุ 25 ปี หรือให้ทยอยจ่ายเป็นงวดๆ คือวิธีที่พ่อแม่ดูแลลูกข้ามเวลาได้",
    11: "เมื่อเกิดเหตุขึ้น ทุกอย่างจะวุ่นวายมาก คนที่รับหน้าที่ต้องรู้ว่าเอกสารอยู่ที่ไหน และจะเปิดเอกสารเอกสารเมื่อไหร่ ต้องตัดสิยใจเรื่องสำคัญอย่างไร บัญชีมีอะไรบ้าง และสิทธิ์อะไรที่ต้องเรียกร้องภายในกี่วัน ถ้าไม่มีคู่มือไว้ล่วงหน้า ความล่าช้าอาจทำให้ครอบครัวเสียสิทธิ์หรือได้รับเงินช้ากว่าที่ควร",
    12: "ไม่มีเอกสารกฎหมายฉบับไหนทดแทนสิ่งนี้ได้ จดหมายที่เราเขียนถึงคนที่รักจะบอกเล่าสิ่งที่ไม่สามารถใส่ในพินัยกรรมได้ ทั้งความรัก เหตุผลเบื้องหลังทุกการตัดสินใจ ค่านิยมที่อยากส่งต่อ และสิ่งที่อยากฝากไว้ให้ลูกเป็นเข็มทิศในการใช้ชีวิต",
}

ISSUE_TITLES = {
    1: "เงินพร้อมใช้ทันทีหลังจากไป",
    2: "เงินรองรับช่วงปรับตัว",
    3: "รายได้ที่หายไปตลอดชีวิต",
    4: "ความต่อเนื่องของกิจการ",
    5: "การพิการหรือทำงานไม่ได้",
    6: "ทรัพย์สินจะถึงมือลูกหรือเปล่า?",
    7: "แผนสำรองหากคู่สมรสจากไปด้วย",
    8: "ใครดูแลลูกหากพ่อแม่จากไปพร้อมกัน?",
    9: "คนดูแลลูก ≠ คนดูแลเงินของลูก",
    10: "ลูกยังเล็กเกินสำหรับเงินก้อนใหญ่",
    11: "คู่มือฉุกเฉินสำหรับคนที่รับช่วงต่อ",
    12: "จดหมายถึงครอบครัว",
}


# ── Drive helpers ─────────────────────────────────────────────────────
def find_folder_id(folder_name: str) -> str:
    service = get_drive_service()
    result = service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)"
    ).execute()
    files = result.get("files", [])
    if not files:
        raise ValueError(f"ไม่พบ folder: {folder_name}")
    return files[0]["id"]


def download_xlsx(folder_id: str) -> str:
    service = get_drive_service()
    result = service.files().list(
        q=f"'{folder_id}' in parents and name contains '.xlsx'",
        fields="files(id, name)"
    ).execute()
    files = result.get("files", [])
    if not files:
        raise ValueError("ไม่พบ xlsx ใน folder")
    file_id = files[0]["id"]
    request = service.files().get_media(fileId=file_id)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    downloader = MediaIoBaseDownload(tmp, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    tmp.close()
    print(f"✅ Downloaded xlsx: {files[0]['name']}")
    return tmp.name


# ── Read xlsx ─────────────────────────────────────────────────────────
def read_xlsx(local_path: str) -> tuple[dict, list]:
    wb = load_workbook(local_path, data_only=True)

    # Tab 1 — ข้อมูลรายช่อง
    ws1 = wb["1 เรื่องราว"]
    client_data = {}
    for row in ws1.iter_rows(values_only=True):
        if row[0] and row[1] and row[0] not in ("เรื่องราวลูกค้า", "ข้อมูลรายช่อง"):
            # skip header rows and story paragraphs
            if isinstance(row[0], str) and len(row[0]) < 40:
                client_data[row[0]] = row[1]

    # Tab 2 — 12 ประเด็น
    ws2 = wb["2 วิเคราะห์"]
    issues = []
    for row in ws2.iter_rows(min_row=3, values_only=True):
        if row[0] and isinstance(row[0], int):
            issues.append({
                "ลำดับ": row[0],
                "แนะนำ": row[4] or "",
                "สถานะ": row[5] or "",
                "ความเห็นคุณพยัต": row[6] or "",
            })
    return client_data, issues


# ── Claude generate paragraph 2 + current status ─────────────────────
def generate_issue_content(client_data: dict, issues: list) -> dict:
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    issues_text = "\n".join([
        f"ประเด็นที่ {i['ลำดับ']}: สถานะ={i['สถานะ']} | แนะนำ={i['แนะนำ']} | ความเห็นคุณพยัต={i['ความเห็นคุณพยัต']}"
        for i in issues
    ])

    prompt = f"""คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกฎหมาย

ข้อมูลลูกค้า:
{json.dumps(client_data, ensure_ascii=False, indent=2)}

ข้อมูล 12 ประเด็น:
{issues_text}

สำหรับแต่ละประเด็น 1-12 ให้เขียน 2 ส่วน:

1. ย่อหน้าลูกค้า (2-3 ประโยค): สรุปวิธีจัดการปัญหานี้ของลูกค้าในปัจจุบันที่เกี่ยวกับประเด็นนี้ ใช้ชื่อจริง ตัวเลขจริง เป็นเรื่องราว ไม่ใช่หลักการทั่วไป

2. สิ่งที่แนะนำ (2-3 ประโยค): สรุปจาก "ย่อหน้าลูกค้า" และ "ความเห็นคุณพยัต" ครบทุกประเด็น ให้อ่านง่าย กระชับ ใส่ตัวเลขถ้ามี

output เป็น JSON เท่านั้น ไม่มีข้อความอื่น:
[{{"ลำดับ": 1, "ย่อหน้าลูกค้า": "...", "สิ่งที่แนะน": "..."}}]"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text
    
    # strip markdown fences
    text = text.replace("```json", "").replace("```", "").strip()
    
    print(f"DEBUG cleaned[:300]: {text[:300]}")
    
    start = text.find("[")
    end = text.rfind("]") + 1
    if start == -1 or end == 0:
        print(f"DEBUG no json: {text}")
        raise ValueError("ไม่พบ JSON")
    result = json.loads(text[start:end])
    return {item["ลำดับ"]: item for item in result}


# ── Build docx ────────────────────────────────────────────────────────
def build_cover(client_data: dict, issues: list, generated: dict, folder_name: str) -> str:
    from docx import Document as DocxDocument
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import copy

    FONT = "TH Sarabun New"
    NAVY_RGB = RGBColor(0x1A, 0x3A, 0x5C)
    LIGHT_BLUE_HEX = "EBF2FA"
    LIGHT_GRAY_HEX = "F5F5F5"
    NAVY_HEX = "1A3A5C"
    DARK_HEX = "111111"
    YELLOW_HEX = "FFFBEA"

    def set_cell_bg(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), hex_color)
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)

    def add_run(para, text, bold=False, size=14, color=None, font=FONT):
        run = para.add_run(text)
        run.font.name = font
        run.font.size = Pt(size)
        run.bold = bold
        if color:
            run.font.color.rgb = color
        return run

    doc = DocxDocument()

    # Page margins
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.79)
        section.right_margin = Inches(0.79)

    # ── Header ────────────────────────────────────────────────────────
    nickname = client_data.get("ชื่อเล่น", folder_name.split("_")[0])
    date_str = folder_name.split("_")[1] if "_" in folder_name else ""

    header_table = doc.add_table(rows=1, cols=1)
    header_table.style = "Table Grid"
    cell = header_table.cell(0, 0)
    set_cell_bg(cell, NAVY_HEX)

    p1 = cell.paragraphs[0]
    add_run(p1, "แผนคุ้มครองครอบครัว", bold=True, size=28, color=RGBColor(0xFF, 0xFF, 0xFF))
    p1.paragraph_format.space_before = Pt(8)
    p1.paragraph_format.space_after = Pt(4)

    p2 = cell.add_paragraph()
    add_run(p2, f"คุณ{nickname}", size=18, color=RGBColor(0xBB, 0xDD, 0xFF))
    p2.paragraph_format.space_after = Pt(2)

    p3 = cell.add_paragraph()
    add_run(p3, f"จัดทำโดย คุณพยัต  |  {date_str[:2]}/{date_str[2:4]}/{date_str[4:]}", size=13, color=RGBColor(0x88, 0xAA, 0xCC))
    p3.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()

    # ── Intro ─────────────────────────────────────────────────────────
    intro_texts = [
        "การทำพินัยกรรมนั้น ถ้ามองแค่วิธีการจัดทำนั้น ไม่ยากเลย เพียงแค่เขียนข้อความบอกว่าอยากให้ทรัพย์สินไปอยู่ที่ไหน ลงวันที่แล้วเซ็นชื่อต่อหน้าพยาน 2 คน เท่านี้ก็ถูกต้องตามกฎหมายทุกอย่าง ทายาทสามารถนำไปให้ศาลออกคำสั่งต่างๆได้แล้ว",
        f"แต่สำหรับพ่อแม่ที่มีลูกเล็ก เรื่องนี้ต้องคิดให้ไกลกว่านั้น เพราะสิ่งที่คุณต้องการจริงๆ ไม่ใช่แค่เอกสารที่ถูกกฎหมาย แต่มันคือแผน ที่ทำให้คุณมั่นใจว่าถ้าวันหนึ่งเกิดเหตุไม่คาดฝัน คู่ชีวิตจะไม่บำลาก ลูกๆจะยังมีคนดูแล มีเงินพอใช้ และได้รับสิ่งที่คุณตั้งใจให้จริงๆ",
        "เอกสารชุดนี้จึงครอบคลุม 12 ประเด็นที่พ่อแม่ลูกเล็กทุกคู่ควรพิจารณา ดังต่อไปนี้",
    ]
    for i, text in enumerate(intro_texts):
        p = doc.add_paragraph()
        run = add_run(p, text, size=15, color=RGBColor(0x11, 0x11, 0x11) if i < 2 else NAVY_RGB, bold=(i == 2))
        p.paragraph_format.space_after = Pt(8)

    doc.add_paragraph()

    # ── 12 Issues ─────────────────────────────────────────────────────
    issue_map = {i["ลำดับ"]: i for i in issues}

    for num in range(1, 13):
        gen = generated.get(num, {})
        issue = issue_map.get(num, {})
        title = ISSUE_TITLES.get(num, "")
        principle = GENERAL_PRINCIPLES.get(num, "")
        client_para = gen.get("ย่อหน้าลูกค้า", "")
        current = gen.get("การจัดการในปัจจุบัน", "")

        # Header row
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        tbl.columns[0].width = Inches(0.5)
        tbl.columns[1].width = Inches(6.5)

        num_cell = tbl.cell(0, 0)
        set_cell_bg(num_cell, NAVY_HEX)
        np = num_cell.paragraphs[0]
        np.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(np, f"{num:02d}", bold=True, size=16, color=RGBColor(0xFF, 0xFF, 0xFF))

        title_cell = tbl.cell(0, 1)
        set_cell_bg(title_cell, LIGHT_BLUE_HEX)
        tp = title_cell.paragraphs[0]
        add_run(tp, title, bold=True, size=16, color=NAVY_RGB)

        # Paragraph 1 — หลักการทั่วไป
        p = doc.add_paragraph()
        add_run(p, principle, size=15, color=RGBColor(0x44, 0x44, 0x44))
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)

        # Paragraph 2 — เรื่องลูกค้า
        if client_para:
            p2 = doc.add_paragraph()
            add_run(p2, client_para, size=15, color=RGBColor(0x22, 0x22, 0x22))
            p2.paragraph_format.space_after = Pt(4)

        # กล่อง การจัดการในปัจจุบัน
        if current:
            box = doc.add_table(rows=1, cols=1)
            box.style = "Table Grid"
            box_cell = box.cell(0, 0)
            set_cell_bg(box_cell, LIGHT_GRAY_HEX)

            bp1 = box_cell.paragraphs[0]
            add_run(bp1, "การจัดการในปัจจุบัน", bold=True, size=14, color=NAVY_RGB)
            bp1.paragraph_format.space_before = Pt(4)

            bp2 = box_cell.add_paragraph()
            add_run(bp2, current, size=14, color=RGBColor(0x33, 0x33, 0x33))
            bp2.paragraph_format.space_after = Pt(4)

        doc.add_paragraph()

    # ── Disclaimer ────────────────────────────────────────────────────
    dis_tbl = doc.add_table(rows=1, cols=1)
    dis_tbl.style = "Table Grid"
    dis_cell = dis_tbl.cell(0, 0)
    set_cell_bg(dis_cell, YELLOW_HEX)
    dp = dis_cell.paragraphs[0]
    add_run(dp, "⚠️  หมายเหตุ: ", bold=True, size=13, color=RGBColor(0x7A, 0x58, 0x00))
    add_run(dp, "คำแนะนำด้านประกันชีวิตเป็นการประเมินเบื้องต้นเพื่อให้เห็นภาพรวมเท่านั้น หากต้องการวางแผนที่เหมาะกับสถานการณ์จริง กรุณาผมติดต่อเพื่อจัดทำแผนประกันภัยเฉพาะบุคคลแยกต่างหาก", size=13, color=RGBColor(0x7A, 0x58, 0x00))

    # Save
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    print(f"✅ Built cover page")
    return tmp.name


# ── Main ──────────────────────────────────────────────────────────────
def run(folder_name: str):
    print(f"น้องดอค เริ่มทำงาน: {folder_name}")

    folder_id = find_folder_id(folder_name)
    print(f"✅ Found folder: {folder_id}")

    xlsx_path = download_xlsx(folder_id)
    client_data, issues = read_xlsx(xlsx_path)
    if os.path.exists(xlsx_path):
        os.unlink(xlsx_path)
    print(f"✅ อ่าน xlsx เรียบร้อย: {len(issues)} ประเด็น")

    print("กำลัง generate เนื้อหา...")
    generated = generate_issue_content(client_data, issues)
    print(f"✅ Generate เสร็จ")

    nickname = client_data.get("ชื่อเล่น", folder_name.split("_")[0])
    cover_path = build_cover(client_data, issues, generated, folder_name)
    cover_filename = f"ใบปะหน้า_{nickname}.docx"

    upload_file_to_folder(cover_path, cover_filename, folder_id)
    if os.path.exists(cover_path):
        os.unlink(cover_path)

    send_line_message(f"✅ เอกสารพร้อมแล้วครับ\n📁 {folder_name}\n📄 {cover_filename}")
    print("เสร็จสิ้น")
