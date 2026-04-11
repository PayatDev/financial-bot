import os
import json
import tempfile
from datetime import datetime

import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from drive_service import upload_file
from email_service import notify_new_client

DEV_MODE = True  # เปลี่ยนเป็น False ตอน prod

SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

NAVY  = "1E3A5F"
BLUE  = "2E6DA4"
LGRAY = "F5F7FA"
MGRAY = "E8ECF0"
LYELL = "FFFDE7"
WHITE = "FFFFFF"
GREEN = "1A7A4A"
RED   = "C0392B"
AMBER = "D35400"

# ── helpers ──────────────────────────────────────────────────────────
def fill(h): return PatternFill("solid", start_color=h, end_color=h)
def fnt(bold=False, size=10, color="222222", italic=False):
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
def thin(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def cal(h="left"):
    return Alignment(horizontal=h, vertical="center", wrap_text=True)
def put(ws, row, col, val="", bg=WHITE, bold=False, color="222222",
        size=10, align="left", italic=False, bc="CCCCCC"):
    c = ws.cell(row=row, column=col, value=val)
    c.fill = fill(bg); c.font = fnt(bold, size, color, italic)
    c.alignment = cal("center" if align == "center" else "left")
    c.border = thin(bc)
    return c


# ── ดึงข้อมูลจาก Sheets ──────────────────────────────────────────────
def get_data_from_sheets() -> dict:
    sa_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(sa_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range="Sheet1!A1:BZ2"
    ).execute()
    values = result.get("values", [])
    if len(values) < 2:
        return {}
    headers = values[0]
    row = values[1]
    row += [""] * (len(headers) - len(row))
    return dict(zip(headers, row))


# ── เรียก Claude API ──────────────────────────────────────────────────
def call_claude(data: dict) -> tuple[str, list]:
    prompt = f"""คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกำลังเรียนกฎหมาย

รับข้อมูลลูกค้า แล้ว output 2 ส่วน คั่นด้วย ---SPLIT---

===== ส่วนที่ 1 =====

เขียน 2 blocks ต่อเนื่องกัน

Block A — เรื่องราวลูกค้า
เล่าเรื่องต่อเนื่อง 3-5 ย่อหน้า ภาษาไทยธรรมดา อ่านเข้าใจง่าย
ครอบคลุมทุก field ที่มีข้อมูล ข้ามข้อมูลที่เป็น "ไม่มี" หรือ "ไม่ได้ระบุ"

Block B — ข้อมูลรายช่อง
ขึ้นบรรทัดใหม่หลัง Block A
แสดงทุก field ในรูปแบบ "ชื่อ: ค่า" ทีละบรรทัด

---SPLIT---

===== ส่วนที่ 2 =====

ประเมิน 16 ประเด็นต่อไปนี้ตามลำดับนี้เท่านั้น
output เป็น JSON array เท่านั้น ไม่มีข้อความอื่น ไม่มี markdown

ประเด็น (ตามลำดับ):
1. เงินฉุกเฉิน 90 วัน
2. เงินสดสำรองคลินิก 6-12 เดือน
3. ประกันชีวิตคุ้มครอง
4. ประกันสุขภาพ / ทุพพลภาพ
5. ประกันอุบัติเหตุ
6. ที่ดินมรดก — ระบุผู้รับใน Will
7. คู่สมรสแต่งงานใหม่
8. I Love You Will — กรณีคู่สมรสตายก่อนหรือพร้อมกัน
9. Business Succession
10. ส่วนแบ่งประกันหนี้ส่วนเกิน
11. Guardian of Person
12. Money Guardian แยกจาก Guardian
13. Money Guardian สำรอง
14. ที่อยู่เอกสารสำคัญ
15. ขั้นตอนแรกหลังเกิดเหตุ
16. จดหมายถึงคู่สมรส

format:
[
  {{
    "ลำดับ": 1,
    "ประเด็น": "ชื่อตามด้านบน",
    "หมวด": "สภาพคล่อง หรือ ประกัน หรือ พินัยกรรม หรือ ผู้ปกครอง หรือ คู่มือฉุกเฉิน",
    "สถานะ": "❌ ขาด หรือ ⚠️ บางส่วน หรือ ✅ ครบ",
    "ความเสี่ยง": "ถ้าปล่อยไว้จะเกิดอะไร ไม่เกิน 2 ประโยค ภาษาธรรมดา",
    "แนะนำ": "ควรทำอะไร ไม่เกิน 1 ประโยค"
  }}
]

กฎ:
- ไม่มีประกันชีวิตคุ้มครอง → ❌ ขาด
- ไม่มีเงินฉุกเฉิน 90 วัน → ❌ ขาด
- ไม่มี Contingency Clause → ❌ ขาด
- Money Guardian = Guardian คนเดียวกัน → ⚠️ บางส่วน
- กังวลคู่สมรสแต่งใหม่แต่ไม่มีแผน → ⚠️ บางส่วน
- ไม่มี Money Guardian สำรอง → ⚠️ บางส่วน
- field ไม่มีข้อมูล → ❌ ขาด
- ประเมินจากข้อมูลจริงเท่านั้น

ข้อมูลลูกค้า:
{json.dumps(data, ensure_ascii=False, indent=2)}"""

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    result = response.content[0].text

    if "---SPLIT---" not in result:
        print(f"DEBUG full response: {result[:500]}")  # เพิ่ม
        raise ValueError("ไม่พบ ---SPLIT--- ใน response")

    part1, part2 = result.split("---SPLIT---", 1)
    part2 = part2.strip()

    # หา JSON array
    start = part2.find("[")
    end = part2.rfind("]") + 1
    if start == -1 or end == 0:
        raise ValueError("ไม่พบ JSON array ใน part2")

    issues = json.loads(part2[start:end])
    return part1.strip(), issues


# ── build xlsx ────────────────────────────────────────────────────────
def build_workbook(data: dict, story: str, issues: list) -> tuple[str, str]:
    wb = Workbook()

    # ── Tab 1: เรื่องราว + ข้อมูลรายช่อง ──
    ws1 = wb.active
    ws1.title = "1 เรื่องราว"
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions["A"].width = 26
    ws1.column_dimensions["B"].width = 54

    # title
    ws1.row_dimensions[1].height = 32
    c = ws1.cell(row=1, column=1,
        value=f"{data.get('nickname','')} | {data.get('occupation','')} | {data.get('age','')} ปี")
    c.fill = fill(NAVY); c.font = fnt(bold=True, size=13, color=WHITE)
    c.alignment = cal(); c.border = thin(NAVY)
    ws1.merge_cells("A1:B1")

    # แยก Block A และ Block B จาก story
    blocks = story.split("Block B")
    block_a = blocks[0].replace("Block A —", "").replace("Block A—", "").strip()
    block_b = blocks[1].strip() if len(blocks) > 1 else ""

    # Block A header
    row = 2
    ws1.row_dimensions[row].height = 18
    c = ws1.cell(row=row, column=1, value="เรื่องราวลูกค้า")
    c.fill = fill(MGRAY); c.font = fnt(bold=True, size=10, color=NAVY)
    c.alignment = cal(); c.border = thin(MGRAY)
    ws1.merge_cells(f"A{row}:B{row}")
    row += 1

    # Block A content
    for para in block_a.split("\n"):
        para = para.strip()
        if not para:
            continue
        ws1.row_dimensions[row].height = 60
        put(ws1, row, 1, "", bg=LGRAY)
        c = ws1.cell(row=row, column=2, value=para)
        c.fill = fill(LGRAY); c.font = fnt(size=10)
        c.alignment = cal(); c.border = thin()
        ws1.merge_cells(f"A{row}:A{row}")
        row += 1

    # Block B header
    ws1.row_dimensions[row].height = 18
    c = ws1.cell(row=row, column=1, value="ข้อมูลรายช่อง")
    c.fill = fill(MGRAY); c.font = fnt(bold=True, size=10, color=NAVY)
    c.alignment = cal(); c.border = thin(MGRAY)
    ws1.merge_cells(f"A{row}:B{row}")
    row += 1

    # Block B content
    for line in block_b.split("\n"):
        line = line.strip().lstrip("—").strip()
        if not line or "Block B" in line:
            continue
        if ": " in line:
            label, _, value = line.partition(": ")
        else:
            label, value = line, ""
        ws1.row_dimensions[row].height = 28
        put(ws1, row, 1, label.strip(), bg=LGRAY, bold=True, color="444444", size=9)
        put(ws1, row, 2, value.strip(), size=10)
        row += 1

    # ── Tab 2: วิเคราะห์ ──
    ws2 = wb.create_sheet("2 วิเคราะห์")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 28
    ws2.column_dimensions["D"].width = 36
    ws2.column_dimensions["E"].width = 30
    ws2.column_dimensions["F"].width = 14
    ws2.column_dimensions["G"].width = 28
    ws2.column_dimensions["H"].width = 14

    ws2.row_dimensions[1].height = 32
    c = ws2.cell(row=1, column=1, value="วิเคราะห์ช่องว่าง")
    c.fill = fill(NAVY); c.font = fnt(bold=True, size=13, color=WHITE)
    c.alignment = cal(); c.border = thin(NAVY)
    ws2.merge_cells("A1:H1")

    ws2.row_dimensions[2].height = 24
    for col, h in [(1,"#"),(2,"หมวด"),(3,"ประเด็น"),(4,"ความเสี่ยง"),(5,"แนะนำ"),(6,"สถานะ"),(7,"ความเห็นคุณพยัต"),(8,"Approve")]:
        c = ws2.cell(row=2, column=col, value=h)
        c.fill = fill(BLUE); c.font = fnt(bold=True, size=9, color=WHITE)
        c.alignment = cal("center"); c.border = thin(WHITE)

    STATUS_STYLE = {
        "❌ ขาด":      ("FDECEA", RED),
        "⚠️ บางส่วน": ("FEF3E8", AMBER),
        "✅ ครบ":      ("E8F5EE", GREEN),
    }

    for i, item in enumerate(issues, 3):
        ws2.row_dimensions[i].height = 50
        bg = WHITE if i % 2 == 1 else LGRAY
        status = item.get("สถานะ", "")
        st_bg, st_fg = STATUS_STYLE.get(status, (LGRAY, NAVY))

        put(ws2, i, 1, item.get("ลำดับ", ""), bg=bg, align="center", color="888888", size=9)
        put(ws2, i, 2, item.get("หมวด", ""), bg=bg, bold=True, color=NAVY, size=9)
        put(ws2, i, 3, item.get("ประเด็น", ""), bg=bg, bold=True, color=NAVY, size=10)
        put(ws2, i, 4, item.get("ความเสี่ยง", ""), bg=bg, color="444444", size=9, italic=True)
        put(ws2, i, 5, item.get("แนะนำ", ""), bg=bg, bold=True, color=GREEN, size=10)

        # สถานะ
        c = ws2.cell(row=i, column=6, value=status)
        c.fill = fill(st_bg); c.font = fnt(bold=True, size=10, color=st_fg)
        c.alignment = cal("center"); c.border = thin(st_fg)

        put(ws2, i, 7, "", bg=LYELL, size=10)   # ช่องคุณพยัต
        put(ws2, i, 8, "", bg=WHITE, align="center", size=12)  # approve

    ws2.freeze_panes = "A3"
    ws1.freeze_panes = "A2"

    # save
    nickname = data.get("nickname", "ลูกค้า")
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{nickname}_{date_str}.xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return tmp.name, filename


# ── main ─────────────────────────────────────────────────────────────
def run(data: dict = None):
    if DEV_MODE or data is None:
        print("🔧 DEV MODE — ดึงข้อมูลจาก Sheets")
        data = get_data_from_sheets()

    print(f"📋 {data.get('nickname')} | {data.get('occupation')} | {data.get('age')} ปี")

    # เรียก Claude
    print("🤖 กำลังวิเคราะห์...")
    story, issues = call_claude(data)
    print(f"✅ ได้ {len(issues)} ประเด็น")

    # สร้างไฟล์
    local_path, filename = build_workbook(data, story, issues)
    print(f"✅ สร้างไฟล์: {filename}")

    # upload Drive
    upload_file(local_path, filename)

    # แจ้ง LINE
    notify_new_client(
        nickname=data.get("nickname", ""),
        occupation=data.get("occupation", ""),
        age=data.get("age", "")
    )

    os.unlink(local_path)
    print("🎉 เสร็จสิ้น")
