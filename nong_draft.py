import os
import json
import tempfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from drive_service import upload_file
from email_service import notify_new_client

DEV_MODE = True  # เปลี่ยนเป็น False ตอน prod

SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# ── helpers ──────────────────────────────────────────────────────────
def fill(h): return PatternFill("solid", start_color=h, end_color=h)
def fnt(bold=False, size=10, color="222222"):
    return Font(name="Arial", bold=bold, size=size, color=color)
def thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)
def cal(h="left"):
    return Alignment(horizontal=h, vertical="center", wrap_text=True)
def put(ws, row, col, val="", bg="FFFFFF", bold=False, color="222222", size=10, align="left"):
    c = ws.cell(row=row, column=col, value=val)
    c.fill = fill(bg); c.font = fnt(bold, size, color)
    c.alignment = cal("center" if align == "center" else "left")
    c.border = thin()
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
    # pad row ให้ยาวเท่า headers
    row += [""] * (len(headers) - len(row))
    return dict(zip(headers, row))


# ── build xlsx เล็กๆ (dev) ───────────────────────────────────────────
def build_workbook_dev(data: dict) -> str:
    wb = Workbook()

    # ── Tab 1: เรื่องราวลูกค้า ──
    ws1 = wb.active
    ws1.title = "1 เรื่องราว"
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions["A"].width = 24
    ws1.column_dimensions["B"].width = 48

    # title
    ws1.row_dimensions[1].height = 30
    c = ws1.cell(row=1, column=1, value=f"เรื่องราวลูกค้า — {data.get('nickname', '')} | {data.get('occupation', '')} | {data.get('age', '')} ปี")
    c.fill = fill("1E3A5F"); c.font = fnt(bold=True, size=13, color="FFFFFF")
    c.alignment = cal(); c.border = thin()
    ws1.merge_cells("A1:B1")

    fields = [
        ("ชื่อเล่น", "nickname"), ("อายุ", "age"), ("เพศ", "gender"),
        ("อาชีพ", "occupation"), ("สุขภาพ", "health"), ("อีเมล", "email"),
        ("คู่สมรส", "spouse_nickname"), ("อาชีพคู่สมรส", "spouse_occupation"),
        ("ลูก", "children"), ("ทรัพย์สิน (เงินสด)", "assets_cash"),
        ("อสังหาริมทรัพย์", "assets_property"), ("หนี้สิน", "debt"),
        ("ประกันชีวิต", "insurance_life"), ("ผู้จัดการมรดก", "estate_executor"),
        ("Guardian หลัก", "guardian_primary"), ("Living Will", "living_will"),
    ]
    for i, (label, key) in enumerate(fields, 2):
        ws1.row_dimensions[i].height = 28
        put(ws1, i, 1, label, bg="F5F7FA", bold=True, color="444444", size=9)
        put(ws1, i, 2, data.get(key, "ไม่ได้ระบุ"), size=10)

    # ── Tab 2: วิเคราะห์ (ช่องให้คุณพยัตกรอก) ──
    ws2 = wb.create_sheet("2 วิเคราะห์")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 24
    ws2.column_dimensions["C"].width = 36
    ws2.column_dimensions["D"].width = 30
    ws2.column_dimensions["E"].width = 16

    ws2.row_dimensions[1].height = 30
    c = ws2.cell(row=1, column=1, value="วิเคราะห์ & ความเห็นคุณพยัต")
    c.fill = fill("2E6DA4"); c.font = fnt(bold=True, size=13, color="FFFFFF")
    c.alignment = cal(); c.border = thin()
    ws2.merge_cells("A1:E1")

    ws2.row_dimensions[2].height = 22
    for col, h in [(1,"#"),(2,"ประเด็น"),(3,"AI วิเคราะห์"),(4,"ความเห็นคุณพยัต"),(5,"Approve")]:
        c = ws2.cell(row=2, column=col, value=h)
        c.fill = fill("2E6DA4"); c.font = fnt(bold=True, size=9, color="FFFFFF")
        c.alignment = cal("center"); c.border = thin()

    issues = [
        ("ประกันชีวิต", f"{'ไม่มี' if not data.get('insurance_life') or data.get('insurance_life') == 'ไม่มี' else 'มี'} — ตรวจสอบความเพียงพอ"),
        ("ผู้รับมรดก", f"Executor: {data.get('estate_executor', 'ไม่ระบุ')}"),
        ("Guardian ลูก", f"{data.get('guardian_primary', 'ไม่ระบุ')} — ตรวจสอบตัวสำรอง"),
        ("Living Will", f"{data.get('living_will', 'ไม่ระบุ')}"),
        ("หนี้สิน", f"{data.get('debt', 'ไม่ระบุ')} — มีความคุ้มครองไหม"),
    ]
    for i, (issue, analysis) in enumerate(issues, 3):
        ws2.row_dimensions[i].height = 40
        bg = "FFFFFF" if i % 2 == 1 else "F5F7FA"
        put(ws2, i, 1, i-2, bg=bg, align="center", color="888888", size=9)
        put(ws2, i, 2, issue, bg=bg, bold=True, color="1E3A5F", size=10)
        put(ws2, i, 3, analysis, bg=bg, color="444444", size=9)
        put(ws2, i, 4, "", bg="FFFDE7", size=10)   # ช่องคุณพยัตกรอก
        put(ws2, i, 5, "", bg="FFFFFF", align="center", size=11)  # approve

    # ── Tab 3: placeholder ──
    for title in ["3 Gap", "4 พินัยกรรม", "5 POA", "6 Living Will", "7 คู่มือฉุกเฉิน"]:
        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        c = ws.cell(row=1, column=1, value=f"{title} — จะ generate เมื่อคุณพยัต approve tab 2")
        c.fill = fill("E8ECF0"); c.font = fnt(bold=True, size=11, color="555555")
        c.alignment = cal(); c.border = thin()
        ws.column_dimensions["A"].width = 60

    # save to temp file
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

    print(f"📋 ข้อมูล: {data.get('nickname')} | {data.get('occupation')} | {data.get('age')} ปี")

    # สร้างไฟล์
    local_path, filename = build_workbook_dev(data)
    print(f"✅ สร้างไฟล์: {filename}")

    # upload Drive
    upload_file(local_path, filename)

    # แจ้ง LINE
    notify_new_client(
        nickname=data.get("nickname", ""),
        occupation=data.get("occupation", ""),
        age=data.get("age", "")
    )

    # ลบไฟล์ temp
    os.unlink(local_path)
    print("🎉 เสร็จสิ้น")


if __name__ == "__main__":
    run()
