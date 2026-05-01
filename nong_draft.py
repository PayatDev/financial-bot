import os
import json
import tempfile
from datetime import datetime

import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

from drive_service import create_folder, upload_file_to_folder
from nong_doc import run as doc_run

DEV_MODE = False
SHEET_ID  = os.environ.get("GOOGLE_SHEET_ID")
SCOPES    = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
MODEL     = "claude-sonnet-4-20250514"

# ── colors ─────────────────────────────────────────────────────────────
NAVY  = "1E3A5F"
BLUE  = "2E6DA4"
LGRAY = "F5F7FA"
MGRAY = "E8ECF0"
WHITE = "FFFFFF"
GREEN = "1A7A4A"
RED   = "C0392B"
AMBER = "D35400"

# ── 12 ประเด็น fixed ───────────────────────────────────────────────────
FIXED_ISSUES = [
    {"ลำดับ": 1,  "ประเด็น": "ค่าใช้จัดการเรื่องหลังเสียชีวิต",
     "อธิบาย": "ทันทีที่เราจากไป จะมีค่าใช้จ่ายเกิดขึ้นทันที ทั้งค่ารักษาพยาบาลครั้งสุดท้าย ค่าจัดงานศพ พิธีทางศาสนา รวมถึงการติดต่อหน่วยงานต่างๆ",
     "กรณีไม่จัดการ": "ผู้จัดการเรื่องต้องจ่ายเงินทดรองไปก่อน อาจกลายเป็นภาระให้คนที่รักในช่วงเวลาที่เจ็บปวดที่สุด"},
    {"ลำดับ": 2,  "ประเด็น": "เงินสดปรับตัว 6-12 เดือน และจัดการหนี้",
     "อธิบาย": "การสูญเสียไม่ได้กระทบแค่ความรู้สึก แต่กระทบวิถีชีวิตทั้งหมด ครอบครัวอาจต้องย้ายบ้าน ลูกอาจต้องเปลี่ยนโรงเรียน ช่วงเปลี่ยนผ่านนี้ต้องใช้เงินก้อนหนึ่ง",
     "กรณีไม่จัดการ": "คู่ชีวิตและลูกๆ ต้องเผชิญกับการเปลี่ยนแปลงครั้งใหญ่ในขณะที่จิตใจยังไม่พร้อมรับมือ"},
    {"ลำดับ": 3,  "ประเด็น": "ประกันชีวิตคุ้มครองรายได้และมูลค่าทางเศรษฐกิจตัวเอง",
     "อธิบาย": "เมื่อเราจากไป รายได้ที่เราจะหาได้ตลอดช่วงชีวิตที่เหลือก็หายไปด้วย ประกันชีวิตที่ดีควรเข้ามาชดเชยรายได้ส่วนนี้ให้ครอบครัว",
     "กรณีไม่จัดการ": "คู่ชีวิตจะต้องแบกภาระหาเลี้ยงครอบครัวเพียงคนเดียว คุณภาพชีวิตจะลดลงในระยะยาว"},
    {"ลำดับ": 4,  "ประเด็น": "Business Succession และเงินหมุนเวียนในกิจการ",
     "อธิบาย": "การบริหารกิจการมักต้องอาศัยลายเซ็นหรืออำนาจของผู้บริหาร หากผู้บริหารจากไปกะทันหัน คนที่เข้ามาดูแลต่อต้องมีอำนาจตามกฎหมาย",
     "กรณีไม่จัดการ": "กิจการอาจหยุดชะงัก อนุมัติงานไม่ได้ ขาดสภาพคล่องในช่วงปรับตัว"},
    {"ลำดับ": 5,  "ประเด็น": "ประกันโรคร้าย / ทุพพลภาพ",
     "อธิบาย": "สิ่งที่น่ากลัวกว่าความตายคือการพิการ เพราะนอกจากหาเงินไม่ได้ ยังมีค่าดูแลรักษาเพิ่มขึ้น และประกันชีวิตยังไม่จ่าย",
     "กรณีไม่จัดการ": "คู่ชีวิตต้องแบกรับทุกอย่าง ทั้งหาเงิน เลี้ยงลูก และดูแลเราในฐานะผู้พิการ"},
    {"ลำดับ": 6,  "ประเด็น": "คู่สมรสแต่งงานใหม่ I Love You Will",
     "อธิบาย": "หากทำพินัยกรรมมอบทุกอย่างให้คู่ชีวิตโดยไม่มีเงื่อนไข สิ่งที่หวังว่าจะตกทอดไปถึงลูกอาจไม่เป็นตามที่คิด",
     "กรณีไม่จัดการ": "ทรัพย์สินที่สร้างมาทั้งชีวิตอาจไหลไปสู่ครอบครัวใหม่ของคู่ชีวิต"},
    {"ลำดับ": 7,  "ประเด็น": "กรณีคู่สมรสตายก่อนหรือพร้อมกัน Contingency Clause",
     "อธิบาย": "หากคู่ชีวิตจากไปก่อนหรือพร้อมกัน มรดกอาจไหลไปตามกฎหมายโดยที่ลูกได้ไม่เต็มเม็ดเต็มหน่วย",
     "กรณีไม่จัดการ": "ปู่ย่าตายายมีสิทธิรับมรดกในลำดับเดียวกับลูก มรดกอาจถูกแบ่งออกไปโดยไม่ตั้งใจ"},
    {"ลำดับ": 8,  "ประเด็น": "Guardian of Person (และสำรอง)",
     "อธิบาย": "กรณีพ่อแม่จากไปพร้อมกัน ลูกจะไปอยู่ที่ไหน ใครจะดูแล และดูแลอย่างที่เราต้องการจริงไหม",
     "กรณีไม่จัดการ": "อาจเกิดการแย่งชิงตัวลูก หรือต่างคนต่างเกี่ยงกันดูแล"},
    {"ลำดับ": 9,  "ประเด็น": "Money Guardian แยกจาก Guardian",
     "อธิบาย": "การแยกคนดูแลลูกออกจากคนคุมเงินของลูกเป็นกลไกสำคัญที่ช่วยตรวจสอบซึ่งกันและกัน",
     "กรณีไม่จัดการ": "หากผู้ปกครองและผู้คุมเงินเป็นคนเดียวกัน ไม่มีใครตรวจสอบว่าเงินถูกใช้เพื่อลูกจริงๆ"},
    {"ลำดับ": 10, "ประเด็น": "เอกสารอยู่ที่ไหนและใครรู้บ้าง",
     "อธิบาย": "เมื่อเกิดเหตุ คนที่รับหน้าที่จัดการจำเป็นต้องรู้ว่าเอกสารอยู่ที่ไหน และต้องทำอะไรก่อนหลัง",
     "กรณีไม่จัดการ": "ผู้จัดการมรดกต้องตามหาเอกสารเองในช่วงที่เจ็บปวดที่สุด อาจตกหล่นสิทธิ์สำคัญ"},
    {"ลำดับ": 11, "ประเด็น": "ทบทวนแผนทุกปี",
     "อธิบาย": "ชีวิตเปลี่ยนตลอดเวลา อาจมีลูกเพิ่ม รายได้เปลี่ยน ผู้ดูแลที่เลือกไว้อาจเสียชีวิตก่อน แผนที่ดีคือแผนที่ปรับให้ทันสถานการณ์",
     "กรณีไม่จัดการ": "พินัยกรรมที่ไม่เคยทบทวนมักเป็นพินัยกรรมที่สร้างปัญหาแทนที่แก้ปัญหา"},
    {"ลำดับ": 12, "ประเด็น": "จดหมายถึงลูกและคู่สมรส",
     "อธิบาย": "จดหมายถึงครอบครัวเป็นสิ่งที่ไม่มีเอกสารทางกฎหมายใดทดแทนได้ ครอบครัวจะได้รับรู้ความรักและเหตุผลเบื้องหลังการตัดสินใจต่างๆ",
     "กรณีไม่จัดการ": "ครอบครัวอาจเข้าใจผิดในสิ่งที่ตัดสินใจ และเสียโอกาสที่จะได้รับสิ่งที่มีค่าที่สุด"},
]

FIELD_LABELS = [
    ("nickname","ชื่อเล่น"),("age","อายุ"),("gender","เพศ"),
    ("occupation","อาชีพ"),("health","สุขภาพ"),("email","อีเมล"),
    ("income_self","รายได้"),("hobbies_and_risks","งานอดิเรก"),
    ("spouse_nickname","คู่สมรส"),("spouse_age","อายุคู่สมรส"),
    ("spouse_occupation","อาชีพคู่สมรส"),("spouse_income","รายได้คู่สมรส"),
    ("spouse_health","สุขภาพคู่สมรส"),("spouse_status","สถานะ"),
    ("children","ลูก"),("children_outside_marriage","บุตรนอกสมรส"),
    ("assets_cash","เงินสด"),("assets_property","อสังหาริมทรัพย์"),
    ("assets_investment","การลงทุน"),("assets_crypto_wallet","คริปโต"),
    ("assets_insurance_savings","ประกันสะสม"),("assets_digital","ทรัพย์สินดิจิทัล"),
    ("assets_business","กิจการ"),("assets_valuables","ของมีค่า"),
    ("debt","หนี้สิน"),("guarantor","ค้ำประกัน"),
    ("insurance_life","ประกันชีวิต"),("insurance_health","ประกันสุขภาพ"),
    ("insurance_group","ประกันกลุ่ม"),("welfare","สวัสดิการ"),
    ("funeral_wishes","ความปรารถนางานศพ"),("emergency_cash_90days","เงินฉุกเฉิน 90 วัน"),
    ("estate_admin_cost","ต้นทุนจัดการมรดก"),("asset_distribution","แผนแบ่งทรัพย์"),
    ("debt_responsibility","หนี้ใครรับผิดชอบ"),("business_succession","แผนกิจการ"),
    ("urgent_manager","ผู้จัดการฉุกเฉิน"),("estate_executor","ผู้จัดการมรดก"),
    ("financial_poa","Financial POA"),("living_will","Living Will"),
    ("surviving_spouse_plan","แผนคู่สมรสที่รอดชีวิต"),
    ("guardian_primary","Guardian หลัก"),("guardian_backup","Guardian สำรอง"),
    ("money_guardian_primary","Money Guardian หลัก"),("money_guardian_backup","Money Guardian สำรอง"),
    ("documents_location","ที่อยู่เอกสาร"),
    ("letter_to_children","จดหมายถึงลูก"),("letter_to_spouse","จดหมายถึงคู่สมรส"),
    ("fullname_self","ชื่อ-นามสกุลจริงเจ้าของแผน"),("id_self","เลขบัตรประชาชนเจ้าของแผน"),
    ("address_self","ที่อยู่ปัจจุบัน"),
    ("fullname_spouse","ชื่อ-นามสกุลจริงคู่สมรส"),("id_spouse","เลขบัตรประชาชนคู่สมรส"),
    ("fullname_children","ชื่อ-นามสกุลจริงลูก (ทุกคน)"),
    ("fullname_executor","ชื่อ-นามสกุลจริงผู้จัดการมรดก"),("id_executor","เลขบัตรประชาชนผู้จัดการมรดก"),
    ("fullname_executor_backup","ชื่อ-นามสกุลจริงผู้จัดการมรดกสำรอง"),("id_executor_backup","เลขบัตรประชาชนผู้จัดการมรดกสำรอง"),
    ("fullname_guardian_primary","ชื่อ-นามสกุลจริงผู้ปกครองหลัก"),("id_guardian_primary","เลขบัตรประชาชนผู้ปกครองหลัก"),
    ("fullname_guardian_backup","ชื่อ-นามสกุลจริงผู้ปกครองสำรอง"),("id_guardian_backup","เลขบัตรประชาชนผู้ปกครองสำรอง"),
    ("fullname_money_guardian_primary","ชื่อ-นามสกุลจริงผู้ดูแลเงินหลัก"),("id_money_guardian_primary","เลขบัตรประชาชนผู้ดูแลเงินหลัก"),
    ("fullname_money_guardian_backup","ชื่อ-นามสกุลจริงผู้ดูแลเงินสำรอง"),("id_money_guardian_backup","เลขบัตรประชาชนผู้ดูแลเงินสำรอง"),
]

# ── openpyxl helpers ───────────────────────────────────────────────────
def fill(h):    return PatternFill("solid", start_color=h, end_color=h)
def fnt(bold=False, size=10, color="222222", italic=False):
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
def thin(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def cal(h="left"):  return Alignment(horizontal=h, vertical="center", wrap_text=True)
def put(ws, row, col, val="", bg=WHITE, bold=False, color="222222",
        size=10, align="left", italic=False, bc="CCCCCC"):
    c = ws.cell(row=row, column=col, value=val)
    c.fill = fill(bg); c.font = fnt(bold, size, color, italic)
    c.alignment = cal("center" if align == "center" else "left")
    c.border = thin(bc); return c


# ── Google Sheets ───────────────────────────────────────────────────────
def get_data_from_sheets():
    creds = Credentials.from_service_account_info(
        json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON")), scopes=SCOPES)
    svc    = build("sheets", "v4", credentials=creds)
    result = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range="Sheet1!A1:BZ2").execute()
    values = result.get("values", [])
    if len(values) < 2: return {}
    headers = values[0]; row = values[1]
    row += [""] * (len(headers) - len(row))
    return dict(zip(headers, row))


# ── Claude prompts ──────────────────────────────────────────────────────
def _story_prompt(data: dict) -> str:
    return f"""คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกฎหมาย

รับข้อมูลลูกค้า แล้ว output 2 ส่วน โดยมี ---SPLIT--- คั่นกลาง

ส่วนที่ 1 — เรื่องราวลูกค้า
เล่าเรื่องต่อเนื่อง 3-5 ย่อหน้า ภาษาไทยธรรมดา อ่านเข้าใจง่าย
ครอบคลุมทุก field ที่มีข้อมูล ข้ามข้อมูลที่เป็น ไม่มี หรือ ไม่ได้ระบุ

จากนั้นพิมพ์ ---SPLIT--- แล้วตามด้วยส่วนที่ 2 ทันที

ส่วนที่ 2 — JSON array เท่านั้น ไม่มีข้อความอื่น ไม่มี markdown

ประเมิน 12 ประเด็นตามลำดับนี้:
1. ค่าใช้จัดการเรื่องหลังเสียชีวิต
2. เงินสดปรับตัว 6-12 เดือน และจัดการหนี้
3. ประกันชีวิตคุ้มครองรายได้และมูลค่าทางเศรษฐกิจตัวเอง
4. Business Succession และเงินหมุนเวียนในกิจการ
5. ประกันโรคร้าย / ทุพพลภาพ
6. คู่สมรสแต่งงานใหม่ I Love You Will
7. กรณีคู่สมรสตายก่อนหรือพร้อมกัน Contingency Clause
8. Guardian of Person (และสำรอง)
9. Money Guardian แยกจาก Guardian
10. เอกสารอยู่ที่ไหนและใครรู้บ้าง
11. ทบทวนแผนทุกปี
12. จดหมายถึงลูกและคู่สมรส

format: [{{"ลำดับ":1,"ประเด็น":"...","ค่าใช้จ่าย":"...","แนะนำ":"...","สถานะ":"..."}}]
สถานะ: ❌ ขาด | ⚠️ บางส่วน | ✅ ครบ

กฎ:
- ข้อ 1: ค่างานศพ × 1.5
- ข้อ 2: หนี้สุทธิ + income×12
- ข้อ 3: HLV = PV(4%/12,(60-age)×12,income,fv=0) แสดงเป็นล้าน
- ข้อ 4: ค่าใช้จ่ายกิจการ×6 (ถ้าไม่มีกิจการ → ไม่มีตัวเลข)
- ข้อ 5: TPD=HLV, CI=income×12
- ข้อ 6-12: ไม่มีตัวเลข
- ทุกข้อที่มีตัวเลข: ต่อท้าย (ประมาณการเบื้องต้น)

ข้อมูลลูกค้า:
{json.dumps(data, ensure_ascii=False, indent=2)}"""


COL_G_RULES = {
    1:  ["- ประเมินค่าใช้จ่ายเร่งด่วน × 1.5 เผื่อบานปลาย",
         "- เทียบกับ เงินสด + สวัสดิการที่จ่ายเร็ว",
         "- พอ → เตือน timing ประกัน 15 วัน + แนะนำระบุสิทธิเบิกธนาคารใน Will",
         "- ไม่พอ → แนะนำประกันตลอดชีพแยกฉบับ ทุนตามที่ขาด"],
    2:  ["- หนี้สุทธิ = หนี้ทั้งหมด - ประกันคุ้มครองหนี้",
         "- Transition = (income_self + income_spouse) × 12",
         "- Cash Need ป.2 = หนี้สุทธิ + Transition แจ้งตัวเลข",
         "- NOTE: ไม่แนะนำขายทรัพย์สิน"],
    3:  ["- HLV = PV(4%/12, (60-age)×12, income_self, fv=0)",
         "- Cash Need รวม = Cash Need ป.2 + HLV หัก ประกันที่มี หัก ทรัพย์สินสภาพคล่อง",
         "- ทุนที่ต้องทำเพิ่ม = Cash Need รวม - ที่หักไป",
         "- แนะนำ Term ปรับลดทุกๆ10ปี ตรวจผู้รับประโยชน์"],
    4:  ["- ไม่มีกิจการ → 1 ประโยค ไม่เกี่ยวข้อง",
         "- มีกิจการ → Keyman = ค่าใช้จ่ายกิจการ × 6 แนะนำ Keyman Term",
         "- มีหุ้นส่วน → แนะนำสัญญาซื้อขายหุ้นกรณีเสียชีวิต"],
    5:  ["- TPD = HLV, CI = income_self × 12",
         "- หักที่มีอยู่แล้ว แสดงส่วนที่ต้องทำเพิ่ม",
         "- แนะนำพ่วง Rider กับประกันหลักป.3"],
    6:  ["- ดู surviving_spouse_plan",
         "- กังวล+ไม่มีแผน → แนะนำแบ่ง Will 2 ส่วน: ให้คู่สมรส / ให้ลูกโดยตรง",
         "- ไม่กังวล → แจ้งความเสี่ยงให้รับทราบ"],
    7:  ["- ดู asset_distribution ว่ามีแผน B ไหม",
         "- ไม่มี → แนะนำเพิ่มประโยคใน Will กรณีคู่สมรสตายก่อน/พร้อมกัน"],
    8:  ["- ระบุชื่อใน Will ไหม? (guardian_primary)",
         "- คุยกับ Guardian แล้วไหม? มีสำรอง? (guardian_backup)"],
    9:  ["- guardian vs money_guardian เดียวกันไหม?",
         "- เดียวกัน → แนะนำแยกทันที มีสำรอง? (money_guardian_backup)"],
    10: ["- ดู documents_location เก็บที่ไหน ใครรู้บ้าง",
         "- แจ้งว่าคุณพยัตจัดทำคู่มือฉุกเฉินให้เป็นส่วนหนึ่งของเอกสารชุดนี้"],
    11: ["- สรุปสถานะแผนตอนนี้ว่าส่วนไหนทำไปแล้ว ส่วนไหนยังต้องทำ",
         "- เน้นว่าทำก่อนแม้ไม่ครบ ดีกว่าไม่มี แล้วค่อยปรับทีหลัง"],
    12: ["- ดู letter_to_spouse, letter_to_children",
         "- ไม่มี → แนะนำเขียน ระบุผู้นำส่ง (estate_executor) และที่เก็บ"],
}

def _col_g_prompt(data: dict, issue_nums: list) -> str:
    sections = []
    for n in issue_nums:
        rules = "\n".join(COL_G_RULES.get(n, []))
        sections.append(f"ประเด็น {n}:\n{rules}")
    return f"""คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกฎหมาย

วิเคราะห์ข้อมูลลูกค้าและเขียนความเห็น Col G สำหรับประเด็น {issue_nums}
output เป็น JSON array เท่านั้น ไม่มีข้อความอื่น ไม่มี markdown
format: [{{"ลำดับ":N,"col_g":"..."}}]

กฎ: ภาษาไทยธรรมดา ใช้ตัวเลขจริง ไม่เกิน 5 ประโยค/ประเด็น ขึ้นต้นด้วยสถานะก่อน

{"=" * 40}
{chr(10).join(sections)}
{"=" * 40}

ข้อมูลลูกค้า:
{json.dumps(data, ensure_ascii=False, indent=2)}"""


# ── call Claude ─────────────────────────────────────────────────────────
def call_claude(data: dict):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    # call 1: story + dynamic issues
    print("กำลังสร้างเรื่องราวและวิเคราะห์...")
    r1 = client.messages.create(
        model=MODEL, max_tokens=8000,
        messages=[{"role": "user", "content": _story_prompt(data)}]
    )
    text1 = r1.content[0].text
    if "---SPLIT---" not in text1:
        raise ValueError("ไม่พบ ---SPLIT---")
    part1, part2 = text1.split("---SPLIT---", 1)
    part2 = part2.strip()
    s = part2.find("["); e = part2.rfind("]") + 1
    dynamic = json.loads(part2[s:e]) if s != -1 and e > 0 else []

    # call 2a: col_g ประเด็น 1-6
    print("กำลัง generate Col G ประเด็น 1-6...")
    col_g_data = []
    for batch in [range(1, 7), range(7, 13)]:
        batch_list = list(batch)
        print(f"  → ประเด็น {batch_list[0]}-{batch_list[-1]}")
        rb = client.messages.create(
            model=MODEL, max_tokens=5000,
            messages=[{"role": "user", "content": _col_g_prompt(data, batch_list)}]
        )
        tb = rb.content[0].text.replace("```json","").replace("```","").strip()
        sb = tb.find("["); eb = tb.rfind("]") + 1
        if sb != -1 and eb > 0:
            try:
                col_g_data += json.loads(tb[sb:eb])
            except json.JSONDecodeError as err:
                print(f"  ⚠️ JSON error: {err}")

    col_g_map = {item.get("ลำดับ"): item.get("col_g","") for item in col_g_data}
    print(f"✅ story + {len(dynamic)} ประเด็น + {len(col_g_data)} Col G")
    return part1.strip(), dynamic, col_g_map


# ── build workbook ──────────────────────────────────────────────────────
def build_workbook(data: dict, story: str, dynamic: list, col_g_map: dict):
    wb = Workbook()

    # Tab 1 — เรื่องราว
    ws1 = wb.active; ws1.title = "1 เรื่องราว"
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions["A"].width = 26
    ws1.column_dimensions["B"].width = 54

    ws1.row_dimensions[1].height = 32
    c = ws1.cell(row=1, column=1,
        value=f"{data.get('nickname','')} | {data.get('occupation','')} | {data.get('age','')} ปี")
    c.fill=fill(NAVY); c.font=fnt(bold=True,size=13,color=WHITE)
    c.alignment=cal(); c.border=thin(NAVY)
    ws1.merge_cells("A1:B1")

    row = 2
    def _header_row(text):
        nonlocal row
        ws1.row_dimensions[row].height = 18
        c = ws1.cell(row=row,column=1,value=text)
        c.fill=fill(MGRAY); c.font=fnt(bold=True,size=10,color=NAVY)
        c.alignment=cal(); c.border=thin(MGRAY)
        ws1.merge_cells(f"A{row}:B{row}")
        row += 1

    _header_row("เรื่องราวลูกค้า")
    for para in story.split("\n"):
        para = para.strip().lstrip("#").lstrip("*").strip()
        if not para: continue
        ws1.row_dimensions[row].height = 60
        c = ws1.cell(row=row,column=1,value=para)
        c.fill=fill("F5F7FA"); c.font=fnt(size=10)
        c.alignment=cal(); c.border=thin()
        ws1.merge_cells(f"A{row}:B{row}")
        row += 1

    _header_row("ข้อมูลรายช่อง")
    for key, label in FIELD_LABELS:
        val = data.get(key,"")
        if not val or val in ("ไม่มี","ไม่ได้ระบุ",""): continue
        ws1.row_dimensions[row].height = 30
        put(ws1,row,1,label,bg="F5F7FA",bold=True,color="444444",size=9)
        put(ws1,row,2,val,size=10)
        row += 1
    ws1.freeze_panes = "A2"

    # Tab 2 — วิเคราะห์
    ws2 = wb.create_sheet("2 วิเคราะห์")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 30
    ws2.column_dimensions["C"].width = 36
    ws2.column_dimensions["D"].width = 28
    ws2.column_dimensions["E"].width = 28
    ws2.column_dimensions["F"].width = 40

    ws2.row_dimensions[1].height = 32
    c = ws2.cell(row=1,column=1,value="วิเคราะห์ช่องว่าง")
    c.fill=fill(NAVY); c.font=fnt(bold=True,size=13,color=WHITE)
    c.alignment=cal(); c.border=thin(NAVY)
    ws2.merge_cells("A1:F1")

    ws2.row_dimensions[2].height = 24
    for col,h in [(1,"#"),(2,"อธิบาย"),(3,"กรณีไม่จัดการ"),
                  (4,"ค่าใช้จ่ายประมาณการ"),(5,"สถานะ"),(6,"ความเห็นคุณพยัต (draft)")]:
        c = ws2.cell(row=2,column=col,value=h)
        c.fill=fill(BLUE); c.font=fnt(bold=True,size=9,color=WHITE)
        c.alignment=cal("center"); c.border=thin(WHITE)

    STATUS_STYLE = {
        "❌ ขาด":      ("FDECEA", RED),
        "⚠️ บางส่วน": ("FEF3E8", AMBER),
        "✅ ครบ":      ("E8F5EE", GREEN),
    }
    dyn_map = {d.get("ลำดับ"): d for d in dynamic}

    for fixed in FIXED_ISSUES:
        i   = fixed["ลำดับ"] + 2
        ws2.row_dimensions[i].height = 85
        bg  = WHITE if fixed["ลำดับ"] % 2 == 1 else LGRAY
        dyn = dyn_map.get(fixed["ลำดับ"], {})
        status = dyn.get("สถานะ","❌ ขาด")
        st_bg, st_fg = STATUS_STYLE.get(status, (LGRAY, NAVY))

        put(ws2,i,1,fixed["ลำดับ"],bg=bg,align="center",color="888888",size=9,bold=True)
        put(ws2,i,2,fixed["อธิบาย"],bg=bg,color="333333",size=9)
        put(ws2,i,3,fixed["กรณีไม่จัดการ"],bg=bg,color="444444",size=9,italic=True)
        put(ws2,i,4,dyn.get("ค่าใช้จ่าย",""),bg=bg,color=NAVY,size=9)
        c = ws2.cell(row=i,column=5,value=status)
        c.fill=fill(st_bg); c.font=fnt(bold=True,size=10,color=st_fg)
        c.alignment=cal("center"); c.border=thin(st_fg)
        c6 = ws2.cell(row=i,column=6,value=col_g_map.get(fixed["ลำดับ"],""))
        c6.fill=fill("FFF8F8"); c6.font=Font(name="Arial",size=9,color=RED,italic=True)
        c6.alignment=cal(); c6.border=thin(RED)
    ws2.freeze_panes = "A3"

    nickname = data.get("nickname","ลูกค้า")
    date_str = datetime.now().strftime("%d%m%Y")
    filename = f"{nickname}_{date_str}.xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return tmp.name, filename


# ── entry point ─────────────────────────────────────────────────────────
def run(data=None):
    if DEV_MODE and data is None:
        print("DEV MODE — ดึงข้อมูลจาก Sheets")
        data = get_data_from_sheets()

    print(f"ข้อมูล: {data.get('nickname')} | {data.get('occupation')} | {data.get('age')} ปี")

    story, dynamic, col_g_map = call_claude(data)

    nickname   = data.get("nickname","ลูกค้า")
    date_str   = datetime.now().strftime("%d%m%Y")
    folder_name = f"{nickname}_{date_str}"
    folder_id   = create_folder(folder_name)

    local_path, filename = build_workbook(data, story, dynamic, col_g_map)
    print(f"สร้างไฟล์: {filename}")
    upload_file_to_folder(local_path, filename, folder_id)
    if os.path.exists(local_path): os.unlink(local_path)

    print("✅ xlsx พร้อม → เริ่มสร้างเอกสาร")
    doc_run(folder_name)
