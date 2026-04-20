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

DEV_MODE = True

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

FIXED_ISSUES = [
    {
        "ลำดับ": 1,
        "ประเด็น": "ค่าใช้จัดการเรื่องหลังเสียชีวิต",
        "อธิบาย": "ทันทีที่เราจากไป จะมีค่าใช้จ่ายเกิดขึ้นทันที ทั้งค่ารักษาพยาบาลครั้งสุดท้าย ค่าจัดงานศพ พิธีทางศาสนา รวมถึงการติดต่อหน่วยงานต่างๆ เช่น การแจ้งตาย การตั้งผู้จัดการมรดก และการรับเงินจากสวัสดิการและสถาบันการเงิน",
        "กรณีไม่จัดการ": "ผู้จัดการเรื่องต้องจ่ายเงินทดรองไปก่อน เพราะยังไม่มีเงินก้อนพร้อมใช้ อาจกลายเป็นภาระให้คนที่รักในช่วงเวลาที่เจ็บปวดที่สุด",
    },
    {
        "ลำดับ": 2,
        "ประเด็น": "เงินสดปรับตัว 6-12 เดือน และจัดการหนี้",
        "อธิบาย": "การสูญเสียไม่ได้กระทบแค่ความรู้สึก แต่กระทบวิถีชีวิตทั้งหมด ครอบครัวอาจต้องย้ายบ้าน ลูกอาจต้องเปลี่ยนคนดูแลหรือเปลี่ยนโรงเรียน คู่ชีวิตอาจต้องลาออกหรือเปลี่ยนงาน ช่วงเปลี่ยนผ่านนี้ต้องใช้เงินก้อนหนึ่งเพื่อไม่ให้คุณภาพชีวิตตกต่ำลงมากนัก และหากยังมีหนี้ค้างอยู่ ก็ยิ่งต้องจัดการให้เสร็จในช่วงนี้ด้วย",
        "กรณีไม่จัดการ": "คู่ชีวิตและลูกๆ ต้องเผชิญกับการเปลี่ยนแปลงครั้งใหญ่ ทั้งเรื่องเงิน เรื่องการใช้ชีวิต และเรื่องมรดก ในขณะที่จิตใจยังไม่พร้อมรับมือ",
    },
    {
        "ลำดับ": 3,
        "ประเด็น": "ประกันชีวิตคุ้มครองรายได้และมูลค่าทางเศรษฐกิจตัวเอง",
        "อธิบาย": "เมื่อเราจากไป ไม่ใช่แค่ตัวเราที่หายไป แต่รายได้ที่เราจะหาได้ตลอดช่วงชีวิตที่เหลือก็หายไปด้วย ประกันชีวิตที่ดีควรเข้ามาชดเชยรายได้ส่วนนี้ให้ครอบครัว เสมือนว่าเราหาเงินให้พวกเขาได้อีกหลายปี",
        "กรณีไม่จัดการ": "คู่ชีวิตจะต้องแบกภาระหาเลี้ยงครอบครัวเพียงคนเดียว คุณภาพชีวิตจะลดลงอย่างหลีกเลี่ยงไม่ได้ โดยเฉพาะในระยะยาว",
    },
    {
        "ลำดับ": 4,
        "ประเด็น": "Business Succession และเงินหมุนเวียนในกิจการ",
        "อธิบาย": "การบริหารกิจการมักต้องอาศัยลายเซ็นหรืออำนาจของผู้บริหารในการอนุมัติงานและทำธุรกรรมต่างๆ หากผู้บริหารจากไปกะทันหัน คนที่เข้ามาดูแลต่อต้องมีอำนาจตามกฎหมายและมีเงินหมุนเวียนพอที่จะพยุงกิจการไว้ระหว่างช่วงเปลี่ยนผ่าน",
        "กรณีไม่จัดการ": "กิจการอาจหยุดชะงัก อนุมัติงานไม่ได้ เบิกจ่ายเงินไม่ได้ หรือขาดสภาพคล่องในช่วงปรับตัว จนส่งผลต่อรายได้ของครอบครัวที่ยังต้องพึ่งพากิจการนั้นอยู่",
    },
    {
        "ลำดับ": 5,
        "ประเด็น": "ประกันโรคร้าย / ทุพพลภาพ",
        "อธิบาย": "สิ่งที่น่ากลัวกว่าความตายคือการพิการหรือทุพพลภาพ เพราะนอกจากจะหาเงินไม่ได้แล้ว ประกันชีวิตยังไม่จ่าย แถมยังมีค่าดูแลรักษาเพิ่มขึ้นมาอีก ประกันหลายประเภทช่วยอุดช่องว่างตรงนี้ได้ ทั้งประกันโรคร้าย ประกันชดเชยทุพพลภาพ และสิทธิ์งดจ่ายเบี้ยประกันชีวิต",
        "กรณีไม่จัดการ": "คู่ชีวิตต้องแบกรับทุกอย่างพร้อมกัน ทั้งหาเงิน เลี้ยงลูก และดูแลเราในฐานะผู้พิการ คุณภาพชีวิตของทั้งครอบครัวอาจตกต่ำถึงขีดสุด",
    },
    {
        "ลำดับ": 6,
        "ประเด็น": "คู่สมรสแต่งงานใหม่ I Love You Will",
        "อธิบาย": "เมื่อเราจากไป คู่ชีวิตอาจตัดสินใจเริ่มต้นชีวิตใหม่ ซึ่งเป็นเรื่องปกติ แต่ถ้าเราทำพินัยกรรมมอบทุกอย่างให้คู่ชีวิตโดยไม่มีเงื่อนไข ที่เรียกว่า I Love You Will สิ่งที่เราหวังว่าจะตกทอดไปถึงลูกอาจไม่เป็นตามที่คิด",
        "กรณีไม่จัดการ": "ทรัพย์สินที่เราสร้างมาทั้งชีวิตอาจไหลไปสู่ครอบครัวใหม่ของคู่ชีวิต แทนที่จะถึงมือลูกแท้ๆ ของเรา เหตุการณ์แบบนี้พบได้บ่อยกว่าที่คิด",
    },
    {
        "ลำดับ": 7,
        "ประเด็น": "กรณีคู่สมรสตายก่อนหรือพร้อมกัน Contingency Clause",
        "อธิบาย": "เราอาจวางแผนให้คู่ชีวิตรับมรดกและดูแลลูกต่อ แต่ถ้าคู่ชีวิตจากไปก่อนเราโดยที่เราไม่ได้แก้พินัยกรรม หรือเราทั้งสองจากไปพร้อมกัน มรดกอาจไหลไปตามกฎหมายโดยที่ลูกได้ไม่เต็มเม็ดเต็มหน่วย",
        "กรณีไม่จัดการ": "ในกฎหมายไทย ปู่ย่าตายายหรือพ่อแม่ของเจ้ามรดกมีสิทธิรับมรดกในลำดับเดียวกับลูก มีโอกาสที่มรดกจะถูกแบ่งออกไปโดยไม่ตั้งใจ และในบางกรณีเจ้าหนี้ก็มีสิทธิฟ้องร้องได้",
    },
    {
        "ลำดับ": 8,
        "ประเด็น": "Guardian of Person (และสำรอง)",
        "อธิบาย": "กรณีที่พ่อแม่จากไปพร้อมกันเป็นเหตุการณ์ที่คาดไม่ถึง แต่ถ้าเกิดขึ้นจริง ลูกจะไปอยู่ที่ไหน ใครจะดูแล และสำคัญกว่านั้นคือเขาจะถูกดูแลอย่างที่เราต้องการจริงไหม",
        "กรณีไม่จัดการ": "หากไม่ได้วางแผนไว้ อาจเกิดการแย่งชิงตัวลูกเมื่อรู้ว่ามีมรดก หรือกลับกันคือต่างคนต่างเกี่ยงกันดูแล และลูกอาจไม่ได้รับการเลี้ยงดูอย่างที่เราตั้งใจ",
    },
    {
        "ลำดับ": 9,
        "ประเด็น": "Money Guardian แยกจาก Guardian",
        "อธิบาย": "หากเราทิ้งมรดกไว้พอสมควร ผู้ที่อาสาเลี้ยงลูกอาจมีเป้าหมายที่ตัวเงินมากกว่าตัวลูก การแยกคนดูแลลูกออกจากคนคุมเงินของลูกจึงเป็นกลไกสำคัญที่ช่วยตรวจสอบซึ่งกันและกัน",
        "กรณีไม่จัดการ": "หากผู้ปกครองและผู้คุมเงินเป็นคนเดียวกัน ไม่มีใครตรวจสอบว่าเงินถูกใช้เพื่อลูกจริงๆ ลูกอาจกินอยู่ไม่ดี เรียนโรงเรียนไม่ดี หรือเงินมรดกถูกยักยอกไปใช้ส่วนตัว",
    },
    {
        "ลำดับ": 10,
        "ประเด็น": "ป้องกันลูกใช้มรดกในทางที่ผิดหรือหมดเร็วเกินไป",
        "อธิบาย": "หากมรดกที่เราทิ้งไว้มีมูลค่ามาก และเรายังไม่แน่ใจว่าลูกจะมีวุฒิภาวะพอรับผิดชอบเงินก้อนใหญ่ได้ การใส่เงื่อนไขในพินัยกรรม เช่น ห้ามโอนที่ดินก่อนอายุ 35 ปี หรือให้บริษัทประกันทยอยจ่ายเป็นงวดๆ ก็เป็นทางเลือกที่น่าพิจารณา",
        "กรณีไม่จัดการ": "ลูกอาจได้รับมรดกก้อนใหญ่ตั้งแต่อายุยังน้อย ยังขาดประสบการณ์และวุฒิภาวะที่จะรับมือกับเงินจำนวนมาก เงินอาจหมดเร็วหรือถูกนำไปใช้ในทางที่ผิด",
    },
    {
        "ลำดับ": 11,
        "ประเด็น": "ขั้นตอนหลังเกิดเหตุ และจดหมายถึงผู้จัดการเรื่อง",
        "อธิบาย": "เมื่อเกิดเหตุไม่คาดฝัน สถานการณ์จะวุ่นวายมาก ทั้งเรื่องด่วน เรื่องเอกสาร เรื่องสถาบันการเงิน และการติดต่อหน่วยงานต่างๆ คู่มือที่เขียนไว้ล่วงหน้าจะช่วยให้ผู้จัดการเรื่องทำงานได้รวดเร็วและไม่ตกหล่น",
        "กรณีไม่จัดการ": "ผู้จัดการเรื่องอาจไม่รู้ว่าต้องติดต่อใคร เอกสารอยู่ที่ไหน หน่วยงานไหนต้องแจ้งก่อน ความล่าช้าอาจทำให้ครอบครัวได้รับเงินช้าลงหรือเสียสิทธิ์บางอย่างไป",
    },
    {
        "ลำดับ": 12,
        "ประเด็น": "จดหมายถึงลูกและคู่สมรส",
        "อธิบาย": "จดหมายถึงครอบครัวเป็นสิ่งที่ไม่มีเอกสารทางกฎหมายใดทดแทนได้ คู่ชีวิตและลูกๆ จะได้รับรู้ความรักของเรา เหตุผลเบื้องหลังการตัดสินใจต่างๆ รวมถึงค่านิยมและแนวทางชีวิตที่เราอยากฝากไว้",
        "กรณีไม่จัดการ": "หากไม่มีจดหมาย ครอบครัวอาจเข้าใจผิดในสิ่งที่เราตัดสินใจโดยไม่มีโอกาสได้อธิบาย และเสียโอกาสที่จะได้รับสิ่งที่มีค่าที่สุดจากเรา นั่นคือตัวตนและความรู้สึกของเรา",
    },
]

FIELD_LABELS = [
    ("nickname", "ชื่อเล่น"), ("age", "อายุ"), ("gender", "เพศ"),
    ("occupation", "อาชีพ"), ("health", "สุขภาพ"), ("email", "อีเมล"),
    ("hobbies_and_risks", "งานอดิเรก"),
    ("spouse_nickname", "คู่สมรส"), ("spouse_age", "อายุคู่สมรส"),
    ("spouse_occupation", "อาชีพคู่สมรส"), ("spouse_income", "รายได้คู่สมรส"),
    ("spouse_health", "สุขภาพคู่สมรส"), ("spouse_status", "สถานะ"),
    ("children", "ลูก"), ("children_outside_marriage", "บุตรนอกสมรส"),
    ("assets_cash", "เงินสด"), ("assets_property", "อสังหาริมทรัพย์"),
    ("assets_investment", "การลงทุน"), ("assets_crypto_wallet", "คริปโต"),
    ("assets_insurance_savings", "ประกันสะสม"), ("assets_digital", "ทรัพย์สินดิจิทัล"),
    ("assets_business", "กิจการ"), ("assets_valuables", "ของมีค่า"),
    ("debt", "หนี้สิน"), ("guarantor", "ค้ำประกัน"),
    ("insurance_life", "ประกันชีวิต"), ("insurance_health", "ประกันสุขภาพ"),
    ("insurance_group", "ประกันกลุ่ม"), ("welfare", "สวัสดิการ"),
    ("funeral_wishes", "ความปรารถนางานศพ"), ("emergency_cash_90days", "เงินฉุกเฉิน 90 วัน"),
    ("estate_admin_cost", "ต้นทุนจัดการมรดก"), ("asset_distribution", "แผนแบ่งทรัพย์"),
    ("debt_responsibility", "หนี้ใครรับผิดชอบ"), ("business_succession", "แผนกิจการ"),
    ("urgent_manager", "ผู้จัดการฉุกเฉิน"), ("estate_executor", "ผู้จัดการมรดก"),
    ("financial_poa", "Financial POA"), ("living_will", "Living Will"),
    ("surviving_spouse_plan", "แผนคู่สมรสที่รอดชีวิต"),
    ("guardian_primary", "Guardian หลัก"), ("guardian_backup", "Guardian สำรอง"),
    ("money_guardian_primary", "Money Guardian หลัก"),
    ("money_guardian_backup", "Money Guardian สำรอง"),
    ("documents_location", "ที่อยู่เอกสาร"),
    ("letter_to_children", "จดหมายถึงลูก"), ("letter_to_spouse", "จดหมายถึงคู่สมรส"),
]


def fill(h):
    return PatternFill("solid", start_color=h, end_color=h)

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
    c.fill = fill(bg)
    c.font = fnt(bold, size, color, italic)
    c.alignment = cal("center" if align == "center" else "left")
    c.border = thin(bc)
    return c


def get_data_from_sheets():
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


def build_story_prompt(data):
    data_str = json.dumps(data, ensure_ascii=False, indent=2)
    lines = [
        "คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกฎหมาย",
        "",
        "รับข้อมูลลูกค้า แล้ว output 2 ส่วน โดยมี ---SPLIT--- คั่นกลาง",
        "",
        "ส่วนที่ 1 — เรื่องราวลูกค้า",
        "เล่าเรื่องต่อเนื่อง 3-5 ย่อหน้า ภาษาไทยธรรมดา อ่านเข้าใจง่าย",
        "ครอบคลุมทุก field ที่มีข้อมูล ข้ามข้อมูลที่เป็น ไม่มี หรือ ไม่ได้ระบุ",
        "",
        "จากนั้นพิมพ์ ---SPLIT--- แล้วตามด้วยส่วนที่ 2 ทันที",
        "",
        "ส่วนที่ 2 — JSON array เท่านั้น ไม่มีข้อความอื่น ไม่มี markdown",
        "",
        "ประเมิน 12 ประเด็นตามลำดับนี้เท่านั้น:",
        "1. ค่าใช้จัดการเรื่องหลังเสียชีวิต",
        "2. เงินสดปรับตัว 6-12 เดือน และจัดการหนี้",
        "3. ประกันชีวิตคุ้มครองรายได้และมูลค่าทางเศรษฐกิจตัวเอง",
        "4. Business Succession และเงินหมุนเวียนในกิจการ",
        "5. ประกันโรคร้าย / ทุพพลภาพ",
        "6. คู่สมรสแต่งงานใหม่ I Love You Will",
        "7. กรณีคู่สมรสตายก่อนหรือพร้อมกัน Contingency Clause",
        "8. Guardian of Person (และสำรอง)",
        "9. Money Guardian แยกจาก Guardian",
        "10. ป้องกันลูกใช้มรดกในทางที่ผิดหรือหมดเร็วเกินไป",
        "11. ขั้นตอนหลังเกิดเหตุ และจดหมายถึงผู้จัดการเรื่อง",
        "12. จดหมายถึงลูกและคู่สมรส",
        "",
        'format: [{"ลำดับ":1,"ประเด็น":"...","ค่าใช้จ่าย":"...","แนะนำ":"...","สถานะ":"..."}]',
        "",
        "สถานะ: ❌ ขาด, ⚠️ บางส่วน, ✅ ครบ",
        "",
        "กฎ:",
        "- ไม่มีประกันชีวิตคุ้มครอง → ❌ ขาด",
        "- ไม่มีเงินฉุกเฉิน → ❌ ขาด",
        "- ไม่มี Contingency Clause → ❌ ขาด",
        "- Money Guardian = Guardian คนเดียวกัน → ⚠️ บางส่วน",
        "- กังวลคู่สมรสแต่งใหม่แต่ไม่มีแผน → ⚠️ บางส่วน",
        "- ไม่มี Money Guardian สำรอง → ⚠️ บางส่วน",
        "- ข้อ 1: ค่างานศพ × 1.5 เผื่อบานปลาย",
        "- ข้อ 2: หนี้สุทธิ + (income_self+income_spouse)×12",
        "- ข้อ 3: HLV = PV(4%/12,(60-age)×12,income_self,fv=0) แสดงเป็นล้าน",
        "- ข้อ 4: ค่าใช้จ่ายกิจการ×6 (ถ้าไม่มีกิจการ → ไม่มีตัวเลข)",
        "- ข้อ 5: TPD=HLV, CI=income_self×12",
        "- ข้อ 6-12: ไม่มีตัวเลข",
        "- ทุกข้อที่มีตัวเลข: ต่อท้ายด้วย (ประมาณการเบื้องต้น)",
        "",
        "ข้อมูลลูกค้า:",
        data_str,
    ]
    return "\n".join(lines)


def build_col_g_prompt(data):
    data_str = json.dumps(data, ensure_ascii=False, indent=2)
    lines = [
        "คุณคือผู้ช่วยของคุณพยัต นักวางแผนการเงินและกฎหมาย",
        "",
        "วิเคราะห์ข้อมูลลูกค้าและเขียนความเห็น Col G สำหรับทั้ง 12 ประเด็น",
        "output เป็น JSON array เท่านั้น ไม่มีข้อความอื่น ไม่มี markdown",
        "",
        'format: [{"ลำดับ":1,"col_g":"..."},{"ลำดับ":2,"col_g":"..."},...]',
        "",
        "กฎทั่วไป:",
        "- ภาษาไทยธรรมดา อ่านเข้าใจง่าย ใช้ตัวเลขจริงของลูกค้าตลอด",
        "- ความยาวแต่ละ col_g ไม่เกิน 5 ประโยค",
        "- ขึ้นต้นด้วยสถานะหรือสรุปก่อนเสมอ",
        "",
        "=" * 50,
        "ประเด็น 1 — ค่าใช้จัดการเรื่องหลังเสียชีวิต",
        "- ประเมินค่าใช้จ่ายเร่งด่วน × 1.5 เผื่อบานปลาย",
        "- เทียบกับ เงินสด + สวัสดิการที่จ่ายเร็ว",
        "- พอ → เตือน timing ประกัน 15 วัน + แนะนำระบุสิทธิเบิกธนาคารใน Will",
        "- ไม่พอ → แนะนำประกันตลอดชีพแยกฉบับ ทุนตามที่ขาด",
        "- NOTE ทุกเคส: ต้องมีคนทดรองเงิน / แยกผู้จัดการศพกับผู้ทดรองเงิน",
        "",
        "=" * 50,
        "ประเด็น 2 — เงินสดปรับตัวและจัดการหนี้",
        "- หนี้สุทธิ = หนี้ทั้งหมด - ประกันคุ้มครองหนี้โดยเฉพาะ",
        "- Transition = (income_self + income_spouse) × 12",
        "- Cash Need ป.2 = หนี้สุทธิ + Transition",
        "- แจ้งตัวเลข Cash Need ป.2 และบอกจะรวมกับ HLV ป.3 ก่อนแนะนำประกัน",
        "- NOTE: ไม่แนะนำขายทรัพย์สิน",
        "",
        "=" * 50,
        "ประเด็น 3 — HLV ประกันชีวิต",
        "- HLV = PV(4%/12, (60-age)×12, income_self, fv=0)",
        "- Cash Need รวม = Cash Need ป.2 + HLV",
        "- หัก ประกันชีวิตที่มี (ไม่รวมประกันคุ้มครองหนี้)",
        "- หัก ทรัพย์สินสภาพคล่อง (เงินสด+หุ้น+กองทุน+ทอง+คริปโต) ยกเว้นอสังหาอาศัย",
        "- ทุนที่ต้องทำเพิ่ม = Cash Need รวม - ที่หักไป",
        "- แนะนำ Term10 ปรับลดทุกๆ10ปี หรือ Unit Linked",
        "- ตรวจผู้รับประโยชน์ตรงกับ Will ไหม",
        "- NOTE: ประมาณการหยาบๆ / คู่สมรสควรทำด้วย",
        "",
        "=" * 50,
        "ประเด็น 4 — Business Succession",
        "- ถ้าไม่มีกิจการ → 1 ประโยค: ไม่มีกิจการ ไม่เกี่ยวข้อง",
        "- ถ้ามีกิจการ → Keyman = ค่าใช้จ่ายกิจการต่อเดือน × 6 แนะนำ Keyman Term",
        "- มีหุ้นส่วน → แนะนำสัญญาซื้อขายหุ้นกรณีเสียชีวิต (ปรึกษาทนาย)",
        "- จดทะเบียนบริษัท → แนะนำตั้งกรรมการที่ไว้ใจมีอำนาจลงนาม",
        "",
        "=" * 50,
        "ประเด็น 5 — TPD/CI",
        "- TPD = HLV (เงินก้อนสุดท้ายดูแลตนเองและครอบครัว)",
        "- CI = income_self × 12 (ค่าใช้จ่ายนอกเหนือค่ารักษา)",
        "- หักที่มีอยู่แล้ว แล้วแสดงส่วนที่ต้องทำเพิ่ม",
        "- แนะนำพ่วง Rider กับประกันหลักป.3",
        "",
        "=" * 50,
        "ประเด็น 6 — I Love You Will",
        "- ดู surviving_spouse_plan",
        "- กังวล+ไม่มีแผน → แนะนำแบ่ง Will 2 ส่วน: ให้คู่สมรส / ให้ลูกโดยตรง",
        "- ไม่กังวล → แจ้งความเสี่ยงให้รับทราบ",
        "",
        "=" * 50,
        "ประเด็น 7 — Contingency Clause",
        "- ดู asset_distribution ว่ามีแผน B ไหม",
        "- ไม่มี → แนะนำเพิ่มประโยคใน Will กรณีคู่สมรสตายก่อน/พร้อมกัน",
        "- มี → ชม + แนะนำทบทวนชื่อสำรอง",
        "",
        "=" * 50,
        "ประเด็น 8 — Guardian of Person",
        "- ระบุชื่อใน Will ไหม? (guardian_primary)",
        "- คุยกับ Guardian แล้วไหม? (ดู gaps_for_payat)",
        "- มีตัวสำรองครอบคลุม 3 กรณีไหม? (guardian_backup)",
        "- NOTE: Guardian ≠ Money Guardian อธิบายในป.9",
        "",
        "=" * 50,
        "ประเด็น 9 — Money Guardian",
        "- guardian_primary vs money_guardian_primary เดียวกันไหม?",
        "- เดียวกัน → แนะนำแยกทันที",
        "- มีตัวสำรองครอบคลุม 3 กรณีไหม? (money_guardian_backup)",
        "- NOTE: Money Guardian รายงานศาลปีละครั้งตามกฎหมายไทย",
        "",
        "=" * 50,
        "ประเด็น 10 — ป้องกันลูกใช้มรดกผิด",
        "- คำนวณมรดกที่จะตกถึงลูก",
        "- น้อยกว่า 5 ล้าน → 1 ประโยค: ยังไม่จำเป็น",
        "- 5 ล้านขึ้นไป → เสนอ 2 ทางเลือก:",
        "  A: ทยอยให้ใน Will อายุ 20/25/30/35 ส่วนละ 25%",
        "  B: Settlement Option ให้บริษัทประกันทยอยจ่าย",
        "",
        "=" * 50,
        "ประเด็น 11 — ขั้นตอนหลังเกิดเหตุ",
        "- แจ้งว่าคุณพยัตจัดทำคู่มือให้เป็นส่วนหนึ่งของเอกสารชุดนี้",
        "- แนะนำ review ทุกปี หรือเมื่อมีการเปลี่ยนแปลง",
        "- ระบุที่อยู่เอกสาร (documents_location) ถ้ามี",
        "",
        "=" * 50,
        "ประเด็น 12 — จดหมายถึงลูกและคู่สมรส",
        "- ดู letter_to_spouse, letter_to_children",
        "- มีแล้ว → ชม แนะนำทบทวนทุกปี",
        "- ไม่มี → แนะนำเขียน: ความรัก เหตุผลใน Will คำแนะนำการใช้ชีวิต",
        "- ระบุผู้นำส่ง (estate_executor) และที่เก็บ",
        "",
        "=" * 50,
        "ข้อมูลลูกค้า:",
        data_str,
    ]
    return "\n".join(lines)


def call_claude(data):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    # Call 1: story + dynamic (D/E/F)
    print("กำลังสร้างเรื่องราวและวิเคราะห์...")
    r1 = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": build_story_prompt(data)}]
    )
    result1 = r1.content[0].text

    if "---SPLIT---" not in result1:
        raise ValueError("ไม่พบ ---SPLIT---")

    part1, part2 = result1.split("---SPLIT---", 1)
    part2 = part2.strip()
    start = part2.find("[")
    end = part2.rfind("]") + 1
    if start == -1 or end == 0:
        raise ValueError("ไม่พบ JSON array ในส่วนที่ 2")
    dynamic = json.loads(part2[start:end])

    # Call 2: Col G
    print("กำลัง generate Col G...")
    r2 = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[{"role": "user", "content": build_col_g_prompt(data)}]
    )
    result2 = r2.content[0].text.strip()
    start2 = result2.find("[")
    end2 = result2.rfind("]") + 1
    if start2 == -1 or end2 == 0:
        raise ValueError("ไม่พบ JSON array ใน Col G")
    col_g_data = json.loads(result2[start2:end2])
    col_g_map = {item.get("ลำดับ"): item.get("col_g", "") for item in col_g_data}

    print(f"✅ story + {len(dynamic)} ประเด็น + {len(col_g_data)} Col G")
    return part1.strip(), dynamic, col_g_map


def build_workbook(data, story, dynamic, col_g_map):
    wb = Workbook()

    # ── Tab 1 ──────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "1 เรื่องราว"
    ws1.sheet_view.showGridLines = False
    ws1.column_dimensions["A"].width = 26
    ws1.column_dimensions["B"].width = 54

    ws1.row_dimensions[1].height = 32
    c = ws1.cell(row=1, column=1,
        value=f"{data.get('nickname','')} | {data.get('occupation','')} | {data.get('age','')} ปี")
    c.fill = fill(NAVY); c.font = fnt(bold=True, size=13, color=WHITE)
    c.alignment = cal(); c.border = thin(NAVY)
    ws1.merge_cells("A1:B1")

    row = 2
    ws1.row_dimensions[row].height = 18
    c = ws1.cell(row=row, column=1, value="เรื่องราวลูกค้า")
    c.fill = fill(MGRAY); c.font = fnt(bold=True, size=10, color=NAVY)
    c.alignment = cal(); c.border = thin(MGRAY)
    ws1.merge_cells(f"A{row}:B{row}")
    row += 1

    for para in story.split("\n"):
        para = para.strip().lstrip("#").lstrip("*").strip()
        if not para:
            continue
        ws1.row_dimensions[row].height = 60
        c = ws1.cell(row=row, column=1, value=para)
        c.fill = fill(LGRAY); c.font = fnt(size=10)
        c.alignment = cal(); c.border = thin()
        ws1.merge_cells(f"A{row}:B{row}")
        row += 1

    ws1.row_dimensions[row].height = 18
    c = ws1.cell(row=row, column=1, value="ข้อมูลรายช่อง")
    c.fill = fill(MGRAY); c.font = fnt(bold=True, size=10, color=NAVY)
    c.alignment = cal(); c.border = thin(MGRAY)
    ws1.merge_cells(f"A{row}:B{row}")
    row += 1

    for key, label in FIELD_LABELS:
        val = data.get(key, "")
        if not val or val in ("ไม่มี", "ไม่ได้ระบุ", ""):
            continue
        ws1.row_dimensions[row].height = 30
        put(ws1, row, 1, label, bg=LGRAY, bold=True, color="444444", size=9)
        put(ws1, row, 2, val, size=10)
        row += 1

    ws1.freeze_panes = "A2"

    # ── Tab 2 ──────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("2 วิเคราะห์")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 30
    ws2.column_dimensions["C"].width = 36
    ws2.column_dimensions["D"].width = 28
    ws2.column_dimensions["E"].width = 28
    ws2.column_dimensions["F"].width = 40
    
    ws2.row_dimensions[1].height = 32
    c = ws2.cell(row=1, column=1, value="วิเคราะห์ช่องว่าง")
    c.fill = fill(NAVY); c.font = fnt(bold=True, size=13, color=WHITE)
    c.alignment = cal(); c.border = thin(NAVY)
    ws2.merge_cells("A1:F1")

    ws2.row_dimensions[2].height = 24
    for col, h in [(1,"#"),(2,"อธิบาย"),(3,"กรณีไม่จัดการ"), 
                   (4,"ค่าใช้จ่ายประมาณการ"),(5,"สถานะ"),(6,"ความเห็นคุณพยัต (draft)")]:
        c = ws2.cell(row=2, column=col, value=h)
        c.fill = fill(BLUE); c.font = fnt(bold=True, size=9, color=WHITE)
        c.alignment = cal("center"); c.border = thin(WHITE)

    STATUS_STYLE = {
        "❌ ขาด":      ("FDECEA", RED),
        "⚠️ บางส่วน": ("FEF3E8", AMBER),
        "✅ ครบ":      ("E8F5EE", GREEN),
    }

    dyn_map = {d.get("ลำดับ"): d for d in dynamic}

    for fixed in FIXED_ISSUES:
        i = fixed["ลำดับ"] + 2
        ws2.row_dimensions[i].height = 85
        bg = WHITE if fixed["ลำดับ"] % 2 == 1 else LGRAY
        dyn = dyn_map.get(fixed["ลำดับ"], {})
        status = dyn.get("สถานะ", "❌ ขาด")
        st_bg, st_fg = STATUS_STYLE.get(status, (LGRAY, NAVY))

        put(ws2, i, 1, fixed["ลำดับ"], bg=bg, align="center", color="888888", size=9, bold=True)
        put(ws2, i, 2, fixed["อธิบาย"], bg=bg, color="333333", size=9)
        put(ws2, i, 3, fixed["กรณีไม่จัดการ"], bg=bg, color="444444", size=9, italic=True)
        put(ws2, i, 4, dyn.get("ค่าใช้จ่าย", ""), bg=bg, color=NAVY, size=9)

        c = ws2.cell(row=i, column=5, value=status)
        c.fill = fill(st_bg); c.font = fnt(bold=True, size=10, color=st_fg)
        c.alignment = cal("center"); c.border = thin(st_fg)

        # Col G — AI draft สีแดง คุณพยัตแก้ทีหลัง
        col_g_text = col_g_map.get(fixed["ลำดับ"], "")
        c6 = ws2.cell(row=i, column=6, value=col_g_text)
        c6.fill = fill("FFF8F8")
        c6.font = Font(name="Arial", size=9, color=RED, italic=True)
        c6.alignment = cal()
        c6.border = thin(RED)

    ws2.freeze_panes = "A3"

    nickname = data.get("nickname", "ลูกค้า")
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{nickname}_{date_str}.xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    return tmp.name, filename


def run(data=None):
    if DEV_MODE and data is None:
        print("DEV MODE — ดึงข้อมูลจาก Sheets")
        data = get_data_from_sheets()

    print(f"ข้อมูล: {data.get('nickname')} | {data.get('occupation')} | {data.get('age')} ปี")

    story, dynamic, col_g_map = call_claude(data)

    local_path, filename = build_workbook(data, story, dynamic, col_g_map)
    print(f"สร้างไฟล์: {filename}")

    upload_file(local_path, filename)

    notify_new_client(
        nickname=data.get("nickname", ""),
        occupation=data.get("occupation", ""),
        age=data.get("age", "")
    )

    os.unlink(local_path)
    print("เสร็จสิ้น")
