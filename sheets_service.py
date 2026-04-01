# sheets_service.py
# บันทึกข้อมูลลูกค้าลง Google Sheets

import os
import json
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from config import GOOGLE_SHEET_ID

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Header columns ตาม JSON fields ของน้องแพลน
HEADERS = [
    "timestamp",
    "line_user_id",
    "name",
    "age",
    "gender",
    "location",
    "lifestyle",
    "hobbies_and_risks",
    "occupation",
    "income_type",
    "marital_status",
    "children",
    "dependents_parents",
    "family_health_history",
    "religion_constraints",
    "income_monthly",
    "expense_monthly",
    "cashflow_monthly",
    "savings",
    "debt",
    "assets",
    "insurance_existing",
    "welfare_benefits",
    "financial_goal",
    "risk_tolerance",
    "investment_constraints",
    "financial_concerns",
    "summary",
]


def get_sheet():
    """เชื่อมต่อ Google Sheets"""
    # อ่าน credentials จาก environment variable
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON not found in environment")

    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
    return sheet


def ensure_headers(sheet):
    """ตรวจสอบว่ามี header row แล้วหรือยัง ถ้าไม่มีให้สร้าง"""
    first_row = sheet.row_values(1)
    if not first_row or first_row[0] != "timestamp":
        sheet.insert_row(HEADERS, index=1)
        # จัด format header ให้ดูง่าย
        sheet.format("1:1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.2, "green": 0.6, "blue": 0.2},
        })


def save_to_sheets(line_user_id: str, data: dict):
    """บันทึกข้อมูลลูกค้าลง Google Sheets"""
    try:
        sheet = get_sheet()
        ensure_headers(sheet)

        # สร้าง row ตาม HEADERS order
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # timestamp
            line_user_id,                                    # line_user_id
            data.get("name", ""),
            data.get("age", ""),
            data.get("gender", ""),
            data.get("location", ""),
            data.get("lifestyle", ""),
            data.get("hobbies_and_risks", ""),
            data.get("occupation", ""),
            data.get("income_type", ""),
            data.get("marital_status", ""),
            data.get("children", ""),
            data.get("dependents_parents", ""),
            data.get("family_health_history", ""),
            data.get("religion_constraints", ""),
            data.get("income_monthly", ""),
            data.get("expense_monthly", ""),
            data.get("cashflow_monthly", ""),
            data.get("savings", ""),
            data.get("debt", ""),
            data.get("assets", ""),
            data.get("insurance_existing", ""),
            data.get("welfare_benefits", ""),
            data.get("financial_goal", ""),
            data.get("risk_tolerance", ""),
            data.get("investment_constraints", ""),
            data.get("financial_concerns", ""),
            data.get("summary", ""),
        ]

        sheet.append_row(row, value_input_option="USER_ENTERED")
        print(f"✅ บันทึกข้อมูลของ {data.get('name', line_user_id)} สำเร็จ")
        return True

    except Exception as e:
        print(f"❌ Error saving to sheets: {e}")
        return False
