# sheets_service.py
# บันทึกข้อมูลลูกค้าลง Google Sheets — น้องแพลน v2

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

HEADERS = [
    "timestamp", "line_user_id",
    "nickname", "age", "gender", "occupation", "health",
    "income_self", "hobbies_and_risks", "email",
    "spouse_nickname", "spouse_age", "spouse_occupation",
    "spouse_income", "spouse_health", "spouse_status",
    "children", "children_outside_marriage",
    "assets_cash", "assets_property", "assets_investment",
    "assets_crypto_wallet", "assets_insurance_savings",
    "assets_digital", "assets_business", "assets_valuables",
    "debt", "guarantor",
    "insurance_life", "insurance_health",
    "insurance_group", "welfare",
    "funeral_wishes", "emergency_cash_90days", "estate_admin_cost",
    "asset_distribution", "debt_responsibility", "business_succession",
    "urgent_manager", "estate_executor", "documents_location", "financial_poa",
    "living_will", "surviving_spouse_plan",
    "guardian_primary", "guardian_backup",
    "money_guardian_primary", "money_guardian_backup",
    "gaps_for_payat", "summary",
]


def get_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON not found in environment")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
    return sheet


def ensure_headers(sheet):
    first_row = sheet.row_values(1)
    if not first_row or first_row[0] != "timestamp":
        sheet.clear()
        sheet.insert_row(HEADERS, index=1)
        sheet.format("1:1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.2, "green": 0.6, "blue": 0.2},
        })


def save_to_sheets(line_user_id: str, data: dict):
    try:
        sheet = get_sheet()
        ensure_headers(sheet)

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            line_user_id,
            data.get("nickname", ""),
            data.get("age", ""),
            data.get("gender", ""),
            data.get("occupation", ""),
            data.get("health", ""),
            data.get("income_self", ""),
            data.get("hobbies_and_risks", ""),
            data.get("email", ""),
            data.get("spouse_nickname", ""),
            data.get("spouse_age", ""),
            data.get("spouse_occupation", ""),
            data.get("spouse_income", ""),
            data.get("spouse_health", ""),
            data.get("spouse_status", ""),
            data.get("children", ""),
            data.get("children_outside_marriage", ""),
            data.get("assets_cash", ""),
            data.get("assets_property", ""),
            data.get("assets_investment", ""),
            data.get("assets_crypto_wallet", ""),
            data.get("assets_insurance_savings", ""),
            data.get("assets_digital", ""),
            data.get("assets_business", ""),
            data.get("assets_valuables", ""),
            data.get("debt", ""),
            data.get("guarantor", ""),
            data.get("insurance_life", ""),
            data.get("insurance_health", ""),
            data.get("insurance_group", ""),
            data.get("welfare", ""),
            data.get("funeral_wishes", ""),
            data.get("emergency_cash_90days", ""),
            data.get("estate_admin_cost", ""),
            data.get("asset_distribution", ""),
            data.get("debt_responsibility", ""),
            data.get("business_succession", ""),
            data.get("urgent_manager", ""),
            data.get("estate_executor", ""),
            data.get("documents_location", ""),
            data.get("financial_poa", ""),
            data.get("living_will", ""),
            data.get("surviving_spouse_plan", ""),
            data.get("guardian_primary", ""),
            data.get("guardian_backup", ""),
            data.get("money_guardian_primary", ""),
            data.get("money_guardian_backup", ""),
            data.get("gaps_for_payat", ""),
            data.get("summary", ""),
        ]

        sheet.append_row(row, value_input_option="USER_ENTERED")
        print(f"✅ บันทึกข้อมูลของ {data.get('nickname', line_user_id)} สำเร็จ")
        return True

    except Exception as e:
        print(f"❌ Error saving to sheets: {e}")
        return False
