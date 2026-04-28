# order_service.py
import os
import json
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ORDER_HEADERS = [
    "timestamp", "name", "email", "phone", "payment", "note", "status"
]

def get_order_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON not found")
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = os.environ.get("ORDER_SHEET_ID")
    return client.open_by_key(sheet_id).sheet1

def ensure_order_headers(sheet):
    first_row = sheet.row_values(1)
    if not first_row or first_row[0] != "timestamp":
        sheet.clear()
        sheet.insert_row(ORDER_HEADERS, index=1)
        sheet.format("1:1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.8},
        })

def save_order(data: dict):
    try:
        sheet = get_order_sheet()
        ensure_order_headers(sheet)
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data.get("name", ""),
            data.get("email", ""),
            data.get("phone", ""),
            data.get("payment", ""),
            data.get("note", ""),
            "รอยืนยัน",
        ]
        sheet.append_row(row, value_input_option="USER_ENTERED")
        print(f"✅ Order saved: {data.get('name', '')}")
        return True
    except Exception as e:
        print(f"❌ Order error: {e}")
        return False
