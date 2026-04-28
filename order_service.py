# order_service.py
import os
from datetime import datetime
from sheets_service import get_sheets_service  # ใช้ service เดิม

ORDER_SHEET_ID = os.environ.get("ORDER_SHEET_ID")
ORDER_SHEET_NAME = "Orders"

def save_order(data: dict):
    service = get_sheets_service()
    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        data.get("name", ""),
        data.get("email", ""),
        data.get("phone", ""),
        data.get("payment", ""),
        data.get("note", ""),
        "รอยืนยัน"  # status
    ]
    service.spreadsheets().values().append(
        spreadsheetId=ORDER_SHEET_ID,
        range=f"{ORDER_SHEET_NAME}!A1",
        valueInputOption="RAW",
        body={"values": [row]}
    ).execute()
