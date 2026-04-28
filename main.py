# main.py

import json, os
from fastapi import FastAPI, Request, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
from linebot.v3.exceptions import InvalidSignatureError

from config import LINE_CHANNEL_SECRET, LINE_CHANNEL_ACCESS_TOKEN
from nong_plan import chat
from session_store import get_history, update_history, clear_session
from sheets_service import save_to_sheets

from drive_service import get_drive_service
from nong_draft import run as draft_run
from nong_doc import run as doc_run

from order_service import save_order
from pydantic import BaseModel

import threading

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://payatdev.github.io"],
    allow_methods=["POST", "GET"],
    allow_headers=["Content-Type"],
)

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

RESET_COMMANDS = ["/reset", "reset", "/เริ่มใหม่", "เริ่มใหม่"]

MAX_TURNS = 80  # ป้องกันค่าใช้จ่ายบาน

# เก็บสถานะ completed ใน memory + session store
# key = user_id, value = True
COMPLETED_USERS: dict = {}

CONTACT_MESSAGE = (
    "น้องแพลนมีหน้าที่เก็บข้อมูลเพียงอย่างเดียวครับ\n"
    "หากมีอะไรสอบถามเพิ่มเติม ติดต่อคุณพยัตได้เลยนะครับ 😊\n\n"
    "📧 Email: payat.jira@gmail.com"
)

# เพิ่ม model
class OrderData(BaseModel):
    name: str
    email: str
    phone: str = ""
    payment: str
    note: str = ""
    
@app.get("/")
def root():
    return {"status": "Financial Bot is running! 🤖"}

@app.post("/order")
async def create_order(data: OrderData):
    try:
        save_order(data.dict())
        return {"status": "ok"}
    except Exception as e:
        print(f"❌ Order error: {e}")
        raise HTTPException(status_code=500, detail="บันทึกไม่สำเร็จ")

@app.get("/test-drive")
def test_drive():
    service = get_drive_service()
    results = service.files().list(
        q=f"'{os.environ.get('GOOGLE_DRIVE_FOLDER_ID')}' in parents",
        fields="files(id, name)"
    ).execute()
    files = results.get("files", [])
    return {"files": files, "count": len(files)}

@app.get("/test-draft")
def test_draft():
    threading.Thread(target=draft_run, daemon=True).start()
    return {"status": "started — ดู log ใน Railway"}

@app.get("/run-doc")
def run_doc(folder: str):
    threading.Thread(target=doc_run, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}

@app.post("/webhook")
async def webhook(request: Request):
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()
    try:
        handler.handle(body.decode(), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")
    return {"status": "ok"}


def reply_to_line(event, text: str):
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=text)],
            )
        )


def is_completed(user_id: str) -> bool:
    """เช็คว่า user นี้ส่งข้อมูลแล้วหรือยัง
    เก็บไว้ใน history เป็น marker เพื่อให้คงอยู่หลัง restart"""
    if user_id in COMPLETED_USERS:
        return True
    # fallback: เช็คใน history ว่ามี COMPLETED marker ไหม
    history = get_history(user_id)
    for msg in history:
        if msg.get("role") == "assistant" and "[COMPLETED]" in msg.get("content", ""):
            COMPLETED_USERS[user_id] = True
            return True
    return False


@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event: MessageEvent):
    user_id = event.source.user_id
    user_message = event.message.text

    # /doc command — ต้องอยู่ก่อน reset และก่อน completed check
    if user_message.strip().startswith("/doc "):
        folder_name = user_message.strip()[5:].strip()
        threading.Thread(target=doc_run, args=(folder_name,), daemon=True).start()
        reply_to_line(event, f"รับทราบครับ กำลังสร้างเอกสารสำหรับ {folder_name}\nรอสักครู่ แล้วจะแจ้งกลับครับ 📄")
        return

    # reset command
    if user_message.strip().lower() in RESET_COMMANDS:
        clear_session(user_id)
        COMPLETED_USERS.pop(user_id, None)
        reply_to_line(event, "ล้างข้อมูลเรียบร้อยแล้วครับ พิมพ์อะไรก็ได้เพื่อเริ่มบทสนทนาใหม่ 😊")
        return

    # สถานะที่ 2: save เสร็จแล้ว — ไม่เรียก Claude
    if is_completed(user_id):
        reply_to_line(event, CONTACT_MESSAGE)
        return

    # สถานะที่ 1: กำลังสัมภาษณ์อยู่
    history = get_history(user_id)
    
    if len(history) >= MAX_TURNS * 2:
        reply_to_line(event,
            "ขออภัยครับ ผมขออนุญาตจบการสนทนานี้นะครับ\n"
            "กรุณาติดต่อคุณพยัตโดยตรงนะครับ 😊\n\n"
            "📧 payat.jira@gmail.com"
        )
        return
        
    try:
        bot_reply = chat(user_id, history, user_message)
    except Exception as e:
        print(f"❌ Claude API error: {e}")
        reply_to_line(event, "ขออภัยครับ ระบบขัดข้องชั่วคราว กรุณาลองใหม่อีกครั้งในอีกสักครู่นะครับ 🙏")
        return

    # เช็ค SAVE_DATA
    if "[SAVE_DATA]" in bot_reply and "[END_SAVE_DATA]" in bot_reply:
        reply_text = bot_reply.split("[SAVE_DATA]")[0].strip()
        try:
            raw = bot_reply.split("[SAVE_DATA]")[1].split("[END_SAVE_DATA]")[0].strip()

            raw = raw.replace("```json", "").replace("```", "").strip()
    
            data = json.loads(raw)  # 🔥 จุดสำคัญ

            # ✅ กันพัง
            if not isinstance(data, dict):
                raise ValueError("SAVE_DATA is not dict")

            EXPECTED_FIELDS = [
                                "email", "nickname", "age", "gender", "occupation",
                                "income_self", "spouse_nickname", "spouse_age",
                                "spouse_occupation", "spouse_income", "spouse_health",
                                "spouse_status", "children", "guardian_primary",
                                "guardian_backup", "money_guardian_primary",
                                "money_guardian_backup", "estate_executor",
                                "urgent_manager", "asset_distribution",
                                "surviving_spouse_plan", "debt_responsibility",
                                "business_succession", "living_will", "financial_poa",
                                "assets_cash", "assets_property", "assets_business",
                                "assets_investment", "assets_insurance_savings",
                                "insurance_life", "insurance_health", "insurance_group",
                                "welfare", "assets_crypto_wallet", "assets_digital",
                                "assets_valuables", "debt", "guarantor",
                                "hobbies_and_risks", "emergency_cash_90days",
                                "estate_admin_cost", "funeral_wishes",
                                "documents_location", "health", "children_outside_marriage"
                                ]

            data = {k: data.get(k, "ไม่ได้ระบุ") for k in EXPECTED_FIELDS}
    
            save_to_sheets(user_id, data)
    
            threading.Thread(target=draft_run, args=(data,), daemon=True).start()
    
            COMPLETED_USERS[user_id] = True
            update_history(user_id, "user", user_message)
            update_history(user_id, "assistant", "[COMPLETED]")
    
            print(f"✅ {user_id} complete — {data.get('nickname', '')}")
    
        except Exception as e:
            print(f"❌ JSON parse error: {e}")
            print(f"RAW SAVE_DATA:\n{raw}")
            reply_text = "ขออภัยครับ ระบบบันทึกข้อมูลมีปัญหา กรุณาลองใหม่อีกครั้งนะครับ 🙏"
    
            update_history(user_id, "user", user_message)
            update_history(user_id, "assistant", bot_reply)
    else:
        reply_text = bot_reply
        update_history(user_id, "user", user_message)
        update_history(user_id, "assistant", bot_reply)

    reply_to_line(event, reply_text)
