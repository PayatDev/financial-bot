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
from linebot.v3.webhooks import MessageEvent, TextMessageContent, FollowEvent
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

from nong_draft import run as draft_run
from nong_doc import find_folder_id, download_xlsx, read_xlsx
from drive_service import upload_file_to_folder
 
from cover_builder import generate_issue_content, build_cover
from will_builder import build_will
from asset_registry_builder import build_asset_registry
from living_will_builder import build_living_will
from emergency_guide_builder import build_emergency_guide

import threading

GREETING_MESSAGE = """สวัสดีค่ะ ดิฉันน้องแพลน
ผู้ช่วยของคุณพยัตค่ะ 😊

วันนี้เราจะคุยกันเพื่อเก็บข้อมูลสำคัญของครอบครัวคุณ

คุณพยัตจะนำข้อมูลที่เราคุยกันวันนี้ ไปจัดทำแผน
ที่คุ้มครองคู่ชีวิตและลูกๆของคุณให้รอบด้าน

ขอทราบชื่อเล่นของคุณหน่อยได้ไหมคะ?"""

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

MAX_TURNS = 120  # ป้องกันค่าใช้จ่ายบาน

EXPECTED_FIELDS = [
                    "email", "nickname", "age", "gender", "occupation",
                    "health", "income_self", "hobbies_and_risks",
                    "spouse_nickname", "spouse_age", "spouse_occupation",
                    "spouse_income", "spouse_health", "spouse_status",
                    "children", "children_outside_marriage",
                    "guardian_primary", "guardian_backup",
                    "money_guardian_primary", "money_guardian_backup",
                    "estate_executor", "urgent_manager",
                    "asset_distribution", "surviving_spouse_plan",
                    "debt_responsibility", "business_succession",
                    "living_will", "financial_poa",
                    "assets_cash", "assets_property", "assets_business",
                    "assets_investment", "assets_insurance_savings",
                    "insurance_life", "insurance_health", "insurance_group",
                    "welfare", "assets_crypto_wallet", "assets_digital",
                    "assets_valuables", "debt", "guarantor",
                    "emergency_cash_90days", "estate_admin_cost",
                    "funeral_wishes", "documents_location",
                    "fullname_self", "id_self", "address_self",
                    "fullname_spouse", "id_spouse",
                    "fullname_children",
                    "fullname_executor", "id_executor",
                    "fullname_executor_backup", "id_executor_backup",
                    "fullname_guardian_primary", "id_guardian_primary",
                    "fullname_guardian_backup", "id_guardian_backup",
                    "fullname_money_guardian_primary", "id_money_guardian_primary",
                    "fullname_money_guardian_backup", "id_money_guardian_backup",
                    "gaps_for_payat", "summary",
                    ]

# เก็บสถานะ completed ใน memory + session store
# key = user_id, value = True
COMPLETED_USERS: dict = {}

CONTACT_MESSAGE = (
    "ดิฉันมีหน้าที่เก็บข้อมูลเพียงอย่างเดียวค่ะ\n"
    "หากมีอะไรสอบถามเพิ่มเติม ติดต่อคุณพยัตได้เลยนะค่ะ 😊\n\n"
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

# ── /test-draft → xlsx เท่านั้น ───────────────────────────────────────
@app.get("/test-draft")
def test_draft():
    threading.Thread(
        target=draft_run, kwargs={"run_doc": False}, daemon=True
    ).start()
    return {"status": "started — xlsx only"}
 
 
# ── helper: โหลด client_data จาก folder ──────────────────────────────
def _load_client_data(folder: str) -> tuple:
    folder_id   = find_folder_id(folder)
    xlsx_path   = download_xlsx(folder_id)
    client_data = read_xlsx(xlsx_path)
    import os; os.unlink(xlsx_path)
    return folder_id, client_data
 
 
# ── /test-cover?folder=xxx ────────────────────────────────────────────
def _run_cover(folder: str):
    folder_id, client_data = _load_client_data(folder)
    nickname  = client_data.get("ชื่อเล่น", folder.split("_")[0])
    print("กำลัง generate cover...")
    generated = generate_issue_content(client_data)
    path = build_cover(client_data, [], generated, folder)
    filename = f"1_แผนครอบครัว_คุณ{nickname}.docx"
    upload_file_to_folder(path, filename, folder_id)
    import os; os.path.exists(path) and os.unlink(path)
    print(f"✅ cover พร้อม → {filename}")
 
@app.get("/test-cover")
def test_cover(folder: str):
    threading.Thread(target=_run_cover, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}
 
 
# ── /test-will?folder=xxx ─────────────────────────────────────────────
def _run_will(folder: str):
    folder_id, client_data = _load_client_data(folder)
    nickname = client_data.get("ชื่อเล่น", folder.split("_")[0])
    path = build_will(client_data)
    filename = f"2_พินัยกรรม_คุณ{nickname}.docx"
    upload_file_to_folder(path, filename, folder_id)
    import os; os.path.exists(path) and os.unlink(path)
    print(f"✅ will พร้อม → {filename}")
 
@app.get("/test-will")
def test_will(folder: str):
    threading.Thread(target=_run_will, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}
 
 
# ── /test-living?folder=xxx ───────────────────────────────────────────
def _run_living(folder: str):
    folder_id, client_data = _load_client_data(folder)
    nickname = client_data.get("ชื่อเล่น", folder.split("_")[0])
    path = build_living_will(client_data)
    filename = f"4_หนังสือแสดงเจตนาการยื้อชีวิต_คุณ{nickname}.docx"
    upload_file_to_folder(path, filename, folder_id)
    import os; os.path.exists(path) and os.unlink(path)
    print(f"✅ living will พร้อม → {filename}")
 
@app.get("/test-living")
def test_living(folder: str):
    threading.Thread(target=_run_living, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}
 
 
# ── /test-guide?folder=xxx ────────────────────────────────────────────
def _run_guide(folder: str):
    folder_id, client_data = _load_client_data(folder)
    nickname = client_data.get("ชื่อเล่น", folder.split("_")[0])
    path = build_emergency_guide(client_data)
    filename = f"5_คู่มือฉุกเฉิน_คุณ{nickname}.docx"
    upload_file_to_folder(path, filename, folder_id)
    import os; os.path.exists(path) and os.unlink(path)
    print(f"✅ guide พร้อม → {filename}")
 
@app.get("/test-guide")
def test_guide(folder: str):
    threading.Thread(target=_run_guide, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}

def _run_asset(folder: str):
    folder_id, client_data = _load_client_data(folder)
    nickname = client_data.get("ชื่อเล่น", folder.split("_")[0])
    path = build_asset_registry(client_data)
    filename = f"3_บัญชีทรัพย์สิน_คุณ{nickname}.docx"
    upload_file_to_folder(path, filename, folder_id)
    import os; os.path.exists(path) and os.unlink(path)
    print(f"✅ asset registry พร้อม → {filename}")

@app.get("/test-asset")
def test_asset(folder: str):
    threading.Thread(target=_run_asset, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}

@app.get("/run-doc")
def run_doc(folder: str):
    threading.Thread(target=doc_run, args=(folder,), daemon=True).start()
    return {"status": "started", "folder": folder}


@app.get("/force-save")
def force_save():
    """ทดสอบ save ลง sheet โดยไม่ต้องคุยกับน้องแพลน"""
    test_data = {
        "email": "test@test.com",
        "nickname": "ทดสอบ",
        "age": "35",
        "gender": "ชาย",
        "occupation": "พนักงานบริษัท",
        "health": "ดี",
        "income_self": "50,000 บาท/เดือน",
        "hobbies_and_risks": "วิ่ง ดูหนัง",
        "spouse_nickname": "แพลน",
        "spouse_age": "32",
        "spouse_occupation": "ครู",
        "spouse_income": "30,000 บาท/เดือน",
        "spouse_health": "ดี",
        "spouse_status": "อยู่ด้วยกัน",
        "children": "น้องทดสอบ 3 ขวบ สุขภาพดี",
        "children_outside_marriage": "ไม่มี",
        "assets_cash": "500,000 บาท",
        "assets_property": "บ้าน 3 ล้าน ปลอดหนี้",
        "assets_investment": "หุ้น 200,000 บาท",
        "assets_crypto_wallet": "ไม่มี",
        "assets_insurance_savings": "ไม่มี",
        "assets_digital": "ไม่มี",
        "assets_business": "ไม่มี",
        "assets_valuables": "รถ 500,000 บาท",
        "debt": "หนี้รถ 200,000 บาท",
        "guarantor": "ไม่มี",
        "insurance_life": "ประกันชีวิต 1 ล้าน",
        "insurance_health": "ประกันสุขภาพเหมาจ่าย",
        "insurance_group": "ไม่มี",
        "welfare": "ประกันสังคม",
        "funeral_wishes": "พิธีพุทธ งบ 100,000 บาท",
        "emergency_cash_90days": "ใช้เงินออม",
        "estate_admin_cost": "ใช้เงินออม",
        "asset_distribution": "ให้ภรรยาทั้งหมด",
        "debt_responsibility": "ภรรยารับผิดชอบ",
        "business_succession": "ไม่มีกิจการ",
        "urgent_manager": "พี่ชาย",
        "estate_executor": "ภรรยา",
        "documents_location": "ตู้เซฟที่บ้าน ภรรยารู้",
        "financial_poa": "ภรรยา",
        "living_will": "ไม่ยื้อชีวิต ภรรยาตัดสินใจ",
        "surviving_spouse_plan": "ไม่กังวล",
        "guardian_primary": "พี่ชาย ยินดี",
        "guardian_backup": "น้องสาวภรรยา",
        "money_guardian_primary": "น้องสาวภรรยา",
        "money_guardian_backup": "ไม่ได้ระบุ",
        "fullname_self": "ทดสอบ ระบบ",
        "id_self": "1234567890123",
        "address_self": "123 ถนนทดสอบ กรุงเทพ 10100",
        "fullname_spouse": "แพลน ระบบ",
        "id_spouse": "ไม่ได้ระบุ",
        "fullname_children": "ทดสอบน้อย ระบบ",
        "fullname_executor": "แพลน ระบบ",
        "id_executor": "ไม่ได้ระบุ",
        "fullname_executor_backup": "ไม่ได้ระบุ",
        "id_executor_backup": "ไม่ได้ระบุ",
        "fullname_guardian_primary": "พี่ชาย ระบบ",
        "id_guardian_primary": "ไม่ได้ระบุ",
        "fullname_guardian_backup": "ไม่ได้ระบุ",
        "id_guardian_backup": "ไม่ได้ระบุ",
        "fullname_money_guardian_primary": "ไม่ได้ระบุ",
        "id_money_guardian_primary": "ไม่ได้ระบุ",
        "fullname_money_guardian_backup": "ไม่ได้ระบุ",
        "id_money_guardian_backup": "ไม่ได้ระบุ",
        "gaps_for_payat": "ทดสอบระบบ",
        "summary": "ข้อมูลทดสอบระบบ force-save"
    }
    result = save_to_sheets("test_force_save", test_data)
    if result:
        return {"status": "✅ บันทึกลง Sheet สำเร็จ"}
    else:
        return {"status": "❌ บันทึกไม่สำเร็จ ดู Railway logs"}


@app.get("/chat-log/{user_id}")
def chat_log(user_id: str):
    history = get_history(user_id)
    return {"user_id": user_id, "messages": history}
    
@app.post("/webhook")
async def webhook(request: Request):
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()
    try:
        handler.handle(body.decode(), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")
    return {"status": "ok"}

@handler.add(FollowEvent)
def handle_follow(event: FollowEvent):
    user_id = event.source.user_id
    clear_session(user_id)  # เคลียร์ session เผื่อเคย reset แล้วแอดใหม่
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=GREETING_MESSAGE)],
                )
            )
    except Exception as e:
        print(f"❌ Follow reply error: {e}")

def reply_to_line(event, text: str):
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=text)],
                )
            )
    except Exception as e:
        print(f"❌ Reply token expired or error: {e}")


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
        reply_to_line(event, f"รับทราบค่ะ กำลังสร้างเอกสารสำหรับ {folder_name}\nรอสักครู่ แล้วจะแจ้งกลับค่ะ 📄")
        return

    # reset command
    if user_message.strip().lower() in RESET_COMMANDS:
        clear_session(user_id)
        COMPLETED_USERS.pop(user_id, None)
        reply_to_line(event, "ล้างข้อมูลเรียบร้อยแล้วค่ะ พิมพ์อะไรก็ได้เพื่อเริ่มบทสนทนาใหม่ 😊")
        return

    # สถานะที่ 2: save เสร็จแล้ว — ไม่เรียก Claude
    if is_completed(user_id):
        reply_to_line(event, CONTACT_MESSAGE)
        return

    if "@" in user_message and "." in user_message:
        reply_to_line(event, 
            "ได้รับ email แล้วค่ะ 😊\n"
            "กำลังบันทึกข้อมูล คุณพยัตจะส่งเอกสารให้เร็วๆ นี้นะคะ"
        )
        history = get_history(user_id)
        threading.Thread(
            target=process_save,
            args=(user_id, history, user_message),
            daemon=True
        ).start()
        return

    # สถานะที่ 1: กำลังสัมภาษณ์อยู่
    history = get_history(user_id)

    if len(history) >= MAX_TURNS * 2:
        reply_to_line(event,
            "ขออภัยค่ะ ดิฉันขออนุญาตจบการสนทนานี้นะค่ะ\n"
            "กรุณาติดต่อคุณพยัตนะค่ะ 😊\n\n"
            "📧 payat.jira@gmail.com"
        )
        return

    try:
        bot_reply = chat(user_id, history, user_message)
    except Exception as e:
        print(f"❌ Claude API error: {e}")
        reply_to_line(event, "ขออภัยค่ะ ระบบขัดข้องชั่วคราว กรุณาลองใหม่อีกครั้งในอีกสักครู่นะค่ะ 🙏")
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
            reply_text = "ขออภัยค่ะ ระบบบันทึกข้อมูลมีปัญหา กรุณาลองใหม่อีกครั้งนะค่ะ 🙏"

            update_history(user_id, "user", user_message)
            update_history(user_id, "assistant", bot_reply)
    else:
        reply_text = bot_reply
        update_history(user_id, "user", user_message)
        update_history(user_id, "assistant", bot_reply)

    reply_to_line(event, reply_text)


def process_save(user_id, history, user_message):
    try:
        bot_reply = chat(user_id, history, user_message)
        if "[SAVE_DATA]" in bot_reply and "[END_SAVE_DATA]" in bot_reply:
            raw = bot_reply.split("[SAVE_DATA]")[1].split("[END_SAVE_DATA]")[0].strip()
            raw = raw.replace("```json", "").replace("```", "").strip()
            data = json.loads(raw)
            data = {k: data.get(k, "ไม่ได้ระบุ") for k in EXPECTED_FIELDS}
            save_to_sheets(user_id, data)
            threading.Thread(target=draft_run, args=(data,), daemon=True).start()
            COMPLETED_USERS[user_id] = True
            update_history(user_id, "user", user_message)
            update_history(user_id, "assistant", "[COMPLETED]")
    except Exception as e:
        print(f"❌ process_save error: {e}")
