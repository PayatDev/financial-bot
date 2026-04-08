# main.py

import json
from fastapi import FastAPI, Request, HTTPException
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
from claude_service import chat
from session_store import get_history, update_history, clear_session
from sheets_service import save_to_sheets

app = FastAPI()

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

RESET_COMMANDS = ["/reset", "reset", "/เริ่มใหม่", "เริ่มใหม่"]

# เก็บสถานะว่า user ไหน save เสร็จแล้ว
COMPLETED_USERS = set()

CONTACT_MESSAGE = (
    "น้องแพลนมีหน้าที่เก็บข้อมูลเพียงอย่างเดียวครับ\n"
    "หากมีอะไรสอบถามเพิ่มเติม ติดต่อคุณพยัตได้เลยนะครับ 😊\n\n"
    "📧 Email: payat@example.com\n"
    "📱 Line: @payat"
)


@app.get("/")
def root():
    return {"status": "Financial Bot is running! 🤖"}


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


@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event: MessageEvent):
    user_id = event.source.user_id
    user_message = event.message.text

    # reset command
    if user_message.strip().lower() in RESET_COMMANDS:
        clear_session(user_id)
        COMPLETED_USERS.discard(user_id)
        reply_to_line(event, "ล้างข้อมูลเรียบร้อยแล้วครับ พิมพ์อะไรก็ได้เพื่อเริ่มบทสนทนาใหม่ 😊")
        return

    # สถานะที่ 2: save เสร็จแล้ว — ไม่เรียก Claude อีก ประหยัด cost
    if user_id in COMPLETED_USERS:
        reply_to_line(event, CONTACT_MESSAGE)
        return

    # สถานะที่ 1: กำลังสัมภาษณ์อยู่
    history = get_history(user_id)
    bot_reply = chat(user_id, history, user_message)

    # เช็ค SAVE_DATA
    if "[SAVE_DATA]" in bot_reply:
        parts = bot_reply.split("[SAVE_DATA]")
        reply_text = parts[0].strip()
        try:
            json_str = parts[1].strip()
            data = json.loads(json_str)
            save_to_sheets(user_id, data)
            # mark user ว่า complete แล้ว ไม่ clear session
            COMPLETED_USERS.add(user_id)
            print(f"✅ {user_id} complete")
        except (json.JSONDecodeError, IndexError) as e:
            print(f"Error parsing SAVE_DATA: {e}")
            reply_text = bot_reply.replace("[SAVE_DATA]", "").strip()
    else:
        reply_text = bot_reply

    # อัพเดท history
    update_history(user_id, "user", user_message)
    update_history(user_id, "assistant", bot_reply)

    # ส่งกลับ LINE
    reply_to_line(event, reply_text)
