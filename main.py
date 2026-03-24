# main.py
# จุดเริ่มต้นของ app ทั้งหมด

import json
import re
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

# LINE SDK setup
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)


@app.get("/")
def root():
    """Health check — Railway ใช้เช็คว่า app ยังทำงานอยู่"""
    return {"status": "Financial Bot is running! 🤖"}


@app.post("/webhook")
async def webhook(request: Request):
    """รับ webhook จาก LINE"""
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()

    try:
        handler.handle(body.decode(), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")

    return {"status": "ok"}


@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event: MessageEvent):
    """จัดการข้อความที่รับมาจาก LINE"""
    user_id = event.source.user_id
    user_message = event.message.text

    # 1. โหลด conversation history ของ user คนนี้
    history = get_history(user_id)

    # 2. ส่งไปให้ Claude
    bot_reply = chat(user_id, history, user_message)

    # 3. เช็คว่า Claude ส่ง [SAVE_DATA] มาหรือเปล่า
    if "[SAVE_DATA]" in bot_reply:
        # แยกข้อความปกติออกจาก JSON
        parts = bot_reply.split("[SAVE_DATA]")
        reply_text = parts[0].strip()  # ข้อความที่ส่งให้ลูกค้า

        # พยายาม parse JSON
        try:
            json_str = parts[1].strip()
            data = json.loads(json_str)
            save_to_sheets(user_id, data)
            clear_session(user_id)  # ล้าง session หลังบันทึกแล้ว
        except (json.JSONDecodeError, IndexError) as e:
            print(f"Error parsing SAVE_DATA: {e}")
            reply_text = bot_reply.replace("[SAVE_DATA]", "").strip()
    else:
        reply_text = bot_reply

    # 4. อัพเดท history
    update_history(user_id, "user", user_message)
    update_history(user_id, "assistant", bot_reply)

    # 5. ส่ง reply กลับ LINE
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)],
            )
        )
