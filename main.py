# main.py
# จุดเริ่มต้นของ app ทั้งหมด

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

    # คำสั่ง reset
    if user_message.strip().lower() in RESET_COMMANDS:
        clear_session(user_id)
        reply_to_line(event, "ล้างข้อมูลเรียบร้อยแล้วครับ พิมพ์อะไรก็ได้เพื่อเริ่มบทสนทนาใหม่ 😊")
        return

    # 1. โหลด history
    history = get_history(user_id)

    # 2. ส่งให้ Claude
    bot_reply = chat(user_id, history, user_message)

    # 3. เช็ค SAVE_DATA
    if "[SAVE_DATA]" in bot_reply:
        parts = bot_reply.split("[SAVE_DATA]")
        reply_text = parts[0].strip()
        try:
            json_str = parts[1].strip()
            data = json.loads(json_str)
            save_to_sheets(user_id, data)
            clear_session(user_id)
        except (json.JSONDecodeError, IndexError) as e:
            print(f"Error parsing SAVE_DATA: {e}")
            reply_text = bot_reply.replace("[SAVE_DATA]", "").strip()
    else:
        reply_text = bot_reply

    # 4. อัพเดท history
    update_history(user_id, "user", user_message)
    update_history(user_id, "assistant", bot_reply)

    # 5. ส่งกลับ LINE
    reply_to_line(event, reply_text)
