import os
import urllib.request
import json

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
PAYAT_USER_ID = "U86c03cd5153459d2dc9ce52adc608147"

def send_line_message(text: str):
    data = json.dumps({
        "to": PAYAT_USER_ID,
        "messages": [{"type": "text", "text": text}]
    }).encode()

    req = urllib.request.Request(
        "https://api.line.me/v2/bot/message/push",
        data=data,
        headers={
            "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
            "Content-Type": "application/json"
        },
        method="POST"
    )
    urllib.request.urlopen(req)
    print(f"✅ LINE sent → {PAYAT_USER_ID}")


def notify_new_client(nickname: str, occupation: str, age: str, folder_name=folder_name, folder_name: str = ""):
    text = (
        f"🆕 ลูกค้าใหม่ครับ\n\n"
        f"ชื่อ: {nickname}\n"
        f"อาชีพ: {occupation}\n"
        f"อายุ: {age} ปี\n\n"
        f"📁 folder: {folder_name}\n\n"
        f"กรอก col G ใน Tab 2 แล้วพิม\n"
        f"/doc {folder_name}"
    )
    send_line_message(text)
