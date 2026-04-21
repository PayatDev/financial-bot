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

