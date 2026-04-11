import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")


def send_email(to: str, subject: str, body: str):
    msg = MIMEMultipart()
    msg["From"] = GMAIL_USER
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, to, msg.as_string())
        print(f"✅ Email sent → {to} | {subject}")


def notify_new_client(nickname: str, occupation: str, age: str):
    subject = f"[ลูกค้าใหม่] {nickname} | {occupation} | {age} ปี"
    body = (
        f"มีลูกค้าใหม่เข้ามาครับ\n\n"
        f"ชื่อ: {nickname}\n"
        f"อาชีพ: {occupation}\n"
        f"อายุ: {age} ปี\n\n"
        f"เปิด Google Drive เพื่อดูไฟล์ได้เลยครับ"
    )
    send_email(GMAIL_USER, subject, body)
