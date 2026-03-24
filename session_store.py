# session_store.py
# เก็บ conversation history ของแต่ละ LINE user ไว้ใน memory
# ถ้า restart server จะหายหมด (เหมาะกับ v.1)

from datetime import datetime, timedelta

# โครงสร้าง: { "line_user_id": { "history": [...], "last_active": datetime } }
sessions = {}

SESSION_TIMEOUT_HOURS = 24


def get_history(user_id: str) -> list:
    """ดึง conversation history ของ user คนนี้"""
    session = sessions.get(user_id)
    if not session:
        return []

    # เช็คว่า timeout หรือยัง
    if datetime.now() - session["last_active"] > timedelta(hours=SESSION_TIMEOUT_HOURS):
        clear_session(user_id)
        return []

    return session["history"]


def update_history(user_id: str, role: str, content: str):
    """เพิ่มข้อความใหม่เข้า history
    role = "user" หรือ "assistant"
    """
    if user_id not in sessions:
        sessions[user_id] = {"history": [], "last_active": datetime.now()}

    sessions[user_id]["history"].append({"role": role, "content": content})
    sessions[user_id]["last_active"] = datetime.now()


def clear_session(user_id: str):
    """ลบ session ของ user คนนี้ (เช่น หลัง save data แล้ว)"""
    if user_id in sessions:
        del sessions[user_id]
