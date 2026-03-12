import os
import time
import requests
from playwright.sync_api import sync_playwright

# ---- Config from environment variables (GitHub Secrets) ----
LOOKER_URL      = os.environ["LOOKER_STUDIO_URL"]   # URL ของ Report (หน้า 14)
GOOGLE_EMAIL    = os.environ["GOOGLE_EMAIL"]
GOOGLE_PASSWORD = os.environ["GOOGLE_PASSWORD"]
LINE_TOKEN      = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
LINE_GROUP_ID   = os.environ["LINE_GROUP_ID"]

SCREENSHOT_PATH = "report_page14.png"


def take_screenshot():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = browser.new_context(
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()

        # --- Login Google ---
        print("🔐 Logging in to Google...")
        page.goto("https://accounts.google.com/signin")
        page.fill('input[type="email"]', GOOGLE_EMAIL)
        page.click('button:has-text("Next"), #identifierNext')
        page.wait_for_timeout(2000)
        page.fill('input[type="password"]', GOOGLE_PASSWORD)
        page.click('button:has-text("Next"), #passwordNext')
        page.wait_for_timeout(4000)

        # --- Open Looker Studio ---
        print(f"📊 Opening Looker Studio: {LOOKER_URL}")
        page.goto(LOOKER_URL, wait_until="networkidle", timeout=60000)
        time.sleep(8)  # รอให้ชาร์ตโหลดครบ

        # --- Screenshot ---
        print("📸 Taking screenshot...")
        page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        print(f"✅ Screenshot saved: {SCREENSHOT_PATH}")

        browser.close()


def upload_image_to_line(image_path: str) -> str:
    """Upload image และรับ URL กลับมา (ใช้ imgbb หรือ LINE rich message)"""
    # ใช้ LINE multipart upload สำหรับ Bot
    # วิธีง่ายสุด: ส่งเป็น imageMessage ผ่าน imgbb (ฟรี)
    imgbb_key = os.environ.get("IMGBB_API_KEY", "")
    if imgbb_key:
        with open(image_path, "rb") as f:
            import base64
            b64 = base64.b64encode(f.read()).decode()
        r = requests.post(
            "https://api.imgbb.com/1/upload",
            data={"key": imgbb_key, "image": b64}
        )
        r.raise_for_status()
        return r.json()["data"]["url"]
    return None


def send_to_line():
    print("📤 Sending to LINE group...")

    image_url = upload_image_to_line(SCREENSHOT_PATH)

    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }

    if image_url:
        # ส่งเป็น imageMessage (แสดงภาพในแชทได้เลย)
        payload = {
            "to": LINE_GROUP_ID,
            "messages": [
                {
                    "type": "image",
                    "originalContentUrl": image_url,
                    "previewImageUrl": image_url
                },
                {
                    "type": "text",
                    "text": "📊 Looker Studio Report - หน้า 14\n🕓 อัปเดต 16:00 น."
                }
            ]
        }
    else:
        # Fallback: ส่งแค่ข้อความ + ลิงก์
        payload = {
            "to": LINE_GROUP_ID,
            "messages": [
                {
                    "type": "text",
                    "text": f"📊 Looker Studio Report - หน้า 14\n🕓 อัปเดต 16:00 น.\n🔗 {LOOKER_URL}"
                }
            ]
        }

    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers=headers,
        json=payload
    )
    resp.raise_for_status()
    print(f"✅ Sent to LINE! Status: {resp.status_code}")


if __name__ == "__main__":
    take_screenshot()
    send_to_line()