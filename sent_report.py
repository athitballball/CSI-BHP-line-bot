import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright

# ---- Config from environment variables (GitHub Secrets) ----
LOOKER_URL      = os.environ["LOOKER_STUDIO_URL"]
GOOGLE_EMAIL    = os.environ["GOOGLE_EMAIL"]
GOOGLE_PASSWORD = os.environ["GOOGLE_PASSWORD"]
LINE_TOKEN      = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
LINE_GROUP_ID   = os.environ["LINE_GROUP_ID"]
GITHUB_TOKEN    = os.environ["GITHUB_TOKEN"]        # มีให้อัตโนมัติใน Actions
GITHUB_REPO     = os.environ["GITHUB_REPOSITORY"]  # เช่น "username/repo-name"

SCREENSHOT_PATH  = "report_page14.png"
GITHUB_FILE_PATH = "report_page14.png"  # path ใน repo


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


def upload_to_github() -> str:
    """Commit ภาพลง repo แล้วคืน raw URL"""
    print("📤 Uploading image to GitHub repo...")

    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    # ตรวจว่าไฟล์มีอยู่แล้วหรือไม่ (ถ้ามีต้องส่ง sha ด้วย)
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json()["sha"]

    payload = {
        "message": "update report screenshot",
        "content": content_b64,
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    print("✅ Uploaded to GitHub!")

    # raw URL สำหรับส่งให้ LINE (เพิ่ม timestamp ป้องกัน cache)
    raw_url = (
        f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FILE_PATH}"
        f"?t={int(time.time())}"
    )
    return raw_url


def send_to_line(image_url: str):
    print("💬 Sending to LINE group...")

    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }

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

    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers=headers,
        json=payload
    )
    resp.raise_for_status()
    print(f"✅ Sent to LINE! Status: {resp.status_code}")


if __name__ == "__main__":
    take_screenshot()
    image_url = upload_to_github()
    send_to_line(image_url)