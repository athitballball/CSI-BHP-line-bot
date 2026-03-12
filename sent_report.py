import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright

# ---- Config from environment variables (GitHub Secrets) ----
LOOKER_URL      = os.environ.get("LOOKER_STUDIO_URL")
GOOGLE_EMAIL    = os.environ.get("GOOGLE_EMAIL")
GOOGLE_PASSWORD = os.environ.get("GOOGLE_PASSWORD")
LINE_TOKEN      = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_GROUP_ID   = os.environ.get("LINE_GROUP_ID")
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN")        
GITHUB_REPO     = os.environ.get("GITHUB_REPOSITORY")  

SCREENSHOT_PATH  = "report_page14.png"
GITHUB_FILE_PATH = "report_page14.png"

def take_screenshot():
    with sync_playwright() as p:
        print("🚀 Launching browser...")
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            viewport={"width": 1440, "height": 1080}, # เพิ่มความสูงให้เห็นชัดขึ้น
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # --- Login Google ---
        print("🔐 Logging in to Google...")
        try:
            page.goto("https://accounts.google.com/signin", timeout=60000)
            page.fill('input[type="email"]', GOOGLE_EMAIL)
            page.click('button:has-text("Next"), #identifierNext')
            page.wait_for_timeout(3000)
            page.fill('input[type="password"]', GOOGLE_PASSWORD)
            page.click('button:has-text("Next"), #passwordNext')
            page.wait_for_load_state("networkidle")
            print("✅ Login step completed.")
        except Exception as e:
            print(f"❌ Login failed: {e}")
            browser.close()
            return

        # --- Open Looker Studio ---
        print(f"📊 Opening Looker Studio: {LOOKER_URL}")
        page.goto(LOOKER_URL, wait_until="networkidle", timeout=90000)
        
        # รอให้ Component ต่างๆ (Chart/NPS) โหลดเสร็จจริง
        print("⏳ Waiting for charts to load...")
        time.sleep(15) 

        # --- Screenshot ---
        print("📸 Taking screenshot...")
        page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        print(f"✅ Screenshot saved: {SCREENSHOT_PATH}")
        browser.close()

def upload_to_github() -> str:
    print("📤 Uploading image to GitHub repo...")
    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}

    # เช็กไฟล์เดิมเพื่อเอา SHA (จำเป็นสำหรับการ Update)
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json()["sha"]

    payload = {"message": "Update report screenshot (Auto)", "content": content_b64}
    if sha: payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    
    raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FILE_PATH}?t={int(time.time())}"
    print(f"✅ Uploaded! Image URL: {raw_url}")
    return raw_url

def send_to_line(image_url: str):
    print("💬 Sending to LINE...")
    headers = {"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "to": LINE_GROUP_ID,
        "messages": [
            {"type": "image", "originalContentUrl": image_url, "previewImageUrl": image_url},
            {"type": "text", "text": "📊 รายงาน Looker Studio - หน้า 14\n🔄 อัปเดตอัตโนมัติเรียบร้อยครับ"}
        ]
    }
    resp = requests.post("https://api.line.me/v2/bot/message/push", headers=headers, json=payload)
    resp.raise_for_status()
    print(f"✅ Line sent! Status: {resp.status_code}")

if __name__ == "__main__":
    if not all([LOOKER_URL, GOOGLE_EMAIL, GOOGLE_PASSWORD, LINE_TOKEN]):
        print("❌ Error: Missing Environment Variables!")
    else:
        take_screenshot()
        if os.path.exists(SCREENSHOT_PATH):
            img_url = upload_to_github()
            send_to_line(img_url)
