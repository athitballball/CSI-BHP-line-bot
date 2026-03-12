import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright

# ---- ดึงค่าจาก GitHub Secrets ----
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
        print("🚀 กำลังเริ่ม Browser...")
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            viewport={"width": 1440, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # --- Login Google ---
        print("🔐 กำลัง Login Google...")
        try:
            page.goto("https://accounts.google.com/signin", timeout=60000)
            page.fill('input[type="email"]', GOOGLE_EMAIL)
            page.click('#identifierNext')
            page.wait_for_timeout(3000)
            page.fill('input[type="password"]', GOOGLE_PASSWORD)
            page.click('#passwordNext')
            page.wait_for_load_state("networkidle")
            print("✅ Login สำเร็จ (เบื้องต้น)")
        except Exception as e:
            print(f"❌ Login พลาด: {e}")
            browser.close()
            return

        # --- เข้า Looker Studio ---
        print(f"📊 กำลังเปิด Looker Studio...")
        page.goto(LOOKER_URL, wait_until="networkidle", timeout=90000)
        
        # รอให้กราฟโหลด (ปรับเวลาได้ตามความอืดของ Report)
        print("⏳ รอชาร์ตโหลด 15 วินาที...")
        time.sleep(15) 

        # --- แคปรูป ---
        print("📸 กำลังแคปหน้าจอ...")
        page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        print(f"✅ บันทึกรูปภาพแล้ว: {SCREENSHOT_PATH}")
        browser.close()

def upload_to_github() -> str:
    print("📤 กำลังอัปโหลดรูปไปที่ GitHub...")
    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}

    # เช็กไฟล์เดิมเพื่อดึง SHA มาอัปเดตทับ
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json()["sha"]

    payload = {"message": "Auto-update report screenshot", "content": content_b64}
    if sha: payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    
    # สร้าง URL รูปภาพแบบ Raw (เติม timestamp เพื่อไม่ให้ LINE จำภาพเก่า)
    raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FILE_PATH}?t={int(time.time())}"
    return raw_url

def send_to_line(image_url: str):
    print("💬 กำลังส่งเข้า LINE...")
    headers = {"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "to": LINE_GROUP_ID,
        "messages": [
            {"type": "image", "originalContentUrl": image_url, "previewImageUrl": image_url},
            {"type": "text", "text": "📊 รายงาน Looker Studio - หน้า 14\n🔄 อัปเดตข้อมูลล่าสุดเรียบร้อยครับ"}
        ]
    }
    resp = requests.post("https://api.line.me/v2/bot/message/push", headers=headers, json=payload)
    resp.raise_for_status()
    print(f"✅ ส่ง LINE เรียบร้อย! (Status: {resp.status_code})")

if __name__ == "__main__":
    # ตรวจสอบว่าตั้งค่า Secret ครบไหม
    if not all([LOOKER_URL, GOOGLE_EMAIL, GOOGLE_PASSWORD]):
        print("❌ Error: ตั้งค่า Secrets ไม่ครบ!")
    else:
        take_screenshot()
        if os.path.exists(SCREENSHOT_PATH):
            img_url = upload_to_github()
            send_to_line(img_url)
