import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright

# ---- ดึงค่าจาก GitHub Secrets (ชื่อ LOOKER_STUDIO_URL ตามที่ตั้งไว้) ----
LOOKER_URL      = os.environ.get("LOOKER_STUDIO_URL")
LINE_TOKEN      = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_GROUP_ID   = os.environ.get("LINE_GROUP_ID")
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN")        
GITHUB_REPO     = os.environ.get("GITHUB_REPOSITORY")  

SCREENSHOT_PATH  = "report_page14.png"
GITHUB_FILE_PATH = "report_page14.png"

def take_screenshot():
    with sync_playwright() as p:
        print("🚀 เริ่มต้นระบบแคปหน้าจอ (Public Mode)...")
        # ใช้ chromium ในการเปิดเว็บ
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            viewport={"width": 1440, "height": 1200}, # ปรับความสูงให้เห็นข้อมูลครบ
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print(f"📊 กำลังเปิดรายงานหน้า 14...")
        # เปิดลิงก์หน้า 14 ที่ระบุไว้ใน Secret
        page.goto(LOOKER_URL, wait_until="networkidle", timeout=90000)
        
        # สำคัญ: รอให้กราฟและค่า NPS โหลดให้เสร็จจริง (เพิ่มเป็น 30 วินาทีเพื่อความชัวร์)
        print("⏳ รอชาร์ตและข้อมูลโหลด 30 วินาที...")
        time.sleep(30) 

        print("📸 กำลังแคปภาพหน้า 14...")
        page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        print(f"✅ บันทึกรูปภาพสำเร็จ: {SCREENSHOT_PATH}")
        browser.close()

def upload_to_github() -> str:
    print("📤 กำลังอัปเดตรูปภาพไปที่ GitHub Repository...")
    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    # ตรวจสอบไฟล์เดิมเพื่อดึง SHA มาใช้ในการ Update ทับไฟล์เก่า
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json()["sha"]

    payload = {
        "message": "Auto-update Report Page 14",
        "content": content_b64
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    
    # สร้าง URL สำหรับส่งให้ LINE (เติม timestamp กัน cache ภาพเก่า)
    raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FILE_PATH}?t={int(time.time())}"
    return raw_url

def send_to_line(image_url: str):
    print("💬 กำลังส่ง Push Message เข้า LINE...")
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }
    
    # ข้อความที่จะส่งพร้อมรูปภาพ
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
                "text": "📊 รายงาน CSI ประจำวัน - หน้า 14\n✅ ข้อมูลอัปเดตล่าสุดเรียบร้อยครับ"
            }
        ]
    }
    resp = requests.post("https://api.line.me/v2/bot/message/push", headers=headers, json=payload)
    resp.raise_for_status()
    print("✅ ส่งเข้า LINE สำเร็จ!")

if __name__ == "__main__":
    if not LOOKER_URL:
        print("❌ Error: ไม่พบค่าใน LOOKER_STUDIO_URL (ตรวจสอบใน GitHub Secrets)")
    else:
        take_screenshot()
        if os.path.exists(SCREENSHOT_PATH):
            img_url = upload_to_github()
            send_to_line(img_url)
