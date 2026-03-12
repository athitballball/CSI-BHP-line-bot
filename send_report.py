import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright

# ---- ดึงค่าจาก GitHub Secrets ----
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
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            viewport={"width": 1440, "height": 1200}, # เพิ่มความสูงเผื่อกราฟยาว
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print(f"📊 กำลังเปิดรายงานหน้า 14...")
        # เปิดลิงก์ที่ตั้งค่าไว้ใน Secret (ควรเป็นลิงก์ที่ระบุหน้า 14 มาแล้ว)
        page.goto(LOOKER_URL, wait_until="networkidle", timeout=90000)
        
        # รอให้ข้อมูล NPS และชาร์ตต่างๆ โหลดเสร็จจริง
        print("⏳ รอชาร์ตโหลด 25 วินาที...")
        time.sleep(25) 

        print("📸 กำลังแคปภาพหน้า 14...")
        page.screenshot(path=SCREENSHOT_PATH, full_page=False)
        print(f"✅ บันทึกรูปภาพสำเร็จ")
        browser.close()

def upload_to_github() -> str:
    print("📤 กำลังอัปเดตรูปภาพไปที่ GitHub...")
    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}

    # ดึงค่า SHA เดิมเพื่อทำการอัปเดตไฟล์ทับ
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json()["sha"]

    payload = {"message": "Update Page 14 Screenshot", "content": content_b64}
    if sha: payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    
    # ส่ง URL แบบ Raw ไปที่ LINE (เติม t= เพื่อป้องกันภาพแคชเก่า)
    raw_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FILE_PATH}?t={int(time.time())}"
    return raw_url

def send_to_line(image_url: str):
    print("💬 กำลังส่งเข้ากลุ่ม LINE...")
    headers = {"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"}
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
    print("✅ ส่ง LINE สำเร็จ!")

if __name__ == "__main__":
    if not LOOKER_URL:
        print("❌ Error: ไม่พบลิงก์ใน LOOKER_STUDIO_URL (Secrets)")
    else:
        take_screenshot()
        if os.path.exists(SCREENSHOT_PATH):
            img_url = upload_to_github()
            send_to_line(img_url)
