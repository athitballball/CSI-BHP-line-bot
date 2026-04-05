import os
import time
import base64
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ---- ดึงค่าจาก GitHub Secrets ----
LOOKER_URL      = os.environ.get("LOOKER_STUDIO_URL", "").strip()
LINE_TOKEN      = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
LINE_GROUP_ID   = os.environ.get("LINE_GROUP_ID", "").strip()
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN", "").strip()
GITHUB_REPO     = os.environ.get("GITHUB_REPOSITORY", "").strip()
GITHUB_BRANCH   = os.environ.get("GITHUB_BRANCH", "main").strip()  # ✅ ปรับ branch ได้

SCREENSHOT_PATH  = "report_page14.png"
GITHUB_FILE_PATH = "report_page14.png"


def check_env():
    """ตรวจสอบ Environment Variables ครบทุกตัวก่อนรัน"""
    missing = []
    for name, val in {
        "LOOKER_STUDIO_URL": LOOKER_URL,
        "LINE_CHANNEL_ACCESS_TOKEN": LINE_TOKEN,
        "LINE_GROUP_ID": LINE_GROUP_ID,
        "GITHUB_TOKEN": GITHUB_TOKEN,
        "GITHUB_REPOSITORY": GITHUB_REPO,
    }.items():
        if not val:
            missing.append(name)
    
    if missing:
        print(f"❌ Error: ไม่พบค่าใน Secrets: {', '.join(missing)}")
        return False
    return True


def take_screenshot() -> bool:
    """เปิด Looker Studio และแคปหน้าจอ"""
    with sync_playwright() as p:
        print("🚀 เริ่มต้นระบบแคปหน้าจอ...")
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

        try:
            print(f"📊 กำลังเปิด Looker Studio หน้า 14...")
            page.goto(LOOKER_URL, wait_until="domcontentloaded", timeout=90000)

            # ✅ รอให้ <iframe> หรือ canvas ของ Looker โหลดก่อน
            print("⏳ รอให้ชาร์ตโหลด...")
            try:
                page.wait_for_selector("canvas, iframe, [data-testid]", timeout=30000)
                print("✅ พบ element หน้าจอแล้ว")
            except PlaywrightTimeout:
                print("⚠️ ไม่พบ selector หลัก — รอเพิ่มเติม 15 วิ")

            # รอเพิ่มอีกนิดให้ข้อมูลใน chart render เสร็จ
            time.sleep(15)

            print("📸 กำลังแคปภาพ...")
            page.screenshot(path=SCREENSHOT_PATH, full_page=False)
            print(f"✅ บันทึกรูปภาพสำเร็จ: {SCREENSHOT_PATH}")
            return True

        except PlaywrightTimeout as e:
            print(f"❌ Timeout เปิดหน้าเว็บ: {e}")
            return False
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาด: {e}")
            return False
        finally:
            browser.close()


def upload_to_github() -> str:
    """อัปโหลดรูปไป GitHub และคืน Raw URL"""
    print("📤 กำลังอัปเดตรูปภาพไปที่ GitHub...")
    with open(SCREENSHOT_PATH, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json"
    }

    # ดึง SHA ของไฟล์เก่า (ถ้ามี) เพื่อ overwrite
    sha = None
    resp = requests.get(api_url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json().get("sha")

    payload = {
        "message": f"Auto-update Report Page 14 [{time.strftime('%Y-%m-%d %H:%M')}]",
        "content": content_b64,
        "branch": GITHUB_BRANCH  # ✅ ใช้ branch จาก env
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(api_url, headers=headers, json=payload)
    resp.raise_for_status()
    print("✅ อัปโหลด GitHub สำเร็จ!")

    # ✅ เติม timestamp กัน LINE cache รูปเก่า
    raw_url = (
        f"https://raw.githubusercontent.com/{GITHUB_REPO}"
        f"/{GITHUB_BRANCH}/{GITHUB_FILE_PATH}?t={int(time.time())}"
    )
    return raw_url


def wait_for_image_ready(url: str, retries: int = 5, delay: int = 5):
    """
    รอให้ GitHub propagate ไฟล์ก่อนส่ง LINE
    (raw.githubusercontent.com อาจใช้เวลาไม่กี่วิหลัง push)
    """
    print("⏳ รอให้ GitHub พร้อมส่งรูป...")
    for i in range(retries):
        resp = requests.head(url)
        if resp.status_code == 200:
            print("✅ รูปพร้อมแล้ว!")
            return True
        print(f"  retry {i+1}/{retries} (status {resp.status_code})")
        time.sleep(delay)
    print("⚠️ GitHub ยังไม่พร้อม — ส่ง LINE ต่อไปเลย")
    return False


def send_to_line(image_url: str):
    """ส่งรูปและข้อความเข้ากลุ่ม LINE"""
    print("💬 กำลังส่ง Push Message เข้า LINE...")
    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }

    # ✅ ต้องใช้ URL ที่ไม่มี query string สำหรับ originalContentUrl / previewImageUrl
    # LINE ไม่รองรับ query string ใน image URL → ตัด ?t= ออก แต่ใส่ใน text แทน
    clean_url = image_url.split("?")[0]
    timestamp_str = time.strftime("%d/%m/%Y %H:%M")

    payload = {
        "to": LINE_GROUP_ID,
        "messages": [
            {
                "type": "image",
                "originalContentUrl": clean_url,
                "previewImageUrl": clean_url
            }
        ]
    }

    resp = requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers=headers,
        json=payload
    )

    print(f"📨 LINE Status: {resp.status_code}")
    print(f"📨 LINE Response: {resp.text}")
    if resp.status_code == 200:

        print("✅ ส่งเข้า LINE สำเร็จ!")
    else:
        print(f"❌ LINE Error {resp.status_code}: {resp.text}")
        resp.raise_for_status()


# ─────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)

    # 1) ตรวจสอบ env
    if not check_env():
        exit(1)

    # 2) แคปหน้าจอ
    success = take_screenshot()
    if not success or not os.path.exists(SCREENSHOT_PATH):
        print("❌ ไม่พบไฟล์รูปภาพ — หยุดการทำงาน")
        exit(1)

    # 3) อัปโหลด GitHub
    img_url = upload_to_github()

    # 4) รอให้ GitHub พร้อม
    wait_for_image_ready(img_url.split("?")[0])

    # 5) ส่ง LINE
    send_to_line(img_url)

    print("=" * 50)
    print("🎉 เสร็จสิ้นทุกขั้นตอน!")
