import requests
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

USERNAME   = os.environ["CSI_USERNAME"]
PASSWORD   = os.environ["CSI_PASSWORD"]
LINE_TOKEN = os.environ["LINE_TOKEN"]
LINE_GROUP = os.environ["LINE_GROUP_ID"]

TARGET_URL = "https://csi-bdms-mgrs.azurewebsites.net/dashboard/BHP?sitecode=BHP"

def scrape_dashboard():

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 30)

        # ---------- LOGIN ----------
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()

        # ---------- รอ redirect กลับ dashboard ----------
        wait.until(EC.url_contains("/dashboard/BHP"))

        # ---------- รอ table โหลด ----------
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")

        data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 2:
                data.append({
                    "form": cols[0].text.strip(),
                    "total": cols[1].text.strip()
                })

        return data

    except Exception as e:
        print("❌ Error:", e)
        print("Current URL:", driver.current_url)
        return []

    finally:
        driver.quit()


def send_line(message):

    url = "https://api.line.me/v2/bot/message/push"

    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "to": LINE_GROUP,
        "messages": [
            {"type": "text", "text": message}
        ]
    }

    r = requests.post(url, headers=headers, json=payload)

    if r.status_code != 200:
        print("❌ LINE Error:", r.text)
    else:
        print("✅ ส่ง LINE สำเร็จ")


def format_message(data):

    today = datetime.now().strftime("%d/%b/%Y")

    lines = [
        "📊 CSI Dashboard BHP",
        f"📅 {today}",
        "──────────────────"
    ]

    for item in data:
        lines.append(f"📋 {item['form']} → {item['total']} รายการ")

    lines.append("──────────────────")
    lines.append("🏥 Bangkok Hospital Pakchong")

    return "\n".join(lines)


if __name__ == "__main__":

    data = scrape_dashboard()

    if not data:
        print("⚠️ ไม่พบข้อมูล")
    else:
        message = format_message(data)
        send_line(message)
