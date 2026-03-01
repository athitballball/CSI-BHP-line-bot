import requests
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time

USERNAME      = os.environ["CSI_USERNAME"]
PASSWORD      = os.environ["CSI_PASSWORD"]
LINE_TOKEN    = os.environ["LINE_TOKEN"]
LINE_GROUP_ID = os.environ["LINE_GROUP_ID"]
LOGIN_URL     = "https://csi-bdms-mgrs.azurewebsites.net"
SITE_CODE     = os.environ.get("SITE_CODE", "BHP")

def scrape_csi():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )
    try:
        wait = WebDriverWait(driver, 20)

        # Login
        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()
        wait.until(EC.url_contains("FirstPage"))
        print("✅ Login สำเร็จ")

        driver.get(f"{LOGIN_URL}/Home/FirstPage/87?userId=87&type=A&mail=***&site=1003&login_type=Y&user_part=N#")
        print("✅ เข้าหน้า BHP แล้ว")
        time.sleep(2)

        driver.get(f"{LOGIN_URL}/dashboard/BHP?sitecode=BHP")
        print("✅ คลิก Dashboard แล้ว")
        time.sleep(3)

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        today = datetime.now().strftime("%d/%b/%Y")
        data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 1:
                text0 = cols[0].text.strip()
                text1 = cols[1].text.strip() if len(cols) > 1 else "0"
                if text0:
                    data.append({"form": text0, "total": text1})
        print(f"📊 พบข้อมูล {len(data)} รายการ")
        return today, data

    finally:
        driver.quit()

def send_line(message):
    requests.post(
        "https://api.line.me/v2/bot/message/push",
        headers={
            "Authorization": f"Bearer {LINE_TOKEN}",
            "Content-Type": "application/json"
        },
        json={
            "to": LINE_GROUP_ID,
            "messages": [{"type": "text", "text": message}]
        }
    )

def format_message(date, data):
    lines = [
        f"\n📊 CSI Daily Report",
        f"📅 {date}",
        f"🏥 Site: BHP",
        "─" * 25,
    ]
    for item in data:
        lines.append(f"📋 {item['form']}  →  {item['total']} รายการ")
    lines.append("─" * 25)
    lines.append("🏥 Bangkok Hospital Pakchong")
    return "\n".join(lines)

today, data = scrape_csi()
if not data:
    print("⚠️ ไม่พบข้อมูลวันนี้")
else:
    message = format_message(today, data)
    print(message)
    send_line(message)
    print("✅ ส่งสำเร็จ")
