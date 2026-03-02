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
        wait = WebDriverWait(driver, 30)

        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()
        wait.until(EC.url_contains("FirstPage"))
        print("✅ Login สำเร็จ:", driver.current_url)

        driver.get(f"{LOGIN_URL}/Home/viewscore/BHP")
        print("✅ เข้าหน้า viewscore แล้ว")
        time.sleep(2)
      
        dashboard_btn = wait.until(EC.element_to_be_clickable(
        (By.ID, "btnDashboard")
        ))
        dashboard_btn.click()
        print("✅ กด Dashboard แล้ว")
        time.sleep(3)

        btn = driver.find_element(By.ID, "btnDashboard")
        driver.execute_script("arguments[0].click();", btn)
        print("🖱️ คลิก Dashboard แล้ว — รอข้อมูลโหลด...")

        def table_has_data(d):
            rows = d.find_elements(By.CSS_SELECTOR, "#resultTable tbody tr")
            for r in rows:
                cells = r.find_elements(By.TAG_NAME, "td")
                if cells and cells[0].get_attribute("innerHTML").strip():
                    text = cells[0].get_attribute("innerText").strip()
                    if text and "No data" not in text:
                        return True
            return False

        try:
            wait.until(table_has_data)
            print("✅ ตารางโหลดข้อมูลแล้ว")
        except Exception:
            print("⚠️ timeout รอตาราง — ลองดึงข้อมูลที่มี")

        time.sleep(1)

        rows = driver.find_elements(By.CSS_SELECTOR, "#resultTable tbody tr")
        print(f"📋 จำนวน rows ทั้งหมด: {len(rows)}")

        today = datetime.now().strftime("%d/%b/%Y")
        data = []
        for i, row in enumerate(rows):
            cells = row.find_elements(By.TAG_NAME, "td")
            # ใช้ innerText แทน .text เพื่อให้ได้ค่าที่ render แล้ว
            texts = [
                driver.execute_script("return arguments[0].innerText;", c).strip()
                for c in cells
            ]
            print(f"  row[{i}]: {texts}")
            if len(texts) >= 3 and texts[1] and "No data" not in texts[1]:
                data.append({"form": texts[1], "total": texts[2]})

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
        f"🏥 Site: {SITE_CODE}",
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
