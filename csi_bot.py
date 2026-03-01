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

USERNAME   = os.environ["CSI_USERNAME"]
PASSWORD   = os.environ["CSI_PASSWORD"]
LINE_TOKEN = os.environ["LINE_TOKEN"]
LINE_GROUP_ID = os.environ["LINE_GROUP_ID"]

LOGIN_URL  = "https://csi-bdms-mgrs.azurewebsites.net"


SITE_CODE  = os.environ.get("SITE_CODE", "BHP")


def scrape_csi():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    chrome_path = os.environ.get("CHROME_PATH")
    if chrome_path:
        options.binary_location = chrome_path

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    try:
        wait = WebDriverWait(driver, 20)

        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()

        wait.until(EC.url_contains("FirstPage"))
        print("✅ Login สำเร็จ:", driver.current_url)

        viewscore_url = f"{LOGIN_URL}/Home/viewscore/{SITE_CODE}"
        driver.get(viewscore_url)
        print("📄 เข้าหน้า viewscore:", driver.current_url)

        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        dashboard_btn = None
        selectors = [
            (By.ID, "btnDashboard"),
            (By.XPATH, "//button[contains(translate(text(),'dashboard','DASHBOARD'),'DASHBOARD')]"),
            (By.XPATH, "//a[contains(translate(text(),'dashboard','DASHBOARD'),'DASHBOARD')]"),
            (By.XPATH, "//button[contains(@onclick,'dashboard') or contains(@id,'dashboard') or contains(@class,'dashboard')]"),
            (By.XPATH, "//input[@type='button' and contains(translate(@value,'dashboard','DASHBOARD'),'DASHBOARD')]"),
        ]

        for by, selector in selectors:
            try:
                dashboard_btn = wait.until(EC.element_to_be_clickable((by, selector)))
                print(f"🔘 พบปุ่ม Dashboard ด้วย selector: {selector}")
                break
            except Exception:
                continue

        if dashboard_btn is None:
            buttons = driver.find_elements(By.XPATH, "//button | //a[contains(@class,'btn')] | //input[@type='button']")
            print("⚠️ ไม่พบปุ่ม Dashboard — ปุ่มที่มีในหน้า:")
            for b in buttons:
                print(f"  tag={b.tag_name}, id={b.get_attribute('id')}, text={b.text[:50]!r}")
            raise RuntimeError("ไม่พบปุ่ม Dashboard")

        dashboard_btn.click()
        print("🖱️ คลิกปุ่ม Dashboard แล้ว")

        wait.until(
            lambda d: len([
                r for r in d.find_elements(By.CSS_SELECTOR, "table tbody tr")
                if "dataTables_empty" not in r.get_attribute("innerHTML")
            ]) > 0
        )
        time.sleep(1)  # buffer เล็กน้อย

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        print(f"📊 จำนวน rows: {len(rows)}")

        today = datetime.now().strftime("%d/%b/%Y")
        data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 2:
                text0 = cols[0].text.strip()
                text1 = cols[1].text.strip()
                if text0 and "dataTables_empty" not in row.get_attribute("class"):
                    data.append({"form": text0, "total": text1})

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


# ── Main ──────────────────────────────────────────────────────────
today, data = scrape_csi()

if not data:
    print("⚠️ ไม่พบข้อมูลวันนี้")
else:
    message = format_message(today, data)
    print(message)
    send_line(message)
    print("✅ ส่งสำเร็จ")
