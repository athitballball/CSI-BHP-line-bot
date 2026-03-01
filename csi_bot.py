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

        # ── 1) Login ──────────────────────────────────────────────
        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.NAME, "username")))
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login").click()
        wait.until(EC.url_contains("FirstPage"))
        print("✅ Login สำเร็จ:", driver.current_url)

        # ── 2) ไปหน้า viewscore ───────────────────────────────────
        driver.get(f"{LOGIN_URL}/Home/viewscore/{SITE_CODE}")
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(3)
        print("📄 อยู่ที่:", driver.current_url)

        # ── 3) DEBUG: print ปุ่มทั้งหมดในหน้า ────────────────────
        all_btns = driver.find_elements(
            By.XPATH,
            "//button | //a[contains(@class,'btn')] | //input[@type='button'] | //input[@type='submit']"
        )
        print(f"🔍 พบปุ่ม {len(all_btns)} ปุ่ม:")
        for b in all_btns:
            print(f"  tag={b.tag_name} | id={b.get_attribute('id')!r} | "
                  f"class={b.get_attribute('class')!r} | text={b.text.strip()[:40]!r}")

        # ── 4) คลิกปุ่ม Dashboard ─────────────────────────────────
        dashboard_btn = None
        selectors = [
            (By.XPATH, "//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'DASHBOARD') and (self::button or self::a or self::input)]"),
            (By.XPATH, "//*[contains(@id,'ashboard') or contains(@id,'ASHBOARD')]"),
            (By.XPATH, "//*[contains(@onclick,'ashboard')]"),
            (By.XPATH, "//*[contains(@class,'ashboard')]"),
        ]

        for by, sel in selectors:
            try:
                elems = driver.find_elements(by, sel)
                if elems:
                    dashboard_btn = elems[0]
                    print(f"🔘 พบ Dashboard ด้วย: {sel}")
                    break
            except Exception:
                continue

        if dashboard_btn:
            driver.execute_script("arguments[0].click();", dashboard_btn)
            print("🖱️ คลิก Dashboard แล้ว")
            time.sleep(4)
        else:
            print("⚠️ ไม่พบปุ่ม Dashboard — ลองดึงข้อมูลจากตารางที่มีอยู่เลย")

        # ── 5) print HTML ส่วนแรกเพื่อ debug ─────────────────────
        body_html = driver.execute_script("return document.body.innerHTML;")
        print("📝 HTML (500 chars):", body_html[:500])

        # ── 6) ดึงข้อมูลจากตาราง ──────────────────────────────────
        table_selectors = [
            "#resultTable tbody tr",
            "table#dataTable tbody tr",
            "table tbody tr",
        ]

        rows = []
        for sel in table_selectors:
            rows = [
                r for r in driver.find_elements(By.CSS_SELECTOR, sel)
                if "dataTables_empty" not in r.get_attribute("innerHTML")
            ]
            if rows:
                print(f"✅ ใช้ selector: {sel} | พบ {len(rows)} rows")
                break

        today = datetime.now().strftime("%d/%b/%Y")
        data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 2:
                text0 = cols[0].text.strip()
                text1 = cols[1].text.strip()
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
