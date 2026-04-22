# =========================
# APPLY LOGIN + DRIVER + ELEMENT INTERACTION FROM SOURCE 1
# INTO SOURCE 2
# =========================

import os
import gc
import json
import time
import pytz
import gspread
import tempfile
import undetected_chromedriver as uc

from datetime import datetime, timezone
from oauth2client.service_account import ServiceAccountCredentials

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from gspread_formatting import *

gc.disable()

# =========================
# CONFIG
# =========================
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM/edit?usp=sharing"

email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")
gcp_credentials_json = os.environ.get("GCP_SA_KEY")


# =========================
# GOOGLE SHEETS
# =========================
def get_gsclient():
    creds_dict = json.loads(gcp_credentials_json)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scopes)
    return gspread.authorize(creds)


# =========================
# SCREENSHOT
# =========================
def save_screenshot(driver, file_name="error.png"):
    try:
        driver.save_screenshot(file_name)
        print(f"📸 Saved: {file_name}")
    except:
        pass


# =========================
# NEW DRIVER FROM SOURCE 1
# =========================
def get_driver():
    options = uc.ChromeOptions()

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    options.page_load_strategy = "eager"
    options.add_argument("--lang=en-GB")

    proxy_url = os.getenv("PROXY_URL")
    if proxy_url:
        options.add_argument(f"--proxy-server={proxy_url}")

    driver = uc.Chrome(options=options, version_main=146)

    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });

                window.navigator.chrome = { runtime: {} };

                Object.defineProperty(navigator, 'plugins', {
                    get: () => [1,2,3,4,5]
                });

                Object.defineProperty(navigator, 'languages', {
                    get: () => ['en-GB','en-US','en']
                });
            """
        },
    )

    return driver


# =========================
# LOGIN FROM SOURCE 1
# =========================
def login():
    driver = get_driver()

    # Truy cập link chuẩn cho Work/School
    driver.get("https://teams.microsoft.com/")
    wait = WebDriverWait(driver, 30)

    try:
        print("⏳ Logging in...")

        # 1. Xử lý nút Sign in (nếu bị đẩy ra trang chờ)
        try:
            sign_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[contains(., "Sign in")] | //a[contains(., "Sign in")] | //button[contains(., "Đăng nhập")]')
                )
            )
            sign_btn.click()
        except:
            pass # Bỏ qua nếu form điền email hiện ra trực tiếp

        # 2. Ô nhập Email (Sử dụng Selector linh hoạt cho Microsoft)
        email_box = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="email"], input[name="loginfmt"]'))
        )
        email_box.send_keys(email)
        email_box.send_keys(Keys.RETURN)

        time.sleep(3)
        # ====== THÊM ĐOẠN NÀY VÀO ======
        # Xử lý trường hợp Microsoft đòi gửi mã code, ép nó quay về dùng Mật khẩu
        try:
            use_pass_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]')
                )
            )
            use_pass_btn.click()
            time.sleep(2)
        except:
            pass # Nếu màn hình đi thẳng tới ô mật khẩu thì cứ bỏ qua bước này
        # ===============================
        # 3. Ô nhập Password
        pass_box = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="password"], input[name="passwd"]'))
        )
        pass_box.send_keys(password)
        pass_box.send_keys(Keys.RETURN)

       # 4. Xử lý nút "Stay signed in?" (Chọn No để không lưu đăng nhập)
        try:
            print("⏳ Đang xử lý màn hình Stay signed in...")
            no_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="declineButton"] | //*[@id="idBtn_Back"] | //*[@value="No"] | //button[contains(., "No")]')
                )
            )
            no_btn.click()
            time.sleep(3)
        except:
            print("⚠️ Không thấy màn hình Stay signed in, tiếp tục...")
            pass

        print("✅ Login success")
        
        # Chờ giao diện Teams load hẳn
        time.sleep(15) 

        return driver

    except Exception as e:
        save_screenshot(driver, "login_error.png")
        print("❌ Login failed:", e)
        driver.quit()
        return None
# =========================
# CREATE SHEET
# =========================
def create_worksheet(title):
    gcx = get_gsclient()
    sheet = gcx.open_by_url(SPREADSHEET_URL)

    names = [x.title for x in sheet.worksheets()]

    if title in names:
        return

    ws = sheet.add_worksheet(title=title, rows=1000, cols=4)

    ws.update("A1:D1", [["NAME", "DATE", "TIME", "CONTENT"]])

    set_column_widths(
        ws,
        [
            ("A", 180),
            ("B", 100),
            ("C", 100),
            ("D", 1000),
        ],
    )

    ws.freeze(rows=1)
    # THÊM 2 DÒNG NÀY VÀO DƯỚI CÙNG:
    fmt = cellFormat(wrapStrategy='WRAP')
    format_cell_range(ws, 'D:D', fmt)

# =========================
# SAVE DATA
# =========================
def save_to_excel(rows, worksheet):
    gcx = get_gsclient()
    sheet = gcx.open_by_url(SPREADSHEET_URL)
    ws = sheet.worksheet(worksheet)

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        print(f"✅ Added {len(rows)} rows -> {worksheet}")


# =========================
# GET MESSAGE
# =========================
def get_messages(driver, worksheet):
    try:
        wait = WebDriverWait(driver, 20)

        pane = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-tid="message-pane-list-runway"]')
            )
        )

        items = pane.find_elements(
            By.CSS_SELECTOR,
            '[data-tid="chat-pane-item"]'
        )

        data = []

        for item in items:
            try:
                name = item.find_element(
                    By.CSS_SELECTOR,
                    '[data-tid="message-author-name"]'
                ).text

                timestamp = item.find_element(
                    By.TAG_NAME,
                    "time"
                ).get_attribute("datetime")

                dt_utc = datetime.strptime(
                    timestamp,
                    "%Y-%m-%dT%H:%M:%S.%fZ"
                ).replace(tzinfo=timezone.utc)

                dt_local = dt_utc.astimezone(local_tz)

                date_str = dt_local.strftime("%Y-%m-%d")
                time_str = dt_local.strftime("%H:%M:%S")

                content = item.find_element(
                    By.CSS_SELECTOR,
                    '[id^="content-"]'
                ).text.strip()
                data.append(
                    [name, date_str, time_str, content]
                )

            except:
                continue

        save_to_excel(data, worksheet)

    except Exception as e:
        print("❌ get_messages error:", e)


# =========================
# SEARCH CHAT FROM SOURCE 1
# =========================
def open_chat_by_search(driver, chat_name):
    wait = WebDriverWait(driver, 20)

    try:
        search_xpath = (
            '//input[@placeholder="Search"]'
            ' | //input[@aria-label="Search"]'
            ' | //input[@id="ms-searchux-input"]'
        )

        search = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, search_xpath)
            )
        )

        search.click()
        search.send_keys(Keys.CONTROL + "a")
        search.send_keys(Keys.BACKSPACE)
        search.send_keys(chat_name)

        time.sleep(3)

        ActionChains(driver)\
            .send_keys(Keys.ARROW_DOWN)\
            .pause(1)\
            .send_keys(Keys.ENTER)\
            .perform()

        time.sleep(5)

        print(f"📂 Opened: {chat_name}")
        return True

    except Exception as e:
        print("❌ Cannot open:", chat_name, e)
        return False


# =========================
# GET ALL GROUPS
# =========================
def get_all_groups(driver):
    wait = WebDriverWait(driver, 20)

    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-tid="chat-list-item"]')
            )
        )

        groups = driver.find_elements(
            By.CSS_SELECTOR,
            '[data-tid="chat-list-item"]'
        )

        names = []

        for g in groups:
            try:
                txt = g.text.strip().split("\n")[0]
                if txt and txt not in names:
                    names.append(txt)
            except:
                pass

        print(f"Found {len(names)} groups")
        return names

    except Exception as e:
        save_screenshot(driver, "error_groups.png")
        print("❌ get_all_groups:", e)
        return []


# =========================
# MAIN
# =========================
if __name__ == "__main__":
    driver = login()

    if driver:

        group_names = get_all_groups(driver)

        for group in group_names:
            try:
                print(f"\n===== {group} =====")

                create_worksheet(group)

                if open_chat_by_search(driver, group):
                    get_messages(driver, group)

                time.sleep(3)

            except Exception as e:
                print("Skip:", group, e)

        driver.quit()
        print("✅ DONE")
