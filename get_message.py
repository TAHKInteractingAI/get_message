import gspread
from oauth2client.service_account import ServiceAccountCredentials
import gc

gc.disable()
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import pytz
from datetime import datetime, timezone
from gspread_formatting import *
import os
import json

local_tz = pytz.timezone("Asia/Ho_Chi_Minh")
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM/edit?usp=sharing"
email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")
gcp_credentials_json = os.environ.get("GCP_SA_KEY")


def get_gsclient():
    """Khởi tạo client Google Sheets sử dụng Service Account"""
    creds_dict = json.loads(gcp_credentials_json)
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPES)
    return gspread.authorize(creds)


def save_screenshot(driver: webdriver.Chrome, file_name: str = "screenshot.png"):
    """Chụp màn hình và lưu lại file"""
    try:
        driver.save_screenshot(file_name)
        print(f"📸 Đã chụp màn hình và lưu tại: {file_name}")
    except Exception as e:
        print(f"❌ Không thể chụp màn hình: {e}")


def login():
    import tempfile

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # Sử dụng mode headless mới ổn định hơn
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    temp_dir = tempfile.mkdtemp()
    options.add_argument(f"--user-data-dir={temp_dir}")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
        },
    )

    # Kiểm tra xem có iframe nào không và switch thử (nếu cần)
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for index, iframe in enumerate(iframes):
        print(f"Thử kiểm tra iframe index: {index}")
        driver.switch_to.frame(index)
        try:
            btn = driver.find_element(By.PARTIAL_LINK_TEXT, "Sign in")
            btn.click()
            break
        except:
            driver.switch_to.default_content()

    driver.get("https://teams.live.com/v2/")
    time.sleep(10)
    save_screenshot(driver, "check_initial_page.png")

    try:
        sign_in_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//*[contains(translate(text(), 'SIGNIN', 'signin'), 'sign in')]",
                )
            )
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", sign_in_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", sign_in_btn)
        print("✅ Đã nhấn nút Sign In bằng JS")
    except Exception as e:
        print(f"❌ Không tìm thấy nút bằng Xpath, thử chuyển hướng trực tiếp...")
        driver.get("https://teams.live.com/v2/?login_hint")

    email_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "usernameEntry"))
    )
    email_input.send_keys(email)
    email_input.send_keys(Keys.RETURN)
    time.sleep(8)

    # Chọn 'Use your password' nếu xuất hiện
    try:
        use_pass_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//span[@role="button" and contains(text(), "Use your password")]',
                )
            )
        )
        use_pass_btn.click()
        time.sleep(3)
    except Exception as e:
        print("Không tìm thấy nút 'Use your password'.")

    # Tiếp tục nhập mật khẩu như cũ
    password_input = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "passwordEntry"))
    )
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    time.sleep(8)

    # driver.save_screenshot("after_email.png")
    # print("Đã chụp màn hình sau khi nhập email.")

    try:
        no_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button[data-testid="secondaryButton"]')
            )
        )
        no_button.click()
        time.sleep(5)
    except Exception as e:
        print("Không tìm thấy nút 'No'.")

    time.sleep(20)

    try:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button[data-testid="primaryButton"]')
            )
        )
        button.click()
        time.sleep(10)
    except:
        pass

    return driver


def get_last_saved_datetime(work_sheet):
    existing_data = work_sheet.get_all_values()
    if len(existing_data) <= 1:
        return None
    last_row = existing_data[-1]
    last_date_str, last_time_str = last_row[1], last_row[2]
    try:
        return datetime.strptime(
            f"{last_date_str} {last_time_str}", "%Y-%m-%d %H:%M:%S"
        )
    except ValueError:
        return None


def save_to_excel(new_data, worksheet_title):
    try:
        gc = get_gsclient()
        sheet = gc.open_by_url(SPREADSHEET_URL)
        work_sheet = sheet.worksheet(worksheet_title)

        last_saved_datetime = get_last_saved_datetime(work_sheet)
        new_messages = []
        for msg in new_data:
            msg_datetime = datetime.strptime(
                f"{msg['DATE']} {msg['TIME']}", "%Y-%m-%d %H:%M:%S"
            )
            if last_saved_datetime is None or msg_datetime > last_saved_datetime:
                new_messages.append(
                    [msg["NAME"], msg["DATE"], msg["TIME"], msg["CONTENT"]]
                )

        if new_messages:
            work_sheet.append_rows(new_messages, value_input_option="USER_ENTERED")
            print(
                f"✅ Đã thêm {len(new_messages)} tin nhắn mới vào worksheet '{worksheet_title}'."
            )
        else:
            print(f"ℹ️ Không có tin nhắn mới trong worksheet '{worksheet_title}'.")
    except Exception as e:
        print(f"❌ Lỗi khi cập nhật Google Sheet: {str(e)}")


def get_messege(driver, worksheet):
    try:
        wait = WebDriverWait(driver, 15)
        chat_list = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-tid="message-pane-list-runway"]')
            )
        )
        message_items = chat_list.find_elements(
            By.CSS_SELECTOR, '[data-tid="chat-pane-item"]'
        )

        data = []
        for item in message_items:
            try:
                name = item.find_element(
                    By.CSS_SELECTOR, '[data-tid="message-author-name"]'
                ).text
                timestamp = item.find_element(By.TAG_NAME, "time").get_attribute(
                    "datetime"
                )
                dt_utc = datetime.strptime(timestamp, "%Y-%m-%dT%H:%M:%S.%fZ").replace(
                    tzinfo=timezone.utc
                )
                dt_local = dt_utc.astimezone(local_tz)
                date_str = dt_local.strftime("%Y-%m-%d")
                time_str = dt_local.strftime("%H:%M:%S")
                content = item.find_element(
                    By.CSS_SELECTOR, '[id^="content-"][aria-label]'
                ).get_attribute("aria-label")
                content = content.replace("\xa0", " ").strip()

                data.append(
                    {
                        "DATE": date_str,
                        "TIME": time_str,
                        "NAME": name,
                        "CONTENT": content,
                    }
                )
            except Exception:
                continue  # Bỏ qua nếu có lỗi với một tin nhắn cụ thể

        if data:
            save_to_excel(data, worksheet)
        else:
            print("Không tìm thấy tin nhắn nào.")

    except Exception as e:
        print(f"Lỗi khi lấy tin nhắn: {e}")


def create_worksheet(title):
    gc = get_gsclient()

    # Mở Google Sheet bằng URL
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM/edit?usp=sharing"
    sheet = gc.open_by_url(spreadsheet_url)
    sheet_names = [s.title for s in sheet.worksheets()]

    if title in sheet_names:
        print(f"Worksheet '{title}' đã tồn tại, không cần tạo thêm.")
    else:
        # Tạo worksheet mới KHÔNG giới hạn hàng/cột
        new_worksheet = sheet.add_worksheet(
            title=title, rows=1000, cols=4
        )  # Bỏ rows và cols

        # Định nghĩa header
        headers = ["NAME", "DATE", "TIME", "CONTENT"]

        # Ghi header vào dòng đầu tiên
        new_worksheet.update(range_name="A1:D1", values=[headers])

        # Định dạng header
        header_format = CellFormat(
            backgroundColor=Color(70, 189, 198),
            textFormat=TextFormat(bold=True),
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            borders=Borders(
                top=Border("SOLID"),
                bottom=Border("SOLID"),
                left=Border("SOLID"),
                right=Border("SOLID"),
            ),
            wrapStrategy="WRAP",
        )

        # Định dạng chung cho TOÀN BỘ WORKSHEET (không giới hạn)
        body_format = CellFormat(
            horizontalAlignment="CENTER",
            verticalAlignment="MIDDLE",
            wrapStrategy="WRAP",
        )

        # Áp dụng định dạng header
        format_cell_range(new_worksheet, "A1:D1", header_format)

        # Áp dụng định dạng cho TOÀN BỘ DỮ LIỆU TƯƠNG LAI (phạm vi động)
        format_cell_range(
            new_worksheet, "A2:D", body_format
        )  # Z là cột xa nhất cần định dạng

        # Đặt độ rộng cột (tùy chọn)
        set_column_widths(
            new_worksheet, [("A", 186), ("B", 100), ("C", 100), ("D", 1020)]
        )
        # Ghim hàng header
        new_worksheet.freeze(rows=1)

        print(f"Đã tạo worksheet '{title}'!")


def get_message_all_group(driver):
    try:
        wait = WebDriverWait(driver, 20)
        print("Đang chờ danh sách chat tải...")
        # Chờ cho ít nhất một mục chat xuất hiện
        wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[starts-with(@id, '19:')]"))
        )
        elements = driver.find_elements(By.XPATH, "//*[starts-with(@id, '19:')]")
        print(f"Tìm thấy {len(elements)} nhóm chat.")

        for i, element in enumerate(elements):
            try:
                # Lấy lại element để tránh StaleElementReferenceException
                current_element = driver.find_elements(
                    By.XPATH, "//*[starts-with(@id, '19:')]"
                )[i]
                element_id = current_element.get_attribute("id")

                # Lấy tên nhóm từ title
                title_id = "title-chat-list-item_" + element_id
                title_element = wait.until(
                    EC.presence_of_element_located((By.ID, title_id))
                )
                title_text = title_element.get_attribute("title")
                print(f"\n--- Bắt đầu xử lý nhóm: {title_text} ---")

                create_worksheet(
                    title_text
                )  # Tạm thời comment out để tập trung vào lấy tin

                # Click vào nhóm chat
                actions = ActionChains(driver)
                actions.move_to_element(title_element).click().perform()
                print(f"Đã click vào nhóm '{title_text}'.")
                time.sleep(3)  # Chờ một chút để tin nhắn tải

                get_messege(driver, title_text)

            except Exception as e:
                print(f"Lỗi khi xử lý một nhóm chat: {e}")
                continue  # Bỏ qua và tiếp tục với nhóm tiếp theo

    except Exception as e:
        print(f"Lỗi khi lấy danh sách nhóm chat: {e}")
        save_screenshot(driver, "error_get_groups.png")


if __name__ == "__main__":
    driver = login()
    if driver:
        get_message_all_group(driver)
        driver.quit()
        print("✅ Hoàn tất công việc!")
