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
import re
import html
from datetime import datetime, timezone
from google.auth import default
from google.oauth2.service_account import Credentials

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from gspread_formatting import *
import platform
import subprocess
import re

from dotenv import load_dotenv
load_dotenv()

gc.disable()

# =========================
# PATCH UNDETECTED CHROMEDRIVER CLEANUP
# =========================
# Sửa lỗi OSError: [WinError 6] The handle is invalid khi Python giải phóng bộ nhớ (garbage collection)
def _patch_uc_del():
    def _safe_del(self):
        try:
            self.quit()
        except Exception:
            pass
    uc.Chrome.__del__ = _safe_del

_patch_uc_del()

# =========================
# CONFIG
# =========================
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM/edit?usp=sharing"

email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")


# =========================
# SCREENSHOT
# =========================
def save_screenshot(driver, file_name="error.png"):
    try:
        driver.save_screenshot(file_name)
        print(f"📸 Đã lưu ảnh màn hình: {file_name}")
    except Exception as e:
        print(f"❌ Không thể lưu ảnh màn hình: {e}")

        
# =========================
# Kiểm tra version Chrome
# =========================       
def get_installed_chrome_major_version():
    """Tự động kiểm tra Major Version của Chrome trên máy"""

    system = platform.system()
    try:
        if system == "Windows":
            import winreg
            # Đọc phiên bản Chrome từ Registry Windows
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            except FileNotFoundError:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return int(version.split('.')[0])

        elif system == "Linux":
            output = subprocess.check_output(["google-chrome", "--version"]).decode("utf-8")
            match = re.search(r"Google Chrome (\d+)\.", output)
            if match:
                return int(match.group(1))

        elif system == "Darwin":  # macOS
            cmd = r"/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version"
            output = subprocess.check_output(cmd, shell=True).decode("utf-8")
            match = re.search(r"Google Chrome (\d+)\.", output)
            if match:
                return int(match.group(1))
    except Exception as e:
        print(f"⚠️ Không thể tự động phát hiện phiên bản Chrome: {e}")
    
    return None


# =========================
# NEW DRIVER FROM SOURCE 1
# =========================
def get_driver():
    options = uc.ChromeOptions()
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    # --- THÊM DÒNG NÀY ĐỂ TẮT PASSKEY ---
    options.add_argument("--disable-features=WebAuthentication,WebAuthenticationUI")
    
    # Cho phép tất cả Cookie (Teams cực kỳ cần cái này để không bị văng)
    prefs = {
        "profile.cookie_controls_mode": 0,
        "credentials_enable_service": False,      # Tắt popup hỏi lưu pass
        "profile.password_manager_enabled": False # Tắt trình quản lý mật khẩu
    } 
    
    options.add_experimental_option("prefs", prefs)

    options.page_load_strategy = "eager"
    options.add_argument("--lang=en-GB")

    proxy_url = os.getenv("PROXY_URL")
    if proxy_url:
        options.add_argument(f"--proxy-server={proxy_url}")
    chrome_version = None
    try:
        # Lệnh này sẽ chạy thành công trên máy chủ Ubuntu của GitHub Actions
        # Lấy output (ví dụ: "Google Chrome 147.0.7727.55")
        result = subprocess.check_output(["google-chrome", "--version"]).decode("utf-8")
        # Dùng Regex để tách lấy con số đầu tiên (147)
        chrome_version = int(re.search(r"\d+", result).group(0))
        print(
            f"✅ Đã tự động nhận diện Chrome trên máy chủ là version: {chrome_version}"
        )
    except Exception:
        # Nếu chạy thủ công trên Windows ở máy tính cá nhân nó sẽ nhảy vào đây
        chrome_version = get_installed_chrome_major_version()

    # Khởi tạo Driver với đúng phiên bản máy chủ đang có
    if chrome_version:
        driver = uc.Chrome(options=options, version_main=chrome_version)
    else:
        driver = uc.Chrome(options=options)
    
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'credentials', {
                    get: () => undefined
                });

                window.PublicKeyCredential = undefined;
            """
        }
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
        print("⏳ Đang đăng nhập...")

        # 1. Xử lý nút Sign in (nếu bị đẩy ra trang chờ)
        try:
            sign_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//button[contains(., "Sign in")] | //a[contains(., "Sign in")] | //button[contains(., "Đăng nhập")]',
                    )
                )
            )
            sign_btn.click()
        except:
            pass  # Bỏ qua nếu form điền email hiện ra trực tiếp

        # 2. Ô nhập Email (Sử dụng Selector linh hoạt cho Microsoft)
        email_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="email"], input[name="loginfmt"]')
            )
        )
        email_box.send_keys(email)
        email_box.send_keys(Keys.RETURN)

        time.sleep(3)
        # ====== THÊM ĐOẠN NÀY VÀO ======
        # Xử lý trường hợp Microsoft đòi gửi mã code, ép nó quay về dùng Mật khẩu
        try:
            use_pass_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]',
                    )
                )
            )
            use_pass_btn.click()
            time.sleep(2)
        except:
            pass  # Nếu màn hình đi thẳng tới ô mật khẩu thì cứ bỏ qua bước này
        # ===============================
        # ====== XỬ LÝ MÀN HÌNH "Sign in another way" -> CHỌN USE YOUR PASSWORD ======
        try:
            # 1. Nếu có màn hình "Other ways to sign in" thì bấm trước
            try:
                other_ways_btn = WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//*[contains(text(), "Other ways to sign in") or contains(text(), "Cách đăng nhập khác")]',
                        )
                    )
                )
                driver.execute_script("arguments[0].click();", other_ways_btn)
                print("👉 Đã chọn: Other ways to sign in")
                time.sleep(2)
            except:
                pass

            # 2. Chọn dòng "Use your password" (Dùng JS click để không bị trượt)
            use_pass_btn = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]/ancestor::div[@role="button"] '
                        '| //*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]',
                    )
                )
            )
            driver.execute_script("arguments[0].click();", use_pass_btn)
            print("👉 Đã kích hoạt: Use your password")
            time.sleep(3)
        except Exception as e:
            print("ℹ️ Bỏ qua chọn phương thức (hoặc đã ở màn hình nhập pass):", e)
        # ===========================================================================
        # 3. Ô nhập Password
        pass_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="password"], input[name="passwd"]')
            )
        )
        pass_box.send_keys(password)
        pass_box.send_keys(Keys.RETURN)

        # 4. Xử lý nút "Stay signed in?" (Chọn No để không lưu đăng nhập)
        try:
            print("⏳ Đang xử lý màn hình Stay signed in...")
            no_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[@id="declineButton"] | //*[@id="idBtn_Back"] | //*[@value="No"] | //button[contains(., "No")]',
                    )
                )
            )
            no_btn.click()
            time.sleep(3)
        except:
            print("⚠️ Không thấy màn hình Stay signed in, tiếp tục...")
            pass

        print("✅ Đăng nhập thành công")

        # Chờ giao diện Teams load hẳn
        time.sleep(15)

        return driver

    except Exception as e:
        save_screenshot(driver, "login_error.png")
        print("❌ Đăng nhập thất bại:", e)
        try:
            driver.quit()
        except Exception:
            pass
        return None


# =========================
# GOOGLE SHEETS AUTHENTICATION
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_gspread_client(credentials_path="gcp-credentials.json"):
    """Khởi tạo và xác thực với Google Sheets API."""
    # 1. Ưu tiên đọc chuỗi JSON từ các biến môi trường trong .env
    env_json = (
        os.getenv("GCP_SA_KEY")
        or os.getenv("GCP_CREDENTIALS_JSON")
        or os.getenv("GOOGLE_CREDENTIALS")
    )
    if env_json and env_json.strip().startswith("{"):
        try:
            info = json.loads(env_json)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
            return gspread.authorize(creds)
        except Exception as e:
            print(f"⚠️ Lỗi đọc credentials từ biến môi trường: {e}")

    # 2. Đọc từ file JSON trên ổ đĩa (thử nhiều đường dẫn)
    possible_paths = [ # Thêm các đường dẫn tương đối để linh hoạt hơn
        credentials_path,
        credentials_path + ".json",
        "gcp-credentials.json",
        "gcp-credentials.json.json",
        "credentials.json",
    ]
    target_path = None
    for path in possible_paths:
        if os.path.exists(path):
            target_path = path
            break

    if target_path:
        creds = Credentials.from_service_account_file(target_path, scopes=SCOPES)
        return gspread.authorize(creds)

    # 3. Thử dùng default credentials (hữu ích trên môi trường Google Cloud)
    try:
        creds, _ = default()
        return gspread.authorize(creds)
    except Exception as e:
        raise FileNotFoundError(
            "Không thể khởi tạo kết nối Google Sheets API. "
            "Vui lòng cấu hình biến GCP_SA_KEY / GCP_CREDENTIALS_JSON trong file .env hoặc thêm file 'gcp-credentials.json' vào thư mục."
        ) from e


# =========================
# CREATE SHEET
# =========================
def create_worksheet(title, gc=None):
    if gc is None:
        gc = get_gspread_client()
    sheet = gc.open_by_url(SPREADSHEET_URL)

    names = [x.title for x in sheet.worksheets()]

    if title in names:
        # Nếu sheet đã có, chúng ta lấy sheet đó để định dạng lại cho chắc chắn
        ws = sheet.worksheet(title)
    else:
        # Nếu chưa có thì mới tạo mới và thêm header
        ws = sheet.add_worksheet(title=title, rows=1000, cols=4)
        ws.update("A1:D1", [["NAME", "DATE", "TIME", "CONTENT"]])
        ws.freeze(rows=1)

    # ĐƯA PHẦN NÀY RA NGOÀI ĐỂ LUÔN THỰC THI:
    set_column_widths(
        ws,
        [
            ("A", 180),
            ("B", 100),
            ("C", 100),
            ("D", 1000),
        ],
    )

    # Ép kiểu xuống dòng (Wrap text) cho toàn bộ cột D
    fmt = cellFormat(wrapStrategy="WRAP")
    format_cell_range(ws, "D:D", fmt)
    print(f"✅ Đã cập nhật định dạng cho sheet: {title}")


# =========================
# SAVE DATA
# =========================
def save_to_excel(rows, worksheet, gc=None):
    if gc is None:
        gc = get_gspread_client()
    sheet = gc.open_by_url(SPREADSHEET_URL)
    ws = sheet.worksheet(worksheet)

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        print(f"✅ Đã thêm {len(rows)} dòng vào sheet: {worksheet}")


# =========================
# GET MESSAGE
# =========================
def get_messages(driver, worksheet, gc=None):
    try:
        wait = WebDriverWait(driver, 20)

        pane = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-tid="message-pane-list-runway"]')
            )
        )

        items = pane.find_elements(By.CSS_SELECTOR, '[data-tid="chat-pane-item"]')

        data = []

        for item in items:
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

                # --- CÁCH XỬ LÝ TẬN GỐC BẰNG HTML ---
                content_el = item.find_element(By.CSS_SELECTOR, '[id^="content-"]')

                raw_html = content_el.get_attribute("innerHTML")

                # 1. MỚI: Xóa các thẻ inline (mention, span, link) TRƯỚC để tránh bị cắt vụn chữ
                text = re.sub(
                    r"</?(span|at|a|strong|b|i|em)[^>]*>",
                    "",
                    raw_html,
                    flags=re.IGNORECASE,
                )

                # 2. Chủ động thay thế các thẻ ngắt dòng phổ biến thành ký tự \n
                text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
                text = re.sub(r"</(div|p)>", "\n", text, flags=re.IGNORECASE)

                # 3. Xóa sạch mọi thẻ HTML còn sót lại
                text = re.sub(r"<[^>]+>", "", text)

                # 4. Dịch các ký tự đặc biệt của web (như &nbsp; thành dấu cách)
                text = html.unescape(text)

                # 5. Dọn dẹp khoảng trắng thừa và nối lại thành đoạn văn hoàn chỉnh
                lines = [line.strip() for line in text.split("\n")]
                content = "\n".join([line for line in lines if line])
                # ------------------------------------
                data.append([name, date_str, time_str, content])
            except:
                continue

        save_to_excel(data, worksheet, gc=gc)

    except Exception as e:
        save_screenshot(driver, "get_messages.png")
        print("❌ Lỗi khi lấy tin nhắn (get_messages):", e)


# =========================
# SEARCH CHAT FROM SOURCE 1
# =========================
def open_chat_by_search(driver, chat_name):
    wait = WebDriverWait(driver, 20)
    chat_item_xpath = '//*[contains(@data-tid, "chat-list") or contains(@data-tid, "chat-item") or @role="treeitem" or @role="listitem"]'

    try:
        # 1. Chờ danh sách tải xong bằng XPath mới
        wait.until(EC.presence_of_element_located((By.XPATH, chat_item_xpath)))
        groups = driver.find_elements(By.XPATH, chat_item_xpath)

        for g in groups:
            # Lấy tất cả các dòng text, bỏ khoảng trắng
            lines = [x.strip() for x in g.text.splitlines() if x.strip()]
            
            if not lines:
                continue
                
            # Loại bỏ nhãn "Unread" nếu có tin nhắn mới
            if lines[0] == "Unread":
                lines.pop(0)
                
            if not lines:
                continue
                
            txt = lines[0]

            if not txt:
                txt = g.get_attribute("aria-label") or ""

            # 3. SO SÁNH
            # Dùng startswith phòng trường hợp Teams hiển thị dấu "..." ở đuôi
            if txt == chat_name or chat_name.startswith(txt.replace("...", "").strip()):
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", g)
                time.sleep(1)
                g.click()
                time.sleep(5)
                print(f"📂 Đã mở đúng nhóm: {chat_name}")
                return True

        print(f"⚠️ Không thấy {chat_name} ở ngoài, thử dùng thanh Filter bên trái...")
        
        # 1. Thử bấm vào icon Kính lúp (Filter) kế bên chữ Chat để mở ô nhập liệu (nếu nó đang ẩn)
        try:
            # Ưu tiên bắt theo data-testid cố định của hệ thống
            filter_icon_xpath = (
                '//button[@data-testid="simple-collab-left-rail-header-sticky-filter-v2-button"]'
                ' | //button[contains(@aria-keyshortcuts, "Ctrl+Shift+F")]'
            )
            filter_icon = driver.find_element(By.XPATH, filter_icon_xpath)
            driver.execute_script("arguments[0].click();", filter_icon)
            time.sleep(1)
        except Exception:
            pass # Bỏ qua nếu ô Filter đã hiển thị sẵn (giống trong ảnh bạn chụp)

        # 2. Định vị chính xác ô input của thanh Filter bên trái
        left_search_xpath = (
            '//input[@data-testid="simple-collab-left-rail-sticky-filter-input"]'
            ' | //input[@id="simple-collab-left-rail-sticky-filter-input-id"]'
        )
        
        search = wait.until(EC.presence_of_element_located((By.XPATH, left_search_xpath)))
        
        # 3. Điền tên nhóm cần tìm
        search.click()
        search.send_keys(Keys.CONTROL + "a")
        search.send_keys(Keys.BACKSPACE)
        search.send_keys(chat_name)

        time.sleep(4)

        # 4. Lấy kết quả chính xác xuất hiện trong danh sách sau khi lọc
        filtered_result_xpath = f"//span[normalize-space(text())='{chat_name}']"
        
        filtered_result = wait.until(
            EC.presence_of_element_located((By.XPATH, filtered_result_xpath))
        )
        
        # Dùng JS Click để đảm bảo không bị lỗi "element not interactable"
        driver.execute_script("arguments[0].click();", filtered_result)

        time.sleep(5)
        print(f"📂 Đã mở nhóm qua thanh Filter bên trái: {chat_name}")
        return True

    except Exception as e:
        save_screenshot(driver, "open_chat_error.png")
        print(f"❌ Không thể mở nhóm {chat_name}:", e)
        return False
# =========================
# GET ALL GROUPS
# =========================
def get_all_groups(driver):

    wait = WebDriverWait(driver, 20)

    # mở tab Chat
    try:
        chat_btn = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//button[contains(@data-tid,"app-bar-chat") '
                    'or contains(@aria-label,"Chat") '
                    'or contains(@aria-label,"Trò chuyện")]'
                )
            )
        )

        driver.execute_script("arguments[0].click()", chat_btn)

    except Exception as e:
        print(f"⚠️ Không thể click nút Chat, có thể đã ở trong tab Chat rồi. Lỗi: {e}")
        pass

    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-item-type="chat"]')
            )
        )
    except Exception as e:
        print("❌ Lỗi timeout khi chờ danh sách chat tải. Đang chụp ảnh màn hình...")
        save_screenshot(driver, "get_all_groups_error.png")
        raise e # Ném lại lỗi để chương trình dừng lại như cũ

    chat_items = driver.find_elements(By.CSS_SELECTOR, '[data-item-type="chat"]')

    groups = []

    for item in chat_items:

        try:

            lines = [
                x.strip()
                for x in item.text.splitlines()
                if x.strip()
            ]

            if not lines:
                continue

            # bỏ chữ Unread
            if lines[0] == "Unread":
                lines.pop(0)

            if not lines:
                continue

            name = lines[0]

            if name not in groups:
                groups.append(name)

        except Exception:
            continue

    print("=" * 60)
    print(f"Tổng cộng: {len(groups)} nhóm")

    for g in groups:
        print(" -", g)
    
    return groups
# =========================
# MAIN
# =========================
if __name__ == "__main__":
    gc_client = None
    try:
        gc_client = get_gspread_client()
        print("✅ Kết nối Google Sheets API thành công!")
    except Exception as e:
        print(f"⚠️ Cảnh báo kết nối Google Sheets: {e}")
        print("💡 Lưu ý: Cần thêm file 'gcp-credentials.json' vào thư mục dự án để ghi dữ liệu lên Google Sheets.")

    driver = login()

    if driver:
        try:
            group_names = get_all_groups(driver)

            for group in group_names:
                try:
                    print(f"\n===== {group} =====")

                    if gc_client:
                        create_worksheet(group, gc=gc_client)
                    else:
                        print(f"⚠️ Bỏ qua ghi Google Sheet cho '{group}' vì chưa có file gcp-credentials.json.")

                    if open_chat_by_search(driver, group):
                        if gc_client:
                            get_messages(driver, group, gc=gc_client)

                    time.sleep(3)

                except Exception as e:
                    print(f"⚠️ Bỏ qua nhóm {group} do lỗi:", e)
        finally:
            try:
                driver.quit()
            except Exception:
                pass
            print("✅ ĐÃ HOÀN THÀNH")
