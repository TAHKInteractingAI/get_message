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
from datetime import datetime, timezone, timedelta, time as dt_time
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
os.environ["TZ"] = "Asia/Ho_Chi_Minh"
if hasattr(time, "tzset"):
    time.tzset()

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
# Kiểm tra version & binary Browser
# =========================       
def get_browser_binary_and_version():
    """Tự động kiểm tra đường dẫn binary và Major Version của trình duyệt."""
    system = platform.system()
    binary_path = None
    major_version = None

    try:
        if system == "Windows":
            import winreg
            # Đọc phiên bản Chrome từ Registry Windows
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            except FileNotFoundError:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            major_version = int(version.split('.')[0])

        elif system == "Linux":
            for cmd_name in ["google-chrome", "google-chrome-stable", "chromium", "brave-browser"]:
                try:
                    output = subprocess.check_output([cmd_name, "--version"]).decode("utf-8")
                    # Ưu tiên bắt chính xác số đứng sau tên trình duyệt
                    match = re.search(r"(?:Google Chrome|Chromium|Brave Browser)\s+(\d+)\.", output, re.IGNORECASE)
                    if not match:
                        match = re.search(r"\b(\d{2,3})\.\d+", output)
                    if match:
                        major_version = int(match.group(1))
                        print(f"🔍 Nhận diện {cmd_name}: {output.strip()} -> major: {major_version}")
                        break
                except Exception:
                    pass

        elif system == "Darwin":  # macOS
            candidates = [
                ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", r"Google Chrome (\d+)\."),
                ("/Applications/Brave Browser.app/Contents/MacOS/Brave Browser", r"Brave Browser (\d+)\."),
                ("/Applications/Chromium.app/Contents/MacOS/Chromium", r"Chromium (\d+)\."),
                ("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", r"Microsoft Edge (\d+)\.")
            ]
            for path, regex in candidates:
                if os.path.exists(path):
                    binary_path = path
                    try:
                        output = subprocess.check_output([path, "--version"]).decode("utf-8")
                        match = re.search(regex, output)
                        if match:
                            major_version = int(match.group(1))
                        break
                    except Exception as e:
                        print(f"⚠️ Không lấy được version của {path}: {e}")

    except Exception as e:
        print(f"⚠️ Không thể tự động phát hiện phiên bản trình duyệt: {e}")
    
    return binary_path, major_version


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

    binary_path, chrome_version = get_browser_binary_and_version()
    if binary_path:
        print(f"🌐 Sử dụng trình duyệt: {binary_path}")
        options.binary_location = binary_path

    if chrome_version:
        print(f"✅ Đã nhận diện phiên bản trình duyệt: {chrome_version}")

    kwargs = {"options": options}
    if chrome_version:
        kwargs["version_main"] = chrome_version
    if binary_path:
        kwargs["browser_executable_path"] = binary_path

    driver = uc.Chrome(**kwargs)

    try:
        driver.execute_cdp_cmd(
            "Emulation.setTimezoneOverride",
            {"timezoneId": "Asia/Ho_Chi_Minh"},
        )
    except Exception:
        pass
    
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

        # 5. KIỂM TRA MÀN HÌNH "TRY AGAIN LATER" DO BỊ MICROSOFT RATE LIMIT
        time.sleep(3)
        page_source = driver.page_source.lower()
        if "try again later" in page_source or "we can't sign you in right now" in page_source:
            save_screenshot(driver, "try_again_later_blocked.png")
            raise Exception("❌ Microsoft tạm khóa đăng nhập (Try again later) do tần suất đăng nhập từ IP này quá nhiều. Vui lòng thử lại sau!")

        # 6. Chờ giao diện Teams load hẳn (Thoát khỏi màn hình Splash Screen tím)
        print("⏳ Đang chờ giao diện Teams nạp xong (tối đa 45s)...")
        start_wait = time.time()
        while time.time() - start_wait < 45:
            current_url = driver.current_url
            # Kiểm tra xem có popup "Use the web app instead" trên Linux không
            try:
                use_web_btn = driver.find_elements(
                    By.XPATH,
                    '//a[contains(text(), "Use the web app") or contains(text(), "Dùng ứng dụng web")] | '
                    '//button[contains(text(), "Use the web app") or contains(text(), "Dùng ứng dụng web")]'
                )
                for btn in use_web_btn:
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].click();", btn)
                        print("👉 Đã bấm nút 'Use the web app instead'")
                        time.sleep(2)
            except Exception:
                pass

            # Kiểm tra nếu đã xuất hiện các nút bấm của giao diện chính Teams
            nav_items = driver.find_elements(
                By.XPATH,
                '//button | //div[@role="button"] | //div[@role="tab"] | //div[@role="listitem"]'
            )
            if len(nav_items) > 5 and "login.microsoftonline.com" not in current_url:
                print(f"✅ Đăng nhập Teams thành công sau {int(time.time() - start_wait)} giây!")
                return driver
            time.sleep(3)

        print("⚠️ Hết thời gian chờ Splash Screen, tiếp tục tiến trình...")
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


def _normalize_row_key(name, date_str, time_str, content):
    """Chuẩn hóa khóa nhận diện tin nhắn để chống trùng lặp chính xác (kể cả khi Google Sheets bỏ số 0 ở đầu giờ)."""
    norm_name = str(name).strip()
    norm_date = str(date_str).strip()
    norm_time = str(time_str).strip().lstrip("0") or "0"
    norm_content = str(content).strip().lstrip("'")
    return (norm_name, norm_date, norm_time, norm_content)


# =========================
# SAVE DATA
# =========================
def save_to_excel(rows, worksheet, gc=None):
    if gc is None:
        gc = get_gspread_client()
    sheet = gc.open_by_url(SPREADSHEET_URL)
    ws = sheet.worksheet(worksheet)

    if not rows:
        print(f"ℹ️ Không có tin nhắn nào trong khung giờ cho sheet: {worksheet}")
        return

    # Kiểm tra dữ liệu đã có trên sheet để tránh ghi trùng lặp
    existing_values = ws.get_all_values()
    existing_set = set()
    for r in existing_values[1:]:  # Bỏ qua dòng tiêu đề
        if len(r) >= 4:
            existing_set.add(_normalize_row_key(r[0], r[1], r[2], r[3]))

    new_rows = []
    for r in rows:
        key = _normalize_row_key(r[0], r[1], r[2], r[3])
        if key not in existing_set:
            existing_set.add(key)
            new_rows.append(r)

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")
        print(
            f"✅ Đã thêm {len(new_rows)} dòng mới vào sheet: {worksheet} "
            f"(bỏ qua {len(rows) - len(new_rows)} dòng trùng)"
        )
    else:
        print(f"ℹ️ Tất cả {len(rows)} tin nhắn trong khung giờ đã có sẵn ở sheet: {worksheet}")


# =========================
# TIME WINDOW (3 CA CRONJOB: 02:00, 13:00, 18:00 - GIỜ HỒ CHÍ MINH)
# =========================
def get_hcm_now():
    """Luôn lấy giờ hiện tại chuẩn theo múi giờ Asia/Ho_Chi_Minh (UTC+7), bất kể giờ hệ thống của máy chủ CI."""
    return datetime.now(timezone.utc).astimezone(local_tz)


def get_time_window(now=None, buffer_minutes=10):
    """
    Xác định khung giờ lấy tin nhắn nối tiếp giữa 3 lần chạy cronjob (theo múi giờ Asia/Ho_Chi_Minh):
    - Ca 02:00 sáng (00:00 -> trước 08:00): Lấy từ 18:00 hôm qua -> hiện tại
    - Ca 13:00 chiều (08:00 -> trước 16:00): Lấy từ 02:00 sáng nay -> hiện tại
    - Ca 18:00 tối (16:00 -> 23:59): Lấy từ 13:00 chiều nay -> hiện tại
    Có trừ lùi `buffer_minutes` (mặc định 10 phút) ở mốc bắt đầu để tránh sót tin nhắn lúc giao ca
    khi CI khởi động trễ vài phút (đã có bộ lọc chống trùng khi ghi vào Sheet).
    """
    if now is None:
        now = get_hcm_now()
    elif now.tzinfo is None:
        now = local_tz.localize(now)
    else:
        now = now.astimezone(local_tz)

    # Nếu CI truyền thời điểm kích hoạt (CI_TRIGGER_TIME), dùng giờ kích hoạt để xác định đúng ca
    ref_time = now
    ci_trigger_str = os.getenv("CI_TRIGGER_TIME", "").strip()
    if ci_trigger_str:
        try:
            parsed_trigger = datetime.strptime(ci_trigger_str, "%Y-%m-%d %H:%M:%S")
            ref_time = local_tz.localize(parsed_trigger)
        except Exception:
            ref_time = now

    ref_date = ref_time.date()
    ref_hour = ref_time.hour

    if ref_hour < 8:
        shift_name = "Ca 02:00 AM (02h sáng)"
        yesterday = ref_date - timedelta(days=1)
        start_naive = datetime.combine(yesterday, dt_time(18, 0, 0))
    elif ref_hour < 16:
        shift_name = "Ca 01:00 PM (13h chiều)"
        start_naive = datetime.combine(ref_date, dt_time(2, 0, 0))
    else:
        shift_name = "Ca 06:00 PM (18h tối)"
        start_naive = datetime.combine(ref_date, dt_time(13, 0, 0))

    start_time = local_tz.localize(start_naive) - timedelta(minutes=buffer_minutes)
    end_time = max(now, ref_time)
    return start_time, end_time, shift_name


def _parse_teams_timestamp(timestamp_str):
    """
    Chuyển chuỗi datetime từ thẻ <time> của Teams sang múi giờ local_tz (Asia/Ho_Chi_Minh).
    Xử lý an toàn cả định dạng 7 chữ số thập phân (.1234567Z của Microsoft), không có mili-giây, hoặc có offset (+00:00).
    """
    if not timestamp_str:
        return None
    ts = timestamp_str.strip()
    # Chuẩn hóa phần thập phân dài hơn 6 chữ số (ví dụ .1234567Z -> .123456Z)
    ts = re.sub(r"\.(\d{6})\d+(Z|[+-]\d{2}:?\d{2})$", r".\1\2", ts)
    try:
        iso_ts = ts[:-1] + "+00:00" if ts.endswith("Z") else ts
        dt_obj = datetime.fromisoformat(iso_ts)
        if dt_obj.tzinfo is None:
            dt_obj = dt_obj.replace(tzinfo=timezone.utc)
        return dt_obj.astimezone(local_tz)
    except Exception:
        pass

    for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ"):
        try:
            dt_utc = datetime.strptime(ts, fmt).replace(tzinfo=timezone.utc)
            return dt_utc.astimezone(local_tz)
        except ValueError:
            continue
    return None


# =========================
# GET MESSAGE
# =========================
def get_messages(driver, worksheet, gc=None, time_window=None, max_scrolls=15):
    try:
        wait = WebDriverWait(driver, 20)

        pane = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, '[data-tid="message-pane-list-runway"]')
            )
        )

        if time_window is None:
            start_time, end_time, _ = get_time_window()
        else:
            start_time, end_time = time_window

        # Lưu các tin nhắn thu thập được theo key để tránh trùng khi cuộn trang (Virtual DOM)
        collected_messages = {}
        prev_oldest_dt = None
        unchanged_scrolls = 0

        for scroll_idx in range(max_scrolls):
            items = pane.find_elements(By.CSS_SELECTOR, '[data-tid="chat-pane-item"]')
            if not items:
                break

            oldest_dt_in_batch = None
            last_author_name = ""
            last_dt_local = None

            for item in items:
                try:
                    try:
                        name = item.find_element(
                            By.CSS_SELECTOR, '[data-tid="message-author-name"]'
                        ).text.strip()
                        if name:
                            last_author_name = name
                    except Exception:
                        name = last_author_name

                    dt_local = None
                    try:
                        timestamp = item.find_element(By.TAG_NAME, "time").get_attribute(
                            "datetime"
                        )
                        dt_local = _parse_teams_timestamp(timestamp)
                        if dt_local is not None:
                            last_dt_local = dt_local
                    except Exception:
                        dt_local = last_dt_local

                    if dt_local is None:
                        continue

                    if oldest_dt_in_batch is None or dt_local < oldest_dt_in_batch:
                        oldest_dt_in_batch = dt_local

                    # Chỉ thu thập tin nhắn nằm trong khung giờ [start_time, end_time]
                    if not (start_time <= dt_local <= end_time):
                        continue

                    date_str = dt_local.strftime("%Y-%m-%d")
                    time_str = dt_local.strftime("%H:%M:%S")

                    # --- CÁCH XỬ LÝ TẬN GỐC BẰNG HTML ---
                    content_el = item.find_element(By.CSS_SELECTOR, '[id^="content-"]')
                    raw_html = content_el.get_attribute("innerHTML")

                    # 1. Xóa các thẻ inline (mention, span, link) TRƯỚC để tránh bị cắt vụn chữ
                    text = re.sub(
                        r"</?(span|at|a|strong|b|i|em|u|s|code|font)[^>]*>",
                        "",
                        raw_html,
                        flags=re.IGNORECASE,
                    )

                    # 2. Chủ động thay thế cả thẻ mở & thẻ đóng dạng block và <br> thành ký tự \n
                    #    (Tránh lỗi dòng đầu là bare text nối trực tiếp vào <div>dòng 2</div> làm dính chữ)
                    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
                    text = re.sub(
                        r"</?(div|p|li|ul|ol|tr|table|blockquote|section|article|h[1-6])[^>]*>",
                        "\n",
                        text,
                        flags=re.IGNORECASE,
                    )

                    # 3. Xóa sạch mọi thẻ HTML còn sót lại
                    text = re.sub(r"<[^>]+>", "", text)

                    # 4. Dịch các ký tự đặc biệt của web (như &nbsp; thành dấu cách)
                    text = html.unescape(text).replace("\xa0", " ")

                    # 5. Dọn dẹp khoảng trắng thừa và nối lại thành đoạn văn hoàn chỉnh
                    lines = [line.strip() for line in text.splitlines()]
                    content = "\n".join([line for line in lines if line])
                    # ------------------------------------

                    if not content:
                        continue

                    msg_key = _normalize_row_key(name, date_str, time_str, content)
                    if msg_key not in collected_messages:
                        collected_messages[msg_key] = (
                            dt_local,
                            [name, date_str, time_str, content],
                        )
                except Exception:
                    continue

            # Kiểm tra điều kiện dừng cuộn
            if oldest_dt_in_batch is None:
                break

            if oldest_dt_in_batch <= start_time:
                # Đã chạm tới các tin nhắn cũ hơn mốc bắt đầu của khung giờ
                break

            if prev_oldest_dt is not None and oldest_dt_in_batch >= prev_oldest_dt:
                unchanged_scrolls += 1
                if unchanged_scrolls >= 2:
                    # Đã cuộn tới đỉnh lịch sử chat, không còn tin nhắn cũ hơn để tải
                    break
            else:
                unchanged_scrolls = 0

            prev_oldest_dt = oldest_dt_in_batch

            # Cuộn lên đầu danh sách để Teams nạp thêm tin nhắn cũ hơn (kết hợp JS scroll + phím PageUp cho Virtual List)
            print(
                f"🔄 Đang cuộn lên tải thêm tin nhắn cũ (mốc cũ nhất hiện tại: "
                f"{oldest_dt_in_batch.strftime('%Y-%m-%d %H:%M:%S')})..."
            )
            try:
                first_item = items[0]
                driver.execute_script(
                    """
                    const firstItem = arguments[0];
                    firstItem.scrollIntoView({block: 'start'});
                    const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
                    if (viewport) {
                        viewport.scrollTop = 0;
                        viewport.dispatchEvent(new Event('scroll', { bubbles: true }));
                    }
                    let parent = firstItem.parentElement;
                    while (parent) {
                        if (parent.scrollHeight > parent.clientHeight) {
                            parent.scrollTop = 0;
                            parent.dispatchEvent(new Event('scroll', { bubbles: true }));
                        }
                        parent = parent.parentElement;
                    }
                    """,
                    first_item,
                )
                try:
                    ActionChains(driver).move_to_element(first_item).click().send_keys(
                        Keys.PAGE_UP
                    ).perform()
                except Exception:
                    pass
            except Exception:
                pass
            time.sleep(2.5)

        # Sắp xếp tin nhắn theo thứ tự thời gian từ cũ đến mới
        sorted_items = sorted(collected_messages.values(), key=lambda x: x[0])
        data = [row for _, row in sorted_items]

        print(
            f"📊 Nhóm '{worksheet}': Lọc được {len(data)} tin nhắn trong khung giờ "
            f"[{start_time.strftime('%Y-%m-%d %H:%M:%S')} -> {end_time.strftime('%Y-%m-%d %H:%M:%S')}]"
        )

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
def dismiss_teams_popups(driver):
    """Tự động đóng các banner và popup làm phiền sau đăng nhập."""
    try:
        # 1. Đóng banner Turn on notifications
        close_banner_btns = driver.find_elements(
            By.XPATH,
            '//button[contains(@aria-label, "Dismiss") or contains(@aria-label, "Đóng") or contains(@aria-label, "Close")]'
        )
        for btn in close_banner_btns:
            try:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
            except Exception:
                pass
    except Exception:
        pass


def get_all_groups(driver):

    wait = WebDriverWait(driver, 25)

    dismiss_teams_popups(driver)

    # mở tab Chat
    try:
        chat_btn = WebDriverWait(driver, 8).until(
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
        time.sleep(2)
    except Exception as e:
        print(f"ℹ️ Tab Chat có thể đã được chọn sẵn. ({e})")
        pass

    # Bộ selector đa năng cho danh sách chat (tương thích cả Teams Classic lẫn Teams 2.1 Mới)
    chat_item_xpath = (
        '//*[contains(@data-tid, "chat-list-item") '
        'or contains(@data-tid, "chat-item") '
        'or @data-item-type="chat" '
        'or @role="listitem" '
        'or @role="treeitem"]'
    )

    try:
        # Chờ danh sách nhóm xuất hiện (tối đa 25 giây)
        wait.until(EC.presence_of_element_located((By.XPATH, chat_item_xpath)))
    except Exception as e:
        print("❌ Lỗi timeout khi chờ danh sách chat tải. Đang chụp ảnh màn hình...")
        save_screenshot(driver, "get_all_groups_error.png")
        raise e # Ném lại lỗi để chương trình dừng lại như cũ

    chat_items = driver.find_elements(By.XPATH, chat_item_xpath)

    groups = []

    for item in chat_items:

        try:
            if not item.is_displayed():
                continue

            lines = [
                x.strip()
                for x in item.text.splitlines()
                if x.strip()
            ]

            if not lines:
                continue

            # Bỏ nhãn Unread hoặc các nhãn hệ thống
            if lines[0] in ["Unread", "Chưa đọc"]:
                lines.pop(0)

            if not lines:
                continue

            name = lines[0]

            # Bỏ qua các tiêu đề phân mục như 'Favorites', 'Chats', 'Pinned'
            if name in ["Favorites", "Chats", "Pinned", "Yêu thích", "Trò chuyện"]:
                continue

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
    utc_now = datetime.now(timezone.utc)
    hcm_now = get_hcm_now()
    start_time, end_time, shift_name = get_time_window(hcm_now)
    ci_trigger_env = os.getenv("CI_TRIGGER_TIME", "N/A (Chạy trực tiếp)")
    ci_event_name = os.getenv("GITHUB_EVENT_NAME", "local_run")

    print("=" * 65)
    print("⏰ KIỂM TRA THỜI GIAN & CA CHẠY (MÚI GIỜ HỒ CHÍ MINH - UTC+7)")
    print(f" - Sự kiện kích hoạt (CI Event) : {ci_event_name}")
    print(f" - Giờ CI nhận lệnh (HCM)       : {ci_trigger_env}")
    print(f" - Giờ máy chủ hiện tại (UTC)   : {utc_now.strftime('%Y-%m-%d %H:%M:%S UTC')}")
    print(f" - Giờ Việt Nam hiện tại (HCM)  : {hcm_now.strftime('%Y-%m-%d %H:%M:%S %Z%z')}")
    print(f" - Ca chạy được nhận diện       : {shift_name}")
    print(
        f" - Khung giờ lấy tin nhắn       : "
        f"{start_time.strftime('%Y-%m-%d %H:%M:%S')} -> {end_time.strftime('%Y-%m-%d %H:%M:%S')} (Asia/Ho_Chi_Minh)"
    )
    print("=" * 65)

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
                            current_window = (start_time, get_hcm_now())
                            get_messages(driver, group, gc=gc_client, time_window=current_window)

                    time.sleep(3)

                except Exception as e:
                    print(f"⚠️ Bỏ qua nhóm {group} do lỗi:", e)
        finally:
            try:
                driver.quit()
            except Exception:
                pass
            print("✅ ĐÃ HOÀN THÀNH")
