import os
import gc
import time
import pytz
import re
from datetime import datetime, timezone, timedelta, time as dt_time
from dotenv import load_dotenv

load_dotenv()

gc.disable()

# =========================
# TIMEZONE CONFIG
# =========================
os.environ["TZ"] = "Asia/Ho_Chi_Minh"
if hasattr(time, "tzset"):
    time.tzset()

local_tz = pytz.timezone("Asia/Ho_Chi_Minh")

# =========================
# APP & CREDENTIAL CONFIG
# =========================
SPREADSHEET_URL = os.environ.get(
    "SPREADSHEET_URL",
    "https://docs.google.com/spreadsheets/d/1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM/edit?usp=sharing",
)

email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")
PROXY_URL = os.getenv("PROXY_URL")

# =========================
# GOOGLE SHEETS SCOPES
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# =========================
# TIME WINDOW HELPERS
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
