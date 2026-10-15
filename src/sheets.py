import os
import json
import gspread
from google.auth import default
from google.oauth2.service_account import Credentials
from gspread_formatting import set_column_widths, cellFormat, format_cell_range
from .config import SCOPES, SPREADSHEET_URL


# =========================
# GOOGLE SHEETS AUTHENTICATION
# =========================
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
    possible_paths = [
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
# CREATE WORKSHEET
# =========================
def create_worksheet(title, gc=None):
    if gc is None:
        gc = get_gspread_client()
    sheet = gc.open_by_url(SPREADSHEET_URL)

    names = [x.title for x in sheet.worksheets()]

    if title in names:
        ws = sheet.worksheet(title)
    else:
        ws = sheet.add_worksheet(title=title, rows=1000, cols=4)
        ws.update("A1:D1", [["NAME", "DATE", "TIME", "CONTENT"]])
        ws.freeze(rows=1)

    set_column_widths(
        ws,
        [
            ("A", 180),
            ("B", 100),
            ("C", 100),
            ("D", 1000),
        ],
    )

    fmt = cellFormat(wrapStrategy="WRAP")
    format_cell_range(ws, "D:D", fmt)
    print(f"✅ Đã cập nhật định dạng cho sheet: {title}")
    return ws


def _normalize_row_key(name, date_str, time_str, content):
    """Chuẩn hóa khóa nhận diện tin nhắn để chống trùng lặp chính xác."""
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

    existing_values = ws.get_all_values()
    existing_set = set()
    for r in existing_values[1:]:
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
        print(
            f"ℹ️ Tất cả {len(rows)} tin nhắn trong khung giờ đã có sẵn ở sheet: {worksheet}"
        )
