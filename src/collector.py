import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src.config import get_hcm_now, get_time_window
    from src.sheets import get_gspread_client, create_worksheet
    from src.teams import login, get_all_groups, open_chat_by_search, get_messages
else:
    from .config import get_hcm_now, get_time_window
    from .sheets import get_gspread_client, create_worksheet
    from .teams import login, get_all_groups, open_chat_by_search, get_messages


def run_collector():
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
        print(
            "💡 Lưu ý: Cần thêm file 'gcp-credentials.json' vào thư mục dự án để ghi dữ liệu lên Google Sheets."
        )

    driver = None
    group_names = []
    for attempt_login in range(5):
        driver = login()
        if driver:
            try:
                group_names = get_all_groups(driver)
                print("✅ Đăng nhập và tải danh sách nhóm Teams thành công!")
                break
            except Exception as e:
                print(
                    f"⚠️ Đăng nhập xong nhưng chưa tải được danh sách nhóm (lần {attempt_login + 1}/5): {e}"
                )
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = None
        print(f"⚠️ Thử đăng nhập lại lần {attempt_login + 1}/5 sau 5 giây...")
        time.sleep(5)

    if not driver:
        print("❌ Đăng nhập Teams không thành công sau 5 lần thử!")
        exit(1)

    try:
        for group in group_names:
            try:
                print(f"\n===== {group} =====")

                if gc_client:
                    create_worksheet(group, gc=gc_client)
                else:
                    print(
                        f"⚠️ Bỏ qua ghi Google Sheet cho '{group}' vì chưa có file gcp-credentials.json."
                    )

                if open_chat_by_search(driver, group):
                    if gc_client:
                        current_window = (start_time, get_hcm_now())
                        get_messages(
                            driver, group, gc=gc_client, time_window=current_window
                        )

                time.sleep(3)

            except Exception as e:
                print(f"⚠️ Bỏ qua nhóm {group} do lỗi:", e)
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        print("✅ ĐÃ HOÀN THÀNH")


if __name__ == "__main__":
    run_collector()
