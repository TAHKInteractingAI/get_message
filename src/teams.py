import time
import html
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from .config import (
    email,
    password,
    get_time_window,
    _parse_teams_timestamp,
    get_hcm_now,
)
from .browser import get_driver, save_screenshot
from .sheets import _normalize_row_key, save_to_excel


# =========================
# LOGIN TEAMS
# =========================
def login():
    driver = get_driver()
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
        except Exception:
            pass

        # 2. Ô nhập Email
        email_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="email"], input[name="loginfmt"]')
            )
        )
        email_box.send_keys(email)
        time.sleep(1)
        email_box.send_keys(Keys.RETURN)

        time.sleep(3)

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
        except Exception:
            pass

        # Xử lý màn hình "Sign in another way" -> Chọn Use your password
        try:
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
            except Exception:
                pass

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

        # 3. Ô nhập Password
        pass_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="password"], input[name="passwd"]')
            )
        )
        pass_box.send_keys(password)
        time.sleep(1)
        pass_box.send_keys(Keys.RETURN)
        time.sleep(3)

        # 4. Xử lý nút "Stay signed in?"
        try:
            print("⏳ Đang xử lý màn hình Stay signed in...")
            no_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[@id="declineButton"] | //*[@id="idBtn_Back"] | //*[@value="No"] '
                        '| //button[@data-testid="secondaryButton"] '
                        '| //button[normalize-space(.)="No" or normalize-space(.)="Không"]',
                    )
                )
            )
            no_btn.click()
            time.sleep(3)
        except Exception:
            print("⚠️ Không thấy màn hình Stay signed in, tiếp tục...")

        # 5. Chờ giao diện Teams load hẳn (tối đa 45s)
        print("⏳ Đang chờ giao diện Teams nạp xong (tối đa 45s)...")
        start_wait = time.time()
        while time.time() - start_wait < 45:
            current_url = (driver.current_url or "").lower()

            try:
                body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                if (
                    "try again later" in body_text
                    or "can't sign you in right now" in body_text
                ):
                    save_screenshot(driver, "try_again_later_blocked.png")
                    raise RuntimeError(
                        "Microsoft tạm khóa phiên đăng nhập (Try again later / We can't sign you in right now)."
                    )
            except RuntimeError:
                raise
            except Exception:
                pass

            try:
                use_web_btn = driver.find_elements(
                    By.XPATH,
                    '//a[contains(text(), "Use the web app") or contains(text(), "Dùng ứng dụng web")] | '
                    '//button[contains(text(), "Use the web app") or contains(text(), "Dùng ứng dụng web")]',
                )
                for btn in use_web_btn:
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].click();", btn)
                        print("👉 Đã bấm nút 'Use the web app instead'")
                        time.sleep(2)
            except Exception:
                pass

            is_login_domain = any(
                d in current_url
                for d in (
                    "login.microsoftonline.com",
                    "login.live.com",
                    "account.live.com",
                )
            )
            nav_items = driver.find_elements(
                By.XPATH,
                '//button | //div[@role="button"] | //div[@role="tab"] | //div[@role="listitem"] | //*[@role="treeitem"]',
            )
            if len(nav_items) > 5 and not is_login_domain:
                print(
                    f"✅ Đăng nhập Teams thành công sau {int(time.time() - start_wait)} giây!"
                )
                return driver
            time.sleep(3)

        current_url = (driver.current_url or "").lower()
        if any(
            d in current_url
            for d in (
                "login.microsoftonline.com",
                "login.live.com",
                "account.live.com",
            )
        ):
            save_screenshot(driver, "login_stuck_on_auth.png")
            raise RuntimeError(
                f"Vẫn kẹt ở trang xác thực Microsoft sau 45s (URL: {driver.current_url})"
            )

        print(
            "⚠️ Hết thời gian chờ Splash Screen, chờ thêm 10s để giao diện ổn định..."
        )
        time.sleep(10)
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
# POPUP DISMISSAL
# =========================
def dismiss_teams_popups(driver):
    """Tự động đóng các banner và popup làm phiền sau đăng nhập."""
    try:
        close_banner_btns = driver.find_elements(
            By.XPATH,
            '//button[contains(@aria-label, "Dismiss") or contains(@aria-label, "Đóng") or contains(@aria-label, "Close")]',
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


# =========================
# GET ALL GROUPS
# =========================
def get_all_groups(driver):
    wait = WebDriverWait(driver, 25)
    dismiss_teams_popups(driver)

    try:
        chat_btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//button[contains(@data-tid,"app-bar-chat") '
                    'or contains(@aria-label,"Chat") '
                    'or contains(@aria-label,"Trò chuyện")]',
                )
            )
        )
        driver.execute_script("arguments[0].click()", chat_btn)
        time.sleep(2)
    except Exception as e:
        print(f"ℹ️ Tab Chat có thể đã được chọn sẵn. ({e})")

    chat_item_xpath = (
        '//*[contains(@data-tid, "chat-list-item") '
        'or contains(@data-tid, "chat-item") '
        'or @data-item-type="chat" '
        'or @role="listitem" '
        'or @role="treeitem"]'
    )

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, chat_item_xpath)))
    except Exception as e:
        print("❌ Lỗi timeout khi chờ danh sách chat tải. Đang chụp ảnh màn hình...")
        save_screenshot(driver, "get_all_groups_error.png")
        raise e

    chat_items = driver.find_elements(By.XPATH, chat_item_xpath)
    groups = []

    for item in chat_items:
        try:
            if not item.is_displayed():
                continue

            lines = [x.strip() for x in item.text.splitlines() if x.strip()]
            if not lines:
                continue

            if lines[0] in ["Unread", "Chưa đọc"]:
                lines.pop(0)

            if not lines:
                continue

            name = lines[0]
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
# OPEN CHAT BY SEARCH
# =========================
def open_chat_by_search(driver, chat_name):
    wait = WebDriverWait(driver, 20)
    chat_item_xpath = '//*[contains(@data-tid, "chat-list") or contains(@data-tid, "chat-item") or @role="treeitem" or @role="listitem"]'

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, chat_item_xpath)))
        groups = driver.find_elements(By.XPATH, chat_item_xpath)

        for g in groups:
            lines = [x.strip() for x in g.text.splitlines() if x.strip()]
            if not lines:
                continue

            if lines[0] == "Unread":
                lines.pop(0)

            if not lines:
                continue

            txt = lines[0]
            if not txt:
                txt = g.get_attribute("aria-label") or ""

            if txt == chat_name or chat_name.startswith(txt.replace("...", "").strip()):
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", g
                )
                time.sleep(1)
                g.click()
                time.sleep(5)
                print(f"📂 Đã mở đúng nhóm: {chat_name}")
                return True

        print(f"⚠️ Không thấy {chat_name} ở ngoài, thử dùng thanh Filter bên trái...")

        try:
            filter_icon_xpath = (
                '//button[@data-testid="simple-collab-left-rail-header-sticky-filter-v2-button"]'
                ' | //button[contains(@aria-keyshortcuts, "Ctrl+Shift+F")]'
            )
            filter_icon = driver.find_element(By.XPATH, filter_icon_xpath)
            driver.execute_script("arguments[0].click();", filter_icon)
            time.sleep(1)
        except Exception:
            pass

        left_search_xpath = (
            '//input[@data-testid="simple-collab-left-rail-sticky-filter-input"]'
            ' | //input[@id="simple-collab-left-rail-sticky-filter-input-id"]'
        )

        search = wait.until(
            EC.presence_of_element_located((By.XPATH, left_search_xpath))
        )
        search.click()
        search.send_keys(Keys.CONTROL + "a")
        search.send_keys(Keys.BACKSPACE)
        search.send_keys(chat_name)

        time.sleep(4)

        filtered_result_xpath = f"//span[normalize-space(text())='{chat_name}']"
        filtered_result = wait.until(
            EC.presence_of_element_located((By.XPATH, filtered_result_xpath))
        )
        driver.execute_script("arguments[0].click();", filtered_result)

        time.sleep(5)
        print(f"📂 Đã mở nhóm qua thanh Filter bên trái: {chat_name}")
        return True

    except Exception as e:
        save_screenshot(driver, "open_chat_error.png")
        print(f"❌ Không thể mở nhóm {chat_name}:", e)
        return False


# =========================
# GET MESSAGES
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
                        timestamp = item.find_element(
                            By.TAG_NAME, "time"
                        ).get_attribute("datetime")
                        dt_local = _parse_teams_timestamp(timestamp)
                        if dt_local is not None:
                            last_dt_local = dt_local
                    except Exception:
                        dt_local = last_dt_local

                    if dt_local is None:
                        continue

                    if oldest_dt_in_batch is None or dt_local < oldest_dt_in_batch:
                        oldest_dt_in_batch = dt_local

                    if not (start_time <= dt_local <= end_time):
                        continue

                    date_str = dt_local.strftime("%Y-%m-%d")
                    time_str = dt_local.strftime("%H:%M:%S")

                    content_el = item.find_element(By.CSS_SELECTOR, '[id^="content-"]')
                    raw_html = content_el.get_attribute("innerHTML")

                    # 1. Xóa các thẻ inline
                    text = re.sub(
                        r"</?(span|at|a|strong|b|i|em|u|s|code|font)[^>]*>",
                        "",
                        raw_html,
                        flags=re.IGNORECASE,
                    )

                    # 2. Thay thẻ block và <br> thành \n
                    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
                    text = re.sub(
                        r"</?(div|p|li|ul|ol|tr|table|blockquote|section|article|h[1-6])[^>]*>",
                        "\n",
                        text,
                        flags=re.IGNORECASE,
                    )

                    # 3. Xóa các thẻ còn lại
                    text = re.sub(r"<[^>]+>", "", text)

                    # 4. Decode HTML entities
                    text = html.unescape(text).replace("\xa0", " ")

                    # 5. Dọn dẹp khoảng trắng
                    lines = [line.strip() for line in text.splitlines()]
                    content = "\n".join([line for line in lines if line])

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

            if oldest_dt_in_batch is None:
                break

            if oldest_dt_in_batch <= start_time:
                break

            if prev_oldest_dt is not None and oldest_dt_in_batch >= prev_oldest_dt:
                unchanged_scrolls += 1
                if unchanged_scrolls >= 2:
                    break
            else:
                unchanged_scrolls = 0

            prev_oldest_dt = oldest_dt_in_batch

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
