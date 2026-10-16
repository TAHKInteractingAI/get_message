import os
import sys
import types
import platform
import subprocess
import re

# Khắc phục lỗi thiếu distutils.version.LooseVersion trên Python 3.12+
try:
    from distutils.version import LooseVersion  # type: ignore
except ImportError:
    class LooseVersion:
        def __init__(self, vstring=None):
            self.vstring = vstring
            if vstring:
                self.version = [int(x) if x.isdigit() else x for x in re.split(r'(\d+)', str(vstring)) if x]
            else:
                self.version = []
        def __repr__(self):
            return f"LooseVersion ('{self.vstring}')"
        def __lt__(self, other):
            return self.version < (other.version if isinstance(other, LooseVersion) else LooseVersion(other).version)
        def __le__(self, other):
            return self.version <= (other.version if isinstance(other, LooseVersion) else LooseVersion(other).version)
        def __eq__(self, other):
            return self.version == (other.version if isinstance(other, LooseVersion) else LooseVersion(other).version)
        def __ge__(self, other):
            return self.version >= (other.version if isinstance(other, LooseVersion) else LooseVersion(other).version)
        def __gt__(self, other):
            return self.version > (other.version if isinstance(other, LooseVersion) else LooseVersion(other).version)

    mod_distutils = types.ModuleType("distutils")
    mod_version = types.ModuleType("distutils.version")
    mod_version.LooseVersion = LooseVersion
    mod_distutils.version = mod_version
    sys.modules.setdefault("distutils", mod_distutils)
    sys.modules.setdefault("distutils.version", mod_version)

import undetected_chromedriver as uc
from .config import PROXY_URL


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
# SCREENSHOT
# =========================
def save_screenshot(driver, file_name="error.png"):
    try:
        driver.save_screenshot(file_name)
        print(f"📸 Đã lưu ảnh màn hình: {file_name}")
    except Exception as e:
        print(f"❌ Không thể lưu ảnh màn hình: {e}")


# =========================
# KIỂM TRA VERSION & BINARY BROWSER
# =========================
def get_browser_binary_and_version():
    """Tự động kiểm tra đường dẫn binary và Major Version của trình duyệt."""
    system = platform.system()
    binary_path = None
    major_version = None

    try:
        if system == "Windows":
            import winreg

            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon"
                )
            except FileNotFoundError:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
                )
            version, _ = winreg.QueryValueEx(key, "version")
            major_version = int(version.split(".")[0])

        elif system == "Linux":
            for cmd_name in [
                "google-chrome",
                "google-chrome-stable",
                "chromium",
                "brave-browser",
            ]:
                try:
                    output = subprocess.check_output([cmd_name, "--version"]).decode(
                        "utf-8"
                    )
                    match = re.search(
                        r"(?:Google Chrome|Chromium|Brave Browser)\s+(\d+)\.",
                        output,
                        re.IGNORECASE,
                    )
                    if not match:
                        match = re.search(r"\b(\d{2,3})\.\d+", output)
                    if match:
                        major_version = int(match.group(1))
                        print(
                            f"🔍 Nhận diện {cmd_name}: {output.strip()} -> major: {major_version}"
                        )
                        break
                except Exception:
                    pass

        elif system == "Darwin":  # macOS
            candidates = [
                (
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                    r"Google Chrome (\d+)\.",
                ),
                (
                    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
                    r"Brave Browser (\d+)\.",
                ),
                (
                    "/Applications/Chromium.app/Contents/MacOS/Chromium",
                    r"Chromium (\d+)\.",
                ),
                (
                    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                    r"Microsoft Edge (\d+)\.",
                ),
            ]
            for path, regex in candidates:
                if os.path.exists(path):
                    binary_path = path
                    try:
                        output = subprocess.check_output([path, "--version"]).decode(
                            "utf-8"
                        )
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
# DRIVER INITIALIZATION
# =========================
def get_driver():
    options = uc.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-background-networking")
    options.add_argument("--window-size=1920,1080")
    options.page_load_strategy = "eager"
    options.add_argument("--lang=en-GB")

    # Cho phép tất cả Cookie
    prefs = {
        "profile.cookie_controls_mode": 0,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)

    proxy_url = os.getenv("PROXY_URL") or PROXY_URL
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

    try:
        driver = uc.Chrome(**kwargs)
    except Exception as e:
        print(f"❌ Lỗi khi khởi động undetected_chromedriver: {e}")
        raise e

    try:
        driver.execute_cdp_cmd(
            "Emulation.setTimezoneOverride",
            {"timezoneId": "Asia/Ho_Chi_Minh"},
        )
    except Exception:
        pass

    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    window.navigator.chrome = { runtime: {} };
                    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
                    Object.defineProperty(navigator, 'languages', { get: () => ['en-GB', 'en-US', 'en'] });
                    Object.defineProperty(navigator, 'credentials', {
                        get: () => undefined
                    });

                    window.PublicKeyCredential = undefined;
                """
            },
        )
    except Exception as e:
        print(f"⚠️ Cảnh báo cấu hình CDP script (bỏ qua): {e}")

    return driver
