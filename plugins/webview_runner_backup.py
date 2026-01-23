import sys
import os
import json
import webview
import ctypes
import traceback
import tempfile
import subprocess
import base64
from json import JSONDecodeError

try:
    from webview.platforms.edgechromium import EdgeChrome

    def _safe_clear_user_data(self):
        return

    EdgeChrome.clear_user_data = _safe_clear_user_data
except Exception:
    pass

try:
    from PySide6.QtGui import QGuiApplication, QImage, QPainter, QColor
    from PySide6.QtSvg import QSvgRenderer
    from PySide6.QtCore import QBuffer, QByteArray, QIODevice
except Exception:
    QGuiApplication = None
    QImage = None
    QPainter = None
    QColor = None
    QSvgRenderer = None
    QBuffer = None
    QByteArray = None
    QIODevice = None

# Windows 11 DWM Attributes
DWMWA_WINDOW_CORNER_PREFERENCE = 33
DWMWCP_ROUND = 2
DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
DWMWA_BORDER_COLOR = 34
DWMWA_CAPTION_COLOR = 35
DWMWA_TEXT_COLOR = 36

def _get_windows_dark_mode():
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return int(val) == 0
    except Exception:
        return False

def _resolve_theme_dark(theme_mode):
    mode = str(theme_mode or "").lower()
    if mode == "dark":
        return True
    if mode == "light":
        return False
    if mode == "auto":
        return _get_windows_dark_mode()
    return False

def _try_get_proc(dll, ordinal):
    try:
        kernel32 = ctypes.windll.kernel32
        kernel32.GetProcAddress.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        kernel32.GetProcAddress.restype = ctypes.c_void_p
        addr = kernel32.GetProcAddress(ctypes.c_void_p(dll._handle), ctypes.c_void_p(ordinal))
        if addr:
            return addr
    except Exception:
        return None
    return None

def _set_preferred_app_mode(is_dark):
    if sys.platform != "win32":
        return
    try:
        uxtheme = ctypes.WinDLL("uxtheme")
        addr = _try_get_proc(uxtheme, 135)
        if not addr:
            return
        func = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int)(addr)
        mode = 2 if is_dark else 3
        func(mode)
    except Exception:
        pass

def _allow_dark_for_window(hwnd, is_dark):
    if sys.platform != "win32" or not hwnd:
        return
    try:
        uxtheme = ctypes.WinDLL("uxtheme")
        addr = _try_get_proc(uxtheme, 133)
        if not addr:
            return
        func = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_bool)(addr)
        func(ctypes.c_void_p(hwnd), bool(is_dark))
    except Exception:
        pass

def _apply_window_theme(hwnd, is_dark):
    if sys.platform != "win32" or not hwnd:
        return
    try:
        dwmapi = ctypes.windll.dwmapi
        uxtheme = ctypes.windll.uxtheme
        user32 = ctypes.windll.user32
        _set_preferred_app_mode(is_dark)
        _allow_dark_for_window(hwnd, is_dark)
        val = ctypes.c_int(1 if is_dark else 0)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(val), ctypes.sizeof(val))
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ctypes.byref(val), ctypes.sizeof(val))
        if is_dark:
            border = ctypes.c_int(0x00202020)
            caption = ctypes.c_int(0x00202020)
            text = ctypes.c_int(0x00FFFFFF)
        else:
            border = ctypes.c_int(-1)
            caption = ctypes.c_int(-1)
            text = ctypes.c_int(-1)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR, ctypes.byref(border), ctypes.sizeof(border))
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ctypes.byref(caption), ctypes.sizeof(caption))
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_TEXT_COLOR, ctypes.byref(text), ctypes.sizeof(text))
        theme = "DarkMode_Explorer" if is_dark else "Explorer"
        uxtheme.SetWindowTheme(hwnd, ctypes.c_wchar_p(theme), None)
        flags = 0x0001 | 0x0002 | 0x0004 | 0x0020
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, flags)
    except Exception:
        pass

def apply_win11_aesthetics(window, theme_mode=None):
    if sys.platform == "win32":
        try:
            hwnd = window.native
            dwmapi = ctypes.windll.dwmapi
            # DWMWA_WINDOW_CORNER_PREFERENCE = 33, DWMWCP_ROUND = 2
            corner_preference = ctypes.c_int(2)
            dwmapi.DwmSetWindowAttribute(
                hwnd, 
                33, 
                ctypes.byref(corner_preference), 
                ctypes.sizeof(corner_preference)
            )
            _apply_window_theme(hwnd, _resolve_theme_dark(theme_mode))
            
            # Set Window Icon (WM_SETICON)
            root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(root_dir, "icons", "settings.png")
            if os.path.exists(icon_path):
                # IMAGE_ICON = 1, LR_LOADFROMFILE = 0x00000010
                hicon = ctypes.windll.user32.LoadImageW(0, icon_path, 1, 0, 0, 0x00000010)
                if hicon:
                    # WM_SETICON = 0x0080, ICON_SMALL = 0, ICON_BIG = 1
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, hicon)
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, hicon)
        except Exception as e:
            pass

class Api:
    def __init__(self, window=None):
        self._window = window
        self.settings = {}
        self.version = {}
        self.dialog_data = {}

    def set_window(self, window):
        self._window = window

    def update_settings(self, settings):
        self.settings = settings
        theme_mode = settings.get("Appearance", {}).get("ThemeMode", "Light")
        theme_id = settings.get("Appearance", {}).get("ThemeId", "default")
            
        if self._window:
            self._window.evaluate_js(
                f"if (typeof updateTheme === 'function') updateTheme({json.dumps(theme_mode)}, {json.dumps(theme_id)})"
            )
            try:
                _apply_window_theme(self._window.native, _resolve_theme_dark(theme_mode))
            except Exception:
                pass

    def set_title(self, title):
        if self._window:
            try:
                self._window.set_title(str(title))
            except Exception:
                pass

    def get_settings(self):
        return self.settings

    def get_version(self):
        return self.version

    def get_toolbar_icon(self, icon_name):
        if not icon_name:
            return ""
        icon_map = {
            "select": "Mouse.svg",
            "pen": "Pen.svg",
            "eraser": "Eraser.svg",
            "clear": "Clear.svg",
            "spotlight": "spotlight.svg",
            "timer": "timer.svg",
            "exit": "Minimize.svg"
        }
        key = str(icon_name).lower()
        icon_file = icon_map.get(key)
        if not icon_file:
            return ""
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(root_dir, "icons", icon_file)
        if os.path.exists(icon_path):
            return f"file:///{icon_path.replace('\\', '/')}"
        return ""

    def get_system_fonts(self):
        if sys.platform != "win32":
            return []
        try:
            import winreg

            keys = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            ]
            fonts = set()
            suffixes = (
                " (TrueType)",
                " (OpenType)",
                " (Type 1)",
                " (All res)",
            )
            for root, path in keys:
                try:
                    with winreg.OpenKey(root, path) as k:
                        try:
                            count = winreg.QueryInfoKey(k)[1]
                        except Exception:
                            count = 0
                        for i in range(count):
                            try:
                                name, _, _ = winreg.EnumValue(k, i)
                            except Exception:
                                continue
                            if not isinstance(name, str):
                                continue
                            display = name.strip()
                            for suf in suffixes:
                                if display.endswith(suf):
                                    display = display[: -len(suf)].strip()
                                    break
                            if display:
                                fonts.add(display)
                except OSError:
                    continue
            return sorted(fonts, key=lambda s: s.lower())
        except Exception:
            return []

    def _get_settings_path(self):
        settings_path = os.environ.get("SETTINGS_PATH")
        if not settings_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            settings_path = os.path.join(base_dir, "settings.json")
        return settings_path

    def get_quick_launch_apps(self):
        settings_path = self._get_settings_path()
        data = {}
        try:
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            else:
                data = self.settings or {}
        except Exception:
            data = self.settings or {}

        toolbar = data.get("Toolbar") or {}
        apps = toolbar.get("QuickLaunchApps") or []
        if not isinstance(apps, list):
            apps = []
        self.settings = data
        return apps

    def add_quick_launch_app(self):
        if not self._window:
            return self.get_quick_launch_apps()

        file_path = None
        
        # Try native Windows dialog via ctypes for maximum "directness" and reliability
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes
                
                class OPENFILENAMEW(ctypes.Structure):
                    _fields_ = [
                        ("lStructSize", wintypes.DWORD),
                        ("hwndOwner", wintypes.HWND),
                        ("hInstance", wintypes.HINSTANCE),
                        ("lpstrFilter", wintypes.LPCWSTR),
                        ("lpstrCustomFilter", wintypes.LPWSTR),
                        ("nMaxCustFilter", wintypes.DWORD),
                        ("nFilterIndex", wintypes.DWORD),
                        ("lpstrFile", wintypes.LPWSTR),
                        ("nMaxFile", wintypes.DWORD),
                        ("lpstrFileTitle", wintypes.LPWSTR),
                        ("nMaxFileTitle", wintypes.DWORD),
                        ("lpstrInitialDir", wintypes.LPCWSTR),
                        ("lpstrTitle", wintypes.LPCWSTR),
                        ("Flags", wintypes.DWORD),
                        ("nFileOffset", wintypes.WORD),
                        ("nFileExtension", wintypes.WORD),
                        ("lpstrDefExt", wintypes.LPCWSTR),
                        ("lCustData", wintypes.LPARAM),
                        ("lpfnHook", ctypes.c_void_p),
                        ("lpTemplateName", wintypes.LPCWSTR),
                        ("pvReserved", ctypes.c_void_p),
                        ("dwReserved", wintypes.DWORD),
                        ("FlagsEx", wintypes.DWORD),
                    ]

                ofn = OPENFILENAMEW()
                ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
                
                # Try to get valid HWND from window.native
                hwnd = 0
                try:
                    if isinstance(self._window.native, int):
                        hwnd = self._window.native
                    elif hasattr(self._window.native, 'value'): # some ctypes objects
                        hwnd = self._window.native.value
                    elif hasattr(self._window.native, 'Handle'): # WinForms
                        hwnd = int(self._window.native.Handle)
                except:
                    pass
                ofn.hwndOwner = hwnd
                
                # Filters: must be a null-terminated sequence of null-terminated strings
                filter_str = "Fixed Items\0*.exe;*.lnk;*.mp3;*.wav;*.mp4;*.mkv;*.png;*.jpg;*.jpeg;*.gif\0All Files\0*.*\0\0"
                filter_buf = ctypes.create_unicode_buffer(filter_str)
                ofn.lpstrFilter = ctypes.cast(filter_buf, wintypes.LPCWSTR)
                
                # Buffer for file path
                buf = ctypes.create_unicode_buffer(260)
                ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
                ofn.nMaxFile = 260
                
                # Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY
                ofn.Flags = 0x00080000 | 0x00001000 | 0x00000004
                
                if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
                    file_path = buf.value
            except Exception as e:
                print(f"Direct Windows dialog failed, falling back to pywebview: {e}")

        # Fallback to pywebview dialog if ctypes failed or not on Windows
        if not file_path:
            try:
                # Use simpler filters for pywebview to avoid format errors
                result = self._window.create_file_dialog(
                    webview.OPEN_DIALOG,
                    file_types=(
                        "Items (*.exe;*.lnk;*.mp3;*.wav;*.mp4;*.mkv;*.png;*.jpg;*.jpeg;*.gif)",
                        "All Files (*.*)"
                    ),
                )
                if result:
                    file_path = result[0]
            except Exception as e:
                print(f"Error opening pywebview file dialog: {e}", file=sys.stderr)
                return self.get_quick_launch_apps()

        if not file_path:
            return self.get_quick_launch_apps()

        name = os.path.splitext(os.path.basename(file_path))[0]
        
        settings_path = self._get_settings_path()
        data = {}
        try:
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            toolbar = data.get("Toolbar") or {}
            apps = toolbar.get("QuickLaunchApps") or []
            if not isinstance(apps, list):
                apps = []

            if any(app.get("path") == file_path for app in apps):
                self.settings = data
                return apps

            apps.append(
                {
                    "name": name,
                    "path": file_path,
                    "icon": "",
                }
            )

            toolbar["QuickLaunchApps"] = apps
            data["Toolbar"] = toolbar

            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            self.settings = data
            return apps
        except Exception:
            return self.get_quick_launch_apps()

    def rename_quick_launch_app(self, path, new_name):
        if not path or not new_name:
            return self.get_quick_launch_apps()

        settings_path = self._get_settings_path()
        data = {}
        try:
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            toolbar = data.get("Toolbar") or {}
            apps = toolbar.get("QuickLaunchApps") or []
            if not isinstance(apps, list):
                apps = []

            for app in apps:
                if app.get("path") == path:
                    app["name"] = new_name
                    break

            toolbar["QuickLaunchApps"] = apps
            data["Toolbar"] = toolbar

            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            self.settings = data
            return apps
        except Exception:
            return self.get_quick_launch_apps()

    def remove_quick_launch_app(self, path):
        settings_path = self._get_settings_path()
        data = {}
        try:
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            toolbar = data.get("Toolbar") or {}
            apps = toolbar.get("QuickLaunchApps") or []
            if not isinstance(apps, list):
                apps = []

            apps = [app for app in apps if app.get("path") != path]
            toolbar["QuickLaunchApps"] = apps
            data["Toolbar"] = toolbar

            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            self.settings = data
            return apps
        except Exception:
            return self.get_quick_launch_apps()

    def save_setting(self, category, key, value):
        settings_path = os.environ.get("SETTINGS_PATH")
        if not settings_path:
            if getattr(sys, "frozen", False):
                settings_path = os.path.join(os.path.dirname(sys.executable), "settings.json")
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                settings_path = os.path.join(os.path.dirname(base_dir), "settings.json")
            
        try:
            data = {}
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            
            if category not in data: data[category] = {}
            data[category][key] = value
            
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            
            self.settings = data
            
            if category == "Appearance" and key in ("ThemeMode", "ThemeId"):
                self.update_settings(data)
                
        except Exception as e:
            print(f"Error saving settings: {e}", file=sys.stderr)

    def show_window(self):
        """Bring the window to the foreground."""
        if self._window:
            if sys.platform == "win32":
                try:
                    import ctypes
                    hwnd = self._window.native
                    if hwnd:
                        # SW_RESTORE = 9, SW_SHOW = 5
                        ctypes.windll.user32.ShowWindow(hwnd, 9)
                        ctypes.windll.user32.SetForegroundWindow(hwnd)
                except Exception as e:
                    print(f"Error bringing window to foreground: {e}", file=sys.stderr)
            else:
                self._window.show()

    def open_browser(self, url):
        import webbrowser
        webbrowser.open(url)

    def trigger_crash(self):
        raise RuntimeError("这是一个手动触发的测试崩溃。")

    def create_dialog(self):
        msg = "君不见，黄河之水天上来，奔流到海不复回！君不见，高堂明镜悲白发，朝如青丝暮成雪！\n人生得意须尽欢，莫使金樽空对月。\n天生我材必有用，千金散尽还复来。\n烹羊宰牛且为乐，会须一饮三百杯。\n岑夫子，丹丘生。将进酒，君莫停。\n与君歌一曲，请君为我倾耳听。\n钟鼓馔玉不足贵，但愿长醉不复醒。\n古来圣贤皆寂寞，惟有饮者留其名。\n陈王昔时宴平乐，斗酒十千恣欢谑。\n主人何为言少钱？径须沽取对君酌。\n五花马，千金裘。呼儿将出换美酒，与尔同销万古愁。"
        base_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(os.path.dirname(base_dir), "ppt_assistant", "ui", "dialog.html")
        
        theme_mode = self.settings.get("Appearance", {}).get("ThemeMode", "Light")
        theme_lower = str(theme_mode).lower()
        if theme_lower == "dark":
            accent = "#E1EBFF"
        elif theme_lower == "auto":
            accent = "#3275F5"
        else:
            accent = "#3275F5"

        dialog_data = {
            "code": "test_dialog",
            "title": "",
            "text": msg,
            "confirmText": "",
            "cancelText": "",
            "theme": theme_lower,
            "accentColor": accent
        }
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump(dialog_data, f)
            temp_path = f.name
            
        subprocess.Popen([sys.executable, __file__, "--dialog", temp_path])

    def show_font_warning(self, font_name=None, font_lang=None):
        theme_mode = self.settings.get("Appearance", {}).get("ThemeMode", "Light")
        theme_lower = str(theme_mode).lower()
        if theme_lower == "dark":
            accent = "#E1EBFF"
        else:
            accent = "#3275F5"

        dialog_data = {
            "code": "font_warning",
            "title": "",
            "text": "",
            "confirmText": "",
            "hideCancel": True,
            "theme": theme_lower,
            "accentColor": accent,
            "targetLang": font_lang # Pass the target language for the font
        }
        
        # If font_name is provided, override the web font in settings passed to the dialog
        # This allows the dialog to preview the font before it's necessarily saved/reloaded
        temp_settings = json.loads(json.dumps(self.settings)) # deep copy
        if font_name:
            if "Fonts" not in temp_settings: temp_settings["Fonts"] = {}
            if "Profiles" not in temp_settings["Fonts"]: temp_settings["Fonts"]["Profiles"] = {}
            
            # Use provided font_lang or fallback to UI language
            lang = font_lang or temp_settings.get("General", {}).get("Language", "zh-CN")
            
            if lang not in temp_settings["Fonts"]["Profiles"]: temp_settings["Fonts"]["Profiles"][lang] = {}
            temp_settings["Fonts"]["Profiles"][lang]["web"] = font_name

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump(dialog_data, f)
            temp_path = f.name

        # We need a way to pass the temp_settings to the new process too
        # The easiest way is to include them in the dialog_data if we want to override
        if font_name:
            dialog_data["overrideSettings"] = temp_settings
            # Rewrite the temp file with overrideSettings
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(dialog_data, f)

        subprocess.Popen([sys.executable, __file__, "--dialog", temp_path])

    def get_dialog_data(self):
        return self.dialog_data

    def on_confirm(self):
        print("DIALOG_CONFIRMED")
        sys.stdout.flush()
        if self._window:
            self._window.destroy()
        sys.exit(0)

    def on_cancel(self):
        print("DIALOG_CANCELLED")
        sys.stdout.flush()
        if self._window:
            self._window.destroy()
        sys.exit(0)

    def get_assets_path(self):
        return os.environ.get("ASSETS_PATH", "")

    def get_timer_state(self):
        return {
            "remaining": int(os.environ.get("TIMER_REMAINING", 0)),
            "is_running": os.environ.get("TIMER_IS_RUNNING", "false") == "true"
        }

    def start_timer(self, seconds):
        print(f"TIMER_START:{seconds}")
        sys.stdout.flush()

    def pause_timer(self):
        print("TIMER_PAUSE")
        sys.stdout.flush()

    def resume_timer(self):
        print("TIMER_RESUME")
        sys.stdout.flush()

    def stop_timer(self):
        print("TIMER_STOP")
        sys.stdout.flush()

    def finish_timer(self):
        print("TIMER_FINISH")
        sys.stdout.flush()

    def select_item(self, item):
        # Generic method to return data to the caller process
        # item can be any serializable object
        print(f"SELECTED_ITEM:{json.dumps(item, ensure_ascii=False)}")
        sys.stdout.flush()
        if self._window:
            self._window.destroy()
        sys.exit(0)

    def get_screen_list(self):
        screens = []
        if sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes

                user32 = ctypes.windll.user32

                class MONITORINFOEXW(ctypes.Structure):
                    _fields_ = [
                        ("cbSize", wintypes.DWORD),
                        ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT),
                        ("dwFlags", wintypes.DWORD),
                        ("szDevice", wintypes.WCHAR * 32)
                    ]

                class DISPLAY_DEVICEW(ctypes.Structure):
                    _fields_ = [
                        ("cb", wintypes.DWORD),
                        ("DeviceName", wintypes.WCHAR * 32),
                        ("DeviceString", wintypes.WCHAR * 128),
                        ("StateFlags", wintypes.DWORD),
                        ("DeviceID", wintypes.WCHAR * 128),
                        ("DeviceKey", wintypes.WCHAR * 128)
                    ]

                monitor_infos = []

                def monitor_enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
                    mi = MONITORINFOEXW()
                    mi.cbSize = ctypes.sizeof(MONITORINFOEXW)
                    if user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi)):
                        monitor_infos.append(mi)
                    return True

                MONITORENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(wintypes.RECT), ctypes.c_double)
                user32.EnumDisplayMonitors(0, 0, MONITORENUMPROC(monitor_enum_proc), 0)

                for i, mi in enumerate(monitor_infos):
                    name = f"Display {i+1}"
                    
                    # Try to get the monitor name
                    dd = DISPLAY_DEVICEW()
                    dd.cb = ctypes.sizeof(DISPLAY_DEVICEW)
                    
                    if user32.EnumDisplayDevicesW(mi.szDevice, 0, ctypes.byref(dd), 0):
                        name = dd.DeviceString
                        
                    screens.append({
                        "id": i,
                        "name": name,
                        "is_primary": (mi.dwFlags & 1) != 0
                    })
            except Exception as e:
                print(f"Error getting screens: {e}", file=sys.stderr)
        
        return screens

def main():
    # Optimization: Only import what's needed for the specific mode
    if "--dialog" in sys.argv or "--crash-file" in sys.argv:
        mode = "--dialog" if "--dialog" in sys.argv else "--crash-file"
        idx = sys.argv.index(mode)
        file_path = sys.argv[idx + 1]
        
        # Load theme settings first for dialog mode
        settings_path = os.environ.get("SETTINGS_PATH")
        if not settings_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            settings_path = os.path.join(base_dir, "settings.json")
        
        default_theme = "auto"
        default_accent = "#3275F5"
        settings = {}
        
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    default_theme = settings.get("Appearance", {}).get("ThemeMode", "Auto").lower()
                    if default_theme == "dark":
                        default_accent = "#E1EBFF"
                    else:
                        default_accent = "#3275F5"
            except:
                settings = {}

        with open(file_path, 'r', encoding='utf-8') as f:
            if mode == "--dialog":
                dialog_data = json.load(f)
                if "theme" not in dialog_data or dialog_data["theme"] == "auto":
                    dialog_data["theme"] = default_theme
                if "accentColor" not in dialog_data:
                    dialog_data["accentColor"] = default_accent
            else:
                error_msg = f.read()
                dialog_data = {
                    "title": "程序崩溃了 (´；ω；`) ",
                    "text": error_msg,
                    "isError": True,
                    "confirmText": "关闭",
                    "cancelText": "复制错误",
                    "theme": default_theme,
                    "accentColor": default_accent
                }
        
        try: os.remove(file_path)
        except: pass

        api = Api()
        api.dialog_data = dialog_data
        
        # Use overridden settings if provided in dialog_data
        if "overrideSettings" in dialog_data:
            api.settings = dialog_data["overrideSettings"]
        else:
            api.settings = settings
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(os.path.dirname(base_dir), "ppt_assistant", "ui", "dialog.html")
        
        if mode == "--crash-file":
            win_width = 900
            win_height = 600
        else:
            win_width = 650
            win_height = 500
        
        window = webview.create_window(
            dialog_data.get("title", "Dialog"),
            html_path,
            js_api=api,
            width=win_width,
            height=win_height,
            frameless=False,
            on_top=True,
            resizable=True
        )
        api.set_window(window)
        storage_path = os.path.join(os.getenv('APPDATA', os.path.expanduser('~')), 'KazuhaRemake', 'WebView')
        if not os.path.exists(storage_path):
            try: os.makedirs(storage_path)
            except: pass
        webview.start(apply_win11_aesthetics, (window, dialog_data.get("theme", default_theme)), gui='edgechromium', storage_path=storage_path)
        return

    if len(sys.argv) < 5:
        return

    url = sys.argv[1]
    title = sys.argv[2]
    width = int(sys.argv[3])
    height = int(sys.argv[4])
    transparent = len(sys.argv) > 5 and sys.argv[5].lower() == "true"

    api = Api()
    
    # Pre-load settings
    settings_path = os.environ.get("SETTINGS_PATH")
    if not settings_path:
        # 兼容逻辑：如果环境变量没传，尝试从程序同级读取
        if getattr(sys, "frozen", False):
            settings_path = os.path.join(os.path.dirname(sys.executable), "settings.json")
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            settings_path = os.path.join(os.path.dirname(base_dir), "settings.json")

    api.settings = {}
    if settings_path and os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                api.settings = json.load(f)
        except:
            pass

    # Load version info
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        root_dir = os.path.dirname(base_dir)
        version_path = os.path.join(root_dir, "version.json")
        if os.path.exists(version_path):
            with open(version_path, "r", encoding="utf-8") as f:
                api.version = json.load(f)
    except:
        api.version = {}

    # Inject settings into HTML for "immediate load"
    final_url = url
    temp_html_path = None
    if os.path.exists(url) and url.endswith(".html"):
        try:
            with open(url, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Inject as a script before other scripts
            settings_json = json.dumps(api.settings, ensure_ascii=False)
            injection = f"\n<script>window.initialSettings = {settings_json};</script>\n"
            
            if "</head>" in content:
                content = content.replace("</head>", injection + "</head>")
            else:
                content = injection + content
                
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(content)
                temp_html_path = f.name
                final_url = temp_html_path
        except Exception as e:
            print(f"Error injecting settings: {e}")

    window = webview.create_window(
        title, 
        final_url, 
        js_api=api,
        width=width, 
        height=height,
        frameless=False
    )
    api.set_window(window)
    
    # Clean up temp html after start
    def on_loaded():
        if temp_html_path and os.path.exists(temp_html_path):
            try: os.remove(temp_html_path)
            except: pass
    
    # Optimization: Use edgechromium directly for faster startup on Windows
    storage_path = os.path.join(os.getenv('APPDATA', os.path.expanduser('~')), 'KazuhaRemake', 'WebView')
    if not os.path.exists(storage_path):
        try: os.makedirs(storage_path)
        except: pass
    
    theme_mode = api.settings.get("Appearance", {}).get("ThemeMode", "Auto")
    webview.start(apply_win11_aesthetics, (window, theme_mode), gui='edgechromium', debug=False, storage_path=storage_path)

if __name__ == "__main__":
    main()
