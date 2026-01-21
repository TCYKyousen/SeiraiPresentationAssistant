import sys
import os
import json
import ctypes
import tempfile
import subprocess
from json import JSONDecodeError

from PySide6.QtWidgets import QApplication, QFileDialog
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineScript, QWebEngineSettings
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot, QUrl, QFile, QIODevice, Qt, QTimer
from PySide6.QtGui import QColor

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

def _apply_window_theme(hwnd, is_dark):
    if sys.platform != "win32" or not hwnd:
        return
    try:
        dwmapi = ctypes.windll.dwmapi
        uxtheme = ctypes.windll.uxtheme
        user32 = ctypes.windll.user32
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

def _apply_chromium_flags():
    flags = [
        "--enable-gpu",
        "--ignore-gpu-blocklist",
        "--enable-zero-copy"
    ]
    current = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "").strip()
    if current:
        merged = current.split()
        for flag in flags:
            if flag not in merged:
                merged.append(flag)
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = " ".join(merged)
    else:
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = " ".join(flags)

class Api(QObject):
    def __init__(self, window=None):
        super().__init__()
        self._window = window
        self.settings = {}
        self.version = {}
        self.dialog_data = {}

    def set_window(self, window):
        self._window = window

    @Slot("QVariant")
    def update_settings(self, settings):
        if not isinstance(settings, dict):
            try:
                settings = settings.toPython()
            except Exception:
                try:
                    settings = json.loads(settings)
                except Exception:
                    settings = {}
        self.settings = settings
        theme_mode = settings.get("Appearance", {}).get("ThemeMode", "Light")
        theme_id = settings.get("Appearance", {}).get("ThemeId", "default")
        if self._window:
            if hasattr(self._window, "update_theme_mode"):
                self._window.update_theme_mode(theme_mode)
            js = f"if (typeof updateTheme === 'function') updateTheme({json.dumps(theme_mode)}, {json.dumps(theme_id)})"
            self._window.page().runJavaScript(js)
            try:
                _apply_window_theme(int(self._window.winId()), _resolve_theme_dark(theme_mode))
            except Exception:
                pass

    @Slot(str)
    def set_title(self, title):
        if self._window:
            self._window.setWindowTitle(str(title))

    @Slot(result="QVariant")
    def get_settings(self):
        return self.settings

    @Slot(result="QVariant")
    def get_version(self):
        return self.version

    @Slot(str, result=str)
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
            return "file:///" + icon_path.replace("\\", "/")
        return ""

    @Slot(result="QVariant")
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
            suffixes = (" (TrueType)", " (OpenType)", " (Type 1)", " (All res)")
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

    @Slot(result="QVariant")
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

    @Slot(result="QVariant")
    def add_quick_launch_app(self):
        file_path = None
        if self._window:
            file_path, _ = QFileDialog.getOpenFileName(
                self._window,
                "Select Application",
                "",
                "Applications (*.exe *.lnk);;Media (*.mp3 *.wav *.mp4 *.mkv *.png *.jpg *.jpeg *.gif);;All Files (*.*)"
            )
        if not file_path:
            return None if self._window else self.get_quick_launch_apps()
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
            apps.append({"name": name, "path": file_path, "icon": ""})
            toolbar["QuickLaunchApps"] = apps
            data["Toolbar"] = toolbar
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.settings = data
            return apps
        except Exception:
            return self.get_quick_launch_apps()

    @Slot(str, str, result="QVariant")
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

    @Slot(str, result="QVariant")
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

    @Slot(str, str, "QVariant")
    def save_setting(self, category, key, value):
        if not isinstance(category, str) or not isinstance(key, str):
            return
        try:
            value = value.toPython()
        except Exception:
            pass
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
                with open(settings_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except JSONDecodeError:
                        data = {}
            if category not in data:
                data[category] = {}
            data[category][key] = value
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.settings = data
            if category == "Appearance" and key in ("ThemeMode", "ThemeId"):
                self.update_settings(data)
        except Exception as e:
            print(f"Error saving settings: {e}", file=sys.stderr)

    @Slot()
    def show_window(self):
        if self._window:
            self._window.show()
            self._window.raise_()
            self._window.activateWindow()

    @Slot(str)
    def open_browser(self, url):
        import webbrowser
        webbrowser.open(url)

    @Slot()
    def trigger_crash(self):
        raise RuntimeError("这是一个手动触发的测试崩溃。")

    @Slot()
    def create_dialog(self):
        msg = "君不见，黄河之水天上来，奔流到海不复回！君不见，高堂明镜悲白发，朝如青丝暮成雪！\n人生得意须尽欢，莫使金樽空对月。\n天生我材必有用，千金散尽还复来。\n烹羊宰牛且为乐，会须一饮三百杯。\n岑夫子，丹丘生。将进酒，君莫停。\n与君歌一曲，请君为我倾耳听。\n钟鼓馔玉不足贵，但愿长醉不复醒。\n古来圣贤皆寂寞，惟有饮者留其名。\n陈王昔时宴平乐，斗酒十千恣欢谑。\n主人何为言少钱？径须沽取对君酌。\n五花马，千金裘。呼儿将出换美酒，与尔同销万古愁。"
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
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(dialog_data, f)
            temp_path = f.name
        subprocess.Popen([sys.executable, __file__, "--dialog", temp_path])

    @Slot(str, str)
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
            "targetLang": font_lang
        }
        temp_settings = json.loads(json.dumps(self.settings))
        if font_name:
            if "Fonts" not in temp_settings:
                temp_settings["Fonts"] = {}
            if "Profiles" not in temp_settings["Fonts"]:
                temp_settings["Fonts"]["Profiles"] = {}
            lang = font_lang or temp_settings.get("General", {}).get("Language", "zh-CN")
            if lang not in temp_settings["Fonts"]["Profiles"]:
                temp_settings["Fonts"]["Profiles"][lang] = {}
            temp_settings["Fonts"]["Profiles"][lang]["web"] = font_name
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(dialog_data, f)
            temp_path = f.name
        if font_name:
            dialog_data["overrideSettings"] = temp_settings
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(dialog_data, f)
        subprocess.Popen([sys.executable, __file__, "--dialog", temp_path])

    @Slot(result="QVariant")
    def get_dialog_data(self):
        return self.dialog_data

    @Slot()
    def on_confirm(self):
        print("DIALOG_CONFIRMED")
        sys.stdout.flush()
        if self._window:
            self._window.close()
        sys.exit(0)

    @Slot()
    def on_cancel(self):
        print("DIALOG_CANCELLED")
        sys.stdout.flush()
        if self._window:
            self._window.close()
        sys.exit(0)

    @Slot(result=str)
    def get_assets_path(self):
        return os.environ.get("ASSETS_PATH", "")

    @Slot(result="QVariant")
    def get_timer_state(self):
        return {
            "remaining": int(os.environ.get("TIMER_REMAINING", 0)),
            "is_running": os.environ.get("TIMER_IS_RUNNING", "false") == "true"
        }

    @Slot(int)
    def start_timer(self, seconds):
        print(f"TIMER_START:{seconds}")
        sys.stdout.flush()

    @Slot()
    def pause_timer(self):
        print("TIMER_PAUSE")
        sys.stdout.flush()

    @Slot()
    def resume_timer(self):
        print("TIMER_RESUME")
        sys.stdout.flush()

    @Slot()
    def stop_timer(self):
        print("TIMER_STOP")
        sys.stdout.flush()

    @Slot()
    def finish_timer(self):
        print("TIMER_FINISH")
        sys.stdout.flush()

    @Slot("QVariant")
    def select_item(self, item):
        print(f"SELECTED_ITEM:{json.dumps(item, ensure_ascii=False)}")
        sys.stdout.flush()
        if self._window:
            self._window.close()
        sys.exit(0)

    @Slot(result="QVariant")
    def get_monet_colors(self):
        try:
            import plugins.monet_utils as monet
            path = monet.get_wallpaper_path()
            if not path:
                return {}
            colors = monet.extract_colors(path)
            return colors or {}
        except Exception as e:
            print(f"Monet error: {e}")
            return {}

    @Slot(result="QVariant")
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

class MainWindow(QWebEngineView):
    def __init__(self, title, url, api, width, height, theme_mode="auto"):
        super().__init__()
        self.setWindowTitle(title)
        self.resize(width, height)
        self._center_on_screen()
        self._theme_mode = theme_mode
        self._apply_page_background()
        settings = self.page().settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.Accelerated2dCanvasEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.WebGLEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        self.api = api
        self.api.set_window(self)
        self.channel = QWebChannel()
        self.channel.registerObject("api", api)
        self.page().setWebChannel(self.channel)
        qwebchannel_js = ""
        f = QFile(":/qtwebchannel/qwebchannel.js")
        if f.open(QIODevice.OpenModeFlag.ReadOnly):
            qwebchannel_js = str(f.readAll(), "utf-8")
            f.close()
        shim_js = """
        new QWebChannel(qt.webChannelTransport, function(channel) {
            window.pywebview = {
                api: new Proxy(channel.objects.api, {
                    get: function(target, prop) {
                        return function(...args) {
                            return new Promise((resolve, reject) => {
                                target[prop](...args, function(result) {
                                    resolve(result);
                                });
                            });
                        }
                    }
                })
            };
            window.dispatchEvent(new Event('pywebviewready'));
        });
        """
        script = QWebEngineScript()
        script.setSourceCode(qwebchannel_js + shim_js)
        script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentCreation)
        script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
        self.page().scripts().insert(script)
        settings_json = json.dumps(api.settings, ensure_ascii=False)
        settings_script = QWebEngineScript()
        settings_script.setSourceCode(f"window.initialSettings = {settings_json};")
        settings_script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentCreation)
        settings_script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
        self.page().scripts().insert(settings_script)
        self.load(QUrl.fromUserInput(url))
        self.loadFinished.connect(lambda *_: self._schedule_backdrop_apply())
        self._schedule_backdrop_apply()

    def _apply_page_background(self):
        is_dark = _resolve_theme_dark(self._theme_mode)
        if is_dark:
            self.page().setBackgroundColor(QColor(24, 24, 24))
        else:
            self.page().setBackgroundColor(QColor(255, 255, 255))

    def _center_on_screen(self):
        screen = QApplication.primaryScreen()
        if not screen:
            return
        geo = screen.availableGeometry()
        x = geo.x() + (geo.width() - self.width()) // 2
        y = geo.y() + (geo.height() - self.height()) // 2
        self.move(x, y)

    def _apply_backdrop(self):
        apply_win11_aesthetics(self, self._theme_mode)
        self.update()

    def _schedule_backdrop_apply(self):
        if sys.platform != "win32":
            return
        for delay in (0, 200, 800):
            QTimer.singleShot(delay, self._apply_backdrop)

    def update_theme_mode(self, theme_mode):
        self._theme_mode = theme_mode
        self._apply_page_background()
        self._schedule_backdrop_apply()

    def showEvent(self, event):
        super().showEvent(event)
        self._schedule_backdrop_apply()

def apply_win11_aesthetics(window, theme_mode=None):
    if sys.platform == "win32":
        try:
            hwnd = int(window.winId())
            dwmapi = ctypes.windll.dwmapi
            corner_preference = ctypes.c_int(2)
            dwmapi.DwmSetWindowAttribute(
                hwnd,
                33,
                ctypes.byref(corner_preference),
                ctypes.sizeof(corner_preference)
            )
            _apply_window_theme(hwnd, _resolve_theme_dark(theme_mode))
            root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(root_dir, "icons", "settings.png")
            if os.path.exists(icon_path):
                hicon = ctypes.windll.user32.LoadImageW(0, icon_path, 1, 0, 0, 0x00000010)
                if hicon:
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, hicon)
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, hicon)
        except Exception:
            pass

def main():
    _apply_chromium_flags()
    app = QApplication(sys.argv)
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    if "--dialog" in sys.argv or "--crash-file" in sys.argv:
        mode = "--dialog" if "--dialog" in sys.argv else "--crash-file"
        try:
            idx = sys.argv.index(mode)
            file_path = sys.argv[idx + 1]
        except ValueError:
            return
        settings_path = os.environ.get("SETTINGS_PATH")
        if not settings_path:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            settings_path = os.path.join(base_dir, "settings.json")
        default_theme = "auto"
        default_accent = "#3275F5"
        settings = {}
        if os.path.exists(settings_path):
            try:
                with open(settings_path, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    default_theme = settings.get("Appearance", {}).get("ThemeMode", "Auto").lower()
                    if default_theme == "dark":
                        default_accent = "#E1EBFF"
                    else:
                        default_accent = "#3275F5"
            except Exception:
                settings = {}
        dialog_data = {}
        try:
            with open(file_path, "r", encoding="utf-8") as f:
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
        except Exception:
            pass
        try:
            os.remove(file_path)
        except Exception:
            pass
        api = Api()
        api.dialog_data = dialog_data
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
        window = MainWindow(
            dialog_data.get("title", "Dialog"),
            html_path,
            api,
            win_width,
            win_height,
            dialog_data.get("theme", default_theme)
        )
        window.show()
    elif len(sys.argv) >= 5:
        url = sys.argv[1]
        title = sys.argv[2]
        width = int(sys.argv[3])
        height = int(sys.argv[4])
        api = Api()
        settings_path = os.environ.get("SETTINGS_PATH")
        if not settings_path:
            if getattr(sys, "frozen", False):
                settings_path = os.path.join(os.path.dirname(sys.executable), "settings.json")
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                settings_path = os.path.join(os.path.dirname(base_dir), "settings.json")
        api.settings = {}
        if settings_path and os.path.exists(settings_path):
            try:
                with open(settings_path, "r", encoding="utf-8") as f:
                    api.settings = json.load(f)
            except Exception:
                pass
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            root_dir = os.path.dirname(base_dir)
            version_path = os.path.join(root_dir, "version.json")
            if os.path.exists(version_path):
                with open(version_path, "r", encoding="utf-8") as f:
                    api.version = json.load(f)
        except Exception:
            api.version = {}
        theme_mode = api.settings.get("Appearance", {}).get("ThemeMode", "Auto")
        window = MainWindow(title, url, api, width, height, theme_mode)
        window.show()
    else:
        return
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
