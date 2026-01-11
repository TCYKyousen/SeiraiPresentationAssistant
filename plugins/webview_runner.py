import sys
import os
import json
import webview
import ctypes
import traceback
import tempfile
import subprocess
from json import JSONDecodeError

try:
    from webview.platforms.edgechromium import EdgeChrome

    def _safe_clear_user_data(self):
        return

    EdgeChrome.clear_user_data = _safe_clear_user_data
except Exception:
    pass

# Windows 11 DWM Attributes
DWMWA_WINDOW_CORNER_PREFERENCE = 33
DWMWCP_ROUND = 2

def apply_win11_aesthetics(window):
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
            
        if self._window:
            self._window.evaluate_js(f"if (typeof updateTheme === 'function') updateTheme('{theme_mode}')")

    def get_settings(self):
        return self.settings

    def get_version(self):
        return self.version

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

        try:
            result = self._window.create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=["Executable Files (*.exe)", "All Files (*.*)"],
            )
        except Exception:
            return self.get_quick_launch_apps()

        if not result:
            return self.get_quick_launch_apps()

        file_path = result[0]
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
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            settings_path = os.path.join(base_dir, "settings.json")
            
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
            
            if category == "Appearance" and key == "ThemeMode":
                self.update_settings(data)
                
        except Exception as e:
            print(f"Error saving settings: {e}", file=sys.stderr)

    def trigger_crash(self):
        raise RuntimeError("这是一个手动触发的测试崩溃。")

    def create_dialog(self):
        msg = "曾许下心愿 等待你的出现\n褪色的秋千 有本书会纪念\n我循着时间 捡起梦的照片\n童话还没有兑现 却需要说再见"
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
        
        if os.path.exists(settings_path):
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    default_theme = settings.get("Appearance", {}).get("ThemeMode", "Auto").lower()
                    if default_theme == "dark":
                        default_accent = "#E1EBFF"
                    else:
                        default_accent = "#3275F5"
            except: pass

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
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(os.path.dirname(base_dir), "ppt_assistant", "ui", "dialog.html")
        
        if mode == "--crash-file":
            win_width = 900
            win_height = 600
        else:
            win_width = 500
            win_height = 400
        
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
        webview.start(apply_win11_aesthetics, (window,), gui='edgechromium', storage_path=storage_path)
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
    
    webview.start(apply_win11_aesthetics, (window,), gui='edgechromium', debug=False, storage_path=storage_path)

if __name__ == "__main__":
    main()
