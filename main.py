import sys
import os
import traceback
import tempfile
import subprocess
import json
import importlib
import importlib.util
import time

# Delay heavy imports or move them inside if __name__ == "__main__" logic
# to allow --webview-runner to start fast and clean.

if __name__ == "__main__":
    if "--webview-runner" in sys.argv:
        # Minimal imports for webview runner
        idx = sys.argv.index("--webview-runner")
        import plugins.webview_runner as _wv
        # Adjust sys.argv so argparse in webview_runner (if any) sees clean args
        sys.argv = ["webview_runner.py"] + sys.argv[idx + 1 :]
        _wv.main()
        sys.exit(0)

from PySide6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, QFrame, QGraphicsDropShadowEffect, QProgressBar
from PySide6.QtCore import Qt, QTimer, Slot, QSize, QPoint
from PySide6.QtGui import QFontDatabase, QFont, QColor, QIcon, QRegion, QPainter, QPen, QBrush

from ppt_assistant.core.ppt_monitor import PPTMonitor
from ppt_assistant.ui.overlay import OverlayWindow
from plugins.builtins.settings.plugin import SettingsPlugin
from plugins.builtins.timer.plugin import TimerPlugin
from ppt_assistant.ui.tray import SystemTray
from ppt_assistant.core.config import cfg, SETTINGS_PATH, PLUGINS_DIR, reload_cfg, _apply_theme_and_color, Theme, qconfig
from ppt_assistant.core.timer_manager import TimerManager
from ppt_assistant.core.i18n import t


SPLASH_I18N = {
    "zh-CN": {
        "initializing": "正在初始化",
        "loading_config": "加载配置",
        "loading_fonts": "加载字体",
        "init_monitor": "启动监视器",
        "init_ui": "创建界面",
        "loading_plugins": "加载插件",
        "loading_settings": "加载设置",
        "loading_timer": "加载计时器",
        "init_tray": "创建托盘图标",
        "finalizing": "完成初始化",
        "watermark.1": "开发中版本",
        "watermark.2": "技术预览版",
        "watermark.3": "Release Preview",
        "watermark.4": "重新评估版本",
        "dev_watermark": "{type}\n不保证最终品质 （{version}）"
    },
    "zh-TW": {
        "initializing": "正在初始化",
        "loading_config": "載入設定",
        "loading_fonts": "載入字型",
        "init_monitor": "啟動監視器",
        "init_ui": "建立介面",
        "loading_plugins": "載入插件",
        "loading_settings": "載入設定",
        "loading_timer": "載入計時器",
        "init_tray": "建立系統匣圖示",
        "finalizing": "完成初始化",
        "watermark.1": "開發中版本",
        "watermark.2": "技術預覽版",
        "watermark.3": "Release Preview",
        "watermark.4": "重新評估版本",
        "dev_watermark": "{type}\n不保證最終品質 （{version}）"
    },
    "ja-JP": {
        "initializing": "初期化中",
        "loading_config": "設定を読み込み中",
        "loading_fonts": "フォントを読み込み中",
        "init_monitor": "モニターを起動中",
        "init_ui": "UIを作成中",
        "loading_plugins": "プラグインを読み込み中",
        "loading_settings": "設定を読み込み中",
        "loading_timer": "タイマーを読み込み中",
        "init_tray": "トレイアイコンを作成中",
        "finalizing": "初期化完了",
        "watermark.1": "開発中バージョン",
        "watermark.2": "テクニカルプレビュー",
        "watermark.3": "Release Preview",
        "watermark.4": "再評価バージョン",
        "dev_watermark": "{type}\n品質は保証されません （{version}）"
    },
    "en-US": {
        "initializing": "Initializing",
        "loading_config": "Loading config",
        "loading_fonts": "Loading fonts",
        "init_monitor": "Starting monitor",
        "init_ui": "Creating UI",
        "loading_plugins": "Loading plugins",
        "loading_settings": "Loading settings",
        "loading_timer": "Loading timer",
        "init_tray": "Creating system tray",
        "finalizing": "Finalizing",
        "watermark.1": "In-Development",
        "watermark.2": "Technical Preview",
        "watermark.3": "Release Preview",
        "watermark.4": "Re-evaluated Version",
        "dev_watermark": "{type}\nFinal quality not guaranteed ({version})"
    }
}


def _get_current_language():
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("General", {}).get("Language", "zh-CN")
    except Exception:
        pass
    return "zh-CN"


def _apply_global_font(app: QApplication):
    root_dir = os.path.dirname(os.path.abspath(__file__))
    font_path = os.path.join(root_dir, "fonts", "MiSansVF.ttf")
    selected_family = ""
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            lang = data.get("General", {}).get("Language", "zh-CN")
            profiles = (data.get("Fonts", {}) or {}).get("Profiles", {}) or {}
            v = (profiles.get(lang, {}) or {}).get("qt", "")
            if isinstance(v, str) and v.strip():
                selected_family = v.strip()
    except Exception:
        selected_family = ""

    base_family = ""
    if os.path.exists(font_path):
        try:
            font_id = QFontDatabase.addApplicationFont(font_path)
            if font_id != -1:
                families = QFontDatabase.applicationFontFamilies(font_id)
                if families:
                    base_family = families[0]
        except Exception:
            base_family = ""

    family = selected_family or base_family
    if not family:
        return
    app.setFont(QFont(family))


def _load_version_info():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    version_path = os.path.join(root_dir, "version.json")
    version = ""
    code_name = ""
    code_name_cn = ""
    if os.path.exists(version_path):
        try:
            with open(version_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            version = data.get("version", "")
            raw_code_name = data.get("code_name", "")
            code_name_cn = data.get("code_name_CN", "")
            mapping = {
                "MomokaKawaragi": "Momoka Kawaragi"
            }
            code_name = mapping.get(raw_code_name, raw_code_name)
        except Exception:
            pass
    return version, code_name, code_name_cn


def _is_dev_preview_version(version: str) -> bool:
    if not version:
        return False
    parts = str(version).strip().split(".")
    if len(parts) < 2:
        return False
    # 除了 .0 是正式版，其他后缀 (.1, .2, .3, .4) 都带水印
    suffix = parts[-1]
    return suffix in ["1", "2", "3", "4"]


class StartupSplash(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self._container = QFrame(self)
        self._container.setObjectName("splashContainer")

        self._version_text, self._code_name_en, self._code_name_cn = _load_version_info()
        self._language = _get_current_language()
        
        # 确定主题
        theme_val = cfg.themeMode.value
        if theme_val == Theme.AUTO:
            from qfluentwidgets import isDarkTheme
            self._is_dark = isDarkTheme()
        else:
            self._is_dark = theme_val == Theme.DARK

        self._build_ui()
        self._apply_styles()
        self._center_on_screen()

        self.set_progress(0, "initializing")

        if _is_dev_preview_version(self._version_text):
            self._dev_watermark = QLabel(self._container)
            i18n_table = SPLASH_I18N.get(self._language, SPLASH_I18N["zh-CN"])
            suffix = self._version_text.split(".")[-1]
            w_type = i18n_table.get(f"watermark.{suffix}", "")
            tmpl = i18n_table.get("dev_watermark", "")
            self._dev_watermark.setText(tmpl.format(type=w_type, version=self._version_text))
            font = QFont()
            font.setPixelSize(11)
            self._dev_watermark.setFont(font)
            self._dev_watermark.setAlignment(Qt.AlignRight | Qt.AlignBottom)
            
            watermark_color = "rgba(255, 255, 255, 100)" if self._is_dark else "rgba(0, 0, 0, 100)"
            self._dev_watermark.setStyleSheet(f"color: {watermark_color};")
            
            self._dev_watermark.resize(320, 36)
            self._dev_watermark.move(self._container.width() - self._dev_watermark.width() - 16,
                                     self._container.height() - self._dev_watermark.height() - 12)

    def _build_ui(self):
        # Redesigned based on QML spec
        # Width: 678, Height: 255
        self._container.setFixedSize(678, 255)
        self._container.move(24, 16) # Offset for shadow
        
        # Logo (kZHTXT_2.png equivalent) - x: 38, y: 37
        self._icon_label = QLabel(self._container)
        self._icon_label.setFixedSize(64, 64) 
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "logo.svg")
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            pix = icon.pixmap(64, 64)
            self._icon_label.setPixmap(pix)
        self._icon_label.move(38, 37)

        # Brand Name (kazuha) - x: 38, y: 111
        self._brand_label = QLabel("Kazuha", self._container)
        brand_font = QFont("HarmonyOS Sans SC")
        brand_font.setPixelSize(32)
        brand_font.setWeight(QFont.Black) 
        self._brand_label.setFont(brand_font)
        self._brand_label.setFixedWidth(328)
        self._brand_label.setFixedHeight(32) 
        self._brand_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self._brand_label.move(38, 111)

        # Version Info Container - x: 38, y: 143
        self._version_info_label = QLabel(self._container)
        
        ver_text = self._version_text or ""
        en_text = self._code_name_en or ""
        
        ver_color = "#FFFFFF" if self._is_dark else "#000000"
        en_color = "rgba(255, 255, 255, 0.47)" if self._is_dark else "rgba(0, 0, 0, 0.47)"
        
        html = f"""
        <div style="line-height: 20px;">
            <span style="font-family: 'MiSans'; font-size: 11px; font-weight: 500; color: {ver_color};">{ver_text}</span>
            <span style="font-family: 'MiSans'; font-size: 11px; font-weight: 300; color: {en_color}; margin-left: 2px;">{en_text}</span>
        </div>
        """
        
        self._version_info_label.setText(html)
        self._version_info_label.adjustSize()
        self._version_info_label.move(38, 143) 

        # Loading Spinner (Vector) - x: 38, y: 199
        spinner_color = QColor("#d9d9d9") if self._is_dark else QColor("#666666")
        self._spinner = IndeterminateSpinner(self._container, color=spinner_color)
        self._spinner.move(38, 199)
        self._spinner.start()

        # Status Text (element_2) - x: 76, y: 203
        init_text = SPLASH_I18N.get(self._language, SPLASH_I18N["zh-CN"])["initializing"]
        self._percent_label = QLabel(f"{init_text} 0%", self._container)
        percent_font = QFont("HarmonyOS Sans SC")
        percent_font.setPixelSize(15)
        percent_font.setBold(True)
        self._percent_label.setFont(percent_font)
        self._percent_label.setFixedWidth(221)
        self._percent_label.setFixedHeight(24)
        self._percent_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self._percent_label.move(76, 200)

        # Progress Bar Foreground (rectangle_31) - y: 247, h: 8, w: 678
        self._progress = QProgressBar(self._container)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        self._progress.setFixedHeight(8)
        self._progress.setFixedWidth(678) 
        self._progress.move(0, 247)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 60))
        self._container.setGraphicsEffect(shadow)

    def _apply_styles(self):
        if self._is_dark:
            bg_color = "rgba(47, 47, 47, 240)"
            border_color = "rgba(255, 255, 255, 0.15)"
            brand_color = "#ffffff"
            percent_color = "#d9d9d9"
            progress_bg = "#454545" 
            chunk_color = "#E1EBFF" 
        else:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.08)"
            brand_color = "#000000"
            percent_color = "#666666"
            progress_bg = "#e5e5e5"
            chunk_color = "#3275F5"

        self.resize(678 + 48, 255 + 48) # Increased for shadow
        
        self._container.setStyleSheet(
            f"QFrame#splashContainer {{"
            f"background-color: {bg_color};"
            f"border: 1px solid {border_color};"
            "border-radius: 8px;"
            "}"
        )
        self._brand_label.setStyleSheet(f"color: {brand_color};")
        self._percent_label.setStyleSheet(f"color: {percent_color};")
        
        self._progress.setStyleSheet(
            "QProgressBar {"
            f"background-color: {progress_bg};"
            "border: none;"
            "border-bottom-left-radius: 8px;"
            "border-bottom-right-radius: 8px;"
            "}"
            "QProgressBar::chunk {"
            f"background-color: {chunk_color};"
            "border-bottom-left-radius: 8px;"
            "border-bottom-right-radius: 8px;"
            "}"
        )

    def _center_on_screen(self):
        screen = QApplication.primaryScreen()
        if not screen:
            return
        screen_geo = screen.geometry()
        w = self.width()
        h = self.height()
        x = screen_geo.x() + (screen_geo.width() - w) // 2
        y = screen_geo.y() + (screen_geo.height() - h) // 2
        self.move(x, y)

    def set_progress(self, value, text_key="initializing"):
        value = min(max(value, 0), 100)
        self._progress.setValue(value)
        
        display_text = SPLASH_I18N.get(self._language, SPLASH_I18N["zh-CN"]).get(text_key, text_key)
        self._percent_label.setText(f"{display_text} {value}%")
        
        # Update spinner if needed, or it spins automatically
        QApplication.processEvents()

    def finish(self):
        self._progress.setValue(100)
        self._percent_label.setText("初始化完成 100%")
        if hasattr(self, '_spinner'):
            self._spinner.stop()
        QTimer.singleShot(250, self.close)

class IndeterminateSpinner(QWidget):
    def __init__(self, parent=None, color=QColor("#d9d9d9")):
        super().__init__(parent)
        self.setFixedSize(27, 27)
        self._angle = 0
        self._color = color
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._rotate)
        self._timer.start(16) # ~60 FPS

    def start(self):
        if not self._timer.isActive():
            self._timer.start()

    def stop(self):
        self._timer.stop()

    def _rotate(self):
        self._angle = (self._angle + 5) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        rect = self.rect()
        cx, cy = rect.center().x(), rect.center().y()
        
        # Outer ring (static)
        pen = QPen(self._color)
        pen.setWidth(3)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        
        # Draw ring. Adjust rect to account for pen width
        ring_radius = 10 
        painter.drawEllipse(QPoint(cx, cy), ring_radius, ring_radius)
        
        # Inner rotating dot
        painter.save()
        painter.translate(cx, cy)
        painter.rotate(self._angle)
        
        # Draw small circle on the orbit
        dot_radius = 2.5
        orbit_radius = ring_radius - 3 - dot_radius + 1 # Fine tuned visual position
        
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(self._color))
        painter.drawEllipse(QPoint(0, -orbit_radius), dot_radius, dot_radius)
        
        painter.restore()
        painter.end()

def show_webview_dialog(title, text, confirm_text="确认", cancel_text="取消", is_error=False, hide_cancel=False, code=None):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    theme = "auto"
    accent = "#3275F5"
    try:
        from ppt_assistant.core.config import cfg, Theme, qconfig
        theme = cfg.themeMode.value.lower() if hasattr(cfg.themeMode, "value") else "auto"
        resolved_theme = theme
        if theme == "auto":
            try:
                if isinstance(qconfig.theme, Theme):
                    resolved_theme = "dark" if qconfig.theme == Theme.DARK else "light"
            except:
                resolved_theme = "light"
        accent = "#E1EBFF" if resolved_theme == "dark" else "#3275F5"
    except:
        pass

    dialog_data = {
        "title": title,
        "text": text,
        "confirmText": confirm_text,
        "cancelText": cancel_text,
        "isError": is_error,
        "hideCancel": hide_cancel,
        "theme": theme,
        "accentColor": accent
    }
    if code is not None:
        dialog_data["code"] = code
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as f:
        json.dump(dialog_data, f)
        temp_path = f.name
        
    root_dir = base_dir
    main_path = os.path.join(root_dir, "main.py")
    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--webview-runner", "--dialog", temp_path]
    else:
        cmd = [sys.executable, main_path, "--webview-runner", "--dialog", temp_path]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, text=True)
    return proc

class CrashHandler:
    def __init__(self, app=None):
        self.app = app
        self.app_instance = None
        self._handling = False
        sys.excepthook = self.handle_exception
        import threading
        threading.excepthook = self.handle_thread_exception

    def set_app_instance(self, instance):
        self.app_instance = instance

    def handle_thread_exception(self, args):
        self.handle_exception(args.exc_type, args.exc_value, args.exc_traceback)

    def handle_exception(self, exc_type, exc_value, exc_traceback):
        if self._handling:
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        self._handling = True
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(f"CRASH DETECTED:\n{error_msg}", file=sys.stderr)
        
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            root_dir = base_dir
            main_path = os.path.join(root_dir, "main.py")
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.log', delete=False, encoding='utf-8') as f:
                f.write(error_msg)
                temp_path = f.name
            
            creationflags = 0x00000008 # DETACHED_PROCESS
            if getattr(sys, "frozen", False):
                cmd = [sys.executable, "--webview-runner", "--crash-file", temp_path]
            else:
                cmd = [sys.executable, main_path, "--webview-runner", "--crash-file", temp_path]
            subprocess.Popen(cmd, creationflags=creationflags, close_fds=True)
        except Exception as e:
            print(f"Failed to launch crash dialog: {e}", file=sys.stderr)
        
        try:
            if self.app_instance is not None:
                self.app_instance.cleanup()
        except Exception as e:
            print(f"Error during crash cleanup: {e}", file=sys.stderr)
        
        import time
        time.sleep(0.5)
        os._exit(1)

def _handle_multi_instance(app: QApplication):
    try:
        import psutil
    except ImportError:
        return

    current_pid = os.getpid()
    main_path = os.path.abspath(__file__)
    pids = []
    for p in psutil.process_iter(["pid", "cmdline"]):
        if p.info.get("pid") == current_pid: continue
        cmd = p.info.get("cmdline") or []
        for part in cmd:
            try:
                if os.path.abspath(part) == main_path or os.path.basename(part).lower() == "main.py":
                    pids.append(p.info.get("pid"))
                    break
            except: continue
    
    if not pids: return
    
    proc = show_webview_dialog(
        title="",
        text="",
        confirm_text="",
        cancel_text="",
        is_error=False,
        hide_cancel=False,
        code="multi_instance"
    )
    
    # Wait for the process to exit and check stdout for result
    stdout, _ = proc.communicate()
    if "DIALOG_CONFIRMED" in stdout:
        for pid in pids:
            try: psutil.Process(pid).terminate()
            except: pass
        return
    
    for pid in pids:
        try: psutil.Process(pid).terminate()
        except: pass
    app.quit()
    sys.exit(0)


def _t(key):
    return key # Simple fallback if i18n is missing

class PPTAssistantApp:
    def __init__(self, app: QApplication, splash=None):
        self.app = app
        self.app.setQuitOnLastWindowClosed(False)
        self._splash = splash
        self._timer_manager = TimerManager()
        self._last_timer_notify_at = 0.0
        self._reloading_overlay = False
        self._reload_timer = QTimer()
        self._reload_timer.setSingleShot(True)
        self._reload_timer.setInterval(150)
        self._reload_timer.timeout.connect(self._reload_overlay)
        
        # Start async initialization
        self._init_gen = self._init_steps()
        QTimer.singleShot(0, self._perform_init_step)

    def _init_steps(self):
        # Step 1: Basic Config
        yield 10, "loading_config"
        _apply_theme_and_color(cfg.themeMode.value)
        
        # Step 2: Fonts
        yield 20, "loading_fonts"
        _apply_global_font(self.app)
        self._current_language = _get_current_language()
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
        profiles = (data.get("Fonts", {}) or {}).get("Profiles", {}) or {}
        lang_profile = profiles.get(self._current_language, {}) or {}
        qt_font = lang_profile.get("qt", "")
        overlay_font = lang_profile.get("overlay", "") or qt_font
        self._current_qt_font = qt_font.strip() if isinstance(qt_font, str) else ""
        self._current_overlay_font = overlay_font.strip() if isinstance(overlay_font, str) else ""

        self._settings_mtime = os.path.getmtime(SETTINGS_PATH) if os.path.exists(SETTINGS_PATH) else 0
        self._settings_timer = QTimer()
        self._settings_timer.setInterval(100)
        self._settings_timer.timeout.connect(self._check_settings_changed)
        self._settings_timer.start()

        self.app.aboutToQuit.connect(self.cleanup)

        # Step 3: Monitor (Non-UI logic)
        yield 30, "init_monitor"
        self.monitor = PPTMonitor()
        
        # Step 4: Overlay (UI creation - expensive)
        yield 40, "init_ui"
        # Yield to event loop BEFORE creating heavy UI to prevent freeze
        # We can split Overlay creation if needed, but yielding before is key
        pass 
        
        self.overlay = OverlayWindow()
        
        # Step 5: Plugins (IO/Process - expensive)
        yield 60, "loading_plugins"
        self._load_plugins()
        
        # Step 6: Tray (UI)
        yield 80, "init_tray"
        self.tray = SystemTray()
        
        # Step 7: Finalize connections
        yield 85, "finalizing"
        self.overlay.set_monitor(self.monitor)

        yield 90, "finalizing"
        self._connect_signals()

        yield 95, "finalizing"
        self.monitor.start_monitoring()

        if self._splash is not None:
            self._splash.finish()

    def _perform_init_step(self):
        try:
            progress, text = next(self._init_gen)
            self.update_splash(progress, text)
            # Schedule next step immediately but allow event loop to breathe
            QTimer.singleShot(0, self._perform_init_step)
        except StopIteration:
            pass # Done
        except Exception as e:
            print(f"Initialization error: {e}")
            sys.exit(1)

    def _load_plugins(self):
        """Dynamic plugin loading from builtins and external directory."""
        self.plugins = []
        
        # 1. Load Builtin Plugins
        builtin_plugins = [
            "plugins.builtins.settings.plugin.SettingsPlugin",
            "plugins.builtins.timer.plugin.TimerPlugin",
            "plugins.builtins.spotlight.plugin.SpotlightPlugin",
            "plugins.builtins.app_launcher.plugin.AppLauncherPlugin"
        ]
        
        for p_path in builtin_plugins:
            try:
                mod_name, cls_name = p_path.rsplit(".", 1)
                mod = importlib.import_module(mod_name)
                cls = getattr(mod, cls_name)
                plugin = cls()
                plugin.set_context(self)
                self.plugins.append(plugin)
                
                # Maintain compatibility with existing code
                if cls_name == "SettingsPlugin":
                    self.settings_plugin = plugin
                elif cls_name == "TimerPlugin":
                    self.timer_plugin = plugin
            except Exception as e:
                print(f"Failed to load builtin plugin {p_path}: {e}")

        # 2. Load External Plugins from PLUGINS_DIR
        if os.path.exists(PLUGINS_DIR):
            for item in os.listdir(PLUGINS_DIR):
                p_dir = os.path.join(PLUGINS_DIR, item)
                if not os.path.isdir(p_dir):
                    continue
                
                # Check for manifest.json
                manifest_path = os.path.join(p_dir, "manifest.json")
                if not os.path.exists(manifest_path):
                    continue
                
                try:
                    with open(manifest_path, "r", encoding="utf-8") as f:
                        manifest = json.load(f)
                    
                    entry_point = manifest.get("entry")
                    if not entry_point:
                        continue
                    
                    # Add plugins dir to sys.path if not present
                    if PLUGINS_DIR not in sys.path:
                        sys.path.insert(0, PLUGINS_DIR)
                    
                    # Import from the specific plugin directory
                    mod_name, cls_name = entry_point.rsplit(".", 1)
                    spec = importlib.util.spec_from_file_location(
                        f"external_plugin_{item}", 
                        os.path.join(p_dir, mod_name + ".py")
                    )
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)
                    
                    # Instantiate the plugin class
                    cls = getattr(mod, cls_name)
                    plugin = cls()
                    plugin.set_context(self)
                    # Set plugin metadata from manifest
                    plugin.manifest = manifest
                    self.plugins.append(plugin)
                    print(f"Loaded external plugin: {manifest.get('name', item)}")
                except Exception as e:
                    print(f"Failed to load external plugin from {p_dir}: {e}")
                    traceback.print_exc()

    def update_splash(self, value, text):
        if self._splash:
            self._splash.set_progress(value, text)

    def _connect_signals(self):
        self.monitor.slideshow_started.connect(self.on_slideshow_start)
        self.monitor.slideshow_ended.connect(self.on_slideshow_end)

        self.overlay.request_next.connect(self.monitor.go_next)
        self.overlay.request_prev.connect(self.monitor.go_previous)
        self.overlay.request_end.connect(self.monitor.end_show)

        self.overlay.request_ptr_arrow.connect(lambda: self.monitor.set_pointer_type(1))
        self.overlay.request_ptr_pen.connect(lambda: self.monitor.set_pointer_type(2))
        self.overlay.request_ptr_eraser.connect(lambda: self.monitor.set_pointer_type(5))
        self.overlay.request_pen_color.connect(self.monitor.set_pen_color)

        self.tray.show_settings.connect(self.settings_plugin.execute)
        self.tray.show_timer.connect(self.timer_plugin.execute)
        self.tray.restart_app.connect(self.restart)
        self.tray.exit_app.connect(self.app.quit)

        self._timer_manager.finished.connect(self._on_timer_finished)

        self.monitor.slide_changed.connect(self.overlay.update_page_info)
        self.monitor.window_geometry_changed.connect(self.overlay.update_geometry)

    @Slot()
    def _on_timer_finished(self):
        now = time.monotonic()
        if now - self._last_timer_notify_at < 1.0:
            return
        self._last_timer_notify_at = now
        if hasattr(self, "tray") and self.tray:
            self.tray.show_message(t("timer.notify.title"), t("timer.notify.body"))
    
    @Slot()
    def on_slideshow_start(self):
        # Cleanup slide thumbnails from previous session
        temp_dir = os.path.join(tempfile.gettempdir(), "kazuha_ppt_thumbs")
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

        if cfg.autoShowOverlay.value:
            self.overlay.show()
            self.overlay.raise_()
            self.tray.show_message("PPT Assistant", "Slideshow detected. Overlay active.")
    
    @Slot()
    def on_slideshow_end(self):
        self.overlay.hide()

    def _check_settings_changed(self):
        if not os.path.exists(SETTINGS_PATH):
            return
        mtime = os.path.getmtime(SETTINGS_PATH)
        if mtime != self._settings_mtime:
            self._settings_mtime = mtime
            
            old_theme = cfg.themeMode.value
            old_lang = getattr(self, "_current_language", "zh-CN")
            old_qt_font = getattr(self, "_current_qt_font", "")
            old_overlay_font = getattr(self, "_current_overlay_font", "")
            old_toolbar_text = cfg.showToolbarText.value
            old_status_bar = cfg.showStatusBar.value
            old_undo_redo = cfg.showUndoRedo.value
            old_spotlight = cfg.showSpotlight.value
            old_timer = cfg.showTimer.value
            old_toolbar_order = cfg.toolbarOrder.value
            old_safe_area = cfg.safeArea.value
            old_scale = cfg.scale.value
            # old_layout_mode = cfg.toolbarLayout.value

            reload_cfg()

            try:
                with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}

            new_lang = (data.get("General", {}) or {}).get("Language", "zh-CN")
            profiles = (data.get("Fonts", {}) or {}).get("Profiles", {}) or {}
            lang_profile = profiles.get(new_lang, {}) or {}
            qt_font = lang_profile.get("qt", "")
            overlay_font = lang_profile.get("overlay", "") or qt_font
            new_qt_font = qt_font.strip() if isinstance(qt_font, str) else ""
            new_overlay_font = overlay_font.strip() if isinstance(overlay_font, str) else ""

            self._current_language = new_lang
            self._current_qt_font = new_qt_font
            self._current_overlay_font = new_overlay_font

            if new_qt_font != old_qt_font:
                _apply_global_font(self.app)
            
            # If language or critical layout changed, reload overlay
            if (new_lang != old_lang or 
                cfg.showToolbarText.value != old_toolbar_text or 
                cfg.showStatusBar.value != old_status_bar or
                cfg.showUndoRedo.value != old_undo_redo or
                cfg.showSpotlight.value != old_spotlight or
                cfg.showTimer.value != old_timer or
                cfg.toolbarOrder.value != old_toolbar_order or
                new_overlay_font != old_overlay_font or
                cfg.safeArea.value != old_safe_area or
                cfg.scale.value != old_scale):
                if not self._reloading_overlay:
                    self._reload_timer.start()
            
            # Layout mode change is now handled by auto-reload above, no restart prompt needed
            
            if cfg.themeMode.value != old_theme:
                if hasattr(self, 'tray'):
                    self.tray._update_icon()

    def _reload_overlay(self):
        """Recreate the overlay window to apply language and layout changes."""
        if self._reloading_overlay:
            return
        self._reloading_overlay = True
        was_visible = self.overlay.isVisible()
        
        try:
            # Import overlay again to refresh module-level LANGUAGE
            import ppt_assistant.ui.overlay as overlay_mod
            importlib.reload(overlay_mod)
            from ppt_assistant.ui.overlay import OverlayWindow
            
            # Create new overlay first (prevent crash if creation fails)
            new_overlay = OverlayWindow()
            new_overlay.set_monitor(self.monitor)
            
            # Re-connect signals
            new_overlay.request_next.connect(self.monitor.go_next)
            new_overlay.request_prev.connect(self.monitor.go_previous)
            new_overlay.request_end.connect(self.monitor.end_show)
            new_overlay.request_ptr_arrow.connect(lambda: self.monitor.set_pointer_type(1))
            new_overlay.request_ptr_pen.connect(lambda: self.monitor.set_pointer_type(2))
            new_overlay.request_ptr_eraser.connect(lambda: self.monitor.set_pointer_type(5))
            new_overlay.request_pen_color.connect(self.monitor.set_pen_color)
            
            # Disconnect old overlay slots before connecting new ones
            try:
                self.monitor.slide_changed.disconnect(self.overlay.update_page_info)
            except Exception:
                pass
            try:
                self.monitor.window_geometry_changed.disconnect(self.overlay.update_geometry)
            except Exception:
                pass
            self.monitor.slide_changed.connect(new_overlay.update_page_info)
            self.monitor.window_geometry_changed.connect(new_overlay.update_geometry)
            
            # Swap overlay
            old_overlay = self.overlay
            self.overlay = new_overlay
            
            # Cleanup old overlay
            old_overlay.cleanup() # Stop threads safely
            old_overlay.hide()
            old_overlay.deleteLater()
            
            if was_visible:
                self.overlay.show()
                self.overlay.raise_()
            
            # Update current page info immediately
            if self.monitor:
                curr, total = self.monitor.get_page_info()
                self.overlay.update_page_info(curr, total)
                self.monitor.force_update_geometry()
                
        except Exception as e:
            print(f"Error reloading overlay: {e}")
            # If failed, keep using the old overlay if it's still alive
            if was_visible and not self.overlay.isVisible():
                 self.overlay.show()
        finally:
            self._reloading_overlay = False

    def restart(self):
        self.cleanup()
        os.execl(sys.executable, sys.executable, *sys.argv)

    def cleanup(self):
        """Cleanup app resources and terminate subprocesses."""
        if hasattr(self, 'monitor'):
            self.monitor.stop_monitoring()
        if hasattr(self, 'settings_plugin'):
            self.settings_plugin.terminate()
        if hasattr(self, 'overlay'):
            self.overlay.cleanup()

    def run(self):
        # sys.exit(self.app.exec())
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    _apply_global_font(app)
    crash_handler = CrashHandler(app)
    # _handle_multi_instance(app)

    show_splash = True
    try:
        mode = cfg.splashMode.value
        if mode == "Never":
            show_splash = False
        elif mode == "HideOnAutoStart":
            args = [a.lower() for a in sys.argv]
            if "--autostart" in args or "-autostart" in args or "--silent" in args:
                show_splash = False
        elif mode == "TimeRange":
            from PySide6.QtCore import QTime
            start_str = cfg.splashStartTime.value
            end_str = cfg.splashEndTime.value
            start_t = QTime.fromString(start_str, "HH:mm")
            end_t = QTime.fromString(end_str, "HH:mm")
            now = QTime.currentTime()
            
            if start_t.isValid() and end_t.isValid():
                if start_t <= end_t:
                    if not (start_t <= now <= end_t):
                        show_splash = False
                else:
                    if not (now >= start_t or now <= end_t):
                        show_splash = False
    except Exception as e:
        print(f"Error determining splash visibility: {e}")
        show_splash = True

    splash = None
    if show_splash:
        splash = StartupSplash()
        splash.show()
        app.processEvents()

    app_instance = PPTAssistantApp(app, splash)
    crash_handler.set_app_instance(app_instance)
    sys.exit(app.exec())
