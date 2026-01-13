import sys
import os
import traceback
import tempfile
import subprocess
import json
import importlib

from PySide6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, QFrame, QGraphicsDropShadowEffect, QProgressBar
from PySide6.QtCore import Qt, QTimer, Slot, QSize
from PySide6.QtGui import QFontDatabase, QFont, QColor, QIcon

from ppt_assistant.core.ppt_monitor import PPTMonitor
from ppt_assistant.ui.overlay import OverlayWindow
from plugins.builtins.settings.plugin import SettingsPlugin
from plugins.builtins.timer.plugin import TimerPlugin
from ppt_assistant.ui.tray import SystemTray
from ppt_assistant.core.config import cfg, SETTINGS_PATH, reload_cfg, _apply_theme_and_color


SPLASH_I18N = {
    "zh-CN": {
        "initializing": "正在初始化",
    },
    "zh-TW": {
        "initializing": "正在初始化",
    },
    "ja-JP": {
        "initializing": "初期化中",
    },
    "en-US": {
        "initializing": "Initializing",
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
    if not os.path.exists(font_path):
        return
    font_id = QFontDatabase.addApplicationFont(font_path)
    if font_id == -1:
        return
    families = QFontDatabase.applicationFontFamilies(font_id)
    if not families:
        return
    family = families[0]
    font = QFont(family)
    app.setFont(font)


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
    return parts[-1] == "1"


class StartupSplash(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self._container = QFrame(self)
        self._container.setObjectName("splashContainer")

        self._version_text, self._code_name_en, self._code_name_cn = _load_version_info()
        self._language = _get_current_language()

        self._build_ui()
        self._apply_styles()
        self._center_on_screen()

        self._progress_timer = QTimer(self)
        self._progress_timer.timeout.connect(self._advance_progress)
        self._progress_timer.start(30)

        if _is_dev_preview_version(self._version_text):
            self._dev_watermark = QLabel(self)
            self._dev_watermark.setText(f"开发中版本/技术预览版本\n不保证最终品质 （{self._version_text}）")
            font = QFont()
            font.setPixelSize(11)
            self._dev_watermark.setFont(font)
            self._dev_watermark.setAlignment(Qt.AlignRight | Qt.AlignBottom)
            self._dev_watermark.setStyleSheet(
                "color: rgba(255, 255, 255, 100);"
            )
            self._dev_watermark.resize(320, 36)
            self._dev_watermark.move(self.width() - self._dev_watermark.width() - 16,
                                     self.height() - self._dev_watermark.height() - 12)

    def _build_ui(self):
        # Using absolute positioning based on QML study
        self._container.setFixedSize(899, 286)
        
        # Logo (kZHTXT_2)
        self._icon_label = QLabel(self._container)
        self._icon_label.setFixedSize(64, 64)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "logo.svg")
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            pix = icon.pixmap(64, 64)
            self._icon_label.setPixmap(pix)
        self._icon_label.move(38, 47)

        # Brand name (kazuha)
        self._brand_label = QLabel("Kazuha", self._container)
        brand_font = QFont("Bahnschrift")
        brand_font.setPixelSize(36)
        brand_font.setBold(True)
        self._brand_label.setFont(brand_font)
        self._brand_label.setFixedWidth(328)
        self._brand_label.move(38, 143)

        # Version/Codename (Right side - following QML's momoka_Kawaragi_ image position)
        self._right_container = QWidget(self._container)
        self._right_container.setFixedWidth(226) # 899 - 635 - 38
        self._right_container.move(635, 47)
        right_layout = QVBoxLayout(self._right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(2)

        self._version_label = QLabel(self._version_text or "")
        version_font = QFont("Bahnschrift")
        version_font.setPixelSize(12)
        version_font.setStretch(75)
        self._version_label.setFont(version_font)
        self._version_label.setAlignment(Qt.AlignRight)

        self._code_name_en_label = QLabel(self._code_name_en or "")
        code_en_font = QFont("Bahnschrift")
        code_en_font.setPixelSize(22)
        code_en_font.setBold(True)
        code_en_font.setStretch(75)
        self._code_name_en_label.setFont(code_en_font)
        self._code_name_en_label.setAlignment(Qt.AlignRight)

        self._code_name_cn_label = QLabel(self._code_name_cn or "")
        code_cn_font = QFont()
        code_cn_font.setPixelSize(11)
        self._code_name_cn_label.setFont(code_cn_font)
        self._code_name_cn_label.setAlignment(Qt.AlignRight)

        right_layout.addWidget(self._version_label)
        right_layout.addWidget(self._code_name_en_label)
        right_layout.addWidget(self._code_name_cn_label)

        # Status text (element)
        init_text = SPLASH_I18N.get(self._language, SPLASH_I18N["zh-CN"])["initializing"]
        self._percent_label = QLabel(f"8% {init_text}", self._container)
        percent_font = QFont()
        percent_font.setPixelSize(13)
        self._percent_label.setFont(percent_font)
        self._percent_label.setFixedWidth(221)
        self._percent_label.setAlignment(Qt.AlignRight)
        self._percent_label.move(643, 220)

        # Progress bar (rectangle_30 & rectangle_31)
        self._progress = QProgressBar(self._container)
        self._progress.setRange(0, 100)
        self._progress.setValue(8)
        self._progress.setTextVisible(False)
        self._progress.setFixedHeight(8)
        self._progress.setFixedWidth(823) # 899 - 38 - 38
        self._progress.move(38, 248)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 40))
        self._container.setGraphicsEffect(shadow)

    def _apply_styles(self):
        from qfluentwidgets import isDarkTheme
        self._is_dark = isDarkTheme()
        
        bg_color = "#202020" if self._is_dark else "white"
        brand_color = "#FFFFFF" if self._is_dark else "#000000"
        version_color = "#E0E0E0" if self._is_dark else "#222222"
        codename_en_color = "#FFFFFF" if self._is_dark else "#000000"
        codename_cn_color = "#AAAAAA" if self._is_dark else "#999999"
        percent_color = "#FFFFFF" if self._is_dark else "#000000"
        progress_bg = "rgba(255, 255, 255, 0.1)" if self._is_dark else "rgba(69, 69, 69, 0.16)"
        chunk_color = "#E1EBFF" if self._is_dark else "#3275F5"

        self.resize(899, 286)
        self._container.setStyleSheet(
            f"QFrame#splashContainer {{"
            f"background-color: {bg_color};"
            "border-radius: 5px;"
            "}"
        )
        self._brand_label.setStyleSheet(f"color: {brand_color};")
        self._version_label.setStyleSheet(f"color: {version_color};")
        self._code_name_en_label.setStyleSheet(f"color: {codename_en_color};")
        self._code_name_cn_label.setStyleSheet(f"color: {codename_cn_color};")
        self._percent_label.setStyleSheet(f"color: {percent_color};")
        self._progress.setStyleSheet(
            "QProgressBar {"
            f"background-color: {progress_bg};"
            "border-radius: 4px;"
            "border: none;"
            "}"
            "QProgressBar::chunk {"
            f"background-color: {chunk_color};"
            "border-radius: 4px;"
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

    def _advance_progress(self):
        value = self._progress.value()
        if value < 95:
            value += 1
            self._progress.setValue(value)
            init_text = SPLASH_I18N.get(self._language, SPLASH_I18N["zh-CN"])["initializing"]
            self._percent_label.setText(f"{value}% {init_text}")
        else:
            self._progress_timer.stop()

    def finish(self):
        self._progress_timer.stop()
        self._progress.setValue(100)
        self._percent_label.setText("100% 初始化完成")
        QTimer.singleShot(250, self.close)

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


class PPTAssistantApp:
    def __init__(self, app: QApplication, splash=None):
        self.app = app
        self.app.setQuitOnLastWindowClosed(False)
        self._splash = splash
        
        _apply_theme_and_color(cfg.themeMode.value)
        
        _apply_global_font(self.app)

        self._settings_mtime = os.path.getmtime(SETTINGS_PATH) if os.path.exists(SETTINGS_PATH) else 0
        self._settings_timer = QTimer()
        self._settings_timer.setInterval(100)
        self._settings_timer.timeout.connect(self._check_settings_changed)
        self._settings_timer.start()

        self.app.aboutToQuit.connect(self.cleanup)

        self.monitor = PPTMonitor()
        self.overlay = OverlayWindow()
        self.settings_plugin = SettingsPlugin()
        self.timer_plugin = TimerPlugin()
        self.tray = SystemTray()
        self.overlay.set_monitor(self.monitor)

        self._connect_signals()

        self.monitor.start_monitoring()

        if self._splash is not None:
            self._splash.finish()

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

        self.monitor.slide_changed.connect(self.overlay.update_page_info)

    @Slot()
    def on_slideshow_start(self):
        if cfg.autoShowOverlay.value:
            self.overlay.showFullScreen()
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
            old_lang = _get_current_language()
            old_toolbar_text = cfg.showToolbarText.value
            old_status_bar = cfg.showStatusBar.value
            old_undo_redo = cfg.showUndoRedo.value

            reload_cfg()
            
            new_lang = _get_current_language()
            
            # If language or critical layout changed, reload overlay
            if (new_lang != old_lang or 
                cfg.showToolbarText.value != old_toolbar_text or 
                cfg.showStatusBar.value != old_status_bar or
                cfg.showUndoRedo.value != old_undo_redo):
                self._reload_overlay()
            
            if cfg.themeMode.value != old_theme:
                if hasattr(self, 'tray'):
                    self.tray._update_icon()

    def _reload_overlay(self):
        """Recreate the overlay window to apply language and layout changes."""
        was_visible = self.overlay.isVisible()
        
        # Cleanup old overlay
        self.overlay.hide()
        self.overlay.deleteLater()
        
        # Import overlay again to refresh module-level LANGUAGE
        import ppt_assistant.ui.overlay as overlay_mod
        importlib.reload(overlay_mod)
        from ppt_assistant.ui.overlay import OverlayWindow
        
        # Create new overlay
        self.overlay = OverlayWindow()
        self.overlay.set_monitor(self.monitor)
        
        # Re-connect signals
        self.overlay.request_next.connect(self.monitor.go_next)
        self.overlay.request_prev.connect(self.monitor.go_previous)
        self.overlay.request_end.connect(self.monitor.end_show)
        self.overlay.request_ptr_arrow.connect(lambda: self.monitor.set_pointer_type(1))
        self.overlay.request_ptr_pen.connect(lambda: self.monitor.set_pointer_type(2))
        self.overlay.request_ptr_eraser.connect(lambda: self.monitor.set_pointer_type(5))
        self.overlay.request_pen_color.connect(self.monitor.set_pen_color)
        
        self.monitor.slide_changed.connect(self.overlay.update_page_info)
        
        if was_visible:
            if self.monitor and self.monitor._running:
                self.overlay.showFullScreen()
            else:
                self.overlay.show()
            self.overlay.raise_()
        
        # Update current page info immediately
        if self.monitor:
            curr, total = self.monitor.get_page_info()
            self.overlay.update_page_info(curr, total)

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
        sys.exit(self.app.exec())


if __name__ == "__main__":
    if "--webview-runner" in sys.argv:
        idx = sys.argv.index("--webview-runner")
        import plugins.webview_runner as _wv
        sys.argv = ["webview_runner.py"] + sys.argv[idx + 1 :]
        _wv.main()
        sys.exit(0)

    app = QApplication(sys.argv)
    _apply_global_font(app)
    crash_handler = CrashHandler(app)
    _handle_multi_instance(app)

    splash = StartupSplash()
    splash.show()
    app.processEvents()

    app_instance = PPTAssistantApp(app, splash)
    crash_handler.set_app_instance(app_instance)
    app_instance.run()

