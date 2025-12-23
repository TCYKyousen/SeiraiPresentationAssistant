import sys
import os
import winreg
import psutil
import subprocess
import ctypes
from pathlib import Path
from datetime import datetime
import win32api
import win32con
import win32gui

from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal, QUrl, QObject, QPoint, QRect
from PyQt6.QtGui import QIcon, QFont
from qfluentwidgets import (
    setTheme,
    Theme,
    SystemTrayMenu,
    Action,
    setThemeColor,
    QConfig,
    OptionsConfigItem,
    OptionsValidator,
    EnumSerializer,
    ColorConfigItem,
    ConfigItem,
    MessageBox,
    PushButton,
    qconfig,
)
from ui.widgets import AnnotationWidget, TimerWindow, ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
from .ppt_client import PPTWorker
from .ppt_core import PPTState
from .version_manager import VersionManager
from .sound_manager import SoundManager
import pythoncom
import os
import time
from typing import Optional
from datetime import datetime

def log(msg):
    try:
        log_dir = os.path.join(os.getenv("APPDATA"), "Kazuha")
        log_path = os.path.join(log_dir, "debug.log")
        with open(log_path, "a") as f:
            f.write(f"{datetime.now()}: [BusinessLogic] {msg}\n")
    except:
        pass

# fallback 到 Temp
CONFIG_DIR = Path(os.getenv("APPDATA", os.getenv("TEMP", "C:\\"))) / "Kazuha"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = CONFIG_DIR / "config.json"


class Config(QConfig):
    themeMode = OptionsConfigItem(
        "Theme",
        "ThemeMode",
        "Auto",
        OptionsValidator(["Light", "Dark", "Auto"]),
        None,
    )
    enableStartUp = ConfigItem("General", "EnableStartUp", False)
    timerPosition = OptionsConfigItem(
        "General",
        "TimerPosition",
        "Center",
        OptionsValidator(
            ["Center", "TopLeft", "TopRight", "BottomLeft", "BottomRight"]
        ),
        None,
    )
    clockPosition = OptionsConfigItem(
        "General",
        "ClockPosition",
        "TopRight",
        OptionsValidator(
            ["TopLeft", "TopRight", "BottomLeft", "BottomRight"]
        ),
        None,
    )
    navPosition = OptionsConfigItem(
        "General",
        "NavPosition",
        "BottomSides",
        OptionsValidator(
            ["BottomSides", "MiddleSides"]
        ),
        None,
    )
    language = OptionsConfigItem(
        "General",
        "Language",
        "Simplified Chinese",
        OptionsValidator(["Simplified Chinese", "Traditional Chinese", "English", "Japanese", "Tibetan"]),
        None,
    )
    clockFontWeight = OptionsConfigItem(
        "Clock",
        "FontWeight",
        "Bold",
        OptionsValidator(["Light", "Normal", "DemiBold", "Bold"]),
        None
    )
    clockShowSeconds = ConfigItem("Clock", "ShowSeconds", False)
    clockShowDate = ConfigItem("Clock", "ShowDate", False)
    clockShowLunar = ConfigItem("Clock", "ShowLunar", False)
    enableClock = ConfigItem("General", "EnableClock", False)


cfg = Config()
cfg.file = CONFIG_PATH
try:
    cfg.load()
except Exception:
    pass

class SlideExportThread(QThread):
    def __init__(self, cache_dir):
        super().__init__()
        self.cache_dir = cache_dir
        
    def run(self):
        pythoncom.CoInitialize()
        try:
            import win32com.client
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except:
                app = win32com.client.Dispatch("PowerPoint.Application")
                
            if app.SlideShowWindows.Count > 0:
                presentation = app.ActivePresentation
                slides_count = presentation.Slides.Count
                
                if not os.path.exists(self.cache_dir):
                    os.makedirs(self.cache_dir)
                    
                for i in range(1, slides_count + 1):
                    thumb_path = os.path.join(self.cache_dir, f"slide_{i}.jpg")
                    if not os.path.exists(thumb_path):
                        try:
                            presentation.Slides(i).Export(thumb_path, "JPG", 320, 180)
                            time.sleep(0.01) # Yield CPU
                        except:
                            pass
        except Exception as e:
            print(f"Slide export error: {e}")
        pythoncom.CoUninitialize()

class BusinessLogicController(QWidget):
    def __init__(self):
        super().__init__()
        self._com_initialized = False
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception:
            self._com_initialized = False

        self.theme_mode = self.load_theme_setting()
        setTheme(self.theme_mode)
        
        # Connect to system theme changes
        qconfig.themeChanged.connect(self.on_system_theme_changed)
        
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(1, 1)
        self.move(-100, -100) 
        
        self.ppt_worker = PPTWorker()
        self.ppt_worker.state_updated.connect(self.on_state_updated)
        self.ppt_worker.start()
        self.last_state = PPTState()
        self.waiting_for_state = False
        
        # Initialize VersionManager with correct path
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        version_config_path = os.path.join(base_dir, "config", "version.json")
        self.version_manager = VersionManager(config_path=version_config_path, repo_owner="Haraguse", repo_name="Kazuha")
        
        # Initialize SoundManager
        sound_dir = os.path.join(base_dir, "SoundEffectResources")
        self.sound_manager = SoundManager(sound_dir)
        
        # Connect cleanup to application exit
        QApplication.instance().aboutToQuit.connect(self.sound_manager.cleanup)
        
        # 批注功能组件
        self.annotation_widget: Optional[AnnotationWidget] = None
        self.timer_window = None
        self.loading_overlay = None
        
        # UI组件引用（将在主程序中设置）
        self.toolbar: Optional[ToolBarWidget] = None
        self.nav_left: Optional[PageNavWidget] = None
        self.nav_right: Optional[PageNavWidget] = None
        self.spotlight: Optional[SpotlightOverlay] = None
        self.clock_widget: Optional[ClockWidget] = None
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_state)
        self.timer.start(500)
        
        self.widgets_visible = False
        self.slides_loaded = False
        self.conflicting_process_running = False

        self.process_check_timer = QTimer(self)
        self.process_check_timer.timeout.connect(self.check_conflicting_processes)
        self.process_check_timer.start(3000)

        # Sync autorun state
        is_auto = self.is_autorun()
        if cfg.enableStartUp.value != is_auto:
            cfg.set(cfg.enableStartUp, is_auto)
            cfg.save()

    def check_conflicting_processes(self):
        try:
            import subprocess
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output('tasklist', startupinfo=startupinfo).decode('gbk', errors='ignore').lower()
            running = ('classisland' in output) or ('classwidgets' in output)

            if running:
                if self.clock_widget and self.clock_widget.isVisible():
                    self.clock_widget.hide()
                if not self.conflicting_process_running and hasattr(self, "tray_icon") and cfg.enableClock.value:
                    self.tray_icon.showMessage("提示", "检测到 ClassIsland/ClassWidgets，已自动隐藏时钟组件。", QSystemTrayIcon.MessageIcon.Information, 2000)
                    self.sound_manager.play("ConflictCICW")
                self.conflicting_process_running = True
            else:
                if self.clock_widget and not self.clock_widget.isVisible() and self.widgets_visible and cfg.enableClock.value:
                    self.clock_widget.show()
                self.conflicting_process_running = False
        except Exception as e:
            pass

    def set_theme_color(self, color_str):
        setThemeColor(color_str)
        self.save_theme_setting(self.theme_mode) # Re-save to persist if needed, though config handles it

    def set_font(self, font_name="Bahnschrift"):
        font = QFont()
        font.setFamilies([font_name, "Microsoft YaHei"])
        font.setPixelSize(14) # Default size
        QApplication.setFont(font)
        # Update all top-level widgets
        for widget in QApplication.topLevelWidgets():
            widget.setFont(font)
        
    def set_animation(self, enabled):
        # qfluentwidgets doesn't have a global switch exposed easily, 
        # but we can try to set it on specific widgets if needed.
        # For now, this is a placeholder or we can implement specific animation toggles.
        pass

    def set_start_up(self, enabled):
        self.toggle_autorun(enabled)

    def toggle_tray_autorun(self, checked=False):
        enabled = bool(checked)
        if cfg.enableStartUp.value != enabled:
            cfg.set(cfg.enableStartUp, enabled)
            cfg.save()
        self.set_start_up(enabled)

    def set_tray_timer_position(self, pos):
        if cfg.timerPosition.value != pos:
            cfg.set(cfg.timerPosition, pos)
            cfg.save()
        if pos == "Center":
            self.timer_pos_center_action.setChecked(True)
            self.timer_pos_tl_action.setChecked(False)
            self.timer_pos_tr_action.setChecked(False)
            self.timer_pos_bl_action.setChecked(False)
            self.timer_pos_br_action.setChecked(False)
        elif pos == "TopLeft":
            self.timer_pos_center_action.setChecked(False)
            self.timer_pos_tl_action.setChecked(True)
            self.timer_pos_tr_action.setChecked(False)
            self.timer_pos_bl_action.setChecked(False)
            self.timer_pos_br_action.setChecked(False)
        elif pos == "TopRight":
            self.timer_pos_center_action.setChecked(False)
            self.timer_pos_tl_action.setChecked(False)
            self.timer_pos_tr_action.setChecked(True)
            self.timer_pos_bl_action.setChecked(False)
            self.timer_pos_br_action.setChecked(False)
        elif pos == "BottomLeft":
            self.timer_pos_center_action.setChecked(False)
            self.timer_pos_tl_action.setChecked(False)
            self.timer_pos_tr_action.setChecked(False)
            self.timer_pos_bl_action.setChecked(True)
            self.timer_pos_br_action.setChecked(False)
        elif pos == "BottomRight":
            self.timer_pos_center_action.setChecked(False)
            self.timer_pos_tl_action.setChecked(False)
            self.timer_pos_tr_action.setChecked(False)
            self.timer_pos_bl_action.setChecked(False)
            self.timer_pos_br_action.setChecked(True)

    def set_tray_clock_position(self, pos):
        if cfg.clockPosition.value != pos:
            cfg.set(cfg.clockPosition, pos)
            cfg.save()
        if pos == "TopLeft":
            self.clock_pos_tl_action.setChecked(True)
            self.clock_pos_tr_action.setChecked(False)
            self.clock_pos_bl_action.setChecked(False)
            self.clock_pos_br_action.setChecked(False)
        elif pos == "TopRight":
            self.clock_pos_tl_action.setChecked(False)
            self.clock_pos_tr_action.setChecked(True)
            self.clock_pos_bl_action.setChecked(False)
            self.clock_pos_br_action.setChecked(False)
        elif pos == "BottomLeft":
            self.clock_pos_tl_action.setChecked(False)
            self.clock_pos_tr_action.setChecked(False)
            self.clock_pos_bl_action.setChecked(True)
            self.clock_pos_br_action.setChecked(False)
        elif pos == "BottomRight":
            self.clock_pos_tl_action.setChecked(False)
            self.clock_pos_tr_action.setChecked(False)
            self.clock_pos_bl_action.setChecked(False)
            self.clock_pos_br_action.setChecked(True)

    def set_tray_nav_position(self, pos):
        # Always save/update to ensure snapping effect even if mode is same
        changed = (cfg.navPosition.value != pos)
        if changed:
            cfg.set(cfg.navPosition, pos)
            cfg.save()
            
        if self.widgets_visible:
            # Force update to snap back to correct position
            self.show_widgets(animate=True)
        elif changed and hasattr(self, "tray_icon"):
            self.tray_icon.showMessage("提示", "翻页组件位置设置已保存，将在下一次放映时生效。", QSystemTrayIcon.MessageIcon.Information, 2000)

        if pos == "BottomSides":
            self.nav_pos_bottom_action.setChecked(True)
            self.nav_pos_middle_action.setChecked(False)
        elif pos == "MiddleSides":
            self.nav_pos_bottom_action.setChecked(False)
            self.nav_pos_middle_action.setChecked(True)

    def set_timer_position(self, pos):
        # Position is handled when window is toggled
        pass

    def set_clock_position(self, pos):
        self.adjust_positions()

    def set_language(self, language):
        if cfg.language.value != language:
            cfg.set(cfg.language, language)
            cfg.save()
        self.show_warning(None, "语言设置将在重启后生效")

    def set_language_zh(self, checked=False):
        self.set_language("Simplified Chinese")
        if hasattr(self, "language_zh_action"):
            self.language_zh_action.setChecked(True)
        if hasattr(self, "language_en_action"):
            self.language_en_action.setChecked(False)

    def set_language_en(self, checked=False):
        self.set_language("English")
        if hasattr(self, "language_zh_action"):
            self.language_zh_action.setChecked(False)
        if hasattr(self, "language_en_action"):
            self.language_en_action.setChecked(True)

    def restart_application(self):
        app_path = os.path.abspath(sys.argv[0])
        if app_path.endswith(".py"):
            python_exe = sys.executable.replace("python.exe", "pythonw.exe")
            if not os.path.exists(python_exe):
                python_exe = sys.executable
            cmd = [python_exe, app_path]
        else:
            cmd = [sys.executable]
        try:
            subprocess.Popen(cmd, close_fds=True)
        except Exception:
            pass
        app = QApplication.instance()
        if app is not None:
            app.quit()

    def crash_test(self):
        raise RuntimeError("Crash test")
    
    def setup_connections(self):
        """设置UI组件与业务逻辑之间的信号连接"""
        if self.toolbar:
            self.toolbar.request_spotlight.connect(self.toggle_spotlight)
            if hasattr(self, 'change_pointer_mode'):
                self.toolbar.request_pointer_mode.connect(self.change_pointer_mode)
            if hasattr(self, 'change_pen_color'):
                self.toolbar.request_pen_color.connect(self.change_pen_color)
            if hasattr(self, 'clear_ink'):
                self.toolbar.request_clear_ink.connect(self.clear_ink)
            if hasattr(self, 'exit_slideshow'):
                self.toolbar.request_exit.connect(self.exit_slideshow)
            if hasattr(self, 'toggle_timer_window'):
                self.toolbar.request_timer.connect(self.toggle_timer_window)
            
        if self.nav_left:
            self.nav_left.prev_clicked.connect(self.prev_page)
            self.nav_left.next_clicked.connect(self.next_page)
            self.nav_left.request_slide_jump.connect(self.jump_to_slide)
            
        if self.nav_right:
            self.nav_right.prev_clicked.connect(self.prev_page)
            self.nav_right.next_clicked.connect(self.next_page)
            self.nav_right.request_slide_jump.connect(self.jump_to_slide)
    
    def setup_tray(self):
        """设置系统托盘图标和菜单"""
        import sys
        import os
        def icon_path(name):
            base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
            if hasattr(sys, "_MEIPASS"):
                return os.path.join(base_dir, "icons", name)
            return os.path.join(os.path.dirname(base_dir), "icons", name)
        
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(icon_path("trayicon.svg")))

        tray_menu = SystemTrayMenu(parent=self)
        
        tray_menu.addActions([
            Action("选项", triggered=self.show_settings),
            Action("退出", triggered=self.exit_application),
        ])
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
    def show_settings(self):
        if not hasattr(self, 'settings_window') or self.settings_window is None:
            from ui.settings_window import SettingsWindow
            self.settings_window = SettingsWindow()
            self.settings_window.configChanged.connect(self.on_settings_changed)
            self.settings_window.checkUpdateClicked.connect(self.check_for_updates)
            if hasattr(self.settings_window, 'set_theme'):
                self.settings_window.set_theme(self.theme_mode)
                
        self.settings_window.show()
        self.settings_window.activateWindow()
        
    def on_settings_changed(self):
        # Handle immediate updates
        self.toggle_autorun(cfg.enableStartUp.value)
        
        # Update clock visibility
        if self.clock_widget:
            if cfg.enableClock.value and self.widgets_visible and not self.conflicting_process_running:
                self.clock_widget.show()
            elif not cfg.enableClock.value:
                self.clock_widget.hide()
        
        # Update window effect for all widgets
        self.update_widgets_theme()
        
        # Apply theme change
        target_theme = cfg.themeMode.value
        if target_theme == "Auto":
            self.set_theme_auto(checked=True)
        elif target_theme == "Light":
            self.set_theme_light(checked=True)
        elif target_theme == "Dark":
            self.set_theme_dark(checked=True)
            
        # Update clock settings
        if self.clock_widget and hasattr(self.clock_widget, 'apply_settings'):
            self.clock_widget.apply_settings(cfg)
        
        # Save settings
        cfg.save()
        
        # Re-layout widgets if positions changed (with animation)
        if self.widgets_visible:
            self.show_widgets(animate=True)
            
    def show_about(self):
        title = "关于 Seirai PPT Assistant"
        content = (f"当前版本: {VersionManager.CURRENT_VERSION}\n"
                   "一个现代化、高性能的 PowerPoint 演示辅助工具。\n\n"
                   "© 2024 Seirai Studio")
                   
        w = MessageBox(title, content, self.settings_window if hasattr(self, 'settings_window') and self.settings_window else None)
        w.yesButton.setText("确定")
        w.cancelButton.hide()
        w.exec()
        self.timer_pos_center_action.setChecked(timer_pos == "Center")
        self.timer_pos_tl_action.setChecked(timer_pos == "TopLeft")
        self.timer_pos_tr_action.setChecked(timer_pos == "TopRight")
        self.timer_pos_bl_action.setChecked(timer_pos == "BottomLeft")
        self.timer_pos_br_action.setChecked(timer_pos == "BottomRight")

        clock_pos = cfg.clockPosition.value
        self.clock_pos_tl_action.setChecked(clock_pos == "TopLeft")
        self.clock_pos_tr_action.setChecked(clock_pos == "TopRight")
        self.clock_pos_bl_action.setChecked(clock_pos == "BottomLeft")
        self.clock_pos_br_action.setChecked(clock_pos == "BottomRight")

        nav_pos = cfg.navPosition.value
        self.nav_pos_bottom_action.setChecked(nav_pos == "BottomSides")
        self.nav_pos_middle_action.setChecked(nav_pos == "MiddleSides")

        self.tray_icon.setContextMenu(tray_menu)
        self.set_theme_mode(self.theme_mode)
        self.tray_icon.show()
        
        # Connect Version Manager Signals
        self.version_manager.update_available.connect(self.on_update_available)
        self.version_manager.update_error.connect(lambda e: self.show_warning(None, f"更新错误: {e}"))
        
        # Auto check for updates
        self.version_manager.check_for_updates()

    def check_for_updates(self):
        self.tray_icon.showMessage("PPT助手", "正在检查更新...", QSystemTrayIcon.MessageIcon.Information, 2000)
        self.version_manager.check_for_updates()

    def on_update_available(self, info):
        title = "发现新版本"
        content = f"版本: {info['version']}\n\n{info['body']}\n\n是否立即更新？"
        
        w = MessageBox(title, content, self.get_parent_for_dialog())
        if w.exec():
            # Find the exe asset
            asset_url = None
            for asset in info['assets']:
                if asset['name'].endswith('.exe'):
                    asset_url = asset['browser_download_url']
                    break
            
            if asset_url:
                self.tray_icon.showMessage("PPT助手", "正在后台下载更新...", QSystemTrayIcon.MessageIcon.Information, 3000)
                self.version_manager.download_and_install(asset_url)
            else:
                self.show_warning(None, "未找到可执行的更新文件")

    def show_about_dialog(self):
        version_info = self.version_manager.current_version_info
        v_name = version_info.get('versionName', 'Unknown')
        
        self.sound_manager.play("WindowPop")
        w = MessageBox("关于 PPT助手", f"当前版本: {v_name}\n\n一个专注于演示辅助的工具。\n\nDesigned by Seirai.", self.get_parent_for_dialog())
        w.exec()

    def get_parent_for_dialog(self):
        # MessageBox needs a parent window, but we are essentially hidden/tray based.
        # We can use a dummy widget or None (if supported by qfluentwidgets)
        # Using None might create a standalone window.
        return None

    def closeEvent(self, event):
        self.timer.stop()
        if hasattr(self, "ppt_worker") and self.ppt_worker:
            try:
                self.ppt_worker.stop()
            except Exception:
                pass
        if getattr(self, "_com_initialized", False):
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        # 清理批注功能
        if self.annotation_widget:
            self.annotation_widget.close()
        event.accept()
    
    def load_theme_setting(self):
        key = None
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Kazuha", 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, "ThemeMode")
            
            if isinstance(value, str):
                v = value.lower()
                if v == "light":
                    return Theme.LIGHT
                if v == "dark":
                    return Theme.DARK
                if v == "auto":
                    return Theme.AUTO
        except Exception:
            pass
        finally:
            if key:
                try:
                    winreg.CloseKey(key)
                except Exception:
                    pass
        return Theme.AUTO
    
    def save_theme_setting(self, theme):
        try:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Kazuha", 0, winreg.KEY_ALL_ACCESS)
            except Exception:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Kazuha")
            
            with key:
                value = getattr(theme, "value", None)
                if not isinstance(value, str):
                    value = str(theme)
                winreg.SetValueEx(key, "ThemeMode", 0, winreg.REG_SZ, value)
        except Exception as e:
            print(f"Error saving theme setting: {e}")
    
    def set_theme_mode(self, theme):
        self.theme_mode = theme
        
        # Ensure qfluentwidgets theme is set
        setTheme(theme)
        
        # Save setting
        self.save_theme_setting(theme)
        
        # Update tray menu actions state
        if hasattr(self, "theme_auto_action"):
            self.theme_auto_action.setChecked(theme == Theme.AUTO)
        if hasattr(self, "theme_light_action"):
            self.theme_light_action.setChecked(theme == Theme.LIGHT)
        if hasattr(self, "theme_dark_action"):
            self.theme_dark_action.setChecked(theme == Theme.DARK)
        
        self.update_widgets_theme()

    def on_system_theme_changed(self):
        """Handle system theme changes when in Auto mode"""
        if self.theme_mode == Theme.AUTO:
            self.update_widgets_theme()

            
    def update_widgets_theme(self):
        """Update theme for all active widgets"""
        theme = self.theme_mode
        widgets = [
            self.toolbar,
            self.nav_left,
            self.nav_right,
            self.timer_window,
            self.spotlight,
            self.clock_widget
        ]
        
        # Add SettingsWindow if it exists
        if hasattr(self, 'settings_window') and self.settings_window:
            widgets.append(self.settings_window)
        
        # Add optional widgets if they exist
        if hasattr(self, 'annotation_widget') and self.annotation_widget:
            widgets.append(self.annotation_widget)
            
        for widget in widgets:
            if widget:
                if hasattr(widget, 'set_theme'):
                    widget.set_theme(theme)
    
    def set_theme_auto(self, checked=False):
        self.set_theme_mode(Theme.AUTO)
    
    def set_theme_light(self, checked=False):
        self.set_theme_mode(Theme.LIGHT)
    
    def set_theme_dark(self, checked=False):
        self.set_theme_mode(Theme.DARK)
    
    def toggle_timer_window(self):
        if not self.timer_window:
            self.timer_window = TimerWindow()
            if hasattr(self.timer_window, 'set_theme'):
                self.timer_window.set_theme(self.theme_mode)
            if self.clock_widget:
                self.timer_window.timer_state_changed.connect(self.on_timer_state_changed)
                self.timer_window.countdown_finished.connect(self.on_countdown_finished)
                self.timer_window.timer_reset.connect(self.on_timer_reset)
                self.timer_window.pre_reminder_triggered.connect(self.sound_manager.speak)
                self.timer_window.emit_state()
        if self.timer_window.isVisible():
            self.timer_window.hide()
        else:
            self.timer_window.show()
            self.sound_manager.play("WindowPop")
            self.timer_window.activateWindow()
            self.timer_window.raise_()
            
            # Position the window
            pos_setting = cfg.timerPosition.value
            screen = QApplication.primaryScreen().geometry() # type: ignore
            w = self.timer_window.width()
            h = self.timer_window.height()
            
            if pos_setting == "Center":
                x = screen.left() + (screen.width() - w) // 2
                y = screen.top() + (screen.height() - h) // 2
            elif pos_setting == "TopLeft":
                x = screen.left() + 20
                y = screen.top() + 20
            elif pos_setting == "TopRight":
                x = screen.left() + screen.width() - w - 20
                y = screen.top() + 20
            elif pos_setting == "BottomLeft":
                x = screen.left() + 20
                y = screen.top() + screen.height() - h - 20
            elif pos_setting == "BottomRight":
                x = screen.left() + screen.width() - w - 20
                y = screen.top() + screen.height() - h - 20
            else: # Default Center
                x = screen.left() + (screen.width() - w) // 2
                y = screen.top() + (screen.height() - h) // 2
            
            self.timer_window.move(x, y)
    
    def on_timer_state_changed(self, up_seconds, up_running, down_remaining, down_running):
        if not self.clock_widget:
            return
        self.clock_widget.update_timer_state(up_seconds, up_running, down_remaining, down_running)
        self.clock_widget.adjustSize()
        if self.widgets_visible:
            self.adjust_positions()
    
    def on_countdown_finished(self):
        if self.timer_window and self.timer_window.strong_reminder_mode and self.timer_window.post_rem_switch.isChecked():
            self.sound_manager.speak("倒计时结束")

        if self.timer_window and self.timer_window.strong_reminder_mode:
            # Check for custom ringtone
            ring_data = self.timer_window.ringtone_combo.currentData()
            if ring_data and ring_data != "StrongTimerRing":
                # It's a path
                self.sound_manager.play("CustomRing", loop=True, custom_path=ring_data)
            else:
                self.sound_manager.play("StrongTimerRing", loop=True)
                
            # Force show timer window
            if not self.timer_window.isVisible():
                self.toggle_timer_window()
            else:
                self.timer_window.activateWindow()
                self.timer_window.raise_()
            
            # Show full screen mask
            if not self.loading_overlay:
                from ui.widgets import LoadingOverlay
                # We reuse LoadingOverlay or create a specific mask. 
                # Let's create a simple black mask if LoadingOverlay is not suitable or use a new MaskOverlay.
                # Since we don't have MaskOverlay, let's use a full screen window with semi-transparent black.
                self.mask_overlay = QWidget()
                self.mask_overlay.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
                self.mask_overlay.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
                self.mask_overlay.setStyleSheet("background-color: rgba(0, 0, 0, 180);")
                screen = QApplication.primaryScreen().geometry()
                self.mask_overlay.setGeometry(screen)
                self.mask_overlay.show()
                # Ensure timer window is above mask
                self.timer_window.raise_()
        else:
            self.sound_manager.play("TimerRing")
            
        if not self.clock_widget:
            return
        self.clock_widget.show_countdown_finished()
        self.clock_widget.adjustSize()
        if self.widgets_visible:
            self.adjust_positions()

    def on_timer_reset(self):
        self.sound_manager.stop("StrongTimerRing")
        self.sound_manager.stop("CustomRing") # Stop custom ring too
        self.sound_manager.stop("TimerRing")
        self.sound_manager.play("Reset")
        if hasattr(self, 'mask_overlay') and self.mask_overlay:
            self.mask_overlay.close()
            self.mask_overlay = None

    def toggle_annotation_mode(self):
        """切换独立批注模式"""
        # 在标准模式下使用AnnotationWidget
        if not self.annotation_widget:
            self.annotation_widget = AnnotationWidget()
            if hasattr(self.annotation_widget, 'set_theme'):
                self.annotation_widget.set_theme(self.theme_mode)
        
        if self.annotation_widget.isVisible():
            self.annotation_widget.hide()
        else:
            self.annotation_widget.showFullScreen()
    
    def check_presentation_processes(self):
        """检查演示进程并控制窗口显示"""
        presentation_detected = False
        
        # 检查PowerPoint或WPS进程
        for proc in psutil.process_iter(['name']):
            try:
                proc_name = proc.info.get('name', '') or ""
                proc_name_lower = proc_name.lower()
                
                # 检查是否为PowerPoint或WPS演示相关进程
                if any(keyword in proc_name_lower for keyword in ['powerpnt', 'wpp', 'wps']):
                    presentation_detected = True
                    break
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        
        return presentation_detected
    
    def is_autorun(self):#设定程序自启动
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, "Kazuha")
            return True
        except Exception:
            return False

    def toggle_autorun(self, checked):#程序未编译下自启动
        app_path = os.path.abspath(sys.argv[0])
        # If running as script
        if app_path.endswith('.py'):
            # Use pythonw.exe to avoid console if available, otherwise sys.executable
            python_exe = sys.executable.replace("python.exe", "pythonw.exe")
            if not os.path.exists(python_exe):
                python_exe = sys.executable
            cmd = f'"{python_exe}" "{app_path}"'
        else:
            # Frozen exe
            cmd = f'"{sys.executable}"'

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            if checked:
                winreg.SetValueEx(key, "Kazuha", 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, "Kazuha")
                except WindowsError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error setting autorun: {e}")

    def ensure_topmost(self):
        """Ensure all widgets are topmost"""
        try:
            flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            widgets = [
                self.toolbar, 
                self.nav_left, 
                self.nav_right, 
                self.clock_widget,
                self.timer_window,
                self.spotlight,
                self.annotation_widget
            ]
            
            for widget in widgets:
                if widget and widget.isVisible():
                    hwnd = widget.winId()
                    # HWND_TOPMOST = -1
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
                    widget.raise_()
        except Exception:
            pass

    def ensure_nav_widgets_orientation(self, animate=False):
        nav_pos = cfg.navPosition.value
        from PyQt6.QtCore import Qt, QPropertyAnimation, QParallelAnimationGroup, QSequentialAnimationGroup, QEasingCurve, QRect
        target_orientation = Qt.Orientation.Vertical if nav_pos == "MiddleSides" else Qt.Orientation.Horizontal
        
        if self.nav_left and hasattr(self.nav_left, 'orientation') and self.nav_left.orientation == target_orientation:
            return
            
        old_left = self.nav_left
        old_right = self.nav_right
        
        from ui.widgets import PageNavWidget
        self.nav_left = PageNavWidget(is_right=False, orientation=target_orientation)
        self.nav_right = PageNavWidget(is_right=True, orientation=target_orientation)
        
        self.nav_left.prev_clicked.connect(self.prev_page)
        self.nav_left.next_clicked.connect(self.next_page)
        self.nav_left.request_slide_jump.connect(self.jump_to_slide)
        
        self.nav_right.prev_clicked.connect(self.prev_page)
        self.nav_right.next_clicked.connect(self.next_page)
        self.nav_right.request_slide_jump.connect(self.jump_to_slide)
        
        if hasattr(self, 'theme_mode'):
             self.nav_left.set_theme(self.theme_mode)
             self.nav_right.set_theme(self.theme_mode)
        
        font = QApplication.font()
        self.nav_left.setFont(font)
        self.nav_right.setFont(font)
        
        if animate and old_left and old_right and self.widgets_visible:
             self.adjust_positions(self.last_state.hwnd)
             
             target_rect_l = self.nav_left.geometry()
             target_rect_r = self.nav_right.geometry()
             
             center_l = target_rect_l.center()
             center_r = target_rect_r.center()
             
             start_rect_l = QRect(center_l.x(), center_l.y(), 0, 0)
             start_rect_r = QRect(center_r.x(), center_r.y(), 0, 0)
             
             self.nav_left.setGeometry(start_rect_l)
             self.nav_right.setGeometry(start_rect_r)
             self.nav_left.setWindowOpacity(0)
             self.nav_right.setWindowOpacity(0)
             self.nav_left.show()
             self.nav_right.show()
             
             self.anim_group = QSequentialAnimationGroup(self)
             
             retract_group = QParallelAnimationGroup()
             
             r_anim_l_geo = QPropertyAnimation(old_left, b"geometry")
             r_anim_l_geo.setDuration(150)
             r_anim_l_geo.setStartValue(old_left.geometry())
             r_anim_l_geo.setEndValue(QRect(old_left.geometry().center().x(), old_left.geometry().center().y(), 0, 0))
             r_anim_l_geo.setEasingCurve(QEasingCurve.Type.InBack)
             
             r_anim_l_op = QPropertyAnimation(old_left, b"windowOpacity")
             r_anim_l_op.setDuration(150)
             r_anim_l_op.setStartValue(1.0)
             r_anim_l_op.setEndValue(0.0)
             
             r_anim_r_geo = QPropertyAnimation(old_right, b"geometry")
             r_anim_r_geo.setDuration(150)
             r_anim_r_geo.setStartValue(old_right.geometry())
             r_anim_r_geo.setEndValue(QRect(old_right.geometry().center().x(), old_right.geometry().center().y(), 0, 0))
             r_anim_r_geo.setEasingCurve(QEasingCurve.Type.InBack)
             
             r_anim_r_op = QPropertyAnimation(old_right, b"windowOpacity")
             r_anim_r_op.setDuration(150)
             r_anim_r_op.setStartValue(1.0)
             r_anim_r_op.setEndValue(0.0)
             
             retract_group.addAnimation(r_anim_l_geo)
             retract_group.addAnimation(r_anim_l_op)
             retract_group.addAnimation(r_anim_r_geo)
             retract_group.addAnimation(r_anim_r_op)
             
             pop_group = QParallelAnimationGroup()
             
             p_anim_l_geo = QPropertyAnimation(self.nav_left, b"geometry")
             p_anim_l_geo.setDuration(250)
             p_anim_l_geo.setStartValue(start_rect_l)
             p_anim_l_geo.setEndValue(target_rect_l)
             p_anim_l_geo.setEasingCurve(QEasingCurve.Type.OutBack)
             
             p_anim_l_op = QPropertyAnimation(self.nav_left, b"windowOpacity")
             p_anim_l_op.setDuration(150)
             p_anim_l_op.setStartValue(0.0)
             p_anim_l_op.setEndValue(1.0)
             
             p_anim_r_geo = QPropertyAnimation(self.nav_right, b"geometry")
             p_anim_r_geo.setDuration(250)
             p_anim_r_geo.setStartValue(start_rect_r)
             p_anim_r_geo.setEndValue(target_rect_r)
             p_anim_r_geo.setEasingCurve(QEasingCurve.Type.OutBack)
             
             p_anim_r_op = QPropertyAnimation(self.nav_right, b"windowOpacity")
             p_anim_r_op.setDuration(150)
             p_anim_r_op.setStartValue(0.0)
             p_anim_r_op.setEndValue(1.0)
             
             pop_group.addAnimation(p_anim_l_geo)
             pop_group.addAnimation(p_anim_l_op)
             pop_group.addAnimation(p_anim_r_geo)
             pop_group.addAnimation(p_anim_r_op)
             
             self.anim_group.addAnimation(retract_group)
             self.anim_group.addAnimation(pop_group)
             
             self.anim_group.finished.connect(lambda l=old_left, r=old_right: self._cleanup_old_nav_widgets(l, r))
             
             self.anim_group.start()
        
        else:
            if old_left:
                old_left.close()
                old_left.deleteLater()
            if old_right:
                old_right.close()
                old_right.deleteLater()

    def _cleanup_old_nav_widgets(self, old_left, old_right):
        if old_left:
            old_left.close()
            old_left.deleteLater()
        if old_right:
            old_right.close()
            old_right.deleteLater()

    def check_state(self):
        if self.waiting_for_state:
            return
        self.waiting_for_state = True
        self.ppt_worker.request_state()

    def on_state_updated(self, state: PPTState):
        self.waiting_for_state = False
        self.last_state = state
        
        if state.is_running:
            if not self.widgets_visible:
                self.show_widgets()
            else:
                self.ensure_topmost()
            
            # Attempt to provide ppt_app for SlideSelector if available
            # Warning: Passing direct COM object might be risky if used improperly
            # but currently widgets use it for some logic. 
            # Ideally we should remove this dependency too.
            # For now, we removed self.ppt_client so we can't pass it.
            # self.nav_left.ppt_app = ... 
            # If widgets need thumbnail generation, they should rely on the cache path we provide.

            if self.nav_left:
                self.nav_left.update_page(state.slide_index, state.slide_count, state.presentation_path)
            if self.nav_right:
                self.nav_right.update_page(state.slide_index, state.slide_count, state.presentation_path)
            
            self.sync_state(state.pointer_type)
            
            if state.presentation_path:
                # Debounce slide loading: only start if not already loading AND path changed or never loaded
                if not getattr(self, 'loading_slides', False):
                     if not self.slides_loaded or getattr(self, 'last_presentation_path', '') != state.presentation_path:
                         self.start_loading_slides(state.presentation_path)
        else:
            if self.widgets_visible:
                self.hide_widgets()

    def start_loading_slides(self, presentation_path):
        try:
            self.loading_slides = True
            self.last_presentation_path = presentation_path
            
            import hashlib
            try:
                path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
            except:
                path_hash = "default"
                
            cache_dir = os.path.join(os.environ['APPDATA'], 'PPTAssistant', 'Cache', path_hash)
            
            self.loader_thread = SlideExportThread(cache_dir)
            self.loader_thread.finished.connect(self.on_slides_loaded)
            self.loader_thread.start()
            
            # self.on_slides_loaded()
            
        except Exception as e:
            print(f"Error starting slide load: {e}")
            self.slides_loaded = True
            self.loading_slides = False

    def on_slides_loaded(self):
        self.slides_loaded = True
        self.loading_slides = False

    def change_pointer_mode(self, mode):
        self.ppt_worker.set_pointer_type(mode)
        self.sound_manager.play("Switch")
        if self.toolbar:
            self.toolbar.set_pointer_mode(mode)

    def hide_widgets(self):
        assert self.toolbar is not None
        assert self.nav_left is not None
        assert self.nav_right is not None
        self.toolbar.hide()
        self.nav_left.hide()
        self.nav_right.hide()
        if self.clock_widget:
            self.clock_widget.hide()
        self.widgets_visible = False
        
    def show_widgets(self, animate=False):
        self.ensure_nav_widgets_orientation(animate=animate)
        assert self.toolbar is not None
        assert self.nav_left is not None
        assert self.nav_right is not None
        self.toolbar.show()
        self.nav_left.show()
        self.nav_right.show()
        if self.clock_widget and not self.conflicting_process_running and cfg.enableClock.value:
            self.clock_widget.show()
            
        self.adjust_positions(self.last_state.hwnd, animate=animate)
        self.ensure_topmost()
        self.widgets_visible = True
        
    def adjust_positions(self, hwnd=0, animate=False):
        assert self.toolbar is not None
        assert self.nav_left is not None
        assert self.nav_right is not None

        screen = None
        try:
            if hwnd:
                monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
                info = win32api.GetMonitorInfo(monitor)
                work = info.get("Work") or info.get("Monitor")
                if work:
                    left, top, right, bottom = work
                    screen = (left, top, right - left, bottom - top)
        except Exception:
            screen = None

        if screen is None:
            if self.toolbar and self.toolbar.screen():
                g = self.toolbar.screen().geometry()  # type: ignore
                screen = (g.left(), g.top(), g.width(), g.height())
            else:
                g = QApplication.primaryScreen().geometry()  # type: ignore
                screen = (g.left(), g.top(), g.width(), g.height())

        MARGIN = 20
        left, top, width, height = screen
        right = left + width
        bottom = top + height
        
        tb_w = self.toolbar.sizeHint().width()
        tb_h = self.toolbar.sizeHint().height()
        tb_w = min(tb_w, width - 2 * MARGIN)
        tb_x = left + (width - tb_w) // 2
        tb_y = top + height - tb_h - MARGIN
        self.toolbar.setGeometry(tb_x, tb_y, tb_w, tb_h)
        
        nav_w = self.nav_left.sizeHint().width()
        nav_h = self.nav_left.sizeHint().height()
        
        nav_pos_setting = cfg.navPosition.value
        if nav_pos_setting == "MiddleSides":
            # Vertical layout centered vertically
            nav_y = top + (height - nav_h) // 2
        else:
            # Bottom sides (default)
            nav_y = top + height - nav_h - MARGIN

        target_nav_l = QRect(left + MARGIN, nav_y, nav_w, nav_h)
        target_nav_r = QRect(right - nav_w - MARGIN, nav_y, nav_w, nav_h)

        if animate and self.widgets_visible and (self.nav_left.geometry() != target_nav_l or self.nav_right.geometry() != target_nav_r):
             from PyQt6.QtCore import QPropertyAnimation, QEasingCurve, QParallelAnimationGroup
             
             self._nav_anim_group = QParallelAnimationGroup(self)
             
             anim_l = QPropertyAnimation(self.nav_left, b"geometry", self.nav_left)
             anim_l.setDuration(400)
             anim_l.setStartValue(self.nav_left.geometry())
             anim_l.setEndValue(target_nav_l)
             anim_l.setEasingCurve(QEasingCurve.Type.OutCubic)
             
             anim_r = QPropertyAnimation(self.nav_right, b"geometry", self.nav_right)
             anim_r.setDuration(400)
             anim_r.setStartValue(self.nav_right.geometry())
             anim_r.setEndValue(target_nav_r)
             anim_r.setEasingCurve(QEasingCurve.Type.OutCubic)
             
             self._nav_anim_group.addAnimation(anim_l)
             self._nav_anim_group.addAnimation(anim_r)
             self._nav_anim_group.start()
        else:
            self.nav_left.setGeometry(target_nav_l)
            self.nav_right.setGeometry(target_nav_r)

        try:
            flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            for w in (self.toolbar, self.nav_left, self.nav_right, self.clock_widget):
                if not w:
                    continue
                win32gui.SetWindowPos(int(w.winId()), win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        except Exception:
            pass
        
        if self.clock_widget:
            self.clock_widget.adjustSize()
            cw = self.clock_widget.width()
            ch = self.clock_widget.height()
            
            # Use configured position
            pos_setting = cfg.clockPosition.value
            
            if pos_setting == "TopLeft":
                x = left + MARGIN
                y = top + MARGIN
            elif pos_setting == "BottomLeft":
                x = left + MARGIN
                y = bottom - ch - MARGIN
                if nav_pos_setting == "BottomSides":
                    y = y - nav_h - MARGIN
            elif pos_setting == "BottomRight":
                x = right - cw - MARGIN
                y = bottom - ch - MARGIN
                if nav_pos_setting == "BottomSides":
                    y = y - nav_h - MARGIN
            else: # Default TopRight
                x = right - cw - MARGIN
                y = top + MARGIN
                
            # Boundary checks (keep within screen)
            if x < left + MARGIN: x = left + MARGIN
            if x > right - cw - MARGIN: x = right - cw - MARGIN
            if y < top + MARGIN: y = top + MARGIN
            if y > bottom - ch - MARGIN: y = bottom - ch - MARGIN
            
            target_geo = QRect(x, y, cw, ch)
            current_geo = self.clock_widget.geometry()
            
            # Animate if requested and visible and position changed significantly
            if animate and self.clock_widget.isVisible() and current_geo != target_geo and (current_geo.topLeft() - target_geo.topLeft()).manhattanLength() > 20:
                 from PyQt6.QtCore import QPropertyAnimation, QEasingCurve
                 anim = QPropertyAnimation(self.clock_widget, b"geometry", self.clock_widget)
                 anim.setDuration(400)
                 anim.setStartValue(current_geo)
                 anim.setEndValue(target_geo)
                 anim.setEasingCurve(QEasingCurve.Type.OutCubic)
                 anim.start()
                 # Keep reference
                 self._clock_anim = anim
            else:
                 self.clock_widget.setGeometry(target_geo)

    def sync_state(self, pointer_type):
        try:
            assert self.toolbar is not None
            if pointer_type == 1:
                self.toolbar.btn_arrow.setChecked(True)
            elif pointer_type == 2:
                self.toolbar.btn_pen.setChecked(True)
            elif pointer_type == 5: # Eraser
                self.toolbar.btn_eraser.setChecked(True)
        except:
            pass


    def go_prev(self):
        self.ppt_worker.prev_slide()

    def go_next(self):
        self.ppt_worker.next_slide()
        
    def next_page(self):
        """下一页"""
        current = self.last_state.slide_index
        total = self.last_state.slide_count
        if current >= total and total > 0:
             # Last page: Trigger exit flow
             self.exit_slideshow()
        else:
             self.go_next()
             # self.sound_manager.play("Cursor")
        
    def prev_page(self):
        """上一页"""
        self.go_prev()
        # self.sound_manager.play("Cursor")
                
    def jump_to_slide(self, index):
        self.ppt_worker.goto_slide(index)
        # self.sound_manager.play("Cursor")
        self.sound_manager.play("Confirm")

    def set_pointer(self, type_id):
        if type_id == 5:
            # Check for ink using last state
            if not self.last_state.has_ink:
                self.show_warning(None, "当前页没有笔迹")
        
        self.ppt_worker.set_pointer_type(type_id)
    
    def set_pen_color(self, color):
        self.ppt_worker.set_pen_color(color)
                
    def change_pen_color(self, color):
        """更改笔颜色"""
        self.set_pen_color(color)
        self.sound_manager.play("Confirm")
                
    def clear_ink(self):
        if not self.last_state.has_ink:
            self.show_warning(None, "当前页没有笔迹")
        self.ppt_worker.erase_ink()
        self.sound_manager.play("ClearAll")
                
    def toggle_spotlight(self):
        assert self.spotlight is not None
        if self.spotlight.isVisible():
            self.spotlight.hide()
        else:
            self.spotlight.showFullScreen()
            self.sound_manager.play("Switch")
            

    def exit_slideshow(self):
        has_ink = self.last_state.has_ink

        if not has_ink:
            self.ppt_worker.exit_show(keep_ink=False)
            return
        
        # Create a transparent full-screen widget as parent
        dummy_parent = QWidget()
        dummy_parent.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        dummy_parent.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        dummy_parent.setGeometry(QApplication.primaryScreen().geometry()) # type: ignore
        dummy_parent.show()
        
        from qfluentwidgets import MessageBox
        
        w = MessageBox("退出放映", "是否保留墨迹注释？", dummy_parent)
        w.yesButton.setText("保留")
        w.cancelButton.setText("丢弃")
        
        if w.exec():
            # Keep ink
            self.ppt_worker.exit_show(keep_ink=True)
        else:
            # Discard ink
            self.ppt_worker.exit_show(keep_ink=False)
            
        dummy_parent.close()

    def exit_application(self):
        """退出应用程序"""
        self.exit_slideshow()
        if hasattr(self, 'ppt_worker'):
            self.ppt_worker.stop()
        app = QApplication.instance()
        if app is not None:
            app.quit()

    def show_warning(self, target, message):
        title = "PPT助手提示"
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Warning, 2000)

