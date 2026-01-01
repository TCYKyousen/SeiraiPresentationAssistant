import sys
import os
import winreg
import psutil
import subprocess
import ctypes
import pythoncom
import time
import win32api
import win32con
import win32gui
import win32com.client
from pathlib import Path
from datetime import datetime
from typing import Optional

from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal, QUrl, QObject, QPoint, QRect, QCoreApplication
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
    FluentIcon as FIF
)
from ui.widgets import AnnotationWidget, TimerWindow, ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
from .ppt_client import PPTWorker
from .ppt_core import PPTState
from .version_manager import VersionManager
from .sound_manager import SoundManager
try:
    from windows_toasts import WindowsToaster, ToastText2, ToastButton, ToastActivatedEventArgs
    HAS_WINDOWS_TOASTS = True
except ImportError:
    HAS_WINDOWS_TOASTS = False


def get_app_base_dir():
    return Path(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0]))))


def tr(text: str) -> str:
    return text

def log(msg):
    try:
        base_dir = get_app_base_dir()
        log_path = base_dir / "debug.log"
        with open(log_path, "a") as f:
            f.write(f"{datetime.now()}: [BusinessLogic] {msg}\n")
    except:
        pass

CONFIG_DIR = get_app_base_dir()
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
    enableSystemNotification = ConfigItem("General", "EnableSystemNotification", True)
    enableGlobalSound = ConfigItem("General", "EnableGlobalSound", True)

    enableGlobalAnimation = ConfigItem("General", "EnableGlobalAnimation", True)
    checkUpdateOnStart = ConfigItem("General", "CheckUpdateOnStart", True)
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
    screenPaddingTop = ConfigItem("Layout", "ScreenPaddingTop", 20)
    screenPaddingBottom = ConfigItem("Layout", "ScreenPaddingBottom", 20)
    screenPaddingLeft = ConfigItem("Layout", "ScreenPaddingLeft", 20)
    screenPaddingRight = ConfigItem("Layout", "ScreenPaddingRight", 20)
    screenPaddingLock = ConfigItem("Layout", "ScreenPaddingLock", True)
    screenPadding = OptionsConfigItem(
        "Layout",
        "ScreenPadding",
        "Normal",
        OptionsValidator(
            ["Small", "Normal", "Large"]
        ),
        None,
    )


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
                            time.sleep(0.01)
                        except:
                            pass
        except Exception as e:
            print(f"Slide export error: {e}")
        pythoncom.CoUninitialize()

class BusinessLogicController(QWidget):
    request_show_settings = pyqtSignal()
    request_show_update = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.request_show_settings.connect(self.show_settings)
        self.request_show_update.connect(self.show_update_interface)
        self._com_initialized = False
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
        except Exception:
            self._com_initialized = False

        tv = cfg.themeMode.value
        if tv == "Light":
            self.theme_mode = Theme.LIGHT
        elif tv == "Dark":
            self.theme_mode = Theme.DARK
        else:
            self.theme_mode = Theme.AUTO
        setTheme(self.theme_mode)
        
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
        
        base_dir = get_app_base_dir()
        version_config_path = base_dir / "config" / "version.json"
        if not version_config_path.exists():
            version_config_path = base_dir / "version.json"
        self.version_manager = VersionManager(config_path=str(version_config_path), repo_owner="Haraguse", repo_name="Kazuha")
        self.version_manager.update_available.connect(self.on_update_available)
        self.version_manager.update_error.connect(lambda e: self.show_warning(None, tr("更新错误: {e}").format(e=e)))

        self.version_manager.update_check_finished.connect(self.on_update_check_finished)
        
        if cfg.checkUpdateOnStart.value:
            QTimer.singleShot(2000, self.check_for_updates)

        sound_dir = base_dir / "resources" / "sound_effects"
        self.sound_manager = SoundManager(str(sound_dir))
        
        QApplication.instance().aboutToQuit.connect(self.sound_manager.cleanup)
        
        self.annotation_widget: Optional[AnnotationWidget] = None
        self.timer_window = None
        self.loading_overlay = None
        
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

        try:
            is_auto = self.is_autorun()
        except Exception:
            is_auto = False
        if cfg.enableStartUp.value != is_auto:
            cfg.set(cfg.enableStartUp, is_auto)
            cfg.save()

    def check_conflicting_processes(self):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output('tasklist', startupinfo=startupinfo).decode('gbk', errors='ignore').lower()
            running = ('classisland' in output) or ('classwidgets' in output)

            if running:
                if self.clock_widget and self.clock_widget.isVisible():
                    self.clock_widget.hide()
                if not self.conflicting_process_running and hasattr(self, "tray_icon") and cfg.enableClock.value:
                    self.tray_show_message("Kazuha", tr("检测到 ClassIsland/ClassWidgets，已自动隐藏时钟组件。"), QSystemTrayIcon.MessageIcon.Information, 2000)
                    self.play_sound("ConflictCICW")
                self.conflicting_process_running = True
            else:
                if self.clock_widget and not self.clock_widget.isVisible() and self.widgets_visible and cfg.enableClock.value:
                    self.clock_widget.show()
                self.conflicting_process_running = False
        except Exception as e:
            pass

    def set_theme_color(self, color_str):
        setThemeColor(color_str)

    def set_font(self, font_name="Bahnschrift"):
        font = QFont()
        font.setFamilies([font_name, "Microsoft YaHei"])
        font.setPixelSize(14)
        QApplication.setFont(font)
        for widget in QApplication.topLevelWidgets():
            widget.setFont(font)
        
    def play_sound(self, name, *args, **kwargs):
        if not cfg.enableGlobalSound.value:
            return
        self.sound_manager.play(name, *args, **kwargs)
        
    def speak_text(self, text):
        if not cfg.enableGlobalSound.value:
            return
        self.sound_manager.speak(text)
        
    def tray_show_message(self, title, message, icon, timeout):
        if not cfg.enableSystemNotification.value:
            return
        if hasattr(self, "tray_icon") and self.tray_icon:
            self.tray_icon.showMessage(title, message, icon, timeout)
        
    def set_animation(self, enabled):
        if cfg.enableGlobalAnimation.value != bool(enabled):
            cfg.set(cfg.enableGlobalAnimation, bool(enabled))
            cfg.save()

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
        changed = (cfg.navPosition.value != pos)
        if changed:
            cfg.set(cfg.navPosition, pos)
            cfg.save()
            
        if self.widgets_visible:
            self.show_widgets(animate=cfg.enableGlobalAnimation.value)
        elif changed and hasattr(self, "tray_icon"):
            self.tray_show_message("Kazuha", tr("翻页组件位置设置已保存，将在下一次放映时生效。"), QSystemTrayIcon.MessageIcon.Information, 2000)

        if pos == "BottomSides":
            self.nav_pos_bottom_action.setChecked(True)
            self.nav_pos_middle_action.setChecked(False)
        elif pos == "MiddleSides":
            self.nav_pos_bottom_action.setChecked(False)
            self.nav_pos_middle_action.setChecked(True)

    def set_timer_position(self, pos):
        pass

    def set_clock_position(self, pos):
        self.adjust_positions()

    def soft_restart(self):
        cfg.save()
        app = QApplication.instance()
        if app is None:
            return
        
        prev_visible = getattr(self, "widgets_visible", False)
        settings_visible = False
        if hasattr(self, "settings_window") and self.settings_window and self.settings_window.isVisible():
            settings_visible = True
            
        windows = [
            ("settings_window", getattr(self, "settings_window", None)),
            ("toolbar", getattr(self, "toolbar", None)),
            ("nav_left", getattr(self, "nav_left", None)),
            ("nav_right", getattr(self, "nav_right", None)),
            ("clock_widget", getattr(self, "clock_widget", None)),
            ("spotlight", getattr(self, "spotlight", None)),
            ("timer_window", getattr(self, "timer_window", None)),
            ("annotation_widget", getattr(self, "annotation_widget", None)),
        ]
        for name, w in windows:
            if w:
                try:
                    w.close()
                    w.deleteLater()
                except Exception:
                    pass
            setattr(self, name, None)

        if hasattr(self, "tray_icon") and self.tray_icon:
            try:
                self.tray_icon.hide()
                self.tray_icon.deleteLater()
            except Exception:
                pass
            self.tray_icon = None

        from ui.widgets import ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
        nav_pos = cfg.navPosition.value
        orientation = Qt.Orientation.Vertical if nav_pos == "MiddleSides" else Qt.Orientation.Horizontal
        self.toolbar = ToolBarWidget()
        self.nav_left = PageNavWidget(is_right=False, orientation=orientation)
        self.nav_right = PageNavWidget(is_right=True, orientation=orientation)
        self.clock_widget = ClockWidget()
        self.clock_widget.apply_settings(cfg)
        self.spotlight = SpotlightOverlay()
        self.setup_connections()
        self.setup_tray()
        self.update_widgets_theme()
        
        if settings_visible:
            self.show_settings()
            
        if prev_visible:
            try:
                self.show_widgets(animate=False)
            except Exception:
                pass

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
        def icon_path(name):
            base_dir = get_app_base_dir()
            return os.path.join(base_dir, "resources", "icons", name)
        
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(icon_path("trayicon.svg")))

        tray_menu = SystemTrayMenu(parent=self)
        
        action_settings = Action(FIF.SETTING, tr("界面和功能设置"), triggered=self.show_settings)
        action_exit = Action(FIF.POWER_BUTTON, tr("退出 Kazuha"), triggered=self.exit_application)
        
        tray_menu.addActions([
            action_settings,
            action_exit
        ])
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
    def show_settings(self):
        from ui.settings_window import SettingsWindow
        if not hasattr(self, 'settings_window') or self.settings_window is None:
            self.settings_window = SettingsWindow()
            self.settings_window.configChanged.connect(self.on_settings_changed)
            self.settings_window.checkUpdateClicked.connect(self.check_for_updates)
            self.settings_window.startUpdateClicked.connect(self.start_update_download)
            if hasattr(self.settings_window, 'set_theme'):
                self.settings_window.set_theme(self.theme_mode)
                
            # Trigger update check immediately when settings window is created
            # This ensures logs are fetched as soon as possible
            QTimer.singleShot(500, self.check_for_updates)
        
        screen = QApplication.primaryScreen().geometry()
        w = self.settings_window.width()
        h = self.settings_window.height()
        x = screen.left() + (screen.width() - w) // 2
        y = screen.top() + (screen.height() - h) // 2
        self.settings_window.move(x, y)
        
        self.settings_window.show()
        self.settings_window.activateWindow()

    def show_update_interface(self):
        self.show_settings()
        if hasattr(self, "settings_window") and self.settings_window:
            if hasattr(self, "_pending_update_info") and self._pending_update_info:
                info = self._pending_update_info
                try:
                    self.settings_window.set_update_info(info.get("version"), info.get("body"))
                except Exception:
                    pass
            try:
                self.settings_window.switch_to_update_page()
            except Exception:
                pass
        
    def on_settings_changed(self):
        self.toggle_autorun(cfg.enableStartUp.value)
        
        if self.clock_widget:
            if cfg.enableClock.value and self.widgets_visible and not self.conflicting_process_running:
                self.clock_widget.show()
            elif not cfg.enableClock.value:
                self.clock_widget.hide()
        
        self.update_widgets_theme()
        
        target_theme = cfg.themeMode.value
        if target_theme == "Auto":
            self.set_theme_auto(checked=True)
        elif target_theme == "Light":
            self.set_theme_light(checked=True)
        elif target_theme == "Dark":
            self.set_theme_dark(checked=True)
            
        if self.clock_widget and hasattr(self.clock_widget, 'apply_settings'):
            self.clock_widget.apply_settings(cfg)
        
        cfg.save()
        
        if self.widgets_visible:
            self.show_widgets(animate=cfg.enableGlobalAnimation.value)
            
    def show_about(self):
        title = tr("关于 Kazuha 演示助手")
        content = (
            tr("当前版本: {version}").format(version=VersionManager.CURRENT_VERSION)
            + "\n"
            + tr("一个现代化、高性能的 PowerPoint 演示辅助工具，用于配合课堂与演示场景。")
            + "\n\n"
            + "© 2024 Seirai Studio"
        )
                   
        w = MessageBox(title, content, self.settings_window if hasattr(self, 'settings_window') and self.settings_window else None)
        w.yesButton.setText(tr("确定"))
        w.cancelButton.hide()
        w.exec()

    def check_for_updates(self):
        self.version_manager.check_for_updates()

    def _on_toast_activated(self, activatedEventArgs):
        if activatedEventArgs.arguments == 'view_details':
            self.request_show_update.emit()

    def on_update_available(self, info):
        self._pending_update_info = info
        
        notification_shown = False
        
        # Only show notification if it is truly a NEW version (remote > local)
        local_v_code = int(self.version_manager.current_version_info.get("versionCode", 0))
        remote_v_code = int(info.get("versionCode", 0))
        
        # Strict check: must be strictly greater
        is_new_version = remote_v_code > local_v_code
        
        if is_new_version:
            if cfg.enableSystemNotification.value and HAS_WINDOWS_TOASTS:
                try:
                    current_v = self.version_manager.current_version_info.get("versionName", "Unknown")
                    new_v = info.get("version", "Unknown")
                    
                    toaster = WindowsToaster('Kazuha')
                    newToast = ToastText2()
                    newToast.SetHeadline(tr("发现新版本"))
                    newToast.SetBody(f"{current_v}->{new_v}")
                    
                    button = ToastButton(tr("点击此处以查看详细信息"), 'view_details')
                    newToast.AddAction(button)
                    
                    newToast.on_activated = self._on_toast_activated
                    toaster.show_toast(newToast)
                    notification_shown = True
                except Exception as e:
                    print(f"Toast error: {e}")
            
            if not notification_shown:
                self.show_settings()
        
        # Always update UI with info (even if it's just changelog for current version)
        if hasattr(self, "settings_window") and self.settings_window:
            if hasattr(self.settings_window, "set_update_info"):
                try:
                    # Pass is_latest flag explicitly based on VersionCode comparison
                    # This avoids string parsing issues in UI
                    self.settings_window.set_update_info(
                        info.get("version"), 
                        info.get("body"),
                        is_latest=(not is_new_version)
                    )
                except Exception:
                    # Fallback for old signature if needed (though we just updated it)
                     self.settings_window.set_update_info(info.get("version"), info.get("body"))

            if hasattr(self.settings_window, "switch_to_update_page"):
                try:
                    if not notification_shown and is_new_version:
                         self.settings_window.switch_to_update_page()
                except Exception:
                    pass

    def start_update_download(self):
        info = getattr(self, "_pending_update_info", None)
        if not info:
             info = self.version_manager.current_version_info
        
        if not info or 'assets' not in info:
             return

        asset_url = None
        for asset in info['assets']:
            if asset['name'].endswith('.exe'):
                asset_url = asset['browser_download_url']
                break
        
        if asset_url:
            self.version_manager.update_progress.connect(self.on_download_progress)
            self.version_manager.download_and_install(asset_url)
        else:
            self.show_warning(None, tr("未找到可执行的更新文件"))

    def on_download_progress(self, val):
        if hasattr(self, "settings_window") and self.settings_window:
             if hasattr(self.settings_window, "set_download_progress"):
                 self.settings_window.set_download_progress(val)

    def on_update_check_finished(self):
        if hasattr(self, "settings_window") and self.settings_window:
            if hasattr(self.settings_window, "stop_update_loading"):
                try:
                    self.settings_window.stop_update_loading()
                except Exception:
                    pass


    def show_about_dialog(self):
        version_info = self.version_manager.current_version_info
        v_name = version_info.get('versionName', 'Unknown')
        
        self.play_sound("WindowPop")
        w = MessageBox(tr("关于 PPT助手"), tr("当前版本: {version}\n\n一个专注于演示辅助的工具。\n\nDesigned by Seirai.").format(version=v_name), self.get_parent_for_dialog())
        w.exec()

    def show_warning(self, parent, text):
        w = MessageBox(tr("提示"), text, parent if parent else None)
        w.yesButton.setText(tr("确定"))
        w.cancelButton.hide()
        w.exec()

    def get_parent_for_dialog(self):
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
        
        setTheme(theme)
        if theme == Theme.LIGHT:
            v = "Light"
        elif theme == Theme.DARK:
            v = "Dark"
        else:
            v = "Auto"
        if cfg.themeMode.value != v:
            cfg.set(cfg.themeMode, v)
            cfg.save()
        
        if hasattr(self, "theme_auto_action"):
            self.theme_auto_action.setChecked(theme == Theme.AUTO)
        if hasattr(self, "theme_light_action"):
            self.theme_light_action.setChecked(theme == Theme.LIGHT)
        if hasattr(self, "theme_dark_action"):
            self.theme_dark_action.setChecked(theme == Theme.DARK)
        
        self.update_widgets_theme()

    def on_system_theme_changed(self):
        if self.theme_mode == Theme.AUTO:
            self.update_widgets_theme()

            
    def update_widgets_theme(self):
        theme = self.theme_mode
        widgets = [
            self.toolbar,
            self.nav_left,
            self.nav_right,
            self.timer_window,
            self.spotlight,
            self.clock_widget
        ]
        
        if hasattr(self, 'settings_window') and self.settings_window:
            widgets.append(self.settings_window)
        
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
                self.timer_window.pre_reminder_triggered.connect(self.speak_text)
                self.timer_window.emit_state()
        if self.timer_window.isVisible():
            self.timer_window.hide()
        else:
            self.timer_window.show()
            self.play_sound("WindowPop")
            self.timer_window.activateWindow()
            self.timer_window.raise_()
            
            pos_setting = cfg.timerPosition.value
            screen = QApplication.primaryScreen().geometry()
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
            else:
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
            self.speak_text(tr("倒计时结束"))

        if self.timer_window and self.timer_window.strong_reminder_mode:
            ring_data = self.timer_window.ringtone_combo.currentData()
            if ring_data and ring_data != "StrongTimerRing":
                self.play_sound("CustomRing", loop=True, custom_path=ring_data)
            else:
                self.play_sound("StrongTimerRing", loop=True)
                
            if not self.timer_window.isVisible():
                self.toggle_timer_window()
            else:
                self.timer_window.activateWindow()
                self.timer_window.raise_()
            
            if not self.loading_overlay:
                from ui.widgets import LoadingOverlay
                self.mask_overlay = QWidget()
                self.mask_overlay.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
                self.mask_overlay.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
                self.mask_overlay.setStyleSheet("background-color: rgba(0, 0, 0, 180);")
                screen = QApplication.primaryScreen().geometry()
                self.mask_overlay.setGeometry(screen)
                self.mask_overlay.show()
                self.timer_window.raise_()
        else:
            self.play_sound("TimerRing")
            
        if not self.clock_widget:
            return
        self.clock_widget.show_countdown_finished()
        self.clock_widget.adjustSize()
        if self.widgets_visible:
            self.adjust_positions()

    def on_timer_reset(self):
        self.sound_manager.stop("StrongTimerRing")
        self.sound_manager.stop("CustomRing")
        self.sound_manager.stop("TimerRing")
        self.play_sound("Reset")
        if hasattr(self, 'mask_overlay') and self.mask_overlay:
            self.mask_overlay.close()
            self.mask_overlay = None

    def toggle_annotation_mode(self):
        if not self.annotation_widget:
            self.annotation_widget = AnnotationWidget()
            if hasattr(self.annotation_widget, 'set_theme'):
                self.annotation_widget.set_theme(self.theme_mode)
        
        if self.annotation_widget.isVisible():
            self.annotation_widget.hide()
        else:
            self.annotation_widget.showFullScreen()
    
    def check_presentation_processes(self):
        presentation_detected = False
        
        for proc in psutil.process_iter(['name']):
            try:
                proc_name = proc.info.get('name', '') or ""
                proc_name_lower = proc_name.lower()
                
                if any(keyword in proc_name_lower for keyword in ['powerpnt', 'wpp', 'wps']):
                    presentation_detected = True
                    break
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        
        return presentation_detected
    
    def is_autorun(self):
        key = None
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
            try:
                winreg.QueryValueEx(key, "Kazuha")
                return True
            except FileNotFoundError:
                try:
                    if key is not None:
                        winreg.CloseKey(key)
                        key = None
                    self.toggle_autorun(True)
                    return True
                except Exception:
                    return False
            except OSError:
                return False
        except Exception:
            return False
        finally:
            if key is not None:
                try:
                    winreg.CloseKey(key)
                except Exception:
                    pass

    def toggle_autorun(self, checked):
        app_path = os.path.abspath(sys.argv[0])
        if app_path.endswith('.py'):
            python_exe = sys.executable.replace("python.exe", "pythonw.exe")
            if not os.path.exists(python_exe):
                python_exe = sys.executable
            cmd = f'"{python_exe}" "{app_path}"'
        else:
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

    def check_state(self):
        if hasattr(self, "ppt_worker") and self.ppt_worker:
            self.ppt_worker.request_state()

    def on_state_updated(self, state):
        self.last_state = state
        
        if state.is_running:
            if not self.widgets_visible:
                self.show_widgets(animate=cfg.enableGlobalAnimation.value)
            
            # Update navigation widgets
            if self.nav_left:
                self.nav_left.update_page(state.slide_index, state.slide_count, state.presentation_path)
            if self.nav_right:
                self.nav_right.update_page(state.slide_index, state.slide_count, state.presentation_path)
            
            # Update toolbar
            if self.toolbar:
                self.toolbar.set_pointer_mode(state.pointer_type)
                
            self.ensure_topmost()
        else:
            if self.widgets_visible:
                self.hide_widgets(animate=cfg.enableGlobalAnimation.value)
                
    def show_widgets(self, animate=True):
        self.widgets_visible = True
        
        if self.toolbar:
            self.toolbar.show()
            
        nav_pos = cfg.navPosition.value
        if nav_pos == "BottomSides":
            if self.nav_left: self.nav_left.show()
            if self.nav_right: self.nav_right.show()
        elif nav_pos == "MiddleSides":
            if self.nav_left: self.nav_left.show()
            if self.nav_right: self.nav_right.show()
            
        if cfg.enableClock.value and self.clock_widget and not self.conflicting_process_running:
            self.clock_widget.show()
            
        self.adjust_positions()

    def hide_widgets(self, animate=True):
        self.widgets_visible = False
        
        if self.toolbar:
            self.toolbar.hide()
        if self.nav_left:
            self.nav_left.hide()
        if self.nav_right:
            self.nav_right.hide()
        if self.clock_widget:
            self.clock_widget.hide()
        if self.spotlight:
            self.spotlight.hide()
        if self.annotation_widget:
            self.annotation_widget.hide()
        if self.timer_window:
            self.timer_window.hide()

    def adjust_positions(self):
        screen = QApplication.primaryScreen().geometry()
        
        # Toolbar (Top Center)
        if self.toolbar:
            w = self.toolbar.width()
            h = self.toolbar.height()
            x = screen.left() + (screen.width() - w) // 2
            y = screen.top() + cfg.screenPaddingTop.value
            self.toolbar.move(x, y)
            
        # Navigation
        nav_pos = cfg.navPosition.value
        padding_h = cfg.screenPaddingLeft.value
        padding_v = cfg.screenPaddingBottom.value
        
        if self.nav_left and self.nav_right:
            if nav_pos == "BottomSides":
                w = self.nav_left.width()
                h = self.nav_left.height()
                y = screen.top() + screen.height() - h - padding_v
                self.nav_left.move(screen.left() + padding_h, y)
                self.nav_right.move(screen.left() + screen.width() - w - padding_h, y)
            elif nav_pos == "MiddleSides":
                w = self.nav_left.width()
                h = self.nav_left.height()
                y = screen.top() + (screen.height() - h) // 2
                self.nav_left.move(screen.left() + padding_h, y)
                self.nav_right.move(screen.left() + screen.width() - w - padding_h, y)

        # Clock
        if self.clock_widget:
            clock_pos = cfg.clockPosition.value
            w = self.clock_widget.width()
            h = self.clock_widget.height()
            padding = 20
            
            if clock_pos == "TopLeft":
                x = screen.left() + padding
                y = screen.top() + padding
            elif clock_pos == "TopRight":
                x = screen.left() + screen.width() - w - padding
                y = screen.top() + padding
            elif clock_pos == "BottomLeft":
                x = screen.left() + padding
                y = screen.top() + screen.height() - h - padding
            elif clock_pos == "BottomRight":
                x = screen.left() + screen.width() - w - padding
                y = screen.top() + screen.height() - h - padding
            
            self.clock_widget.move(x, y)

    def prev_page(self):
        if self.ppt_worker:
            self.ppt_worker.prev_slide()
            
    def next_page(self):
        if self.ppt_worker:
            self.ppt_worker.next_slide()
            
    def jump_to_slide(self, index):
        if self.ppt_worker:
            self.ppt_worker.goto_slide(index)
            
    def toggle_spotlight(self):
        if not self.spotlight:
            self.spotlight = SpotlightOverlay()
            self.spotlight.set_theme(self.theme_mode)
        
        if self.spotlight.isVisible():
            self.spotlight.hide()
        else:
            self.spotlight.showFullScreen()
            
    def exit_application(self):
        if hasattr(self, "ppt_worker") and self.ppt_worker:
            self.ppt_worker.stop()
        QApplication.quit()
        
    def change_pointer_mode(self, mode):
        if self.ppt_worker:
            self.ppt_worker.set_pointer_type(mode)
            
    def change_pen_color(self, color):
        if self.ppt_worker:
            self.ppt_worker.set_pen_color(color)
            
    def clear_ink(self):
        if self.ppt_worker:
            self.ppt_worker.erase_ink()
            
    def exit_slideshow(self):
        if self.ppt_worker:
            self.ppt_worker.exit_show()

    def ensure_topmost(self):
        try:
            flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
            widgets = [
                self.toolbar, 
                self.nav_left, 
                self.nav_right, 
                self.clock_widget
            ]
            for w in widgets:
                if w and w.isVisible():
                    hwnd = int(w.winId())
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        except Exception:
            pass
