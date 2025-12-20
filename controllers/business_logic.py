import sys
import os
import winreg
import psutil
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont
from qfluentwidgets import (
    setTheme,
    Theme,
    SystemTrayMenu,
    Action,
    setThemeColor,
    QConfig,
    qconfig,
    OptionsConfigItem,
    OptionsValidator,
    EnumSerializer,
    ColorConfigItem,
    ConfigItem,
    MessageBox,
    PushButton,
)
from ui.widgets import AnnotationWidget, TimerWindow, ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
from .ppt_client import PPTClient
from .version_manager import VersionManager
import pythoncom
import os
from typing import Optional

# fallback 到 Temp
CONFIG_DIR = Path(os.getenv("APPDATA", os.getenv("TEMP", "C:\\"))) / "SeiraiPPTAssistant"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = CONFIG_DIR / "config.json"


class Config(QConfig):
    windowEffect = OptionsConfigItem(
        "Theme",
        "WindowEffect",
        "Mica",
        OptionsValidator(["Mica", "Acrylic", "Normal"]),
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
    language = OptionsConfigItem(
        "General",
        "Language",
        "Simplified Chinese",
        OptionsValidator(["Simplified Chinese", "English"]),
        None,
    )


cfg = Config()
qconfig.load(str(CONFIG_PATH), cfg)

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
                        except:
                            pass
        except Exception as e:
            print(f"Slide export error: {e}")
        pythoncom.CoUninitialize()

class BusinessLogicController(QWidget):
    def __init__(self):
        super().__init__()
        self.theme_mode = self.load_theme_setting()
        setTheme(self.theme_mode)
        
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(1, 1)
        self.move(-100, -100) 
        
        self.ppt_client = PPTClient()
        self.version_manager = VersionManager()
        
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
            qconfig.save()

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
                if not self.conflicting_process_running and hasattr(self, "tray_icon"):
                    self.tray_icon.showMessage("提示", "检测到 ClassIsland/ClassWidgets，已自动隐藏时钟组件。", QSystemTrayIcon.MessageIcon.Information, 2000)
                self.conflicting_process_running = True
            else:
                if self.clock_widget and not self.clock_widget.isVisible() and self.widgets_visible:
                    self.clock_widget.show()
                self.conflicting_process_running = False
        except Exception as e:
            pass

    def set_theme_color(self, color_str):
        setThemeColor(color_str)
        self.save_theme_setting(self.theme_mode) # Re-save to persist if needed, though config handles it

    def set_font(self, font_name):
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

    def set_window_effect(self, effect_name):
        # Update config if needed
        if cfg.windowEffect.value != effect_name:
            cfg.set(cfg.windowEffect, effect_name)
        self.update_widgets_theme()

    def set_context_menu_style(self, style):
        self.set_window_effect(style)

    def apply_window_effect_to_widget(self, widget):
        if not widget:
            return
            
        effect_name = cfg.windowEffect.value
        try:
            hwnd = int(widget.winId())
            # Try to import WindowEffect
            try:
                from qfluentwidgets.window import WindowEffect
                from qfluentwidgets import isDarkTheme
            except ImportError:
                from qfluentwidgets import WindowEffect, isDarkTheme
            
            effect = WindowEffect()
            
            if effect_name == "Mica":
                effect.setMicaEffect(hwnd, isDarkTheme())
            elif effect_name == "Acrylic":
                # Acrylic requires a gradient color or hex
                color = "000000" if isDarkTheme() else "ffffff"
                effect.setAcrylicEffect(hwnd, gradientColor=color, enableShadow=True)
            else:
                effect.removeBackgroundEffect(hwnd)
                
        except Exception as e:
            # Fail silently if effect not supported
            pass

    def set_start_up(self, enabled):
        self.toggle_autorun(enabled)

    def toggle_tray_autorun(self, checked=False):
        enabled = bool(checked)
        if cfg.enableStartUp.value != enabled:
            cfg.set(cfg.enableStartUp, enabled)
            qconfig.save()
        self.set_start_up(enabled)

    def set_tray_timer_position(self, pos):
        if cfg.timerPosition.value != pos:
            cfg.set(cfg.timerPosition, pos)
            qconfig.save()
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
            qconfig.save()
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
    def set_timer_position(self, pos):
        # Position is handled when window is toggled
        pass

    def set_clock_position(self, pos):
        self.adjust_positions()

    def set_language(self, language):
        if cfg.language.value != language:
            cfg.set(cfg.language, language)
            qconfig.save()
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
    
    def setup_connections(self):
        """设置UI组件与业务逻辑之间的信号连接"""
        if self.toolbar:
            self.toolbar.request_spotlight.connect(self.toggle_spotlight)
            self.toolbar.request_pointer_mode.connect(self.change_pointer_mode)
            self.toolbar.request_pen_color.connect(self.change_pen_color)
            self.toolbar.request_clear_ink.connect(self.clear_ink)
            self.toolbar.request_exit.connect(self.exit_slideshow)
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

        annotation_action = Action(self, text="独立批注", triggered=self.toggle_annotation_mode)
        timer_action = Action(self, text="计时器", triggered=self.toggle_timer_window)

        self.theme_auto_action = Action(self, text="跟随系统", checkable=True, triggered=self.set_theme_auto)
        self.theme_light_action = Action(self, text="浅色模式", checkable=True, triggered=self.set_theme_light)
        self.theme_dark_action = Action(self, text="深色模式", checkable=True, triggered=self.set_theme_dark)

        language_menu = SystemTrayMenu("语言", parent=tray_menu)
        self.language_zh_action = Action(self, text="简体中文", checkable=True, triggered=self.set_language_zh)
        self.language_en_action = Action(self, text="English", checkable=True, triggered=self.set_language_en)
        language_menu.addAction(self.language_zh_action)
        language_menu.addAction(self.language_en_action)

        self.autorun_action = Action(self, text="开机自启动", checkable=True, triggered=self.toggle_tray_autorun)

        timer_pos_menu = SystemTrayMenu("计时器位置", parent=tray_menu)
        self.timer_pos_center_action = Action(self, text="居中", checkable=True, triggered=lambda: self.set_tray_timer_position("Center"))
        self.timer_pos_tl_action = Action(self, text="左上角", checkable=True, triggered=lambda: self.set_tray_timer_position("TopLeft"))
        self.timer_pos_tr_action = Action(self, text="右上角", checkable=True, triggered=lambda: self.set_tray_timer_position("TopRight"))
        self.timer_pos_bl_action = Action(self, text="左下角", checkable=True, triggered=lambda: self.set_tray_timer_position("BottomLeft"))
        self.timer_pos_br_action = Action(self, text="右下角", checkable=True, triggered=lambda: self.set_tray_timer_position("BottomRight"))
        timer_pos_menu.addAction(self.timer_pos_center_action)
        timer_pos_menu.addAction(self.timer_pos_tl_action)
        timer_pos_menu.addAction(self.timer_pos_tr_action)
        timer_pos_menu.addAction(self.timer_pos_bl_action)
        timer_pos_menu.addAction(self.timer_pos_br_action)

        clock_pos_menu = SystemTrayMenu("时钟位置", parent=tray_menu)
        self.clock_pos_tl_action = Action(self, text="左上角", checkable=True, triggered=lambda: self.set_tray_clock_position("TopLeft"))
        self.clock_pos_tr_action = Action(self, text="右上角", checkable=True, triggered=lambda: self.set_tray_clock_position("TopRight"))
        self.clock_pos_bl_action = Action(self, text="左下角", checkable=True, triggered=lambda: self.set_tray_clock_position("BottomLeft"))
        self.clock_pos_br_action = Action(self, text="右下角", checkable=True, triggered=lambda: self.set_tray_clock_position("BottomRight"))
        clock_pos_menu.addAction(self.clock_pos_tl_action)
        clock_pos_menu.addAction(self.clock_pos_tr_action)
        clock_pos_menu.addAction(self.clock_pos_bl_action)
        clock_pos_menu.addAction(self.clock_pos_br_action)

        restart_action = Action(self, text="重启", triggered=self.restart_application)
        exit_action = Action(self, text="退出", triggered=self.exit_application)

        tray_menu.addAction(annotation_action)
        tray_menu.addAction(timer_action)
        tray_menu.addSeparator()
        tray_menu.addAction(self.theme_auto_action)
        tray_menu.addAction(self.theme_light_action)
        tray_menu.addAction(self.theme_dark_action)
        tray_menu.addSeparator()
        tray_menu.addMenu(language_menu)
        tray_menu.addMenu(timer_pos_menu)
        tray_menu.addMenu(clock_pos_menu)
        tray_menu.addAction(self.autorun_action)
        tray_menu.addSeparator()

        about_menu = SystemTrayMenu("关于", parent=tray_menu)
        check_update_action = Action(self, text="检查更新", triggered=self.check_for_updates)
        about_action = Action(self, text="版本信息", triggered=self.show_about_dialog)
        
        about_menu.addAction(check_update_action)
        about_menu.addAction(about_action)
        tray_menu.addMenu(about_menu)
        
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)

        current_language = cfg.language.value
        if current_language == "English":
            self.language_en_action.setChecked(True)
            self.language_zh_action.setChecked(False)
        else:
            self.language_zh_action.setChecked(True)
            self.language_en_action.setChecked(False)

        if cfg.enableStartUp.value:
            self.autorun_action.setChecked(True)

        timer_pos = cfg.timerPosition.value
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
        
        w = MessageBox("关于 PPT助手", f"当前版本: {v_name}\n\n一个专注于演示辅助的工具。\n\nDesigned by Seirai.", self.get_parent_for_dialog())
        w.exec()

    def get_parent_for_dialog(self):
        # MessageBox needs a parent window, but we are essentially hidden/tray based.
        # We can use a dummy widget or None (if supported by qfluentwidgets)
        # Using None might create a standalone window.
        return None

    def closeEvent(self, event):
        self.timer.stop()
        if self.ppt_client.app:
            try:
                # self.ppt_client.app.Quit() # Should not quit PPT app on helper exit?
                pass
            except:
                pass
        # 清理批注功能
        if self.annotation_widget:
            self.annotation_widget.close()
        event.accept()
    
    def load_theme_setting(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant", 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, "ThemeMode")
            winreg.CloseKey(key)
            if isinstance(value, str):
                v = value.lower()
                if v == "light":
                    return Theme.LIGHT
                if v == "dark":
                    return Theme.DARK
                if v == "auto":
                    return Theme.AUTO
        except WindowsError:
            pass
        return Theme.AUTO
    
    def save_theme_setting(self, theme):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant", 0, winreg.KEY_ALL_ACCESS)
        except WindowsError:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant")
        
        try:
            value = getattr(theme, "value", None)
            if not isinstance(value, str):
                value = str(theme)
            winreg.SetValueEx(key, "ThemeMode", 0, winreg.REG_SZ, value)
            winreg.CloseKey(key)
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
        
        # Add optional widgets if they exist
        if hasattr(self, 'annotation_widget') and self.annotation_widget:
            widgets.append(self.annotation_widget)
            
        for widget in widgets:
            if widget:
                if hasattr(widget, 'set_theme'):
                    widget.set_theme(theme)
                
                # Apply window effect only to UI panels, not full-screen overlays
                # AnnotationWidget and SpotlightOverlay should remain transparent/dimmed
                if widget not in [self.spotlight, self.annotation_widget]:
                    self.apply_window_effect_to_widget(widget)
    
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
                self.timer_window.emit_state()
        if self.timer_window.isVisible():
            self.timer_window.hide()
        else:
            self.timer_window.show()
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
        if not self.clock_widget:
            return
        self.clock_widget.show_countdown_finished()
        self.clock_widget.adjustSize()
        if self.widgets_visible:
            self.adjust_positions()
    
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
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, "SeiraiPPTAssistant")
            winreg.CloseKey(key)
            return True
        except WindowsError:
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
                winreg.SetValueEx(key, "SeiraiPPTAssistant", 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, "SeiraiPPTAssistant")
                except WindowsError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error setting autorun: {e}")

    def check_state(self):
        # 仅在有活动放映视图时显示控件
        view = None
        has_process = self.check_presentation_processes()
        
        if has_process:
            view = self.ppt_client.get_active_view()
            
        if view:
            if not self.widgets_visible:
                self.show_widgets()
            
            if self.ppt_client.app:
                assert self.nav_left is not None
                assert self.nav_right is not None
                self.nav_left.ppt_app = self.ppt_client.app
                self.nav_right.ppt_app = self.ppt_client.app
                self.update_page_num(view)
                self.sync_state(view)
        else:
            if self.widgets_visible:
                self.hide_widgets()

    def show_widgets(self):
        assert self.toolbar is not None
        assert self.nav_left is not None
        assert self.nav_right is not None
        self.toolbar.show()
        self.nav_left.show()
        self.nav_right.show()
        if self.clock_widget and not self.conflicting_process_running:
            self.clock_widget.show()
        self.adjust_positions()
        self.widgets_visible = True
        
        if not self.slides_loaded:
            self.start_loading_slides()

    def start_loading_slides(self):
        try:
            if not self.ppt_client.app:
                return
                
            presentation = self.ppt_client.app.ActivePresentation
            presentation_path = presentation.FullName
            
            if hasattr(self, 'last_presentation_path') and self.last_presentation_path != presentation_path:
                self.slides_loaded = False
            self.last_presentation_path = presentation_path
            
            if self.slides_loaded:
                return

            import hashlib
            path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
            cache_dir = os.path.join(os.environ['APPDATA'], 'PPTAssistant', 'Cache', path_hash)
            
            self.loader_thread = SlideExportThread(cache_dir)
            self.loader_thread.finished.connect(self.on_slides_loaded)
            self.loader_thread.start()
            
        except Exception as e:
            print(f"Error starting slide load: {e}")
            self.slides_loaded = True

    def on_slides_loaded(self):
        self.slides_loaded = True

    def change_pointer_mode(self, mode):
        self.ppt_client.set_pointer_type(mode)

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

    def adjust_positions(self):
        assert self.toolbar is not None
        assert self.nav_left is not None
        assert self.nav_right is not None
        if self.toolbar and self.toolbar.screen():
            screen = self.toolbar.screen().geometry() # type: ignore
        else:
            screen = QApplication.primaryScreen().geometry() # type: ignore
        MARGIN = 20
        left = screen.left()
        top = screen.top()
        width = screen.width()
        height = screen.height()
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
        nav_y = top + height - nav_h - MARGIN
        self.nav_left.setGeometry(left + MARGIN, nav_y, nav_w, nav_h)
        self.nav_right.setGeometry(right - nav_w - MARGIN, nav_y, nav_w, nav_h)
        
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
            elif pos_setting == "BottomRight":
                x = right - cw - MARGIN
                y = bottom - ch - MARGIN
            else: # Default TopRight
                x = right - cw - MARGIN
                y = top + MARGIN
                
            # Boundary checks (keep within screen)
            if x < left + MARGIN: x = left + MARGIN
            if x > right - cw - MARGIN: x = right - cw - MARGIN
            if y < top + MARGIN: y = top + MARGIN
            if y > bottom - ch - MARGIN: y = bottom - ch - MARGIN
                
            self.clock_widget.setGeometry(x, y, cw, ch)

    def sync_state(self, view):
        try:
            assert self.toolbar is not None
            pt = view.PointerType
            if pt == 1:
                self.toolbar.btn_arrow.setChecked(True)
            elif pt == 2:
                self.toolbar.btn_pen.setChecked(True)
            elif pt == 5: # Eraser
                self.toolbar.btn_eraser.setChecked(True)
        except:
            pass

    def update_page_num(self, view):
        try:
            assert self.nav_left is not None
            assert self.nav_right is not None
            current = view.Slide.SlideIndex
            total = self.ppt_client.get_slide_count()
            self.nav_left.update_page(current, total)
            self.nav_right.update_page(current, total)
        except:
            pass

    def go_prev(self):
        self.ppt_client.prev_slide()

    def go_next(self):
        self.ppt_client.next_slide()
                
    def next_page(self):
        """下一页"""
        self.go_next()
        
    def prev_page(self):
        """上一页"""
        self.go_prev()
                
    def jump_to_slide(self, index):
        self.ppt_client.goto_slide(index)

    def set_pointer(self, type_id):
        # 强制使用COM接口
        if type_id == 5:
            # Check for ink but DO NOT BLOCK
            if not self.ppt_client.has_ink():
                self.show_warning(None, "当前页没有笔迹")
        
        self.ppt_client.set_pointer_type(type_id)
    
    def set_pen_color(self, color):
        self.ppt_client.set_pen_color(color)
                
    def change_pen_color(self, color):
        """更改笔颜色"""
        self.set_pen_color(color)
                
    def clear_ink(self):
        if not self.ppt_client.has_ink():
            self.show_warning(None, "当前页没有笔迹")
        self.ppt_client.erase_ink()
                
    def toggle_spotlight(self):
        assert self.spotlight is not None
        if self.spotlight.isVisible():
            self.spotlight.hide()
        else:
            self.spotlight.showFullScreen()
            

    def exit_slideshow(self):
        self.ppt_client.exit_show()
                
    def exit_application(self):
        """退出应用程序"""
        self.exit_slideshow()
        app = QApplication.instance()
        if app is not None:
            app.quit()

    def show_warning(self, target, message):
        title = "PPT助手提示"
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Warning, 2000)
