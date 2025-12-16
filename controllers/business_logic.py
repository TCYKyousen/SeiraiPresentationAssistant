import sys
import os
import winreg
import psutil

from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from qfluentwidgets import setTheme, Theme, SystemTrayMenu, Action
from ui.widgets import AnnotationWidget, TimerWindow, LoadingOverlay
from .ppt_client import PPTClient
import pythoncom
import os

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
        
        # 批注功能组件
        self.annotation_widget = None
        self.timer_window = None
        self.loading_overlay = None
        
        # UI组件引用（将在主程序中设置）
        self.toolbar = None
        self.nav_left = None
        self.nav_right = None
        self.spotlight = None
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_state)
        self.timer.start(500)
        
        self.widgets_visible = False
        self.slides_loaded = False
    
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
        # 导入图标路径函数
        import sys
        import os
        def icon_path(name):
            base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
            # 返回上级目录的icons文件夹路径
            return os.path.join(os.path.dirname(base_dir), "icons", name)
        
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(icon_path("trayicon.svg")))

        tray_menu = SystemTrayMenu(parent=self)

        # Removed Compatibility Mode Checkbox
        annotation_action = Action(self, text="独立批注", triggered=self.toggle_annotation_mode)
        timer_action = Action(self, text="计时器", triggered=self.toggle_timer_window)

        self.theme_auto_action = Action(self, text="跟随系统", checkable=True, triggered=self.set_theme_auto)
        self.theme_light_action = Action(self, text="浅色模式", checkable=True, triggered=self.set_theme_light)
        self.theme_dark_action = Action(self, text="深色模式", checkable=True, triggered=self.set_theme_dark)

        exit_action = Action(self, text="退出", triggered=self.exit_application)

        tray_menu.addAction(annotation_action)
        tray_menu.addAction(timer_action)
        tray_menu.addSeparator()
        tray_menu.addAction(self.theme_auto_action)
        tray_menu.addAction(self.theme_light_action)
        tray_menu.addAction(self.theme_dark_action)
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.set_theme_mode(self.theme_mode)
        self.tray_icon.show()

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
        setTheme(theme)
        self.save_theme_setting(theme)
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
            self.spotlight
        ]
        
        # Add optional widgets if they exist
        if hasattr(self, 'annotation_widget') and self.annotation_widget:
            widgets.append(self.annotation_widget)
            
        for widget in widgets:
            if widget and hasattr(widget, 'set_theme'):
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
        if self.timer_window.isVisible():
            self.timer_window.hide()
        else:
            self.timer_window.show()
            self.timer_window.activateWindow()
            self.timer_window.raise_()
    
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
                self.nav_left.ppt_app = self.ppt_client.app
                self.nav_right.ppt_app = self.ppt_client.app
                self.update_page_num(view)
                self.sync_state(view)
        else:
            if self.widgets_visible:
                self.hide_widgets()

    def show_widgets(self):
        self.toolbar.show()
        self.nav_left.show()
        self.nav_right.show()
        self.adjust_positions()
        self.widgets_visible = True
        
        # Trigger slide loading with overlay
        if not self.slides_loaded:
            self.start_loading_slides()

    def start_loading_slides(self):
        try:
            if not self.ppt_client.app:
                return
                
            presentation = self.ppt_client.app.ActivePresentation
            presentation_path = presentation.FullName
            
            # Reset if path changed (simple check)
            if hasattr(self, 'last_presentation_path') and self.last_presentation_path != presentation_path:
                self.slides_loaded = False
            self.last_presentation_path = presentation_path
            
            if self.slides_loaded:
                return

            import hashlib
            path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
            cache_dir = os.path.join(os.environ['APPDATA'], 'PPTAssistant', 'Cache', path_hash)
            
            if not self.loading_overlay:
                self.loading_overlay = LoadingOverlay()
            
            # Force show overlay
            self.loading_overlay.show()
            QApplication.processEvents() # Ensure it renders
            
            self.loader_thread = SlideExportThread(cache_dir)
            self.loader_thread.finished.connect(self.on_slides_loaded)
            self.loader_thread.start()
            
        except Exception as e:
            print(f"Error starting slide load: {e}")
            self.slides_loaded = True # Prevent loop

    def on_slides_loaded(self):
        if self.loading_overlay:
            self.loading_overlay.hide()
        self.slides_loaded = True

    def change_pointer_mode(self, mode):
        self.ppt_client.set_pointer_type(mode)
        # Update button state if needed, but toolbar handles its own exclusive group
        # We might want to sync back if PPT changes mode externally, but that's for sync_state

    def hide_widgets(self):
        self.toolbar.hide()
        self.nav_left.hide()
        self.nav_right.hide()
        self.widgets_visible = False

    def adjust_positions(self):
        screen = QApplication.primaryScreen().geometry()
        MARGIN = 20
        
        # Toolbar: Bottom Center
        tb_w = self.toolbar.sizeHint().width()
        tb_h = self.toolbar.sizeHint().height()
        
        self.toolbar.setGeometry(
            (screen.width() - tb_w) // 2,
            screen.height() - tb_h - MARGIN, # Flush bottom
            tb_w, tb_h
        )
        nav_w = self.nav_left.sizeHint().width()
        nav_h = self.nav_left.sizeHint().height()
        y = screen.height() - nav_h - MARGIN
        
        self.nav_left.setGeometry(
            MARGIN,
            y,
            nav_w, nav_h
        )
        
        self.nav_right.setGeometry(
            screen.width() - nav_w - MARGIN,
            y,
            nav_w, nav_h
        )

    def sync_state(self, view):
        try:
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
