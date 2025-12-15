import sys
import os
import winreg
import win32com.client
import win32gui
import pyautogui
import psutil

from PyQt6.QtWidgets import QApplication, QWidget, QSystemTrayIcon, QLabel
from PyQt6.QtCore import QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QIcon
from qfluentwidgets import setTheme, Theme, Flyout, FlyoutView, FlyoutAnimationType, SystemTrayMenu, Action
from qfluentwidgets.components.material import AcrylicFlyout
from ui.widgets import AnnotationWidget, CompatibilityAnnotationWidget, TimerWindow

class BusinessLogicController(QWidget):
    def __init__(self):
        super().__init__()
        setTheme(Theme.DARK) 
        
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(1, 1)
        self.move(-100, -100) 
        
        self.ppt_app = None
        self.current_view = None
        # 兼容模式设置
        self.compatibility_mode = self.load_compatibility_mode_setting()
        
        # 批注功能组件
        self.annotation_widget = None
        self.timer_window = None
        
        # UI组件引用（将在主程序中设置）
        self.toolbar = None
        self.nav_left = None
        self.nav_right = None
        self.spotlight = None
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_state)
        self.timer.start(500)
        
        self.widgets_visible = False
    
    def setup_connections(self):
        """设置UI组件与业务逻辑之间的信号连接"""
        if self.toolbar:
            self.toolbar.request_spotlight.connect(self.toggle_spotlight)
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
        self.tray_icon.setIcon(QIcon(icon_path("Mouse.svg")))

        tray_menu = SystemTrayMenu(parent=self)

        self.compatibility_checkbox = Action(self, text="com接口模式", checkable=True, checked=not self.compatibility_mode, triggered=self.toggle_compatibility_mode)
        annotation_action = Action(self, text="独立批注", triggered=self.toggle_annotation_mode)
        timer_action = Action(self, text="计时器", triggered=self.toggle_timer_window)
        exit_action = Action(self, text="退出", triggered=self.exit_application)

        tray_menu.addAction(self.compatibility_checkbox)
        tray_menu.addAction(annotation_action)
        tray_menu.addAction(timer_action)
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def closeEvent(self, event):
        self.timer.stop()
        if self.ppt_app:
            try:
                self.ppt_app.Quit()
            except:
                pass
        # 清理批注功能
        if self.annotation_widget:
            self.annotation_widget.close()
        event.accept()

    def load_compatibility_mode_setting(self):
        """加载兼容模式设置"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant", 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, "CompatibilityMode")
            winreg.CloseKey(key)
            return bool(value)
        except WindowsError:
            return False  # 默认关闭兼容模式
    
    def save_compatibility_mode_setting(self, enabled):
        """保存兼容模式设置"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant", 0, winreg.KEY_ALL_ACCESS)
        except WindowsError:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\SeiraiPPTAssistant")
        
        try:
            winreg.SetValueEx(key, "CompatibilityMode", 0, winreg.REG_DWORD, int(enabled))
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error saving compatibility mode setting: {e}")
    
    def toggle_compatibility_mode(self):
        checked = False
        if hasattr(self, "compatibility_checkbox") and self.compatibility_checkbox is not None:
            checked = self.compatibility_checkbox.isChecked()
        self.compatibility_mode = not checked
        self.save_compatibility_mode_setting(self.compatibility_mode)
    
    def toggle_timer_window(self):
        if not self.timer_window:
            self.timer_window = TimerWindow()
        if self.timer_window.isVisible():
            self.timer_window.hide()
        else:
            self.timer_window.show()
            self.timer_window.activateWindow()
            self.timer_window.raise_()
    
    def toggle_annotation_mode(self):
        """切换独立批注模式"""
        if self.compatibility_mode:
            # 在兼容模式下使用CompatibilityAnnotationWidget
            if not hasattr(self, 'compatibility_annotation_widget') or self.compatibility_annotation_widget is None:
                self.compatibility_annotation_widget = CompatibilityAnnotationWidget()
            
            if self.compatibility_annotation_widget.isVisible():
                self.compatibility_annotation_widget.hide()
            else:
                self.compatibility_annotation_widget.showFullScreen()
        else:
            # 在标准模式下使用原来的AnnotationWidget
            if not self.annotation_widget:
                self.annotation_widget = AnnotationWidget()
            
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
                # 增加空值检查后再调用lower()
                proc_name_lower = proc_name.lower()
                
                # 检查是否为PowerPoint或WPS演示相关进程
                if any(keyword in proc_name_lower for keyword in ['powerpnt', 'wpp', 'wps']):
                    presentation_detected = True
                    break
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        
        return presentation_detected
    
    def find_presentation_window(self):
        """查找WPS或PowerPoint的放映窗口"""
        windows = []
        title_keywords = ['wps', 'powerpoint', '演示', '幻灯片', 'slide show', 'slideshow']
        class_keywords = ['wpp', 'powerpnt', 'presentation', 'screenclass', 'ppt', 'kwpp']
        
        def enum_windows_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                window_text = (win32gui.GetWindowText(hwnd) or "").lower()
                class_name = (win32gui.GetClassName(hwnd) or "").lower()
                if (any(keyword in window_text for keyword in title_keywords) or 
                    any(keyword in class_name for keyword in class_keywords)):
                    extra.append(hwnd)
            return True
        
        win32gui.EnumWindows(enum_windows_callback, windows)
        return windows[0] if windows else None
    
    def simulate_up_key(self):
        """模拟上一页按键"""
        # 查找并激活演示窗口
        hwnd = self.find_presentation_window()
        if hwnd:
            # 激活窗口
            import win32con
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            
        # 模拟按下向上键
        pyautogui.press('up')
    
    def simulate_down_key(self):
        """模拟下一页按键"""
        # 查找并激活演示窗口
        hwnd = self.find_presentation_window()
        if hwnd:
            # 激活窗口
            import win32con
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            
        # 模拟按下向下键
        pyautogui.press('down')
    
    def simulate_esc_key(self):
        """模拟ESC键退出演示"""
        # 查找并激活演示窗口
        hwnd = self.find_presentation_window()
        if hwnd:
            # 激活窗口
            import win32con
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            
        # 模拟按下ESC键
        pyautogui.press('esc')
    
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

    def get_ppt_view(self):#获取ppt全部页面
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            if self.ppt_app.SlideShowWindows.Count > 0:
                return self.ppt_app.SlideShowWindows(1).View
            else:
                return None
        except Exception:
            return None

    def check_state(self):
        has_presentation = self.check_presentation_processes()
        if has_presentation:
            if not self.widgets_visible:
                self.show_widgets()
        else:
            if self.widgets_visible:
                self.hide_widgets()
        
        if self.compatibility_mode:
            return
        
        view = self.get_ppt_view()
        if view and self.ppt_app:
            self.nav_left.ppt_app = self.ppt_app
            self.nav_right.ppt_app = self.ppt_app
            self.sync_state(view)
            self.update_page_num(view)

    def show_widgets(self):
        self.toolbar.show()
        self.nav_left.show()
        self.nav_right.show()
        self.adjust_positions()
        self.widgets_visible = True

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
            total = self.ppt_app.ActivePresentation.Slides.Count
            self.nav_left.update_page(current, total)
            self.nav_right.update_page(current, total)
        except:
            pass

    def go_prev(self):
        # 如果启用了兼容模式，则使用pyautogui模拟按键
        if self.compatibility_mode:
            # 检查是否有演示进程在运行
            if self.check_presentation_processes():
                self.simulate_up_key()
            return
        
        view = self.get_ppt_view()
        if view:
            try:
                view.Previous()
            except:
                pass
        else:
            if self.check_presentation_processes():
                self.simulate_up_key()

    def go_next(self):
        # 如果启用了兼容模式，则使用pyautogui模拟按键
        if self.compatibility_mode:
            # 检查是否有演示进程在运行
            if self.check_presentation_processes():
                self.simulate_down_key()
            return
        
        view = self.get_ppt_view()
        if view:
            try:
                view.Next()
            except:
                pass
        else:
            if self.check_presentation_processes():
                self.simulate_down_key()
                
    def next_page(self):
        """下一页"""
        self.go_next()
        
    def prev_page(self):
        """上一页"""
        self.go_prev()
                
    def jump_to_slide(self, index):
        view = self.get_ppt_view()
        if view:
            try:
                view.GotoSlide(index)
            except:
                pass

    def set_pointer(self, type_id):
        # 如果在兼容模式下，控制CompatibilityAnnotationWidget
        if self.compatibility_mode and hasattr(self, 'compatibility_annotation_widget') and self.compatibility_annotation_widget:
            try:
                if type_id == 2:  # 笔模式
                    self.compatibility_annotation_widget.set_pen_mode()
                elif type_id == 5:  # 橡皮擦模式
                    self.compatibility_annotation_widget.set_eraser_mode()
                return
            except:
                pass
        
        # 否则使用COM接口
        view = self.get_ppt_view()
        if view:
            try:
                # If switching to eraser (5)
                if type_id == 5:
                    # Check for ink but DO NOT BLOCK
                    if not self.has_ink():
                        self.show_warning(None, "当前页没有笔迹")
                
                view.PointerType = type_id
                self.activate_ppt_window()
            except:
                pass
    
    def set_pen_color(self, color):
        # 如果在兼容模式下，控制CompatibilityAnnotationWidget
        if self.compatibility_mode and hasattr(self, 'compatibility_annotation_widget') and self.compatibility_annotation_widget:
            try:
                # 将RGB整数转换为Qt颜色
                from PyQt6.QtCore import Qt
                if color == 255:  # 红色 (R=255, G=0, B=0)
                    qt_color = Qt.GlobalColor.red
                elif color == 65280:  # 绿色 (R=0, G=255, B=0)
                    qt_color = Qt.GlobalColor.green
                elif color == 16711680:  # 蓝色 (R=0, G=0, B=255)
                    qt_color = Qt.GlobalColor.blue
                else:
                    # 默认使用红色
                    qt_color = Qt.GlobalColor.red
                
                self.compatibility_annotation_widget.set_pen_color(qt_color)
                return
            except:
                pass
        
        # 否则使用COM接口
        view = self.get_ppt_view()
        if view:
            try:
                view.PointerType = 2 # Switch to pen first
                view.PointerColor.RGB = color
                self.activate_ppt_window()
            except:
                pass
                
    def change_pen_color(self, color):
        """更改笔颜色"""
        self.set_pen_color(color)

    def activate_ppt_window(self):
        try:
            # Try to get the window handle of the slide show
            hwnd = self.ppt_app.SlideShowWindows(1).HWND
            # Force bring to foreground
            win32gui.SetForegroundWindow(hwnd)
        except:
            pass
                
    def clear_ink(self):
        # 如果在兼容模式下，控制CompatibilityAnnotationWidget
        if self.compatibility_mode and hasattr(self, 'compatibility_annotation_widget') and self.compatibility_annotation_widget:
            try:
                self.compatibility_annotation_widget.clear_all()
                return
            except:
                pass
        
        # 否则使用COM接口
        view = self.get_ppt_view()
        if view:
            try:
                if not self.has_ink():
                    self.show_warning(None, "当前页没有笔迹")
                view.EraseDrawing()
            except:
                pass
                
    def toggle_spotlight(self):
        if self.spotlight.isVisible():
            self.spotlight.hide()
        else:
            self.spotlight.showFullScreen()
            
    def exit_slideshow(self):
        # 如果启用了兼容模式，则使用pyautogui模拟ESC键退出演示
        if self.compatibility_mode:
            # 检查是否有演示进程在运行
            if self.check_presentation_processes():
                self.simulate_esc_key()
            return
        
        view = self.get_ppt_view()
        if view:
            try:
                view.Exit()
            except:
                pass
        else:
            if self.check_presentation_processes():
                self.simulate_esc_key()
                
    def exit_application(self):
        """退出应用程序"""
        self.exit_slideshow()
        app = QApplication.instance()
        if app is not None:
            app.quit()

    def has_ink(self):
        try:
            view = self.get_ppt_view()
            if not view:
                return False
            slide = view.Slide
            if slide.Shapes.Count == 0:
                return False
            for shape in slide.Shapes:
                if shape.Type == 22: # msoInk
                    return True
            return False
        except:
            # If any error occurs during check (e.g. COM busy), 
            # fail safe to True to allow eraser usage (don't block user)
            return True

    def show_warning(self, target, message):
        title = "PPT助手提示"
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Warning, 2000)
