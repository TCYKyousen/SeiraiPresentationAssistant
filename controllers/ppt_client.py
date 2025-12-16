import win32com.client
import pythoncom
import win32gui
import win32con

class PPTClient:
    def __init__(self):
        self.app = None
        self.app_type = None # 'office' or 'wps'

    def connect(self):
        """尝试连接到 PowerPoint 或 WPS"""
        self.app = None
        self.app_type = None
        
        # 尝试连接 Office PowerPoint
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.app_type = 'office'
            return True
        except Exception:
            pass

        # 尝试连接 WPS Presentation
        # WPS 有多种 ProgID: Kwpp.Application, Wpp.Application
        wps_prog_ids = ["Kwpp.Application", "Wpp.Application"]
        for prog_id in wps_prog_ids:
            try:
                self.app = win32com.client.GetActiveObject(prog_id)
                self.app_type = 'wps'
                return True
            except Exception:
                continue
                
        return False

    def activate_window(self):
        try:
            if self.app and self.app.SlideShowWindows.Count > 0:
                hwnd = self.app.SlideShowWindows(1).HWND
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
        except:
            pass
        return False

    def get_active_view(self):
        """获取当前放映视图"""
        if not self.app:
            if not self.connect():
                return None

        try:
            if self.app.SlideShowWindows.Count > 0:
                # 无论是 Office 还是 WPS，通常 SlideShowWindows(1).View 都是可用的
                return self.app.SlideShowWindows(1).View
        except Exception:
            # 连接可能已断开，尝试重新连接
            if self.connect():
                try:
                    if self.app.SlideShowWindows.Count > 0:
                        return self.app.SlideShowWindows(1).View
                except Exception:
                    pass
        return None

    def get_slide_count(self):
        try:
            if self.app and self.app.ActivePresentation:
                return self.app.ActivePresentation.Slides.Count
        except:
            pass
        return 0

    def get_current_slide_index(self):
        view = self.get_active_view()
        if view:
            try:
                return view.Slide.SlideIndex
            except:
                pass
        return 0

    def next_slide(self):
        view = self.get_active_view()
        if view:
            try:
                view.Next()
                return True
            except:
                pass
        return False

    def prev_slide(self):
        view = self.get_active_view()
        if view:
            try:
                view.Previous()
                return True
            except:
                pass
        return False

    def goto_slide(self, index):
        view = self.get_active_view()
        if view:
            try:
                view.GotoSlide(index)
                return True
            except:
                pass
        return False
        
    def get_pointer_type(self):
        view = self.get_active_view()
        if view:
            try:
                return view.PointerType
            except:
                pass
        return 0

    def set_pointer_type(self, type_id):
        view = self.get_active_view()
        if view:
            try:
                view.PointerType = type_id
                self.activate_window()
                return True
            except:
                pass
        return False

    def set_pen_color(self, rgb_color):
        view = self.get_active_view()
        if view:
            try:
                # Ensure pen mode is active
                view.PointerType = 2 
                view.PointerColor.RGB = rgb_color
                self.activate_window()
                return True
            except:
                pass
        return False

    def erase_ink(self):
        view = self.get_active_view()
        if view:
            try:
                view.EraseDrawing()
                return True
            except:
                pass
        return False

    def exit_show(self):
        view = self.get_active_view()
        if view:
            try:
                view.Exit()
                return True
            except:
                pass
        return False

    def has_ink(self):
        try:
            view = self.get_active_view()
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
            return True # Fail safe
