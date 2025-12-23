import multiprocessing
import queue
import time
import win32com.client
import pythoncom
import win32api
import win32gui
import win32con
import os
from datetime import datetime
from dataclasses import dataclass

@dataclass
class PPTState:
    is_running: bool = False
    slide_index: int = 0
    slide_count: int = 0
    pointer_type: int = 0
    has_ink: bool = False
    presentation_path: str = ""
    hwnd: int = 0

class PPTClient:
    def __init__(self):
        self.app = None
        self.app_type = None # 'office' or 'wps'

    def connect(self):
        """尝试连接到 PowerPoint 或 WPS"""
        self.app = None
        self.app_type = None
        
        try:
            pythoncom.CoInitialize()
        except:
            pass

        # 尝试连接 Office PowerPoint
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.app_type = 'office'
            return True
        except Exception as e:
            pass

        # 尝试连接 WPS Presentation
        wps_prog_ids = ["Kwpp.Application", "Wpp.Application"]
        for prog_id in wps_prog_ids:
            try:
                self.app = win32com.client.GetActiveObject(prog_id)
                self.app_type = 'wps'
                return True
            except Exception as e:
                continue
        
        return False

    def get_slideshow_window_hwnd(self):
        if not self.app:
            if not self.connect():
                return 0

        try:
            try:
                if self.app.ActivePresentation and self.app.ActivePresentation.SlideShowWindow:
                    hwnd = int(self.app.ActivePresentation.SlideShowWindow.HWND)
                    if hwnd > 0 and win32gui.IsWindow(hwnd):
                        return hwnd
            except:
                pass
            
            count = int(self.app.SlideShowWindows.Count)
            for i in range(1, count + 1):
                try:
                    win = self.app.SlideShowWindows(i)
                    hwnd = int(win.HWND)
                    if hwnd > 0 and win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                        return hwnd
                except:
                    continue
        except Exception:
            self.app = None
            pass

        return 0

    def activate_window(self):
        hwnd = self.get_slideshow_window_hwnd()
        if hwnd:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                return True
            except:
                pass
        return False

    def get_active_view(self):
        if not self.app:
            if not self.connect():
                return None

        try:
            try:
                if self.app.ActivePresentation and self.app.ActivePresentation.SlideShowWindow:
                    view = self.app.ActivePresentation.SlideShowWindow.View
                    if view:
                        return view
            except:
                pass

            if self.app.SlideShowWindows.Count > 0:
                for i in range(1, self.app.SlideShowWindows.Count + 1):
                    try:
                        view = self.app.SlideShowWindows(i).View
                        if view:
                            return view
                    except:
                        continue
        except Exception:
            if self.connect():
                try:
                    if self.app.SlideShowWindows.Count > 0:
                        return self.app.SlideShowWindows(1).View
                except Exception:
                    pass
        return None

    def get_office_fullscreen_view(self):
        return self.get_active_view()

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

    def exit_show(self, keep_ink=None):
        view = self.get_active_view()
        if view:
            try:
                if keep_ink is not None:
                    try:
                        self.app.DisplayAlerts = 0
                    except:
                        pass
                        
                    if keep_ink:
                        try:
                            if self.app.ActivePresentation:
                                self.app.ActivePresentation.Save()
                        except:
                            pass
                            
                    view.Exit()
                    
                    try:
                        self.app.DisplayAlerts = 1
                    except:
                        pass
                else:
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
            return True

def ppt_worker_process(cmd_queue, result_queue):
    """
    Independent process function for handling PPT COM automation.
    """
    pythoncom.CoInitialize()
    client = PPTClient()
    
    # Debug log
    log_file = os.path.join(os.getenv("APPDATA"), "Kazuha", "worker_core.log")
    def worker_log(msg):
        try:
            with open(log_file, "a", encoding='utf-8') as f:
                f.write(f"{datetime.now()}: {msg}\n")
        except:
            pass
            
    worker_log("Worker process started (Core)")
    
    while True:
        try:
            task = cmd_queue.get(timeout=0.05)
        except queue.Empty:
            continue
            
        if task is None: # Poison pill
            worker_log("Exiting")
            break
            
        cmd_type = task.get('type')
        args = task.get('args', {})
        
        try:
            if cmd_type == 'check_state':
                if not client.app:
                    client.connect()
                
                state_data = {
                    'is_running': False,
                    'slide_index': 0,
                    'slide_count': 0,
                    'pointer_type': 0,
                    'has_ink': False,
                    'presentation_path': "",
                    'hwnd': 0
                }
                
                view = client.get_office_fullscreen_view()
                if view:
                    state_data['is_running'] = True
                    try:
                        state_data['slide_index'] = view.Slide.SlideIndex
                        state_data['slide_count'] = client.get_slide_count()
                        state_data['pointer_type'] = view.PointerType
                        state_data['has_ink'] = client.has_ink()
                        state_data['hwnd'] = client.get_slideshow_window_hwnd()
                        
                        try:
                            if client.app and client.app.ActivePresentation:
                                state_data['presentation_path'] = client.app.ActivePresentation.FullName
                        except:
                            pass
                    except Exception as e:
                        worker_log(f"Error reading state: {e}")
                        state_data['is_running'] = False
                
                result_queue.put(('state_update', state_data))
                
            elif cmd_type == 'next':
                client.next_slide()
            elif cmd_type == 'prev':
                client.prev_slide()
            elif cmd_type == 'goto':
                client.goto_slide(args.get('index'))
            elif cmd_type == 'set_pointer':
                client.set_pointer_type(args.get('mode'))
            elif cmd_type == 'set_pen_color':
                client.set_pen_color(args.get('color'))
            elif cmd_type == 'erase_ink':
                client.erase_ink()
            elif cmd_type == 'exit_show':
                client.exit_show(args.get('keep_ink'))
                
        except Exception as e:
            worker_log(f"Error processing {cmd_type}: {e}")
            pass
            
    pythoncom.CoUninitialize()
