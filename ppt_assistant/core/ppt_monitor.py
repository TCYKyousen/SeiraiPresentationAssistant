import win32com.client
import pythoncom
from PySide6.QtCore import QObject, Signal, QThread, QTimer, QPoint
from PySide6.QtGui import QGuiApplication
import time
from ppt_assistant.core.config import cfg

try:
    import win32gui
except ImportError:
    win32gui = None

class PPTMonitor(QObject):
    """
    Monitors PowerPoint state and emits signals when slideshow starts/ends.
    Also provides methods to control slides.
    """
    slideshow_started = Signal()
    slideshow_ended = Signal()
    slide_changed = Signal(int, int) # current, total
    window_geometry_changed = Signal(int, int, int, int) # left, top, width, height
    
    def __init__(self):
        super().__init__()
        self.ppt_app = None
        self.wps_app = None
        self._running = False
        self._monitoring_active = False
        self._current_slide = 0
        self._total_slides = 0
        self._video_position = 0.0
        self._video_length = 0.0
        self._last_win_rect = (0, 0, 0, 0)
        self._active_kind = None

    def get_page_info(self):
        """Returns (current_slide, total_slides)."""
        return self._current_slide, self._total_slides

    def start_monitoring(self):
        self._monitoring_active = True
        self._check_timer = QTimer(self)
        self._check_timer.timeout.connect(self._check_ppt_state)
        self._check_timer.start(200)  # Check every 200ms

    def stop_monitoring(self):
        self._monitoring_active = False
        if hasattr(self, '_check_timer'):
            self._check_timer.stop()

    def _check_ppt_state(self):
        try:
            try:
                self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                self.ppt_app = None
                self._check_wps_com_state()
                return

            if self.ppt_app.SlideShowWindows.Count > 0:
                try:
                    state = self.ppt_app.SlideShowWindows(1).View.State
                    if state in [1, 2]:
                        if not self._running:
                            self._running = True
                            self._active_kind = "ppt"
                            self.slideshow_started.emit()
                        
                        try:
                            ss_win = self.ppt_app.SlideShowWindows(1)
                            view = ss_win.View
                            current = view.Slide.SlideIndex
                            presentation = self.ppt_app.SlideShowWindows(1).Presentation
                            total = presentation.Slides.Count
                            
                            if current != self._current_slide or total != self._total_slides:
                                self._current_slide = current
                                self._total_slides = total
                                self.slide_changed.emit(current, total)
                            
                            self._update_window_rect(ss_win)
                            self._update_video_state()
                        except Exception:
                            pass
                    else:
                        if state == 5 and self._running:
                             self._running = False
                             if self._active_kind == "ppt":
                                 self._active_kind = None
                             self.slideshow_ended.emit()
                except Exception:
                    pass
            else:
                if self._running:
                    self._running = False
                    if self._active_kind == "ppt":
                        self._active_kind = None
                    self.slideshow_ended.emit()

        except Exception as e:
            pass

        if not self._running:
            self._check_wps_com_state()

    def _check_wps_com_state(self):
        try:
            self.wps_app = win32com.client.GetActiveObject("KWPP.Application")
        except Exception:
            self.wps_app = None
            if self._running and self._active_kind == "wps":
                self._running = False
                self._active_kind = None
                self.slideshow_ended.emit()
            return False

        try:
            if self.wps_app.SlideShowWindows.Count > 0:
                try:
                    view = self.wps_app.SlideShowWindows(1).View
                    state = getattr(view, "State", 1)
                    if state in [1, 2]:
                        if not self._running:
                            self._running = True
                            self._active_kind = "wps"
                            self.slideshow_started.emit()

                        try:
                            ss_win = self.wps_app.SlideShowWindows(1)
                            view = ss_win.View
                            current = view.Slide.SlideIndex
                            presentation = ss_win.Presentation
                            total = presentation.Slides.Count

                            if current != self._current_slide or total != self._total_slides:
                                self._current_slide = current
                                self._total_slides = total
                                self.slide_changed.emit(current, total)

                            self._update_window_rect(ss_win)
                            self._update_video_state()
                        except Exception:
                            pass
                    else:
                        if state == 5 and self._running and self._active_kind == "wps":
                             self._running = False
                             self._active_kind = None
                             self.slideshow_ended.emit()
                except Exception:
                    pass
            else:
                if self._running and self._active_kind == "wps":
                    self._running = False
                    self._active_kind = None
                    self.slideshow_ended.emit()
        except Exception:
            return False

        return self._running and self._active_kind == "wps"

    def _get_active_app(self):
        if self._active_kind == "ppt" and self.ppt_app is not None:
            return self.ppt_app
        if self._active_kind == "wps" and self.wps_app is not None:
            return self.wps_app
        if self.ppt_app is not None:
            return self.ppt_app
        if self.wps_app is not None:
            return self.wps_app
        return None

    def go_next(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.Next()
        except Exception as e:
            print(f"Error going next: {e}")

    def go_previous(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.Previous()
        except Exception as e:
            print(f"Error going previous: {e}")

    def end_show(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                if cfg.autoHandleInk.value and self._active_kind == "ppt":
                    try:
                        app.DisplayAlerts = 1
                    except Exception as e:
                        print(f"Error setting DisplayAlerts: {e}")

                app.SlideShowWindows(1).View.Exit()

                if cfg.autoHandleInk.value and self._active_kind == "ppt":
                    try:
                        app.DisplayAlerts = -1
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error ending show: {e}")

    def set_pointer_type(self, pointer_type):
        try:
            pythoncom.CoInitialize()
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                ss_win = app.SlideShowWindows(1)
                ss_win.View.PointerType = pointer_type
                ss_win.Activate() 
        except Exception as e:
            print(f"Error setting pointer type: {e}")

    def set_pen_color(self, r, g, b):
        try:
            pythoncom.CoInitialize()
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                view = app.SlideShowWindows(1).View
                rgb_value = r + (g << 8) + (b << 16)
                view.PointerColor.RGB = rgb_value
        except Exception as e:
            print(f"Error setting pen color: {e}")

    def get_total_slides(self):
        return self._total_slides

    def go_to_slide(self, index):
        try:
            pythoncom.CoInitialize()
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                view = app.SlideShowWindows(1).View
                view.GotoSlide(index)
        except Exception as e:
            print(f"Error going to slide: {e}")

    def export_slide_thumbnail(self, index, path):
        try:
            pythoncom.CoInitialize()
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                presentation = app.SlideShowWindows(1).Presentation
                slides = presentation.Slides
                if 1 <= index <= slides.Count:
                    slide = slides(index)
                    slide.Export(path, "PNG", 320, 180)
        except Exception as e:
            print(f"Error exporting slide thumbnail: {e}")

    def _update_window_rect(self, ss_win):
        try:
            hwnd = getattr(ss_win, "HWND", None)
            l_left, l_top, l_width, l_height = 0, 0, 0, 0
            
            if hwnd and win32gui is not None:
                # 1. Get Physical Coordinates from Windows API
                # GetWindowRect returns physical pixels (including shadow if DWM is on, but usually close enough for fullscreen)
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                
                p_x = left
                p_y = top
                p_w = right - left
                p_h = bottom - top

                # 2. Convert Physical to Logical using Qt's screen awareness
                # Find which screen the window center is on
                center_x = p_x + p_w // 2
                center_y = p_y + p_h // 2
                
                screen = QGuiApplication.screenAt(QPoint(center_x, center_y))
                if not screen:
                    screen = QGuiApplication.primaryScreen()
                
                if screen:
                    dpr = screen.devicePixelRatio()
                    # Convert physical to logical
                    l_left = int(p_x / dpr)
                    l_top = int(p_y / dpr)
                    l_width = int(p_w / dpr)
                    l_height = int(p_h / dpr)
                else:
                    # Fallback if no screen found (rare)
                    l_left, l_top, l_width, l_height = p_x, p_y, p_w, p_h
                    
                final_rect = (l_left, l_top, l_width, l_height)
            else:
                # Fallback to COM properties (usually points, which map to logical pixels mostly)
                l_left = int(getattr(ss_win, "Left", 0))
                l_top = int(getattr(ss_win, "Top", 0))
                l_width = int(getattr(ss_win, "Width", 0))
                l_height = int(getattr(ss_win, "Height", 0))
                final_rect = (l_left, l_top, l_width, l_height)

            if final_rect != self._last_win_rect and l_width > 0 and l_height > 0:
                self._last_win_rect = final_rect
                self.window_geometry_changed.emit(l_left, l_top, l_width, l_height)
        except Exception:
            pass

    def _update_video_state(self):
        self._video_position = 0.0
        self._video_length = 0.0
        try:
            pythoncom.CoInitialize()
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                view = app.SlideShowWindows(1).View
                slide = view.Slide
                shapes = slide.Shapes
                count = shapes.Count
                for i in range(1, count + 1):
                    shape = shapes.Item(i)
                    media = getattr(shape, "MediaFormat", None)
                    if media is None:
                        continue
                    length = getattr(media, "Length", None)
                    position = getattr(media, "Position", None)
                    if length and float(length) > 0:
                        self._video_length = float(length)
                        self._video_position = float(position or 0.0)
                        break
        except Exception:
            pass

    def get_video_progress(self):
        if self._video_length and self._video_length > 0:
            ratio = self._video_position / self._video_length
            if ratio < 0:
                ratio = 0.0
            if ratio > 1:
                ratio = 1.0
            return ratio, self._video_position, self._video_length
        return None, None, None
