import win32com.client
import pythoncom
from PySide6.QtCore import QObject, Signal, QThread, QTimer, QPoint, QRect, Slot
from PySide6.QtGui import QGuiApplication
import time
from ppt_assistant.core.config import cfg

try:
    import win32gui
    import win32api
    import win32con
except ImportError:
    win32gui = None
    win32api = None
    win32con = None

class PPTWorker(QObject):
    """
    Worker thread for PPT COM operations to prevent blocking the main UI.
    """
    # Signals to Main Thread
    slideshow_started = Signal()
    slideshow_ended = Signal()
    slide_changed = Signal(int, int) # current, total
    window_geometry_changed = Signal(object, object) # QRect, QScreen
    video_state_changed = Signal(float, float, float) # ratio, pos, length
    thumbnail_generated = Signal(int, str) # index, path

    def __init__(self):
        super().__init__()
        self.ppt_app = None
        self.wps_app = None
        self._running = False
        self._current_slide = 0
        self._total_slides = 0
        self._last_win_rect = (0, 0, 0, 0)
        self._active_kind = None
        self._last_screen = None
        self._timer = None
        self._com_initialized = False

    @Slot()
    def start(self):
        if not self._com_initialized:
            pythoncom.CoInitialize()
            self._com_initialized = True
            
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._check_ppt_state)
        self._timer.start(200)

    @Slot()
    def stop(self):
        if self._timer:
            self._timer.stop()
        if self._com_initialized:
            pythoncom.CoUninitialize()
            self._com_initialized = False

    def _get_active_app(self):
        # Helper to get the currently tracked app
        if self._active_kind == "ppt" and self.ppt_app: return self.ppt_app
        if self._active_kind == "wps" and self.wps_app: return self.wps_app
        # Fallback
        if self.ppt_app: return self.ppt_app
        if self.wps_app: return self.wps_app
        return None

    def _check_ppt_state(self):
        try:
            # 1. Try PowerPoint
            try:
                self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                self.ppt_app = None
                self._check_wps_state()
                return

            if self.ppt_app.SlideShowWindows.Count > 0:
                try:
                    # Find the best slide show window (avoiding Presenter View if possible)
                    ss_win = None
                    count = self.ppt_app.SlideShowWindows.Count
                    
                    if count == 1:
                        ss_win = self.ppt_app.SlideShowWindows(1)
                    else:
                        # Try to find the one with class name "screenClass"
                        for i in range(1, count + 1):
                            try:
                                tmp_win = self.ppt_app.SlideShowWindows(i)
                                hwnd = getattr(tmp_win, "HWND", 0)
                                if hwnd:
                                    class_name = win32gui.GetClassName(int(hwnd))
                                    if class_name == "screenClass":
                                        ss_win = tmp_win
                                        break
                            except:
                                continue
                        
                        # Fallback to the first one if not found
                        if ss_win is None:
                            ss_win = self.ppt_app.SlideShowWindows(1)

                    view = ss_win.View
                    state = view.State
                    
                    if state in [1, 2]: # Running or Paused
                        if not self._running:
                            self._running = True
                            self._active_kind = "ppt"
                            self.slideshow_started.emit()
                        
                        try:
                            current = view.Slide.SlideIndex
                            presentation = ss_win.Presentation
                            total = presentation.Slides.Count
                            
                            if current != self._current_slide or total != self._total_slides:
                                self._current_slide = current
                                self._total_slides = total
                                self.slide_changed.emit(current, total)
                            
                            self._update_window_rect(ss_win)
                            self._update_video_state(ss_win)
                        except Exception:
                            pass
                    else:
                        self._handle_stop("ppt")
                except Exception:
                    pass
            else:
                self._handle_stop("ppt")

        except Exception:
            pass
            
        # If not running PPT, check WPS
        if not self._running:
            self._check_wps_state()

    def _check_wps_state(self):
        try:
            self.wps_app = win32com.client.GetActiveObject("KWPP.Application")
        except Exception:
            self.wps_app = None
            self._handle_stop("wps")
            return

        try:
            if self.wps_app.SlideShowWindows.Count > 0:
                ss_win = None
                count = self.wps_app.SlideShowWindows.Count
                
                if count == 1:
                    ss_win = self.wps_app.SlideShowWindows(1)
                else:
                    # Try to find the slideshow window (avoiding presenter view)
                    # WPS slideshow window class is usually "wppSlideShowWindowClass"
                    for i in range(1, count + 1):
                        try:
                            tmp_win = self.wps_app.SlideShowWindows(i)
                            hwnd = getattr(tmp_win, "HWND", 0)
                            if hwnd:
                                class_name = win32gui.GetClassName(int(hwnd))
                                if class_name == "wppSlideShowWindowClass":
                                    ss_win = tmp_win
                                    break
                        except:
                            continue
                    
                    if ss_win is None:
                        ss_win = self.wps_app.SlideShowWindows(1)

                view = ss_win.View
                # WPS State might differ, usually 1=Running
                state = getattr(view, "State", 1)
                
                if state in [1, 2]:
                    if not self._running:
                        self._running = True
                        self._active_kind = "wps"
                        self.slideshow_started.emit()

                    try:
                        current = view.Slide.SlideIndex
                        presentation = ss_win.Presentation
                        total = presentation.Slides.Count

                        if current != self._current_slide or total != self._total_slides:
                            self._current_slide = current
                            self._total_slides = total
                            self.slide_changed.emit(current, total)

                        self._update_window_rect(ss_win)
                        self._update_video_state(ss_win)
                    except Exception:
                        pass
                else:
                    self._handle_stop("wps")
            else:
                self._handle_stop("wps")
        except Exception:
            self._handle_stop("wps")

    def _handle_stop(self, kind):
        if self._running and (self._active_kind == kind or self._active_kind is None):
            self._running = False
            self._active_kind = None
            self.slideshow_ended.emit()

    def _update_window_rect(self, ss_win):
        try:
            l_left, l_top, l_width, l_height = 0, 0, 0, 0
            screen = None
            success = False
            
            # 1. Try Win32 API
            if win32gui:
                try:
                    hwnd = getattr(ss_win, "HWND", 0)
                    if hwnd:
                        hwnd = int(hwnd)
                        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                        w, h = right - left, bottom - top
                        cx, cy = left + w // 2, top + h // 2
                        
                        # Note: QGuiApplication calls are not thread-safe if they access GUI
                        # But screenAt/primaryScreen are generally okay.
                        # However, strictly we should calculate rect here and let Main Thread determine Screen.
                        # To be safe, we just emit the Rect and let Main Thread handle Screen resolution if possible.
                        # OR: We trust QGuiApplication read-only methods.
                        
                        # Optimization: Just send raw rect, let UI thread figure out DPI/Screen
                        # But existing logic does DPI scaling here. 
                        # We will assume DPI unawareness in worker and let Main Thread handle scaling if needed?
                        # Actually, raw pixels are better.
                        
                        # REVERTING to existing logic but being careful.
                        # Accessing QGuiApplication in thread is risky for some operations.
                        # Let's try to get screen in main thread.
                        # But wait, we need screen for DPI.
                        
                        # Let's emit raw global coords and let main thread map it.
                        final_rect = (left, top, w, h)
                        success = True
                except Exception:
                    pass

            # 2. Fallback to COM
            if not success:
                try:
                    l_left = int(getattr(ss_win, "Left", 0))
                    l_top = int(getattr(ss_win, "Top", 0))
                    l_width = int(getattr(ss_win, "Width", 0))
                    l_height = int(getattr(ss_win, "Height", 0))
                    final_rect = (l_left, l_top, l_width, l_height)
                    success = True
                except Exception:
                    pass

            if success:
                if final_rect != self._last_win_rect:
                    self._last_win_rect = final_rect
                    # We send RAW rect (x, y, w, h). Main thread converts to QRect and finds Screen.
                    self.window_geometry_changed.emit(QRect(*final_rect), None)
                    
        except Exception:
            pass

    def _update_video_state(self, ss_win):
        try:
            view = ss_win.View
            try:
                slide = view.Slide
                shapes = slide.Shapes
                count = shapes.Count
                found_video = False
                for i in range(1, count + 1):
                    shape = shapes.Item(i)
                    media = getattr(shape, "MediaFormat", None)
                    if media is None: continue
                    
                    length = getattr(media, "Length", 0)
                    position = getattr(media, "Position", 0)
                    
                    if length and float(length) > 0:
                        l = float(length)
                        p = float(position or 0.0)
                        ratio = p / l if l > 0 else 0
                        self.video_state_changed.emit(ratio, p, l)
                        found_video = True
                        break
                
                if not found_video:
                    self.video_state_changed.emit(0.0, 0.0, 0.0)
            except Exception:
                pass
        except Exception:
            pass

    # --- Control Slots ---
    @Slot()
    def go_next(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.Next()
        except Exception:
            pass

    @Slot()
    def go_previous(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.Previous()
        except Exception:
            pass

    @Slot()
    def clear_screen(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                ss_win = app.SlideShowWindows(1)
                hwnd = int(getattr(ss_win, "HWND", 0) or 0)
                if hwnd and win32gui:
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                if win32api and win32con:
                    try:
                        vk = ord("E")
                        win32api.keybd_event(vk, 0, 0, 0)
                        win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
                    except Exception:
                        pass
        except Exception:
            pass

    @Slot()
    def end_show(self):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                if cfg.autoHandleInk.value and self._active_kind == "ppt":
                    try: app.DisplayAlerts = 1 
                    except: pass
                
                app.SlideShowWindows(1).View.Exit()

                if cfg.autoHandleInk.value and self._active_kind == "ppt":
                    try: app.DisplayAlerts = -1
                    except: pass
        except Exception:
            pass

    @Slot(int)
    def set_pointer_type(self, pointer_type):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.PointerType = pointer_type
        except Exception:
            pass

    @Slot(int, int, int)
    def set_pen_color(self, r, g, b):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                rgb = r + (g << 8) + (b << 16)
                app.SlideShowWindows(1).View.PointerColor.RGB = rgb
        except Exception:
            pass

    @Slot(int)
    def go_to_slide(self, index):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.GotoSlide(index)
        except Exception:
            pass
            
    @Slot(int, str)
    def export_slide_thumbnail(self, index, path):
        try:
            app = self._get_active_app()
            if app and app.SlideShowWindows.Count > 0:
                pres = app.SlideShowWindows(1).Presentation
                if 1 <= index <= pres.Slides.Count:
                    pres.Slides(index).Export(path, "PNG", 320, 180)
                    self.thumbnail_generated.emit(index, path)
        except Exception:
            pass


class PPTMonitor(QObject):
    """
    Facade for PPTWorker. Runs worker in a separate thread.
    """
    slideshow_started = Signal()
    slideshow_ended = Signal()
    slide_changed = Signal(int, int)
    window_geometry_changed = Signal(object, object)
    video_state_changed = Signal(float, float, float)
    thumbnail_generated = Signal(int, str)
    
    # Internal signals to worker
    _req_start = Signal()
    _req_stop = Signal()
    _req_next = Signal()
    _req_prev = Signal()
    _req_clear = Signal()
    _req_end = Signal()
    _req_ptr_type = Signal(int)
    _req_pen_color = Signal(int, int, int)
    _req_goto = Signal(int)
    _req_export = Signal(int, str)

    def __init__(self):
        super().__init__()
        self._thread = QThread()
        self._worker = PPTWorker()
        self._worker.moveToThread(self._thread)

        # Wire up signals (Worker -> Self)
        self._worker.slideshow_started.connect(self.slideshow_started)
        self._worker.slideshow_ended.connect(self.slideshow_ended)
        self._worker.slide_changed.connect(self._on_slide_changed)
        self._worker.window_geometry_changed.connect(self._on_geometry_changed)
        self._worker.video_state_changed.connect(self.video_state_changed)
        self._worker.video_state_changed.connect(self._update_local_video_state)
        self._worker.thumbnail_generated.connect(self.thumbnail_generated)

        # Wire up requests (Self -> Worker)
        self._req_start.connect(self._worker.start)
        self._req_stop.connect(self._worker.stop)
        self._req_next.connect(self._worker.go_next)
        self._req_prev.connect(self._worker.go_previous)
        self._req_clear.connect(self._worker.clear_screen)
        self._req_end.connect(self._worker.end_show)
        self._req_ptr_type.connect(self._worker.set_pointer_type)
        self._req_pen_color.connect(self._worker.set_pen_color)
        self._req_goto.connect(self._worker.go_to_slide)
        self._req_export.connect(self._worker.export_slide_thumbnail)
        
        # Local state cache (for synchronous getters if needed)
        self._current = 0
        self._total = 0
        self._video_ratio = 0.0
        self._video_pos = 0.0
        self._video_len = 0.0
        
        self._thread.start()

    def start_monitoring(self):
        self._req_start.emit()

    def stop_monitoring(self):
        self._req_stop.emit()
        self._thread.quit()
        self._thread.wait()

    # --- Public API (Async) ---
    def go_next(self):
        self._req_next.emit()

    def go_previous(self):
        self._req_prev.emit()

    def clear_screen(self):
        self._req_clear.emit()

    def end_show(self):
        self._req_end.emit()

    def set_pointer_type(self, ptr_type):
        self._req_ptr_type.emit(ptr_type)

    def set_pen_color(self, r, g, b):
        self._req_pen_color.emit(r, g, b)

    def go_to_slide(self, index):
        self._req_goto.emit(index)

    def export_slide_thumbnail(self, index, path):
        self._req_export.emit(index, path)
        
    def force_update_geometry(self):
        # We can't force update easily from main thread without roundtrip
        # But we can re-emit last known if we cached it.
        # For now, ignore or implement caching if critical.
        pass

    # --- State Handling ---
    def _on_slide_changed(self, current, total):
        self._current = current
        self._total = total
        self.slide_changed.emit(current, total)

    def _on_geometry_changed(self, rect_raw, _):
        x, y, w, h = rect_raw.x(), rect_raw.y(), rect_raw.width(), rect_raw.height()
        cx, cy = x + w // 2, y + h // 2
        
        target_mode = cfg.overlayScreen.value
        screens = QGuiApplication.screens()
        
        ppt_screen = None
        ppt_p_origin = (0, 0)
        
        try:
            hmonitor = win32api.MonitorFromPoint((cx, cy), win32con.MONITOR_DEFAULTTONEAREST)
            m_info = win32api.GetMonitorInfo(hmonitor)
            m_name = m_info['Device']
            ppt_p_origin = (m_info['Monitor'][0], m_info['Monitor'][1])
            
            for s in screens:
                if s.name() == m_name:
                    ppt_screen = s
                    break
        except Exception:
            pass

        if not ppt_screen:
            ppt_screen = QGuiApplication.primaryScreen()

        display_screen = None
        if target_mode == "Primary":
            display_screen = QGuiApplication.primaryScreen()
        elif target_mode.startswith("Screen "):
            try:
                idx = int(target_mode.split(" ")[1]) - 1
                if 0 <= idx < len(screens):
                    display_screen = screens[idx]
            except:
                pass
        
        if not display_screen or target_mode == "Auto":
            display_screen = ppt_screen

        if display_screen and ppt_screen:
            if display_screen == ppt_screen:
                dpr = ppt_screen.devicePixelRatio()
                geo = ppt_screen.geometry()
                lx = geo.x() + (x - ppt_p_origin[0]) / dpr
                ly = geo.y() + (y - ppt_p_origin[1]) / dpr
                lw = w / dpr
                lh = h / dpr
                rect_logical = QRect(int(lx), int(ly), int(lw), int(lh))
            else:
                rect_logical = display_screen.geometry()
            
            self.window_geometry_changed.emit(rect_logical, display_screen)
        else:
            self.window_geometry_changed.emit(rect_raw, None)

    def _update_local_video_state(self, ratio, pos, length):
        self._video_ratio = ratio
        self._video_pos = pos
        self._video_len = length

    # --- Getters (Cached) ---
    def get_page_info(self):
        return self._current, self._total

    def get_total_slides(self):
        return self._total
        
    def get_video_progress(self):
        return self._video_ratio, self._video_pos, self._video_len

