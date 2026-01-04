import ctypes
from ctypes import wintypes
import win32con
import win32gui

from PyQt6.QtWidgets import QWidget, QApplication
from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import QCursor, QIcon

class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", wintypes.POINT),
        ("lPrivate", wintypes.DWORD), # Start of private data
    ]
    # Note: MSG structure might vary slightly but first 6 fields are standard.
    # Actually, on 64-bit, pointers are 64-bit. WPARAM/LPARAM are 64-bit.
    # The definition above uses wintypes which should be correct.

class OverlayWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        
        # Initialize geometry to full screen of primary screen
        self.update_geometry()

        self.setWindowTitle("Kazuha Overlay")
        self.setWindowIcon(QIcon("resources/icons/app_icon.ico"))

    def update_geometry(self):
        screen = QApplication.primaryScreen()
        if screen:
            self.setGeometry(screen.geometry())

    def nativeEvent(self, eventType, message):
        # 暂时屏蔽鼠标穿透逻辑，避免内存访问冲突导致静默崩溃
        # if eventType == "windows_generic_MSG":
        #     try:
        #         msg = MSG.from_address(int(message))
        #         if msg.message == win32con.WM_NCHITTEST:
        #             # Check if cursor is over a child widget
        #             pos = QCursor.pos()
        #             local_pos = self.mapFromGlobal(pos)
        #             
        #             child = self.childAt(local_pos)
        #             
        #             # If no child (or child is self), let events pass through
        #             if not child or child is self:
        #                 return True, win32con.HTTRANSPARENT
        #     except Exception:
        #         pass
                
        return super().nativeEvent(eventType, message)

    def paintEvent(self, event):
        # Do nothing to keep it transparent
        pass
