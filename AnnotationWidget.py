from PyQt6.QtWidgets import QWidget, QApplication
from PyQt6.QtGui import QPainter, QPen, QImage, QCursor
from PyQt6.QtCore import Qt, QPoint, QTimer
import sys
import win32api
import win32con
import win32gui
import win32process
import win32event
from ctypes import windll, c_int, c_long, byref


class AnnotationWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # ========== 1. Windows底层设置（核心：拦截所有输入） ==========
        # 禁用Windows的底层输入传递
        self.user32 = windll.user32
        self.user32.BlockInput(True)  # 临时屏蔽系统输入（仅鼠标+键盘）
        
        # ========== 2. Qt窗口标志（极简且稳定） ==========
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        
        # ========== 3. 关键属性 ==========
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_MouseTracking, True)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        
        # ========== 4. 全屏覆盖 ==========
        screen = QApplication.primaryScreen()
        self.screen_geo = screen.geometry()
        self.setGeometry(self.screen_geo)
        print(f"AnnotationWidget geometry: {self.geometry()}")
        
        # ========== 5. 绘图初始化 ==========
        self.drawing = False
        self.pen_size = 5
        self.pen_color = Qt.GlobalColor.red
        self.last_point = QPoint()
        
        self.image = QImage(self.screen_geo.size(), QImage.Format.Format_ARGB32)
        self.image.fill(Qt.GlobalColor.transparent)
        
        # 撤销栈
        self.image_stack = [self.image.copy()]
        self.current_stack_pos = 0
        self.stack_limit = 50
        
        # ========== 6. 强制激活+捕获 ==========
        self.show()
        self.raise_()
        self.grabMouse(QCursor(Qt.CursorShape.CrossCursor))
        
        # 延时验证捕获状态
        QTimer.singleShot(200, self._check_mouse_grab)

    def _check_mouse_grab(self):
        """验证鼠标捕获状态"""
        if self.mouseGrabber() == self:
            print("✅ 鼠标捕获成功，无法点击其他窗口")
        else:
            print("⚠️ 鼠标捕获失败，重试...")
            self.grabMouse()
            self.user32.BlockInput(True)  # 再次屏蔽系统输入

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drawing = True
            self.last_point = event.pos()
            with QPainter(self.image) as painter:
                pen = QPen(self.pen_color, self.pen_size, Qt.PenStyle.SolidLine,
                          Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
                painter.setPen(pen)
                painter.drawPoint(self.last_point)
            self.update()
        event.accept()

    def mouseMoveEvent(self, event):
        if self.drawing and event.buttons() & Qt.MouseButton.LeftButton:
            with QPainter(self.image) as painter:
                pen = QPen(self.pen_color, self.pen_size, Qt.PenStyle.SolidLine,
                          Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
                painter.setPen(pen)
                painter.drawLine(self.last_point, event.pos())
            self.last_point = event.pos()
            self.update()
        event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.drawing:
            self.drawing = False
            self._push_stack()
        event.accept()

    def keyPressEvent(self, event):
        """ESC退出（恢复系统输入），Ctrl+Z撤销，Ctrl+Y重做"""
        if event.key() == Qt.Key.Key_Escape:
            # 恢复系统输入+释放鼠标
            self.user32.BlockInput(False)
            if self.mouseGrabber() == self:
                self.releaseMouse()
            self.close()
        elif event.key() == Qt.Key.Key_Z and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.undo()
        elif event.key() == Qt.Key.Key_Y and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.redo()
        event.accept()

    def closeEvent(self, event):
        """关闭时必须恢复系统输入！"""
        self.user32.BlockInput(False)  # 关键：解除系统输入屏蔽
        if self.mouseGrabber() == self:
            self.releaseMouse()
        event.accept()

    def _push_stack(self):
        if self.current_stack_pos < len(self.image_stack) - 1:
            self.image_stack = self.image_stack[:self.current_stack_pos + 1]
        self.image_stack.append(self.image.copy())
        if len(self.image_stack) > self.stack_limit:
            self.image_stack.pop(0)
        self.current_stack_pos = len(self.image_stack) - 1

    def undo(self):
        if self.current_stack_pos > 0:
            self.current_stack_pos -= 1
            self.image = self.image_stack[self.current_stack_pos].copy()
            self.update()

    def redo(self):
        if self.current_stack_pos < len(self.image_stack) - 1:
            self.current_stack_pos += 1
            self.image = self.image_stack[self.current_stack_pos].copy()
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.drawImage(self.rect(), self.image, self.image.rect())
        painter.end()


if __name__ == "__main__":
    # 高DPI适配（Windows）
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except:
        pass
    
    app = QApplication(sys.argv)
    # 强制设置应用为前台
    hwnd = win32gui.GetForegroundWindow()
    win32process.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    
    widget = AnnotationWidget()
    sys.exit(app.exec())