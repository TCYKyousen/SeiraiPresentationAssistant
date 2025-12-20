from PyQt6.QtWidgets import QWidget, QVBoxLayout, QApplication, QGraphicsDropShadowEffect
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QColor
from qfluentwidgets import isDarkTheme


class DetachedFlyoutWindow(QWidget):
    def __init__(self, content_widget, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.show_anim = None
        self.hide_anim = None
        self._closing = False

        self.setWindowFlags(Qt.WindowType.Popup)

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)

        self.background = QWidget(self)
        self.background.setObjectName("Background")
        inner_layout = QVBoxLayout(self.background)
        inner_layout.setContentsMargins(10, 10, 10, 10)

        self.content = content_widget
        inner_layout.addWidget(self.content)
        root_layout.addWidget(self.background)

        if isDarkTheme():
            mica_color = "rgba(32, 32, 32, 0.95)"
            border = "1px solid rgba(255, 255, 255, 0.08)"
        else:
            mica_color = "rgba(243, 243, 243, 0.95)"
            border = "1px solid rgba(0, 0, 0, 0.05)"
            
        self.background.setStyleSheet(f"""
            QWidget#Background {{ 
                background-color: {mica_color}; 
                border-radius: 8px; 
                border: {border};
            }}
        """)

        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(32)
        self.shadow.setColor(QColor(0, 0, 0, 60))
        self.shadow.setOffset(0, 8)
        self.background.setGraphicsEffect(self.shadow)

    def show_at(self, target_widget):
        self.adjustSize()
        w = self.width()
        h = self.height()

        rect = target_widget.rect()
        global_pos = target_widget.mapToGlobal(rect.topLeft())

        x = global_pos.x() + rect.width() // 2 - w // 2
        y = global_pos.y() + rect.height() + 5

        screen = QApplication.primaryScreen().geometry()

        if x < screen.left():
            x = screen.left() + 5
        if x + w > screen.right():
            x = screen.right() - w - 5

        if y + h > screen.bottom():
            y = global_pos.y() - h - 5
        
        self.move(x, y)
        self.setWindowOpacity(0.0)
        self.show()
        self.show_anim = QPropertyAnimation(self, b"windowOpacity", self)
        self.show_anim.setDuration(160)
        self.show_anim.setStartValue(0.0)
        self.show_anim.setEndValue(1.0)
        self.show_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.show_anim.start()

    def focusOutEvent(self, event):
        self.close()
        super().focusOutEvent(event)

    def closeEvent(self, event):
        if self._closing:
            super().closeEvent(event)
            return
        event.ignore()
        self.hide_anim = QPropertyAnimation(self, b"windowOpacity", self)
        self.hide_anim.setDuration(160)
        self.hide_anim.setStartValue(self.windowOpacity())
        self.hide_anim.setEndValue(0.0)
        self.hide_anim.setEasingCurve(QEasingCurve.Type.InCubic)
        self.hide_anim.finished.connect(self._finish_close)
        self.hide_anim.start()

    def _finish_close(self):
        self._closing = True
        self.close()
        self._closing = False
