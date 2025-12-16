from PyQt6.QtWidgets import QWidget, QVBoxLayout, QApplication, QGraphicsDropShadowEffect
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from qfluentwidgets import isDarkTheme


class DetachedFlyoutWindow(QWidget):
    def __init__(self, content_widget, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

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
            mica_color = "#202020"
        else:
            mica_color = "#f3f3f3"
        self.background.setStyleSheet(f"QWidget#Background {{ background-color: {mica_color}; border-radius: 12px; }}")

        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(40)
        self.shadow.setColor(QColor(0, 0, 0, 80))
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
        self.show()

    def focusOutEvent(self, event):
        self.close()
        super().focusOutEvent(event)
