from PySide6.QtWidgets import QSystemTrayIcon, QMenu, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame
from PySide6.QtGui import QIcon, QAction, QPainter, QColor, QPixmap, QGuiApplication, QFont, QCursor, QPainterPath, QPen, QBrush
from PySide6.QtCore import Signal, QObject, Qt, QTimer, QPoint, QSize, QDateTime, QRectF
from PySide6.QtSvg import QSvgRenderer
import os
import time
from qfluentwidgets import themeColor

ICON_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "icons")

_ICON_CACHE = {}

def get_icon_pixmap(icon_path, size=18, color=None):
    cache_key = (icon_path, size, color.rgba() if color else None)
    if cache_key in _ICON_CACHE:
        return _ICON_CACHE[cache_key]
    
    if not os.path.exists(icon_path):
        return QPixmap()
        
    renderer = QSvgRenderer(icon_path)
    pixmap = QPixmap(size * 2, size * 2)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    renderer.render(painter)
    
    if color:
        painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
        painter.fillRect(pixmap.rect(), color)
        
    painter.end()
    
    _ICON_CACHE[cache_key] = pixmap
    return pixmap

class TrayMenuItem(QFrame):
    clicked = Signal()

    def __init__(self, icon_path, text, parent=None, is_exit=False, is_first=False, is_last=False):
        super().__init__(parent)
        self.is_exit = is_exit
        self.is_first = is_first
        self.is_last = is_last
        self.is_hovered = False
        self.text = text
        self.setFixedHeight(38)
        self.setCursor(Qt.PointingHandCursor)
        self.pixmap = get_icon_pixmap(icon_path, 18, QColor(255, 255, 255))
        
        if self.is_exit:
            self.hover_bg = QColor(255, 69, 58, 40)
        else:
            self.hover_bg = QColor(255, 255, 255, 15)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        rect = self.rect()
        radius = 12
        
        if self.is_hovered:
            painter.setBrush(self.hover_bg)
            painter.setPen(Qt.NoPen)
            if self.is_first and self.is_last:
                painter.drawRoundedRect(rect, radius, radius)
            elif self.is_first:
                path = QPainterPath()
                path.addRoundedRect(rect, radius, radius)
                path.addRect(0, rect.height() - radius, rect.width(), radius)
                painter.drawPath(path.simplified())
            elif self.is_last:
                path = QPainterPath()
                path.addRoundedRect(rect, radius, radius)
                path.addRect(0, 0, rect.width(), radius)
                painter.drawPath(path.simplified())
            else:
                painter.drawRect(rect)
        
        if not self.pixmap.isNull():
            target_rect = QRectF(12, (self.height() - 18) / 2, 18, 18)
            painter.drawPixmap(target_rect, self.pixmap, QRectF(self.pixmap.rect()))
            
        painter.setPen(QColor(255, 255, 255))
        font = QFont("MiSans", 10)
        font.setWeight(QFont.Medium)
        painter.setFont(font)
        text_rect = QRectF(40, 0, self.width() - 50, self.height())
        painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, self.text)
        
        if not self.is_last:
            painter.setPen(QColor(255, 255, 255, 15))
            painter.drawLine(12, self.height() - 1, self.width() - 12, self.height() - 1)

    def enterEvent(self, event):
        self.is_hovered = True
        self.update()

    def leaveEvent(self, event):
        self.is_hovered = False
        self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()

class TrayMenuGroup(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor(255, 255, 255, 10))
        painter.setPen(QPen(QColor(255, 255, 255, 20), 1))
        painter.drawRoundedRect(self.rect().adjusted(0, 0, -1, -1), 12, 12)

    def add_item(self, item):
        self.layout.addWidget(item)

class TrayMenu(QWidget):
    show_settings = Signal()
    show_timer = Signal()
    restart_app = Signal()
    exit_app = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(220, 260)
        self._last_hide_time = 0
        
        self.main_frame = QFrame(self)
        self.main_frame.setGeometry(10, 10, 200, 240)
        
        layout = QVBoxLayout(self.main_frame)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(6, 2, 6, 2)
        self.time_label = QLabel()
        self.time_label.setStyleSheet("color: rgba(255, 255, 255, 0.8); font-size: 15px; font-weight: 300; font-family: 'MiSans VF', 'MiSans', 'Segoe UI Variable Display', 'Segoe UI', sans-serif; background: transparent;")
        header_layout.addWidget(self.time_label)
        header_layout.addStretch()
        
        self.logo_label = QLabel()
        logo_path = os.path.join(ICON_DIR, "logo.svg")
        logo_pixmap = get_icon_pixmap(logo_path, 16)
        if not logo_pixmap.isNull():
            self.logo_label.setPixmap(logo_pixmap)
            self.logo_label.setFixedSize(16, 16)
            self.logo_label.setScaledContents(True)
        header_layout.addWidget(self.logo_label)
        layout.addLayout(header_layout)
        
        group1 = TrayMenuGroup(self.main_frame)
        self.item_settings = TrayMenuItem(os.path.join(ICON_DIR, "Settings.svg"), "设置", is_first=True)
        self.item_timer = TrayMenuItem(os.path.join(ICON_DIR, "Timer.svg"), "Timer 插件", is_last=True)
        
        self.item_settings.clicked.connect(self._on_settings)
        self.item_timer.clicked.connect(self._on_timer)
        
        group1.add_item(self.item_settings)
        group1.add_item(self.item_timer)
        layout.addWidget(group1)
        
        layout.addStretch()
        
        group2 = TrayMenuGroup(self.main_frame)
        self.item_restart = TrayMenuItem(os.path.join(ICON_DIR, "Next.svg"), "重新启动程序", is_first=True)
        self.item_exit = TrayMenuItem(os.path.join(ICON_DIR, "Clear.svg"), "退出程序", is_exit=True, is_last=True)
        
        self.item_restart.clicked.connect(self._on_restart)
        self.item_exit.clicked.connect(self._on_exit)
        
        group2.add_item(self.item_restart)
        group2.add_item(self.item_exit)
        layout.addWidget(group2)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Simple fast shadow
        shadow_rect = QRectF(self.main_frame.geometry()).adjusted(-2, -2, 2, 2)
        painter.setBrush(QColor(0, 0, 0, 40))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(shadow_rect, 22, 22)
        
        # Main background
        painter.setBrush(QColor(44, 44, 44))
        painter.setPen(QPen(QColor(255, 255, 255, 20), 1))
        painter.drawRoundedRect(self.main_frame.geometry(), 20, 20)

    def update_time(self):
        self.time_label.setText(QDateTime.currentDateTime().toString("H:mm"))

    def _on_settings(self):
        self.hide()
        self.show_settings.emit()

    def _on_timer(self):
        self.hide()
        self.show_timer.emit()

    def _on_restart(self):
        self.hide()
        self.restart_app.emit()

    def _on_exit(self):
        self.hide()
        self.exit_app.emit()

    def hideEvent(self, event):
        import time
        self._last_hide_time = time.time()
        super().hideEvent(event)

class SystemTray(QObject):
    show_settings = Signal()
    show_timer = Signal()
    restart_app = Signal()
    exit_app = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.tray_icon = QSystemTrayIcon(parent)
        self._update_icon()
        
        self.tray_icon.setToolTip("Kazuha 助手")
        
        self.menu = TrayMenu()
        self.menu.show_settings.connect(self.show_settings.emit)
        self.menu.show_timer.connect(self.show_timer.emit)
        self.menu.restart_app.connect(self.restart_app.emit)
        self.menu.exit_app.connect(self.exit_app.emit)
        
        self.tray_icon.activated.connect(self._on_activated)
        self.tray_icon.show()

    def _on_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.Context):
            import time
            now = time.time()
            if self.menu.isVisible() or (now - self.menu._last_hide_time < 0.2):
                self.menu.hide()
                return
                
            pos = QCursor.pos()
            screen = QGuiApplication.screenAt(pos)
            if screen:
                screen_geo = screen.geometry()
                x = pos.x() - self.menu.width() // 2
                y = pos.y() - self.menu.height()
                
                x = max(screen_geo.left(), min(x, screen_geo.right() - self.menu.width()))
                y = max(screen_geo.top(), min(y, screen_geo.bottom() - self.menu.height()))
                
                self.menu.update_time()
                self.menu.move(x, y)
                self.menu.show()
                self.menu.activateWindow()

    def _update_icon(self):
        logo_path = os.path.join(ICON_DIR, "logo.svg")
        if not os.path.exists(logo_path):
             logo_path = os.path.join(ICON_DIR, "Pen.svg")
        
        color = themeColor()
        
        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        renderer = QSvgRenderer(logo_path)
        renderer.render(painter)
        
        painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
        painter.fillRect(pixmap.rect(), color)
        painter.end()
        
        self.tray_icon.setIcon(QIcon(pixmap))
    
    def show_message(self, title, message):
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 2000)
