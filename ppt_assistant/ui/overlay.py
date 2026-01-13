from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QFrame, QApplication, QLabel, QPushButton, QSwipeGesture, QGestureEvent, QGridLayout, QStyleOption, QStyle, QGraphicsDropShadowEffect, QMenu
from PySide6.QtCore import Qt, Signal, QSize, QPoint, QEvent, QTimer, QTime
from PySide6.QtGui import QColor, QIcon, QPainter, QBrush, QPen, QPixmap, QGuiApplication, QFont, QPalette, QLinearGradient, QAction
from PySide6.QtSvg import QSvgRenderer
import os
import importlib.util
import sys
import tempfile
import json
import subprocess
from ppt_assistant.core.config import cfg, SETTINGS_PATH
from qfluentwidgets import FluentWidget, FluentIcon as FIF, BodyLabel, IconWidget, themeColor

try:
    import psutil
except ImportError:
    psutil = None

ICON_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "icons")
PLUGIN_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "plugins", "builtins")


def _load_language():
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("General", {}).get("Language", "zh-CN")
    except Exception:
        return "zh-CN"
    return "zh-CN"


LANGUAGE = _load_language()

_TRANSLATIONS = {
    "zh-CN": {
        "status.media_length": "媒体时长",
        "overlay.title": "顶层效果窗口",
        "toolbar.select": "选择",
        "toolbar.pen": "画笔",
        "toolbar.eraser": "橡皮",
        "toolbar.undo": "上一步",
        "toolbar.redo": "下一步",
        "toolbar.end_show": "结束放映",
        "toolbar.page": "页码",
        "toolbar.theme_colors": "主题颜色",
        "toolbar.standard_colors": "标准颜色",
    },
    "zh-TW": {
        "status.media_length": "媒體時長",
        "overlay.title": "頂層效果視窗",
        "toolbar.select": "選取",
        "toolbar.pen": "畫筆",
        "toolbar.eraser": "橡皮擦",
        "toolbar.undo": "上一步",
        "toolbar.redo": "下一步",
        "toolbar.end_show": "結束播放",
        "toolbar.page": "頁碼",
        "toolbar.theme_colors": "主題顏色",
        "toolbar.standard_colors": "标准颜色",
    },
    "ja-JP": {
        "status.media_length": "メディア長さ",
        "overlay.title": "オーバーレイウィンドウ",
        "toolbar.select": "選択",
        "toolbar.pen": "ペン",
        "toolbar.eraser": "消しゴム",
        "toolbar.undo": "戻る",
        "toolbar.redo": "進む",
        "toolbar.end_show": "スライド終了",
        "toolbar.page": "ページ番号",
        "toolbar.theme_colors": "テーマの色",
        "toolbar.standard_colors": "標準の色",
    },
    "en-US": {
        "status.media_length": "Media duration",
        "overlay.title": "Overlay window",
        "toolbar.select": "Select",
        "toolbar.pen": "Pen",
        "toolbar.eraser": "Eraser",
        "toolbar.undo": "Undo",
        "toolbar.redo": "Redo",
        "toolbar.end_show": "End Show",
        "toolbar.page": "Page number",
        "toolbar.theme_colors": "Theme Colors",
        "toolbar.standard_colors": "Standard Colors",
    },
}


def _get_app_version():
    try:
        root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        version_path = os.path.join(root_dir, "version.json")
        if not os.path.exists(version_path):
            return ""
        with open(version_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return str(data.get("version", "")).strip()
    except Exception:
        return ""


def _is_dev_preview_version(version: str) -> bool:
    if not version:
        return False
    parts = str(version).strip().split(".")
    if len(parts) < 2:
        return False
    return parts[-1] == "1"


def _t(key: str) -> str:
    lang = LANGUAGE
    table = _TRANSLATIONS.get(lang) or _TRANSLATIONS["zh-CN"]
    if key in table:
        return table[key]
    default = _TRANSLATIONS["zh-CN"]
    return default.get(key, key)


class StatusBarWidget(QFrame):
    is_light_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(26)
        self._is_light = False
        self._monitor = None
        self._network_kind = "offline"
        self._volume_supported = False
        self._build_ui()
        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._update_time)
        self._clock_timer.start(1000)
        self._color_timer = QTimer(self)
        self._color_timer.timeout.connect(self._update_color_from_screen)
        self._color_timer.start(2000)
        self._video_timer = QTimer(self)
        self._video_timer.timeout.connect(self._update_video)
        self._video_timer.start(500)
        self._network_timer = QTimer(self)
        self._network_timer.timeout.connect(self._update_network)
        self._network_timer.start(5000)
        self._volume_timer = QTimer(self)
        self._volume_timer.timeout.connect(self._update_volume)
        self._volume_timer.start(1000)
        self._update_time()
        self._update_palette()
        self._update_volume()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(8)

        self.time_label = QLabel("", self)
        self.time_label.setObjectName("TimeLabel")
        layout.addWidget(self.time_label)

        self.center_widget = QWidget(self)
        center_layout = QHBoxLayout(self.center_widget)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(6)
        self.progress_value = QLabel("", self.center_widget)
        self.progress_value.setObjectName("ProgressValue")
        self.progress_caption = QLabel(_t("status.media_length"), self.center_widget)
        self.progress_caption.setObjectName("ProgressCaption")
        self.progress_caption.hide()
        self.progress_value.hide()
        center_layout.addWidget(self.progress_value)
        center_layout.addWidget(self.progress_caption, 0, Qt.AlignBottom)

        layout.addStretch(1)
        layout.addWidget(self.center_widget)
        layout.addStretch(1)

        self.net_icon = IconWidget(FIF.WIFI, self)
        self.net_icon.setFixedSize(18, 18)
        layout.addWidget(self.net_icon)

        self.volume_icon = IconWidget(FIF.VOLUME, self)
        self.volume_icon.setFixedSize(18, 18)
        layout.addWidget(self.volume_icon)

    def _update_time(self):
        self.time_label.setText(QTime.currentTime().toString("HH:mm"))

    def set_monitor(self, monitor):
        self._monitor = monitor

    def _update_video(self):
        if not self._monitor:
            self.progress_caption.hide()
            self.progress_value.hide()
            return
        try:
            ratio, pos, length = self._monitor.get_video_progress()
        except Exception:
            self.progress_caption.hide()
            self.progress_value.hide()
            return
        if length is None or length <= 0:
            self.progress_caption.hide()
            self.progress_value.hide()
            return
        length_sec = length or 0.0
        if length_sec > 36000:
            length_sec = length_sec / 1000.0
        total_text = self._format_seconds(length_sec)
        self.progress_value.setText(total_text)
        self.progress_caption.show()
        self.progress_value.show()

    def _format_seconds(self, value):
        secs = int(float(value))
        if secs < 0:
            secs = 0
        m = secs // 60
        s = secs % 60
        return f"{m:02}:{s:02}"

    def _update_color_from_screen(self):
        # Disabled adaptive logic as per user request for forced light theme
        pass

    def _update_network(self):
        kind = "offline"
        if psutil is not None:
            try:
                stats = psutil.net_if_stats()
                for name, st in stats.items():
                    if not st.isup:
                        continue
                    lname = name.lower()
                    if "wi-fi" in lname or "wifi" in lname or "wlan" in lname:
                        kind = "wifi"
                        break
                    if "ethernet" in lname or "eth" in lname or "lan" in lname:
                        kind = "wired"
                if kind == "offline" and any(st.isup for st in stats.values()):
                    kind = "wired"
            except Exception:
                pass
        if kind == "offline":
            try:
                out = subprocess.check_output(
                    ["netsh", "wlan", "show", "interfaces"],
                    encoding="utf-8",
                    errors="ignore",
                )
                if "state" in out.lower() and "connected" in out.lower():
                    kind = "wifi"
            except Exception:
                pass
        self._network_kind = kind
        if kind == "wifi":
            self.net_icon.setIcon(FIF.WIFI)
        elif kind == "wired":
            icon = getattr(FIF, "ETHERNET", FIF.WIFI)
            self.net_icon.setIcon(icon)
        else:
            self.net_icon.setIcon(FIF.WIFI)

    def _update_volume(self):
        try:
            self.volume_icon.setIcon(FIF.VOLUME)
        except Exception:
            pass

    def _update_palette(self, is_light=False):
        self._is_light = is_light
        if is_light:
            fg = "#191919"
            bg = "rgba(255, 255, 255, 0.8)"
            border = "rgba(0, 0, 0, 0.1)"
        else:
            fg = "#FFFFFF"
            bg = "rgba(28, 28, 30, 0.82)"
            border = "rgba(255, 255, 255, 0.15)"
            
        self.setStyleSheet(
            f"""
            StatusBarWidget {{
                background-color: {bg};
                border-bottom: 0.5px solid {border};
            }}
            QLabel {{
                color: {fg};
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
                font-size: 12px;
                background: transparent;
            }}
            QLabel#TimeLabel, QLabel#ProgressValue {{
                font-weight: 900;
            }}
            QLabel#ProgressCaption {{
                font-size: 9px;
            }}
            IconWidget {{
                color: {fg};
            }}
        """
        )

class ClickableLabel(QLabel):
    clicked = Signal(QPoint)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if hasattr(event, "globalPosition"):
                pos = event.globalPosition().toPoint()
            else:
                pos = event.globalPos()
            self.clicked.emit(pos)
        super().mousePressEvent(event)

class MarqueeLabel(QLabel):
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self._offset = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._update_offset)
        self._timer.setInterval(30)
        self._spacing = 40
        self._should_scroll = False
        self._text_width = 0
        self._is_paused = False
        self._pause_timer = QTimer(self)
        self._pause_timer.setSingleShot(True)
        self._pause_timer.timeout.connect(self._resume_scroll)

    def setText(self, text):
        super().setText(text)
        self._update_scroll_state()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_scroll_state()

    def _update_scroll_state(self):
        font_metrics = self.fontMetrics()
        self._text_width = font_metrics.horizontalAdvance(self.text())
        if self._text_width > self.width() and self.width() > 0:
            self._should_scroll = True
            self._offset = 0
            self._is_paused = False
            if not self._timer.isActive():
                self._timer.start()
        else:
            self._should_scroll = False
            self._timer.stop()
            self._offset = 0
            self.update()

    def _update_offset(self):
        if not self._should_scroll or self._is_paused:
            return
        
        self._offset += 1
        if self._offset >= self._text_width + self._spacing:
            self._offset = 0
            self._is_paused = True
            self._pause_timer.start(2000) # Pause for 2 seconds
        
        self.update()

    def _resume_scroll(self):
        self._is_paused = False

    def paintEvent(self, event):
        if not self._should_scroll:
            super().paintEvent(event)
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        
        rect = self.rect()
        dpr = self.devicePixelRatioF()
        
        # Create a linear gradient mask for the fade effect
        gradient = QLinearGradient(0, 0, rect.width(), 0)
        gradient.setColorAt(0, Qt.transparent)
        gradient.setColorAt(0.15, Qt.black)
        gradient.setColorAt(0.85, Qt.black)
        gradient.setColorAt(1, Qt.transparent)
        
        # Draw everything into a temporary pixmap with High DPI support
        pixmap = QPixmap(rect.size() * dpr)
        pixmap.setDevicePixelRatio(dpr)
        pixmap.fill(Qt.transparent)
        
        p = QPainter(pixmap)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        
        # Get color from palette
        color = self.palette().color(QPalette.WindowText)
        p.setPen(color)
        p.setFont(self.font())
        
        y = (rect.height() + self.fontMetrics().ascent() - self.fontMetrics().descent()) // 2
        
        x = -self._offset
        p.drawText(x, y, self.text())
        if self._text_width > 0:
            p.drawText(x + self._text_width + self._spacing, y, self.text())
        p.end()
        
        # Apply the gradient mask
        mask_painter = QPainter(pixmap)
        mask_painter.setCompositionMode(QPainter.CompositionMode_DestinationIn)
        mask_painter.fillRect(pixmap.rect(), gradient)
        mask_painter.end()
        
        # Draw the final result
        painter.drawPixmap(0, 0, pixmap)
        painter.end()

class PenColorPopup(QFrame):
    color_selected = Signal(int, int, int)

    def __init__(self, parent=None, colors=None, is_light=False):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_StyledBackground, True)
        
        # Office Standard Colors + Theme Colors (approximate)
        # Standard Colors
        self.standard_colors = [
            (192, 0, 0), (255, 0, 0), (255, 192, 0), (255, 255, 0), (146, 208, 80),
            (0, 176, 80), (0, 176, 240), (0, 112, 192), (0, 32, 96), (112, 48, 160)
        ]
        
        # Theme Base Colors (Office default theme)
        self.theme_bases = [
            (255, 255, 255), (0, 0, 0), (231, 230, 230), (68, 84, 106), (68, 114, 196),
            (237, 125, 49), (165, 165, 165), (255, 192, 0), (91, 155, 213), (112, 173, 71)
        ]
        
        bg = "white" if is_light else "rgb(32, 32, 32)"
        border = "rgba(0, 0, 0, 0.08)" if is_light else "rgba(255, 255, 255, 0.1)"
        fg = "black" if is_light else "white"
        
        # Main layout for the popup window
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Container for content and background
        self.container = QFrame(self)
        self.container.setObjectName("PenColorContainer")
        self.container.setStyleSheet(f"""
            #PenColorContainer {{
                background-color: {bg};
                border-radius: 12px;
                border: 1px solid {border};
            }}
            QLabel {{
                color: {fg};
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
                font-size: 11px;
                font-weight: 500;
                background: transparent;
            }}
        """)
        self.layout.addWidget(self.container)

        main_layout = QVBoxLayout(self.container)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # Theme Colors
        theme_label = QLabel(_t("toolbar.theme_colors"), self.container)
        main_layout.addWidget(theme_label)
        
        theme_grid = QGridLayout()
        theme_grid.setSpacing(4)
        for i, (r, g, b) in enumerate(self.theme_bases):
            btn = self._create_color_btn(r, g, b)
            theme_grid.addWidget(btn, 0, i)
        main_layout.addLayout(theme_grid)

        # Standard Colors
        std_label = QLabel(_t("toolbar.standard_colors"), self.container)
        main_layout.addWidget(std_label)
        
        std_grid = QGridLayout()
        std_grid.setSpacing(4)
        for i, (r, g, b) in enumerate(self.standard_colors):
            btn = self._create_color_btn(r, g, b)
            std_grid.addWidget(btn, 0, i)
        main_layout.addLayout(std_grid)

    def _create_color_btn(self, r, g, b):
        btn = QPushButton(self.container)
        btn.setFixedSize(20, 20)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: rgb({r}, {g}, {b});
                border-radius: 4px;
                border: 1px solid rgba(0, 0, 0, 0.1);
            }}
            QPushButton:hover {{
                border: 2px solid white;
            }}
        """)
        btn.clicked.connect(lambda: self._select_color(r, g, b))
        return btn

    def _select_color(self, r, g, b):
        self.color_selected.emit(r, g, b)
        self.close()

    def paintEvent(self, event):
        # Ensure QSS is applied even for custom widgets
        opt = QStyleOption()
        opt.initFrom(self)
        painter = QPainter(self)
        self.style().drawPrimitive(QStyle.PE_Widget, opt, painter, self)
        super().paintEvent(event)

class CustomToolButton(QFrame):
    clicked = Signal()

    def __init__(self, icon_name, tooltip, parent=None, is_exit=False, tool_name=None, text="", pixmap=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.is_exit = is_exit
        self.tool_name = tool_name
        self.icon_name = icon_name
        self.text = text
        self.pixmap = pixmap
        self.is_active = False
        
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip(tooltip)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(4, 4, 4, 4)
        self.layout.setSpacing(6)
        self.layout.setAlignment(Qt.AlignCenter)
        
        self.icon_label = QLabel(self)
        self.icon_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.icon_label)
        
        self.text_label = MarqueeLabel(text, self)
        self.text_label.setAlignment(Qt.AlignCenter)
        self.text_label.setStyleSheet("""
            QLabel {
                font-size: 8px;
                font-weight: 300;
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
                color: rgba(255, 255, 255, 0.6);
                background: transparent;
            }
        """)
        self.text_label.setVisible(bool(text) and cfg.showToolbarText.value)
        self.layout.addWidget(self.text_label)
        
        # Connect context menu (placeholder for future use)
        # self.customContextMenuRequested.connect(self._show_context_menu)
        
        self.update_size()
        self.update_style(False)
        self.set_icon_color(False)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)

    def update_size(self):
        show_text = cfg.showToolbarText.value and bool(self.text)
        if show_text:
            # Match HTML: .toolbar-preview-btn-capsule { min-width: 68px; height: 40px; }
            # Increased width to 92 for horizontal layout and marquee room
            self.setFixedSize(92, 38)
            self.icon_size = 18
            self.layout.setContentsMargins(12, 4, 12, 4)
        else:
            # Match HTML: .toolbar-preview-btn { min-width: 38px; height: 38px; }
            self.setFixedSize(38, 38)
            self.icon_size = 20
            self.layout.setContentsMargins(4, 4, 4, 4)
        
        self.set_icon_color(False)

    def set_icon_color(self, is_light):
        s = self.icon_size
        if self.pixmap:
            scaled_pixmap = self.pixmap.scaled(s, s, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.icon_label.setPixmap(scaled_pixmap)
            self.icon_label.setFixedSize(s, s)
            return

        icon_path = os.path.join(ICON_DIR, self.icon_name)
        if not os.path.exists(icon_path):
            return
            
        color = QColor(255, 255, 255)
        renderer = QSvgRenderer(icon_path)
        if not renderer.isValid():
            return
            
        device_size = s * 2
        pixmap = QPixmap(device_size, device_size)
        pixmap.fill(Qt.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        
        if self.is_exit:
            # Render icon with red color directly, no background circle
            renderer.render(painter)
            painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
            painter.fillRect(pixmap.rect(), QColor("#FF453A"))
        else:
            # Normal icon colorization
            renderer.render(painter)
            painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
            painter.fillRect(pixmap.rect(), color)
            
        painter.end()
        
        scaled_pixmap = pixmap.scaled(s, s, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.icon_label.setPixmap(scaled_pixmap)
        self.icon_label.setFixedSize(s, s)

    def update_style(self, is_active, is_light=False, use_indicator=False):
        self.is_active = is_active
        show_text = cfg.showToolbarText.value and bool(self.text)
        
        # Hover colors based on theme
        if is_light:
            hover_bg = "rgba(0, 0, 0, 0.06)"
            active_bg = "rgba(0, 0, 0, 0.12)"
        else:
            hover_bg = "rgba(255, 255, 255, 0.08)"
            active_bg = "rgba(255, 255, 255, 0.15)"

        if is_active:
            bg = active_bg
            fg_alpha = 1.0
        else:
            bg = "transparent"
            fg_alpha = 0.9
        
        radius = self.height() // 2
            
        self.setStyleSheet(f"""
            CustomToolButton {{
                background-color: {bg};
                border-radius: {radius}px;
                border: none;
            }}
            CustomToolButton:hover {{
                background-color: {hover_bg};
            }}
        """)
        
        # Match HTML: .toolbar-preview-btn-capsule .toolbar-preview-btn-icon { background: rgba(255, 255, 255, 0.1); }
        if show_text:
            self.icon_label.setStyleSheet("background: transparent; border-radius: 0; padding: 0;")
        else:
            self.icon_label.setStyleSheet("background: transparent; border-radius: 0; padding: 0;")
        
        # Update text style
        self.text_label.setStyleSheet(f"""
            QLabel {{
                font-size: 11px;
                font-weight: 400;
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
                color: rgba(255, 255, 255, {fg_alpha});
                background: transparent;
            }}
        """)
        
        # Sync palette for MarqueeLabel's custom painting
        palette = self.text_label.palette()
        palette.setColor(QPalette.WindowText, QColor(255, 255, 255, int(255 * fg_alpha)))
        self.text_label.setPalette(palette)
        
        self.text_label.setVisible(show_text)
        
        if show_text:
            self.layout.setContentsMargins(10, 0, 12, 0)
            self.layout.setSpacing(6)
            # Constrain text width to trigger marquee if needed
            self.text_label.setFixedWidth(46)
        else:
            self.layout.setContentsMargins(0, 0, 0, 0)
            self.layout.setSpacing(0)
            self.text_label.setFixedWidth(0) # Hide effectively if needed, though setVisible handles it

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

class SlidePreviewPopup(FluentWidget):
    def __init__(self, parent=None, monitor=None, is_light=False):
        super().__init__(parent=parent)
        self.monitor = monitor
        self._is_light = is_light
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint | Qt.Tool)
        self.cards = []
        self.slide_indices = []
        self.current_index = 0
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)
        bg = "white" if self._is_light else "rgb(32, 32, 32)"
        border = "rgba(0, 0, 0, 0.12)" if self._is_light else "rgba(255, 255, 255, 0.18)"
        fg = "#191919" if self._is_light else "#FFFFFF"
        self.setStyleSheet(f"""
            SlidePreviewPopup {{
                background-color: {bg};
                border-radius: 10px;
                border: 1px solid {border};
            }}
            QLabel {{
                color: {fg};
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
            }}
        """)
        self.card_container = QWidget(self)
        self.card_layout = QHBoxLayout(self.card_container)
        self.card_layout.setContentsMargins(0, 0, 0, 0)
        self.card_layout.setSpacing(6)
        layout.addWidget(self.card_container)
        self.page_label = QLabel(self)
        self.page_label.setAlignment(Qt.AlignCenter)
        self.page_label.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.page_label)
        self._load_slides()
        self._update_page_label()

    def _load_slides(self):
        if not self.monitor or not hasattr(self.monitor, "get_total_slides"):
            return
        total = self.monitor.get_total_slides()
        if not total:
            return
        temp_dir = os.path.join(tempfile.gettempdir(), "kazuha_ppt_thumbs")
        os.makedirs(temp_dir, exist_ok=True)
        for slide_num in range(1, total + 1):
            path = os.path.join(temp_dir, f"slide_{slide_num}.png")
            try:
                if hasattr(self.monitor, "export_slide_thumbnail"):
                    self.monitor.export_slide_thumbnail(slide_num, path)
            except Exception:
                continue
            pix = QPixmap(path)
            if pix.isNull():
                continue
            btn = QPushButton(self.card_container)
            btn.setFlat(True)
            btn.setIcon(QIcon(pix))
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                "QPushButton { border-radius: 14px; border: 1px solid rgba(0,0,0,0.05); background-color: #F5F6F8; }"
                "QPushButton:hover { border: 1px solid rgba(50,117,245,0.5); background-color: #FFFFFF; }"
            )
            index_in_row = len(self.cards)
            btn.clicked.connect(lambda _, idx=index_in_row: self._on_card_clicked(idx))
            self.card_layout.addWidget(btn)
            self.cards.append(btn)
            self.slide_indices.append(slide_num)
        self._update_cards()

    def _update_cards(self):
        for idx, btn in enumerate(self.cards):
            w, h = 180, 110
            btn.setFixedSize(w, h)
            btn.setIconSize(QSize(w, h))

    def _update_page_label(self):
        total = len(self.cards)
        if total == 0:
            self.page_label.setText("")
        else:
            self.page_label.setText(f"{self.current_index + 1}/{total}")

    def _go_prev(self):
        if not self.cards:
            return
        if self.current_index > 0:
            self.current_index -= 1
            self._update_cards()
            self._update_page_label()

    def _go_next(self):
        if not self.cards:
            return
        if self.current_index < len(self.cards) - 1:
            self.current_index += 1
            self._update_cards()
            self._update_page_label()

    def _activate_current(self):
        if not self.cards or not self.slide_indices:
            return
        index = self.slide_indices[self.current_index]
        if self.monitor and hasattr(self.monitor, "go_to_slide"):
            self.monitor.go_to_slide(index)
        self.close()

    def _on_card_clicked(self, idx):
        if idx < 0 or idx >= len(self.cards):
            return
        self.current_index = idx
        self._update_cards()
        self._update_page_label()
        self._activate_current()

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        if delta > 0:
            self._go_prev()
        elif delta < 0:
            self._go_next()

    def keyPressEvent(self, event):
        key = event.key()
        if key in (Qt.Key_Left, Qt.Key_A):
            self._go_prev()
        elif key in (Qt.Key_Right, Qt.Key_D):
            self._go_next()
        elif key in (Qt.Key_Return, Qt.Key_Enter):
            self._activate_current()
        elif key == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    def event(self, e):
        if e.type() == QEvent.Gesture:
            ge = e
            if isinstance(ge, QGestureEvent):
                swipe = ge.gesture(Qt.SwipeGesture)
                if isinstance(swipe, QSwipeGesture):
                    if swipe.horizontalDirection() == QSwipeGesture.Left:
                        self._go_next()
                        return True
                    if swipe.horizontalDirection() == QSwipeGesture.Right:
                        self._go_prev()
                        return True
        return super().event(e)

class OverlayWindow(QWidget):
    request_next = Signal()
    request_prev = Signal()
    request_end = Signal()
    request_ptr_arrow = Signal()
    request_ptr_pen = Signal()
    request_ptr_eraser = Signal()
    request_pen_color = Signal(int, int, int)
    
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        # Add this attribute to fix UpdateLayeredWindowIndirect error on some systems
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WA_PaintOnScreen, False)
        
        icon_path = os.path.join(ICON_DIR, "overlayicon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.setWindowTitle(_t("overlay.title"))
        self.status_bar = None
        self.plugins = []
        self.monitor = None
        self.slide_preview = None
        self._dev_watermark = None
        version = _get_app_version()
        if _is_dev_preview_version(version):
            label = QLabel(self)
            label.setText(f"开发中版本/技术预览版本\n不保证最终品质 （{version}）")
            font = QFont()
            font.setPixelSize(11)
            label.setFont(font)
            label.setAlignment(Qt.AlignRight | Qt.AlignBottom)
            label.setStyleSheet(
                "color: rgba(255, 255, 255, 120);"
            )
            label.resize(340, 36)
            self._dev_watermark = label
        self.load_plugins()
        self.init_ui()
        self.update_layout()

    def bind_monitor_signals(self):
        if self.monitor:
            return

    def set_monitor(self, monitor):
        self.monitor = monitor
        if self.status_bar:
            self.status_bar.set_monitor(monitor)
        self.bind_monitor_signals()

    def show_slide_preview(self, global_pos=None):
        if not self.monitor:
            return
        self.slide_preview = SlidePreviewPopup(self, self.monitor, getattr(self, "_is_light", True))
        self.slide_preview.adjustSize()
        if isinstance(global_pos, QPoint):
            screen = QGuiApplication.screenAt(global_pos)
            if not screen:
                screen = QGuiApplication.primaryScreen()
            geo = screen.availableGeometry()
            w = self.slide_preview.width()
            h = self.slide_preview.height()
            x = global_pos.x() - w // 2
            y = global_pos.y() - h - 12
            if x < geo.left():
                x = geo.left() + 8
            if x + w > geo.right():
                x = geo.right() - w - 8
            if y < geo.top():
                y = global_pos.y() + 12
            self.slide_preview.move(x, y)
        else:
            center = self.rect().center()
            center_global = self.mapToGlobal(center)
            x = center_global.x() - self.slide_preview.width() // 2
            y = center_global.y() - self.slide_preview.height() // 2
            self.slide_preview.move(x, y)
        self.slide_preview.show()
        self.slide_preview.raise_()
        self.slide_preview.setFocus()

    def load_plugins(self):
        if not os.path.exists(PLUGIN_DIR):
            return
        for entry in os.listdir(PLUGIN_DIR):
            plugin_dir = os.path.join(PLUGIN_DIR, entry)
            if not os.path.isdir(plugin_dir):
                continue
            manifest_path = os.path.join(plugin_dir, "manifest.json")
            if not os.path.exists(manifest_path):
                continue
            try:
                with open(manifest_path, "r", encoding="utf-8") as f:
                    manifest = json.load(f)
                entry_point = manifest.get("entry")
                if not entry_point:
                    continue
                module_name, class_name = entry_point.rsplit(".", 1)
                module_path = os.path.join(plugin_dir, module_name + ".py")
                if not os.path.exists(module_path):
                    continue
                spec = importlib.util.spec_from_file_location(f"plugins.builtins.{entry}.{module_name}", module_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                plugin_cls = getattr(module, class_name, None)
                if plugin_cls is None:
                    continue
                plugin_instance = plugin_cls()
                plugin_instance.manifest = manifest
                plugin_instance.set_context(self)
                self.plugins.append(plugin_instance)
                print(f"Loaded plugin: {plugin_instance.get_name()}")
            except Exception as e:
                print(f"Error loading plugin {entry}: {e}")

    def init_ui(self):
        self.toolbar_height = 45 # Default/minimum height
        has_status_plugin = any(
            hasattr(p, "manifest")
            and isinstance(p.manifest, dict)
            and p.manifest.get("type") == "status_bar"
            for p in self.plugins
        )
        self._has_status_plugin = has_status_plugin
        if cfg.showStatusBar.value and has_status_plugin:
            self.status_bar = StatusBarWidget(self)
            self.status_bar.show()
        cfg.showStatusBar.valueChanged.connect(self._on_status_bar_visibility_changed)
        
        self.toolbar = ToolbarWidget(self, self.plugins) 
        self.toolbar.prev_clicked.connect(self.request_prev.emit)
        self.toolbar.next_clicked.connect(self.request_next.emit)
        self.toolbar.end_clicked.connect(self.request_end.emit)
        
        self.toolbar.select_clicked.connect(self.request_ptr_arrow.emit)
        self.toolbar.pen_clicked.connect(self.request_ptr_pen.emit)
        self.toolbar.eraser_clicked.connect(self.request_ptr_eraser.emit)
        self.toolbar.pen_color_changed.connect(self.request_pen_color.emit)
        
        self.left_flipper = PageFlipWidget("Left", self, height=self.toolbar_height)
        self.left_flipper.clicked_prev.connect(self.request_prev.emit)
        self.left_flipper.clicked_next.connect(self.request_next.emit)
        self.left_flipper.page_clicked.connect(self.show_slide_preview)
        
        self.right_flipper = PageFlipWidget("Right", self, height=self.toolbar_height)
        self.right_flipper.clicked_prev.connect(self.request_prev.emit)
        self.right_flipper.clicked_next.connect(self.request_next.emit)
        self.right_flipper.page_clicked.connect(self.show_slide_preview)

        # Connect adaptive theme signal
        if self.status_bar:
            self.status_bar.is_light_changed.connect(self._on_theme_changed)
        
        self.left_flipper.show()
        self.right_flipper.show()
        self.toolbar.show()
        
        self._update_theme_from_cfg()
        cfg.themeMode.valueChanged.connect(self._update_theme_from_cfg)
        cfg.showToolbarText.valueChanged.connect(self.update_layout)

    def _update_theme_from_cfg(self):
        theme = cfg.themeMode.value
        if theme == "Auto":
            from qfluentwidgets import isDarkTheme
            is_light = not isDarkTheme()
        else:
            is_light = (theme == "Light")
        self._on_theme_changed(is_light)

    def _on_theme_changed(self, is_light):
        self._is_light = is_light
        self.toolbar.update_style(is_light)
        self.left_flipper.update_style(is_light)
        self.right_flipper.update_style(is_light)
        if self.status_bar:
            self.status_bar._update_palette(is_light)

    def _on_status_bar_visibility_changed(self, visible: bool):
        if visible and getattr(self, "_has_status_plugin", False):
            if self.status_bar is None:
                self.status_bar = StatusBarWidget(self)
                if self.monitor:
                    self.status_bar.set_monitor(self.monitor)
            self.status_bar.show()
        else:
            if self.status_bar is not None:
                self.status_bar.hide()
        self.update_layout()

    def showEvent(self, event):
        super().showEvent(event)
        self.update_layout()

    def cleanup(self):
        """Clean up resources and terminate plugins."""
        for plugin in self.plugins:
            try:
                plugin.terminate()
            except Exception as e:
                print(f"Error terminating plugin {plugin.get_name()}: {e}")

    def update_toolbar(self):
        if hasattr(self, 'toolbar') and self.toolbar:
            self.toolbar.refresh_dynamic_tools()

    def update_layout(self):
        # Prevent crash during initialization
        if not hasattr(self, "toolbar") or self.toolbar is None:
            return
        if not hasattr(self, "left_flipper") or not hasattr(self, "right_flipper"):
            return
            
        w = self.width()
        h = self.height()
        margin = 16
        
        if w <= 100 or h <= 100:
            return

        # Let the toolbar determine its own ideal size based on content
        self.toolbar.adjustSize()
        tb_size = self.toolbar.size()
        tb_w = tb_size.width()
        tb_h = tb_size.height()
        
        # Flipper height matches toolbar height for consistency
        flipper_w = 140
        self.left_flipper.setFixedSize(flipper_w, tb_h)
        self.right_flipper.setFixedSize(flipper_w, tb_h)
        # Re-trigger style update to ensure radius is correct
        self.left_flipper.h_val = tb_h
        self.right_flipper.h_val = tb_h
        self.left_flipper.update_style(getattr(self, "_is_light", False))
        self.right_flipper.update_style(getattr(self, "_is_light", False))

        y_pos = h - tb_h - 14 
        
        self.toolbar.move((w - tb_w) // 2, y_pos)
        self.left_flipper.move(margin, y_pos)
        self.right_flipper.move(w - flipper_w - margin, y_pos)

        if self.status_bar and self.status_bar.isVisible():
            self.status_bar.setFixedWidth(w)
            self.status_bar.move(0, 0)

        if self._dev_watermark:
            self._dev_watermark.adjustSize()
            wm_w = self._dev_watermark.width()
            wm_h = self._dev_watermark.height()
            self._dev_watermark.move(w - wm_w - 16, h - wm_h - 12)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_layout()

    # paintEvent removed to ensure full transparency of the overlay container

    def update_page_info(self, current, total):
        self.left_flipper.set_page_info(current, total)
        self.right_flipper.set_page_info(current, total)
        
        # Force update to clean artifacts if any
        self.update()

class ToolbarWidget(QFrame):
    prev_clicked = Signal()
    next_clicked = Signal()
    end_clicked = Signal()
    select_clicked = Signal()
    pen_clicked = Signal()
    eraser_clicked = Signal()
    pen_color_changed = Signal(int, int, int)

    def __init__(self, parent=None, plugins=[], height=64):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.plugins = plugins
        self.current_tool = "select"
        self.pen_popup = None
        self._is_light = False
        self._dynamic_widgets = []

        self.indicator = QFrame(self)
        self.indicator.setObjectName("Indicator")
        self.indicator.hide()
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)
        
        self.init_ui()
        self.update_layout_style()
        
        cfg.showToolbarText.valueChanged.connect(self.update_layout_style)
    
    def sizeHint(self):
        return self.layout.sizeHint()

    def update_layout_style(self):
        show_text = cfg.showToolbarText.value
        
        # Match HTML: .toolbar-preview-bar { background: rgba(28, 28, 30, 0.82); border: 0.5px solid rgba(255, 255, 255, 0.15); }
        bg = "rgba(28, 28, 30, 0.82)" 
        border = "rgba(255, 255, 255, 0.15)"
        line_color = "rgba(255, 255, 255, 0.15)"

        self.setStyleSheet(f"""
            ToolbarWidget {{
                background-color: {bg};
                border: 0.5px solid {border};
            }}
        """)
        
        # Update children
        for line in self.findChildren(QFrame):
            if line.frameShape() == QFrame.VLine:
                line.setFixedWidth(1)
                line.setStyleSheet(f"background-color: {line_color}; border: none; margin: 10px 0;")
                
        for btn in self.findChildren(CustomToolButton):
            btn.update_size()
            btn.update_style(btn.tool_name == self.current_tool, False)
        
        # Force layout recalculation and update radius
        self.layout.activate()
        hint = self.layout.sizeHint()
        # Match HTML: .toolbar-preview-bar { border-radius: 999px; }
        radius = hint.height() // 2
        
        self.setStyleSheet(self.styleSheet() + f"\nToolbarWidget {{ border-radius: {radius}px; }}")
        
        # Match HTML: .toolbar-preview-bar { box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4); }
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 12)
        self.setGraphicsEffect(shadow)

        # Notify parent to update position
        p = self.parent()
        if p and hasattr(p, "update_layout"):
            p.update_layout()

    def update_style(self, is_light=False):
        # Compatibility with old calls
        self._is_light = is_light
        self.update_layout_style()

    def _on_undo_redo_visibility_changed(self, visible: bool):
        self.btn_undo.setVisible(visible)
        self.btn_redo.setVisible(visible)
        self.line1.setVisible(visible)
        QTimer.singleShot(10, self._update_indicator_now)

    def _update_indicator_now(self):
        pass # Indicator is hidden in new design

    def _ensure_pen_popup(self):
        if self.pen_popup is None:
            self.pen_popup = PenColorPopup(self.window(), is_light=self._is_light)
            self.pen_popup.color_selected.connect(self._on_pen_color_selected)

    def _position_pen_popup(self):
        if not self.pen_popup:
            return
        self.pen_popup.adjustSize()
        btn_center = self.btn_pen.mapToGlobal(self.btn_pen.rect().center())
        # Use parent window to position popup
        x = btn_center.x() - self.pen_popup.width() // 2
        y = btn_center.y() - self.btn_pen.height() // 2 - self.pen_popup.height() - 12
        self.pen_popup.move(x, y)

    def _toggle_pen_popup(self):
        self._ensure_pen_popup()
        if self.pen_popup.isVisible():
            self.pen_popup.hide()
        else:
            self._position_pen_popup()
            self.pen_popup.show()

    def _on_pen_button_clicked(self):
        if self.current_tool == "pen":
            self._toggle_pen_popup()
        else:
            self._on_tool_changed("pen", self.pen_clicked)

    def _on_pen_color_selected(self, r, g, b):
        self.pen_color_changed.emit(r, g, b)
        # In new version, PenColorPopup.close() is called inside popup
        # Re-sync style if needed
        self.btn_pen.update_style(True, self._is_light, use_indicator=True)

    def _on_tool_changed(self, tool_name, signal):
        old_tool = self.current_tool
        self.current_tool = tool_name
        
        for btn in [self.btn_select, self.btn_pen, self.btn_eraser]:
            if isinstance(btn, CustomToolButton):
                btn.update_style(btn.tool_name == self.current_tool, self._is_light, use_indicator=True)
        
        if tool_name != "pen" and self.pen_popup and self.pen_popup.isVisible():
            self.pen_popup.hide()
        if signal:
            signal.emit()

    def init_ui(self):
        self.layout = QHBoxLayout(self)
        # Match HTML: .toolbar-preview-bar { padding: 8px 12px; gap: 4px; }
        self.layout.setContentsMargins(12, 8, 12, 8) 
        self.layout.setSpacing(4) 
        self.layout.setSizeConstraint(QHBoxLayout.SetFixedSize)
        self.layout.setAlignment(Qt.AlignCenter)

        # Base Tools
        self.btn_select = CustomToolButton("Mouse.svg", _t("toolbar.select"), self, tool_name="select", text=_t("toolbar.select"))
        self.btn_select.clicked.connect(lambda: self._on_tool_changed("select", self.select_clicked))
        self.layout.addWidget(self.btn_select)

        self.btn_pen = CustomToolButton("Pen.svg", _t("toolbar.pen"), self, tool_name="pen", text=_t("toolbar.pen"))
        self.btn_pen.clicked.connect(self._on_pen_button_clicked)
        self.layout.addWidget(self.btn_pen)

        self.btn_eraser = CustomToolButton("Eraser.svg", _t("toolbar.eraser"), self, tool_name="eraser", text=_t("toolbar.eraser"))
        self.btn_eraser.clicked.connect(lambda: self._on_tool_changed("eraser", self.eraser_clicked))
        self.layout.addWidget(self.btn_eraser)

        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.VLine)
        self.line1.setFixedHeight(24)
        self.layout.addWidget(self.line1)

        self.btn_undo = CustomToolButton("Previous.svg", _t("toolbar.undo"), self, text=_t("toolbar.undo"))
        self.btn_undo.clicked.connect(self.prev_clicked.emit)
        self.layout.addWidget(self.btn_undo)

        self.btn_redo = CustomToolButton("Next.svg", _t("toolbar.redo"), self, text=_t("toolbar.redo"))
        self.btn_redo.clicked.connect(self.next_clicked.emit)
        self.layout.addWidget(self.btn_redo)

        self.btn_undo.setVisible(cfg.showUndoRedo.value)
        self.btn_redo.setVisible(cfg.showUndoRedo.value)
        self.line1.setVisible(cfg.showUndoRedo.value)
        cfg.showUndoRedo.valueChanged.connect(self._on_undo_redo_visibility_changed)

        # Plugins and Dynamic Content
        self.dynamic_container = QWidget(self)
        self.dynamic_layout = QHBoxLayout(self.dynamic_container)
        self.dynamic_layout.setContentsMargins(0, 0, 0, 0)
        self.dynamic_layout.setSpacing(4)
        self.layout.addWidget(self.dynamic_container)
        
        self.refresh_dynamic_tools()

        self.line3 = QFrame()
        self.line3.setFrameShape(QFrame.VLine)
        self.line3.setFixedHeight(24)
        self.layout.addWidget(self.line3)

        self.btn_end = CustomToolButton("Minimize.svg", _t("toolbar.end_show"), self, is_exit=True, text=_t("toolbar.end_show"))
        self.btn_end.clicked.connect(self.end_clicked.emit)
        self.layout.addWidget(self.btn_end)
        
        QTimer.singleShot(0, self._update_indicator_now)

    def refresh_dynamic_tools(self):
        while self.dynamic_layout.count():
            item = self.dynamic_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        toolbar_plugins = []
        app_launcher = None
        for plugin in self.plugins:
            plugin_type = "toolbar"
            if hasattr(plugin, "get_type"):
                plugin_type = plugin.get_type()

            name = ""
            if hasattr(plugin, "get_name"):
                name = plugin.get_name()

            if isinstance(name, str) and ("工具栏固定项" in name or "应用启动器" in name or "App Launcher" in name):
                app_launcher = plugin
                continue

            if isinstance(plugin_type, str) and plugin_type.startswith("toolbar"):
                toolbar_plugins.append(plugin)

        if toolbar_plugins:
            line = QFrame()
            line.setFrameShape(QFrame.VLine)
            line.setFixedHeight(24)
            self.dynamic_layout.addWidget(line)
            for plugin in toolbar_plugins:
                btn = CustomToolButton(plugin.get_icon() or "More.svg", plugin.get_name(), self, text=plugin.get_name())
                btn.clicked.connect(plugin.execute)
                self.dynamic_layout.addWidget(btn)

        # 2. Add Quick Launch Apps
        apps = cfg.quickLaunchApps.value
        if apps:
            line = QFrame()
            line.setFrameShape(QFrame.VLine)
            line.setFixedHeight(24)
            self.dynamic_layout.addWidget(line)
            for app in apps:
                app_path = app['path']
                pixmap = None
                if app_launcher:
                    pixmap = app_launcher.get_app_icon(app_path)
                
                btn = CustomToolButton("More.svg", app['name'], self, text=app['name'], pixmap=pixmap)
                btn.clicked.connect(lambda checked=False, p=app_path: app_launcher.execute_app(p) if app_launcher else None)
                self.dynamic_layout.addWidget(btn)

        self.adjustSize()
        if hasattr(self.parent(), 'update_layout'):
            self.parent().update_layout()

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._update_indicator_now)

class PageFlipButton(QFrame):
    btn_clicked = Signal()

    def __init__(self, icon_name, parent=None):
        super().__init__(parent)
        # Match CustomToolButton circular size
        self.setFixedSize(38, 38)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("""
            QFrame {
                background-color: transparent;
                border-radius: 19px;
            }
            QFrame:hover {
                background-color: rgba(255, 255, 255, 0.08);
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.setAlignment(Qt.AlignCenter)
        
        self.icon_label = QLabel(self)
        # Match CustomToolButton circular icon size
        self.icon_label.setFixedSize(20, 20)
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        icon_path = os.path.join(ICON_DIR, icon_name)
        if os.path.exists(icon_path):
            renderer = QSvgRenderer(icon_path)
            if renderer.isValid():
                pixmap = QPixmap(40, 40)
                pixmap.fill(Qt.transparent)
                painter = QPainter(pixmap)
                painter.setRenderHint(QPainter.Antialiasing)
                renderer.render(painter)
                painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
                painter.fillRect(pixmap.rect(), QColor(255, 255, 255))
                painter.end()
                self.icon_label.setPixmap(pixmap.scaled(20, 20, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        
        layout.addWidget(self.icon_label)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.btn_clicked.emit()
        super().mousePressEvent(event)

class PageFlipWidget(QFrame):
    clicked_prev = Signal()
    clicked_next = Signal()
    page_clicked = Signal(QPoint)

    def __init__(self, side="Left", parent=None, height=56):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.side = side
        self.h_val = height
        
        self.update_style()
        self.setFixedSize(160, self.h_val)
        
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(5, 0, 5, 0)
        self.layout.setSpacing(0)
        
        self.btn_prev = PageFlipButton("Previous.svg", self)
        self.btn_prev.btn_clicked.connect(self.clicked_prev.emit)
        
        self.btn_next = PageFlipButton("Next.svg", self)
        self.btn_next.btn_clicked.connect(self.clicked_next.emit)
        
        # Center container that takes all remaining space
        self.page_container = QFrame(self)
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setContentsMargins(0, 0, 0, 0)
        self.page_layout.setSpacing(0)
        self.page_layout.setAlignment(Qt.AlignCenter)
        
        self.lbl_page = ClickableLabel("0/0", self.page_container)
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.lbl_page.clicked.connect(self.page_clicked.emit)
        
        self.lbl_hint = QLabel(_t("toolbar.page"), self.page_container)
        self.lbl_hint.setAlignment(Qt.AlignCenter)
        self.lbl_hint.setObjectName("PageHint")
        self.lbl_hint.setVisible(True)
        
        self.page_layout.addWidget(self.lbl_page)
        self.page_layout.addWidget(self.lbl_hint)
        
        self.layout.addWidget(self.btn_prev)
        self.layout.addWidget(self.page_container, 1) # Stretch factor 1
        self.layout.addWidget(self.btn_next)

    def _on_show_text_changed(self, show):
        # Hint is now always visible, no need to toggle
        pass

    def set_page_info(self, current, total):
        hint_fg = "rgba(255, 255, 255, 0.6)" if not hasattr(self, "_is_light") or not self._is_light else "rgba(0, 0, 0, 0.5)"
        self.lbl_page.setText(f'<span style="font-size: 16px; font-weight: 900;">{current}</span>'
                              f'<span style="font-size: 10px; font-weight: 400; color: {hint_fg};">/{total}</span>')

    def update_style(self, is_light=False):
        # Match HTML toolbar preview design
        self._is_light = is_light
        bg = "rgba(28, 28, 30, 0.82)"
        border = "rgba(255, 255, 255, 0.15)"
        fg = "white"
        hint_fg = "rgba(255, 255, 255, 0.6)"
        
        # Calculate radius to ensure it's always a capsule (pill shape)
        radius = self.h_val // 2
        
        self.setStyleSheet(f"""
            PageFlipWidget {{
                background-color: {bg};
                border-radius: {radius}px;
                border: 0.5px solid {border};
            }}
            QLabel {{
                color: {fg};
                font-family: 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', 'PingFang SC', sans-serif;
                font-size: 12px;
                font-weight: 900;
                background: transparent;
                border: none;
            }}
            QLabel#PageHint {{
                font-size: 9px;
                font-weight: 400;
                color: {hint_fg};
            }}
        """)

        # Match HTML: .toolbar-preview-bar { box-shadow: 0 12px 40px rgba(0, 0, 0, 0.4); }
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 12)
        self.setGraphicsEffect(shadow)
