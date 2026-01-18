from PySide6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QFrame, QApplication, QLabel, QPushButton, QSwipeGesture, QGestureEvent, QGridLayout, QStyleOption, QStyle, QGraphicsDropShadowEffect, QMenu
from PySide6.QtCore import Qt, Signal, QSize, QPoint, QEvent, QTimer, QTime, QDateTime, QLocale, QThread, QObject, QPropertyAnimation, QEasingCurve, QParallelAnimationGroup, QRect
from PySide6.QtGui import QColor, QIcon, QPainter, QBrush, QPen, QPixmap, QGuiApplication, QFont, QPalette, QLinearGradient, QAction, QRegion
from PySide6.QtSvg import QSvgRenderer
import os
import importlib.util
import sys
import tempfile
import json
import subprocess
import time
import math
import shiboken6
from ppt_assistant.core.config import cfg, SETTINGS_PATH
from ppt_assistant.core.timer_manager import TimerManager
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


def _get_overlay_font_stack():
    base = "'SF Pro', '苹方-简', 'PingFang SC', 'MiSans Latin', 'Segoe UI', 'Microsoft YaHei', sans-serif"
    try:
        if os.path.exists(SETTINGS_PATH):
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            lang = data.get("General", {}).get("Language", "zh-CN")
            fonts = data.get("Fonts", {}) or {}
            profiles = fonts.get("Profiles", {}) or {}
            lang_profile = profiles.get(lang, {}) or {}
            v = lang_profile.get("overlay") or lang_profile.get("qt") or ""
            if isinstance(v, str):
                v = v.strip()
            else:
                v = ""
            if not v:
                return base
            safe = v.replace("'", "\\'")
            return f"'{safe}', {base}"
    except Exception:
        return base
    return base


LANGUAGE = _load_language()

_TRANSLATIONS = {
    "zh-CN": {
        "status.media_length": "媒体时长",
        "overlay.title": "顶层效果窗口",
        "toolbar.select": "选择",
        "toolbar.pen": "画笔",
        "toolbar.eraser": "橡皮",
        "toolbar.spotlight": "聚光灯",
        "toolbar.timer": "计时器",
        "toolbar.undo": "上一步",
        "toolbar.redo": "下一步",
        "toolbar.end_show": "结束放映",
        "toolbar.page": "页码",
        "toolbar.theme_colors": "主题颜色",
        "toolbar.standard_colors": "标准颜色",
        "overlay.dev_watermark": "开发中版本/技术预览版本\n不保证最终品质 （{version}）",
    },
    "zh-TW": {
        "status.media_length": "媒體時長",
        "overlay.title": "頂層效果視窗",
        "toolbar.select": "選取",
        "toolbar.pen": "畫筆",
        "toolbar.eraser": "橡皮擦",
        "toolbar.spotlight": "聚光燈",
        "toolbar.timer": "計時器",
        "toolbar.undo": "上一步",
        "toolbar.redo": "下一步",
        "toolbar.end_show": "結束播放",
        "toolbar.page": "頁碼",
        "toolbar.theme_colors": "主題顏色",
        "toolbar.standard_colors": "标准颜色",
        "overlay.dev_watermark": "開發中版本/技術預覽版本\n不保證最終品質 （{version}）",
    },
    "ja-JP": {
        "status.media_length": "メディア長さ",
        "overlay.title": "オーバーレイウィンドウ",
        "toolbar.select": "選択",
        "toolbar.pen": "ペン",
        "toolbar.eraser": "消しゴム",
        "toolbar.spotlight": "スポットライト",
        "toolbar.timer": "タイマー",
        "toolbar.undo": "戻る",
        "toolbar.redo": "進む",
        "toolbar.end_show": "スライド終了",
        "toolbar.page": "ページ番号",
        "toolbar.theme_colors": "テーマの色",
        "toolbar.standard_colors": "標準の色",
        "overlay.dev_watermark": "開発中バージョン/テクニカルプレビュー\n品質は保証されません （{version}）",
    },
    "en-US": {
        "status.media_length": "Media duration",
        "overlay.title": "Overlay window",
        "toolbar.select": "Select",
        "toolbar.pen": "Pen",
        "toolbar.eraser": "Eraser",
        "toolbar.spotlight": "Spotlight",
        "toolbar.timer": "Timer",
        "toolbar.undo": "Undo",
        "toolbar.redo": "Redo",
        "toolbar.end_show": "End Show",
        "toolbar.page": "Page number",
        "toolbar.theme_colors": "Theme Colors",
        "toolbar.standard_colors": "Standard Colors",
        "overlay.dev_watermark": "In-Development/Technical Preview\nFinal quality not guaranteed ({version})",
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


class GlobalIconCache:
    _cache = {}

    @classmethod
    def get(cls, key):
        return cls._cache.get(key)

    @classmethod
    def set(cls, key, pixmap):
        cls._cache[key] = pixmap

class NetworkCheckThread(QThread):
    status_changed = Signal(str)

    def run(self):
        while not self.isInterruptionRequested():
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
                    # Non-blocking check in thread
                    if sys.platform == "win32":
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        out = subprocess.check_output(
                            ["netsh", "wlan", "show", "interfaces"],
                            encoding="utf-8",
                            errors="ignore",
                            startupinfo=startupinfo
                        )
                        if "state" in out.lower() and "connected" in out.lower():
                            kind = "wifi"
                except Exception:
                    pass
            
            self.status_changed.emit(kind)
            self.sleep(5)

class StatusBarWidget(QFrame):
    is_light_changed = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        scale = cfg.scale.value
        self.setFixedHeight(int(30 * scale))
        self._is_light = False
        self._monitor = None
        self._network_kind = "offline"
        self._volume_supported = False
        self._timer_manager = TimerManager()
        self._build_ui()
        
        # Single master timer for UI updates
        self._master_timer = QTimer(self)
        self._master_timer.timeout.connect(self._on_master_tick)
        self._master_timer.start(500) # 2Hz update rate
        
        # Network check in background thread
        self._network_thread = NetworkCheckThread(self)
        self._network_thread.status_changed.connect(self._on_network_status_changed)
        self._network_thread.start()

        # Connect to timer manager
        self._timer_manager.updated.connect(self._update_countdown)
        self._timer_manager.state_changed.connect(self._on_timer_state_changed)

        self._update_time()
        self._update_palette()
        self._update_volume()
        self._on_timer_state_changed(self._timer_manager.is_running)

    def _on_master_tick(self):
        self._update_time()
        self._update_video()
        self._update_volume()
        # Color sampling disabled as per user request

    def _on_network_status_changed(self, kind):
        self._network_kind = kind
        if kind == "wifi":
            self.net_icon.setIcon(FIF.WIFI)
        elif kind == "wired":
            icon = getattr(FIF, "ETHERNET", FIF.WIFI)
            self.net_icon.setIcon(icon)
        else:
            self.net_icon.setIcon(FIF.WIFI)

    def closeEvent(self, event):
        if self._network_thread.isRunning():
            self._network_thread.requestInterruption()
            self._network_thread.wait()
        super().closeEvent(event)

    def _build_ui(self):
        layout = QHBoxLayout(self)
        scale = cfg.scale.value
        layout.setContentsMargins(int(16 * scale), 0, int(16 * scale), 0)
        layout.setSpacing(int(12 * scale))

        self.time_label = QLabel("", self)
        self.time_label.setObjectName("TimeLabel")
        layout.addWidget(self.time_label)

        # Countdown
        self.countdown_container = QWidget(self)
        countdown_layout = QHBoxLayout(self.countdown_container)
        countdown_layout.setContentsMargins(0, 0, 0, 0)
        countdown_layout.setSpacing(int(8 * scale))
        self.countdown_separator = QFrame(self)
        self.countdown_separator.setFixedWidth(1)
        self.countdown_separator.setFixedHeight(int(12 * scale))
        self.countdown_separator.setObjectName("Separator")
        self.countdown_label = QLabel("", self)
        self.countdown_label.setObjectName("CountdownLabel")
        countdown_layout.addWidget(self.countdown_separator)
        countdown_layout.addWidget(self.countdown_label)
        layout.addWidget(self.countdown_container)
        self.countdown_container.hide()

        # Video progress
        self.video_container = QWidget(self)
        video_layout = QHBoxLayout(self.video_container)
        video_layout.setContentsMargins(0, 0, 0, 0)
        video_layout.setSpacing(int(8 * scale))
        self.separator = QFrame(self)
        self.separator.setFixedWidth(1)
        self.separator.setFixedHeight(int(12 * scale))
        self.separator.setObjectName("Separator")
        self.progress_value = QLabel("", self)
        self.progress_value.setObjectName("ProgressValue")
        self.progress_caption = QLabel(_t("status.media_length"), self)
        self.progress_caption.setObjectName("ProgressCaption")
        video_layout.addWidget(self.separator)
        video_layout.addWidget(self.progress_value)
        video_layout.addWidget(self.progress_caption)
        layout.addWidget(self.video_container)
        self.video_container.hide()

        layout.addStretch(1)

        self.net_icon = IconWidget(FIF.WIFI, self)
        self.net_icon.setFixedSize(int(18 * scale), int(18 * scale))
        layout.addWidget(self.net_icon)

        self.volume_icon = IconWidget(FIF.VOLUME, self)
        self.volume_icon.setFixedSize(int(18 * scale), int(18 * scale))
        layout.addWidget(self.volume_icon)

    def _update_countdown(self, seconds):
        if seconds > 0:
            self.countdown_label.setText(f"倒计时 {self._timer_manager.get_remaining_time_str()}")
            self.countdown_container.show()
        else:
            self.countdown_container.hide()

    def _on_timer_state_changed(self, is_running):
        if is_running:
            self.countdown_container.show()
            self._update_countdown(self._timer_manager.remaining_seconds)
        else:
            self.countdown_container.hide()

    def _update_time(self):
        now = QDateTime.currentDateTime()
        locale = QLocale(QLocale.Chinese, QLocale.China)
        time_str = locale.toString(now, "H:mm M月d日 ddd")
        self.time_label.setText(time_str)

    def set_monitor(self, monitor):
        self._monitor = monitor

    def _update_video(self):
        if not self._monitor:
            self.video_container.hide()
            return
        try:
            ratio, pos, length = self._monitor.get_video_progress()
        except Exception:
            self.video_container.hide()
            return
        if length is None or length <= 0:
            self.video_container.hide()
            return
        length_sec = length or 0.0
        if length_sec > 36000:
            length_sec = length_sec / 1000.0
        total_text = self._format_seconds(length_sec)
        self.progress_value.setText(total_text)
        self.video_container.show()

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

    def cleanup(self):
        if hasattr(self, "_network_thread") and self._network_thread.isRunning():
            self._network_thread.requestInterruption()
            self._network_thread.quit()
            self._network_thread.wait()

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
        fg = "#FFFFFF"
        bg = "#40000000"
        font_stack = _get_overlay_font_stack()
        scale = cfg.scale.value

        self.setStyleSheet(
            f"""
            StatusBarWidget {{
                background-color: {bg};
                border: none;
            }}
            QLabel {{
                color: {fg};
                font-family: {font_stack};
                background: transparent;
            }}
            QLabel#TimeLabel {{
                font-size: {int(13 * scale)}px;
                font-weight: bold;
            }}
            QLabel#ProgressValue, QLabel#ProgressCaption, QLabel#CountdownLabel {{
                font-size: {int(11 * scale)}px;
                font-weight: 500;
            }}
            QFrame#Separator {{
                background-color: rgba(255, 255, 255, 0.3);
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
        font_stack = _get_overlay_font_stack()
        
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
                font-family: {font_stack};
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
        self.text_label.setStyleSheet(f"""
            QLabel {{
                font-size: 8px;
                font-weight: 300;
                font-family: {_get_overlay_font_stack()};
                color: rgba(255, 255, 255, 0.6);
                background: transparent;
            }}
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
        scale = cfg.scale.value
        if show_text:
            # Match HTML: .toolbar-preview-btn-capsule { min-width: 68px; height: 40px; }
            # Increased width to 92 for horizontal layout and marquee room
            self.setFixedSize(int(92 * scale), int(38 * scale))
            self.icon_size = int(18 * scale)
            self.layout.setContentsMargins(int(12 * scale), int(4 * scale), int(12 * scale), int(4 * scale))
        else:
            # Match HTML: .toolbar-preview-btn { min-width: 38px; height: 38px; }
            self.setFixedSize(int(38 * scale), int(38 * scale))
            self.icon_size = int(20 * scale)
            self.layout.setContentsMargins(int(4 * scale), int(4 * scale), int(4 * scale), int(4 * scale))
        
        if show_text:
            self.text_label.setStyleSheet(f"""
                QLabel {{
                    font-size: {int(8 * scale)}px;
                    font-weight: 300;
                    font-family: {_get_overlay_font_stack()};
                    color: rgba(255, 255, 255, 0.6);
                    background: transparent;
                }}
            """)
        
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
            
        color_hex = "#FFFFFF"
        if self.is_exit:
            color_hex = "#FF453A"

        cache_key = (self.icon_name, color_hex, s)
        cached_pixmap = GlobalIconCache.get(cache_key)
        if cached_pixmap:
            self.icon_label.setPixmap(cached_pixmap)
            self.icon_label.setFixedSize(s, s)
            return
            
        color = QColor(color_hex)
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
            # Render icon with red color directly
            renderer.render(painter)
            painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
            painter.fillRect(pixmap.rect(), color)
        else:
            # Normal icon colorization
            renderer.render(painter)
            painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
            painter.fillRect(pixmap.rect(), color)
            
        painter.end()
        
        scaled_pixmap = pixmap.scaled(s, s, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        GlobalIconCache.set(cache_key, scaled_pixmap)
        
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
                font-family: {_get_overlay_font_stack()};
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
                font-family: 'MiSans Latin', 'HarmonyOS Sans SC', 'SF Pro', '苹方-简', 'PingFang SC', 'Segoe UI', 'Microsoft YaHei', sans-serif;
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
        
        self.slide_map.clear()
        
        for slide_num in range(1, total + 1):
            path = os.path.join(temp_dir, f"slide_{slide_num}.png")
            
            # Create button first with placeholder
            btn = QPushButton(self.card_container)
            btn.setFlat(True)
            # Set a simple placeholder or empty icon
            btn.setStyleSheet(
                "QPushButton { border-radius: 14px; border: 1px solid rgba(0,0,0,0.05); background-color: #F5F6F8; }"
                "QPushButton:hover { border: 1px solid rgba(50,117,245,0.5); background-color: #FFFFFF; }"
            )
            
            # Check if exists in cache/disk first for speed
            if os.path.exists(path):
                 pix = QPixmap(path)
                 if not pix.isNull():
                     btn.setIcon(QIcon(pix))
            
            # Request update/export
            try:
                if hasattr(self.monitor, "export_slide_thumbnail"):
                    self.monitor.export_slide_thumbnail(slide_num, path)
            except Exception:
                pass

            index_in_row = len(self.cards)
            btn.clicked.connect(lambda _, idx=index_in_row: self._on_card_clicked(idx))
            self.card_layout.addWidget(btn)
            self.cards.append(btn)
            self.slide_indices.append(slide_num)
            self.slide_map[slide_num] = btn
            
        self._update_cards()

    def _on_thumbnail_generated(self, index, path):
        if index in self.slide_map:
            btn = self.slide_map[index]
            if os.path.exists(path):
                pix = QPixmap(path)
                if not pix.isNull():
                    w, h = 180, 110
                    btn.setIcon(QIcon(pix))
                    btn.setIconSize(QSize(w, h))

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

class IndeterminateSpinner(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self.setFixedSize(27, 27)

    def start(self):
        if not self._timer.isActive():
            self._timer.start(16)

    def stop(self):
        if self._timer.isActive():
            self._timer.stop()

    def _tick(self):
        self._angle = (self._angle + 5) % 360
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        self.start()

    def hideEvent(self, event):
        self.stop()
        super().hideEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        center_x = self.width() / 2
        center_y = self.height() / 2

        ring_color = QColor("#d9d9d9")
        pen = QPen(ring_color)
        pen.setWidth(3)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)

        ring_radius = 10
        painter.drawEllipse(QPoint(int(center_x), int(center_y)), ring_radius, ring_radius)

        dot_radius = 2.5
        orbit_radius = ring_radius - 3 - dot_radius + 1
        angle_rad = math.radians(self._angle)
        dot_x = center_x + orbit_radius * math.cos(angle_rad)
        dot_y = center_y + orbit_radius * math.sin(angle_rad)

        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(ring_color))
        painter.drawEllipse(QPoint(int(dot_x), int(dot_y)), dot_radius, dot_radius)

class ReloadMask(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        font_stack = _get_overlay_font_stack()
        self.setStyleSheet(
            "ReloadMask { background-color: rgba(0, 0, 0, 110); }"
            "QFrame { background-color: rgba(30, 30, 30, 220); border-radius: 16px; }"
            f"QLabel {{ color: rgba(255, 255, 255, 0.92); font-size: 14px; font-weight: 500; font-family: {font_stack}; }}"
        )

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.setAlignment(Qt.AlignCenter)

        card = QFrame(self)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(22, 18, 22, 18)
        card_layout.setSpacing(12)
        card_layout.setAlignment(Qt.AlignCenter)

        self.spinner = IndeterminateSpinner(card)
        card_layout.addWidget(self.spinner, 0, Qt.AlignCenter)

        self.label = QLabel("正在重载页面", card)
        self.label.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(self.label, 0, Qt.AlignCenter)

        outer.addWidget(card, 0, Qt.AlignCenter)

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
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WA_PaintOnScreen, False)
        
        scr = QGuiApplication.primaryScreen()
        if scr:
            self.setGeometry(scr.geometry())
        
        icon_path = os.path.join(ICON_DIR, "overlayicon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.setWindowTitle(_t("overlay.title"))
        self.status_bar = None
        self.plugins = []
        self.monitor = None
        self.slide_widgets = {} # { slide_index: [widgets] }
        self._current_page_idx = 1
        self.slide_preview = None
        self._dev_watermark = None
        self._wheel_acc = 0 # For OverlayWindow scroll handling
        self._reload_mask = None
        self._pending_timers = []
        version = _get_app_version()
        if _is_dev_preview_version(version):
            label = QLabel(self)
            label.setText(_t("overlay.dev_watermark").format(version=version))
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

    def update_geometry(self, rect, screen):
        target_screen = screen
        if not target_screen and rect and not rect.isEmpty():
            s = QGuiApplication.screenAt(rect.center())
            if not s:
                s = QGuiApplication.primaryScreen()
            target_screen = s
        if target_screen:
            geo = target_screen.geometry()
            if not geo.isEmpty():
                if self.windowHandle():
                    self.windowHandle().setScreen(target_screen)
                self.setGeometry(geo)
        QTimer.singleShot(0, self.update_layout)

    def set_monitor(self, monitor):
        self.monitor = monitor
        if self.status_bar:
            self.status_bar.set_monitor(monitor)
        self.bind_monitor_signals()

    def show_reload_mask(self, text="正在重载页面"):
        if self._reload_mask is None:
            self._reload_mask = ReloadMask(self)
            self._reload_mask.setGeometry(self.rect())
        if text:
            self._reload_mask.label.setText(str(text))
        self._reload_mask.show()
        self._reload_mask.raise_()
        self._reload_mask.activateWindow()

    def hide_reload_mask(self):
        if self._reload_mask is not None:
            self._reload_mask.hide()

    def _defer(self, callback):
        t = QTimer(self)
        t.setSingleShot(True)
        self._pending_timers.append(t)
        def _run():
            try:
                callback()
            finally:
                if t in self._pending_timers:
                    self._pending_timers.remove(t)
                t.deleteLater()
        t.timeout.connect(_run)
        t.start(0)

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
        
        self._init_flippers()
        
        self.update_layout()

    def showEvent(self, event):
        super().showEvent(event)
        self.update_layout()
        # Animation disabled for instant feedback
        # self.start_fly_in_animation()

    def hide(self):
        super().hide()
        # Animation disabled for instant feedback
        # if not self.isVisible():
        #     super().hide()
        #     return
        # self.start_fly_out_animation()

    def start_fly_in_animation(self):
        if hasattr(self, "_anim_group") and self._anim_group.state() == QParallelAnimationGroup.Running:
            self._anim_group.stop()
            
        self._is_animating = False
        # Ensure layout is up to date to get correct end positions
        self.update_layout()
        self._is_animating = True
        
        self._anim_group = QParallelAnimationGroup(self)
        
        def add_anim(widget, start_pos, end_pos):
            if not widget or not widget.isVisible():
                return
            widget.move(start_pos)
            anim = QPropertyAnimation(widget, b"pos")
            anim.setDuration(400)
            anim.setStartValue(start_pos)
            anim.setEndValue(end_pos)
            anim.setEasingCurve(QEasingCurve.OutCubic)
            self._anim_group.addAnimation(anim)

        # 1. Status Bar (Fly from Top)
        if self.status_bar and self.status_bar.isVisible():
            end_pos = self.status_bar.pos()
            start_pos = QPoint(end_pos.x(), -self.status_bar.height())
            add_anim(self.status_bar, start_pos, end_pos)
            
        # 2. Bottom Widgets (Fly from Bottom)
        bottom_widgets = []
        if hasattr(self, "toolbar"): bottom_widgets.append(self.toolbar)
        if hasattr(self, "left_flipper"): bottom_widgets.append(self.left_flipper)
        if hasattr(self, "right_flipper"): bottom_widgets.append(self.right_flipper)
        if hasattr(self, "_dev_watermark"): bottom_widgets.append(self._dev_watermark)
        
        h = self.height()
        for w in bottom_widgets:
            if w and w.isVisible():
                end_pos = w.pos()
                start_pos = QPoint(end_pos.x(), h)
                add_anim(w, start_pos, end_pos)
                
        self._anim_group.finished.connect(lambda: setattr(self, "_is_animating", False))
        self._anim_group.start()

    def start_fly_out_animation(self):
        if hasattr(self, "_anim_group") and self._anim_group.state() == QParallelAnimationGroup.Running:
            self._anim_group.stop()
            
        self._is_animating = True
        self._anim_group = QParallelAnimationGroup(self)
        
        def add_anim(widget, target_y):
            if not widget or not widget.isVisible():
                return
            anim = QPropertyAnimation(widget, b"pos")
            anim.setDuration(300)
            anim.setEndValue(QPoint(widget.x(), target_y))
            anim.setEasingCurve(QEasingCurve.InCubic)
            self._anim_group.addAnimation(anim)

        # 1. Status Bar
        if self.status_bar and self.status_bar.isVisible():
            add_anim(self.status_bar, -self.status_bar.height())
            
        # 2. Bottom Widgets
        bottom_widgets = []
        if hasattr(self, "toolbar"): bottom_widgets.append(self.toolbar)
        if hasattr(self, "left_flipper"): bottom_widgets.append(self.left_flipper)
        if hasattr(self, "right_flipper"): bottom_widgets.append(self.right_flipper)
        if hasattr(self, "_dev_watermark"): bottom_widgets.append(self._dev_watermark)
        
        h = self.height()
        for w in bottom_widgets:
            add_anim(w, h)
            
        def on_finished():
            self._is_animating = False
            super(OverlayWindow, self).hide()
            
        self._anim_group.finished.connect(on_finished)
        self._anim_group.start()

    def cleanup(self):
        if hasattr(self, 'status_bar') and self.status_bar:
            self.status_bar.cleanup()
    
    def _on_status_bar_visibility_changed(self, visible: bool):
        if visible and getattr(self, "_has_status_plugin", False):
            if self.status_bar is None:
                self.status_bar = StatusBarWidget(self)
            self.status_bar.show()
        else:
            if self.status_bar is not None:
                self.status_bar.hide()
        self.update_layout()

    def update_toolbar(self):
        if hasattr(self, 'toolbar') and self.toolbar:
            self.show_reload_mask("正在重载页面")
            def _do_refresh():
                if not shiboken6.isValid(self.toolbar):
                    return
                self.toolbar.refresh_dynamic_tools()
                self.hide_reload_mask()
            self._defer(_do_refresh)
    
    def _init_flippers(self):
        # layout_mode = cfg.toolbarLayout.value
        # Force VerticalSplit layout as per user request
        orientation = "Vertical"
        if hasattr(self, "left_flipper") and self.left_flipper:
            self.left_flipper.deleteLater()
        if hasattr(self, "right_flipper") and self.right_flipper:
            self.right_flipper.deleteLater()
        self.left_flipper = PageFlipWidget("Left", self, height=self.toolbar_height, orientation=orientation)
        self.right_flipper = PageFlipWidget("Right", self, height=self.toolbar_height, orientation=orientation)
        self.left_flipper.clicked_prev.connect(self.request_prev.emit)
        self.left_flipper.clicked_next.connect(self.request_next.emit)
        self.right_flipper.clicked_prev.connect(self.request_prev.emit)
        self.right_flipper.clicked_next.connect(self.request_next.emit)
        
        self.left_flipper.show()
        self.right_flipper.show()

    def update_layout(self):
        # Prevent crash during initialization
        if not hasattr(self, "toolbar") or self.toolbar is None:
            return
        if not hasattr(self, "left_flipper") or not hasattr(self, "right_flipper"):
            return
        if not hasattr(self, "_layout_updating"):
            self._layout_updating = False
        if self._layout_updating or getattr(self, "_is_animating", False):
            return
        self._layout_updating = True
        try:
            w = self.width()
            h = self.height()
            scale = cfg.scale.value
            safe_area = cfg.safeArea.value
            margin = int(16 * scale) + safe_area

            if w <= 100 or h <= 100:
                return

            self.toolbar.adjustSize()
            tb_size = self.toolbar.size()
            tb_w = tb_size.width()
            tb_h = tb_size.height()
            if tb_h < 32 or tb_h > 240:
                tb_h = max(self.toolbar.sizeHint().height(), int(45 * scale))

            flipper_h = int(160 * scale)
            self.left_flipper.setFixedSize(tb_h, flipper_h)
            self.right_flipper.setFixedSize(tb_h, flipper_h)

            self.left_flipper.h_val = tb_h
            self.right_flipper.h_val = tb_h
            self.left_flipper.update_style(getattr(self, "_is_light", False))
            self.right_flipper.update_style(getattr(self, "_is_light", False))

            y_pos = (h - flipper_h) // 2
            self.toolbar.move((w - tb_w) // 2, h - tb_h - int(14 * scale) - safe_area)
            self.left_flipper.move(margin, y_pos)
            self.right_flipper.move(w - self.right_flipper.width() - margin, y_pos)

            self.left_flipper.show()
            self.right_flipper.show()

            if self.status_bar and self.status_bar.isVisible():
                self.status_bar.setFixedWidth(w)
                self.status_bar.move(0, 0)

            if self._dev_watermark:
                self._dev_watermark.move(w - self._dev_watermark.width() - margin, h - self._dev_watermark.height() - int(12 * scale) - safe_area)

            if self._reload_mask is not None and self._reload_mask.isVisible():
                self._reload_mask.setGeometry(self.rect())
                self._reload_mask.raise_()

            self.update_mask()
        finally:
            self._layout_updating = False

    def update_mask(self):
        # Optimization: Mask out empty areas to reduce DWM composition overhead
        # This makes the "transparent" pixels truly pass-through for performance
        if not self.isVisible():
            return

        if self._reload_mask is not None and self._reload_mask.isVisible():
            self.clearMask()
            return
            
        region = QRegion()
        
        # Helper to add widget geometry with margin for shadows
        def add_widget(w, margin=40):
            if w and w.isVisible():
                geo = w.geometry()
                # Expand for shadow (approximate)
                geo.adjust(-margin, -margin, margin, margin)
                # In PySide6/Qt6, unite is deprecated/removed in favor of united or using += operator
                # QRegion.united returns a new region, it does not modify in-place
                nonlocal region
                region = region.united(QRegion(geo))

        add_widget(self.toolbar)
        add_widget(self.left_flipper)
        add_widget(self.right_flipper)
        add_widget(self.status_bar, margin=0) # Status bar has no large shadow
        add_widget(self._dev_watermark, margin=0)
        
        # Add slide widgets
        for widgets in self.slide_widgets.values():
            for w in widgets:
                add_widget(w)
                
        # If no widgets, set empty mask (fully transparent)
        if region.isEmpty():
            # Keep a 1x1 pixel so window exists? Or just empty.
            pass
            
        self.setMask(region)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._layout_updating = False
        if self._reload_mask is not None and self._reload_mask.isVisible():
            self._reload_mask.setGeometry(self.rect())
            self._reload_mask.raise_()
        self.update_mask()
        
    def add_slide_widget(self, widget):
        if not widget:
            return
        
        # Ensure we have the latest page index from monitor if possible
        if self.monitor:
            try:
                curr, _ = self.monitor.get_page_info()
                if curr > 0:
                    # If we were out of sync, just update the index
                    # We assume what's currently on screen matches 'curr'
                    self._current_page_idx = int(curr)
            except Exception:
                pass
        
        # Use current page index
        page_idx = int(self._current_page_idx)
        if page_idx not in self.slide_widgets:
            self.slide_widgets[page_idx] = []
        
        self.slide_widgets[page_idx].append(widget)
        widget.setParent(self)
        
        # Connect close signal if available
        if hasattr(widget, "request_close"):
            widget.request_close.connect(lambda: self._remove_slide_widget(widget))
            
        widget.show()

    def _remove_slide_widget(self, widget):
        # Find which page it belongs to
        for page_idx, widgets in self.slide_widgets.items():
            if widget in widgets:
                widgets.remove(widget)
                widget.hide()
                widget.deleteLater()
                break

    def update_page_info(self, current, total):
        # Ensure int
        try:
            current = int(current)
        except ValueError:
            return

        # Handle slide-bound widgets visibility
        if self._current_page_idx != current:
            # Hide old widgets
            if self._current_page_idx in self.slide_widgets:
                for w in self.slide_widgets[self._current_page_idx]:
                    w.hide()
            
            self._current_page_idx = current
            
            # Show new widgets
            if current in self.slide_widgets:
                for w in self.slide_widgets[current]:
                    w.show()

        if hasattr(self, 'left_flipper'):
            self.left_flipper.set_page_info(current, total)
        if hasattr(self, 'right_flipper'):
            self.right_flipper.set_page_info(current, total)
            
        if self.isVisible():
            QTimer.singleShot(0, self.update)

        # Re-trigger style update to ensure radius is correct
        tb_h = 56
        if hasattr(self, 'toolbar'):
            tb_h = self.toolbar.height()
            
        if hasattr(self, 'left_flipper'):
            self.left_flipper.h_val = tb_h
            self.left_flipper.update_style(getattr(self, "_is_light", False))
        if hasattr(self, 'right_flipper'):
            self.right_flipper.h_val = tb_h
            self.right_flipper.update_style(getattr(self, "_is_light", False))

class ToolbarWidget(QWidget):
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
        self._style_update_pending = False
        self._style_update_timer = QTimer(self)
        self._style_update_timer.setSingleShot(True)
        self._style_update_timer.timeout.connect(self._apply_layout_style)

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
        if self._style_update_timer.isActive():
            self._style_update_timer.stop()
        self._style_update_pending = True
        p = self.parent()
        if shiboken6.isValid(p) and hasattr(p, "show_reload_mask"):
            p.show_reload_mask("正在重载页面")
        self._style_update_timer.start(0)

    def _apply_layout_style(self):
        try:
            if not shiboken6.isValid(self) or not hasattr(self, "layout") or self.layout is None:
                return
            show_text = cfg.showToolbarText.value

            if self._is_light:
                bg = "#FFFFFF"
                border = "rgba(0, 0, 0, 0.08)"
                line_color = "rgba(0, 0, 0, 0.08)"
                shadow_color = QColor(0, 0, 0, 15)
            else:
                bg = "#202020"
                border = "rgba(255, 255, 255, 0.08)"
                line_color = "rgba(255, 255, 255, 0.15)"
                shadow_color = QColor(0, 0, 0, 80)

            self.layout.activate()
            self.adjustSize()

            for line in self.findChildren(QFrame):
                if line.frameShape() == QFrame.VLine:
                    line.setFixedWidth(1)
                    line.setStyleSheet(f"background-color: {line_color}; border: none; margin: 10px 0;")

            for btn in self.findChildren(CustomToolButton):
                btn.update_size()
                btn.update_style(btn.tool_name == self.current_tool, self._is_light)

            self.layout.activate()
            self.adjustSize()
            hint = self.sizeHint()

            radius = hint.height() // 2

            self.setStyleSheet(f"""
                ToolbarWidget {{
                    background-color: {bg};
                    border: 1px solid {border};
                    border-radius: {radius}px;
                }}
            """)

            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(40)
            shadow.setColor(shadow_color)
            shadow.setOffset(0, 8)
            self.setGraphicsEffect(shadow)

            p = self.parent()
            if shiboken6.isValid(p) and hasattr(p, "update_layout"):
                p.update_layout()
        finally:
            self._style_update_pending = False
            p = self.parent()
            if shiboken6.isValid(p) and hasattr(p, "hide_reload_mask"):
                p.hide_reload_mask()

    def update_style(self, is_light=False):
        # Compatibility with old calls
        self._is_light = is_light
        self.update_layout_style()

    def _on_toolbar_visibility_changed(self, value):
        self.update_toolbar_layout()
        QTimer.singleShot(10, self._update_indicator_now)

    def _execute_plugin_by_name(self, name):
        for plugin in self.plugins:
            if hasattr(plugin, "get_name") and plugin.get_name() == name:
                plugin.execute()
                return

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

        # Initialize all buttons
        self.btn_select = CustomToolButton("Mouse.svg", _t("toolbar.select"), self, tool_name="select", text=_t("toolbar.select"))
        self.btn_select.clicked.connect(lambda: self._on_tool_changed("select", self.select_clicked))

        self.btn_pen = CustomToolButton("Pen.svg", _t("toolbar.pen"), self, tool_name="pen", text=_t("toolbar.pen"))
        self.btn_pen.clicked.connect(self._on_pen_button_clicked)

        self.btn_eraser = CustomToolButton("Eraser.svg", _t("toolbar.eraser"), self, tool_name="eraser", text=_t("toolbar.eraser"))
        self.btn_eraser.clicked.connect(lambda: self._on_tool_changed("eraser", self.eraser_clicked))

        self.btn_undo = CustomToolButton("Previous.svg", _t("toolbar.undo"), self, text=_t("toolbar.undo"))
        self.btn_undo.clicked.connect(self.prev_clicked.emit)

        self.btn_redo = CustomToolButton("Next.svg", _t("toolbar.redo"), self, text=_t("toolbar.redo"))
        self.btn_redo.clicked.connect(self.next_clicked.emit)

        self.btn_spotlight = CustomToolButton("spotlight.svg", _t("toolbar.spotlight"), self, text=_t("toolbar.spotlight"))
        self.btn_spotlight.clicked.connect(lambda: self._execute_plugin_by_name("聚光灯"))

        self.btn_timer = CustomToolButton("timer.svg", _t("toolbar.timer"), self, text=_t("toolbar.timer"))
        self.btn_timer.clicked.connect(lambda: self._execute_plugin_by_name("计时器"))

        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.VLine)
        self.line1.setFixedHeight(24)

        self.dynamic_container = QWidget(self)
        self.dynamic_layout = QHBoxLayout(self.dynamic_container)
        self.dynamic_layout.setContentsMargins(0, 0, 0, 0)
        self.dynamic_layout.setSpacing(4)

        self.line3 = QFrame()
        self.line3.setFrameShape(QFrame.VLine)
        self.line3.setFixedHeight(24)

        self.btn_end = CustomToolButton("Minimize.svg", _t("toolbar.end_show"), self, is_exit=True, text=_t("toolbar.end_show"))
        self.btn_end.clicked.connect(self.end_clicked.emit)

        # Add to layout based on order
        self.update_toolbar_layout()

        # Connect signals
        cfg.showUndoRedo.valueChanged.connect(self._on_toolbar_visibility_changed)
        cfg.showSpotlight.valueChanged.connect(self._on_toolbar_visibility_changed)
        cfg.showTimer.valueChanged.connect(self._on_toolbar_visibility_changed)
        cfg.toolbarOrder.valueChanged.connect(self.update_toolbar_layout)

        QTimer.singleShot(0, self._update_indicator_now)

    def update_toolbar_layout(self):
        # Clear layout
        while self.layout.count():
            item = self.layout.takeAt(0)
            if item.widget():
                item.widget().setParent(None)
        
        order = cfg.toolbarOrder.value
        has_undo_redo = cfg.showUndoRedo.value
        has_spotlight = cfg.showSpotlight.value
        has_timer = cfg.showTimer.value

        for item_id in order:
            if item_id == "select":
                self.layout.addWidget(self.btn_select)
            elif item_id == "pen":
                self.layout.addWidget(self.btn_pen)
            elif item_id == "eraser":
                self.layout.addWidget(self.btn_eraser)
            elif item_id == "undo":
                self.btn_undo.setVisible(has_undo_redo)
                if has_undo_redo:
                    self.layout.addWidget(self.btn_undo)
            elif item_id == "redo":
                self.btn_redo.setVisible(has_undo_redo)
                if has_undo_redo:
                    self.layout.addWidget(self.btn_redo)
            elif item_id == "spotlight":
                self.btn_spotlight.setVisible(has_spotlight)
                if has_spotlight:
                    self.layout.addWidget(self.btn_spotlight)
            elif item_id == "timer":
                self.btn_timer.setVisible(has_timer)
                if has_timer:
                    self.layout.addWidget(self.btn_timer)
            elif item_id == "apps":
                self.layout.addWidget(self.dynamic_container)
            
        # Add separator line and end button
        self.layout.addWidget(self.line3)
        self.layout.addWidget(self.btn_end)
        
        self.refresh_dynamic_tools()
        self.update_layout_style()

    def refresh_dynamic_tools(self):
        p = self.parent()
        if shiboken6.isValid(p) and hasattr(p, "show_reload_mask"):
            p.show_reload_mask("正在重载页面")
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

            if name in ["聚光灯", "计时器"]:
                continue

            if isinstance(plugin_type, str) and plugin_type.startswith("toolbar"):
                toolbar_plugins.append(plugin)

        if toolbar_plugins:
            line = QFrame()
            line.setFrameShape(QFrame.VLine)
            line.setFixedHeight(24)
            self.dynamic_layout.addWidget(line)
            for plugin in toolbar_plugins:
                pixmap = None
                if hasattr(plugin, 'get_pixmap'):
                    pixmap = plugin.get_pixmap()
                
                btn = CustomToolButton(plugin.get_icon() or "More.svg", plugin.get_name(), self, text=plugin.get_name(), pixmap=pixmap)
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
        p = self.parent()
        if shiboken6.isValid(p) and hasattr(p, "update_layout"):
            p.update_layout()
        if shiboken6.isValid(p) and hasattr(p, "_defer") and hasattr(p, "hide_reload_mask"):
            p._defer(p.hide_reload_mask)
        elif shiboken6.isValid(p) and hasattr(p, "hide_reload_mask"):
            p.hide_reload_mask()

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._update_indicator_now)

class PageFlipButton(QFrame):
    btn_clicked = Signal()

    def __init__(self, icon_name, parent=None, rotation=0):
        super().__init__(parent)
        self.icon_name = icon_name
        self.rotation = rotation
        scale = cfg.scale.value
        self.setFixedSize(int(38 * scale), int(38 * scale))
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"""
            QFrame {{
                background-color: transparent;
                border-radius: {int(19 * scale)}px;
            }}
            QFrame:hover {{
                background-color: rgba(255, 255, 255, 0.08);
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.setAlignment(Qt.AlignCenter)
        
        self.icon_label = QLabel(self)
        self.icon_label.setFixedSize(int(20 * scale), int(20 * scale))
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        layout.addWidget(self.icon_label)
        self.update_icon_color(QColor(255, 255, 255))

    def update_icon_color(self, color):
        icon_path = os.path.join(ICON_DIR, self.icon_name)
        if not os.path.exists(icon_path):
            return
            
        scale = cfg.scale.value
        s = int(20 * scale)
        cache_key = (self.icon_name, color.name(), s, self.rotation)
        cached_pixmap = GlobalIconCache.get(cache_key)
        if cached_pixmap:
            self.icon_label.setPixmap(cached_pixmap)
            return

        renderer = QSvgRenderer(icon_path)
        if renderer.isValid():
            device_size = int(40 * scale)
            base = QPixmap(device_size, device_size)
            base.fill(Qt.transparent)
            p = QPainter(base)
            p.setRenderHint(QPainter.Antialiasing)
            p.translate(base.width() / 2, base.height() / 2)
            if self.rotation:
                p.rotate(self.rotation)
            p.translate(-base.width() / 2, -base.height() / 2)
            renderer.render(p)
            p.setCompositionMode(QPainter.CompositionMode_SourceIn)
            p.fillRect(base.rect(), color)
            p.end()
            scaled = base.scaled(s, s, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            GlobalIconCache.set(cache_key, scaled)
            self.icon_label.setPixmap(scaled)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.btn_clicked.emit()
        super().mousePressEvent(event)

class PageFlipWidget(QFrame):
    clicked_prev = Signal()
    clicked_next = Signal()
    page_clicked = Signal(QPoint)

    def __init__(self, side="Left", parent=None, height=56, orientation="Horizontal"):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.side = side
        scale = cfg.scale.value
        self.h_val = int(height * scale)
        self.orientation = orientation
        
        self.update_style()
        
        # Center container that takes all remaining space
        self.page_container = QFrame(self)
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setContentsMargins(0, 0, 0, 0)
        self.page_layout.setSpacing(0)
        self.page_layout.setAlignment(Qt.AlignCenter)
        
        self.lbl_page = ClickableLabel("0/0", self.page_container)
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.lbl_page.clicked.connect(self.page_clicked.emit)
        # Ensure label repaints cleanly when content changes
        self.lbl_page.setAttribute(Qt.WA_OpaquePaintEvent, False)
        
        self.lbl_hint = QLabel(_t("toolbar.page"), self.page_container)
        self.lbl_hint.setAlignment(Qt.AlignCenter)
        self.lbl_hint.setObjectName("PageHint")
        self.lbl_hint.setVisible(True)
        
        self.page_layout.addWidget(self.lbl_page)
        self.page_layout.addWidget(self.lbl_hint)

        if orientation == "Vertical":
             self.setFixedSize(self.h_val, int(160 * scale))
             self.layout = QVBoxLayout(self)
             self.layout.setContentsMargins(0, int(5 * scale), 0, int(5 * scale))
             
             # Up arrow (Previous) - Rotate 90 deg
             self.btn_prev = PageFlipButton("Previous.svg", self, rotation=90)
             
             # Down arrow (Next) - Rotate 90 deg
             self.btn_next = PageFlipButton("Next.svg", self, rotation=90)
             
             self.layout.addWidget(self.btn_prev, 0, Qt.AlignHCenter)
             self.layout.addWidget(self.page_container, 1)
             self.layout.addWidget(self.btn_next, 0, Qt.AlignHCenter)
        else:
             self.setFixedSize(int(160 * scale), self.h_val)
             self.layout = QHBoxLayout(self)
             self.layout.setContentsMargins(int(5 * scale), 0, int(5 * scale), 0)
             self.btn_prev = PageFlipButton("Previous.svg", self)
             self.btn_next = PageFlipButton("Next.svg", self)
             
             self.layout.addWidget(self.btn_prev, 0, Qt.AlignVCenter)
             self.layout.addWidget(self.page_container, 1)
             self.layout.addWidget(self.btn_next, 0, Qt.AlignVCenter)
             
        self.layout.setSpacing(0)
        
        self.btn_prev.btn_clicked.connect(self.clicked_prev.emit)
        self.btn_next.btn_clicked.connect(self.clicked_next.emit)

    def _on_show_text_changed(self, show):
        # Hint is now always visible, no need to toggle
        pass

    def set_page_info(self, current, total):
        hint_fg = "rgba(255, 255, 255, 0.6)" if not hasattr(self, "_is_light") or not self._is_light else "rgba(0, 0, 0, 0.5)"
        scale = cfg.scale.value
        
        # Calculate optimal font size based on number length to prevent overflow
        num_len = len(str(current))
        font_size = int(16 * scale)
        if num_len > 3:
            font_size = int(12 * scale)
        elif num_len > 2:
            font_size = int(14 * scale)
            
        self.lbl_page.setText(f'<span style="font-size: {font_size}px; font-weight: 900;">{current}</span>'
                              f'<span style="font-size: {int(10 * scale)}px; font-weight: 400; color: {hint_fg};">/{total}</span>')
        self.lbl_page.repaint()

    def update_style(self, is_light=False):
        self._is_light = is_light
        scale = cfg.scale.value
        
        # Settings Design Style
        if self._is_light:
            bg = "#FFFFFF"
            border = "rgba(0, 0, 0, 0.08)"
            fg = "#191919"
            hint_fg = "rgba(0, 0, 0, 0.5)"
            hover_bg = "rgba(0, 0, 0, 0.05)"
            shadow_color = QColor(0, 0, 0, 15)
        else:
            bg = "#202020"
            border = "rgba(255, 255, 255, 0.08)"
            fg = "white"
            hint_fg = "rgba(255, 255, 255, 0.6)"
            hover_bg = "rgba(255, 255, 255, 0.08)"
            shadow_color = QColor(0, 0, 0, 80)
        
        # Calculate radius to ensure it's always a capsule (pill shape)
        # Use min dimension for radius
        radius = min(self.width(), self.height()) // 2
        
        self.setStyleSheet(f"""
            PageFlipWidget {{
                background-color: {bg};
                border-radius: {radius}px;
                border: 1px solid {border};
            }}
            QLabel {{
                color: {fg};
                font-family: 'MiSans Latin', 'HarmonyOS Sans SC', 'SF Pro', '苹方-简', 'PingFang SC', 'Segoe UI', 'Microsoft YaHei', sans-serif;
                font-size: {int(12 * scale)}px;
                font-weight: 900;
                background: transparent;
                border: none;
            }}
            QLabel#PageHint {{
                font-size: {int(9 * scale)}px;
                font-weight: 400;
                color: {hint_fg};
            }}
        """)

        # Update buttons hover style
        for btn in self.findChildren(PageFlipButton):
            btn.setStyleSheet(f"""
                QFrame {{
                    background-color: transparent;
                    border-radius: {int(19 * scale)}px;
                }}
                QFrame:hover {{
                    background-color: {hover_bg};
                }}
            """)
            btn.update_icon_color(QColor(fg))

        # Shadow
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setColor(shadow_color)
        shadow.setOffset(0, 8)
        self.setGraphicsEffect(shadow)
