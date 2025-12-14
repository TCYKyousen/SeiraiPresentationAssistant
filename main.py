import sys
import os
import winreg
import win32com.client
import win32gui
import hashlib
from PyQt6.QtWidgets import (QApplication, QWidget, QHBoxLayout, QVBoxLayout, QButtonGroup, 
                             QSystemTrayIcon, QMenu, QLabel, QFrame, QScrollArea, QGridLayout,
                             QColorDialog, QFileDialog, QPushButton)
from PyQt6.QtCore import Qt, QTimer, QSize, QPoint, QPointF, pyqtSignal, QEvent, QRect, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import (QIcon, QPainter, QColor, QAction, QPixmap, QPen, QBrush, 
                         QPolygon, QPolygonF, QMouseEvent)
from qfluentwidgets import (PushButton, TransparentToolButton, ToolButton, 
                            setTheme, Theme, isDarkTheme, ToolTipFilter, ToolTipPosition,
                            Flyout, FlyoutView, FlyoutAnimationType)
from qfluentwidgets.components.material import AcrylicFlyout


def icon_path(name):
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, "icons", name)

class IconFactory:
    @staticmethod
    def draw_cursor(color):
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Draw a cursor arrow
        path = QPolygonF([
            QPointF(10, 6),
            QPointF(10, 26),
            QPointF(15, 21),
            QPointF(22, 28),
            QPointF(24, 26),
            QPointF(17, 19),
            QPointF(24, 19)
        ])
        
        painter.setPen(QPen(QColor("white"), 1.5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        painter.setBrush(QBrush(QColor(color)))
        painter.drawPolygon(path)
        
        painter.end()
        return QIcon(pixmap)

    @staticmethod
    def draw_arrow(color, direction='left'):
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        pen = QPen(QColor(color))
        pen.setWidth(3)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        
        if direction == 'left':
            points = [QPoint(20, 8), QPoint(12, 16), QPoint(20, 24)]
            painter.drawPolyline(points)
        elif direction == 'right':
            points = [QPoint(12, 8), QPoint(20, 16), QPoint(12, 24)]
            painter.drawPolyline(points)
            
        painter.end()
        return QIcon(pixmap)

class SlidePreviewCard(QWidget):
    clicked = pyqtSignal(int)
    
    def __init__(self, index, image_path, parent=None):
        super().__init__(parent)
        self.index = index
        self.setFixedSize(200, 140) 
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        self.img_label = QLabel()
        self.img_label.setFixedSize(190, 107) # 16:9 approx
        self.img_label.setStyleSheet("background-color: #333333; border-radius: 6px; border: 1px solid #444444;")
        self.img_label.setScaledContents(True)
        if image_path and os.path.exists(image_path):
            self.img_label.setPixmap(QPixmap(image_path))
            
        self.txt_label = QLabel(f"{index}")
        self.txt_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.txt_label.setStyleSheet("font-size: 14px; color: #dddddd; font-weight: bold;")
        
        layout.addWidget(self.txt_label)
        layout.addWidget(self.img_label)
        
    def mousePressEvent(self, event):
        self.clicked.emit(self.index)

class SlideSelectorFlyout(QWidget):
    slide_selected = pyqtSignal(int)
    
    def __init__(self, ppt_app, parent=None):
        super().__init__(parent)
        self.ppt_app = ppt_app
        self.setStyleSheet("background-color: rgba(30, 30, 30, 240); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        
        title = QLabel("幻灯片预览")
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px; color: white; border: none; background: transparent;")
        layout.addWidget(title)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedSize(450, 500) 
        scroll.setStyleSheet("""
            QScrollArea { border: none; background-color: rgba(30, 30, 30, 240); }
            QScrollBar:vertical {
                border: none;
                background: transparent;
                width: 8px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 255, 255, 0.2);
                min-height: 20px;
                border-radius: 0px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)
        
        container = QWidget()
        container.setStyleSheet("background-color: rgba(30, 30, 30, 240);")
        self.grid = QGridLayout(container)
        self.grid.setSpacing(15)
        
        self.load_slides()
        
        scroll.setWidget(container)
        layout.addWidget(scroll)
        
    def get_cache_dir(self, presentation_path):
        path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
        cache_dir = os.path.join(os.environ['APPDATA'], 'PPTAssistant', 'Cache', path_hash)
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
        return cache_dir

    def load_slides(self):
        try:
            presentation = self.ppt_app.ActivePresentation
            slides_count = presentation.Slides.Count
            presentation_path = presentation.FullName
            cache_dir = self.get_cache_dir(presentation_path)
            
            for i in range(1, slides_count + 1):
                slide = presentation.Slides(i)
                thumb_path = os.path.join(cache_dir, f"slide_{i}.jpg")
                
                if not os.path.exists(thumb_path):
                    try:
                        slide.Export(thumb_path, "JPG", 640, 360) 
                    except:
                        pass
                    
                card = SlidePreviewCard(i, thumb_path)
                card.clicked.connect(self.on_card_clicked)
                row = (i - 1) // 2
                col = (i - 1) % 2
                self.grid.addWidget(card, row, col)
                
        except Exception as e:
            print(f"Error loading slides: {e}")
            
    def on_card_clicked(self, index):
        self.slide_selected.emit(index)
        parent = self.parent()
        while parent:
            if isinstance(parent, Flyout):
                parent.close()
                break
            parent = parent.parent()

class PenSettingsFlyout(QWidget):
    color_selected = pyqtSignal(int) # Returns RGB integer for PPT
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: rgba(30, 30, 30, 240); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        
        title = QLabel("笔颜色")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px; color: white; border: none; background: transparent;")
        layout.addWidget(title)
        
        # Grid of colors
        grid_widget = QWidget()
        grid = QGridLayout(grid_widget)
        grid.setSpacing(10)
        
        colors = [
            (255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0),
            (255, 0, 255), (0, 255, 255), (0, 0, 0), (255, 255, 255),
            (255, 165, 0), (128, 0, 128)
        ]
        
        for i, rgb in enumerate(colors):
            btn = QPushButton()
            btn.setFixedSize(40, 40)
            color_hex = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color_hex};
                    border: 2px solid #555;
                    border-radius: 0px;
                }}
                QPushButton:hover {{
                    border: 2px solid white;
                }}
            """)
            # PPT uses RGB integer: R + (G << 8) + (B << 16)
            ppt_color = rgb[0] + (rgb[1] << 8) + (rgb[2] << 16)
            btn.clicked.connect(lambda checked, c=ppt_color: self.on_color_clicked(c))
            row = i // 5
            col = i % 5
            grid.addWidget(btn, row, col)
            
        layout.addWidget(grid_widget)
        
    def on_color_clicked(self, color):
        self.color_selected.emit(color)
        # Close parent flyout
        parent = self.parent()
        while parent:
            if isinstance(parent, Flyout):
                parent.close()
                break
            parent = parent.parent()

class EraserSettingsFlyout(QWidget):
    clear_all_clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: rgba(30, 30, 30, 240); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        
        btn = PushButton("清除当前页笔迹")
        btn.setFixedSize(200, 40)
        btn.clicked.connect(self.on_clicked)
        layout.addWidget(btn)
        
    def on_clicked(self):
        self.clear_all_clicked.emit()
        parent = self.parent()
        while parent:
            if isinstance(parent, Flyout):
                parent.close()
                break
            parent = parent.parent()


class SpotlightOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        self.selection_rect = QRect()
        self.is_selecting = False
        self.has_selection = False
        
        # Close button (context aware)
        self.btn_close = QPushButton("X", self)
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.setStyleSheet("background-color: red; color: white; border-radius: 0px; font-weight: bold;")
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.RightButton:
            self.close()
        elif event.button() == Qt.MouseButton.LeftButton:
            self.selection_rect.setTopLeft(event.pos())
            self.selection_rect.setBottomRight(event.pos())
            self.is_selecting = True
            self.has_selection = False
            self.btn_close.hide()
            self.update()
            
    def mouseMoveEvent(self, event: QMouseEvent):
        if self.is_selecting:
            self.selection_rect.setBottomRight(event.pos())
            self.update()
            
    def mouseReleaseEvent(self, event: QMouseEvent):
        if self.is_selecting:
            self.is_selecting = False
            self.has_selection = True
            # Show close button near top-right of selection
            normalized_rect = self.selection_rect.normalized()
            self.btn_close.move(normalized_rect.topRight() + QPoint(10, -15))
            self.btn_close.show()
            self.update()
            
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fill screen with semi-transparent black
        painter.setBrush(QColor(0, 0, 0, 180))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRect(self.rect())
        
        if self.has_selection or self.is_selecting:
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
            painter.setBrush(Qt.GlobalColor.transparent)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(self.selection_rect, 10, 10)
            
            # Draw border
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            pen = QPen(QColor("#00cc7a"))
            pen.setWidth(3)
            painter.setPen(pen)
            painter.drawRoundedRect(self.selection_rect, 10, 10)

class PageNavWidget(QWidget):
    request_slide_jump = pyqtSignal(int)
    
    def __init__(self, parent=None, is_right=False):
        super().__init__(parent)
        self.ppt_app = None 
        self.is_right = is_right
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.layout = QHBoxLayout()
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        self.container = QWidget()
        # Dark Theme Style
        
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: rgba(30, 30, 30, 240);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-bottom: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 12px;
            }}
            QLabel {{
                font-family: "Segoe UI", "Microsoft YaHei";
                color: white;
            }}
        """)
        self.container.setObjectName("Container")
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(8, 6, 8, 6) 
        inner_layout.setSpacing(15) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(52)
        self.btn_prev = TransparentToolButton(parent=self)
        self.btn_prev.setIcon(QIcon(icon_path("Previous.svg")))
        self.btn_prev.setFixedSize(36, 36) 
        self.btn_prev.setIconSize(QSize(18, 18))
        self.btn_prev.setToolTip("上一页")
        self.btn_prev.installEventFilter(ToolTipFilter(self.btn_prev, 1000, ToolTipPosition.TOP))
        self.style_nav_btn(self.btn_prev)
        
        self.btn_next = TransparentToolButton(parent=self)
        self.btn_next.setIcon(QIcon(icon_path("Next.svg")))
        self.btn_next.setFixedSize(36, 36) 
        self.btn_next.setIconSize(QSize(18, 18))
        self.btn_next.setToolTip("下一页")
        self.btn_next.installEventFilter(ToolTipFilter(self.btn_next, 1000, ToolTipPosition.TOP))
        self.style_nav_btn(self.btn_next)

        # Page Info Clickable Area
        self.page_info_widget = QWidget()
        self.page_info_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.page_info_widget.installEventFilter(self)
        
        info_layout = QVBoxLayout(self.page_info_widget)
        info_layout.setContentsMargins(10, 0, 10, 0)
        info_layout.setSpacing(2)
        info_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_page_num = QLabel("1/--")
        self.lbl_page_num.setStyleSheet("font-size: 18px; font-weight: bold; color: white;")
        self.lbl_page_text = QLabel("页码")
        self.lbl_page_text.setStyleSheet("font-size: 12px; color: #aaaaaa;")
        
        info_layout.addWidget(self.lbl_page_num, 0, Qt.AlignmentFlag.AlignCenter)
        info_layout.addWidget(self.lbl_page_text, 0, Qt.AlignmentFlag.AlignCenter)
        
        inner_layout.addWidget(self.btn_prev)
        
        line1 = QFrame()
        line1.setFrameShape(QFrame.Shape.VLine)
        line1.setStyleSheet("color: rgba(255, 255, 255, 0.2);")
        inner_layout.addWidget(line1)
        
        inner_layout.addWidget(self.page_info_widget)
        
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.VLine)
        line2.setStyleSheet("color: rgba(255, 255, 255, 0.2);")
        inner_layout.addWidget(line2)
        
        inner_layout.addWidget(self.btn_next)
        
        self.layout.addWidget(self.container)
        self.setLayout(self.layout)

    def style_nav_btn(self, btn):
        btn.setStyleSheet("""
            TransparentToolButton {
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: white;
            }
            TransparentToolButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
            }
            TransparentToolButton:pressed {
                background-color: rgba(255, 255, 255, 0.2);
            }
        """)
    
    def eventFilter(self, obj, event):
        if obj == self.page_info_widget:
            if event.type() == QEvent.Type.MouseButtonDblClick:
                return super().eventFilter(obj, event)
            elif event.type() == QEvent.Type.MouseButtonRelease:
                self.show_slide_selector()
                return True
        return super().eventFilter(obj, event)

    def show_slide_selector(self):
        if not self.ppt_app:
            return
            
        view = SlideSelectorFlyout(self.ppt_app)
        view.slide_selected.connect(self.request_slide_jump.emit)
        
        flyout = AcrylicFlyout(view, self.window())
        flyout.exec(self.page_info_widget.mapToGlobal(self.page_info_widget.rect().bottomLeft()), FlyoutAnimationType.PULL_UP)

    def update_page(self, current, total):
        self.lbl_page_num.setText(f"{current}/{total}")
    
    def apply_settings(self):
        self.btn_prev.setToolTip("上一页")
        self.btn_next.setToolTip("下一页")
        self.lbl_page_text.setText("页码")
        self.style_nav_btn(self.btn_prev)
        self.style_nav_btn(self.btn_next)

class ToolBarWidget(QWidget):
    request_spotlight = pyqtSignal()
    request_pen_color = pyqtSignal(int)
    request_clear_ink = pyqtSignal()
    request_exit = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.was_checked = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        # Dark Theme Style
        self.container.setStyleSheet("""
            QWidget#Container {
                background-color: rgba(30, 30, 30, 240);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-bottom: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 12px;
            }
        """)
        
        container_layout = QHBoxLayout(self.container)
        container_layout.setContentsMargins(8, 6, 8, 6) 
        container_layout.setSpacing(12) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(56)
        
        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        
        self.btn_arrow = self.create_tool_btn("选择", QIcon(icon_path("Mouse.svg")))
        self.btn_pen = self.create_tool_btn("笔", QIcon(icon_path("Pen.svg")))
        self.btn_eraser = self.create_tool_btn("橡皮", QIcon(icon_path("Eraser.svg")))
        self.btn_clear = self.create_action_btn("一键清除", QIcon(icon_path("Clear.svg")))
        self.btn_clear.clicked.connect(self.request_clear_ink.emit)
        
        self.group.addButton(self.btn_arrow)
        self.group.addButton(self.btn_pen)
        self.group.addButton(self.btn_eraser)
        
        self.btn_spotlight = self.create_action_btn("聚焦", QIcon(icon_path("Select.svg")))
        self.btn_spotlight.clicked.connect(self.request_spotlight.emit)
        
        self.btn_exit = self.create_action_btn("结束放映", QIcon(icon_path("Minimaze.svg")))
        self.btn_exit.clicked.connect(self.request_exit.emit)
        self.btn_exit.setStyleSheet("""
            TransparentToolButton {
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: white;
            }
            TransparentToolButton:hover {
                background-color: rgba(255, 50, 50, 0.3);
            }
            TransparentToolButton:pressed {
                background-color: rgba(255, 50, 50, 0.5);
            }
        """)
        
        container_layout.addWidget(self.btn_arrow)
        container_layout.addWidget(self.btn_pen)
        container_layout.addWidget(self.btn_eraser)
        container_layout.addWidget(self.btn_clear)
        
        line1 = QFrame()
        line1.setFrameShape(QFrame.Shape.VLine)
        line1.setStyleSheet("color: rgba(255, 255, 255, 0.2);")
        container_layout.addWidget(line1)
        
        container_layout.addWidget(self.btn_spotlight)
        
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.VLine)
        line2.setStyleSheet("color: rgba(255, 255, 255, 0.2);")
        container_layout.addWidget(line2)
        
        container_layout.addWidget(self.btn_exit)
        
        layout.addWidget(self.container)
        self.setLayout(layout)
        
        self.btn_arrow.setChecked(True)
        
        # Install event filter to detect second click for expansion
        self.btn_pen.installEventFilter(self)
        self.btn_eraser.installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            if obj in [self.btn_pen, self.btn_eraser]:
                self.was_checked = obj.isChecked()
        elif event.type() == QEvent.Type.MouseButtonRelease:
            if obj == self.btn_pen and self.was_checked and self.btn_pen.isChecked():
                self.show_pen_settings()
            elif obj == self.btn_eraser and self.was_checked and self.btn_eraser.isChecked():
                self.show_eraser_settings()
            self.was_checked = False
        return super().eventFilter(obj, event)

    def show_pen_settings(self):
        view = PenSettingsFlyout()
        view.color_selected.connect(self.request_pen_color.emit)
        flyout = AcrylicFlyout(view, self.window())
        flyout.exec(self.btn_pen.mapToGlobal(self.btn_pen.rect().bottomLeft()), FlyoutAnimationType.PULL_UP)

    def show_eraser_settings(self):
        view = EraserSettingsFlyout()
        view.clear_all_clicked.connect(self.request_clear_ink.emit)
        flyout = AcrylicFlyout(view, self.window())
        flyout.exec(self.btn_eraser.mapToGlobal(self.btn_eraser.rect().bottomLeft()), FlyoutAnimationType.PULL_UP)
    
    def apply_settings(self):
        self.btn_arrow.setToolTip("选择")
        self.btn_pen.setToolTip("笔")
        self.btn_eraser.setToolTip("橡皮")
        self.btn_clear.setToolTip("一键清除")
        self.btn_spotlight.setToolTip("聚焦")
        self.btn_exit.setToolTip("结束放映")
        self.style_tool_btn(self.btn_arrow)
        self.style_tool_btn(self.btn_pen)
        self.style_tool_btn(self.btn_eraser)
        self.style_action_btn(self.btn_clear)
        self.style_action_btn(self.btn_spotlight)
        self.style_action_btn(self.btn_exit)

    def create_tool_btn(self, text, icon):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(icon)
        btn.setFixedSize(40, 40) 
        btn.setIconSize(QSize(20, 20))
        btn.setCheckable(True)
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        self.style_tool_btn(btn)
        return btn
        
    def create_action_btn(self, text, icon):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(icon)
        btn.setFixedSize(40, 40)
        btn.setIconSize(QSize(20, 20))
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        self.style_action_btn(btn)
        return btn
    
    def style_tool_btn(self, btn):
        btn.setStyleSheet("""
            TransparentToolButton {
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: white;
                margin-bottom: 2px;
            }
            TransparentToolButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
            }
            TransparentToolButton:checked {
                background-color: rgba(255, 255, 255, 0.1);
                color: white;
                border-bottom: 3px solid #00cc7a;
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }
            TransparentToolButton:checked:hover {
                background-color: rgba(255, 255, 255, 0.15);
            }
        """)
    
    def style_action_btn(self, btn):
        btn.setStyleSheet("""
            TransparentToolButton {
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: white;
            }
            TransparentToolButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
            }
            TransparentToolButton:pressed {
                background-color: rgba(255, 255, 255, 0.2);
            }
        """)

class MainController(QWidget):
    def __init__(self):
        super().__init__()
        setTheme(Theme.DARK) 
        
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(1, 1)
        self.move(-100, -100) 
        
        self.ppt_app = None
        self.current_view = None
        
        self.toolbar = ToolBarWidget()
        self.nav_left = PageNavWidget(is_right=False)
        self.nav_right = PageNavWidget(is_right=True)
        self.spotlight = SpotlightOverlay()
        
        self.setup_connections()
        self.setup_tray()
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_state)
        self.timer.start(500)
        
        self.widgets_visible = False

    def setup_connections(self):
        self.toolbar.btn_arrow.clicked.connect(lambda: self.set_pointer(1))
        self.toolbar.btn_pen.clicked.connect(lambda: self.set_pointer(2))
        self.toolbar.btn_eraser.clicked.connect(lambda: self.set_pointer(5))
        self.toolbar.request_exit.connect(self.exit_slideshow)
        
        self.toolbar.request_spotlight.connect(self.toggle_spotlight)
        self.toolbar.request_pen_color.connect(self.set_pen_color)
        self.toolbar.request_clear_ink.connect(self.clear_ink)
        
        for nav in [self.nav_left, self.nav_right]:
            nav.btn_prev.clicked.connect(self.go_prev)
            nav.btn_next.clicked.connect(self.go_next)
            nav.request_slide_jump.connect(self.jump_to_slide)

    def has_ink(self):
        try:
            view = self.get_ppt_view()
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
            # If any error occurs during check (e.g. COM busy), 
            # fail safe to True to allow eraser usage (don't block user)
            return True

    def show_warning(self, target, message):
        title = "PPT助手提示"
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.MessageIcon.Warning, 2000)

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        # Use a more visible icon (using our arrow drawer for consistency)
        self.tray_icon.setIcon(IconFactory.draw_arrow("#00cc7a", 'right'))
        self.tray_icon.setToolTip("PPT演示助手")
        
        menu = QMenu()
        
        # Autorun Action
        self.autorun_action = QAction("开机自启动", self)
        self.autorun_action.setCheckable(True)
        self.autorun_action.setChecked(self.is_autorun())
        self.autorun_action.triggered.connect(self.toggle_autorun)
        menu.addAction(self.autorun_action)
        
        self.quit_action = QAction("退出", self)
        self.quit_action.triggered.connect(QApplication.instance().quit)
        menu.addAction(self.quit_action)
        
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()
    
    def is_autorun(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, "SeiraiPPTAssistant")
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False

    def toggle_autorun(self, checked):
        app_path = os.path.abspath(sys.argv[0])
        # If running as script
        if app_path.endswith('.py'):
            # Use pythonw.exe to avoid console if available, otherwise sys.executable
            python_exe = sys.executable.replace("python.exe", "pythonw.exe")
            if not os.path.exists(python_exe):
                python_exe = sys.executable
            cmd = f'"{python_exe}" "{app_path}"'
        else:
            # Frozen exe
            cmd = f'"{sys.executable}"'

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            if checked:
                winreg.SetValueEx(key, "SeiraiPPTAssistant", 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, "SeiraiPPTAssistant")
                except WindowsError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error setting autorun: {e}")

    def get_ppt_view(self):
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            if self.ppt_app.SlideShowWindows.Count > 0:
                return self.ppt_app.SlideShowWindows(1).View
            else:
                return None
        except Exception:
            return None

    def check_state(self):
        view = self.get_ppt_view()
        if view:
            if not self.widgets_visible:
                self.show_widgets()
            
            # Update ppt_app reference for nav widgets
            if self.ppt_app:
                self.nav_left.ppt_app = self.ppt_app
                self.nav_right.ppt_app = self.ppt_app
            
            self.sync_state(view)
            self.update_page_num(view)
        else:
            if self.widgets_visible:
                self.hide_widgets()

    def show_widgets(self):
        self.toolbar.show()
        self.nav_left.show()
        self.nav_right.show()
        self.adjust_positions()
        self.widgets_visible = True

    def hide_widgets(self):
        self.toolbar.hide()
        self.nav_left.hide()
        self.nav_right.hide()
        self.widgets_visible = False

    def adjust_positions(self):
        screen = QApplication.primaryScreen().geometry()
        MARGIN = 20
        
        # Toolbar: Bottom Center
        tb_w = self.toolbar.sizeHint().width()
        tb_h = self.toolbar.sizeHint().height()
        
        self.toolbar.setGeometry(
            (screen.width() - tb_w) // 2,
            screen.height() - tb_h - MARGIN, # Flush bottom
            tb_w, tb_h
        )
        nav_w = self.nav_left.sizeHint().width()
        nav_h = self.nav_left.sizeHint().height()
        y = screen.height() - nav_h - MARGIN
        
        self.nav_left.setGeometry(
            MARGIN,
            y,
            nav_w, nav_h
        )
        
        self.nav_right.setGeometry(
            screen.width() - nav_w - MARGIN,
            y,
            nav_w, nav_h
        )

    def sync_state(self, view):
        try:
            pt = view.PointerType
            if pt == 1:
                self.toolbar.btn_arrow.setChecked(True)
            elif pt == 2:
                self.toolbar.btn_pen.setChecked(True)
            elif pt == 5: # Eraser
                self.toolbar.btn_eraser.setChecked(True)
        except:
            pass

    def update_page_num(self, view):
        try:
            current = view.Slide.SlideIndex
            total = self.ppt_app.ActivePresentation.Slides.Count
            self.nav_left.update_page(current, total)
            self.nav_right.update_page(current, total)
        except:
            pass

    def go_prev(self):
        view = self.get_ppt_view()
        if view:
            try:
                view.Previous()
            except:
                pass

    def go_next(self):
        view = self.get_ppt_view()
        if view:
            try:
                view.Next()
            except:
                pass
                
    def jump_to_slide(self, index):
        view = self.get_ppt_view()
        if view:
            try:
                view.GotoSlide(index)
            except:
                pass

    def set_pointer(self, type_id):
        view = self.get_ppt_view()
        if view:
            try:
                # If switching to eraser (5)
                if type_id == 5:
                    # Check for ink but DO NOT BLOCK
                    if not self.has_ink():
                        self.show_warning(None, "当前页没有笔迹")
                
                view.PointerType = type_id
                self.activate_ppt_window()
            except:
                pass
    
    def set_pen_color(self, color):
        view = self.get_ppt_view()
        if view:
            try:
                view.PointerType = 2 # Switch to pen first
                view.PointerColor.RGB = color
                self.activate_ppt_window()
            except:
                pass

    def activate_ppt_window(self):
        try:
            # Try to get the window handle of the slide show
            hwnd = self.ppt_app.SlideShowWindows(1).HWND
            # Force bring to foreground
            win32gui.SetForegroundWindow(hwnd)
        except:
            pass
                
    def clear_ink(self):
        view = self.get_ppt_view()
        if view:
            try:
                if not self.has_ink():
                    self.show_warning(None, "当前页没有笔迹")
                view.EraseDrawing()
            except:
                pass
                
    def toggle_spotlight(self):
        if self.spotlight.isVisible():
            self.spotlight.hide()
        else:
            self.spotlight.showFullScreen()
            
    def exit_slideshow(self):
        view = self.get_ppt_view()
        if view:
            try:
                view.Exit()
            except:
                pass

if __name__ == '__main__':
    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    controller = MainController()
    sys.exit(app.exec())

