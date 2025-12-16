import sys
import os
from PyQt6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QButtonGroup, 
                             QLabel, QFrame, QPushButton, QGridLayout, QStackedWidget,
                             QScrollArea)
from PyQt6.QtCore import (Qt, QSize, pyqtSignal, QEvent, QRect, QPropertyAnimation, 
                          QEasingCurve, QAbstractAnimation, QTimer, QSequentialAnimationGroup, 
                          QPoint, QThread)
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen
from qfluentwidgets import (TransparentToolButton, ToolButton, SpinBox,
                            PrimaryPushButton, PushButton, TabWidget,
                            ToolTipFilter, ToolTipPosition, Flyout, FlyoutAnimationType,
                            Pivot, SegmentedWidget, TimePicker, Theme, isDarkTheme,
                            FluentIcon, StrongBodyLabel, TitleLabel, LargeTitleLabel,
                            BodyLabel, CaptionLabel, IndeterminateProgressRing,
                            SmoothScrollArea, FlowLayout)
from qfluentwidgets.components.material import AcrylicFlyout
from .detached_flyout import DetachedFlyoutWindow

def icon_path(name):
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # 返回上级目录的icons文件夹路径
    return os.path.join(os.path.dirname(base_dir), "icons", name)

def get_icon(name, theme=Theme.DARK):
    path = icon_path(name)
    if not os.path.exists(path):
        return QIcon()
        
    if theme == Theme.LIGHT:
        # Read SVG and replace white with black
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            # Simple replacement of fill="white" or fill="#ffffff" or stroke="white"
            # This is a heuristic, might need adjustment
            content = content.replace('fill="white"', 'fill="#333333"')
            content = content.replace('fill="#ffffff"', 'fill="#333333"')
            content = content.replace('stroke="white"', 'stroke="#333333"')
            content = content.replace('stroke="#ffffff"', 'stroke="#333333"')
            
            pixmap = QPixmap()
            pixmap.loadFromData(content.encode('utf-8'))
            return QIcon(pixmap)
        except:
            return QIcon(path)
    else:
        return QIcon(path)


class IconFactory:
    @staticmethod
    def draw_cursor(color):#笔相关类
        from PyQt6.QtGui import QPixmap, QPainter, QPen, QBrush, QColor
        from PyQt6.QtCore import Qt, QPointF
        from PyQt6.QtGui import QPolygonF
        
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
    def draw_arrow(color, direction='left'):#绘画托盘菜单里的箭头图标……
        from PyQt6.QtGui import QPixmap, QPainter, QPen, QColor, QPoint
        from PyQt6.QtCore import Qt
        
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


class PenSettingsFlyout(QWidget):
    color_selected = pyqtSignal(int)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(240, 200)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        title = StrongBodyLabel("墨迹颜色", self)
        layout.addWidget(title)
        
        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)
        self.grid_layout.setSpacing(8)
        
        colors = [
            (255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0),
            (255, 0, 255), (0, 255, 255), (0, 0, 0), (255, 255, 255),
            (255, 165, 0), (128, 0, 128)
        ]
        
        for i, rgb in enumerate(colors):
            btn = QPushButton()
            btn.setFixedSize(32, 32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            
            color_hex = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            ppt_color = rgb[0] + (rgb[1] << 8) + (rgb[2] << 16)
            
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color_hex};
                    border: 2px solid rgba(128, 128, 128, 0.2);
                    border-radius: 16px;
                }}
                QPushButton:hover {{
                    border: 2px solid rgba(128, 128, 128, 0.8);
                }}
                QPushButton:pressed {{
                    border: 2px solid white;
                }}
            """)
            
            btn.clicked.connect(lambda checked, c=ppt_color: self.on_color_clicked(c))
            row = i // 5
            col = i % 5
            self.grid_layout.addWidget(btn, row, col)
            
        layout.addWidget(self.grid_widget)
        layout.addStretch()
        
    def on_color_clicked(self, color):
        self.color_selected.emit(color)


class EraserSettingsFlyout(QWidget):
    clear_all_clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(220, 100)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        title = StrongBodyLabel("橡皮擦选项", self)
        layout.addWidget(title)
        
        btn = PrimaryPushButton("清除当前页笔迹", self)
        btn.setFixedSize(188, 32)
        btn.clicked.connect(self.on_clicked)
        layout.addWidget(btn)
        
    def on_clicked(self):
        self.clear_all_clicked.emit()


class SlidePreviewCard(QFrame):
    clicked = pyqtSignal(int)
    
    def __init__(self, index, image_path, parent=None):
        super().__init__(parent)
        self.index = index
        self.setFixedSize(140, 100) 
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet("""
            SlidePreviewCard {
                background-color: transparent;
                border: 1px solid rgba(0, 0, 0, 0.1);
                border-radius: 6px;
            }
            SlidePreviewCard:hover {
                background-color: rgba(0, 0, 0, 0.05);
                border: 1px solid rgba(0, 0, 0, 0.2);
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)
        
        self.img_label = QLabel()
        self.img_label.setFixedSize(130, 73) # 16:9 approx
        self.img_label.setStyleSheet("background-color: rgba(128, 128, 128, 0.2); border-radius: 4px;")
        self.img_label.setScaledContents(True)
        if image_path and os.path.exists(image_path):
            self.img_label.setPixmap(QPixmap(image_path))
            
        self.txt_label = CaptionLabel(f"{index}", self)
        self.txt_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(self.img_label)
        layout.addWidget(self.txt_label)
        
    def mousePressEvent(self, event):
        self.clicked.emit(self.index)


class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.ring = IndeterminateProgressRing(self)
        self.ring.setFixedSize(60, 60)
        
        layout.addWidget(self.ring)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(0, 0, 0, 80)) # Reduced opacity (more transparent)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRect(self.rect())


class SlideSelectorFlyout(QWidget):
    slide_selected = pyqtSignal(int)
    
    def __init__(self, ppt_app, parent=None):
        super().__init__(parent)
        self.ppt_app = ppt_app
        self.setFixedSize(480, 520)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        
        title = StrongBodyLabel("幻灯片预览", self)
        layout.addWidget(title)
        
        self.scroll = SmoothScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background-color: transparent; border: none;")
        
        self.container = QWidget()
        self.container.setStyleSheet("background-color: transparent;")
        self.flow = FlowLayout(self.container)
        self.flow.setContentsMargins(0, 0, 0, 0)
        self.flow.setHorizontalSpacing(12)
        self.flow.setVerticalSpacing(12)
        
        self.scroll.setWidget(self.container)
        layout.addWidget(self.scroll)
        
        # We will load slides in background if possible, but here we just trigger load
        QTimer.singleShot(10, self.load_slides)
        
    def get_cache_dir(self, presentation_path):
        import hashlib
        import os
        try:
            path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
        except:
            path_hash = "default"
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
                
                # Check if exists (assuming cache is handled by BusinessLogic or previous run)
                # If not, we might lag here exporting. 
                # Ideally BusinessLogic pre-caches.
                if not os.path.exists(thumb_path):
                    try:
                        slide.Export(thumb_path, "JPG", 320, 180) 
                    except:
                        pass
                    
                card = SlidePreviewCard(i, thumb_path)
                card.clicked.connect(self.on_card_clicked)
                self.flow.addWidget(card)
                
        except Exception as e:
            print(f"Error loading slides: {e}")
            
    def on_card_clicked(self, index):
        self.slide_selected.emit(index)


class CompatibilityAnnotationWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        # 绘图属性
        self.drawing = False
        self.erasing = False
        self.last_point = None
        self.pen_color = Qt.GlobalColor.red
        self.pen_width = 3
        
        # 存储绘制的线条
        self.lines = []
        self.current_line = None
        
        # 橡皮擦大小
        self.eraser_size = 20
        
        self.current_theme = Theme.DARK
        
        # 工具栏
        self.setup_toolbar()
        
    def setup_toolbar(self):
        """设置工具栏"""
        toolbar_width = 40
        toolbar_height = 240
        margin = 20
        
        # 创建工具栏容器
        self.toolbar = QWidget(self)
        self.toolbar.setGeometry(margin, margin, toolbar_width, toolbar_height)
        
        # 工具栏布局
        from PyQt6.QtWidgets import QVBoxLayout
        layout = QVBoxLayout(self.toolbar)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        # 笔工具按钮
        self.btn_pen = QPushButton()
        self.btn_pen.setIcon(get_icon("Pen.svg", self.current_theme))
        self.btn_pen.setFixedSize(30, 30)
        self.btn_pen.setCheckable(True)
        self.btn_pen.clicked.connect(self.set_pen_mode)
        
        # 橡皮擦按钮
        self.btn_eraser = QPushButton()
        self.btn_eraser.setIcon(get_icon("Eraser.svg", self.current_theme))
        self.btn_eraser.setFixedSize(30, 30)
        self.btn_eraser.setCheckable(True)
        self.btn_eraser.clicked.connect(self.set_eraser_mode)
        
        # 清除按钮
        self.btn_clear = QPushButton()
        self.btn_clear.setIcon(get_icon("Clear.svg", self.current_theme))
        self.btn_clear.setFixedSize(30, 30)
        self.btn_clear.clicked.connect(self.clear_all)
        
        # 关闭按钮
        self.btn_close = QPushButton("X")
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.clicked.connect(self.close)
        
        # 颜色选择按钮
        self.btn_color_red = QPushButton()
        self.btn_color_red.setFixedSize(30, 30)
        self.btn_color_red.setStyleSheet("background-color: red; border-radius: 4px; border: none;")
        self.btn_color_red.clicked.connect(lambda: self.set_pen_color(Qt.GlobalColor.red))
        
        self.btn_color_blue = QPushButton()
        self.btn_color_blue.setFixedSize(30, 30)
        self.btn_color_blue.setStyleSheet("background-color: blue; border-radius: 4px; border: none;")
        self.btn_color_blue.clicked.connect(lambda: self.set_pen_color(Qt.GlobalColor.blue))
        
        self.btn_color_green = QPushButton()
        self.btn_color_green.setFixedSize(30, 30)
        self.btn_color_green.setStyleSheet("background-color: green; border-radius: 4px; border: none;")
        self.btn_color_green.clicked.connect(lambda: self.set_pen_color(Qt.GlobalColor.green))
        
        # 添加按钮到布局
        layout.addWidget(self.btn_pen)
        layout.addWidget(self.btn_eraser)
        layout.addWidget(self.btn_clear)
        layout.addWidget(self.btn_color_red)
        layout.addWidget(self.btn_color_blue)
        layout.addWidget(self.btn_color_green)
        layout.addWidget(self.btn_close)
        
        # 默认选择笔工具
        self.btn_pen.setChecked(True)
        
        self.set_theme(Theme.DARK)

    def set_theme(self, theme):
        self.current_theme = theme
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            btn_bg = "rgba(0, 0, 0, 0.05)"
            btn_checked = "#00cc7a"
            btn_hover = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            btn_bg = "rgba(255, 255, 255, 0.1)"
            btn_checked = "#00cc7a"
            btn_hover = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            
        self.toolbar.setStyleSheet(f"""
            QWidget {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        btn_style = f"""
            QPushButton {{
                background-color: {btn_bg};
                color: {text_color};
                border: none;
                border-radius: 4px;
            }}
            QPushButton:checked {{
                background-color: {btn_checked};
                color: white;
            }}
            QPushButton:hover {{
                background-color: {btn_hover};
            }}
        """
        
        self.btn_pen.setStyleSheet(btn_style)
        self.btn_eraser.setStyleSheet(btn_style)
        
        # Update icons
        self.btn_pen.setIcon(get_icon("Pen.svg", theme))
        self.btn_eraser.setIcon(get_icon("Eraser.svg", theme))
        self.btn_clear.setIcon(get_icon("Clear.svg", theme))
        
        # Clear button style (reddish)
        self.btn_clear.setStyleSheet(f"""
            QPushButton {{
                background-color: rgba(255, 50, 50, 0.3);
                color: {text_color};
                border: none;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: rgba(255, 50, 50, 0.5);
            }}
        """)
        
        # Close button style
        self.btn_close.setStyleSheet(f"""
            QPushButton {{
                background-color: rgba(255, 50, 50, 0.3);
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: rgba(255, 50, 50, 0.5);
            }}
        """)
        
    def set_pen_mode(self):
        """设置笔模式"""
        self.btn_pen.setChecked(True)
        self.btn_eraser.setChecked(False)
        self.drawing = True
        self.erasing = False
        
    def set_eraser_mode(self):
        """设置橡皮擦模式"""
        self.btn_eraser.setChecked(True)
        self.btn_pen.setChecked(False)
        self.drawing = False
        self.erasing = True
        
    def set_pen_color(self, color):
        """设置笔颜色"""
        self.pen_color = color
        
    def clear_all(self):
        """清除所有批注"""
        self.lines = []
        self.update()
        
    def mousePressEvent(self, event):
        from PyQt6.QtCore import Qt
        if event.button() == Qt.MouseButton.RightButton:
            self.close()
        elif event.button() == Qt.MouseButton.LeftButton:
            if self.drawing:
                # 开始绘制新线条
                self.current_line = {
                    'points': [event.pos()],
                    'color': self.pen_color,
                    'width': self.pen_width
                }
            elif self.erasing:
                # 开始橡皮擦操作
                self.erase_at_point(event.pos())
                
    def mouseMoveEvent(self, event):
        if self.drawing and self.current_line:
            # 添加点到当前线条
            self.current_line['points'].append(event.pos())
            self.update()
        elif self.erasing:
            # 橡皮擦操作
            self.erase_at_point(event.pos())
            
    def mouseReleaseEvent(self, event):
        from PyQt6.QtCore import Qt, QPoint
        if event.button() == Qt.MouseButton.LeftButton:
            if self.drawing and self.current_line:
                # 完成当前线条
                if len(self.current_line['points']) > 1:
                    self.lines.append(self.current_line)
                self.current_line = None
                
    def erase_at_point(self, point):
        """在指定点进行擦除操作"""
        # 简单实现：移除靠近点击点的线条
        lines_to_remove = []
        for line in self.lines:
            for p in line['points']:
                # 计算点之间的距离
                distance = ((p.x() - point.x()) ** 2 + (p.y() - point.y()) ** 2) ** 0.5
                if distance < self.eraser_size:
                    lines_to_remove.append(line)
                    break
                    
        # 移除需要擦除的线条
        for line in lines_to_remove:
            if line in self.lines:
                self.lines.remove(line)
                
        self.update()
        
    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter, QPen
        from PyQt6.QtCore import Qt
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 绘制已保存的线条
        for line in self.lines:
            if len(line['points']) > 1:
                pen = QPen(line['color'], line['width'], Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
                painter.setPen(pen)
                for i in range(len(line['points']) - 1):
                    painter.drawLine(line['points'][i], line['points'][i + 1])
                    
        # 绘制当前正在绘制的线条
        if self.current_line and len(self.current_line['points']) > 1:
            pen = QPen(self.current_line['color'], self.current_line['width'], Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            for i in range(len(self.current_line['points']) - 1):
                painter.drawLine(self.current_line['points'][i], self.current_line['points'][i + 1])


class AnnotationWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        self.drawing = False
        self.last_point = None
        self.pen_color = Qt.GlobalColor.red
        self.pen_width = 3
        
        # Close button
        self.btn_close = QPushButton("X", self)
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.setStyleSheet("background-color: red; color: white; border-radius: 0px; font-weight: bold;")
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
    def set_pen_color(self, color):
        self.pen_color = color
        
    def set_pen_width(self, width):
        self.pen_width = width
        
    def mousePressEvent(self, event):
        from PyQt6.QtCore import Qt
        if event.button() == Qt.MouseButton.RightButton:
            self.close()
        elif event.button() == Qt.MouseButton.LeftButton:
            self.drawing = True
            self.last_point = event.pos()
            
    def mouseMoveEvent(self, event):
        if self.drawing and self.last_point:
            from PyQt6.QtGui import QPainter, QPen
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            pen = QPen(self.pen_color, self.pen_width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            painter.drawLine(self.last_point, event.pos())
            self.last_point = event.pos()
            self.update()
            
    def mouseReleaseEvent(self, event):
        from PyQt6.QtCore import Qt, QPoint
        if event.button() == Qt.MouseButton.LeftButton:
            self.drawing = False
            self.last_point = None
            
    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        # Paint implementation would go here if needed


class SpotlightOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        self.selection_rect = QRect()
        self.is_selecting = False
        self.has_selection = False
        self.current_theme = Theme.DARK
        
        # Close button (context aware)
        self.btn_close = TransparentToolButton(FluentIcon.CLOSE, self)
        self.btn_close.setFixedSize(32, 32)
        self.btn_close.setIconSize(QSize(12, 12))
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
        self.set_theme(Theme.DARK)

    def set_theme(self, theme):
        self.current_theme = theme
        self.update()
        if theme == Theme.LIGHT:
            # Light mode style for close button (when visible on top of overlay)
            # Since overlay is dark (dimmed screen), close button should probably remain light/white for visibility?
            # Or if it's on top of content...
            # The overlay background is black with alpha. So white icon is best.
            pass
        
        # We'll keep the close button style consistent: white on red or just white?
        # Standard WinUI close is usually red on hover.
        # TransparentToolButton handles this.
        # But since we are on a dark overlay (0,0,0,180), we should force dark theme style for the button
        # so it has white icon.
        self.btn_close.setIcon(FluentIcon.CLOSE)
        self.btn_close.setStyleSheet("""
            TransparentToolButton {
                background-color: #cc0000;
                border: none;
                border-radius: 4px;
                color: white;
            }
            TransparentToolButton:hover {
                background-color: #e60000;
            }
            TransparentToolButton:pressed {
                background-color: #b30000;
            }
        """)

    def mousePressEvent(self, event):
        from PyQt6.QtGui import QMouseEvent
        from PyQt6.QtCore import Qt
        if event.button() == Qt.MouseButton.RightButton:
            self.close()
        elif event.button() == Qt.MouseButton.LeftButton:
            self.selection_rect.setTopLeft(event.pos())
            self.selection_rect.setBottomRight(event.pos())
            self.is_selecting = True
            self.has_selection = False
            self.btn_close.hide()
            self.update()
            
    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.selection_rect.setBottomRight(event.pos())
            self.update()
            
    def mouseReleaseEvent(self, event):
        from PyQt6.QtCore import Qt, QPoint
        if self.is_selecting:
            self.is_selecting = False
            self.has_selection = True
            # Show close button near top-right of selection
            normalized_rect = self.selection_rect.normalized()
            self.btn_close.move(normalized_rect.topRight() + QPoint(10, -15))
            self.btn_close.show()
            self.update()
            
    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter, QColor, QPen
        from PyQt6.QtCore import Qt
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fill screen with semi-transparent color based on theme
        if self.current_theme == Theme.LIGHT:
            painter.setBrush(QColor(255, 255, 255, 180))
        else:
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
    prev_clicked = pyqtSignal()
    next_clicked = pyqtSignal()
    
    def __init__(self, parent=None, is_right=False):
        super().__init__(parent)
        self.ppt_app = None 
        self.is_right = is_right
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(8, 6, 8, 6) 
        inner_layout.setSpacing(15) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(52)
        self.btn_prev = TransparentToolButton(parent=self)
        self.btn_prev.setFixedSize(36, 36) 
        self.btn_prev.setIconSize(QSize(18, 18))
        self.btn_prev.setToolTip("上一页")
        self.btn_prev.installEventFilter(ToolTipFilter(self.btn_prev, 1000, ToolTipPosition.TOP))
        self.btn_prev.clicked.connect(self.prev_clicked.emit)
        
        self.btn_next = TransparentToolButton(parent=self)
        self.btn_next.setFixedSize(36, 36) 
        self.btn_next.setIconSize(QSize(18, 18))
        self.btn_next.setToolTip("下一页")
        self.btn_next.installEventFilter(ToolTipFilter(self.btn_next, 1000, ToolTipPosition.TOP))
        self.btn_next.clicked.connect(self.next_clicked.emit)

        # Page Info Clickable Area
        self.page_info_widget = QWidget()
        self.page_info_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.page_info_widget.installEventFilter(self)
        
        from PyQt6.QtWidgets import QVBoxLayout
        info_layout = QVBoxLayout(self.page_info_widget)
        info_layout.setContentsMargins(10, 0, 10, 0)
        info_layout.setSpacing(2)
        info_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_page_num = QLabel("1/--")
        self.lbl_page_text = QLabel("页码")
        
        info_layout.addWidget(self.lbl_page_num, 0, Qt.AlignmentFlag.AlignCenter)
        info_layout.addWidget(self.lbl_page_text, 0, Qt.AlignmentFlag.AlignCenter)
        
        inner_layout.addWidget(self.btn_prev)
        
        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.Shape.VLine)
        inner_layout.addWidget(self.line1)
        
        inner_layout.addWidget(self.page_info_widget)
        
        self.line2 = QFrame()
        self.line2.setFrameShape(QFrame.Shape.VLine)
        inner_layout.addWidget(self.line2)
        
        inner_layout.addWidget(self.btn_next)
        
        layout.addWidget(self.container)
        self.setLayout(layout)

        self.setup_click_feedback(self.btn_prev, QSize(18, 18))
        self.setup_click_feedback(self.btn_next, QSize(18, 18))
        
        self.set_theme(Theme.AUTO)
        
    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            subtext_color = "#666666"
            line_color = "rgba(0, 0, 0, 0.1)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            subtext_color = "#aaaaaa"
            line_color = "rgba(255, 255, 255, 0.2)"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 12px;
            }}
            QLabel {{
                font-family: "Segoe UI", "Microsoft YaHei";
                color: {text_color};
            }}
        """)
        
        self.lbl_page_num.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {text_color};")
        self.lbl_page_text.setStyleSheet(f"font-size: 12px; color: {subtext_color};")
        self.line1.setStyleSheet(f"color: {line_color};")
        self.line2.setStyleSheet(f"color: {line_color};")
        
        self.btn_prev.setIcon(get_icon("Previous.svg", theme))
        self.btn_next.setIcon(get_icon("Next.svg", theme))
        
        self.style_nav_btn(self.btn_prev, theme)
        self.style_nav_btn(self.btn_next, theme)

    def style_nav_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            pressed_bg = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            pressed_bg = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: {text_color};
            }}
            TransparentToolButton:hover {{
                background-color: {hover_bg};
            }}
            TransparentToolButton:pressed {{
                background-color: {pressed_bg};
            }}
        """)
    
    def setup_click_feedback(self, btn, base_size):
        anim = QPropertyAnimation(btn, b"iconSize", self)
        anim.setDuration(120)
        anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

        def on_pressed():
            if anim.state() == QAbstractAnimation.State.Running:
                anim.stop()
            shrink = QSize(int(base_size.width() * 0.85), int(base_size.height() * 0.85))
            anim.setStartValue(base_size)
            anim.setKeyValueAt(0.5, shrink)
            anim.setEndValue(base_size)
            anim.start()

        btn.pressed.connect(on_pressed)
    
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
        
        # Ensure view has a background
        if self.current_theme == Theme.LIGHT:
            bg_color = "#f3f3f3"
        else:
            bg_color = "#202020"
        view.setStyleSheet(f"SlideSelectorFlyout {{ background-color: {bg_color}; border-radius: 8px; }}")

        win = DetachedFlyoutWindow(view, self)
        view.slide_selected.connect(win.close)
        win.show_at(self.page_info_widget)

    def update_page(self, current, total):
        self.lbl_page_num.setText(f"{current}/{total}")
    
    def apply_settings(self):
        self.btn_prev.setToolTip("上一页")
        self.btn_next.setToolTip("下一页")
        self.lbl_page_text.setText("页码")
        self.style_nav_btn(self.btn_prev, self.current_theme)
        self.style_nav_btn(self.btn_next, self.current_theme)


class ToolBarWidget(QWidget):
    request_spotlight = pyqtSignal()
    request_pointer_mode = pyqtSignal(int)
    request_pen_color = pyqtSignal(int)
    request_clear_ink = pyqtSignal()
    request_exit = pyqtSignal()
    request_timer = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.current_theme = Theme.DARK
        self.was_checked = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        container_layout = QHBoxLayout(self.container)
        container_layout.setContentsMargins(8, 6, 8, 6) 
        container_layout.setSpacing(12) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(56)
        
        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        
        self.btn_arrow = self.create_tool_btn("选择", "Mouse.svg")
        self.btn_arrow.clicked.connect(lambda: self.request_pointer_mode.emit(1))
        
        self.btn_pen = self.create_tool_btn("笔", "Pen.svg")
        self.btn_pen.clicked.connect(lambda: self.request_pointer_mode.emit(2))
        
        self.btn_eraser = self.create_tool_btn("橡皮", "Eraser.svg")
        self.btn_eraser.clicked.connect(lambda: self.request_pointer_mode.emit(5))
        
        self.btn_clear = self.create_action_btn("一键清除", "Clear.svg")
        self.btn_clear.clicked.connect(self.request_clear_ink.emit)
        
        self.group.addButton(self.btn_arrow)
        self.group.addButton(self.btn_pen)
        self.group.addButton(self.btn_eraser)
        
        self.btn_spotlight = self.create_action_btn("聚焦", "Select.svg")
        self.btn_spotlight.clicked.connect(self.request_spotlight.emit)
        self.btn_timer = self.create_action_btn("计时器", "timer.svg")
        self.btn_timer.clicked.connect(self.request_timer.emit)

        self.btn_exit = self.create_action_btn("结束放映", "Minimaze.svg")
        self.btn_exit.clicked.connect(self.request_exit.emit)
        
        container_layout.addWidget(self.btn_arrow)
        container_layout.addWidget(self.btn_pen)
        container_layout.addWidget(self.btn_eraser)
        container_layout.addWidget(self.btn_clear)
        
        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.Shape.VLine)
        container_layout.addWidget(self.line1)
        
        container_layout.addWidget(self.btn_spotlight)
        container_layout.addWidget(self.btn_timer)

        self.line2 = QFrame()
        self.line2.setFrameShape(QFrame.Shape.VLine)
        container_layout.addWidget(self.line2)

        container_layout.addWidget(self.btn_exit)
        
        layout.addWidget(self.container)
        self.setLayout(layout)
        
        self.btn_arrow.setChecked(True)
        
        # Install event filter to detect second click for expansion
        self.btn_pen.installEventFilter(self)
        self.btn_eraser.installEventFilter(self)

        self.setup_click_feedback(self.btn_arrow, QSize(20, 20))
        self.setup_click_feedback(self.btn_pen, QSize(20, 20))
        self.setup_click_feedback(self.btn_eraser, QSize(20, 20))
        self.setup_click_feedback(self.btn_clear, QSize(20, 20))
        self.setup_click_feedback(self.btn_spotlight, QSize(20, 20))
        self.setup_click_feedback(self.btn_timer, QSize(20, 20))
        self.setup_click_feedback(self.btn_exit, QSize(20, 20))
        
        self.set_theme(Theme.AUTO)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            if event.button() == Qt.MouseButton.RightButton:
                if obj == self.btn_pen:
                    self.show_pen_settings()
                    return True
                elif obj == self.btn_eraser:
                    self.show_eraser_settings()
                    return True
            elif event.button() == Qt.MouseButton.LeftButton:
                # If clicking Pen button while it is already checked -> Show settings
                if obj == self.btn_pen and self.btn_pen.isChecked():
                    self.show_pen_settings()
                    return True

        elif event.type() == QEvent.Type.MouseButtonRelease:
            pass
                    
        return super().eventFilter(obj, event)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)

    def show_pen_settings(self):
        view = PenSettingsFlyout(self)
        view.color_selected.connect(self.request_pen_color.emit)
        
        view.setStyleSheet(f"background-color: {self.get_flyout_bg_color()}; border-radius: 8px;")
        
        win = DetachedFlyoutWindow(view, self)
        view.color_selected.connect(win.close)
        win.show_at(self.btn_pen)

    def show_eraser_settings(self):
        view = EraserSettingsFlyout(self)
        view.clear_all_clicked.connect(self.request_clear_ink.emit)
        
        view.setStyleSheet(f"background-color: {self.get_flyout_bg_color()}; border-radius: 8px;")
        
        win = DetachedFlyoutWindow(view, self)
        view.clear_all_clicked.connect(win.close)
        win.show_at(self.btn_eraser)
        
    def get_flyout_bg_color(self):
        if self.current_theme == Theme.LIGHT:
            return "#f3f3f3"
        else:
            return "#202020"

    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        # Update container style
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            line_color = "rgba(0, 0, 0, 0.1)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            line_color = "rgba(255, 255, 255, 0.2)"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        self.line1.setStyleSheet(f"color: {line_color};")
        self.line2.setStyleSheet(f"color: {line_color};")
        
        # Update icons and button styles
        for btn, icon_name in [
            (self.btn_arrow, "Mouse.svg"),
            (self.btn_pen, "Pen.svg"),
            (self.btn_eraser, "Eraser.svg"),
            (self.btn_clear, "Clear.svg"),
            (self.btn_spotlight, "Select.svg"),
            (self.btn_timer, "timer.svg"),
            (self.btn_exit, "Minimaze.svg")
        ]:
            btn.setIcon(get_icon(icon_name, theme))
            self.style_tool_btn(btn, theme) if btn.isCheckable() else self.style_action_btn(btn, theme)

    def create_tool_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(40, 40) 
        btn.setIconSize(QSize(20, 20))
        btn.setCheckable(True)
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        # Style will be set in set_theme
        return btn
        
    def create_action_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(40, 40)
        btn.setIconSize(QSize(20, 20))
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        # Style will be set in set_theme
        return btn
    
    def style_tool_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            checked_bg = "rgba(0, 0, 0, 0.05)"
            text_color = "#333333"
            border_bottom = "#00cc7a"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            checked_bg = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            border_bottom = "#00cc7a"
            
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: {text_color};
                margin-bottom: 2px;
            }}
            TransparentToolButton:hover {{
                background-color: {hover_bg};
            }}
            TransparentToolButton:checked {{
                background-color: {checked_bg};
                color: {text_color};
                border-bottom: 3px solid {border_bottom};
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }}
            TransparentToolButton:checked:hover {{
                background-color: {hover_bg};
            }}
        """)
    
    def style_action_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            pressed_bg = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            pressed_bg = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            
        # Special case for exit button hover color if needed, but keeping it simple for now
        # If it's exit button, we might want red hover. 
        # But the previous code had specific style for btn_exit. 
        # Let's handle btn_exit specifically in set_theme loop or check here.
        
        if btn == self.btn_exit:
             btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 6px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    background-color: rgba(255, 50, 50, 0.3);
                }}
                TransparentToolButton:pressed {{
                    background-color: rgba(255, 50, 50, 0.5);
                }}
            """)
        else:
            btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 6px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    background-color: {hover_bg};
                }}
                TransparentToolButton:pressed {{
                    background-color: {pressed_bg};
                }}
            """)

    def setup_click_feedback(self, btn, base_size):
        anim = QPropertyAnimation(btn, b"iconSize", self)
        anim.setDuration(120)
        anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

        def on_pressed():
            if anim.state() == QAbstractAnimation.State.Running:
                anim.stop()
            shrink = QSize(int(base_size.width() * 0.85), int(base_size.height() * 0.85))
            anim.setStartValue(base_size)
            anim.setKeyValueAt(0.5, shrink)
            anim.setEndValue(base_size)
            anim.start()

        btn.pressed.connect(on_pressed)


class TimerWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(340, 280)
        
        self.up_seconds = 0
        self.up_running = False
        self.down_total_seconds = 0
        self.down_remaining = 0
        self.down_running = False
        self.sound_effect = None
        self.drag_pos = None
        
        # Main Container
        self.container = QWidget(self)
        self.container.setObjectName("Container")
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.addWidget(self.container)
        
        self.layout = QVBoxLayout(self.container)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)
        
        # Header
        header_layout = QHBoxLayout()
        self.title_label = TitleLabel("计时工具", self)
        self.close_btn = TransparentToolButton(FluentIcon.CLOSE, self)
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setIconSize(QSize(12, 12))
        self.close_btn.clicked.connect(self.close)
        header_layout.addWidget(self.title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.close_btn)
        self.layout.addLayout(header_layout)
        
        # Segmented Widget for switching modes
        self.pivot = SegmentedWidget(self)
        self.pivot.addItem("up", "正计时")
        self.pivot.addItem("down", "倒计时")
        self.pivot.currentItemChanged.connect(self.on_pivot_changed)
        self.layout.addWidget(self.pivot)
        
        # Content Stack
        self.stack = QStackedWidget(self)
        self.up_page = QWidget()
        self.down_page = QWidget()
        self.completed_page = QWidget()
        self.setup_up_page()
        self.setup_down_page()
        self.setup_completed_page()
        self.stack.addWidget(self.up_page)
        self.stack.addWidget(self.down_page)
        self.stack.addWidget(self.completed_page)
        self.layout.addWidget(self.stack)
        
        self.init_timers()
        self.init_sound()
        self.set_theme(Theme.DARK)
        
        self.pivot.setCurrentItem("up")

    def on_pivot_changed(self, route_key):
        if route_key == "up":
            self.stack.setCurrentWidget(self.up_page)
        else:
            self.stack.setCurrentWidget(self.down_page)

    def setup_up_page(self):
        layout = QVBoxLayout(self.up_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        self.up_label = QLabel("00:00")
        self.up_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addStretch()
        layout.addWidget(self.up_label)
        layout.addStretch()
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        self.up_start_btn = PrimaryPushButton("开始", self.up_page)
        self.up_start_btn.setFixedWidth(100)
        self.up_reset_btn = PushButton("重置", self.up_page)
        self.up_reset_btn.setFixedWidth(100)
        
        self.up_start_btn.clicked.connect(self.toggle_up)
        self.up_reset_btn.clicked.connect(self.reset_up)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.up_start_btn)
        btn_layout.addWidget(self.up_reset_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addSpacing(10)

    def setup_completed_page(self):
        layout = QVBoxLayout(self.completed_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        self.completed_label = QLabel("倒计时已结束")
        self.completed_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addStretch()
        layout.addWidget(self.completed_label)
        layout.addSpacing(20)
        
        self.back_btn = PrimaryPushButton("返回", self.completed_page)
        self.back_btn.setFixedWidth(120)
        self.back_btn.clicked.connect(self.on_completed_back)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(self.back_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addStretch()

    def on_completed_back(self):
        self.reset_down()
        self.stack.setCurrentWidget(self.down_page)

    def setup_down_page(self):
        layout = QVBoxLayout(self.down_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        # Time input area
        self.input_widget = QWidget()
        input_layout = QHBoxLayout(self.input_widget)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(10)
        
        self.down_min_spin = SpinBox()
        self.down_min_spin.setRange(0, 999)
        self.down_min_spin.setSuffix(" 分")
        self.down_min_spin.setFixedWidth(110)
        self.down_min_spin.setValue(5) # Default 5 mins
        
        self.down_sec_spin = SpinBox()
        self.down_sec_spin.setRange(0, 59)
        self.down_sec_spin.setSuffix(" 秒")
        self.down_sec_spin.setFixedWidth(110)
        
        input_layout.addStretch()
        input_layout.addWidget(self.down_min_spin)
        input_layout.addWidget(self.down_sec_spin)
        input_layout.addStretch()
        
        self.down_label = QLabel("00:00")
        self.down_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.down_label.hide()
        
        layout.addStretch()
        layout.addWidget(self.input_widget)
        layout.addWidget(self.down_label)
        layout.addStretch()
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        self.down_start_btn = PrimaryPushButton("开始", self.down_page)
        self.down_start_btn.setFixedWidth(100)
        self.down_reset_btn = PushButton("重置", self.down_page)
        self.down_reset_btn.setFixedWidth(100)
        
        self.down_start_btn.clicked.connect(self.toggle_down)
        self.down_reset_btn.clicked.connect(self.reset_down)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.down_start_btn)
        btn_layout.addWidget(self.down_reset_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addSpacing(10)

    def set_theme(self, theme):
        self.current_theme = theme
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 248)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            self.title_label.setTextColor("#333333", "#333333")
        else:
            bg_color = "rgba(32, 32, 32, 248)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            self.title_label.setTextColor("white", "white")
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        font_style = f"font-size: 56px; font-weight: bold; color: {text_color}; font-family: 'Segoe UI', 'Microsoft YaHei';"
        self.up_label.setStyleSheet(font_style)
        self.down_label.setStyleSheet(font_style)
        
        completed_style = f"font-size: 24px; font-weight: bold; color: {text_color}; font-family: 'Segoe UI', 'Microsoft YaHei';"
        if hasattr(self, 'completed_label'):
            self.completed_label.setStyleSheet(completed_style)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self.drag_pos:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()
            
    def mouseReleaseEvent(self, event):
        self.drag_pos = None

    def shake_window(self):
        self.shake_anim = QSequentialAnimationGroup(self)
        original_pos = self.pos()
        offset = 10
        
        for _ in range(2):
            anim1 = QPropertyAnimation(self, b"pos")
            anim1.setDuration(50)
            anim1.setStartValue(original_pos)
            anim1.setEndValue(original_pos + QPoint(offset, 0))
            anim1.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            anim2 = QPropertyAnimation(self, b"pos")
            anim2.setDuration(50)
            anim2.setStartValue(original_pos + QPoint(offset, 0))
            anim2.setEndValue(original_pos - QPoint(offset, 0))
            anim2.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            anim3 = QPropertyAnimation(self, b"pos")
            anim3.setDuration(50)
            anim3.setStartValue(original_pos - QPoint(offset, 0))
            anim3.setEndValue(original_pos)
            anim3.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            self.shake_anim.addAnimation(anim1)
            self.shake_anim.addAnimation(anim2)
            self.shake_anim.addAnimation(anim3)
            
        self.shake_anim.start()

    def init_timers(self):
        self.up_timer = QTimer(self)
        self.up_timer.setInterval(1000)
        self.up_timer.timeout.connect(self.update_up)
        self.down_timer = QTimer(self)
        self.down_timer.setInterval(1000)
        self.down_timer.timeout.connect(self.update_down)

    def init_sound(self):
        try:
            from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
            from PyQt6.QtCore import QUrl
        except Exception:
            self.player = None
            return
            
        self.player = QMediaPlayer(self)
        self.audio_output = QAudioOutput(self)
        self.player.setAudioOutput(self.audio_output)
        
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base_dir, "timer_ring.ogg")
        
        if os.path.exists(path):
            self.player.setSource(QUrl.fromLocalFile(path))
            self.audio_output.setVolume(1.0)

    def play_ring(self):
        if hasattr(self, 'player') and self.player:
            self.player.stop()
            self.player.play()

    def format_time(self, seconds):
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def toggle_up(self):
        if not self.up_running:
            self.up_timer.start()
            self.up_running = True
            self.up_start_btn.setText("暂停")
        else:
            self.up_timer.stop()
            self.up_running = False
            self.up_start_btn.setText("开始")

    def reset_up(self):
        self.up_timer.stop()
        self.up_running = False
        self.up_seconds = 0
        self.up_start_btn.setText("开始")
        self.up_label.setText(self.format_time(self.up_seconds))

    def update_up(self):
        self.up_seconds += 1
        self.up_label.setText(self.format_time(self.up_seconds))

    def toggle_down(self):
        if not self.down_running:
            # Check if we are resuming or starting new
            if self.down_remaining <= 0 and self.down_total_seconds == 0:
                # Start new
                minutes = self.down_min_spin.value()
                seconds = self.down_sec_spin.value()
                total = minutes * 60 + seconds
                if total <= 0:
                    return
                self.down_total_seconds = total
                self.down_remaining = total
            
            # Switch to label view
            self.input_widget.hide()
            self.down_label.show()
            self.down_label.setText(self.format_time(self.down_remaining))
            
            self.down_timer.start()
            self.down_running = True
            self.down_start_btn.setText("暂停")
        else:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")

    def reset_down(self):
        self.down_timer.stop()
        self.down_running = False
        self.down_remaining = 0
        self.down_total_seconds = 0
        self.down_start_btn.setText("开始")
        
        # Show input again
        self.down_label.hide()
        self.input_widget.show()

    def update_down(self):
        if self.down_remaining > 0:
            self.down_remaining -= 1
            self.down_label.setText(self.format_time(self.down_remaining))
        
        if self.down_remaining <= 0 and self.down_running:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")
            
            self.stack.setCurrentWidget(self.completed_page)
            self.play_ring()
            self.shake_window()
