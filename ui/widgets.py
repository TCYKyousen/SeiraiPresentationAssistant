import sys
import os
from PyQt6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QButtonGroup, 
                             QLabel, QFrame, QPushButton, QGridLayout, QStackedWidget,
                             QScrollArea)
from PyQt6.QtCore import (Qt, QSize, pyqtSignal, QEvent, QRect, QPropertyAnimation, 
                          QEasingCurve, QAbstractAnimation, QTimer, QSequentialAnimationGroup, QDateTime, 
                          QPoint, QThread, pyqtProperty, QPointF)
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QRadialGradient
from qfluentwidgets import (TransparentToolButton, ToolButton, SpinBox,
                            PrimaryPushButton, PushButton, TabWidget,
                            ToolTipFilter, ToolTipPosition, Flyout, FlyoutAnimationType,
                            Pivot, SegmentedWidget, TimePicker, Theme, isDarkTheme,
                            FluentIcon, StrongBodyLabel, TitleLabel, LargeTitleLabel,
                            BodyLabel, CaptionLabel, IndeterminateProgressRing,
                            SmoothScrollArea, FlowLayout, SwitchButton, MessageBox, MessageDialog)
from qfluentwidgets.components.material import AcrylicFlyout
HEIGHT_BAR = 60

try:
    from .detached_flyout import DetachedFlyoutWindow
except ImportError:
    from detached_flyout import DetachedFlyoutWindow

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
        font = title.font()
        font.setFamily("Meiryo UI")
        title.setFont(font)
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
        layout.setSpacing(20)
        
        self.ring = IndeterminateProgressRing(self)
        self.ring.setFixedSize(48, 48)
        self.ring.setStrokeWidth(4)
        
        self.label = BodyLabel("正在加载幻灯片预览...", self)
        self.label.setStyleSheet("color: white; font-size: 14px;")
        
        layout.addWidget(self.ring)
        layout.addWidget(self.label, 0, Qt.AlignmentFlag.AlignCenter)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        # Smoke layer
        painter.setBrush(QColor(0, 0, 0, 128)) 
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
        
        self.btn_close = TransparentToolButton(FluentIcon.CLOSE, self)
        self.btn_close.setFixedSize(32, 32)
        self.btn_close.setIconSize(QSize(12, 12))
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
        self.set_theme(Theme.DARK)

    def set_theme(self, theme):
        self.current_theme = theme
        self.update()
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
        if self.is_selecting:
            self.is_selecting = False
            self.has_selection = True
            normalized_rect = self.selection_rect.normalized()
            from PyQt6.QtCore import QPoint as _QPoint
            self.btn_close.move(normalized_rect.topRight() + _QPoint(10, -15))
            self.btn_close.show()
            self.update()
            
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
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
            
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            pen = QPen(QColor("#00cc7a"))
            pen.setWidth(3)
            painter.setPen(pen)
            painter.drawRoundedRect(self.selection_rect, 10, 10)


class RippleOverlay(QWidget):
    def __init__(self, center, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen)
        self.center = QPoint(center.x() - screen.left(), center.y() - screen.top())
        self._radius = 0.0
        self.max_radius = max(screen.width(), screen.height()) * 1.2
        self.anim = QPropertyAnimation(self, b"radius", self)
        self.anim.setDuration(800)
        self.anim.setStartValue(0.0)
        self.anim.setEndValue(self.max_radius)
        self.anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.anim.finished.connect(self.close)

    def start(self):
        self.anim.start()

    def get_radius(self):
        return self._radius

    def set_radius(self, value):
        self._radius = float(value)
        self.update()

    radius = pyqtProperty(float, fget=get_radius, fset=set_radius)

    def paintEvent(self, event):
        if self._radius <= 0:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        center_f = QPointF(self.center)
        gradient = QRadialGradient(center_f, self._radius)
        color = QColor(212, 165, 165, 180)
        gradient.setColorAt(0.0, color)
        gradient.setColorAt(1.0, QColor(212, 165, 165, 0))
        painter.setBrush(gradient)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(center_f, self._radius, self._radius)


class PageNavWidget(QWidget):
    request_slide_jump = pyqtSignal(int)
    prev_clicked = pyqtSignal()
    next_clicked = pyqtSignal()
    
    def __init__(self, parent=None, is_right=False):
        super().__init__(parent)
        self.ppt_app = None 
        self.is_right = is_right
        self.current_theme = Theme.DARK
        self.slide_selector_window = None
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        # [修改] 统一边距：左右10，上下5。配合总高度60，内部高度为50
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(8, 0, 8, 0)
        inner_layout.setSpacing(15)
        
        # [修改] 强制锁死内部容器高度为 50
        self.container.setFixedHeight(50)
        
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

        self.page_info_widget = QWidget()
        self.page_info_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.page_info_widget.setObjectName("PageInfo")
        self.page_info_widget.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.page_info_widget.installEventFilter(self)
        
        from PyQt6.QtWidgets import QVBoxLayout
        info_layout = QVBoxLayout(self.page_info_widget)
        info_layout.setContentsMargins(6, 0, 6, 0)
        info_layout.setSpacing(0) # [修改] 减小垂直间距
        info_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_page_num = QLabel("1/--")
        self.lbl_page_num.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_page_num.setFixedWidth(80)
        self.lbl_page_text = QLabel("页码")
        self.lbl_page_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        info_layout.addWidget(self.lbl_page_num, 0, Qt.AlignmentFlag.AlignCenter)
        info_layout.addWidget(self.lbl_page_text, 0, Qt.AlignmentFlag.AlignCenter)
        
        inner_layout.addWidget(self.btn_prev)
        
        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.Shape.VLine)
        self.line1.setFixedHeight(24)
        inner_layout.addWidget(self.line1)
        
        inner_layout.addWidget(self.page_info_widget)
        
        self.line2 = QFrame()
        self.line2.setFrameShape(QFrame.Shape.VLine)
        self.line2.setFixedHeight(24)
        inner_layout.addWidget(self.line2)
        
        inner_layout.addWidget(self.btn_next)
        
        layout.addWidget(self.container)
        self.setFixedHeight(HEIGHT_BAR)
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
            indicator_color = "#D4A5A5"
            indicator_hover = "rgba(212, 165, 165, 0.2)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            subtext_color = "#aaaaaa"
            line_color = "rgba(255, 255, 255, 0.2)"
            indicator_color = "#B38F8F"
            indicator_hover = "rgba(179, 143, 143, 0.3)"
            
        # [修改] 统一圆角为 25px (对应 50px 高度)
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 25px;
            }}
            QWidget#PageInfo {{
                border-radius: 12px;
                background-color: transparent;
            }}
            QWidget#PageInfo:hover {{
                background-color: {indicator_hover};
                background: qradialgradient(spread:pad, cx:0.5, cy:0.5, radius:0.17, fx:0.5, fy:0.5, stop:0 {indicator_color}, stop:1 transparent);
            }}
            QLabel {{
                color: {text_color};
            }}
        """)
        
        # [修改] 修复字体基线问题：
        # 1. 使用 Segoe UI 字体
        # 2. 增加 padding-top 强制下压
        self.lbl_page_num.setStyleSheet(f"""
            font-family: 'Bahnschrift', sans-serif;
            font-size: 16px; 
            font-weight: bold; 
            color: {text_color}; 
            padding-top: 4px;
            background-color: transparent;
        """)
        
        # [修改] 微调下方小字位置
        self.lbl_page_text.setStyleSheet(f"font-size: 10px; color: {subtext_color}; padding-bottom: 2px;")
        
        self.line1.setStyleSheet(f"color: {line_color};")
        self.line2.setStyleSheet(f"color: {line_color};")
        
        self.btn_prev.setIcon(get_icon("Previous.svg", theme))
        self.btn_next.setIcon(get_icon("Next.svg", theme))
        
        self.style_nav_btn(self.btn_prev, theme)
        self.style_nav_btn(self.btn_next, theme)

    def style_nav_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            dot_color = "#D4A5A5"
            hover_bg = "rgba(212, 165, 165, 0.2)"
            pressed_bg = "rgba(212, 165, 165, 0.4)"
            text_color = "#333333"
        else:
            dot_color = "#B38F8F"
            hover_bg = "rgba(179, 143, 143, 0.3)"
            pressed_bg = "rgba(179, 143, 143, 0.5)"
            text_color = "white"
        
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 18px;
                border: none;
                background-color: transparent;
                color: {text_color};
                padding: 0;
            }}
            TransparentToolButton:hover {{
                border-radius: 18px;
                background-color: {hover_bg};
                padding: 0;
            }}
            TransparentToolButton:pressed {{
                border-radius: 18px;
                background-color: {pressed_bg};
                padding: 0;
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
        if self.slide_selector_window and self.slide_selector_window.isVisible():
            self.slide_selector_window.close()
            self.slide_selector_window = None
            return
        view = SlideSelectorFlyout(self.ppt_app)
        view.slide_selected.connect(self.request_slide_jump.emit)
        view.setStyleSheet("SlideSelectorFlyout { background-color: transparent; }")
        win = DetachedFlyoutWindow(view, self)
        view.slide_selected.connect(win.close)
        win.destroyed.connect(self.on_slide_selector_closed)
        self.slide_selector_window = win
        win.show_at(self.page_info_widget)

    def on_slide_selector_closed(self):
        self.slide_selector_window = None

    def update_page(self, current, total):
        from PyQt6.QtCore import QPoint
        # 保持 HTML 结构以支持大小字体，但由于 stylesheet 设置了 font-family 和 padding，这里主要控制相对大小
        html = f"<html><head/><body><p><span style='font-size:16pt;'>{current}</span><span style='font-size:10pt;'>/{total}</span></p></body></html>"
        if not hasattr(self, "page_base_pos") or self.page_base_pos is None:
            self.lbl_page_num.setText(html)
            self.page_base_pos = self.lbl_page_num.pos()
            self.last_page_value = current
            return
        if hasattr(self, "last_page_value") and self.last_page_value == current:
            self.lbl_page_num.setText(html)
            return
        direction = 1
        if hasattr(self, "last_page_value") and self.last_page_value is not None and current < self.last_page_value:
            direction = -1
        self.last_page_value = current
        self.lbl_page_num.setText(html)
        if not hasattr(self, "page_anim") or self.page_anim is None:
            self.page_anim = QPropertyAnimation(self.lbl_page_num, b"pos", self)
            self.page_anim.setDuration(150)
            self.page_anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self.page_anim.stop()
        offset = 6
        start_pos = self.page_base_pos + QPoint(0, offset * direction)
        self.lbl_page_num.move(start_pos)
        self.page_anim.setStartValue(start_pos)
        self.page_anim.setEndValue(self.page_base_pos)
        self.page_anim.start()
    
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
        self.indicator = None
        self.was_checked = False
        self.current_pen_color = None
        self.pen_settings_win = None
        self.eraser_settings_win = None
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        # [修改] 与 PageNavWidget 保持完全一致的边距 (10, 5, 10, 5)
        layout.setContentsMargins(10, 5, 10, 5)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        container_layout = QHBoxLayout(self.container)
        container_layout.setContentsMargins(10, 6, 10, 6)
        container_layout.setSpacing(10)
        
        # [修改] 锁死高度为 50，与页码栏一致
        self.container.setFixedHeight(50)
        
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
        self.setFixedHeight(HEIGHT_BAR)
        
        self.btn_arrow.setChecked(True)
        self.indicator = QFrame(self.container)
        self.indicator.setFixedHeight(2)
        self.indicator.hide()
        
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
        self.update_indicator_for_current()

    def set_pointer_mode(self, mode):
        if mode == 1:
            self.btn_arrow.setChecked(True)
        elif mode == 2:
            self.btn_pen.setChecked(True)
        elif mode == 5:
            self.btn_eraser.setChecked(True)
        self.update_indicator_for_current()

    def show_pen_settings(self):
        if self.pen_settings_win and self.pen_settings_win.isVisible():
            self.pen_settings_win.close()
            self.pen_settings_win = None
            return
        view = PenSettingsFlyout(self)
        view.color_selected.connect(self.on_pen_color_selected)
        view.setStyleSheet("background-color: transparent;")
        win = DetachedFlyoutWindow(view, self)
        view.color_selected.connect(win.close)
        win.destroyed.connect(self.on_pen_settings_closed)
        self.pen_settings_win = win
        win.show_at(self.btn_pen)
    
    def on_pen_color_selected(self, color):
        self.update_pen_display(color)
        self.request_pen_color.emit(color)

    def show_eraser_settings(self):
        if self.eraser_settings_win and self.eraser_settings_win.isVisible():
            self.eraser_settings_win.close()
            self.eraser_settings_win = None
            return
        view = EraserSettingsFlyout(self)
        view.clear_all_clicked.connect(self.request_clear_ink.emit)
        view.setStyleSheet("background-color: transparent;")
        win = DetachedFlyoutWindow(view, self)
        view.clear_all_clicked.connect(win.close)
        win.destroyed.connect(self.on_eraser_settings_closed)
        self.eraser_settings_win = win
        win.show_at(self.btn_eraser)

    def on_pen_settings_closed(self):
        self.pen_settings_win = None

    def on_eraser_settings_closed(self):
        self.eraser_settings_win = None

    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        # Check for window effect
        effect = self.property("windowEffect")
        is_transparent = effect in ["Mica", "Acrylic"]
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)" if not is_transparent else "rgba(255, 255, 255, 10)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            line_color = "rgba(0, 0, 0, 0.1)"
        else:
            bg_color = "rgba(30, 30, 30, 240)" if not is_transparent else "rgba(30, 30, 30, 10)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            line_color = "rgba(255, 255, 255, 0.2)"
            
        # [修改] 统一圆角为 25px，对应 50px 高度
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 25px;
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
        
        indicator_color = "#D4A5A5" if theme == Theme.LIGHT else "#B38F8F"
        if self.indicator:
            self.indicator.setStyleSheet(f"background-color: {indicator_color}; border-radius: 1px;")
            self.update_indicator_for_current()
        self.update_pen_icon_state()

    def create_tool_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(36, 36)
        btn.setIconSize(QSize(20, 20))
        btn.setCheckable(True)
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        btn.toggled.connect(self.on_tool_btn_toggled)
        return btn
        
    def create_action_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(36, 36)
        btn.setIconSize(QSize(20, 20))
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        # Style will be set in set_theme
        return btn
    
    def style_tool_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            checked_bg = "rgba(0, 0, 0, 0.08)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(255, 255, 255, 0.05)"
            checked_bg = "rgba(255, 255, 255, 0.08)"
            text_color = "white"
            
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 18px;
                border: none;
                background-color: transparent;
                color: {text_color};
            }}
            TransparentToolButton:hover {{
                border-radius: 18px;
                background-color: {hover_bg};
            }}
            TransparentToolButton:checked {{
                border-radius: 18px;
                background-color: {checked_bg};
                color: {text_color};
            }}
            TransparentToolButton:checked:hover {{
                border-radius: 18px;
                background-color: {checked_bg};
            }}
        """)
    
    def get_pen_color_char(self, r, g, b):
        if (r, g, b) == (255, 0, 0):
            return "紅"
        if (r, g, b) == (0, 255, 0):
            return "綠"
        if (r, g, b) == (0, 0, 255):
            return "藍"
        if (r, g, b) == (255, 255, 0):
            return "黃"
        if (r, g, b) == (0, 0, 0):
            return "黑"
        if (r, g, b) == (255, 255, 255):
            return "白"
        if (r, g, b) == (255, 165, 0):
            return "橙"
        if (r, g, b) == (128, 0, 128):
            return "紫"
        if (r, g, b) == (255, 0, 255):
            return "紫"
        if (r, g, b) == (0, 255, 255):
            return "青"
        return "筆"
    
    def update_pen_display(self, color):
        self.current_pen_color = color
        r = color & 0xFF
        g = (color >> 8) & 0xFF
        b = (color >> 16) & 0xFF
        char = self.get_pen_color_char(r, g, b)
        self.btn_pen.setIcon(QIcon())
        self.btn_pen.setText(char)
        font = self.btn_pen.font()
        font.setFamily("Meiryo UI")
        font.setPixelSize(16)
        font.setBold(True)
        self.btn_pen.setFont(font)
        self.style_tool_btn(self.btn_pen, self.current_theme)
    
    def on_tool_btn_toggled(self, checked):
        if not checked:
            return
        btn = self.sender()
        self.update_indicator_geometry(btn)
        self.update_pen_icon_state()

    def update_pen_icon_state(self):
        if self.btn_pen.isChecked() and self.current_pen_color is not None:
            self.update_pen_display(self.current_pen_color)
        else:
            self.btn_pen.setText("")
            self.btn_pen.setIcon(get_icon("Pen.svg", self.current_theme))
            self.style_tool_btn(self.btn_pen, self.current_theme)
    
    def update_indicator_for_current(self):
        for btn in (self.btn_arrow, self.btn_pen, self.btn_eraser):
            if btn.isChecked():
                self.update_indicator_geometry(btn)
                return
        if self.indicator:
            self.indicator.hide()
    
    def update_indicator_geometry(self, btn):
        if not self.indicator or not btn or not btn.isVisible():
            return
        indicator_width = int(btn.width() * 0.6)
        if indicator_width <= 0:
            indicator_width = btn.width()
        h = self.indicator.height()
        x = btn.x() + (btn.width() - indicator_width) // 2
        y = self.container.height() - h - 4
        self.indicator.setGeometry(x, y, indicator_width, h)
        self.indicator.show()
    
    def style_action_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(212, 165, 165, 0.2)"
            pressed_bg = "rgba(212, 165, 165, 0.3)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(179, 143, 143, 0.3)"
            pressed_bg = "rgba(179, 143, 143, 0.4)"
            text_color = "white"
            
        if btn == self.btn_exit:
             btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 18px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    border-radius: 18px;
                    background-color: rgba(255, 50, 50, 0.3);
                }}
                TransparentToolButton:pressed {{
                    border-radius: 18px;
                    background-color: rgba(255, 50, 50, 0.5);
                }}
            """)
        else:
            btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 18px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    border-radius: 18px;
                    background-color: {hover_bg};
                }}
                TransparentToolButton:pressed {{
                    border-radius: 18px;
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


class ClockWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = Theme.DARK
        self.countdown_finished = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.container.setMinimumHeight(40)
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(16, 8, 16, 8)
        inner_layout.setSpacing(6)
        
        self.lbl_time = QLabel()
        self.lbl_time.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_extra = QLabel()
        self.lbl_extra.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_extra.setObjectName("Extra")
        self.lbl_extra.hide()
        inner_layout.addWidget(self.lbl_time)
        inner_layout.addWidget(self.lbl_extra)
        
        layout.addWidget(self.container)
        self.setLayout(layout)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        
        self.update_time()
        self.set_theme(Theme.AUTO)
        
    def update_time(self):
        now = QDateTime.currentDateTime()
        hour = now.time().hour()
        minute = now.time().minute()
        period = "上午" if hour < 12 else "下午"
        time_str = f"{period} {hour:02d}:{minute:02d}"
        self.lbl_time.setText(time_str)
        
    def format_time(self, seconds):
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def update_timer_state(self, up_seconds, up_running, down_remaining, down_running):
        if up_running or down_running:
            self.countdown_finished = False
        elif self.countdown_finished:
            self.lbl_extra.setText("倒计时已结束")
            self.lbl_extra.setVisible(True)
            return
        parts = []
        if up_running:
            parts.append(f"↑ {self.format_time(up_seconds)}")
        if down_running:
            parts.append(f"↓ {self.format_time(down_remaining)}")
        text = "  ".join(parts)
        self.lbl_extra.setText(text)
        self.lbl_extra.setVisible(bool(text))
    
    def show_countdown_finished(self):
        self.countdown_finished = True
        self.lbl_extra.setText("倒计时已结束")
        self.lbl_extra.setVisible(True)
        
    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        # Check for window effect
        effect = self.property("windowEffect")
        is_transparent = effect in ["Mica", "Acrylic"]
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)" if not is_transparent else "rgba(255, 255, 255, 10)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            bg_color = "rgba(30, 30, 30, 240)" if not is_transparent else "rgba(30, 30, 30, 10)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 20px;
            }}
            QLabel {{
                font-size: 16px;
                font-weight: bold;
                color: {text_color};
            }}
            QLabel#Extra {{
                font-size: 11px;
                font-weight: normal;
            }}
        """)


class TimerWindow(QWidget):
    timer_state_changed = pyqtSignal(int, bool, int, bool)
    countdown_finished = pyqtSignal()
    timer_reset = pyqtSignal()
    pre_reminder_triggered = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(460, 360)
        
        self.up_seconds = 0
        self.up_running = False
        self.down_total_seconds = 0
        self.down_remaining = 0
        self.down_running = False
        self.strong_reminder_mode = False
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
        self.set_theme(Theme.DARK)
        
        self.pivot.setCurrentItem("up")
        self.emit_state()

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
        self.up_start_btn.setFixedWidth(130)
        self.up_reset_btn = PushButton("重置", self.up_page)
        self.up_reset_btn.setFixedWidth(130)
        
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
        self.back_btn.setFixedWidth(150)
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
        # Stop shaking if it was loop
        if hasattr(self, 'shake_anim'):
            self.shake_anim.stop()
        # Signals for controller to stop sound/mask
        self.timer_reset.emit()

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
        self.down_min_spin.setFixedWidth(140)
        self.down_min_spin.setValue(5) # Default 5 mins
        
        self.down_sec_spin = SpinBox()
        self.down_sec_spin.setRange(0, 59)
        self.down_sec_spin.setSuffix(" 秒")
        self.down_sec_spin.setFixedWidth(140)
        
        input_layout.addStretch()
        input_layout.addWidget(self.down_min_spin)
        input_layout.addWidget(self.down_sec_spin)
        input_layout.addStretch()
        
        self.down_label = QLabel("00:00")
        self.down_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.down_label.hide()
        
        # Strong reminder settings
        self.strong_reminder_switch = SwitchButton(self.down_page)
        self.strong_reminder_switch.setOnText("强提醒已开启")
        self.strong_reminder_switch.setOffText("强提醒已关闭")
        self.strong_reminder_switch.checkedChanged.connect(self.on_strong_reminder_toggled)

        # Voice settings container (hidden by default)
        self.voice_settings_container = QWidget()
        voice_layout = QVBoxLayout(self.voice_settings_container)
        voice_layout.setContentsMargins(10, 5, 10, 5)
        voice_layout.setSpacing(10)
        self.voice_settings_container.hide()

        # Separator line
        self.separator_line = QFrame()
        self.separator_line.setFrameShape(QFrame.Shape.HLine)
        self.separator_line.setFrameShadow(QFrame.Shadow.Sunken)
        self.separator_line.setStyleSheet("background-color: rgba(255, 255, 255, 0.1); max-height: 1px;")
        voice_layout.addWidget(self.separator_line)

        # Pre-reminder settings
        pre_layout = QHBoxLayout()
        pre_layout.setContentsMargins(0,0,0,0)
        pre_label = BodyLabel("剩余播报:", self)
        
        self.pre_rem_min = SpinBox()
        self.pre_rem_min.setRange(0, 59)
        self.pre_rem_min.setSuffix("分")
        self.pre_rem_min.setFixedWidth(130)
        
        self.pre_rem_sec = SpinBox()
        self.pre_rem_sec.setRange(0, 59)
        self.pre_rem_sec.setSuffix("秒")
        self.pre_rem_sec.setFixedWidth(130)
        
        pre_layout.addWidget(pre_label)
        pre_layout.addStretch()
        pre_layout.addWidget(self.pre_rem_min)
        pre_layout.addWidget(self.pre_rem_sec)

        # Post-reminder settings
        post_layout = QHBoxLayout()
        post_layout.setContentsMargins(0,0,0,0)
        post_label = BodyLabel("结束播报:", self)
        self.post_rem_switch = SwitchButton()
        self.post_rem_switch.setOnText("开启")
        self.post_rem_switch.setOffText("关闭")
        self.post_rem_switch.setChecked(True)
        
        post_layout.addWidget(post_label)
        post_layout.addStretch()
        post_layout.addWidget(self.post_rem_switch)
        
        voice_layout.addLayout(pre_layout)
        voice_layout.addLayout(post_layout)
        
        layout.addStretch()
        layout.addWidget(self.input_widget)
        layout.addWidget(self.down_label)
        layout.addSpacing(10)
        layout.addWidget(self.strong_reminder_switch, 0, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.voice_settings_container)
        layout.addStretch()
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        self.down_start_btn = PrimaryPushButton("开始", self.down_page)
        self.down_start_btn.setFixedWidth(130)
        self.down_reset_btn = PushButton("重置", self.down_page)
        self.down_reset_btn.setFixedWidth(130)
        
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
        
        # Check for window effect
        effect = self.property("windowEffect")
        is_transparent = effect in ["Mica", "Acrylic"]
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(243, 243, 243, 0.95)" if not is_transparent else "rgba(243, 243, 243, 0.1)"
            border_color = "rgba(0, 0, 0, 0.05)"
            text_color = "#333333"
            self.title_label.setTextColor("#333333", "#333333")
        else:
            bg_color = "rgba(32, 32, 32, 0.95)" if not is_transparent else "rgba(32, 32, 32, 0.1)"
            border_color = "rgba(255, 255, 255, 0.08)"
            text_color = "white"
            self.title_label.setTextColor("white", "white")
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        font_style = f"font-size: 56px; font-weight: bold; color: {text_color};"
        self.up_label.setStyleSheet(font_style)
        self.down_label.setStyleSheet(font_style)
        
        completed_style = f"font-size: 24px; font-weight: bold; color: {text_color};"
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
        
        # In strong reminder mode, loop indefinitely
        loop_count = -1 if self.strong_reminder_mode else 2
        
        # Create one cycle of shake
        cycle_anim = QSequentialAnimationGroup(self)
        
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
        
        cycle_anim.addAnimation(anim1)
        cycle_anim.addAnimation(anim2)
        cycle_anim.addAnimation(anim3)
        
        if loop_count == -1:
            self.shake_anim.setLoopCount(0) # 0 means infinite in some contexts, but let's check docs or use loop
            # QSequentialAnimationGroup doesn't have setLoopCount directly in all Qt versions, 
            # but we can re-start on finished if needed.
            # Actually QAbstractAnimation has setLoopCount.
            self.shake_anim.addAnimation(cycle_anim)
            self.shake_anim.setLoopCount(-1) 
        else:
             # Add multiple cycles
             for _ in range(loop_count):
                 self.shake_anim.addAnimation(cycle_anim)
            
        self.shake_anim.start()

    def on_strong_reminder_toggled(self, checked):
        if checked:
            w = MessageDialog(
                "开启强提醒",
                "强提醒模式下，倒计时结束时将：\n\n1. 强制弹出计时窗口\n2. 播放高强度警报音\n3. 窗口持续抖动\n4. 显示全屏遮罩\n\n只有点击“返回”按钮才能停止提醒。",
                self.window()
            )
            w.yesButton.setText("开启")
            w.cancelButton.setText("取消")
            if w.exec():
                self.strong_reminder_mode = True
                self.voice_settings_container.show()
                # Expand window slightly if needed, or let layout handle it
            else:
                self.strong_reminder_switch.setChecked(False)
                self.strong_reminder_mode = False
                self.voice_settings_container.hide()
        else:
            self.strong_reminder_mode = False
            self.voice_settings_container.hide()

    def init_timers(self):
        self.up_timer = QTimer(self)
        self.up_timer.setInterval(1000)
        self.up_timer.timeout.connect(self.update_up)
        self.down_timer = QTimer(self)
        self.down_timer.setInterval(1000)
        self.down_timer.timeout.connect(self.update_down)


    def emit_state(self):
        self.timer_state_changed.emit(self.up_seconds, self.up_running, self.down_remaining, self.down_running)

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
        self.emit_state()

    def reset_up(self):
        self.up_timer.stop()
        self.up_running = False
        self.up_seconds = 0
        self.up_start_btn.setText("开始")
        self.up_label.setText(self.format_time(self.up_seconds))
        self.emit_state()
        self.timer_reset.emit()

    def update_up(self):
        self.up_seconds += 1
        self.up_label.setText(self.format_time(self.up_seconds))
        self.emit_state()

    def toggle_down(self):
        if not self.down_running:
            if self.down_remaining <= 0 and self.down_total_seconds == 0:
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
            self.strong_reminder_switch.setEnabled(False)
        else:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")
            self.strong_reminder_switch.setEnabled(True)
        self.emit_state()

    def reset_down(self):
        self.down_timer.stop()
        self.down_running = False
        self.down_remaining = 0
        self.down_total_seconds = 0
        self.down_start_btn.setText("开始")
        self.strong_reminder_switch.setEnabled(True)
        
        self.down_label.hide()
        self.input_widget.show()
        self.emit_state()
        self.timer_reset.emit()

    def update_down(self):
        if self.down_remaining > 0:
            self.down_remaining -= 1
            self.down_label.setText(self.format_time(self.down_remaining))
            self.emit_state()
            
            # Pre-reminder logic
            if self.strong_reminder_mode:
                rem_min = self.pre_rem_min.value()
                rem_sec = self.pre_rem_sec.value()
                total_rem = rem_min * 60 + rem_sec
                if total_rem > 0 and self.down_remaining == total_rem:
                    msg = f"倒计时剩余{rem_min}分{rem_sec}秒" if rem_min > 0 else f"倒计时剩余{rem_sec}秒"
                    # If seconds is 0, just say minutes
                    if rem_sec == 0 and rem_min > 0:
                        msg = f"倒计时剩余{rem_min}分钟"
                    self.pre_reminder_triggered.emit(msg)
        
        if self.down_remaining <= 0 and self.down_running:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")
            
            self.stack.setCurrentWidget(self.completed_page)
            self.shake_window()
            self.countdown_finished.emit()
            self.emit_state()
