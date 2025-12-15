import sys
import os
from PyQt6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QButtonGroup, 
                             QLabel, QFrame, QPushButton, QGridLayout)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QEvent, QRect, QPropertyAnimation, QEasingCurve, QAbstractAnimation, QTimer
from PyQt6.QtGui import QIcon, QPixmap
from qfluentwidgets import (TransparentToolButton, ToolButton, SpinBox,
                            PrimaryPushButton, PushButton, TabWidget,
                            ToolTipFilter, ToolTipPosition, Flyout, FlyoutAnimationType)
from qfluentwidgets.components.material import AcrylicFlyout

def icon_path(name):
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # 返回上级目录的icons文件夹路径
    return os.path.join(os.path.dirname(base_dir), "icons", name)


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
        
        btn = QPushButton("清除当前页笔迹")  # 使用标准按钮而非qfluentwidgets.PushButton
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
        
        from PyQt6.QtWidgets import QScrollArea, QGridLayout
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
        import hashlib
        import os
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
        
        # 工具栏
        self.setup_toolbar()
        
    def setup_toolbar(self):
        """设置工具栏"""
        toolbar_width = 40
        toolbar_height = 200
        margin = 20
        
        # 创建工具栏容器
        self.toolbar = QWidget(self)
        self.toolbar.setGeometry(margin, margin, toolbar_width, toolbar_height)
        self.toolbar.setStyleSheet("""
            QWidget {
                background-color: rgba(30, 30, 30, 240);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 12px;
            }
        """)
        
        # 工具栏布局
        from PyQt6.QtWidgets import QVBoxLayout
        layout = QVBoxLayout(self.toolbar)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        
        # 笔工具按钮
        self.btn_pen = QPushButton("笔")
        self.btn_pen.setFixedSize(30, 30)
        self.btn_pen.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 0.1);
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:checked {
                background-color: #00cc7a;
            }
        """)
        self.btn_pen.setCheckable(True)
        self.btn_pen.clicked.connect(self.set_pen_mode)
        
        # 橡皮擦按钮
        self.btn_eraser = QPushButton("擦")
        self.btn_eraser.setFixedSize(30, 30)
        self.btn_eraser.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 0.1);
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:checked {
                background-color: #00cc7a;
            }
        """)
        self.btn_eraser.setCheckable(True)
        self.btn_eraser.clicked.connect(self.set_eraser_mode)
        
        # 清除按钮
        self.btn_clear = QPushButton("清")
        self.btn_clear.setFixedSize(30, 30)
        self.btn_clear.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 50, 50, 0.3);
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: rgba(255, 50, 50, 0.5);
            }
        """)
        self.btn_clear.clicked.connect(self.clear_all)
        
        # 关闭按钮
        self.btn_close = QPushButton("X")
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 50, 50, 0.3);
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: rgba(255, 50, 50, 0.5);
            }
        """)
        self.btn_close.clicked.connect(self.close)
        
        # 颜色选择按钮
        self.btn_color_red = QPushButton()
        self.btn_color_red.setFixedSize(30, 30)
        self.btn_color_red.setStyleSheet("background-color: red; border-radius: 4px;")
        self.btn_color_red.clicked.connect(lambda: self.set_pen_color(Qt.GlobalColor.red))
        
        self.btn_color_blue = QPushButton()
        self.btn_color_blue.setFixedSize(30, 30)
        self.btn_color_blue.setStyleSheet("background-color: blue; border-radius: 4px;")
        self.btn_color_blue.clicked.connect(lambda: self.set_pen_color(Qt.GlobalColor.blue))
        
        self.btn_color_green = QPushButton()
        self.btn_color_green.setFixedSize(30, 30)
        self.btn_color_green.setStyleSheet("background-color: green; border-radius: 4px;")
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
        
        # Close button (context aware)
        self.btn_close = QPushButton("X", self)
        self.btn_close.setFixedSize(30, 30)
        self.btn_close.setStyleSheet("background-color: red; color: white; border-radius: 0px; font-weight: bold;")
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
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
    prev_clicked = pyqtSignal()
    next_clicked = pyqtSignal()
    
    def __init__(self, parent=None, is_right=False):
        super().__init__(parent)
        self.ppt_app = None 
        self.is_right = is_right
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
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
        self.btn_prev.clicked.connect(self.prev_clicked.emit)
        self.style_nav_btn(self.btn_prev)
        
        self.btn_next = TransparentToolButton(parent=self)
        self.btn_next.setIcon(QIcon(icon_path("Next.svg")))
        self.btn_next.setFixedSize(36, 36) 
        self.btn_next.setIconSize(QSize(18, 18))
        self.btn_next.setToolTip("下一页")
        self.btn_next.installEventFilter(ToolTipFilter(self.btn_next, 1000, ToolTipPosition.TOP))
        self.btn_next.clicked.connect(self.next_clicked.emit)
        self.style_nav_btn(self.btn_next)

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
        
        layout.addWidget(self.container)
        self.setLayout(layout)

        self.setup_click_feedback(self.btn_prev, QSize(18, 18))
        self.setup_click_feedback(self.btn_next, QSize(18, 18))

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
    request_timer = pyqtSignal()
    
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
        self.btn_timer = self.create_action_btn("计时器", QIcon(icon_path("timer.svg")))
        self.btn_timer.clicked.connect(self.request_timer.emit)

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
        container_layout.addWidget(self.btn_timer)

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

        self.setup_click_feedback(self.btn_arrow, QSize(20, 20))
        self.setup_click_feedback(self.btn_pen, QSize(20, 20))
        self.setup_click_feedback(self.btn_eraser, QSize(20, 20))
        self.setup_click_feedback(self.btn_clear, QSize(20, 20))
        self.setup_click_feedback(self.btn_spotlight, QSize(20, 20))
        self.setup_click_feedback(self.btn_timer, QSize(20, 20))
        self.setup_click_feedback(self.btn_exit, QSize(20, 20))

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

    def show_pen_settings(self):#
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
        self.btn_timer.setToolTip("计时器")
        self.btn_exit.setToolTip("结束放映")
        self.style_tool_btn(self.btn_arrow)
        self.style_tool_btn(self.btn_pen)
        self.style_tool_btn(self.btn_eraser)
        self.style_action_btn(self.btn_clear)
        self.style_action_btn(self.btn_spotlight)
        self.style_action_btn(self.btn_timer)
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
        self.setWindowTitle("计时器")
        flags = Qt.WindowType.Window | Qt.WindowType.WindowTitleHint | Qt.WindowType.CustomizeWindowHint | Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.resize(320, 220)
        self.up_seconds = 0
        self.up_running = False
        self.down_total_seconds = 0
        self.down_remaining = 0
        self.down_running = False
        self.sound_effect = None
        self.init_ui()
        self.init_timers()
        self.init_sound()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        self.up_page = QWidget()
        self.down_page = QWidget()
        self.setup_up_page()
        self.setup_down_page()
        self.tab_widget = TabWidget(self)
        self.tab_widget.addPage(self.up_page, "正计时")
        self.tab_widget.addPage(self.down_page, "倒计时")
        layout.addWidget(self.tab_widget)

    def setup_up_page(self):
        layout = QVBoxLayout(self.up_page)
        layout.setContentsMargins(0, 16, 0, 0)
        layout.setSpacing(16)
        self.up_label = QLabel("00:00", self.up_page)
        self.up_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.up_label.setStyleSheet("font-size: 32px; color: white;")
        layout.addWidget(self.up_label)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        self.up_start_btn = PrimaryPushButton(self.up_page)
        self.up_start_btn.setText("开始")
        self.up_reset_btn = PushButton(self.up_page)
        self.up_reset_btn.setText("重置")
        self.up_start_btn.clicked.connect(self.toggle_up)
        self.up_reset_btn.clicked.connect(self.reset_up)
        btn_layout.addStretch()
        btn_layout.addWidget(self.up_start_btn)
        btn_layout.addWidget(self.up_reset_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def setup_down_page(self):
        layout = QVBoxLayout(self.down_page)
        layout.setContentsMargins(0, 16, 0, 0)
        layout.setSpacing(16)
        input_layout = QHBoxLayout()
        input_layout.setSpacing(8)
        self.down_min_spin = SpinBox(self.down_page)
        self.down_min_spin.setRange(0, 999)
        self.down_min_spin.setSuffix(" 分")
        self.down_sec_spin = SpinBox(self.down_page)
        self.down_sec_spin.setRange(0, 59)
        self.down_sec_spin.setSuffix(" 秒")
        input_layout.addStretch()
        input_layout.addWidget(self.down_min_spin)
        input_layout.addWidget(self.down_sec_spin)
        input_layout.addStretch()
        layout.addLayout(input_layout)
        self.down_label = QLabel("00:00", self.down_page)
        self.down_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.down_label.setStyleSheet("font-size: 32px; color: white;")
        layout.addWidget(self.down_label)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        self.down_start_btn = ToolButton(self.down_page)
        self.down_start_btn.setText("开始")
        self.down_reset_btn = ToolButton(self.down_page)
        self.down_reset_btn.setText("重置")
        self.down_start_btn.clicked.connect(self.toggle_down)
        self.down_reset_btn.clicked.connect(self.reset_down)
        btn_layout.addStretch()
        btn_layout.addWidget(self.down_start_btn)
        btn_layout.addWidget(self.down_reset_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def init_timers(self):
        self.up_timer = QTimer(self)
        self.up_timer.setInterval(1000)
        self.up_timer.timeout.connect(self.update_up)
        self.down_timer = QTimer(self)
        self.down_timer.setInterval(1000)
        self.down_timer.timeout.connect(self.update_down)

    def init_sound(self):
        try:
            from PyQt6.QtMultimedia import QSoundEffect
            from PyQt6.QtCore import QUrl
        except Exception:
            self.sound_effect = None
            return
        self.sound_effect = QSoundEffect(self)
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base_dir, "timer_ring.ogg")
        self.sound_effect.setSource(QUrl.fromLocalFile(path))
        self.sound_effect.setLoopCount(1)
        self.sound_effect.setVolume(0.9)

    def play_ring(self):
        if not self.sound_effect:
            return
        self.sound_effect.stop()
        self.sound_effect.play()

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
            minutes = self.down_min_spin.value()
            seconds = self.down_sec_spin.value()
            total = minutes * 60 + seconds
            if total <= 0:
                return
            self.down_total_seconds = total
            self.down_remaining = total
            self.down_label.setText(self.format_time(self.down_remaining))
            self.down_min_spin.setEnabled(False)
            self.down_sec_spin.setEnabled(False)
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
        self.down_start_btn.setText("开始")
        self.down_label.setText("00:00")
        self.down_min_spin.setEnabled(True)
        self.down_sec_spin.setEnabled(True)

    def update_down(self):
        if self.down_remaining > 0:
            self.down_remaining -= 1
            self.down_label.setText(self.format_time(self.down_remaining))
        if self.down_remaining <= 0 and self.down_running:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")
            self.down_min_spin.setEnabled(True)
            self.down_sec_spin.setEnabled(True)
            self.play_ring()
