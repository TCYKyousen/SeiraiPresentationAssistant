from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PyQt6.QtCore import Qt, QRectF, QPoint, QPropertyAnimation, QEasingCurve, pyqtProperty
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QPainterPath
from qfluentwidgets import OptionsSettingCard, Theme, isDarkTheme, themeColor

class SchematicOptionButton(QWidget):
    """ Individual schematic option button with hover/press animations """
    def __init__(self, value, text, schematic_type, parent=None):
        super().__init__(parent)
        self.value = value
        self.text_label = text
        self.schematic_type = schematic_type
        self.setFixedSize(140, 110)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self._is_checked = False
        self._hover_progress = 0.0
        self._click_scale = 1.0
        
        # Animations
        self.hover_anim = QPropertyAnimation(self, b"hoverProgress", self)
        self.hover_anim.setDuration(150)
        self.scale_anim = QPropertyAnimation(self, b"clickScale", self)
        self.scale_anim.setDuration(100)
        
    def get_hover_progress(self):
        return self._hover_progress
        
    def set_hover_progress(self, v):
        self._hover_progress = v
        self.update()
        
    hoverProgress = pyqtProperty(float, get_hover_progress, set_hover_progress)
    
    def get_click_scale(self):
        return self._click_scale
        
    def set_click_scale(self, v):
        self._click_scale = v
        self.update()
        
    clickScale = pyqtProperty(float, get_click_scale, set_click_scale)
    
    def setChecked(self, checked):
        if self._is_checked != checked:
            self._is_checked = checked
            self.update()
            
    def enterEvent(self, e):
        self.hover_anim.setStartValue(self._hover_progress)
        self.hover_anim.setEndValue(1.0)
        self.hover_anim.start()
        
    def leaveEvent(self, e):
        self.hover_anim.setStartValue(self._hover_progress)
        self.hover_anim.setEndValue(0.0)
        self.hover_anim.start()
        
    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.scale_anim.setStartValue(self._click_scale)
            self.scale_anim.setEndValue(0.95)
            self.scale_anim.setEasingCurve(QEasingCurve.Type.OutQuad)
            self.scale_anim.start()
        super().mousePressEvent(e)
            
    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.scale_anim.setStartValue(self._click_scale)
            self.scale_anim.setEndValue(1.0)
            self.scale_anim.setEasingCurve(QEasingCurve.Type.OutElastic)
            self.scale_anim.start()
            # Trigger click logic in parent or via signal
            if self.parent():
                self.parent().on_option_clicked(self.value)
        super().mouseReleaseEvent(e)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Scale transform for click effect
        painter.translate(self.width() / 2, self.height() / 2)
        painter.scale(self._click_scale, self._click_scale)
        painter.translate(-self.width() / 2, -self.height() / 2)
        
        rect = QRectF(5, 5, self.width() - 10, self.height() - 30) # Reserve bottom for text
        
        is_dark = isDarkTheme()
        accent = themeColor()
        
        # 1. Background & Border
        bg_color = QColor(0, 0, 0, 0)
        border_color = QColor(255, 255, 255, 20) if is_dark else QColor(0, 0, 0, 10)
        border_width = 1.0
        
        if self._is_checked:
            border_color = accent
            border_width = 2.0
            bg_color = QColor(accent)
            bg_color.setAlpha(20)
        elif self._hover_progress > 0.01:
            # Interpolate hover background
            alpha = int(10 * self._hover_progress)
            bg_color = QColor(255, 255, 255, alpha) if is_dark else QColor(0, 0, 0, alpha)
            border_color = QColor(255, 255, 255, 50) if is_dark else QColor(0, 0, 0, 30)
            
        path = QPainterPath()
        path.addRoundedRect(rect, 8, 8)
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg_color)
        painter.drawPath(path)
        
        # Draw border
        pen = QPen(border_color, border_width)
        # Avoid clipping border
        if border_width > 1:
            pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawPath(path)
        
        # 2. Schematic Content
        content_rect = rect.adjusted(4, 4, -4, -4)
        self.draw_content(painter, content_rect)
        
        # 3. Text Label
        painter.setPen(QColor(255, 255, 255) if is_dark else QColor(0, 0, 0))
        font = painter.font()
        font.setPixelSize(13)
        if self._is_checked:
            font.setBold(True)
        painter.setFont(font)
        text_rect = QRectF(0, rect.bottom() + 5, self.width(), 20)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, self.text_label)

    def draw_content(self, painter, rect):
        if self.schematic_type == "theme":
            self.draw_theme_content(painter, rect)
        elif self.schematic_type == "nav_pos":
            self.draw_nav_content(painter, rect)
            
    def draw_theme_content(self, painter, rect):
        if self.value == "Light":
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(243, 243, 243))
            painter.drawRoundedRect(rect, 4, 4)
            # Lines
            painter.setBrush(QColor(0, 0, 0, 20))
            painter.drawRoundedRect(rect.adjusted(5, 5, -20, -rect.height()+10), 2, 2)
            painter.drawRoundedRect(rect.adjusted(5, 15, -10, -rect.height()+20), 2, 2)
            
        elif self.value == "Dark":
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(32, 32, 32))
            painter.drawRoundedRect(rect, 4, 4)
            # Lines
            painter.setBrush(QColor(255, 255, 255, 20))
            painter.drawRoundedRect(rect.adjusted(5, 5, -20, -rect.height()+10), 2, 2)
            painter.drawRoundedRect(rect.adjusted(5, 15, -10, -rect.height()+20), 2, 2)
            
        elif self.value == "Auto":
            path = QPainterPath()
            path.addRoundedRect(rect, 4, 4)
            painter.save()
            painter.setClipPath(path)
            
            # Light half
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(243, 243, 243))
            painter.drawRect(rect) 
            
            # Dark half
            path_tri = QPainterPath()
            path_tri.moveTo(rect.topRight())
            path_tri.lineTo(rect.bottomRight())
            path_tri.lineTo(rect.bottomLeft())
            path_tri.closeSubpath()
            painter.setBrush(QColor(32, 32, 32))
            painter.drawPath(path_tri)
            
            painter.restore()

    def draw_nav_content(self, painter, rect):
        is_dark = isDarkTheme()
        screen_color = QColor(60, 60, 60) if is_dark else QColor(255, 255, 255)
        btn_color = QColor(0, 204, 122)
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(screen_color)
        painter.drawRoundedRect(rect, 4, 4)
        
        btn_size = 8
        margin = 6
        painter.setBrush(btn_color)
        
        if self.value == "BottomSides":
            painter.drawEllipse(QRectF(rect.left() + margin, rect.bottom() - margin - btn_size, btn_size, btn_size))
            painter.drawEllipse(QRectF(rect.right() - margin - btn_size, rect.bottom() - margin - btn_size, btn_size, btn_size))
        elif self.value == "MiddleSides":
            painter.drawEllipse(QRectF(rect.left() + margin, rect.center().y() - btn_size/2, btn_size, btn_size))
            painter.drawEllipse(QRectF(rect.right() - margin - btn_size, rect.center().y() - btn_size/2, btn_size, btn_size))


class SchematicWidget(QWidget):
    """ Container for schematic option buttons """
    def __init__(self, schematic_type="theme", configItem=None, parent=None):
        super().__init__(parent)
        self.schematic_type = schematic_type
        self.configItem = configItem
        self.setFixedHeight(140)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(15)
        self.layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.buttons = []
        self.setup_options()
        
    def setup_options(self):
        if self.schematic_type == "theme":
            options = [("Light", "浅色"), ("Dark", "深色"), ("Auto", "跟随系统")]
        elif self.schematic_type == "nav_pos":
            options = [("BottomSides", "底部两端"), ("MiddleSides", "中部两侧")]
        else:
            options = []
            
        for val, text in options:
            btn = SchematicOptionButton(val, text, self.schematic_type, self)
            self.layout.addWidget(btn)
            self.buttons.append(btn)
            
        self.update_state()
        
    def update_state(self):
        if not self.configItem:
            return
            
        current_value = self.configItem.value
        # Handle Enum conversion
        if hasattr(current_value, "value"):
            current_value = current_value.value
        if isinstance(current_value, Theme):
            if current_value == Theme.LIGHT: current_value = "Light"
            elif current_value == Theme.DARK: current_value = "Dark"
            elif current_value == Theme.AUTO: current_value = "Auto"
            
        for btn in self.buttons:
            btn.setChecked(btn.value == current_value)
            
    def on_option_clicked(self, val):
        if not self.configItem:
            return
            
        # Determine the target value type based on current config value type
        current = self.configItem.value
        final_val = val
        
        if self.schematic_type == "theme":
            if isinstance(current, Theme):
                if val == "Light": final_val = Theme.LIGHT
                elif val == "Dark": final_val = Theme.DARK
                elif val == "Auto": final_val = Theme.AUTO
            else:
                # Assume string if not Enum
                final_val = val
            
        if self.configItem.value != final_val:
            self.configItem.value = final_val
            self.update_state()

    def update(self):
        self.update_state()
        super().update()


class SchematicOptionsSettingCard(OptionsSettingCard):
    """ OptionsSettingCard with a schematic illustration """
    def __init__(self, configItem, icon, title, content, texts, schematic_type="theme", parent=None):
        super().__init__(configItem, icon, title, content, texts, parent)
        # Pass configItem to widget
        self.schematic = SchematicWidget(schematic_type, configItem, self)
        
        # Connect signal to update schematic when value changes externally
        configItem.valueChanged.connect(lambda v: self.schematic.update())
        
        # Insert schematic at the top of the expanded view layout
        self.viewLayout.insertWidget(0, self.schematic)
        self.viewLayout.insertSpacing(1, 10)
        
        # Hide original radio buttons
        for i in range(self.viewLayout.count()):
            item = self.viewLayout.itemAt(i)
            widget = item.widget()
            if widget and widget != self.schematic:
                widget.hide()
