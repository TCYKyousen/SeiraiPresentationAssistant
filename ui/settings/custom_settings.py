from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel
from PyQt6.QtCore import Qt, QRectF, QPoint, QPropertyAnimation, QEasingCurve, pyqtProperty, QParallelAnimationGroup, QVariantAnimation
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QPainterPath
from qfluentwidgets import OptionsSettingCard, Theme, isDarkTheme, themeColor, SettingCard, PushButton, PrimaryPushButton, SpinBox, SwitchButton
from PyQt6.QtWidgets import QGraphicsBlurEffect

class SchematicOptionButton(QWidget):
    """ Individual schematic option button with hover/press animations """
    def __init__(self, value, text, schematic_type, parent=None):
        super().__init__(parent)
        self.value = value
        self.schematic_type = schematic_type
        self.setFixedSize(140, 110)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.text_label = QLabel(text, self)
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setStyleSheet("font-size: 14px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        self.set_theme()
        
        self._is_checked = False
        self._hover_progress = 0.0
        self._click_scale = 1.0
        
        # 建立动画组
        self.click_group = QParallelAnimationGroup(self)
        
        # Animations
        self.hover_anim = QPropertyAnimation(self, b"hoverProgress", self)
        self.hover_anim.setDuration(100)
        self.hover_anim.setEasingCurve(QEasingCurve.Type.OutQuart)
        
        self.scale_anim = QPropertyAnimation(self, b"clickScale", self.click_group)
        self.scale_anim.setDuration(60)
        
        # 动态模糊
        self.blur_effect = QGraphicsBlurEffect(self)
        self.blur_effect.setBlurRadius(0)
        self.setGraphicsEffect(self.blur_effect)
        
        self.blur_anim = QVariantAnimation(self.click_group)
        self.blur_anim.setDuration(60)
        self.blur_anim.setStartValue(0.0)
        self.blur_anim.setKeyValueAt(0.5, 2.0)
        self.blur_anim.setEndValue(0.0)
        self.blur_anim.valueChanged.connect(lambda v: self.blur_effect.setBlurRadius(v))
        
        self.click_group.addAnimation(self.scale_anim)
        self.click_group.addAnimation(self.blur_anim)
        
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
    
    def set_theme(self, theme=None):
        if theme is None or theme == Theme.AUTO:
            theme = Theme.DARK if isDarkTheme() else Theme.LIGHT
            
        text_color = "white" if theme == Theme.DARK else "black"
        style = self.text_label.styleSheet()
        import re
        style = re.sub(r'color\s*:\s*[^;]+;', '', style)
        self.text_label.setStyleSheet(style + f" color: {text_color};")

    def setChecked(self, checked):
        if self._is_checked != checked:
            self._is_checked = checked
            # Update label style
            font = self.text_label.font()
            font.setBold(checked)
            self.text_label.setFont(font)
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
            if self.click_group.state() == QParallelAnimationGroup.State.Running:
                self.click_group.stop()
            self.scale_anim.setStartValue(self._click_scale)
            self.scale_anim.setEndValue(0.95)
            self.scale_anim.setEasingCurve(QEasingCurve.Type.OutQuad)
            self.click_group.start()
        super().mousePressEvent(e)
            
    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            if self.click_group.state() == QParallelAnimationGroup.State.Running:
                self.click_group.stop()
            self.scale_anim.setStartValue(self._click_scale)
            self.scale_anim.setEndValue(1.0)
            self.scale_anim.setEasingCurve(QEasingCurve.Type.OutBack)
            self.click_group.start()
            # Trigger click logic in parent or via signal
            if self.parent():
                self.parent().on_option_clicked(self.value)
        super().mouseReleaseEvent(e)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.text_label.setGeometry(0, self.height() - 25, self.width(), 20)

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

    def draw_content(self, painter, rect):
        if self.schematic_type == "theme":
            self.draw_theme_content(painter, rect)
        elif self.schematic_type == "nav_pos":
            self.draw_nav_content(painter, rect)
        elif self.schematic_type == "toolbar_pos":
            self.draw_toolbar_content(painter, rect)
            
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

    def draw_toolbar_content(self, painter, rect):
        is_dark = isDarkTheme()
        screen_color = QColor(60, 60, 60) if is_dark else QColor(255, 255, 255)
        bar_color = QColor(0, 204, 122)
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(screen_color)
        painter.drawRoundedRect(rect, 4, 4)
        
        margin_x = 8
        margin_y = 8
        bar_thickness = 6
        
        v = "Bottom"
        h = "Center"
        if "Top" in self.value:
            v = "Top"
        elif "Middle" in self.value:
            v = "Middle"
        elif "Bottom" in self.value:
            v = "Bottom"
        if "Left" in self.value:
            h = "Left"
        elif "Right" in self.value:
            h = "Right"
        elif "Center" in self.value:
            h = "Center"
        
        is_vertical = v == "Middle" and h in ("Left", "Right")
        
        painter.setBrush(bar_color)
        
        if is_vertical:
            bar_w = bar_thickness
            bar_h = rect.height() - 2 * margin_y
            y = rect.top() + margin_y
            if h == "Left":
                x = rect.left() + margin_x
            else:
                x = rect.right() - margin_x - bar_w
        else:
            bar_h = bar_thickness
            bar_w = rect.width() - 2 * margin_x
            if v == "Top":
                y = rect.top() + margin_y
            elif v == "Middle":
                y = rect.center().y() - bar_h / 2
            else:
                y = rect.bottom() - margin_y - bar_h
            x = rect.left() + (rect.width() - bar_w) / 2
        
        painter.drawRoundedRect(QRectF(x, y, bar_w, bar_h), 3, 3)


class SchematicWidget(QWidget):
    """ Container for schematic option buttons """
    def __init__(self, schematic_type="theme", configItem=None, parent=None):
        super().__init__(parent)
        self.schematic_type = schematic_type
        self.configItem = configItem
        
        # Adjust height based on type
        if schematic_type == "toolbar_pos":
            self.setFixedHeight(260) # Increased height for grid layout
        else:
            self.setFixedHeight(140)
            
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # Use QGridLayout for toolbar_pos, QHBoxLayout for others
        if schematic_type == "toolbar_pos":
            self.layout = QGridLayout(self)
            self.layout.setContentsMargins(0, 0, 0, 0)
            self.layout.setSpacing(15)
            self.layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            self.layout = QHBoxLayout(self)
            self.layout.setContentsMargins(0, 0, 0, 0)
            self.layout.setSpacing(15)
            self.layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.buttons = []
        self.setup_options()
        
    def setup_options(self):
        if self.schematic_type == "theme":
            options = [("Light", "浅色"), ("Dark", "深色"), ("Auto", "跟随系统")]
            for val, text in options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn)
                self.buttons.append(btn)
                
        elif self.schematic_type == "nav_pos":
            options = [("BottomSides", "底部两端"), ("MiddleSides", "中部两侧")]
            for val, text in options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn)
                self.buttons.append(btn)
                
        elif self.schematic_type == "toolbar_pos":
            # Grid layout for toolbar position
            # Row 0: MiddleLeft, Empty, MiddleRight
            # Row 1: BottomLeft, BottomCenter, BottomRight
            
            # (value, text, row, col)
            grid_options = [
                ("MiddleLeft", "中部左侧", 0, 0),
                # ("MiddleCenter", "中部居中", 0, 1), # Not supported/allowed
                ("MiddleRight", "中部右侧", 0, 2),
                ("BottomLeft", "底部左侧", 1, 0),
                ("BottomCenter", "底部居中", 1, 1),
                ("BottomRight", "底部右侧", 1, 2),
            ]
            
            for val, text, row, col in grid_options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn, row, col)
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


class ScreenPaddingPreview(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.top = 20
        self.bottom = 20
        self.left = 20
        self.right = 20

    def set_margins(self, t, b, l, r):
        self.top = max(0, t)
        self.bottom = max(0, b)
        self.left = max(0, l)
        self.right = max(0, r)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(10, 10, -10, -10)
        if rect.width() <= 0 or rect.height() <= 0:
            return
        is_dark = isDarkTheme()
        bg = QColor(20, 20, 20) if is_dark else QColor(245, 245, 245)
        screen_color = QColor(40, 40, 40) if is_dark else QColor(255, 255, 255)
        margin_color = QColor(0, 120, 215, 80)
        painter.fillRect(rect, bg)
        inner = QRectF(rect.adjusted(6, 6, -6, -6))
        painter.setPen(Qt.PenStyle.NoPen)
        from PyQt6.QtWidgets import QApplication
        screen = QApplication.primaryScreen()
        if screen is not None:
            g = screen.geometry()
            sw = float(g.width())
            sh = float(g.height())
        else:
            sw = 1920.0
            sh = 1080.0
        if sw <= 0 or sh <= 0:
            sw = 1920.0
            sh = 1080.0
        screen_aspect = sw / sh
        iw = inner.width()
        ih = inner.height()
        target_w = iw
        target_h = target_w / screen_aspect
        if target_h > ih:
            target_h = ih
            target_w = target_h * screen_aspect
        screen_rect = QRectF(
            inner.center().x() - target_w / 2,
            inner.center().y() - target_h / 2,
            target_w,
            target_h,
        )
        painter.setBrush(screen_color)
        painter.drawRoundedRect(screen_rect, 8, 8)
        mt = float(self.top)
        mb = float(self.bottom)
        ml = float(self.left)
        mr = float(self.right)
        scale_x = screen_rect.width() / sw
        scale_y = screen_rect.height() / sh
        st = mt * scale_y
        sb = mb * scale_y
        sl = ml * scale_x
        sr = mr * scale_x
        st = max(0.0, min(st, screen_rect.height() * 0.49))
        sb = max(0.0, min(sb, screen_rect.height() * 0.49))
        sl = max(0.0, min(sl, screen_rect.width() * 0.49))
        sr = max(0.0, min(sr, screen_rect.width() * 0.49))
        painter.setBrush(margin_color)
        painter.drawRect(QRectF(screen_rect.left(), screen_rect.top(), screen_rect.width(), st))
        painter.drawRect(QRectF(screen_rect.left(), screen_rect.bottom() - sb, screen_rect.width(), sb))
        painter.drawRect(QRectF(screen_rect.left(), screen_rect.top() + st, sl, screen_rect.height() - st - sb))
        painter.drawRect(QRectF(screen_rect.right() - sr, screen_rect.top() + st, sr, screen_rect.height() - st - sb))
        content = QRectF(
            screen_rect.left() + sl,
            screen_rect.top() + st,
            max(0.0, screen_rect.width() - sl - sr),
            max(0.0, screen_rect.height() - st - sb),
        )
        painter.setBrush(QColor(0, 0, 0, 0))
        pen = QPen(QColor(0, 120, 215))
        pen.setStyle(Qt.PenStyle.DashLine)
        painter.setPen(pen)
        painter.drawRoundedRect(content, 6, 6)


class ScreenPaddingSettingCard(SettingCard):
    def __init__(self, icon, title, content, parent=None):
        super().__init__(icon, title, content, parent)
        self.configBtn = PushButton("调整", self)
        self.configBtn.setFixedWidth(120)
        self.configBtn.clicked.connect(self._show_config_overlay)
        self.hBoxLayout.addWidget(self.configBtn, 0, Qt.AlignmentFlag.AlignRight)
        self.hBoxLayout.addSpacing(16)

    def _show_config_overlay(self):
        from PyQt6.QtWidgets import QWidget
        from PyQt6.QtCore import QSize
        from qfluentwidgets import MessageBoxBase
        from controllers.business_logic import cfg

        class PaddingConfigDialog(MessageBoxBase):
            def __init__(self, parent=None):
                super().__init__(parent)
                title_label = QLabel("组件屏幕边距 (Margin)", self)
                title_label.setStyleSheet("font-size: 18px; font-weight: bold; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
                self.viewLayout.addWidget(title_label)
                container = QWidget(self)
                c_layout = QVBoxLayout(container)
                c_layout.setContentsMargins(0, 0, 0, 0)
                c_layout.setSpacing(16)
                self.viewLayout.addWidget(container)
                self.preview = ScreenPaddingPreview(container)
                self.preview.setMinimumHeight(220)
                self.preview.setMaximumHeight(260)
                c_layout.addWidget(self.preview)
                controls = QGridLayout()
                controls.setHorizontalSpacing(12)
                controls.setVerticalSpacing(8)
                top_label = QLabel("上", container)
                bottom_label = QLabel("下", container)
                left_label = QLabel("左", container)
                right_label = QLabel("右", container)
                for lbl in (top_label, bottom_label, left_label, right_label):
                    lbl.setStyleSheet("font-size: 14px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
                self.spin_top = SpinBox(container)
                self.spin_bottom = SpinBox(container)
                self.spin_left = SpinBox(container)
                self.spin_right = SpinBox(container)
                for s in (self.spin_top, self.spin_bottom, self.spin_left, self.spin_right):
                    s.setRange(0, 200)
                    s.setFixedWidth(80)
                self.lockSwitch = SwitchButton(container)
                self.lockSwitch.setOnText("锁定相同")
                self.lockSwitch.setOffText("锁定相同")
                controls.addWidget(top_label, 0, 0)
                controls.addWidget(self.spin_top, 0, 1)
                controls.addWidget(bottom_label, 0, 2)
                controls.addWidget(self.spin_bottom, 0, 3)
                controls.addWidget(left_label, 1, 0)
                controls.addWidget(self.spin_left, 1, 1)
                controls.addWidget(right_label, 1, 2)
                controls.addWidget(self.spin_right, 1, 3)
                controls.addWidget(self.lockSwitch, 0, 4, 2, 1)
                c_layout.addLayout(controls)
                preset_row = QHBoxLayout()
                preset_row.setSpacing(12)
                preset_label = QLabel("预设", container)
                preset_label.setStyleSheet("font-size: 14px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
                preset_row.addWidget(preset_label)
                self.btn_small = PushButton("小", container)
                self.btn_medium = PushButton("中", container)
                self.btn_large = PushButton("大", container)
                for b in (self.btn_small, self.btn_medium, self.btn_large):
                    b.setFixedWidth(60)
                preset_row.addWidget(self.btn_small)
                preset_row.addWidget(self.btn_medium)
                preset_row.addWidget(self.btn_large)
                preset_row.addStretch()
                c_layout.addLayout(preset_row)

                # Update text colors based on theme
                text_color = "white" if isDarkTheme() else "black"
                for lbl in (title_label, top_label, bottom_label, left_label, right_label, preset_label):
                    style = lbl.styleSheet()
                    import re
                    style = re.sub(r'color\s*:\s*[^;]+;', '', style)
                    lbl.setStyleSheet(style + f" color: {text_color};")
                t = int(cfg.screenPaddingTop.value)
                b = int(cfg.screenPaddingBottom.value)
                l = int(cfg.screenPaddingLeft.value)
                r = int(cfg.screenPaddingRight.value)
                lock = bool(cfg.screenPaddingLock.value)
                self.spin_top.setValue(t)
                self.spin_bottom.setValue(b)
                self.spin_left.setValue(l)
                self.spin_right.setValue(r)
                self.lockSwitch.setChecked(lock)
                self.preview.set_margins(t, b, l, r)
                self.spin_top.valueChanged.connect(self._on_spin_changed)
                self.spin_bottom.valueChanged.connect(self._on_spin_changed)
                self.spin_left.valueChanged.connect(self._on_spin_changed)
                self.spin_right.valueChanged.connect(self._on_spin_changed)
                self.lockSwitch.checkedChanged.connect(self._on_lock_changed)
                self.btn_small.clicked.connect(lambda: self._apply_preset(10))
                self.btn_medium.clicked.connect(lambda: self._apply_preset(20))
                self.btn_large.clicked.connect(lambda: self._apply_preset(40))
                self.cancelButton.setText("取消")
                self.yesButton.setText("应用")
                self.yesButton.clicked.connect(self._apply_and_close)

            def _sync_preview(self):
                t = self.spin_top.value()
                b = self.spin_bottom.value()
                l = self.spin_left.value()
                r = self.spin_right.value()
                self.preview.set_margins(t, b, l, r)

            def _on_spin_changed(self, _):
                if self.lockSwitch.isChecked():
                    sender = self.sender()
                    if sender is not None:
                        v = sender.value()
                        self.spin_top.blockSignals(True)
                        self.spin_bottom.blockSignals(True)
                        self.spin_left.blockSignals(True)
                        self.spin_right.blockSignals(True)
                        self.spin_top.setValue(v)
                        self.spin_bottom.setValue(v)
                        self.spin_left.setValue(v)
                        self.spin_right.setValue(v)
                        self.spin_top.blockSignals(False)
                        self.spin_bottom.blockSignals(False)
                        self.spin_left.blockSignals(False)
                        self.spin_right.blockSignals(False)
                self._sync_preview()

            def _on_lock_changed(self, checked):
                if checked:
                    v = self.spin_top.value()
                    self.spin_bottom.setValue(v)
                    self.spin_left.setValue(v)
                    self.spin_right.setValue(v)
                self._sync_preview()

            def _apply_preset(self, margin):
                self.spin_top.setValue(margin)
                self.spin_bottom.setValue(margin)
                self.spin_left.setValue(margin)
                self.spin_right.setValue(margin)
                self._sync_preview()

            def _apply_and_close(self):
                t = self.spin_top.value()
                b = self.spin_bottom.value()
                l = self.spin_left.value()
                r = self.spin_right.value()
                lock = self.lockSwitch.isChecked()
                cfg.screenPaddingTop.value = t
                cfg.screenPaddingBottom.value = b
                cfg.screenPaddingLeft.value = l
                cfg.screenPaddingRight.value = r
                cfg.screenPaddingLock.value = lock
                cfg.save()
                
                # Notify settings window that config changed
                parent = self.parent()
                while parent and not hasattr(parent, 'on_config_changed'):
                    parent = parent.parent()
                if parent:
                    parent.on_config_changed()
                    
                self.accept()

        w = PaddingConfigDialog(self.window())
        w.exec()
