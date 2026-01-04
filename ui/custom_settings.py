from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGridLayout
from PyQt6.QtCore import Qt, QRectF, QPropertyAnimation, QEasingCurve, pyqtProperty as Property, QCoreApplication
from PyQt6.QtGui import QPainter, QColor, QPen, QPainterPath
from qfluentwidgets import OptionsSettingCard, Theme, isDarkTheme, themeColor, SettingCard, PushButton, BodyLabel, SpinBox, SwitchButton


def tr(text: str) -> str:
    return QCoreApplication.translate("CustomSettings", text)


class SchematicOptionButton(QWidget):
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
        
        self.hover_anim = QPropertyAnimation(self, b"hoverProgress", self)
        self.hover_anim.setDuration(150)
        self.scale_anim = QPropertyAnimation(self, b"clickScale", self)
        self.scale_anim.setDuration(100)
        
    def get_hover_progress(self):
        return self._hover_progress
        
    def set_hover_progress(self, v):
        self._hover_progress = v
        self.update()
        
    hoverProgress = Property(float, get_hover_progress, set_hover_progress)
    
    def get_click_scale(self):
        return self._click_scale
        
    def set_click_scale(self, v):
        self._click_scale = v
        self.update()
        
    clickScale = Property(float, get_click_scale, set_click_scale)
    
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
            if self.parent():
                self.parent().on_option_clicked(self.value)
        super().mouseReleaseEvent(e)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        painter.translate(self.width() / 2, self.height() / 2)
        painter.scale(self._click_scale, self._click_scale)
        painter.translate(-self.width() / 2, -self.height() / 2)
        
        rect = QRectF(5, 5, self.width() - 10, self.height() - 30)
        
        is_dark = isDarkTheme()
        accent = themeColor()
        
        bg_color = QColor(0, 0, 0, 0)
        border_color = QColor(255, 255, 255, 20) if is_dark else QColor(0, 0, 0, 10)
        border_width = 1.0
        
        if self._is_checked:
            border_color = accent
            border_width = 2.0
            bg_color = QColor(accent)
            bg_color.setAlpha(20)
        elif self._hover_progress > 0.01:
            alpha = int(10 * self._hover_progress)
            bg_color = QColor(255, 255, 255, alpha) if is_dark else QColor(0, 0, 0, alpha)
            border_color = QColor(255, 255, 255, 50) if is_dark else QColor(0, 0, 0, 30)
            
        path = QPainterPath()
        path.addRoundedRect(rect, 8, 8)
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg_color)
        painter.drawPath(path)
        
        pen = QPen(border_color, border_width)
        if border_width > 1:
            pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawPath(path)
        
        content_rect = rect.adjusted(4, 4, -4, -4)
        self.draw_content(painter, content_rect)
        
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
        elif self.schematic_type == "toolbar_pos":
            self.draw_toolbar_content(painter, rect)
        elif self.schematic_type == "language":
            self.draw_language_content(painter, rect)
            
    def draw_theme_content(self, painter, rect):
        if self.value == "Light":
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(243, 243, 243))
            painter.drawRoundedRect(rect, 4, 4)
            painter.setBrush(QColor(0, 0, 0, 20))
            painter.drawRoundedRect(rect.adjusted(5, 5, -20, -rect.height()+10), 2, 2)
            painter.drawRoundedRect(rect.adjusted(5, 15, -10, -rect.height()+20), 2, 2)
            
        elif self.value == "Dark":
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(32, 32, 32))
            painter.drawRoundedRect(rect, 4, 4)
            painter.setBrush(QColor(255, 255, 255, 20))
            painter.drawRoundedRect(rect.adjusted(5, 5, -20, -rect.height()+10), 2, 2)
            painter.drawRoundedRect(rect.adjusted(5, 15, -10, -rect.height()+20), 2, 2)
            
        elif self.value == "Auto":
            path = QPainterPath()
            path.addRoundedRect(rect, 4, 4)
            painter.save()
            painter.setClipPath(path)
            
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(243, 243, 243))
            painter.drawRect(rect) 
            
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

    def draw_language_content(self, painter, rect):
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        is_dark = isDarkTheme()
        bg = QColor(32, 32, 32) if is_dark else QColor(243, 243, 243)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg)
        painter.drawRoundedRect(rect, 4, 4)
        if self.value == "Simplified Chinese":
            code = "简"
        elif self.value == "Traditional Chinese":
            code = "繁"
        elif self.value == "English":
            code = "EN"
        elif self.value == "Japanese":
            code = "あ"
        elif self.value == "Tibetan":
            code = "བོད"
        else:
            code = "?"
        painter.setPen(QColor(255, 255, 255) if is_dark else QColor(0, 0, 0))
        font = painter.font()
        font.setPixelSize(20)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, code)
        painter.restore()


class SchematicWidget(QWidget):
    def __init__(self, schematic_type="theme", configItem=None, parent=None):
        super().__init__(parent)
        self.schematic_type = schematic_type
        self.configItem = configItem
        
        if schematic_type == "toolbar_pos":
            self.setFixedHeight(260)
        else:
            self.setFixedHeight(140)
            
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
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
            options = [
                ("Light", tr("浅色主题")),
                ("Dark", tr("深色主题")),
                ("Auto", tr("跟随系统主题")),
            ]
            for val, text in options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn)
                self.buttons.append(btn)
                
        elif self.schematic_type == "nav_pos":
            options = [
                ("BottomSides", tr("底部两端")),
                ("MiddleSides", tr("中部两侧")),
            ]
            for val, text in options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn)
                self.buttons.append(btn)
                
        elif self.schematic_type == "toolbar_pos":
            grid_options = [
                ("MiddleLeft", tr("中部左侧"), 0, 0),
                ("MiddleRight", tr("中部右侧"), 0, 2),
                ("BottomLeft", tr("底部左侧"), 1, 0),
                ("BottomCenter", tr("底部居中"), 1, 1),
                ("BottomRight", tr("底部右侧"), 1, 2),
            ]
            
            for val, text, row, col in grid_options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn, row, col)
                self.buttons.append(btn)
        elif self.schematic_type == "language":
            options = [
                ("Simplified Chinese", tr("简体中文")),
                ("Traditional Chinese", tr("繁體中文")),
                ("English", tr("English")),
                ("Japanese", tr("日本語")),
                ("Tibetan", tr("བོད་ཡིག")),
            ]
            for val, text in options:
                btn = SchematicOptionButton(val, text, self.schematic_type, self)
                self.layout.addWidget(btn)
                self.buttons.append(btn)
            
        self.update_state()
        
    def update_state(self):
        if not self.configItem:
            return
            
        current_value = self.configItem.value
        if hasattr(current_value, "value"):
            current_value = current_value.value
        if isinstance(current_value, Theme):
            if current_value == Theme.LIGHT: current_value = "Light"
            elif current_value == Theme.DARK: current_value = "Dark"
            elif current_value == Theme.AUTO: current_value = "Auto"
            
        for btn in self.buttons:
            try:
                if btn and not btn.parent():
                    pass
                btn.setChecked(btn.value == current_value)
            except RuntimeError:
                continue
            
    def on_option_clicked(self, val):
        if not self.configItem:
            return
            
        current = self.configItem.value
        final_val = val
        
        if self.schematic_type == "theme":
            if isinstance(current, Theme):
                if val == "Light": final_val = Theme.LIGHT
                elif val == "Dark": final_val = Theme.DARK
                elif val == "Auto": final_val = Theme.AUTO
            else:
                final_val = val
            
        if self.configItem.value != final_val:
            self.configItem.value = final_val
            from controllers.business_logic import cfg
            cfg.save()
            self.update_state()

    def update(self):
        self.update_state()
        super().update()


class SchematicOptionsSettingCard(OptionsSettingCard):
    def __init__(self, configItem, icon, title, content, texts, schematic_type="theme", parent=None):
        super().__init__(configItem, icon, title, content, texts, parent)
        self.schematic = SchematicWidget(schematic_type, configItem, self)
        
        self._config_item = configItem
        self._config_item.valueChanged.connect(self._on_config_changed)
        
        self.viewLayout.insertWidget(0, self.schematic)
        self.viewLayout.insertSpacing(1, 10)
        
        for i in range(self.viewLayout.count()):
            item = self.viewLayout.itemAt(i)
            widget = item.widget()
            if widget and widget != self.schematic:
                widget.hide()

    def _on_config_changed(self, v):
        try:
            if hasattr(self, "schematic") and self.schematic:
                self.schematic.update()
        except RuntimeError:
            try:
                self._config_item.valueChanged.disconnect(self._on_config_changed)
            except Exception:
                pass

    def closeEvent(self, event):
        try:
            self._config_item.valueChanged.disconnect(self._on_config_changed)
        except Exception:
            pass
        super().closeEvent(event)


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
        from PySide6.QtWidgets import QApplication
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
        self.configBtn = PushButton(tr("调整边距"), self)
        self.configBtn.setFixedWidth(120)
        self.configBtn.clicked.connect(self._show_config_overlay)
        self.hBoxLayout.addWidget(self.configBtn, 0, Qt.AlignmentFlag.AlignRight)
        self.hBoxLayout.addSpacing(16)

    def _show_config_overlay(self):
        from PySide6.QtWidgets import QWidget
        from PySide6.QtCore import QSize
        from qfluentwidgets import MessageBoxBase, SubtitleLabel
        from controllers.business_logic import cfg

        class PaddingConfigDialog(MessageBoxBase):
            def __init__(self, parent=None):
                super().__init__(parent)
                title_label = SubtitleLabel(tr("屏幕边距 (Margin) 设置"), self)
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
                top_label = BodyLabel(tr("Top Margin"), container)
                bottom_label = BodyLabel(tr("Bottom Margin"), container)
                left_label = BodyLabel(tr("Left Margin"), container)
                right_label = BodyLabel(tr("Right Margin"), container)
                self.spin_top = SpinBox(container)
                self.spin_bottom = SpinBox(container)
                self.spin_left = SpinBox(container)
                self.spin_right = SpinBox(container)
                for s in (self.spin_top, self.spin_bottom, self.spin_left, self.spin_right):
                    s.setRange(0, 200)
                    s.setFixedWidth(80)
                self.lockSwitch = SwitchButton(container)
                self.lockSwitch.setOnText(tr("锁定四个边距相同"))
                self.lockSwitch.setOffText(tr("锁定四个边距相同"))
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
                preset_row.addWidget(BodyLabel(tr("快速预设"), container))
                self.btn_small = PushButton(tr("小边距"), container)
                self.btn_medium = PushButton(tr("中等边距"), container)
                self.btn_large = PushButton(tr("大边距"), container)
                for b in (self.btn_small, self.btn_medium, self.btn_large):
                    b.setFixedWidth(60)
                preset_row.addWidget(self.btn_small)
                preset_row.addWidget(self.btn_medium)
                preset_row.addWidget(self.btn_large)
                preset_row.addStretch()
                c_layout.addLayout(preset_row)
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
                self.cancelButton.setText(tr("取消"))
                self.yesButton.setText(tr("应用当前设置"))
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
                self.accept()

        w = PaddingConfigDialog(self.window())
        w.exec()
