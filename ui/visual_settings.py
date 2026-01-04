from PyQt6.QtWidgets import QWidget, QHBoxLayout, QVBoxLayout, QLabel, QFrame
from PyQt6.QtCore import Qt, QTimer, QDateTime, QCoreApplication, QRectF
from PyQt6.QtGui import QFont, QRegion, QPainterPath
from qfluentwidgets import SettingCard, Theme, isDarkTheme, ComboBox, SwitchButton, PushButton, BodyLabel


def tr(text: str) -> str:
    return QCoreApplication.translate("ClockSettings", text)

class ClockPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedHeight(100)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(16, 6, 16, 6)
        inner_layout.setSpacing(12)
        inner_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_time = QLabel()
        self.lbl_time.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_date = QLabel()
        self.lbl_date.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_lunar = QLabel()
        self.lbl_lunar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        inner_layout.addWidget(self.lbl_time)
        inner_layout.addWidget(self.lbl_date)
        inner_layout.addWidget(self.lbl_lunar)
        
        layout.addWidget(self.container)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        
        self.show_seconds = False
        self.show_date = False
        self.show_lunar = False
        self.font_weight = "Bold"
        
        self.update_time()
        self.update_style()
        QTimer.singleShot(0, self._update_radius)

    def update_settings(self, seconds, date, lunar, weight):
        self.show_seconds = seconds
        self.show_date = date
        self.show_lunar = lunar
        self.font_weight = weight
        self.update_style()
        self.update_time()
        QTimer.singleShot(0, self._update_radius)

    def _update_radius(self):
        h = self.container.height()
        if h <= 0:
            h = self.container.sizeHint().height()
        
        radius = int(h / 2)
        
        style = self.container.styleSheet()
        import re
        new_style = re.sub(r"border-radius: .*?;", f"border-radius: {radius}px;", style)
        self.container.setStyleSheet(new_style)

        rect = self.container.rect()
        if rect.width() > 0 and rect.height() > 0:
            r = min(radius, int(rect.height() / 2))
            path = QPainterPath()
            path.addRoundedRect(QRectF(rect), float(r), float(r))
            self.container.setMask(QRegion(path.toFillPolygon().toPolygon()))
        else:
            self.container.clearMask()

    def update_style(self):
        theme = Theme.DARK if isDarkTheme() else Theme.LIGHT
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "black"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 10px;
            }}
            QLabel {{
                color: {text_color};
                background: transparent;
                border: none;
            }}
        """)
        
        weight_map = {
            "Light": QFont.Weight.Light,
            "Normal": QFont.Weight.Normal,
            "DemiBold": QFont.Weight.DemiBold,
            "Bold": QFont.Weight.Bold
        }
        w = weight_map.get(self.font_weight, QFont.Weight.Bold)
        font = self.lbl_time.font()
        font.setWeight(w)
        font.setPixelSize(20)
        self.lbl_time.setFont(font)
        
        font_small = self.lbl_date.font()
        font_small.setWeight(QFont.Weight.Normal)
        font_small.setPixelSize(10)
        self.lbl_date.setFont(font_small)
        self.lbl_lunar.setFont(font_small)

    def update_time(self):
        now = QDateTime.currentDateTime()
        hour = now.time().hour()
        minute = now.time().minute()
        period = tr("上午") if hour < 12 else tr("下午")
        time_str = f"{period} {hour:02d}:{minute:02d}"
        
        if self.show_seconds:
             time_str += f":{now.time().second():02d}"
             
        self.lbl_time.setText(time_str)
        
        if self.show_date:
            date_str = tr("{year}年{month}月{day}日").format(
                year=now.date().year(),
                month=now.date().month(),
                day=now.date().day(),
            )
            self.lbl_date.setText(date_str)
            self.lbl_date.setVisible(True)
        else:
            self.lbl_date.setVisible(False)
            
        if self.show_lunar:
            self.lbl_lunar.setText(tr("农历（示例）"))
            self.lbl_lunar.setVisible(True)
        else:
            self.lbl_lunar.setVisible(False)
class ClockSettingCard(SettingCard):
    def __init__(self, icon, title, content, parent=None):
        super().__init__(icon, title, content, parent)
        self.configBtn = PushButton(tr("打开时钟详细设置"), self)
        self.configBtn.setFixedWidth(120)
        self.configBtn.clicked.connect(self._show_config_flyout)
        self.hBoxLayout.addWidget(self.configBtn)
        self.hBoxLayout.addSpacing(16)

    def _show_config_flyout(self):
        self._show_overlay_menu()

    def _show_overlay_menu(self):
        from qfluentwidgets import MessageBoxBase, SubtitleLabel

        class ClockConfigDialog(MessageBoxBase):
            def __init__(self, parent=None):
                super().__init__(parent)
                self.titleLabel = SubtitleLabel(tr("时钟显示内容与样式配置"), self)
                self.viewLayout.addWidget(self.titleLabel)

                self.preview = ClockPreviewWidget()
                self.viewLayout.addWidget(self.preview)

                self.controlsWidget = QWidget()
                self.controlsLayout = QVBoxLayout(self.controlsWidget)
                self.controlsLayout.setSpacing(10)

                row_weight = QHBoxLayout()
                row_weight.addWidget(BodyLabel(tr("时钟时间字体粗细")))
                self.weightCombo = ComboBox()
                self.weightCombo.addItems([tr("细"), tr("常规"), tr("半粗"), tr("粗体")])
                self.weightCombo.setFixedWidth(120)
                row_weight.addWidget(self.weightCombo)
                self.controlsLayout.addLayout(row_weight)

                self.secSwitch = SwitchButton()
                self.secSwitch.setOnText(tr("显示秒"))
                self.secSwitch.setOffText(tr("显示秒"))

                self.dateSwitch = SwitchButton()
                self.dateSwitch.setOnText(tr("显示日期"))
                self.dateSwitch.setOffText(tr("显示日期"))

                self.lunarSwitch = SwitchButton()
                self.lunarSwitch.setOnText(tr("显示农历"))
                self.lunarSwitch.setOffText(tr("显示农历"))

                row1 = QHBoxLayout()
                row1.addWidget(BodyLabel(tr("是否在时钟中显示秒数")))
                row1.addWidget(self.secSwitch)

                row2 = QHBoxLayout()
                row2.addWidget(BodyLabel(tr("是否在时钟下方显示公历日期")))
                row2.addWidget(self.dateSwitch)

                row3 = QHBoxLayout()
                row3.addWidget(BodyLabel(tr("是否在时钟下方显示农历信息")))
                row3.addWidget(self.lunarSwitch)

                self.controlsLayout.addLayout(row1)
                self.controlsLayout.addLayout(row2)
                self.controlsLayout.addLayout(row3)

                self.viewLayout.addWidget(self.controlsWidget)

                from controllers.business_logic import cfg
                weight_map_rev = {
                    "Light": tr("细"),
                    "Normal": tr("常规"),
                    "DemiBold": tr("半粗"),
                    "Bold": tr("粗体")
                }
                self.weightCombo.setCurrentText(weight_map_rev.get(cfg.clockFontWeight.value, tr("粗体")))
                self.secSwitch.setChecked(cfg.clockShowSeconds.value)
                self.dateSwitch.setChecked(cfg.clockShowDate.value)
                self.lunarSwitch.setChecked(cfg.clockShowLunar.value)

                self.update_preview()

                self.weightCombo.currentTextChanged.connect(self.update_cfg)
                self.secSwitch.checkedChanged.connect(self.update_cfg)
                self.dateSwitch.checkedChanged.connect(self.update_cfg)
                self.lunarSwitch.checkedChanged.connect(self.update_cfg)

                self.cancelButton.hide()
                self.yesButton.setText(tr("完成"))

            def update_cfg(self):
                from controllers.business_logic import cfg
                weight_map = {
                    tr("细"): "Light",
                    tr("常规"): "Normal",
                    tr("半粗"): "DemiBold",
                    tr("粗体"): "Bold"
                }
                cfg.clockFontWeight.value = weight_map.get(self.weightCombo.currentText(), "Bold")
                cfg.clockShowSeconds.value = self.secSwitch.isChecked()
                cfg.clockShowDate.value = self.dateSwitch.isChecked()
                cfg.clockShowLunar.value = self.lunarSwitch.isChecked()
                cfg.save()
                self.update_preview()

            def update_preview(self):
                weight_map = {
                    tr("细"): "Light",
                    tr("常规"): "Normal",
                    tr("半粗"): "DemiBold",
                    tr("粗体"): "Bold"
                }
                self.preview.update_settings(
                    self.secSwitch.isChecked(),
                    self.dateSwitch.isChecked(),
                    self.lunarSwitch.isChecked(),
                    weight_map.get(self.weightCombo.currentText(), "Bold"),
                )

        w = ClockConfigDialog(self.window())
        w.exec()
