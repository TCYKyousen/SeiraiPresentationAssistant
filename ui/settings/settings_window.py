from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QUrl
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QFrame, QHBoxLayout, QSizePolicy, QLabel
from PyQt6.QtGui import QFont, QPainter, QColor, QLinearGradient, QPixmap, QDesktopServices
import json
import os

from qfluentwidgets import (
    FluentIcon as FIF,
    SwitchSettingCard, OptionsSettingCard, PushSettingCard,
    SettingCard, PrimaryPushButton, PushButton,
    SmoothScrollArea, ExpandLayout, Theme, setTheme, setThemeColor,
    FluentWindow, NavigationItemPosition, isDarkTheme,
    IndeterminateProgressRing,
    FluentLabelBase, ImageLabel,
    InfoBar, InfoBarPosition, CardWidget, IconWidget, TextEdit,
    ColorSettingCard, RangeSettingCard, ComboBox
)
from ui.settings.custom_settings import SchematicOptionsSettingCard, ScreenPaddingSettingCard
from controllers.business_logic import cfg, logger, log, BusinessLogicController
from controllers.i18n_manager import tr
from ui.crash_dialog import CrashDialog, trigger_crash
from ui.settings.visual_settings import ClockSettingCard


def _create_page(parent: QWidget):
    page = QWidget(parent)
    page.setStyleSheet("background-color: transparent;")
    layout = QVBoxLayout(page)
    layout.setContentsMargins(24, 24, 24, 24)
    layout.setSpacing(0)

    scroll = SmoothScrollArea(page)
    scroll.setObjectName("scrollInterface")
    scroll.setStyleSheet("SmoothScrollArea { background-color: transparent; border: none; }")

    content = QWidget()
    content.setStyleSheet("background-color: transparent;")
    content.setObjectName("scrollWidget")
    expand_layout = ExpandLayout(content)

    scroll.setWidget(content)
    scroll.setWidgetResizable(True)

    layout.addWidget(scroll)
    return page, content, expand_layout


class UpdateManagerWidget(CardWidget):
    checkUpdateClicked = pyqtSignal()
    
    def __init__(self, current_version, code_name="", build_date="", parent=None):
        super().__init__(parent)
        # self.setFixedHeight(380) # Removed fixed height to prevent truncation
        self.vBoxLayout = QVBoxLayout(self)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setSpacing(0)
        
        # 背景装饰 (渐变或微光效果可以通过样式模拟)
        self.headerFrame = QFrame(self)
        self.headerFrame.setFixedHeight(160)
        self.headerLayout = QVBoxLayout(self.headerFrame)
        self.headerLayout.setContentsMargins(30, 30, 30, 30)
        self.headerLayout.setSpacing(10)
        
        # 居中对齐的状态信息
        self.statusIcon = IconWidget(FIF.SYNC, self.headerFrame)
        self.statusIcon.setFixedSize(48, 48)
        
        self.statusLabel = QLabel(tr("Check Update"), self.headerFrame)
        self.statusLabel.setStyleSheet("font-size: 18px; font-weight: bold; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        
        text_lines = []
        text_lines.append(f"{tr('Current Version')}: {current_version}")
        if code_name:
            text_lines.append(f"{tr('Code Name')}: {code_name}")
        if build_date:
            text_lines.append(f"{tr('Build Date')}: {build_date}")
        self.versionLabel = QLabel("\n".join(text_lines), self.headerFrame)
        self.versionLabel.setStyleSheet("font-size: 12px; color: rgba(128, 128, 128, 0.8); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        
        self.headerLayout.addWidget(self.statusIcon, 0, Qt.AlignmentFlag.AlignCenter)
        self.headerLayout.addWidget(self.statusLabel, 0, Qt.AlignmentFlag.AlignCenter)
        self.headerLayout.addWidget(self.versionLabel, 0, Qt.AlignmentFlag.AlignCenter)
        
        self.vBoxLayout.addWidget(self.headerFrame)
        
        # 内容区域容器
        self.contentArea = QWidget(self)
        self.contentLayout = QVBoxLayout(self.contentArea)
        self.contentLayout.setContentsMargins(20, 0, 20, 20)
        self.contentLayout.setSpacing(15)
        
        # 进度环 (悬浮或居中)
        self.progressRing = IndeterminateProgressRing(self.contentArea)
        self.progressRing.setFixedSize(32, 32)
        self.progressRing.hide()
        self.contentLayout.addWidget(self.progressRing, 0, Qt.AlignmentFlag.AlignCenter)
        
        # 更新日志卡片 (内嵌在主卡片中)
        self.logCard = QFrame(self.contentArea)
        self.logCard.setObjectName("logCard")
        self.logCardLayout = QVBoxLayout(self.logCard)
        self.logCardLayout.setContentsMargins(15, 15, 15, 15)
        self.logCardLayout.setSpacing(8)
        
        is_dark = isDarkTheme()
        bg = "rgba(255, 255, 255, 0.05)" if is_dark else "rgba(0, 0, 0, 0.03)"
        border = "rgba(255, 255, 255, 0.1)" if is_dark else "rgba(0, 0, 0, 0.08)"
        self.logCard.setStyleSheet(f"""
            #logCard {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 10px;
            }}
        """)
        
        self.logScroll = SmoothScrollArea(self.logCard)
        self.logScroll.setWidgetResizable(True)
        self.logScroll.setStyleSheet("background: transparent; border: none;")
        
        self.logContent = QLabel(tr("Check Update Desc"), self.logScroll)
        self.logContent.setWordWrap(True)
        self.logContent.setStyleSheet("font-size: 14px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        self.logScroll.setWidget(self.logContent)
        
        self.logCardLayout.addWidget(self.logScroll)
        self.contentLayout.addWidget(self.logCard, 1)
        
        # 底部操作栏
        self.bottomBar = QHBoxLayout()
        self.checkBtn = PrimaryPushButton(tr("Check Now"), self.contentArea)
        self.checkBtn.setFixedWidth(140)
        self.checkBtn.clicked.connect(self._on_check_clicked)
        
        self.bottomBar.addStretch(1)
        self.bottomBar.addWidget(self.checkBtn)
        self.bottomBar.addStretch(1)
        
        self.contentLayout.addLayout(self.bottomBar)
        self.vBoxLayout.addWidget(self.contentArea)

        self.setStyleSheet(f"""
            UpdateManagerWidget {{
                background-color: {"#2d2d2d" if is_dark else "#ffffff"};
                border: 1px solid {"#3f3f3f" if is_dark else "#e5e5e5"};
                border-radius: 12px;
            }}
        """)
        
    def _on_check_clicked(self):
        self.statusLabel.setText("正在连接服务器...")
        self.statusIcon.setIcon(FIF.SYNC)
        self.checkBtn.setEnabled(False)
        self.checkBtn.setText("请稍候")
        self.progressRing.show()
        self.logCard.hide()
        self.checkUpdateClicked.emit()
        
    def on_check_finished(self, success, message):
        self.checkBtn.setEnabled(True)
        self.checkBtn.setText("重新检查")
        self.progressRing.hide()
        self.logCard.show()
        
        if success:
            if "最新" in message:
                self.statusLabel.setText(tr("Latest Version"))
                self.statusIcon.setIcon(FIF.ACCEPT)
                self.logContent.setText("您当前使用的是 Kazuha 的最新稳定版本，无需更新。")
            else:
                self.statusLabel.setText(tr("New Version"))
                self.statusIcon.setIcon(FIF.UPDATE)
        else:
            self.statusLabel.setText(tr("Check Failed"))
            self.statusIcon.setIcon(FIF.CLOSE)
            self.logContent.setText(f"无法获取更新信息：\n{message}")
            
    def set_update_info(self, version, body):
        self.statusLabel.setText(f"发现新版本 v{version}")
        self.statusIcon.setIcon(FIF.UPDATE)
        self.checkBtn.setText("获取更新")
        self.logCard.show()
        
        # 使用富文本简单模拟列表
        formatted_body = body.replace("\n", "<br>")
        self.logContent.setText(f"<b>v{version} 更新详情:</b><br><br>{formatted_body}")


class SettingsWindow(FluentWindow):
    configChanged = pyqtSignal()
    checkUpdateClicked = pyqtSignal()
    restartClicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("界面设置 - Kazuha")
        # 125% 缩放时最小宽度 1471 -> 逻辑宽度 = 1471 / 1.25 = 1176.8
        self.setMinimumWidth(1177)
        self.resize(1200, 800)

        try:
            self.setMicaEffectEnabled(True)
        except Exception:
            pass

        self.generalInterface, generalContent, generalLayout = _create_page(self)
        self.generalInterface.setObjectName("settings-general")

        self.personalInterface, personalContent, personalLayout = _create_page(self)
        self.personalInterface.setObjectName("settings-personal")

        self.clockInterface, clockContent, clockLayout = _create_page(self)
        self.clockInterface.setObjectName("settings-clock")
        self.clockLayout = clockLayout
        self.clockConflictInfoBar = None

        self.updateInterface, updateContent, updateLayout = _create_page(self)
        self.updateInterface.setObjectName("settings-update")
        
        self.debugInterface, debugContent, debugLayout = _create_page(self)
        self.debugInterface.setObjectName("settings-debug")

        self.aboutInterface, aboutContent, aboutLayout = _create_page(self)
        self.aboutInterface.setObjectName("settings-about")

        self._scrollWidgets = [generalContent, personalContent, clockContent, updateContent, aboutContent, debugContent]
        self._scale_preview_timer = QTimer(self)
        self._scale_preview_timer.setSingleShot(True)
        self._scale_preview_timer.timeout.connect(self._trigger_scale_preview)

        self.generalPageTitle = QLabel(tr("General"), generalContent)
        self.generalPageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        generalLayout.addWidget(self.generalPageTitle)

        self.generalHeader = QLabel(tr("General"), generalContent)
        self.generalHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        generalLayout.addWidget(self.generalHeader)

        self.startupCard = SwitchSettingCard(
            FIF.POWER_BUTTON,
            tr("Enable StartUp"),
            tr("Enable StartUp Desc"),
            configItem=cfg.enableStartUp,
            parent=generalContent
        )

        generalLayout.addWidget(self.startupCard)

        self.languageCard = OptionsSettingCard(
            cfg.language,
            FIF.LANGUAGE,
            tr("Language"),
            tr("Language Desc"),
            texts=["Simplified Chinese", "Traditional Chinese", "English", "Japanese"],
            parent=generalContent
        )
        generalLayout.addWidget(self.languageCard)
        cfg.language.valueChanged.connect(self.on_config_changed)
        
        self.systemNotificationCard = SwitchSettingCard(
            FIF.MESSAGE,
            tr("System Notification"),
            tr("System Notification Desc"),
            configItem=cfg.enableSystemNotification,
            parent=generalContent
        )
        generalLayout.addWidget(self.systemNotificationCard)
        
        self.globalSoundCard = SwitchSettingCard(
            FIF.MUSIC,
            tr("Global Sound"),
            tr("Global Sound Desc"),
            configItem=cfg.enableGlobalSound,
            parent=generalContent
        )
        generalLayout.addWidget(self.globalSoundCard)
        
        self.globalAnimationCard = SwitchSettingCard(
            FIF.SPEED_MEDIUM,
            tr("Global Animation"),
            tr("Global Animation Desc"),
            configItem=cfg.enableGlobalAnimation,
            parent=generalContent
        )
        generalLayout.addWidget(self.globalAnimationCard)

        self.personalPageTitle = QLabel(tr("Personalization"), personalContent)
        self.personalPageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        personalLayout.addWidget(self.personalPageTitle)

        self.themeHeader = QLabel(tr("Personalization"), personalContent)
        self.themeHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        personalLayout.addWidget(self.themeHeader)

        self.themeModeCard = SchematicOptionsSettingCard(
            cfg.themeMode,
            FIF.BRUSH,
            tr("App Theme"),
            tr("App Theme Desc"),
            texts=[tr("Light"), tr("Dark"), tr("Auto")],
            schematic_type="theme",
            parent=personalContent
        )
        personalLayout.addWidget(self.themeModeCard)

        self.themeColorModeCard = OptionsSettingCard(
            cfg.themeColorMode,
            FIF.PALETTE,
            tr("Color Mode"),
            tr("Color Mode Desc"),
            texts=[tr("Screen Color"), tr("Wallpaper Color"), tr("System Accent"), tr("Custom")],
            parent=personalContent
        )
        personalLayout.addWidget(self.themeColorModeCard)

        self.themeColorCustomCard = ColorSettingCard(
            cfg.themeColorCustom,
            FIF.PALETTE,
            tr("Custom Color"),
            tr("Custom Color Desc"),
            parent=personalContent
        )
        personalLayout.addWidget(self.themeColorCustomCard)

        self.themeColorIntervalCard = RangeSettingCard(
            cfg.themeColorInterval,
            FIF.HISTORY,
            tr("Update Interval"),
            tr("Update Interval Desc"),
            parent=personalContent
        )
        personalLayout.addWidget(self.themeColorIntervalCard)
        
        # Connect visibility
        self._update_theme_color_cards_visibility()
        cfg.themeColorMode.valueChanged.connect(self._update_theme_color_cards_visibility)

        self.layoutHeader = QLabel(tr("Interface Layout"), personalContent)
        self.layoutHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        personalLayout.addWidget(self.layoutHeader)
        
        self.navPosCard = SchematicOptionsSettingCard(
            cfg.navPosition,
            FIF.ALIGNMENT,
            tr("Nav Position"),
            tr("Nav Position Desc"),
            texts=[tr("Bottom Sides"), tr("Middle Sides")],
            schematic_type="nav_pos",
            parent=personalContent
        )

        self.componentScaleCard = RangeSettingCard(
            cfg.componentScale,
            FIF.ZOOM,
            tr("Component Scale"),
            tr("Component Scale Desc"),
            parent=personalContent
        )

        self.paddingCard = ScreenPaddingSettingCard(
            FIF.FULL_SCREEN,
            tr("Screen Margin"),
            tr("Screen Margin Desc"),
            parent=personalContent
        )
        
        self.timerPosCard = OptionsSettingCard(
            cfg.timerPosition,
            FIF.SPEED_HIGH, 
            tr("Timer Position"),
            tr("Timer Position Desc"),
            texts=[tr("Center"), tr("Top Left"), tr("Top Right"), tr("Bottom Left"), tr("Bottom Right")],
            parent=personalContent
        )
        
        personalLayout.addWidget(self.navPosCard)
        personalLayout.addWidget(self.componentScaleCard)
        personalLayout.addWidget(self.paddingCard)
        personalLayout.addWidget(self.timerPosCard)

        self.clockPageTitle = QLabel(tr("Clock"), clockContent)
        self.clockPageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        clockLayout.addWidget(self.clockPageTitle)

        self.clockHeader = QLabel(tr("Clock Configuration"), clockContent)
        self.clockHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        clockLayout.addWidget(self.clockHeader)

        self.clockEnableCard = SwitchSettingCard(
            FIF.DATE_TIME,
            tr("Show Clock"),
            tr("Show Clock Desc"),
            configItem=cfg.enableClock,
            parent=clockContent
        )

        self.clockPosCard = OptionsSettingCard(
            cfg.clockPosition,
            FIF.HISTORY,
            tr("Clock Position"),
            tr("Clock Position Desc"),
            texts=[tr("Top Left"), tr("Top Right"), tr("Bottom Left"), tr("Bottom Right")],
            parent=clockContent
        )

        self.clockSettingCard = ClockSettingCard(
            FIF.DATE_TIME,
            tr("Clock Style"),
            tr("Clock Style Desc"),
            parent=clockContent
        )

        clockLayout.addWidget(self.clockEnableCard)
        clockLayout.addWidget(self.clockPosCard)
        clockLayout.addWidget(self.clockSettingCard)

        self._update_clock_settings_for_cicw()
        
        # --- Debug Page Setup ---
        self.debugPageTitle = QLabel(tr("Debug"), debugContent)
        self.debugPageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        debugLayout.addWidget(self.debugPageTitle)
        
        self.debugDangerHeader = QLabel(tr("Danger Zone"), debugContent)
        self.debugDangerHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        debugLayout.addWidget(self.debugDangerHeader)
        
        # Move crash card to debug page
        self.crashCard = PushSettingCard(
            "触发",
            FIF.DELETE,
            "崩溃测试",
            "仅用于开发调试，请勿在演示时点击",
            parent=debugContent
        )
        debugLayout.addWidget(self.crashCard)
        
        self.logHeader = QLabel(tr("Logs"), debugContent)
        self.logHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        debugLayout.addWidget(self.logHeader)

        self.logBtnWidget = QWidget(debugContent)
        self.logBtnLayout = QHBoxLayout(self.logBtnWidget)
        self.testLogBtn = PushButton("发送测试日志", debugContent, FIF.SEND)
        self.clearLogBtn = PushButton("清空日志", debugContent, FIF.DELETE)
        self.logBtnLayout.addWidget(self.testLogBtn)
        self.logBtnLayout.addWidget(self.clearLogBtn)
        self.logBtnLayout.addStretch()
        debugLayout.addWidget(self.logBtnWidget)

        self.logTextEdit = TextEdit(debugContent)
        self.logTextEdit.setReadOnly(True)
        self.logTextEdit.setPlaceholderText("等待日志输出...")
        self.logTextEdit.setFixedHeight(400)
        debugLayout.addWidget(self.logTextEdit)
        
        logger.log_signal.connect(self._append_log)
        self.testLogBtn.clicked.connect(lambda: log("这是一条测试日志"))
        self.clearLogBtn.clicked.connect(lambda: self.logTextEdit.clear())
        
        self._load_existing_logs()
        log("调试界面已加载")

        self.updatePageTitle = QLabel(tr("Update"), updateContent)
        self.updatePageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        updateLayout.addWidget(self.updatePageTitle)

        self.updateHeader = QLabel(tr("Update Status"), updateContent)
        self.updateHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        updateLayout.addWidget(self.updateHeader)
        
        version_name = "Unknown"
        version_code = None
        code_name = ""
        build_date_display = ""
        try:
            import sys
            base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
            v_path = os.path.join(base_dir, "config", "version.json")
            with open(v_path, 'r', encoding='utf-8') as f:
                info = json.load(f)
                version_name = info.get("versionName", "Unknown")
                version_code = info.get("versionCode", None)
                code_name = info.get("codeName", "")
                raw_date = info.get("buildDate", "")
                if isinstance(raw_date, str) and raw_date:
                    parts = raw_date.split("T", 1)
                    if len(parts) == 2:
                        time_part = parts[1].split(".", 1)[0]
                        build_date_display = parts[0] + " " + time_part[:5]
                    else:
                        build_date_display = raw_date
        except:
            pass
        self.currentVersion = version_name
        self.currentVersionCode = version_code
        self.currentCodeName = code_name
        self.currentBuildDate = build_date_display

        # 在标题栏插入版本号
        self.titleBarVersionLabel = QLabel(f"v{self.currentVersion}", self.titleBar)
        self.titleBarVersionLabel.setStyleSheet("font-size: 12px; color: rgba(128, 128, 128, 0.7); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        
        # 寻找三键组合的起始位置（通常在最后一个伸缩条之后）
        layout = self.titleBar.layout()
        insert_index = layout.count() - 1
        for i in range(layout.count() - 1, -1, -1):
            item = layout.itemAt(i)
            if item.widget() and item.widget().__class__.__name__ in ['TitleBarButton', 'MinimizeButton', 'MaximizeButton', 'CloseButton']:
                insert_index = i
            else:
                break
        
        layout.insertWidget(insert_index, self.titleBarVersionLabel, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.insertSpacing(insert_index + 1, 12)

        self.updateManagerWidget = UpdateManagerWidget(version_name, code_name, build_date_display, updateContent)
        updateLayout.addWidget(self.updateManagerWidget)
        
        self.aboutPageTitle = QLabel(tr("About"), aboutContent)
        self.aboutPageTitle.setStyleSheet("font-size: 28px; font-weight: bold; margin-bottom: 12px; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        aboutLayout.addWidget(self.aboutPageTitle)

        self.aboutHeader = QLabel(tr("Software Information"), aboutContent)
        self.aboutHeader.setStyleSheet("font-size: 18px; font-weight: bold; margin-top: 20px; margin-bottom: 10px; font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        aboutLayout.addWidget(self.aboutHeader)

        self.aboutContentWidget = QWidget(aboutContent)
        aboutContentLayout = QVBoxLayout(self.aboutContentWidget)
        aboutContentLayout.setContentsMargins(24, 24, 24, 24)
        aboutContentLayout.setSpacing(20)

        try:
            import sys
            banner_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            banner_path = os.path.join(banner_base, "resources", "icons", "banner.png")
            if os.path.exists(banner_path):
                self.aboutBanner = ImageLabel(banner_path, self.aboutContentWidget)
                self.aboutBanner.setBorderRadius(8, 8, 8, 8)
                self.aboutBanner.scaledToHeight(300)
                aboutContentLayout.addWidget(self.aboutBanner)
        except Exception:
            pass

        self.aboutTitle = QLabel("Kazuha", self.aboutContentWidget)
        self.aboutTitle.setStyleSheet("font-size: 32px; font-weight: bold; font-family: 'Bahnschrift', 'Segoe UI Variable Display', 'Segoe UI', 'Microsoft YaHei';")
        aboutContentLayout.addWidget(self.aboutTitle)

        self.aboutSubtitle = QLabel(tr("App Description"), self.aboutContentWidget)
        self.aboutSubtitle.setStyleSheet("font-size: 16px; font-weight: normal; color: rgba(128, 128, 128, 0.8); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        self.aboutSubtitle.setWordWrap(True)
        aboutContentLayout.addWidget(self.aboutSubtitle)

        infoLayout = QHBoxLayout()
        infoLeft = QVBoxLayout()
        infoRight = QVBoxLayout()

        version_text = f"{tr('Current Version')}: {self.currentVersion}" if getattr(self, "currentVersion", None) else f"{tr('Current Version')}: {tr('Unknown')}"
        self.aboutVersionLabel = QLabel(version_text, self.aboutContentWidget)
        self.aboutVersionLabel.setStyleSheet("font-size: 14px; font-weight: normal; color: rgba(128, 128, 128, 0.9); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
        infoLeft.addWidget(self.aboutVersionLabel)

        if getattr(self, "currentCodeName", ""):
            self.aboutCodeNameLabel = QLabel(f"{tr('Code Name')}: {self.currentCodeName}", self.aboutContentWidget)
            self.aboutCodeNameLabel.setStyleSheet("font-size: 14px; font-weight: normal; color: rgba(128, 128, 128, 0.9); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
            infoLeft.addWidget(self.aboutCodeNameLabel)

        if getattr(self, "currentVersionCode", None) is not None:
            self.aboutVersionCodeLabel = QLabel(f"{tr('Version Code')}: {self.currentVersionCode}", self.aboutContentWidget)
            self.aboutVersionCodeLabel.setStyleSheet("font-size: 14px; font-weight: normal; color: rgba(128, 128, 128, 0.9); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
            infoLeft.addWidget(self.aboutVersionCodeLabel)

        if getattr(self, "currentBuildDate", ""):
            self.aboutBuildDateLabel = QLabel(f"{tr('Build Date')}: {self.currentBuildDate}", self.aboutContentWidget)
            self.aboutBuildDateLabel.setStyleSheet("font-size: 14px; font-weight: normal; color: rgba(128, 128, 128, 0.9); font-family: 'Bahnschrift', 'Segoe UI', 'Microsoft YaHei';")
            infoRight.addWidget(self.aboutBuildDateLabel)

        infoLayout.addLayout(infoLeft)
        infoLayout.addSpacing(32)
        infoLayout.addLayout(infoRight)
        infoLayout.addStretch()
        aboutContentLayout.addLayout(infoLayout)

        buttonRow = QHBoxLayout()
        self.githubBtn = PrimaryPushButton(tr("GitHub Repo"), self.aboutContentWidget)
        self.githubBtn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/Haraguse/Kazuha")))
        self.issueBtn = PushButton(tr("Feedback"), self.aboutContentWidget)
        self.issueBtn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/Haraguse/Kazuha/issues")))
        buttonRow.addWidget(self.githubBtn)
        buttonRow.addWidget(self.issueBtn)
        buttonRow.addStretch()
        aboutContentLayout.addLayout(buttonRow)
        
        aboutLayout.addWidget(self.aboutContentWidget)

        self.updateManagerWidget.checkUpdateClicked.connect(self._on_update_card_clicked)
        QTimer.singleShot(0, self._auto_check_update)
        
        self.crashCard.clicked.connect(self.show_crash_dialog)
        
        cfg.themeMode.valueChanged.connect(self.on_config_changed)
        cfg.navPosition.valueChanged.connect(self.on_config_changed)
        cfg.clockPosition.valueChanged.connect(self.on_config_changed)
        cfg.clockFontWeight.valueChanged.connect(self.on_config_changed)
        cfg.clockShowSeconds.valueChanged.connect(self.on_config_changed)
        cfg.clockShowDate.valueChanged.connect(self.on_config_changed)
        cfg.clockShowLunar.valueChanged.connect(self.on_config_changed)
        cfg.timerPosition.valueChanged.connect(self.on_config_changed)
        cfg.enableStartUp.valueChanged.connect(self.on_config_changed)
        cfg.enableClock.valueChanged.connect(self.on_config_changed)
        cfg.enableSystemNotification.valueChanged.connect(self.on_config_changed)
        cfg.enableGlobalSound.valueChanged.connect(self.on_config_changed)
        cfg.enableGlobalAnimation.valueChanged.connect(self.on_config_changed)
        cfg.screenPadding.valueChanged.connect(self.on_config_changed)
        cfg.componentScale.valueChanged.connect(self._on_component_scale_changed)

        self.addSubInterface(self.generalInterface, FIF.SETTING, tr("General"))
        self.addSubInterface(self.personalInterface, FIF.BRUSH, tr("Personalization"))
        self.addSubInterface(self.clockInterface, FIF.DATE_TIME, tr("Clock"))
        self.addSubInterface(self.updateInterface, FIF.SYNC, tr("Update"))
        self.aboutNavItem = self.addSubInterface(self.aboutInterface, FIF.INFO, tr("About"), NavigationItemPosition.BOTTOM)
        
        # Setup click tracking for About navigation item
        self._about_click_count = 0
        self._debug_unlocked = False
        
        # Use an event filter on the navigation panel to detect clicks on the About item
        # In QFluentWindow, the navigation interface is accessible via self.navigationInterface
        # but sometimes it's initialized later. We'll use self.stackedWidget.currentChanged 
        # as a fallback and combine it with a more robust check.
        self.stackedWidget.currentChanged.connect(self._on_page_changed)
        
        # Apply initial theme
        self.set_theme(cfg.themeMode.value)

    def _update_theme_color_cards_visibility(self):
        mode = cfg.themeColorMode.value
        self.themeColorCustomCard.setVisible(mode == "Custom")
        self.themeColorIntervalCard.setVisible(mode in ["Screen", "Wallpaper"])

    def _on_component_scale_changed(self, value):
        self.on_config_changed()
        if self._scale_preview_timer.isActive():
            self._scale_preview_timer.stop()
        self._scale_preview_timer.start(220)

    def _trigger_scale_preview(self):
        controller = BusinessLogicController.instance()
        if controller is None:
            return
        controller.show_layout_preview()

    def _on_page_changed(self, index):
        # This will trigger when the page actually changes
        current_widget = self.stackedWidget.widget(index)
        if current_widget is self.aboutInterface:
             self._on_about_clicked()

    def _on_about_clicked(self):
        if self._debug_unlocked:
            return
            
        self._about_click_count += 1
        if self._about_click_count >= 5:
            self._unlock_debug_mode()

    def _unlock_debug_mode(self):
        self._debug_unlocked = True
        self.addSubInterface(self.debugInterface, FIF.DEVELOPER_TOOLS, tr("Debug"), NavigationItemPosition.BOTTOM)
        
        InfoBar.success(
            title="开发者模式已开启",
            content="调试选项已在侧边栏显示",
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP_RIGHT,
            duration=2000,
            parent=self
        )

    def _load_existing_logs(self):
        try:
            from controllers.business_logic import get_app_base_dir
            log_path = get_app_base_dir() / "debug.log"
            if log_path.exists():
                with open(log_path, "r", encoding='utf-8') as f:
                    # Load last 100 lines to avoid freezing
                    lines = f.readlines()
                    self.logTextEdit.setPlainText("".join(lines[-100:]))
                    # Scroll to bottom
                    self._append_log("") 
        except Exception as e:
            self.logTextEdit.append(f"加载历史日志失败: {e}")

    def _append_log(self, msg):
        self.logTextEdit.append(msg)
        # Scroll to bottom
        cursor = self.logTextEdit.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.logTextEdit.setTextCursor(cursor)
        
    def on_config_changed(self):
        self.configChanged.emit()
        self._show_restart_infobar()
        # Update native labels theme immediately
        self.set_theme(cfg.themeMode.value)

    def _show_restart_infobar(self):
        from controllers.business_logic import BusinessLogicController
        if getattr(self, "_restart_required", False):
            return
        self._restart_required = True
        if not hasattr(self, "restartButton"):
            self.restartButton = PushButton(tr("Restart Now"), self.titleBar)
            self.restartButton.clicked.connect(BusinessLogicController.instance().restart_app)
            self.restartButton.setObjectName("restartRequiredButton")
            layout = self.titleBar.layout()
            insert_index = layout.count() - 1
            for i in range(layout.count() - 1, -1, -1):
                item = layout.itemAt(i)
                if item.widget() and item.widget().__class__.__name__ in ['TitleBarButton', 'MinimizeButton', 'MaximizeButton', 'CloseButton']:
                    insert_index = i
                else:
                    break
            layout.insertWidget(insert_index, self.restartButton, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            layout.insertSpacing(insert_index + 1, 8)
        self.restartButton.show()
        
    def show_crash_dialog(self):
        w = CrashDialog(self)
        if w.exec():
            settings = w.get_settings()
            if settings['countdown']:
                QTimer.singleShot(3000, lambda: trigger_crash(settings))
            else:
                trigger_crash(settings)
    
    def _update_clock_settings_for_cicw(self):
        running = False
        try:
            import subprocess
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            output = subprocess.check_output('tasklist', startupinfo=startupinfo).decode('gbk', errors='ignore').lower()
            running = ('classisland' in output) or ('classwidgets' in output)
        except Exception:
            running = False
        
        widgets = [
            getattr(self, "clockEnableCard", None),
            getattr(self, "clockPosCard", None),
            getattr(self, "clockSettingCard", None),
        ]
        for w in widgets:
            if w is not None:
                w.setEnabled(not running)
        
        if running:
            if not self.clockConflictInfoBar:
                bar = InfoBar.warning(
                    title='检测到 ClassIsland/Class Widgets 正在运行',
                    content="ClassIsland/Class Widgets 一部分具有和时钟组件相同的功能，且另一部分功能甚至可以超出时钟组件所能做到的范围。\n故此，时钟组件现在是不可用的。",
                    orient=Qt.Orientation.Horizontal,
                    isClosable=False,
                    position=InfoBarPosition.BOTTOM,
                    duration=-1,
                    parent=self.clockInterface
                )
                for label in bar.findChildren(FluentLabelBase):
                    label.setWordWrap(True)
                self.clockConflictInfoBar = bar
                bar.show()
            else:
                self.clockConflictInfoBar.show()
        else:
            if self.clockConflictInfoBar:
                self.clockConflictInfoBar.hide()
                
    def set_theme(self, theme):
        if theme == Theme.AUTO:
            theme = Theme.DARK if isDarkTheme() else Theme.LIGHT
        
        text_color = "white" if theme == Theme.DARK else "black"
        header_color = "rgba(255, 255, 255, 0.9)" if theme == Theme.DARK else "rgba(0, 0, 0, 0.85)"
        subtext_color = "rgba(255, 255, 255, 0.7)" if theme == Theme.DARK else "rgba(0, 0, 0, 0.6)"
        
        # List of title labels
        titles = [
            self.generalPageTitle, self.personalPageTitle, self.clockPageTitle,
            self.updatePageTitle, self.aboutPageTitle, self.debugPageTitle,
            self.aboutTitle, self.updateManagerWidget.statusLabel
        ]
        for title in titles:
            if title:
                style = title.styleSheet()
                import re
                style = re.sub(r'color\s*:\s*[^;]+;', '', style)
                title.setStyleSheet(style + f" color: {text_color};")
                
        # List of header labels
        headers = [
            self.generalHeader, self.themeHeader, self.layoutHeader,
            self.clockHeader, self.debugDangerHeader, self.logHeader,
            self.updateHeader, self.aboutHeader
        ]
        for header in headers:
            if header:
                style = header.styleSheet()
                import re
                style = re.sub(r'color\s*:\s*[^;]+;', '', style)
                header.setStyleSheet(style + f" color: {header_color};")

        # List of subtext labels
        subtexts = [
            self.aboutSubtitle, self.aboutVersionLabel, 
            self.aboutCodeNameLabel if hasattr(self, 'aboutCodeNameLabel') else None,
            self.aboutVersionCodeLabel if hasattr(self, 'aboutVersionCodeLabel') else None,
            self.aboutBuildDateLabel if hasattr(self, 'aboutBuildDateLabel') else None,
            self.updateManagerWidget.versionLabel,
            self.updateManagerWidget.logContent,
            self.titleBarVersionLabel
        ]
        for subtext in subtexts:
            if subtext:
                style = subtext.styleSheet()
                import re
                style = re.sub(r'color\s*:\s*[^;]+;', '', style)
                subtext.setStyleSheet(style + f" color: {subtext_color};")

        # Update custom components
        from ui.settings.custom_settings import SchematicOptionButton
        for btn in self.findChildren(SchematicOptionButton):
            btn.set_theme()

    def _on_update_card_clicked(self):
        self.checkUpdateClicked.emit()

    def _auto_check_update(self):
        # 仅触发信号，UI 状态由 UpdateManagerWidget 内部处理
        self._on_update_card_clicked()
