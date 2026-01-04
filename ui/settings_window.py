from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QUrl, QCoreApplication, QSize
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFrame, QHBoxLayout, QSizePolicy, QApplication
from PyQt6.QtGui import QFont, QPainter, QColor, QLinearGradient, QDesktopServices, QIcon
import json
import os
import datetime
import subprocess
import json
import os
import datetime
from qfluentwidgets import (
    SwitchSettingCard, OptionsSettingCard, PushSettingCard,
    PrimaryPushSettingCard, SettingCard, PrimaryPushButton, PushButton,
    SmoothScrollArea, ExpandLayout, Theme, setTheme, setThemeColor,
    FluentWindow, NavigationItemPosition, isDarkTheme,
    LargeTitleLabel, TitleLabel, BodyLabel, CaptionLabel, IndeterminateProgressRing,
    InfoBar, InfoBarPosition, MessageBoxBase, SubtitleLabel, ProgressBar,
    SegmentedWidget, Pivot, FluentIcon as FIF
)
from ui.custom_settings import SchematicOptionsSettingCard, ScreenPaddingSettingCard
from ui.crash_dialog import CrashDialog, trigger_crash
from ui.visual_settings import ClockSettingCard
from ui.about_page import AboutPage


def tr(text: str):
    return QCoreApplication.translate("SettingsWindow", text)


class UpdateDialog(MessageBoxBase):
    def __init__(self, version_info, parent=None):
        super().__init__(parent)
        self.version_info = version_info
        
        self.titleLabel = SubtitleLabel(tr("更新"), self)
        self.titleLabel.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 24px;")
        
        self.statusWidget = QWidget(self)
        self.statusLayout = QHBoxLayout(self.statusWidget)
        self.statusLayout.setContentsMargins(0, 0, 0, 0)
        self.statusLayout.setSpacing(16)
        
        self.iconLabel = QLabel(self)
        self.iconLabel.setFixedSize(64, 64)
        icon_pixmap = FIF.SYNC.icon().pixmap(QSize(64, 64))
        self.iconLabel.setPixmap(icon_pixmap)
        
        self.textLayout = QVBoxLayout()
        self.textLayout.setSpacing(2)
        name = version_info.get('name', '')
        self.updateStatusLabel = BodyLabel(tr("有可用更新: {name}").format(name=name), self)
        self.updateStatusLabel.setStyleSheet("font-size: 16px; font-weight: bold;")
        
        now_str = datetime.datetime.now().strftime("%H:%M")
        self.lastCheckLabel = CaptionLabel(tr("上次检查时间: 今天, {time}").format(time=now_str), self)
        self.lastCheckLabel.setStyleSheet("color: gray;")
        
        self.textLayout.addWidget(self.updateStatusLabel)
        self.textLayout.addWidget(self.lastCheckLabel)
        self.textLayout.addStretch(1)
        
        self.actionButton = PrimaryPushButton(tr("立即更新"), self)
        self.actionButton.setFixedWidth(120)
        self.actionButton.clicked.connect(self.accept)
        
        self.statusLayout.addWidget(self.iconLabel)
        self.statusLayout.addLayout(self.textLayout)
        self.statusLayout.addStretch(1)
        self.statusLayout.addWidget(self.actionButton)
        
        self.pivot = Pivot(self)
        self.pivot.setStyleSheet("""
            Pivot { background-color: transparent; }
            PivotItem { font-size: 14px; padding: 10px 0; margin-right: 20px; }
        """)
        
        self.logWidget = QWidget()
        self.logLayout = QVBoxLayout(self.logWidget)
        self.logLayout.setContentsMargins(0, 10, 0, 0)
        self.logEdit = BodyLabel(version_info.get('body', ''), self)
        self.logEdit.setWordWrap(True)
        self.logEdit.setStyleSheet("color: gray; font-size: 13px; line-height: 1.5;")
        self.logLayout.addWidget(self.logEdit)
        self.logLayout.addStretch(1)
        
        self.settingsWidget = QWidget()
        self.settingsLayout = QVBoxLayout(self.settingsWidget)
        self.settingsLayout.setContentsMargins(0, 10, 0, 0)
        self.autoInstallCard = SwitchSettingCard(FIF.UPDATE, tr("自动安装更新"), tr("下载完成后立即执行静默安装"), parent=self.settingsWidget)
        self.settingsLayout.addWidget(self.autoInstallCard)
        self.settingsLayout.addStretch(1)
        
        self.pivot.addItem("logs", tr("更新日志"), lambda: self.show_tab(0))
        self.pivot.addItem("settings", tr("更新设置"), lambda: self.show_tab(1))
        
        self.progressBar = ProgressBar(self)
        self.progressBar.hide()
        self.statusLabel = CaptionLabel("", self)
        self.statusLabel.hide()
        
        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addWidget(self.statusWidget)
        self.viewLayout.addSpacing(20)
        self.viewLayout.addWidget(self.pivot)
        self.viewLayout.addWidget(self.logWidget)
        self.viewLayout.addWidget(self.settingsWidget)
        self.viewLayout.addWidget(self.progressBar)
        self.viewLayout.addWidget(self.statusLabel)
        
        self.settingsWidget.hide()
        
        self.yesButton.hide()
        self.cancelButton.setText(tr("以后再说"))
        
        self.widget.setMinimumWidth(600)
        self.widget.setMinimumHeight(500)

    def show_tab(self, index):
        if index == 0:
            self.logWidget.show()
            self.settingsWidget.hide()
        else:
            self.logWidget.hide()
            self.settingsWidget.show()

    def set_progress(self, val):
        self.progressBar.show()
        self.statusLabel.show()
        self.progressBar.setValue(val)
        self.statusLabel.setText(tr("正在下载: {0}%").format(val))
        self.actionButton.setEnabled(False)

from ui.custom_settings import SchematicOptionsSettingCard, ScreenPaddingSettingCard
from controllers.business_logic import cfg, get_app_base_dir

def _create_page(parent: QWidget):
    page = QWidget(parent)
    page.setStyleSheet("background-color: transparent;")
    layout = QVBoxLayout(page)
    layout.setContentsMargins(24, 10, 24, 24)
    layout.setSpacing(0)

    scroll = SmoothScrollArea(page)
    scroll.setObjectName("scrollInterface")
    scroll.setStyleSheet("SmoothScrollArea { background-color: transparent; border: none; }")
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    content = QWidget()
    content.setStyleSheet("background-color: transparent;")
    content.setObjectName("scrollWidget")
    expand_layout = ExpandLayout(content)

    scroll.setWidget(content)
    scroll.setWidgetResizable(True)

    layout.addWidget(scroll)
    return page, content, expand_layout


def _apply_title_style(label: QLabel):
    f = label.font()
    f.setPointSize(22)
    f.setWeight(QFont.Weight.Bold)
    label.setFont(f)


def _apply_body_strong_style(label: QLabel):
    f = label.font()
    f.setPointSize(12)
    f.setWeight(QFont.Weight.Normal)
    label.setFont(f)


class VersionInfoCard(SettingCard):
    def __init__(self, icon, title, parent=None):
        super().__init__(icon, title, "", parent)
        box = QVBoxLayout()
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)
        self.latestLabel = BodyLabel(tr("最新可用 Release 版本：尚未执行检查"), self)
        self.currentLabel = BodyLabel(tr("当前正在运行的版本：未知"), self)
        self.latestLabel.setWordWrap(True)
        self.currentLabel.setWordWrap(True)
        box.addWidget(self.latestLabel)
        box.addWidget(self.currentLabel)
        self.hBoxLayout.addLayout(box, 1)

    def set_versions(self, latest, current):
        self.latestLabel.setText(latest)
        self.currentLabel.setText(current)


class UpdateLogCard(SettingCard):
    def __init__(self, icon, title, parent=None):
        super().__init__(icon, title, "", parent)
        box = QVBoxLayout()
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)
        self.latestLogLabel = BodyLabel(tr("最新 Release 版本更新日志：尚未加载"), self)
        self.currentLogLabel = BodyLabel(tr("当前版本对应的更新日志：尚未加载"), self)
        self.latestLogLabel.setWordWrap(True)
        self.currentLogLabel.setWordWrap(True)
        box.addWidget(self.latestLogLabel)
        box.addWidget(self.currentLogLabel)
        self.hBoxLayout.addLayout(box, 1)

    def set_logs(self, latest, current):
        self.latestLogLabel.setText(latest)
        self.currentLogLabel.setText(current)


class SettingsWindow(FluentWindow):
    configChanged = pyqtSignal()
    checkUpdateClicked = pyqtSignal()
    startUpdateClicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        base_dir = get_app_base_dir()
        icon_path = base_dir / "resources" / "icons" / "trayicon.svg"
        self.setWindowIcon(QIcon(str(icon_path)))
        self.setWindowTitle(tr("界面设置 - Kazuha"))
        self.setMinimumWidth(1153)
        self.resize(900, 640)

        try:
            self.setMicaEffectEnabled(True)
        except Exception:
            pass

        font = QFont("Bahnschrift", 14)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        self.setFont(font)

        self.generalInterface, generalContent, generalLayout = _create_page(self)
        self.generalInterface.setObjectName("settings-general")
        
        self.generalPageTitle = LargeTitleLabel(tr("常规"), generalContent)
        _apply_title_style(self.generalPageTitle)
        generalLayout.addWidget(self.generalPageTitle)

        self.startupCard = SwitchSettingCard(
            FIF.INFO,
            tr("开机自启动"),
            tr("在 Windows 登录时自动启动应用"),
            configItem=cfg.enableStartUp,
            parent=generalContent
        )
        self.notificationCard = SwitchSettingCard(
            FIF.FEEDBACK,
            tr("系统通知"),
            tr("允许应用向系统发送状态通知"),
            configItem=cfg.enableSystemNotification,
            parent=generalContent
        )
        self.soundCard = SwitchSettingCard(
            FIF.MUSIC,
            tr("全局音效"),
            tr("在操作时播放交互音效"),
            configItem=cfg.enableGlobalSound,
            parent=generalContent
        )
        self.animationCard = SwitchSettingCard(
            FIF.MOVE,
            tr("全局动画"),
            tr("启用界面过渡和状态切换动画"),
            configItem=cfg.enableGlobalAnimation,
            parent=generalContent
        )
        self.languageCard = OptionsSettingCard(
            cfg.language,
            FIF.LANGUAGE,
            tr("界面语言"),
            tr("设置应用界面的显示语言（需要重启）"),
            texts=[tr("简体中文"), tr("繁體中文"), tr("English"), tr("日本語"), tr("Bod-yig")],
            parent=generalContent
        )
        
        generalLayout.addWidget(self.startupCard)
        generalLayout.addWidget(self.notificationCard)
        generalLayout.addWidget(self.soundCard)
        generalLayout.addWidget(self.animationCard)
        generalLayout.addWidget(self.languageCard)

        self.personalInterface, personalContent, personalLayout = _create_page(self)
        self.personalInterface.setObjectName("settings-personal")
        
        self.personalPageTitle = LargeTitleLabel(tr("个性化"), personalContent)
        _apply_title_style(self.personalPageTitle)
        personalLayout.addWidget(self.personalPageTitle)
        
        self.themeCard = OptionsSettingCard(
            cfg.themeMode,
            FIF.BRUSH,
            tr("应用主题"),
            tr("切换应用的主题颜色模式"),
            texts=[tr("浅色"), tr("深色"), tr("跟随系统")],
            parent=personalContent
        )
        self.navPosCard = OptionsSettingCard(
            cfg.navPosition,
            FIF.MENU,
            tr("导航栏位置"),
            tr("调整侧边栏导航项目的显示位置"),
            texts=[tr("底部两侧"), tr("居中两侧")],
            parent=personalContent
        )
        self.paddingCard = OptionsSettingCard(
            cfg.screenPadding,
            FIF.LAYOUT,
            tr("屏幕边距"),
            tr("调整应用窗口与屏幕边缘的间距"),
            texts=[tr("较小"), tr("正常"), tr("较大")],
            parent=personalContent
        )
        
        personalLayout.addWidget(self.themeCard)
        personalLayout.addWidget(self.navPosCard)
        personalLayout.addWidget(self.paddingCard)

        self.clockInterface, clockContent, clockLayout = _create_page(self)
        self.clockInterface.setObjectName("settings-clock")
        self.clockLayout = clockLayout
        self.clockConflictInfoBar = None

        self.clockPageTitle = LargeTitleLabel(tr("桌面悬浮时钟组件"), clockContent)
        _apply_title_style(self.clockPageTitle)
        clockLayout.addWidget(self.clockPageTitle)

        self.clockHeader = TitleLabel(tr("显示内容与样式"), clockContent)
        _apply_body_strong_style(self.clockHeader)
        clockLayout.addWidget(self.clockHeader)

        self.clockEnableCard = SwitchSettingCard(
            FIF.DATE_TIME,
            tr("显示悬浮时钟"),
            tr("在屏幕上显示一个可自定义的悬浮时钟窗口"),
            configItem=cfg.enableClock,
            parent=clockContent
        )

        self.clockPosCard = OptionsSettingCard(
            cfg.clockPosition,
            FIF.HISTORY,
            tr("时钟位置"),
            tr("调整桌面悬浮时钟在屏幕四角的显示位置"),
            texts=[tr("左上角"), tr("右上角"), tr("左下角"), tr("右下角")],
            parent=clockContent
        )

        self.clockSettingCard = ClockSettingCard(
            FIF.DATE_TIME,
            tr("时钟样式"),
            tr("自定义悬浮时钟显示的内容、字体和粗细等外观"),
            parent=clockContent
        )

        clockLayout.addWidget(self.clockEnableCard)
        clockLayout.addWidget(self.clockPosCard)
        clockLayout.addWidget(self.clockSettingCard)

        self._update_clock_settings_for_cicw()

        self.debugInterface, debugContent, debugLayout = _create_page(self)
        self.debugInterface.setObjectName("settings-debug")
        
        self.dangerHeader = TitleLabel(tr("危险功能（仅限测试）"), debugContent)
        _apply_body_strong_style(self.dangerHeader)
        debugLayout.addWidget(self.dangerHeader)
        self.crashCard = PushSettingCard(
            tr("触发"),
            FIF.DELETE,
            tr("崩溃测试"),
            tr("仅用于开发调试环境，请勿在正式课堂或演示时点击"),
            parent=debugContent
        )
        debugLayout.addWidget(self.crashCard)

        self.aboutInterface = AboutPage(self)

        self.updateInterface, updateContent, updateLayout = self._create_update_interface()

        version = "Unknown"
        try:
            base_dir = get_app_base_dir()
            v_path = base_dir / "config" / "version.json"
            with open(v_path, "r", encoding="utf-8") as f:
                info = json.load(f)
                version = info.get("versionName", "Unknown")
        except:
            pass

        self.currentVersion = version
        if hasattr(self, "titleBar"):
            version_text = version
            if version_text and version_text != "Unknown":
                version_text = f"v{version_text}"
                self.titleVersionLabel = CaptionLabel(version_text, self.titleBar)
                self.titleVersionLabel.setObjectName("titleVersionLabel")
                self.titleVersionLabel.setStyleSheet("padding: 0 8px;")
                layout = self.titleBar.hBoxLayout
                layout.insertWidget(layout.count() - 1, self.titleVersionLabel)

        self.addSubInterface(self.generalInterface, FIF.SETTING, tr("常规"))
        self.addSubInterface(self.personalInterface, FIF.BRUSH, tr("个性化"))
        self.addSubInterface(self.clockInterface, FIF.DATE_TIME, tr("时钟组件"))
        self.addSubInterface(self.updateInterface, FIF.SYNC, tr("更新"))
        self.addSubInterface(self.aboutInterface, FIF.INFO, "About", position=NavigationItemPosition.BOTTOM)

        self._restart_required = False
        self._debug_toggle_count = 0
        self._debug_unlocked = False
        self._last_toggle_widget = None
        self.stackedWidget.currentChanged.connect(self._on_stack_page_changed)
        
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
        cfg.language.valueChanged.connect(self.on_config_changed)
        cfg.language.valueChanged.connect(self._on_language_changed)

    def _create_update_interface(self):
        page, content, layout = _create_page(self)
        page.setObjectName("settings-update")
        
        title = LargeTitleLabel(tr("更新"), content)
        _apply_title_style(title)
        layout.addWidget(title)
        
        statusWidget = QWidget(content)
        statusLayout = QHBoxLayout(statusWidget)
        statusLayout.setContentsMargins(0, 0, 0, 0)
        statusLayout.setSpacing(20)
        
        self.updateStatusIcon = QLabel(statusWidget)
        self.updateStatusIcon.setFixedSize(64, 64)
        self.updateStatusIcon.setPixmap(FIF.COMPLETED.icon().pixmap(QSize(64, 64)))
        
        textLayout = QVBoxLayout()
        textLayout.setSpacing(4)
        self.updateStatusTitle = BodyLabel(tr("你使用的是最新版本"), statusWidget)
        self.updateStatusTitle.setStyleSheet("font-size: 18px; font-weight: bold;")
        
        now_str = datetime.datetime.now().strftime("%H:%M")
        self.updateLastCheckLabel = CaptionLabel(tr("上次检查时间: 今天, {time}").format(time=now_str), statusWidget)
        self.updateLastCheckLabel.setStyleSheet("color: gray;")
        
        textLayout.addWidget(self.updateStatusTitle)
        textLayout.addWidget(self.updateLastCheckLabel)
        textLayout.addStretch(1)
        
        self.checkUpdateButton = PushButton(tr("检查更新"), statusWidget)
        self.checkUpdateButton.setFixedWidth(120)
        self.checkUpdateButton.clicked.connect(self._on_update_card_clicked)
        
        statusLayout.addWidget(self.updateStatusIcon)
        statusLayout.addLayout(textLayout)
        statusLayout.addStretch(1)
        statusLayout.addWidget(self.checkUpdateButton)
        
        layout.addWidget(statusWidget)
        
        pivot = Pivot(content)
        pivot.setStyleSheet("""
            Pivot { background-color: transparent; }
            PivotItem { font-size: 14px; padding: 10px 0; margin-right: 20px; }
        """)
        
        logWidget = QWidget(content)
        logLayout = QVBoxLayout(logWidget)
        logLayout.setContentsMargins(0, 0, 0, 0)
        
        self.latestReleaseLogLabel = BodyLabel(tr("暂无更新日志信息"), logWidget)
        self.latestReleaseLogLabel.setWordWrap(True)
        self.latestReleaseLogLabel.setStyleSheet("color: gray; font-size: 13px; line-height: 1.5;")
        
        logLayout.addWidget(self.latestReleaseLogLabel)
        logLayout.addStretch(1)
        
        settingsWidget = QWidget(content)
        settingsLayout = QVBoxLayout(settingsWidget)
        settingsLayout.setContentsMargins(0, 0, 0, 0)
        
        self.autoCheckCard = SwitchSettingCard(
            FIF.SYNC,
            tr("自动检查更新"),
            tr("在启动时自动检查新版本"),
            parent=settingsWidget
        )
        self.autoCheckCard.setChecked(True)
        self.autoCheckCard.setEnabled(False)
        
        settingsLayout.addWidget(self.autoCheckCard)
        settingsLayout.addStretch(1)
        
        layout.addWidget(pivot)
        layout.addWidget(logWidget)
        layout.addWidget(settingsWidget)
        
        logWidget.show()
        settingsWidget.hide()
        
        pivot.addItem("logs", tr("更新日志"), lambda: [logWidget.show(), settingsWidget.hide()])
        pivot.addItem("settings", tr("更新设置"), lambda: [logWidget.hide(), settingsWidget.show()])
        
        self.updateLoadingWidget = QWidget(content)
        loadingLayout = QHBoxLayout(self.updateLoadingWidget)
        loadingLayout.setContentsMargins(0, 12, 0, 0)
        loadingLayout.setSpacing(8)
        self.updateRing = IndeterminateProgressRing(self.updateLoadingWidget)
        self.updateRing.setFixedSize(20, 20)
        self.updateRing.setStrokeWidth(3)
        self.updateLoadingLabel = BodyLabel(tr("正在从网络检查更新..."), self.updateLoadingWidget)
        loadingLayout.addWidget(self.updateRing)
        loadingLayout.addWidget(self.updateLoadingLabel)
        loadingLayout.addStretch()
        self.updateLoadingWidget.hide()
        layout.addWidget(self.updateLoadingWidget)

        return page, content, layout

    def set_update_info(self, latest_version, latest_log):
        v = latest_version or tr("未知")
        
        if latest_version and self.currentVersion and latest_version != self.currentVersion:
             self.updateStatusTitle.setText(tr("有可用更新: {version}").format(version=v))
             self.updateStatusIcon.setPixmap(FIF.SYNC.icon().pixmap(QSize(64, 64)))
             self.checkUpdateButton.setText(tr("立即下载"))
        else:
             self.updateStatusTitle.setText(tr("你使用的是最新版本"))
             self.updateStatusIcon.setPixmap(FIF.COMPLETED.icon().pixmap(QSize(64, 64)))
             self.checkUpdateButton.setText(tr("检查更新"))

        text = latest_log.strip() if latest_log else tr("无更新日志")
        if self.currentVersion and latest_version and self.currentVersion == latest_version:
             self.latestReleaseLogLabel.setText(tr("当前版本更新日志：\n{text}").format(text=text))
        else:
             self.latestReleaseLogLabel.setText(tr("最新 Release 版本更新日志：\n{text}").format(text=text))
             
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.hide()

    def on_config_changed(self):
        self.configChanged.emit()

    def _on_language_changed(self, value):
        self._set_restart_required(True)

    def _set_restart_required(self, required: bool):
        self._restart_required = required
        if required:
            if not hasattr(self, "restartButton"):
                btn = PushButton(tr("立即重启"), self.titleBar)
                btn.setObjectName("restartRequiredButton")
                btn.setFixedHeight(26)
                try:
                    btn.setIcon(FIF.SYNC.icon())
                    btn.setIconSize(QSize(14, 14))
                except Exception:
                    pass
                btn.setStyleSheet(
                    "#restartRequiredButton {"
                    " padding: 2px 10px;"
                    " border-radius: 6px;"
                    " background-color: rgba(165, 110, 60, 0.15);"
                    " border: 1px solid rgba(165, 110, 60, 0.4);"
                    "}"
                    "#restartRequiredButton:hover {"
                    " background-color: rgba(165, 110, 60, 0.25);"
                    "}"
                    "#restartRequiredButton:pressed {"
                    " background-color: rgba(165, 110, 60, 0.35);"
                    "}"
                )
                layout = self.titleBar.hBoxLayout
                layout.insertWidget(layout.count() - 1, btn)
                btn.clicked.connect(self._on_restart_button_clicked)
                self.restartButton = btn
            self.restartButton.show()
        else:
            if hasattr(self, "restartButton"):
                self.restartButton.hide()

    def _on_restart_button_clicked(self):
        try:
            from controllers.business_logic import BusinessLogicController
        except Exception:
            BusinessLogicController = None
        app = QApplication.instance()
        if app is not None and BusinessLogicController is not None:
            for w in app.topLevelWidgets():
                if isinstance(w, BusinessLogicController):
                    if hasattr(w, "soft_restart"):
                        w.soft_restart()
                    else:
                        w.restart_application()
                    return
        if BusinessLogicController is not None:
            controller = BusinessLogicController()
            if hasattr(controller, "soft_restart"):
                controller.soft_restart()
            else:
                controller.restart_application()

    def _on_stack_page_changed(self, index):
        widget = self.stackedWidget.widget(index)
        if widget not in (self.generalInterface, self.aboutInterface):
            self._last_toggle_widget = widget
            return
        if self._last_toggle_widget in (self.generalInterface, self.aboutInterface) and self._last_toggle_widget is not widget:
            if not self._debug_unlocked:
                self._debug_toggle_count += 1
                if self._debug_toggle_count >= 5:
                    self._unlock_debug_page()
        self._last_toggle_widget = widget

    def _unlock_debug_page(self):
        if self._debug_unlocked:
            return
        self.addSubInterface(self.debugInterface, FIF.DEVELOPER_TOOLS, tr("调试"), position=NavigationItemPosition.BOTTOM)
        self._debug_unlocked = True

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
                    title=tr("检测到 ClassIsland/Class Widgets 正在运行"),
                    content=tr("ClassIsland/Class Widgets 的部分功能与本应用的时钟组件存在重叠，且另一部分功能甚至可以超出时钟组件所能做到的范围。\n为避免冲突，当前时钟组件已被临时禁用。"),
                    orient=Qt.Orientation.Horizontal,
                    isClosable=False,
                    position=InfoBarPosition.BOTTOM,
                    duration=-1,
                    parent=self.clockInterface
                )
                for label in bar.findChildren(QLabel):
                    label.setWordWrap(True)
                self.clockConflictInfoBar = bar
                bar.show()
            else:
                self.clockConflictInfoBar.show()
        else:
            if self.clockConflictInfoBar:
                self.clockConflictInfoBar.hide()
                
    def set_theme(self, theme):
        pass

    def _on_update_card_clicked(self):
        if hasattr(self, "latestReleaseVersionLabel"):
            self.latestReleaseVersionLabel.setText(tr("最新可用 Release 版本：正在检查..."))
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.show()
        self.checkUpdateClicked.emit()

    def _auto_check_update(self):
        self._on_update_card_clicked()

    def stop_update_loading(self):
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.hide()
