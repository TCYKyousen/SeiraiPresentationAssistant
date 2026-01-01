import markdown
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QUrl, QCoreApplication, QSize
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFrame, QHBoxLayout, QSizePolicy, QApplication, QStackedWidget, QTextBrowser
from PyQt6.QtGui import QFont, QPainter, QColor, QLinearGradient, QDesktopServices, QIcon, QAbstractTextDocumentLayout
import json
import os
import datetime
import subprocess

from qfluentwidgets import (
    FluentIcon as FIF,
    SwitchSettingCard, OptionsSettingCard, PushSettingCard,
    PrimaryPushSettingCard, SettingCard, PrimaryPushButton, PushButton,
    SmoothScrollArea, ExpandLayout, Theme, setTheme, setThemeColor,
    MSFluentWindow, NavigationItemPosition, isDarkTheme,
    LargeTitleLabel, TitleLabel, BodyLabel, CaptionLabel, IndeterminateProgressRing,
    InfoBar, InfoBarPosition, MessageBoxBase, SubtitleLabel, ProgressBar,
    SegmentedWidget, Pivot, SimpleCardWidget
)
from ui.custom_settings import SchematicOptionsSettingCard, ScreenPaddingSettingCard
from controllers.business_logic import cfg, get_app_base_dir
from ui.crash_dialog import CrashDialog, trigger_crash
from ui.visual_settings import ClockSettingCard
from ui.about_page import AboutPage



def tr(text: str) -> str:
    return text


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


from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEngineProfile, QWebEnginePage

class AutoHeightTextBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setOpenExternalLinks(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setStyleSheet("background-color: transparent;")
        self.document().documentLayout().documentSizeChanged.connect(self._adjust_height)

    def _adjust_height(self):
        # Force layout to recalculate
        self.document().adjustSize()
        doc_height = self.document().size().height()
        
        # Ensure minimum height for visibility
        if doc_height < 100:
            doc_height = 100
            
        # Add some padding
        self.setFixedHeight(int(doc_height) + 30)
        
        # Force parent layout to update
        parent = self.parent()
        if parent and hasattr(parent, 'layout') and parent.layout():
            parent.layout().invalidate()
            parent.layout().activate()
            
        # If we have a container parent, also update that
        container = self.parent()
        if container and hasattr(container, 'parent') and container.parent():
            parent_container = container.parent()
            if parent_container and hasattr(parent_container, 'layout') and parent_container.layout():
                parent_container.layout().invalidate()
                parent_container.layout().activate()
        
    def setMarkdown(self, text: str):
        # User requested plain text fallback if markdown fails
        # Let's try to be robust: Just show the text.
        # We can still try basic markdown but if it's causing issues, plain text is safer.
        # But wait, user said "plain text is fine".
        # Let's use setPlainText to be 100% sure content is there.
        # But we still want clickable links if possible.
        # QTextBrowser automatically detects links in plain text usually? No.
        
        # Let's try setHtml with simple <pre> or just setPlainText.
        # Given the frustration, setPlainText is the safest bet to prove data exists.
        
        # However, we can do a hybrid: 
        # Convert newlines to <br> and setHtml, so at least it wraps?
        # No, setPlainText handles wrapping if lineWrapMode is on (default).
        
        self.setPlainText(text)
        self._update_style()
        
        # Force immediate height adjustment
        self.document().adjustSize()
        self._adjust_height()

    def _update_style(self):
        color = "white" if isDarkTheme() else "black"
        # minimal style
        self.setStyleSheet(f"""
            QTextBrowser {{
                background-color: transparent;
                color: {color};
                border: none;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
                font-size: 14px;
            }}
        """)

    def showEvent(self, event):
        super().showEvent(event)
        self._update_style()
        self._adjust_height()


class SettingsWindow(MSFluentWindow):
    configChanged = pyqtSignal()
    checkUpdateClicked = pyqtSignal()
    startUpdateClicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_update_available = False
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
        
        self.generalTitle = LargeTitleLabel(tr("常规设置"), generalContent)
        _apply_title_style(self.generalTitle)
        generalLayout.addWidget(self.generalTitle)
        
        self.startUpCard = SwitchSettingCard(
            FIF.POWER_BUTTON,
            tr("开机自启"),
            tr("在系统启动时自动运行 Kazuha"),
            configItem=cfg.enableStartUp,
            parent=generalContent
        )
        self.notificationCard = SwitchSettingCard(
            FIF.RINGER,
            tr("系统通知"),
            tr("允许 Kazuha 发送系统通知消息"),
            configItem=cfg.enableSystemNotification,
            parent=generalContent
        )
        self.soundCard = SwitchSettingCard(
            FIF.MUSIC,
            tr("全局音效"),
            tr("开启或关闭应用内的所有音效"),
            configItem=cfg.enableGlobalSound,
            parent=generalContent
        )
        self.animationCard = SwitchSettingCard(
            FIF.VIDEO,
            tr("全局动画"),
            tr("开启或关闭界面过渡动画"),
            configItem=cfg.enableGlobalAnimation,
            parent=generalContent
        )
        
        generalLayout.addWidget(self.startUpCard)
        generalLayout.addWidget(self.notificationCard)
        generalLayout.addWidget(self.soundCard)
        generalLayout.addWidget(self.animationCard)

        self.personalInterface, personalContent, personalLayout = _create_page(self)
        self.personalInterface.setObjectName("settings-personal")
        
        self.personalTitle = LargeTitleLabel(tr("个性化"), personalContent)
        _apply_title_style(self.personalTitle)
        personalLayout.addWidget(self.personalTitle)
        
        self.themeCard = OptionsSettingCard(
            cfg.themeMode,
            FIF.BRUSH,
            tr("应用主题"),
            tr("调整应用的外观颜色模式"),
            texts=[tr("浅色"), tr("深色"), tr("跟随系统")],
            parent=personalContent
        )
        self.navPosCard = OptionsSettingCard(
            cfg.navPosition,
            FIF.LAYOUT,
            tr("翻页组件位置"),
            tr("调整翻页按钮在屏幕上的显示位置"),
            texts=[tr("底部两侧"), tr("中部两侧")],
            parent=personalContent
        )
        self.paddingCard = OptionsSettingCard(
            cfg.screenPadding,
            FIF.ZOOM_IN,
            tr("屏幕边距"),
            tr("调整所有组件距离屏幕边缘的距离"),
            texts=[tr("较小"), tr("正常"), tr("较大")],
            parent=personalContent
        )
        self.timerPosCard = OptionsSettingCard(
            cfg.timerPosition,
            FIF.HISTORY,
            tr("倒计时位置"),
            tr("调整倒计时窗口的默认显示位置"),
            texts=[tr("居中"), tr("左上角"), tr("右上角"), tr("左下角"), tr("右下角")],
            parent=personalContent
        )
        
        personalLayout.addWidget(self.themeCard)
        personalLayout.addWidget(self.navPosCard)
        personalLayout.addWidget(self.paddingCard)
        personalLayout.addWidget(self.timerPosCard)

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

    def _create_update_interface(self):
        page, content, layout = _create_page(self)
        page.setObjectName("settings-update")
        
        title = LargeTitleLabel(tr("更新"), content)
        _apply_title_style(title)
        layout.addWidget(title)
        
        spacer1 = QWidget()
        spacer1.setFixedHeight(20)
        layout.addWidget(spacer1)
        
        # Status Card
        statusCard = QWidget(content)
        statusCard.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        statusCard.setMinimumHeight(120)  # Increase minimum height to accommodate content
        statusLayout = QHBoxLayout(statusCard)
        statusLayout.setContentsMargins(24, 24, 24, 24)
        statusLayout.setSpacing(20)
        
        self.updateStatusIcon = QLabel(statusCard)
        self.updateStatusIcon.setFixedSize(64, 64)
        self.updateStatusIcon.setPixmap(FIF.COMPLETED.icon().pixmap(QSize(64, 64)))
        
        infoLayout = QVBoxLayout()
        infoLayout.setSpacing(4)
        infoLayout.addStretch(1)
        
        self.updateStatusTitle = BodyLabel(tr("你使用的是最新版本"), statusCard)
        self.updateStatusTitle.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.updateStatusTitle.setWordWrap(True)
        
        now_str = datetime.datetime.now().strftime("%H:%M")
        self.updateLastCheckLabel = CaptionLabel(tr("上次检查时间: 今天, {time}").format(time=now_str), statusCard)
        self.updateLastCheckLabel.setStyleSheet("color: gray;")
        
        # Loading Widget (Inside Info Layout)
        self.updateLoadingWidget = QWidget(statusCard)
        loadingLayout = QHBoxLayout(self.updateLoadingWidget)
        loadingLayout.setContentsMargins(0, 0, 0, 0)
        loadingLayout.setSpacing(8)
        self.updateRing = IndeterminateProgressRing(self.updateLoadingWidget)
        self.updateRing.setFixedSize(16, 16)
        self.updateRing.setStrokeWidth(2)
        self.updateLoadingLabel = BodyLabel(tr("正在从网络检查更新..."), self.updateLoadingWidget)
        loadingLayout.addWidget(self.updateRing)
        loadingLayout.addWidget(self.updateLoadingLabel)
        loadingLayout.addStretch(1)
        self.updateLoadingWidget.hide()
        
        infoLayout.addWidget(self.updateStatusTitle)
        infoLayout.addWidget(self.updateLastCheckLabel)
        infoLayout.addWidget(self.updateLoadingWidget)
        
        infoLayout.addStretch(1)
        
        self.checkUpdateButton = PushButton(tr("检查更新"), statusCard)
        self.checkUpdateButton.setFixedWidth(120)
        self.checkUpdateButton.clicked.connect(self._on_update_card_clicked)
        
        statusLayout.addWidget(self.updateStatusIcon, 0, Qt.AlignmentFlag.AlignVCenter)
        statusLayout.addLayout(infoLayout, 1)
        statusLayout.addWidget(self.checkUpdateButton, 0, Qt.AlignmentFlag.AlignVCenter)
        
        layout.addWidget(statusCard)
        
        spacer2 = QWidget()
        spacer2.setFixedHeight(20)
        layout.addWidget(spacer2)
        
        # Segment
        self.updateSegment = SegmentedWidget(content)
        self.updateSegment.addItem("logs", tr("更新日志"))
        self.updateSegment.addItem("settings", tr("更新设置"))
        self.updateSegment.setCurrentItem("logs")
        
        layout.addWidget(self.updateSegment)
        
        spacer3 = QWidget()
        spacer3.setFixedHeight(10)
        layout.addWidget(spacer3)
        
        # Stack
        self.logPage = QWidget(content)
        logLayout = QVBoxLayout(self.logPage)
        logLayout.setContentsMargins(0, 0, 0, 0)
        
        # Create a container widget for the log label to better control height
        self.logContainer = QWidget(self.logPage)
        self.logContainer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        logContainerLayout = QVBoxLayout(self.logContainer)
        logContainerLayout.setContentsMargins(0, 0, 0, 0)
        
        self.latestReleaseLogLabel = AutoHeightTextBrowser(self.logContainer)
        self.latestReleaseLogLabel.setPlainText(tr("暂无更新日志信息"))
        
        logContainerLayout.addWidget(self.latestReleaseLogLabel)
        
        logLayout.addWidget(self.logContainer)
        logLayout.addStretch(1)
        
        # Settings Page
        self.settingsPage = QWidget(content)
        settingsLayout = QVBoxLayout(self.settingsPage)
        settingsLayout.setContentsMargins(0, 0, 0, 0)
        
        # Create a container widget for settings card to better control height
        self.settingsContainer = QWidget(self.settingsPage)
        self.settingsContainer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        settingsContainerLayout = QVBoxLayout(self.settingsContainer)
        settingsContainerLayout.setContentsMargins(0, 0, 0, 0)
        
        self.autoCheckCard = SwitchSettingCard(
            FIF.SYNC,
            tr("自动检查更新"),
            tr("在启动时自动检查新版本"),
            configItem=cfg.checkUpdateOnStart,
            parent=self.settingsContainer
        )
        self.autoCheckCard.setChecked(cfg.checkUpdateOnStart.value)
        self.autoCheckCard.setEnabled(True)
        self.autoCheckCard.setMinimumHeight(80)  # Increased minimum height
        
        settingsContainerLayout.addWidget(self.autoCheckCard)
        settingsContainerLayout.addStretch(1)
        
        settingsLayout.addWidget(self.settingsContainer)
        settingsLayout.addStretch(1)
        
        self.settingsPage.hide()
        
        layout.addWidget(self.logPage)
        layout.addWidget(self.settingsPage)
        
        self.updateSegment.currentItemChanged.connect(lambda k: self._on_update_tab_changed(k))
        
        return page, content, layout

    def _on_update_tab_changed(self, k):
        if k == "logs":
            self.settingsPage.hide()
            self.logPage.show()
        else:
            self.logPage.hide()
            self.settingsPage.show()

    def set_update_info(self, version, changelog, is_latest=False):
        self._current_version = version
        self._current_changelog = changelog
        self._is_latest = is_latest
        v = version or tr("未知")
        
        # If is_latest is not explicitly passed (e.g. from old call), try to infer, 
        # but preferring the explicit flag from BusinessLogic which did the numeric comparison
        if not is_latest and self.currentVersion and version:
             v_remote = version.lower().lstrip('v')
             v_local = self.currentVersion.lower().lstrip('v')
             if v_remote == v_local:
                 is_latest = True

        if not is_latest:
             self.updateStatusTitle.setText(tr("发现新版本: {version}").format(version=v))
             self.updateStatusIcon.setPixmap(FIF.SYNC.icon().pixmap(QSize(64, 64)))
             self.checkUpdateButton.setText(tr("立即下载"))
             self.checkUpdateButton.clicked.disconnect()
             self.checkUpdateButton.clicked.connect(self._on_start_update_clicked)
        else:
             self.updateStatusTitle.setText(tr("你使用的是最新版本"))
             self.updateStatusIcon.setPixmap(FIF.COMPLETED.icon().pixmap(QSize(64, 64)))
             self.checkUpdateButton.setText(tr("检查更新"))
             self.checkUpdateButton.clicked.disconnect()
             self.checkUpdateButton.clicked.connect(self._on_update_card_clicked)

        self.updateLoadingWidget.hide()
        self.updateLastCheckLabel.show()
        
        # Ensure changelog is string
        if not isinstance(changelog, str):
            changelog = str(changelog) if changelog is not None else ""
            
        # Update log text
        text = changelog.strip() if changelog else tr("无更新日志")
        
        # Just show text, avoid markdown syntax since renderer is disabled
        self.latestReleaseLogLabel.setPlainText(text)
        
        # Force refresh height to avoid truncation
        # Give it a bit more time for layout to settle, and force a larger minimum height if needed
        QTimer.singleShot(200, lambda: self.latestReleaseLogLabel._adjust_height())
            
    def switch_to_update_page(self):
        self.switchTo(self.updateInterface)
        
    def set_download_progress(self, val):
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.show()
            self.updateLoadingLabel.setText(tr("正在下载: {0}%").format(val))
            if hasattr(self, "checkUpdateButton"):
                self.checkUpdateButton.setEnabled(False)

    def on_config_changed(self):
        self.configChanged.emit()
        
        # Reload HTML to update colors for dark/light mode
        if hasattr(self, 'latestReleaseLogLabel') and hasattr(self.latestReleaseLogLabel, '_raw_markdown'):
            self.latestReleaseLogLabel.setMarkdown(self.latestReleaseLogLabel._raw_markdown)

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
        if theme == Theme.AUTO:
            theme = Theme.DARK if isDarkTheme() else Theme.LIGHT
            
        color = "white" if theme == Theme.DARK else "black"
        self.latestReleaseLogLabel.setStyleSheet(f"background-color: transparent; color: {color}; border: none;")
        
        # Re-render markdown with new theme color if there is content
        if hasattr(self, "_current_changelog") and self._current_changelog:
            self.set_update_info(
                getattr(self, "_current_version", ""),
                self._current_changelog,
                getattr(self, "_is_latest", False)
            )

    def _on_update_card_clicked(self):
        if self._is_update_available:
            self.startUpdateClicked.emit()
            return
            
        if hasattr(self, "latestReleaseVersionLabel"):
            self.latestReleaseVersionLabel.setText(tr("最新可用 Release 版本：正在检查..."))
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.show()
            self.updateLoadingLabel.setText(tr("正在从网络检查更新..."))
            if hasattr(self, "checkUpdateButton"):
                self.checkUpdateButton.setEnabled(True)
        self.checkUpdateClicked.emit()

    def _auto_check_update(self):
        self._on_update_card_clicked()

    def stop_update_loading(self):
        if hasattr(self, "updateLoadingWidget"):
            self.updateLoadingWidget.hide()
