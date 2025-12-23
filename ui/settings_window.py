from PyQt6.QtCore import Qt, pyqtSignal, QTimer
from PyQt6.QtWidgets import QWidget, QVBoxLayout
from PyQt6.QtGui import QFont
import json
import os

from qfluentwidgets import (
    FluentIcon as FIF,
    SettingCardGroup, SwitchSettingCard, OptionsSettingCard, PushSettingCard,
    SmoothScrollArea, ExpandLayout, Theme, setTheme, setThemeColor
)
from ui.custom_settings import SchematicOptionsSettingCard
from controllers.business_logic import cfg
from ui.crash_dialog import CrashDialog, trigger_crash
from ui.visual_settings import ClockSettingCard

class SettingsWindow(QWidget):
    configChanged = pyqtSignal()
    checkUpdateClicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选项")
        self.resize(800, 600)
        
        # Apply font
        font = QFont("Bahnschrift", 14)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        self.setFont(font)
        
        # Main Layout
        self.mainLayout = QVBoxLayout(self)
        self.mainLayout.setContentsMargins(24, 24, 24, 24)
        self.mainLayout.setSpacing(0)
        
        # Create a single scrolling interface
        self.scrollInterface = SmoothScrollArea(self)
        self.scrollInterface.setObjectName("scrollInterface")
        self.scrollInterface.setStyleSheet("SmoothScrollArea { background-color: transparent; border: none; }")
        
        self.scrollWidget = QWidget()
        self.scrollWidget.setObjectName("scrollWidget")
        self.expandLayout = ExpandLayout(self.scrollWidget)
        
        self.scrollInterface.setWidget(self.scrollWidget)
        self.scrollInterface.setWidgetResizable(True)
        
        self.mainLayout.addWidget(self.scrollInterface)
        
        # --- Groups ---
        
        # 1. General Group
        self.generalGroup = SettingCardGroup("常规", self.scrollWidget)
        
        self.startupCard = SwitchSettingCard(
            FIF.POWER_BUTTON,
            "开机自启",
            "跟随系统启动自动运行",
            configItem=cfg.enableStartUp,
            parent=self.generalGroup
        )
        
        self.themeCard = SchematicOptionsSettingCard(
            cfg.themeMode,
            FIF.BRUSH,
            "应用主题",
            "调整应用外观",
            texts=["浅色", "深色", "跟随系统"],
            schematic_type="theme",
            parent=self.generalGroup
        )
        
        self.generalGroup.addSettingCard(self.startupCard)
        self.generalGroup.addSettingCard(self.themeCard)
        self.expandLayout.addWidget(self.generalGroup)
        
        # 2. Layout Group
        self.layoutGroup = SettingCardGroup("布局与位置", self.scrollWidget)
        
        self.navPosCard = SchematicOptionsSettingCard(
            cfg.navPosition,
            FIF.ALIGNMENT,
            "翻页导航位置",
            "调整翻页按钮在屏幕上的位置",
            texts=["底部两端", "中部两侧"],
            schematic_type="nav_pos",
            parent=self.layoutGroup
        )
        
        self.clockEnableCard = SwitchSettingCard(
            FIF.DATE_TIME,
            "显示时钟",
            "是否显示桌面悬浮时钟组件",
            configItem=cfg.enableClock,
            parent=self.layoutGroup
        )
        
        self.clockPosCard = OptionsSettingCard(
            cfg.clockPosition,
            FIF.HISTORY, 
            "时钟位置",
            "调整悬浮时钟的显示位置",
            texts=["左上角", "右上角", "左下角", "右下角"],
            parent=self.layoutGroup
        )
        
        self.clockSettingCard = ClockSettingCard(
            FIF.DATE_TIME,
            "时钟样式",
            "自定义悬浮时钟的显示内容和外观",
            parent=self.layoutGroup
        )
        
        self.timerPosCard = OptionsSettingCard(
            cfg.timerPosition,
            FIF.SPEED_HIGH, 
            "计时器位置",
            "调整倒计时窗口的显示位置",
            texts=["屏幕中央", "左上角", "右上角", "左下角", "右下角"],
            parent=self.layoutGroup
        )
        
        self.layoutGroup.addSettingCard(self.navPosCard)
        self.layoutGroup.addSettingCard(self.clockEnableCard)
        self.layoutGroup.addSettingCard(self.clockPosCard)
        self.layoutGroup.addSettingCard(self.clockSettingCard)
        self.layoutGroup.addSettingCard(self.timerPosCard)
        self.expandLayout.addWidget(self.layoutGroup)
        
        # 3. Danger Group
        self.dangerGroup = SettingCardGroup("危险功能", self.scrollWidget)
        self.crashCard = PushSettingCard(
            "触发",
            FIF.DELETE,
            "崩溃测试",
            "仅用于开发调试，请勿在演示时点击",
            parent=self.dangerGroup
        )
        self.dangerGroup.addSettingCard(self.crashCard)
        self.expandLayout.addWidget(self.dangerGroup)

        # 4. About Group
        self.aboutGroup = SettingCardGroup("关于", self.scrollWidget)
        
        # Get version
        version = "Unknown"
        try:
             import sys
             base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
             v_path = os.path.join(base_dir, "config", "version.json")
             with open(v_path, 'r', encoding='utf-8') as f:
                 info = json.load(f)
                 version = info.get("versionName", "Unknown")
        except:
             pass

        self.aboutCard = PushSettingCard(
            "检查更新",
            FIF.INFO,
            "Seirai PPT Assistant",
            f"当前版本: {version}  © 2024 Seirai Studio",
            parent=self.aboutGroup
        )
        
        self.aboutGroup.addSettingCard(self.aboutCard)
        self.expandLayout.addWidget(self.aboutGroup)
        
        self.aboutCard.clicked.connect(self.checkUpdateClicked)
        
        # Connect signals
        self.crashCard.clicked.connect(self.show_crash_dialog)
        
        # Config changed propagation
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
        
    def on_config_changed(self):
        self.configChanged.emit()
        
    def show_crash_dialog(self):
        w = CrashDialog(self)
        if w.exec():
            settings = w.get_settings()
            if settings['countdown']:
                QTimer.singleShot(3000, lambda: trigger_crash(settings))
            else:
                trigger_crash(settings)
                
    def set_theme(self, theme):
        # Apply standard background colors for QWidget based on theme
        try:
            import qfluentwidgets
            actual = qfluentwidgets.theme() if theme == Theme.AUTO else theme
        except Exception:
            actual = theme
        
        # Use standard Fluent UI window background colors
        if actual == Theme.LIGHT:
            bg = "rgb(243, 243, 243)" # Standard Light (Mica Alt approximation)
        else:
            bg = "rgb(32, 32, 32)"    # Standard Dark
            
        self.setStyleSheet(f"SettingsWindow {{ background-color: {bg}; }}")
        self.scrollWidget.setStyleSheet(f"QWidget#scrollWidget {{ background-color: {bg}; }}")
