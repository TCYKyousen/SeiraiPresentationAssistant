import os
import json
from PyQt6.QtCore import Qt, QCoreApplication, QRect
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout
from PyQt6.QtGui import QPixmap, QPainter, QColor, QLinearGradient
from qfluentwidgets import (ScrollArea, LargeTitleLabel, BodyLabel, CaptionLabel, 
                            PrimaryPushSettingCard, SettingCard, FluentIcon as FIF,
                            TransparentPushButton, TitleLabel, isDarkTheme, ImageLabel,
                            ExpandGroupSettingCard, HyperlinkButton, MessageBoxBase,
                            SubtitleLabel, TextEdit)
from controllers.business_logic import cfg, get_app_base_dir

def tr(text: str) -> str:
    return QCoreApplication.translate("AboutPage", text)

class ThirdPartyLibsDialog(MessageBoxBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel(tr("第三方库"), self)
        self.textEdit = TextEdit(self)
        self.textEdit.setReadOnly(True)
        self.textEdit.setFixedHeight(300)
        
        libs_info = (
            "PySide6: https://www.qt.io/qt-for-python\n"
            "PySide6-Fluent-Widgets: https://github.com/zhiyiYo/PySide6-Fluent-Widgets\n"
            "PySide6-Charts: https://pypi.org/project/PySide6-Charts/\n"
            "PySide6-WebEngine: https://pypi.org/project/PySide6-WebEngine/\n"
            "requests: https://github.com/psf/requests\n"
            "Pillow: https://github.com/python-pillow/Pillow\n"
            "python-pptx: https://github.com/scanny/python-pptx\n"
            "psutil: https://github.com/giampaolo/psutil\n"
            "watchdog: https://github.com/gorakhargosh/watchdog"
        )
        self.textEdit.setText(libs_info)
        
        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addWidget(self.textEdit)
        
        self.yesButton.setText(tr("关闭"))
        self.cancelButton.hide()
        self.widget.setFixedWidth(450)

class BannerLabel(QWidget):
    def __init__(self, image_path, parent=None):
        super().__init__(parent)
        self.setFixedHeight(280)
        self.image = QPixmap(image_path)
        
        self.aboutTitleLabel = TitleLabel("Kazuha", self)
        self.aboutTitleLabel.setTextColor(QColor(255, 255, 255))
        font = self.aboutTitleLabel.font()
        font.setPixelSize(28)
        font.setBold(True)
        self.aboutTitleLabel.setFont(font)
        
    def resizeEvent(self, event):
        self.aboutTitleLabel.move(24, 220)
        super().resizeEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        
        w, h = self.width(), self.height()
        bg_color = QColor("#202020")
        
        painter.fillRect(self.rect(), bg_color)
        
        if not self.image.isNull():
            img_h = h
            img_w = int(self.image.width() * (img_h / self.image.height()))
            
            x = (w - img_w) // 2
            target_rect = QRect(x, 0, img_w, img_h)
            
            left_gradient = QLinearGradient(0, 0, x + 20, 0)
            left_gradient.setColorAt(0, bg_color)
            left_gradient.setColorAt(1, Qt.GlobalColor.transparent)
            
            right_gradient = QLinearGradient(w, 0, x + img_w - 20, 0)
            right_gradient.setColorAt(0, bg_color)
            right_gradient.setColorAt(1, Qt.GlobalColor.transparent)
            
            painter.drawPixmap(target_rect, self.image)
            
            painter.fillRect(0, 0, x + 20, h, left_gradient)
            painter.fillRect(x + img_w - 20, 0, w - (x + img_w - 20), h, right_gradient)

class AboutPage(ScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("AboutPage")
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        self.view = QWidget()
        self.view.setObjectName("view")
        self.setWidget(self.view)
        
        self.mainLayout = QVBoxLayout(self.view)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.mainLayout.setSpacing(0)
        
        version = "Unknown"
        code_name = ""
        try:
            v_path = get_app_base_dir() / "config" / "version.json"
            if not v_path.exists():
                v_path = get_app_base_dir() / "version.json"
            if v_path.exists():
                with open(v_path, "r", encoding="utf-8") as f:
                    v_data = json.load(f)
                    version = v_data.get("versionName", "Unknown")
                    code_name = v_data.get("codeName", "")
        except Exception:
            pass

        banner_path = ""
        try:
            possible_paths = [
                os.path.join(os.path.dirname(__file__), "banner.png"),
                str(get_app_base_dir() / "resources" / "banner.png"),
                str(get_app_base_dir() / "resources" / "icons" / "banner.png")
            ]
            for p in possible_paths:
                if os.path.exists(p):
                    banner_path = p
                    break
        except Exception:
            pass
                
        self.banner = BannerLabel(banner_path, self.view)
        self.mainLayout.addWidget(self.banner)
        
        self.contentWidget = QWidget(self.view)
        self.contentLayout = QVBoxLayout(self.contentWidget)
        self.contentLayout.setContentsMargins(24, 24, 24, 24)
        self.contentLayout.setSpacing(12)
        
        title = f"Kazuha {version}"
        if code_name:
            title += f" ({code_name})"
            
        self.appInfoExpander = ExpandGroupSettingCard(
            FIF.INFO,
            title,
            tr("一个用于教室和演示场景的现代 PPT 演示辅助工具。"),
            parent=self.contentWidget
        )
        self.appInfoExpander.viewLayout.setContentsMargins(0, 0, 0, 0)
        self.appInfoExpander.viewLayout.setSpacing(0)
        
        license_btn = HyperlinkButton(
            url="https://github.com/Haraguse/Kazuha/blob/main/LICENSE",
            text=tr("查看许可证"),
            parent=self.appInfoExpander
        )
        self.appInfoExpander.addGroup(
            FIF.PEOPLE,
            tr("版权与许可证"),
            tr("版权所有 © 2025 Haraguse。基于 MIT 许可证开源。"),
            license_btn
        )
        
        links_widget = QWidget(self.appInfoExpander)
        links_layout = QHBoxLayout(links_widget)
        links_layout.setContentsMargins(0, 0, 0, 0)
        links_layout.setSpacing(8)
        
        github_link = HyperlinkButton(
            url="https://github.com/Haraguse/Kazuha",
            text="GitHub",
            parent=links_widget
        )
        qq_link = HyperlinkButton(
            url="https://qm.qq.com/cgi-bin/qm/qr?k=YOUR_KEY",
            text=tr("QQ 群"),
            parent=links_widget
        )
        
        links_layout.addWidget(github_link)
        links_layout.addWidget(qq_link)
        
        self.appInfoExpander.addGroup(
            FIF.LINK,
            tr("快速链接"),
            tr("加入我们的社区并关注更新。"),
            links_widget
        )
        
        self.contentLayout.addWidget(self.appInfoExpander)

        self.thanksExpander = ExpandGroupSettingCard(
            FIF.HEART,
            tr("鸣谢"),
            tr("特别感谢所有让 Kazuha 成为可能的项目和资源。"),
            parent=self.contentWidget
        )
        self.thanksExpander.viewLayout.setContentsMargins(0, 0, 0, 0)
        self.thanksExpander.viewLayout.setSpacing(0)
        
        thanks_widget = QWidget(self.thanksExpander)
        thanks_layout = QVBoxLayout(thanks_widget)
        thanks_layout.setContentsMargins(16, 12, 16, 12)
        thanks_layout.setSpacing(12)
        
        acknowledgment_text = tr(
            "本程序灵感来源于 https://easinote.seewo.com/（希沃白板 5）和 希沃课堂助手、鸿合演示助手。\n"
            "音效资源提取自 TCL® CyberUI 且版权归其所有。\n"
            "Emoji 资源来源于米游社和其他与 miHoYo® 相关游戏的社区等，本程序的名称，图标灵感来源，Emoji 资源的版权归其所有。"
        )
        
        text_label = BodyLabel(acknowledgment_text, thanks_widget)
        text_label.setWordWrap(True)
        thanks_layout.addWidget(text_label)
        
        self.thanksExpander.addGroup(
            FIF.INFO,
            tr("致谢"),
            tr("灵感与资源"),
            thanks_widget
        )
        
        libs_widget = QWidget(self.thanksExpander)
        libs_layout = QHBoxLayout(libs_widget)
        libs_layout.setContentsMargins(0, 0, 0, 0)
        libs_layout.setSpacing(4)
        
        libs_pre_text = BodyLabel(tr("鸣谢所有第三方框架 ("), libs_widget)
        libs_btn = HyperlinkButton(url="", text=tr("查看列表"), parent=libs_widget)
        libs_btn.clicked.connect(self._show_libs_dialog)
        libs_post_text = BodyLabel(")", libs_widget)
        
        libs_layout.addWidget(libs_pre_text)
        libs_layout.addWidget(libs_btn)
        libs_layout.addWidget(libs_post_text)
        libs_layout.addStretch(1)
        
        self.thanksExpander.addGroup(
            FIF.DOCUMENT,
            tr("开源项目"),
            tr("本项目使用的第三方库"),
            libs_widget
        )
        
        self.contentLayout.addWidget(self.thanksExpander)
        
        self.mainLayout.addWidget(self.contentWidget)
        self.mainLayout.addStretch(1)

        self._update_style()
        
        cfg.themeMode.valueChanged.connect(self._update_style)

    def _update_style(self):
        self.view.setStyleSheet("QWidget#view{background: transparent}")

    def _show_libs_dialog(self):
        w = ThirdPartyLibsDialog(self.window())
        w.exec()

