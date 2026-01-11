import os
import subprocess
from PySide6.QtWidgets import QFileDialog, QMessageBox, QFileIconProvider
from PySide6.QtCore import QFileInfo
from plugins.interface import AssistantPlugin
from ppt_assistant.core.config import cfg, _save_cfg

class AppLauncherPlugin(AssistantPlugin):
    def __init__(self, parent=None):
        super().__init__(parent)

    def get_name(self):
        return "应用启动器"

    def get_icon(self):
        return "apps.svg"

    def get_apps(self):
        return cfg.quickLaunchApps.value

    def add_app(self):
        file_path, _ = QFileDialog.getOpenFileName(
            None, "选择要添加的程序", "", "Executable Files (*.exe);;All Files (*.*)"
        )
        if file_path:
            name = os.path.splitext(os.path.basename(file_path))[0]
            apps = cfg.quickLaunchApps.value.copy()
            # Check if already exists
            if any(app['path'] == file_path for app in apps):
                return
            
            apps.append({
                "name": name,
                "path": file_path,
                "icon": "" # We could try to extract icon later
            })
            cfg.quickLaunchApps.value = apps
            _save_cfg()
            
            # Notify toolbar to refresh
            if self.context and hasattr(self.context, 'update_toolbar'):
                self.context.update_toolbar()

    def remove_app(self, path):
        apps = cfg.quickLaunchApps.value.copy()
        apps = [app for app in apps if app['path'] != path]
        cfg.quickLaunchApps.value = apps
        _save_cfg()
        if self.context and hasattr(self.context, 'update_toolbar'):
            self.context.update_toolbar()

    def get_apps(self):
        return cfg.quickLaunchApps.value

    def get_app_icon(self, path):
        if not os.path.exists(path):
            return None
        file_info = QFileInfo(path)
        icon_provider = QFileIconProvider()
        icon = icon_provider.icon(file_info)
        if not icon.isNull():
            return icon.pixmap(32, 32)
        return None

    def execute_app(self, path):
        try:
            if os.path.exists(path):
                # Use shell=True for some apps to launch correctly, but Popen is better
                subprocess.Popen(path, shell=True)
            else:
                QMessageBox.warning(None, "错误", f"找不到程序: {path}")
        except Exception as e:
            QMessageBox.critical(None, "错误", f"启动程序失败: {str(e)}")

    def execute(self):
        self.add_app()
