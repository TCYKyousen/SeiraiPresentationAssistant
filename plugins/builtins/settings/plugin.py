import os
import sys
import subprocess
from PySide6.QtWidgets import QWidget, QApplication
from plugins.interface import AssistantPlugin
from ppt_assistant.core.config import SETTINGS_PATH

class SettingsPlugin(AssistantPlugin):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.process = None

    def get_name(self):
        return "设置"

    def get_icon(self):
        return "settings.svg" 

    def execute(self):
        if self.process and self.process.poll() is None:
            if sys.platform == "win32":
                try:
                    import ctypes
                    # Find window by title "Settings"
                    hwnd = ctypes.windll.user32.FindWindowW(None, "Settings")
                    if hwnd:
                        # SW_RESTORE = 9
                        ctypes.windll.user32.ShowWindow(hwnd, 9)
                        ctypes.windll.user32.SetForegroundWindow(hwnd)
                except:
                    pass
            return

        base_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(base_dir, "settings.html")
        root_dir = os.path.dirname(os.path.dirname(os.path.dirname(base_dir)))
        main_path = os.path.join(root_dir, "main.py")
        
        # Sizing
        screen = QApplication.primaryScreen()
        screen_geo = screen.geometry() if screen else QWidget().screen().geometry()
        width = str(int(min(1280, screen_geo.width() * 0.9)))
        height = str(int(min(800, screen_geo.height() * 0.9)))

        env = os.environ.copy()
        env["SETTINGS_PATH"] = SETTINGS_PATH

        if getattr(sys, "frozen", False):
            cmd = [
                sys.executable,
                "--webview-runner",
                html_path,
                "Settings",
                width,
                height,
                "true",
            ]
        else:
            cmd = [
                sys.executable,
                main_path,
                "--webview-runner",
                html_path,
                "Settings",
                width,
                height,
                "true",
            ]

        self.process = subprocess.Popen(cmd, env=env)

    def terminate(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.process = None
