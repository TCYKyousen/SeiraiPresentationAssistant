import os
import sys
import subprocess
from PySide6.QtWidgets import QWidget, QApplication
from plugins.interface import AssistantPlugin
from ppt_assistant.core.config import SETTINGS_PATH

class TimerPlugin(AssistantPlugin):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.process = None

    def get_name(self):
        return "计时器"

    def get_icon(self):
        return "timer.svg"

    def execute(self):
        if self.process and self.process.poll() is None:
            if sys.platform == "win32":
                try:
                    import ctypes
                    # Find window by title "Kazuha Timer Plugin"
                    hwnd = ctypes.windll.user32.FindWindowW(None, "Kazuha Timer Plugin")
                    if hwnd:
                        # SW_RESTORE = 9
                        ctypes.windll.user32.ShowWindow(hwnd, 9)
                        ctypes.windll.user32.SetForegroundWindow(hwnd)
                except:
                    pass
            return

        base_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(base_dir, "timer.html")
        root_dir = os.path.dirname(os.path.dirname(os.path.dirname(base_dir)))
        main_path = os.path.join(root_dir, "main.py")
        assets_path = os.path.join(base_dir, "assets", "timer_ring.ogg")
        
        screen = QApplication.primaryScreen()
        screen_geo = screen.geometry() if screen else QWidget().screen().geometry()
        width = str(int(min(max(600, screen_geo.width() * 0.35), screen_geo.width() * 0.5)))
        height = str(int(min(max(500, screen_geo.height() * 0.45), screen_geo.height() * 0.6)))

        env = os.environ.copy()
        env["SETTINGS_PATH"] = SETTINGS_PATH
        env["ASSETS_PATH"] = assets_path

        if getattr(sys, "frozen", False):
            cmd = [
                sys.executable,
                "--webview-runner",
                html_path,
                "Kazuha Timer Plugin",
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
                "Kazuha Timer Plugin",
                width,
                height,
                "true",
            ]

        self.process = subprocess.Popen(cmd, env=env)

    def terminate(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.process = None
