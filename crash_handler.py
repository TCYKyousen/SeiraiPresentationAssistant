import os
import subprocess
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime

from PyQt6.QtCore import Qt, QSize, QCoreApplication
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtWidgets import QApplication, QDialog, QHBoxLayout, QLabel, QVBoxLayout
from qfluentwidgets import BodyLabel, FluentIcon, PlainTextEdit, PrimaryPushButton, PushButton, TitleLabel, isDarkTheme

def tr(text: str) -> str:
    return QCoreApplication.translate("CrashHandler", text)


@dataclass(frozen=True)
class CrashInfo:
    title: str
    details: str


def _default_log_path() -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
    return os.path.join(base_dir, "debug.log")


def _format_exception(exc_type, exc, tb) -> str:
    parts = []
    parts.append(f"Time: {datetime.now().isoformat(sep=' ', timespec='seconds')}")
    parts.append(f"Python: {sys.version}")
    parts.append(f"Executable: {sys.executable}")
    parts.append(f"Args: {sys.argv}")
    parts.append("")
    parts.extend(traceback.format_exception(exc_type, exc, tb))
    return "\n".join(parts)


def _creation_flags() -> int:
    flags = 0
    if hasattr(subprocess, "DETACHED_PROCESS"):
        flags |= subprocess.DETACHED_PROCESS
    if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
        flags |= subprocess.CREATE_NEW_PROCESS_GROUP
    return flags


def _restart_cmd() -> list[str]:
    app_path = os.path.abspath(sys.argv[0])
    args = sys.argv[1:]

    if app_path.lower().endswith(".py"):
        python_exe = sys.executable.replace("python.exe", "pythonw.exe")
        if not os.path.exists(python_exe):
            python_exe = sys.executable
        return [python_exe, app_path, *args]

    return [sys.executable, *args]


def _icon_path(filename: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, "icons", filename)


def _fluent_icon(name: str) -> QIcon:
    icon = getattr(FluentIcon, name, None)
    if icon is None:
        return QIcon()
    return icon.icon()


class CrashWindow(QDialog):
    def __init__(self, crash: CrashInfo, parent=None):
        super().__init__(parent)
        self._details = crash.details

        self.setWindowTitle(crash.title)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.resize(860, 560)

        from PyQt6.QtGui import QFont
        font = QFont("Bahnschrift", 14)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        self.setFont(font)

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        if isDarkTheme():
            self.setStyleSheet("QDialog { background-color: rgb(32, 32, 32); color: white; }")
        else:
            self.setStyleSheet("QDialog { background-color: rgb(243, 243, 243); color: black; }")

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(10)

        icon_label = QLabel(self)
        icon_label.setFixedSize(40, 40)
        icon_label.setScaledContents(True)
        for name in ("resources/icons/ErrorEmoji.png", "resources/icons/ErrorEmoji.png"):
            p = _icon_path(name)
            if os.path.exists(p):
                pm = QPixmap(p)
                if not pm.isNull():
                    icon_label.setPixmap(pm.scaled(icon_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
                break

        title = TitleLabel(tr("Kazuha 崩溃啦！"), self)
        title_row.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignVCenter)
        title_row.addWidget(title, 0, Qt.AlignmentFlag.AlignVCenter)
        title_row.addStretch(1)

        subtitle = BodyLabel(tr("程序发生错误并已停止运行。你可以复制错误信息用于反馈。"), self)

        self.details_edit = PlainTextEdit(self)
        self.details_edit.setReadOnly(True)
        self.details_edit.setPlainText(crash.details)
        
        if isDarkTheme():
            self.details_edit.setStyleSheet("QPlainTextEdit { background-color: rgb(43, 43, 43); color: #e0e0e0; border: 1px solid rgb(50, 50, 50); border-radius: 5px; }")
        else:
            self.details_edit.setStyleSheet("QPlainTextEdit { background-color: rgb(255, 255, 255); color: #333333; border: 1px solid rgb(229, 229, 229); border-radius: 5px; }")

        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, 0, 0, 0)
        btn_row.setSpacing(10)

        self.exit_btn = PushButton(tr("退出程序"), self)
        self.copy_btn = PushButton(tr("复制错误日志"), self)
        self.restart_btn = PrimaryPushButton(tr("重启程序"), self)
        self.exit_btn.setIcon(_fluent_icon("POWER_BUTTON"))
        self.copy_btn.setIcon(_fluent_icon("COPY"))
        self.restart_btn.setIcon(_fluent_icon("SYNC"))
        self.exit_btn.setIconSize(QSize(16, 16))
        self.copy_btn.setIconSize(QSize(16, 16))
        self.restart_btn.setIconSize(QSize(16, 16))

        btn_row.addWidget(self.exit_btn)
        btn_row.addWidget(self.copy_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(self.restart_btn)

        root.addLayout(title_row)
        root.addWidget(subtitle)
        root.addWidget(self.details_edit, 1)
        root.addLayout(btn_row)

    def details_text(self) -> str:
        return self._details


class CrashHandler:
    def __init__(self, log_path: str | None = None):
        self._log_path = log_path or _default_log_path()
        self._handling = False

    def install(self):
        sys.excepthook = self._sys_excepthook
        try:
            import threading

            if hasattr(threading, "excepthook"):
                threading.excepthook = self._threading_excepthook
        except Exception:
            pass

    def _threading_excepthook(self, args):
        self.handle(args.exc_type, args.exc_value, args.exc_traceback)

    def _sys_excepthook(self, exc_type, exc, tb):
        self.handle(exc_type, exc, tb)

    def handle(self, exc_type, exc, tb):
        if self._handling:
            try:
                os._exit(1)
            except Exception:
                sys.exit(1)

        self._handling = True
        details = _format_exception(exc_type, exc, tb)
        try:
            print(details, file=sys.stderr)
        except Exception:
            pass
        try:
            with open(self._log_path, "a", encoding="utf-8") as f:
                f.write(details)
                f.write("\n\n")
        except Exception:
            pass
        try:
            crash = CrashInfo(title=tr("Kazuha 崩溃啦！"), details=details)
            app = QApplication.instance()
            if app is None:
                app = QApplication(sys.argv[:1])
            
            win = CrashWindow(crash, parent=None)
            handler = self

            def do_exit():
                try:
                    app.quit()
                finally:
                    os._exit(1)

            def do_copy():
                clipboard = app.clipboard()
                if clipboard is not None:
                    clipboard.setText(win.details_text())

            def do_restart():
                handler.restart()

            win.exit_btn.clicked.connect(do_exit)
            win.copy_btn.clicked.connect(do_copy)
            win.restart_btn.clicked.connect(do_restart)
            win.exec()
        except Exception:
            pass
        try:
            os._exit(1)
        except Exception:
            sys.exit(1)

    def restart(self):
        try:
            cmd = _restart_cmd()
            flags = _creation_flags()
            if flags:
                subprocess.Popen(cmd, cwd=os.getcwd(), close_fds=True, creationflags=flags)
            else:
                subprocess.Popen(cmd, cwd=os.getcwd(), close_fds=True)
        except Exception:
            try:
                subprocess.Popen(_restart_cmd(), cwd=os.getcwd(), close_fds=True)
            except Exception:
                pass

        os._exit(0)

    def test_crash(self):
        raise RuntimeError("Crash test")


class CrashAwareApplication(QApplication):
    def __init__(self, argv, crash_handler: CrashHandler):
        super().__init__(argv)
        self._crash_handler = crash_handler

    def notify(self, receiver, event):
        try:
            return super().notify(receiver, event)
        except Exception:
            exc_type, exc, tb = sys.exc_info()
            if exc_type is not None and exc is not None and tb is not None:
                self._crash_handler.handle(exc_type, exc, tb)
            return False


def run_watchdog_process(parent_pid: int, log_path: str | None = None):
    import ctypes
    import os
    log_path = log_path or _default_log_path()
    kernel32 = ctypes.windll.kernel32
    SYNCHRONIZE = 0x00100000
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    WAIT_OBJECT_0 = 0x00000000
    handle = kernel32.OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, False, int(parent_pid))
    if not handle:
        return
    exit_code = ctypes.c_ulong(259)
    rc = kernel32.WaitForSingleObject(handle, 0xFFFFFFFF)
    if rc != WAIT_OBJECT_0:
        kernel32.CloseHandle(handle)
        return
    if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
        kernel32.CloseHandle(handle)
        return
    code = int(exit_code.value)
    kernel32.CloseHandle(handle)
    if code == 0:
        return
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            details = f.read()
    except Exception:
        details = tr("Process {pid} exited with code {code}").format(pid=parent_pid, code=code)
    crash = CrashInfo(title=tr("Kazuha 崩溃啦！"), details=details)
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv[:1])
    win = CrashWindow(crash, parent=None)
    handler = CrashHandler(log_path=log_path)

    def do_exit():
        try:
            app.quit()
        finally:
            os._exit(1)

    def do_copy():
        clipboard = app.clipboard()
        if clipboard is not None:
            clipboard.setText(win.details_text())

    def do_restart():
        handler.restart()

    win.exit_btn.clicked.connect(do_exit)
    win.copy_btn.clicked.connect(do_copy)
    win.restart_btn.clicked.connect(do_restart)
    win.exec()
