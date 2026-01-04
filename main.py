import sys
import os
import traceback
import ctypes
import multiprocessing
from datetime import datetime
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTranslator, QCoreApplication, Qt
from PyQt6.QtGui import QFont
from qfluentwidgets import SplashScreen

from controllers.business_logic import BusinessLogicController, cfg, get_app_base_dir
from ui.widgets import ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
from ui.overlay_window import OverlayWindow
from crash_handler import CrashAwareApplication, CrashHandler

def tr(text):
    return text

def setup_logging():
    try:
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
        log_path = os.path.join(base_dir, "debug.log")
        
        with open(log_path, "w") as f:
            f.write(f"Session started at {datetime.now()}\n")
            
        return log_path
    except:
        return None

def log_message(msg):
    try:
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
        log_path = os.path.join(base_dir, "debug.log")
        with open(log_path, "a") as f:
            f.write(f"{datetime.now()}: {msg}\n")
    except:
        pass

def main():
    log_path = setup_logging()
    
    # Redirect stdout and stderr to capture startup crashes
    try:
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0])))
        sys.stdout = open(os.path.join(base_dir, "stdout.log"), "w", encoding="utf-8", buffering=1)
        sys.stderr = open(os.path.join(base_dir, "stderr.log"), "w", encoding="utf-8", buffering=1)
    except:
        pass

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseDesktopOpenGL)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
    
    try:
        try:
            app_id = tr("万演 主程序通知")
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass

        crash_handler = CrashHandler(log_path=log_path)
        crash_handler.install()
        app = CrashAwareApplication(sys.argv, crash_handler)
        # app = QApplication(sys.argv)
        app.setQuitOnLastWindowClosed(False)

        app.setApplicationName("Kazuha")
        app.setApplicationDisplayName("Kazuha.Sou.Settings")

        font = QFont()
        font.setFamilies(["Bahnschrift", "Microsoft YaHei"])
        font.setPixelSize(14)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        app.setFont(font)

        try:
            base_dir = get_app_base_dir()
            splash_icon = base_dir / "resources" / "icons" / "trayicon.svg"
            splash = SplashScreen(str(splash_icon))
            splash.show()
        except Exception:
            splash = None

        log_message("Initializing Controller...")
        controller = BusinessLogicController()
        controller.set_font("Bahnschrift")
        
        log_message("Initializing UI...")
        
        nav_pos = cfg.navPosition.value
        orientation = Qt.Orientation.Vertical if nav_pos == "MiddleSides" else Qt.Orientation.Horizontal
        
        controller.overlay = OverlayWindow()
        controller.toolbar = ToolBarWidget(parent=controller.overlay)
        controller.nav_left = PageNavWidget(is_right=False, orientation=orientation, parent=controller.overlay)
        controller.nav_right = PageNavWidget(is_right=True, orientation=orientation, parent=controller.overlay)
        controller.clock_widget = ClockWidget(parent=controller.overlay)
        controller.clock_widget.apply_settings(cfg)
        controller.spotlight = SpotlightOverlay(parent=controller.overlay)
        
        log_message("Setting up connections...")
        controller.setup_connections()
        controller.setup_tray()
        
        log_message("Showing Controller...")
        controller.show()
        if splash is not None:
            try:
                splash.finish()
            except Exception as e:
                log_message(f"Splash finish error: {e}")
        
        log_message("Entering Event Loop...")
        ret = app.exec()
        log_message(f"Event Loop exited with code: {ret}")
        sys.exit(ret)
    except Exception as e:
        log_message(f"CRITICAL ERROR: {str(e)}\n{traceback.format_exc()}")
        try:
            app = QApplication.instance()
            if app is None:
                crash_handler = CrashHandler(log_path=log_path)
                app = CrashAwareApplication(sys.argv, crash_handler)
            else:
                crash_handler = CrashHandler(log_path=log_path)

            crash_handler.handle(type(e), e, e.__traceback__)
        except Exception:
            pass
        sys.exit(1)

if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()
