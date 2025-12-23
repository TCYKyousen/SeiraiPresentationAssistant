import sys
import os
import traceback
from datetime import datetime
from PyQt6.QtWidgets import QApplication
from controllers.business_logic import BusinessLogicController, cfg
from ui.widgets import ToolBarWidget, PageNavWidget, SpotlightOverlay, ClockWidget
from crash_handler import CrashAwareApplication, CrashHandler

def setup_logging():
    try:
        log_dir = os.path.join(os.getenv("APPDATA"), "Kazuha")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        log_path = os.path.join(log_dir, "debug.log")
        
        # Reset log file
        with open(log_path, "w") as f:
            f.write(f"Session started at {datetime.now()}\n")
            
        return log_path
    except:
        return None

def log_message(msg):
    log_dir = os.path.join(os.getenv("APPDATA"), "Kazuha")
    log_path = os.path.join(log_dir, "debug.log")
    try:
        with open(log_path, "a") as f:
            f.write(f"{datetime.now()}: {msg}\n")
    except:
        pass

def main():
    log_path = setup_logging()
    
    # Enable high DPI scaling
    from PyQt6.QtCore import Qt
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseDesktopOpenGL)
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
    
    try:
        crash_handler = CrashHandler(log_path=log_path)
        crash_handler.install()
        app = CrashAwareApplication(sys.argv, crash_handler)
        app.setQuitOnLastWindowClosed(False)

        # Set global font
        from PyQt6.QtGui import QFont
        font = QFont()
        font.setFamilies(["Bahnschrift", "Microsoft YaHei"])
        font.setPixelSize(14)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        app.setFont(font)
        
        log_message("Initializing Controller...")
        controller = BusinessLogicController()
        controller.set_font("Bahnschrift")
        
        # Initialize UI components
        log_message("Initializing UI...")
        
        nav_pos = cfg.navPosition.value
        orientation = Qt.Orientation.Vertical if nav_pos == "MiddleSides" else Qt.Orientation.Horizontal
        
        controller.toolbar = ToolBarWidget()
        controller.nav_left = PageNavWidget(is_right=False, orientation=orientation)
        controller.nav_right = PageNavWidget(is_right=True, orientation=orientation)
        controller.clock_widget = ClockWidget()
        controller.clock_widget.apply_settings(cfg)
        controller.spotlight = SpotlightOverlay()
        
        # Setup connections between UI and business logic
        log_message("Setting up connections...")
        controller.setup_connections()
        controller.setup_tray()
        
        # Show the application
        log_message("Showing Controller...")
        controller.show()
        
        log_message("Entering Event Loop...")
        sys.exit(app.exec())
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
    # Ensure multiprocessing support for Windows (freeze_support)
    import multiprocessing
    multiprocessing.freeze_support()
    main()
