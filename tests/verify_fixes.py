import sys
import os
import unittest
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

# Add root to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from controllers.ppt_core import PPTState
# We need to mock get_theme/setTheme/etc if they are used at module level in widgets.py
# But widgets.py imports them from qfluentwidgets. 
# Hopefully it runs without full environment setup.

# Need to ensure QApplication exists before importing widgets that might create QPixmaps etc.
app = QApplication(sys.argv)

from ui.widgets import TimerWindow

class TestFixes(unittest.TestCase):
    def test_ppt_state(self):
        print("Testing PPTState fields...")
        state = PPTState()
        self.assertTrue(hasattr(state, 'rect'))
        self.assertTrue(hasattr(state, 'is_fullscreen'))
        print("PPTState fields verified.")
        
    def test_timer_reset(self):
        print("Testing TimerWindow reset logic...")
        w = TimerWindow()
        # Simulate state
        w.down_running = True
        w.down_remaining = 10
        # w.stack.setCurrentWidget(w.down_label) # This might fail if down_label is not in stack directly but shown/hidden
        
        # Call reset
        w.reset_down()
        
        self.assertFalse(w.down_running)
        self.assertEqual(w.down_remaining, 0)
        self.assertEqual(w.stack.currentWidget(), w.down_page)
        print("TimerWindow reset verified.")
        w.close()

if __name__ == '__main__':
    unittest.main()
