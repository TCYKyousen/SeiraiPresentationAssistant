import unittest
import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from controllers.ppt_core import PPTState
from ui.widgets import PageNavWidget
from controllers.business_logic import Config

class TestBasic(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not QApplication.instance():
            cls.app = QApplication(sys.argv)
        else:
            cls.app = QApplication.instance()

    def test_ppt_state_defaults(self):
        state = PPTState()
        self.assertFalse(state.is_running)
        self.assertEqual(state.slide_index, 0)
        self.assertEqual(state.rect, (0, 0, 0, 0))

    def test_config_defaults(self):
        # We can't easily test Config without init, but we can check if attributes exist
        self.assertTrue(hasattr(Config, 'screenPaddingBottom'))
        self.assertTrue(hasattr(Config, 'navPosition'))

    def test_page_nav_widget_creation(self):
        w = PageNavWidget(orientation=Qt.Orientation.Horizontal)
        self.assertEqual(w.orientation, Qt.Orientation.Horizontal)
        w.close()

        w_vert = PageNavWidget(orientation=Qt.Orientation.Vertical)
        self.assertEqual(w_vert.orientation, Qt.Orientation.Vertical)
        w_vert.close()

if __name__ == '__main__':
    unittest.main()
