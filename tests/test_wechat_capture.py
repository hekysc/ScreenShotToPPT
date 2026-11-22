import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
if sys.platform != "win32":
    for module in ("win32gui", "win32ui", "win32con"):
        sys.modules.setdefault(module, Mock())

from PIL import Image
from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication
from main import MainApp
from region_capture import RegionPicker


class CaptureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_targets_and_saved_frames(self):
        image = Image.new("RGB", (180, 100), "teal")
        with tempfile.TemporaryDirectory() as output, \
                patch("main.get_window_list", return_value=[("微信", 101), ("微信", 202)]), \
                patch("main.capture_window", return_value=(image, "Succeed")) as grab, \
                patch.object(MainApp, "load_settings"):
            window = MainApp()
            window.capture_folder = output
            window.capture_mode.setCurrentIndex(2)
            window.combo.setCurrentIndex(1)
            window.refresh_window_list()
            window.capture_loop()
            grab.assert_called_with(202)
            window.combo.setCurrentIndex(0)
            window.capture_loop()
            grab.assert_called_with(101)
            window.capture_mode.setCurrentIndex(1)
            window.region = (-120, 10, -40, 60)
            cropped = Image.new("RGB", (80, 50), "orange")
            with patch("main.ImageGrab.grab", return_value=cropped) as desktop:
                window.capture_loop()
                window.capture_loop()
                desktop.assert_called_with(bbox=window.region, all_screens=True)
            files = list(Path(output).glob("*.png"))
            self.assertEqual(len(files), 2)
            with Image.open(next(Path(output).glob("视频区域*.png"))) as saved:
                self.assertEqual(saved.tobytes(), cropped.tobytes())

    def test_reverse_drag_on_left_monitor(self):
        picker = RegionPicker(Image.new("RGB", (1600, 800)), QPoint(-800, 0))
        picker.resize(800, 600)
        rect = picker.image_rect()
        QTest.mousePress(picker, Qt.LeftButton, pos=rect.topLeft() + QPoint(199, 149))
        QTest.mouseRelease(picker, Qt.LeftButton, pos=rect.topLeft() + QPoint(100, 50))
        self.assertEqual(picker.bounds, (-600, 100, -400, 300))
