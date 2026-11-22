"""Windows desktop integration test: the covered window changes behind an opaque cover."""
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PIL import ImageGrab
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget
import win32gui
from background_capture import capture_background


def pump(seconds):
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        app.processEvents()
        time.sleep(0.01)


app = QApplication([])
target = QWidget(flags=Qt.FramelessWindowHint)
cover = QWidget(flags=Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
target.setGeometry(100, 100, 300, 200)
cover.setGeometry(80, 80, 340, 240)
target.setStyleSheet("background-color: red")
cover.setStyleSheet("background-color: blue")
target.show()
cover.show()
cover.raise_()
pump(0.4)
hwnd = int(target.winId())
left, top, right, bottom = win32gui.GetWindowRect(hwnd)
x, y = (left + right) // 2, (top + bottom) // 2
with ThreadPoolExecutor(1) as worker:
    for css, expected in [("red", (255, 0, 0)), ("lime", (0, 255, 0))]:
        target.setStyleSheet(f"background-color: {css}")
        target.repaint()
        pump(0.25)
        actual = ImageGrab.grab((x, y, x + 1, y + 1), all_screens=True).getpixel((0, 0))
        assert actual == (0, 0, 255), f"Cover missing: {actual}"
        future = worker.submit(capture_background, hwnd)
        while not future.done():
            pump(0.02)
        image = future.result()
        actual = image.getpixel((image.width // 2, image.height // 2))
        assert actual == expected, f"Expected {expected}, captured {actual}"
        print(f"HWND {hwnd:#x}, foreground blue, captured {actual}")
