import ctypes

from PIL import ImageGrab
from PyQt5.QtCore import QPoint, QRect, Qt
from PyQt5.QtGui import QPainter, QPen
from PyQt5.QtWidgets import QDialog

from img_convertor import pil_image_to_qpixmap


def desktop_snapshot():
    user32 = ctypes.windll.user32
    origin = QPoint(user32.GetSystemMetrics(76), user32.GetSystemMetrics(77))
    return ImageGrab.grab(all_screens=True), origin


class RegionPicker(QDialog):
    """Select physical desktop pixels on an aspect-preserving desktop preview."""

    def __init__(self, snapshot, origin, parent=None):
        super().__init__(parent)
        self.setWindowTitle("拖动框选视频画面 · Esc 取消")
        self.setCursor(Qt.CrossCursor)
        self.pixmap = pil_image_to_qpixmap(snapshot)
        self.origin = origin
        self.anchor = None
        self.selection = QRect()
        self.bounds = None
        self.resize(960, 600)

    def image_rect(self):
        size = self.pixmap.size().scaled(self.size(), Qt.KeepAspectRatio)
        rect = QRect(QPoint(), size)
        rect.moveCenter(self.rect().center())
        return rect

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), Qt.darkGray)
        painter.drawPixmap(self.image_rect(), self.pixmap)
        painter.setPen(QPen(Qt.yellow, 2))
        painter.drawRect(self.selection)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.image_rect().contains(event.pos()):
            self.anchor = event.pos()
            self.selection = QRect(self.anchor, self.anchor)
            self.update()

    def mouseMoveEvent(self, event):
        if self.anchor is not None:
            self.selection = QRect(self.anchor, event.pos()).normalized().intersected(self.image_rect())
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.LeftButton or self.anchor is None:
            return
        self.mouseMoveEvent(event)
        self.anchor = None
        rect = self.image_rect()
        selected = self.selection.translated(-rect.topLeft())
        if selected.width() < 3 or selected.height() < 3:
            return
        sx = self.pixmap.width() / rect.width()
        sy = self.pixmap.height() / rect.height()
        self.bounds = (
            self.origin.x() + round(selected.left() * sx),
            self.origin.y() + round(selected.top() * sy),
            self.origin.x() + round((selected.right() + 1) * sx),
            self.origin.y() + round((selected.bottom() + 1) * sy),
        )
        self.accept()
