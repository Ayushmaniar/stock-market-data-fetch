"""UI components for the StockWatch application."""

from PyQt5.QtWidgets import QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QDesktopServices

class ClickableLabel(QLabel):
    """A custom QLabel that can be clicked to open a URL."""
    
    def __init__(self, text, url, parent=None):
        super().__init__(text, parent)
        self.url = url
        self.setStyleSheet("color: blue; text-decoration: underline;")
        self.setCursor(QCursor(Qt.PointingHandCursor))

        # Set a larger font size
        font = self.font()
        font.setPointSize(9)  # You can adjust this value as needed
        self.setFont(font)

    def mousePressEvent(self, event):
        """Handle mouse press events to open URL when clicked."""
        if event.button() == Qt.LeftButton:
            QDesktopServices.openUrl(QUrl(self.url))