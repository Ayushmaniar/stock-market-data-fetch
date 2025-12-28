"""
StockWatch - Main application entry point.

This script launches the StockWatch application.
"""

import sys
from PyQt5.QtWidgets import QApplication
from stockwatch.views.main_window import StockWatchlistApp

def main():
    """Launch the StockWatch application."""
    app = QApplication(sys.argv)
    window = StockWatchlistApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()