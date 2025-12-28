"""Tests for the main application window."""

import unittest
from unittest.mock import MagicMock, patch
import os
import pandas as pd
from datetime import datetime

from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt, QCoreApplication

import sys
from stockwatch.views.main_window import StockWatchlistApp

# Create QApplication instance for testing Qt components
app = QApplication.instance()
if not app:
    app = QApplication(sys.argv)

class TestMainWindow(unittest.TestCase):
    """Test case for the main application window."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Patch QMessageBox to avoid showing actual dialog boxes
        self.patcher = patch('stockwatch.views.main_window.QMessageBox')
        self.mock_message_box = self.patcher.start()
        self.mock_message_box.warning.return_value = QMessageBox.Ok
        self.mock_message_box.information.return_value = QMessageBox.Ok
        self.mock_message_box.critical.return_value = QMessageBox.Ok
        self.mock_message_box.question.return_value = QMessageBox.No
        
        # Mock the data download thread
        self.thread_patcher = patch('stockwatch.views.main_window.DataDownloadThread')
        self.mock_thread_class = self.thread_patcher.start()
        self.mock_thread = MagicMock()
        self.mock_thread_class.return_value = self.mock_thread
        
        # Create test DataFrame
        self.test_df = pd.DataFrame({
            'SYMBOL': ['AAPL.NS', 'MSFT.NS', 'GOOGL.NS'],
            'Date': ['2025-04-05', '2025-04-05', '2025-04-05'],
            'Open': [150.0, 250.0, 2000.0],
            'High': [155.0, 255.0, 2050.0],
            'Low': [148.0, 248.0, 1950.0],
            'Close': [152.0, 252.0, 2020.0],
            'Adj Close': [152.0, 252.0, 2020.0],
            'Volume': [1000000, 500000, 200000],
            'Previous_Close': [151.0, 251.0, 2010.0],
            '1D': [0.66, 0.40, 0.50],
            '5D': [1.50, 1.20, 1.80],
            '1M': [3.20, 2.50, 4.00]
        })
        
        # Create a mock loader that returns our test DataFrame
        self.data_patcher = patch.object(StockWatchlistApp, 'load_data', return_value=self.test_df)
        self.data_patcher.start()
        
        # Create the main window
        self.window = StockWatchlistApp()
        
    def tearDown(self):
        """Tear down test fixtures."""
        self.patcher.stop()
        self.thread_patcher.stop()
        self.data_patcher.stop()
        
    def test_init_ui(self):
        """Test that the UI is initialized correctly."""
        # Check that the window title is set correctly
        self.assertEqual(self.window.windowTitle(), "Stock Watchlist")
        
        # Check that the table is created with the correct number of columns
        self.assertEqual(self.window.table.columnCount(), 13)
        
        # Check that the refresh buttons are created
        self.assertTrue(hasattr(self.window, 'refresh_all_button'))
        self.assertTrue(hasattr(self.window, 'quick_refresh_button'))
        
        # Check the visibility of progress indicators
        self.assertFalse(self.window.progress_bar.isVisible())
        self.assertFalse(self.window.status_text.isVisible())
        
    def test_add_stock(self):
        """Test adding a stock to the watchlist."""
        # Set up test input
        self.window.stock_input.setText("AAPL.NS")
        
        # Call the add stock function
        self.window.add_stock()
        
        # Check that the stock was added to the watchlist
        self.assertIn("AAPL.NS", self.window.watchlist)
        
        # Test adding a duplicate stock
        self.window.stock_input.setText("AAPL.NS")
        self.window.add_stock()
        
        # Check that a message was shown
        self.mock_message_box.information.assert_called_once()
        
        # Test adding an invalid stock
        self.window.stock_input.setText("INVALID.NS")
        self.window.add_stock()
        
        # Check that a warning was shown
        self.mock_message_box.warning.assert_called_once()
        
    def test_delete_stock(self):
        """Test removing a stock from the watchlist."""
        # Add a stock first
        self.window.watchlist.append("AAPL.NS")
        
        # Make sure it's there
        self.assertIn("AAPL.NS", self.window.watchlist)
        
        # Delete it
        self.window.delete_stock("AAPL.NS")
        
        # Check that it was removed
        self.assertNotIn("AAPL.NS", self.window.watchlist)
        
    def test_refresh_functionality(self):
        """Test the refresh functionality."""
        # Test refresh all data
        self.window.refresh_all_data()
        
        # Check that the data download thread was started with the right parameters
        self.mock_thread_class.assert_called_once()
        self.mock_thread.start.assert_called_once()
        
        # Reset mocks
        self.mock_thread_class.reset_mock()
        self.mock_thread.start.reset_mock()
        
        # Test quick refresh data
        self.window.watchlist = ["AAPL.NS"]
        self.window.quick_refresh_data()
        
        # Check that the data download thread was started with the right parameters
        self.mock_thread_class.assert_called_once_with(["AAPL.NS"], datetime.today().strftime('%Y-%m-%d'))
        self.mock_thread.start.assert_called_once()
        
        # Test quick refresh with no stocks
        self.mock_thread_class.reset_mock()
        self.mock_thread.start.reset_mock()
        self.window.watchlist = []
        self.window.quick_refresh_data()
        
        # Check that a warning was shown and no thread was started
        self.mock_message_box.warning.assert_called()
        self.mock_thread_class.assert_not_called()
        
    def test_save_load_watchlist(self):
        """Test saving and loading the watchlist."""
        # Create a test watchlist
        test_watchlist = ["AAPL.NS", "MSFT.NS", "GOOGL.NS"]
        self.window.watchlist = test_watchlist.copy()
        
        # Mock the file operations
        mock_open = unittest.mock.mock_open()
        with patch('builtins.open', mock_open):
            # Save the watchlist
            self.window.save_watchlist()
            
            # Check that open was called with the correct parameters
            mock_open.assert_called_once_with("watchlist.txt", "w")
            
            # Check that the correct data was written
            write_calls = [unittest.mock.call(f"{stock}\n") for stock in test_watchlist]
            mock_open().write.assert_has_calls(write_calls)
            
            # Check that a message was shown
            self.mock_message_box.information.assert_called_once()
            
            # Reset mocks
            mock_open.reset_mock()
            self.mock_message_box.information.reset_mock()
            
            # Create a new window to test loading
            self.window = StockWatchlistApp()
            
            # Set up the mock file to return the test watchlist
            mock_open().readlines.return_value = [f"{stock}\n" for stock in test_watchlist]
            
            # Load the watchlist
            with patch('builtins.open', mock_open):
                self.window.load_watchlist()
                
                # Check that the watchlist was loaded correctly
                self.assertEqual(self.window.watchlist, test_watchlist)

if __name__ == '__main__':
    unittest.main()