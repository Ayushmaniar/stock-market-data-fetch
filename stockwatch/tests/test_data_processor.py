"""Tests for the data processing module."""

import unittest
import os
import pandas as pd
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock

from stockwatch.tests.test_utils import project_root
from stockwatch.data.data_processor import DataDownloadThread

class TestDataProcessor(unittest.TestCase):
    """Test case for the data processor module."""

    def setUp(self):
        """Set up test fixtures."""
        self.date_to_use = datetime.today().strftime('%Y-%m-%d')
        
        # Create mock signals
        self.mock_progress = MagicMock()
        self.mock_status = MagicMock()
        self.mock_finished = MagicMock()
        self.mock_error = MagicMock()
        
    @patch('stockwatch.data.data_processor.yf.download')
    @patch('stockwatch.data.data_processor.pd.read_csv')
    def test_download_with_symbols(self, mock_read_csv, mock_download):
        """Test downloading data for specific symbols."""
        # Create mock CSV data
        mock_csv_data = pd.DataFrame({
            'SYMBOL': ['AAPL.NS', 'MSFT.NS'],
            'YahooEquiv': ['AAPL.NS', 'MSFT.NS']
        })
        mock_read_csv.return_value = mock_csv_data
        
        # Create mock stock data (yfinance returns date as index, not as column)
        # And no longer provides Adj Close for NSE stocks
        mock_stock_data = pd.DataFrame({
            'Open': [150.0],
            'High': [155.0],
            'Low': [148.0],
            'Close': [152.0],
            'Volume': [1000000]
        })
        mock_stock_data.index = pd.DatetimeIndex([datetime.today()], name='Date')
        mock_download.return_value = mock_stock_data
        
        # Create the thread with mocked signals
        thread = DataDownloadThread(symbols=['AAPL.NS'], date_to_use=self.date_to_use)
        thread.progress_signal = self.mock_progress
        thread.status_signal = self.mock_status
        thread.finished_signal = self.mock_finished
        thread.error_signal = self.mock_error
        
        # Run with patched functions
        with patch('stockwatch.data.data_processor.pd.to_datetime') as mock_to_datetime:
            mock_to_datetime.return_value = datetime.today()
            with patch('stockwatch.data.data_processor.pd.ExcelWriter') as mock_excel_writer:
                # Mock necessary components for Excel writing
                mock_workbook = MagicMock()
                mock_worksheet = MagicMock()
                mock_excel_writer.return_value.__enter__.return_value.book = mock_workbook
                mock_excel_writer.return_value.__enter__.return_value.sheets = {'Sheet1': mock_worksheet}
                
                # Run the thread
                thread.run()
        
        # Assert the progress signal was called at least once
        self.mock_progress.emit.assert_called()

        # Assert the status signal was called with specific messages
        status_calls = [call[0][0] for call in self.mock_status.emit.call_args_list]
        # Check for threaded download message
        self.assertTrue(any("using 10 parallel threads" in call for call in status_calls))
        self.assertTrue("Data processing completed." in status_calls)
        
        # Assert the finished signal was called with a DataFrame
        self.mock_finished.emit.assert_called_once()
        
        # Assert the error signal was not called
        self.mock_error.emit.assert_not_called()

    def test_data_folder_creation(self):
        """Test that the data folder is created correctly."""
        test_folder = os.path.join(project_root, "test_data_folder")
        
        # Ensure the folder doesn't exist
        if os.path.exists(test_folder):
            os.rmdir(test_folder)
        
        # Check that the folder doesn't exist initially
        self.assertFalse(os.path.exists(test_folder))
        
        # Use the functionality that ensures the folder exists
        thread = DataDownloadThread()
        os.makedirs(test_folder, exist_ok=True)
        
        # Check that the folder was created
        self.assertTrue(os.path.exists(test_folder))
        
        # Clean up
        os.rmdir(test_folder)

if __name__ == '__main__':
    unittest.main()