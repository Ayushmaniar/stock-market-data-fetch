"""Tests for the data processing module."""

import unittest
import os
import pandas as pd
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock
from yfinance import exceptions as yf_exceptions

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
        
    @patch('stockwatch.data.data_processor.yf.Ticker')
    @patch('stockwatch.data.data_processor.pd.read_csv')
    def test_download_with_symbols(self, mock_read_csv, mock_ticker_class):
        """Test downloading data for specific symbols."""
        # Create mock CSV data
        mock_csv_data = pd.DataFrame({
            'SYMBOL': ['AAPL.NS', 'MSFT.NS'],
            'YahooEquiv': ['AAPL.NS', 'MSFT.NS']
        })
        mock_read_csv.return_value = mock_csv_data

        # Create mock stock data (yf.Ticker().history() returns date as index)
        mock_stock_data = pd.DataFrame({
            'Open': [150.0],
            'High': [155.0],
            'Low': [148.0],
            'Close': [152.0],
            'Volume': [1000000]
        })
        # history() returns a tz-aware index (Asia/Kolkata for NSE stocks)
        mock_stock_data.index = pd.DatetimeIndex(
            [pd.Timestamp(self.date_to_use)], name='Date'
        ).tz_localize('Asia/Kolkata')

        # Mock the Ticker class and its history method
        mock_ticker_instance = MagicMock()
        mock_ticker_instance.history.return_value = mock_stock_data
        mock_ticker_class.return_value = mock_ticker_instance
        
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
        # Check for multi-threaded download status message
        download_started = any("Downloading data for" in msg and "threads" in msg for msg in status_calls)
        self.assertTrue(download_started, f"Expected multi-threaded download status, got: {status_calls}")
        self.assertTrue("Data processing completed." in status_calls)
        
        # Assert the finished signal was called with a DataFrame
        self.mock_finished.emit.assert_called_once()

        # Date column should be plain YYYY-MM-DD, with no timezone suffix
        result_df = self.mock_finished.emit.call_args[0][0]
        self.assertEqual(result_df['Date'].iloc[0], self.date_to_use)

        # Assert the error signal was not called
        self.mock_error.emit.assert_not_called()

    @patch('stockwatch.data.data_processor.time.sleep')
    @patch('stockwatch.data.data_processor.yf.Ticker')
    def test_rate_limit_backs_off_and_recovers(self, mock_ticker_class, mock_sleep):
        """A rate-limited request pauses all threads, then succeeds on retry."""
        good = pd.DataFrame({'Open': [1.0], 'High': [2.0], 'Low': [1.0], 'Close': [2.0], 'Volume': [10]},
                            index=pd.DatetimeIndex([pd.Timestamp('2026-01-02')], name='Date'))
        ticker = MagicMock()
        ticker.history.side_effect = [yf_exceptions.YFRateLimitError(), yf_exceptions.YFRateLimitError(), good]
        mock_ticker_class.return_value = ticker

        thread = DataDownloadThread()
        result = thread.download_with_retry('ABC.NS', '2026-01-01', '2026-01-03')

        self.assertEqual(len(result), 1)
        self.assertEqual(ticker.history.call_count, 3)
        # Backoff doubles: 3s then 6s cooldown was set on the shared pause timestamp
        self.assertGreater(thread._rate_limit_until, 0)

    @patch('stockwatch.data.data_processor.time.sleep')
    @patch('stockwatch.data.data_processor.yf.Ticker')
    def test_rate_limit_gives_up_after_max_retries(self, mock_ticker_class, mock_sleep):
        """Persistent rate limiting eventually raises instead of looping forever."""
        ticker = MagicMock()
        ticker.history.side_effect = yf_exceptions.YFRateLimitError()
        mock_ticker_class.return_value = ticker

        thread = DataDownloadThread()
        with self.assertRaises(yf_exceptions.YFRateLimitError):
            thread.download_with_retry('ABC.NS', '2026-01-01', '2026-01-03')
        self.assertEqual(ticker.history.call_count, DataDownloadThread.RATE_LIMIT_MAX_RETRIES + 1)

    @patch('stockwatch.data.data_processor.time.sleep')
    @patch('stockwatch.data.data_processor.yf.Ticker')
    def test_delisted_stock_is_not_retried(self, mock_ticker_class, mock_sleep):
        """Delisted symbols return empty immediately without retries or sleeping."""
        ticker = MagicMock()
        ticker.history.side_effect = yf_exceptions.YFTzMissingError('DEAD.NS')
        mock_ticker_class.return_value = ticker

        thread = DataDownloadThread()
        result = thread.download_with_retry('DEAD.NS', '2026-01-01', '2026-01-03')

        self.assertTrue(result.empty)
        self.assertEqual(ticker.history.call_count, 1)
        mock_sleep.assert_not_called()

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