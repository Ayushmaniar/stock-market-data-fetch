"""
Daily Stock Data Download Script

A simple command-line script to download stock data without opening the GUI.
Perfect for scheduling with Task Scheduler (Windows) or cron (Linux/Mac).

Usage:
    python download_daily.py              # Download all stocks for today
    python download_daily.py --date 2025-12-25  # Download for specific date
"""

import sys
import os
import argparse
from datetime import datetime
import pandas as pd
from pandas.tseries.offsets import BDay

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from stockwatch.data.data_processor import DataDownloadThread


class DownloadProgressTracker:
    """Simple progress tracker for command-line downloads."""

    def __init__(self):
        self.last_progress = 0

    def on_progress(self, value):
        """Handle progress updates."""
        if value % 10 == 0 and value != self.last_progress:
            print(f"Progress: {value}%")
            self.last_progress = value

    def on_status(self, message):
        """Handle status messages."""
        # Only print important messages to reduce clutter
        if any(keyword in message.lower() for keyword in ['downloading data for', 'completed', 'failed', 'error']):
            if 'Downloading data for' in message:
                # Don't print every single stock
                return
            print(message)

    def on_finished(self, dataframe):
        """Handle completion."""
        print(f"\n{'='*60}")
        print("DOWNLOAD COMPLETE!")
        print(f"{'='*60}")
        print(f"Total stocks downloaded: {len(dataframe)}")
        print(f"Data saved to yahoo_finance_data/ folder")
        print(f"{'='*60}\n")

    def on_error(self, error_message):
        """Handle errors."""
        print(f"\n{'='*60}")
        print("ERROR!")
        print(f"{'='*60}")
        print(f"Error: {error_message}")
        print(f"{'='*60}\n")


def download_stock_data(date_str=None):
    """
    Download stock data for all NSE stocks.

    Args:
        date_str: Date string in YYYY-MM-DD format. If None, uses today's date.

    Returns:
        bool: True if download was successful, False otherwise.
    """
    # Determine the date to use
    if date_str:
        try:
            date_to_use = datetime.strptime(date_str, '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            print(f"Error: Invalid date format '{date_str}'. Use YYYY-MM-DD format.")
            return False
    else:
        today = datetime.today()
        pandas_today = pd.Timestamp(today)

        # Handle weekends
        if today.weekday() >= 5:  # Saturday or Sunday
            date_to_use = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
            print(f"Today is {today.strftime('%A')}. Using last business day: {date_to_use}")
        else:
            date_to_use = pandas_today.strftime('%Y-%m-%d')

    print(f"\n{'='*60}")
    print("STOCK DATA DOWNLOAD SCRIPT")
    print(f"{'='*60}")
    print(f"Date: {date_to_use}")
    print(f"Started at: {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'='*60}\n")

    # Create progress tracker
    tracker = DownloadProgressTracker()

    # Create download thread
    thread = DataDownloadThread(symbols=None, date_to_use=date_to_use)

    # Connect signals to tracker
    thread.progress_signal.connect(tracker.on_progress)
    thread.status_signal.connect(tracker.on_status)
    thread.finished_signal.connect(tracker.on_finished)
    thread.error_signal.connect(tracker.on_error)

    # Track if there was an error
    error_occurred = [False]

    def on_error_wrapper(msg):
        error_occurred[0] = True
        tracker.on_error(msg)

    thread.error_signal.connect(on_error_wrapper)

    # Run the download (synchronously)
    print("Starting download... This may take 10-15 minutes for all stocks.\n")
    thread.run()

    # Return success status
    return not error_occurred[0]


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description='Download NSE stock data for a specific date.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python download_daily.py                    # Download today's data
  python download_daily.py --date 2025-12-25  # Download for specific date

Note: This script downloads data for ALL NSE stocks (~1500+) and takes 10-15 minutes.
        """
    )
    parser.add_argument(
        '--date',
        type=str,
        help='Date to download data for (format: YYYY-MM-DD). Defaults to today.',
        default=None
    )

    args = parser.parse_args()

    # Run the download
    success = download_stock_data(args.date)

    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
