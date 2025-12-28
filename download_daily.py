"""
Daily Stock Data Download Script

A simple command-line script to download stock data without opening the GUI.
Perfect for scheduling with Task Scheduler (Windows) or cron (Linux/Mac).

Features:
- Parallel downloading with multiple threads (3x faster!)
- Automatic weekend handling
- Progress tracking

Usage:
    python download_daily.py                      # Download all stocks for today (3 threads)
    python download_daily.py --date 2025-12-25    # Download for specific date
    python download_daily.py --threads 5          # Use 5 parallel threads
"""

import sys
import os
import argparse
from datetime import datetime
import pandas as pd
from pandas.tseries.offsets import BDay
from PyQt5.QtCore import QCoreApplication, QTimer
from urllib.parse import quote_plus
import threading

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from stockwatch.data.data_processor import DataDownloadThread


class ParallelDownloadManager:
    """Manages parallel downloads using multiple threads."""

    def __init__(self, num_threads=3, date_to_use=None):
        """
        Initialize the parallel download manager.

        Args:
            num_threads: Number of parallel download threads to use
            date_to_use: Date string in YYYY-MM-DD format
        """
        self.num_threads = num_threads
        self.date_to_use = date_to_use
        self.threads = []
        self.results = []
        self.completed_threads = 0
        self.lock = threading.Lock()
        self.error_occurred = False
        self.app = None

    def split_symbols(self, symbols):
        """Split symbols into N roughly equal chunks."""
        chunk_size = len(symbols) // self.num_threads
        remainder = len(symbols) % self.num_threads

        chunks = []
        start = 0
        for i in range(self.num_threads):
            # Distribute remainder across first few chunks
            size = chunk_size + (1 if i < remainder else 0)
            chunks.append(symbols[start:start + size])
            start += size

        return chunks

    def on_progress(self, thread_id, value):
        """Handle progress updates from individual threads."""
        if value % 20 == 0:
            print(f"  Thread {thread_id + 1}: {value}%")

    def on_status(self, thread_id, message):
        """Handle status messages from individual threads."""
        # Only print important messages
        if 'completed' in message.lower() or 'error' in message.lower():
            print(f"  Thread {thread_id + 1}: {message}")

    def on_finished(self, thread_id, dataframe):
        """Handle completion of individual thread."""
        with self.lock:
            self.results.append(dataframe)
            self.completed_threads += 1
            print(f"\n{'='*60}")
            print(f"Thread {thread_id + 1} COMPLETE! ({len(dataframe)} stocks)")
            print(f"{'='*60}")

            if self.completed_threads == self.num_threads:
                # All threads completed
                QTimer.singleShot(0, self.save_combined_results)

    def on_error(self, thread_id, error_message):
        """Handle errors from individual threads."""
        print(f"\n{'='*60}")
        print(f"Thread {thread_id + 1} ERROR!")
        print(f"{'='*60}")
        print(f"Error: {error_message}")
        print(f"{'='*60}\n")

        with self.lock:
            self.error_occurred = True
            self.completed_threads += 1

            if self.completed_threads == self.num_threads:
                QTimer.singleShot(0, self.save_combined_results)

    def save_combined_results(self):
        """Combine results from all threads and save to Excel."""
        try:
            if not self.results:
                print("\nNo data to save!")
                self.app.quit()
                return

            print(f"\n{'='*60}")
            print("COMBINING RESULTS FROM ALL THREADS...")
            print(f"{'='*60}")

            # Combine all dataframes
            all_stock_data = pd.concat(self.results, ignore_index=True)

            print(f"Total stocks: {len(all_stock_data)}")
            print(f"Sorting by 1D performance...")

            # Sort by 1D performance
            try:
                all_stock_data.sort_values(by='1D', ascending=False, inplace=True)
            except KeyError:
                print("Warning: Could not sort by 1D (column missing)")

            # Save to Excel
            os.makedirs("yahoo_finance_data", exist_ok=True)
            output_file_name = f'yahoo_finance_data/{self.date_to_use}_stock_market_data.xlsx'

            print(f"Saving to {output_file_name}...")

            # Create Google search URLs
            try:
                google_search_urls = 'https://www.google.com/search?q=' + all_stock_data['SYMBOL'].str.replace('.NS','').apply(quote_plus) + '+share+price'
                google_search_urls = google_search_urls.sort_index()
            except Exception as e:
                print(f"Warning: Could not create Google search URLs: {e}")
                google_search_urls = None

            # Write to Excel with formatting
            with pd.ExcelWriter(output_file_name, engine='xlsxwriter') as writer:
                all_stock_data.to_excel(writer, index=False, sheet_name='Sheet1')

                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Add clickable links for stock symbols
                if google_search_urls is not None:
                    for i, url in enumerate(google_search_urls):
                        cell = f'A{i+2}'
                        worksheet.write_url(cell, url, string=all_stock_data.loc[i, 'SYMBOL'])

                # Add red background for close column
                red_format = workbook.add_format({'bg_color': '#FFC7CE'})
                if 'Close' in all_stock_data.columns:
                    close_col_idx = all_stock_data.columns.get_loc('Close')
                    close_col_letter = chr(ord('A') + close_col_idx)
                    worksheet.conditional_format(f'{close_col_letter}2:{close_col_letter}{len(all_stock_data) + 1}',
                                                {'type': 'no_blanks', 'format': red_format})

            print(f"\n{'='*60}")
            print("DOWNLOAD COMPLETE!")
            print(f"{'='*60}")
            print(f"Total stocks downloaded: {len(all_stock_data)}")
            print(f"Data saved to: {output_file_name}")
            print(f"{'='*60}\n")

        except Exception as e:
            print(f"\nError saving combined results: {e}")
            import traceback
            traceback.print_exc()
            self.error_occurred = True
        finally:
            self.app.quit()

    def download(self):
        """Start parallel downloads."""
        # Create QCoreApplication (required for QThread)
        self.app = QCoreApplication(sys.argv)

        # Load stock symbols
        try:
            # TODO: Temporary - using short version for testing. Revert to EQUITY_L.csv later
            csv_path = "EQUITY_L_short.csv"
            if not os.path.exists(csv_path):
                print(f"Error: {csv_path} not found!")
                return False

            stocks_df = pd.read_csv(csv_path)
            stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'
            all_symbols = list(stocks_df['YahooEquiv'])

            print(f"Total stocks to download: {len(all_symbols)}")
            print(f"Using {self.num_threads} parallel threads")
            print(f"Approximately {len(all_symbols) // self.num_threads} stocks per thread\n")

        except Exception as e:
            print(f"Error loading stock symbols: {e}")
            return False

        # Split symbols into chunks
        symbol_chunks = self.split_symbols(all_symbols)

        # Create and start download threads
        for i, chunk in enumerate(symbol_chunks):
            print(f"Starting Thread {i + 1} with {len(chunk)} stocks...")

            thread = DataDownloadThread(symbols=chunk, date_to_use=self.date_to_use)

            # Connect signals
            thread.progress_signal.connect(lambda val, tid=i: self.on_progress(tid, val))
            thread.status_signal.connect(lambda msg, tid=i: self.on_status(tid, msg))
            thread.finished_signal.connect(lambda df, tid=i: self.on_finished(tid, df))
            thread.error_signal.connect(lambda msg, tid=i: self.on_error(tid, msg))

            self.threads.append(thread)
            thread.start()

        print(f"\n{'='*60}")
        print("All threads started! Downloading in parallel...")
        print(f"{'='*60}\n")

        # Run the event loop
        self.app.exec_()

        return not self.error_occurred


def download_stock_data(date_str=None, num_threads=3):
    """
    Download stock data for all NSE stocks using parallel threads.

    Args:
        date_str: Date string in YYYY-MM-DD format. If None, uses today's date.
        num_threads: Number of parallel download threads to use (default: 3)

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
    print("PARALLEL STOCK DATA DOWNLOAD")
    print(f"{'='*60}")
    print(f"Date: {date_to_use}")
    print(f"Threads: {num_threads}")
    print(f"Started at: {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'='*60}\n")

    # Create manager and start download
    manager = ParallelDownloadManager(num_threads=num_threads, date_to_use=date_to_use)
    success = manager.download()

    if success:
        print(f"\nFinished at: {datetime.now().strftime('%H:%M:%S')}")

    return success


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description='Download NSE stock data using parallel threads.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python download_daily.py                      # Download today's data (3 threads)
  python download_daily.py --date 2025-12-25    # Download for specific date
  python download_daily.py --threads 5          # Use 5 parallel threads

Note: More threads = faster download, but higher CPU/network usage.
      Recommended: 3-5 threads for optimal balance.
        """
    )
    parser.add_argument(
        '--date',
        type=str,
        help='Date to download data for (format: YYYY-MM-DD). Defaults to today.',
        default=None
    )
    parser.add_argument(
        '--threads',
        type=int,
        help='Number of parallel download threads (default: 3)',
        default=3
    )

    args = parser.parse_args()

    # Validate threads
    if args.threads < 1 or args.threads > 10:
        print("Error: Number of threads must be between 1 and 10")
        sys.exit(1)

    # Run the download
    success = download_stock_data(args.date, args.threads)

    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
