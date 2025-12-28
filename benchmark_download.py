"""
Download Speed Benchmark - Sequential vs Threaded

Tests the download speed difference between sequential and parallel downloads
for a subset of stocks (200 by default).

Usage:
    python benchmark_download.py
    python benchmark_download.py --stocks 100  # Test with 100 stocks
"""

import pandas as pd
import yfinance as yf
import time
import argparse
from datetime import datetime
from pandas.tseries.offsets import BDay
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def download_single_stock(symbol, start_date, end_date, max_retries=3, retry_delay=1):
    """Download data for a single stock with retry logic."""
    for attempt in range(max_retries):
        try:
            data = yf.download(symbol, start=start_date, end=end_date, progress=False)

            if isinstance(data.columns, pd.MultiIndex):
                data.columns = [f"{col[0]}" for col in data.columns]

            # Ensure columns are in the expected order
            if not data.empty:
                expected_cols = ['Open', 'High', 'Low', 'Close', 'Volume']
                available_cols = [col for col in expected_cols if col in data.columns]
                if available_cols:
                    data = data[available_cols]

            return symbol, data

        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
            else:
                return symbol, pd.DataFrame()

    return symbol, pd.DataFrame()


def download_sequential(symbols, start_date, end_date):
    """Download stocks sequentially (current method)."""
    print(f"\n{'='*60}")
    print("SEQUENTIAL DOWNLOAD")
    print(f"{'='*60}")
    print(f"Downloading {len(symbols)} stocks...")

    start_time = time.time()
    stock_data = {}
    failed = []

    for i, symbol in enumerate(symbols, 1):
        if i % 20 == 0:
            elapsed = time.time() - start_time
            rate = i / elapsed
            remaining = (len(symbols) - i) / rate
            print(f"Progress: {i}/{len(symbols)} ({i*100//len(symbols)}%) - "
                  f"Rate: {rate:.1f} stocks/sec - ETA: {remaining:.0f}s")

        time.sleep(0.01)  # Small delay like in original code
        symbol_name, data = download_single_stock(symbol, start_date, end_date)

        if not data.empty:
            stock_data[symbol_name] = data
        else:
            failed.append(symbol_name)

    elapsed_time = time.time() - start_time

    print(f"\n{'='*60}")
    print(f"Sequential Download Complete")
    print(f"{'='*60}")
    print(f"Time taken: {elapsed_time:.2f} seconds ({elapsed_time/60:.2f} minutes)")
    print(f"Successfully downloaded: {len(stock_data)} stocks")
    print(f"Failed: {len(failed)} stocks")
    print(f"Download rate: {len(symbols)/elapsed_time:.2f} stocks/second")
    print(f"{'='*60}\n")

    return stock_data, elapsed_time


def download_threaded(symbols, start_date, end_date, max_workers=10):
    """Download stocks in parallel using ThreadPoolExecutor."""
    print(f"\n{'='*60}")
    print(f"THREADED DOWNLOAD (Workers: {max_workers})")
    print(f"{'='*60}")
    print(f"Downloading {len(symbols)} stocks...")

    start_time = time.time()
    stock_data = {}
    failed = []
    completed = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all download tasks
        future_to_symbol = {
            executor.submit(download_single_stock, symbol, start_date, end_date): symbol
            for symbol in symbols
        }

        # Process completed downloads
        for future in as_completed(future_to_symbol):
            completed += 1

            if completed % 20 == 0:
                elapsed = time.time() - start_time
                rate = completed / elapsed
                remaining = (len(symbols) - completed) / rate
                print(f"Progress: {completed}/{len(symbols)} ({completed*100//len(symbols)}%) - "
                      f"Rate: {rate:.1f} stocks/sec - ETA: {remaining:.0f}s")

            try:
                symbol_name, data = future.result()
                if not data.empty:
                    stock_data[symbol_name] = data
                else:
                    failed.append(symbol_name)
            except Exception as e:
                symbol = future_to_symbol[future]
                failed.append(symbol)

    elapsed_time = time.time() - start_time

    print(f"\n{'='*60}")
    print(f"Threaded Download Complete")
    print(f"{'='*60}")
    print(f"Time taken: {elapsed_time:.2f} seconds ({elapsed_time/60:.2f} minutes)")
    print(f"Successfully downloaded: {len(stock_data)} stocks")
    print(f"Failed: {len(failed)} stocks")
    print(f"Download rate: {len(symbols)/elapsed_time:.2f} stocks/second")
    print(f"{'='*60}\n")

    return stock_data, elapsed_time


def run_benchmark(num_stocks=200, max_workers=10):
    """Run the benchmark comparison."""
    print(f"\n{'#'*60}")
    print(f"DOWNLOAD SPEED BENCHMARK")
    print(f"{'#'*60}")
    print(f"Test size: {num_stocks} stocks")
    print(f"Thread workers: {max_workers}")
    print(f"Started at: {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'#'*60}\n")

    # Load stock symbols
    csv_path = "EQUITY_L.csv"
    if not os.path.exists(csv_path):
        print(f"Error: {csv_path} not found!")
        return

    stocks_df = pd.read_csv(csv_path)
    stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'

    # Get first N stocks
    test_symbols = list(stocks_df['YahooEquiv'].head(num_stocks))
    print(f"Testing with {len(test_symbols)} stocks: {test_symbols[0]} to {test_symbols[-1]}\n")

    # Calculate dates
    today = datetime.today()
    pandas_today = pd.Timestamp(today)

    if today.weekday() >= 5:
        date_to_use = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
    else:
        date_to_use = pandas_today.strftime('%Y-%m-%d')

    one_month_ago = pandas_today - pd.DateOffset(months=1)
    if one_month_ago.weekday() > 4:
        one_month_ago = one_month_ago + pd.DateOffset(days=(7 - one_month_ago.weekday()))
    one_month_ago = one_month_ago.strftime('%Y-%m-%d')

    print(f"Date range: {one_month_ago} to {date_to_use}\n")

    # Test 1: Sequential download
    seq_data, seq_time = download_sequential(test_symbols, one_month_ago, date_to_use)

    # Wait a bit before next test
    print("Waiting 5 seconds before threaded test...\n")
    time.sleep(5)

    # Test 2: Threaded download
    thread_data, thread_time = download_threaded(test_symbols, one_month_ago, date_to_use, max_workers)

    # Compare results
    print(f"\n{'#'*60}")
    print(f"BENCHMARK RESULTS")
    print(f"{'#'*60}")
    print(f"Sequential Time:  {seq_time:.2f}s ({seq_time/60:.2f} min)")
    print(f"Threaded Time:    {thread_time:.2f}s ({thread_time/60:.2f} min)")
    print(f"\nSpeedup:          {seq_time/thread_time:.2f}x faster")
    print(f"Time saved:       {seq_time - thread_time:.2f}s ({(seq_time - thread_time)/60:.2f} min)")
    print(f"\nSequential Rate:  {num_stocks/seq_time:.2f} stocks/sec")
    print(f"Threaded Rate:    {num_stocks/thread_time:.2f} stocks/sec")

    # Estimate time for full download
    total_stocks = len(stocks_df)
    est_seq_time = (seq_time / num_stocks) * total_stocks
    est_thread_time = (thread_time / num_stocks) * total_stocks

    print(f"\n{'='*60}")
    print(f"ESTIMATED TIME FOR ALL {total_stocks} STOCKS:")
    print(f"{'='*60}")
    print(f"Sequential:       {est_seq_time/60:.1f} minutes")
    print(f"Threaded:         {est_thread_time/60:.1f} minutes")
    print(f"Time saved:       {(est_seq_time - est_thread_time)/60:.1f} minutes")
    print(f"{'#'*60}\n")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Benchmark sequential vs threaded stock downloads.',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        '--stocks',
        type=int,
        default=200,
        help='Number of stocks to test with (default: 200)'
    )
    parser.add_argument(
        '--workers',
        type=int,
        default=10,
        help='Number of thread workers (default: 10)'
    )

    args = parser.parse_args()

    run_benchmark(args.stocks, args.workers)


if __name__ == "__main__":
    main()
