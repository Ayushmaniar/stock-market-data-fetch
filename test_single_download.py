"""
Quick test script to verify if single stock download works correctly.
Tests the same download logic used in the main application.
"""

import pandas as pd
import yfinance as yf
import time
from datetime import datetime
from pandas.tseries.offsets import BDay
import traceback

def download_with_retry(symbol, start_date, end_date, max_retries=3, retry_delay=2):
    """Download data with retry mechanism to handle transient errors"""
    print(f"\n{'='*60}")
    print(f"Attempting to download: {symbol}")
    print(f"Date range: {start_date} to {end_date}")
    print(f"{'='*60}\n")

    for attempt in range(max_retries):
        try:
            print(f"Attempt {attempt + 1}/{max_retries}...")
            start_time = time.time()

            data = yf.download(
                symbol,
                start=start_date,
                end=end_date,
                progress=True  # Show progress for debugging
            )

            download_time = time.time() - start_time
            print(f"Download completed in {download_time:.2f} seconds")

            # Handle MultiIndex columns
            if isinstance(data.columns, pd.MultiIndex):
                print("Converting MultiIndex columns to single level")
                data.columns = [f"{col[0]}" for col in data.columns]

            if data.empty:
                print(f"WARNING: Empty dataframe received for {symbol}")
                if attempt < max_retries - 1:
                    print(f"Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    continue
                else:
                    print(f"FAILED: No data after {max_retries} attempts")
                    return None

            print(f"SUCCESS: Downloaded {len(data)} rows of data")
            print(f"Date range in data: {data.index.min()} to {data.index.max()}")
            print(f"Columns: {data.columns.tolist()}")
            return data

        except Exception as e:
            error_msg = f"Error for {symbol} (attempt {attempt + 1}): {str(e)}"
            print(error_msg)
            print(f"Traceback: {traceback.format_exc()}")

            if attempt < max_retries - 1:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                print(f"FAILED: Could not download {symbol} after {max_retries} attempts")
                return None

    return None

def test_single_stock(symbol_name="RELIANCE"):
    """Test downloading data for a single stock"""
    print("\n" + "="*60)
    print("STOCK DOWNLOAD TEST SCRIPT")
    print("="*60)

    # Prepare the symbol (add .NS for NSE stocks)
    yahoo_symbol = f"{symbol_name}.NS"

    # Calculate dates (same logic as main app)
    today = datetime.today()
    pandas_today = pd.Timestamp(today)

    # Get date to use (handle weekends)
    if today.weekday() >= 5:  # Weekend
        date_to_use = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
        print(f"Today is weekend, using last business day: {date_to_use}")
    else:
        date_to_use = pandas_today.strftime('%Y-%m-%d')
        print(f"Using today's date: {date_to_use}")

    # Calculate historical dates for returns calculation
    last_trading_day = (pd.Timestamp(date_to_use) - BDay(1)).strftime('%Y-%m-%d')
    five_trading_days_ago = (pd.Timestamp(date_to_use) - BDay(5)).strftime('%Y-%m-%d')

    one_month_ago = pd.Timestamp(date_to_use) - pd.DateOffset(months=1)
    if one_month_ago.weekday() > 4:
        one_month_ago = one_month_ago + pd.DateOffset(days=(7 - one_month_ago.weekday()))
    one_month_ago = one_month_ago.strftime('%Y-%m-%d')

    print(f"\nCalculated dates:")
    print(f"  Target date: {date_to_use}")
    print(f"  1D ago: {last_trading_day}")
    print(f"  5D ago: {five_trading_days_ago}")
    print(f"  1M ago: {one_month_ago}")

    # Download data
    start_download = time.time()
    data = download_with_retry(
        yahoo_symbol,
        start_date=one_month_ago,
        end_date=pd.to_datetime(date_to_use) + pd.Timedelta(days=1)
    )
    total_time = time.time() - start_download

    if data is None or data.empty:
        print("\n" + "="*60)
        print("TEST FAILED: Could not download data")
        print("="*60)
        return False

    # Process the data (same as main app)
    print("\n" + "-"*60)
    print("PROCESSING DATA")
    print("-"*60)

    data = data.reset_index()
    data['SYMBOL'] = yahoo_symbol

    print(f"\nAvailable dates in downloaded data:")
    print(data['Date'].dt.strftime('%Y-%m-%d').tolist())

    # Find the target date or closest available
    if date_to_use in data['Date'].dt.strftime('%Y-%m-%d').tolist():
        target_date = date_to_use
        print(f"\nUsing requested date: {target_date}")
    else:
        available_dates = data['Date'].dt.strftime('%Y-%m-%d').tolist()
        available_dates.sort(reverse=True)
        if available_dates:
            target_date = available_dates[0]
            print(f"\nRequested date not available, using most recent: {target_date}")
        else:
            print("\nERROR: No valid dates available")
            return False

    # Get the row for target date
    single_row = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == target_date]

    if single_row.empty:
        print(f"\nERROR: No data for target date {target_date}")
        return False

    # Calculate returns
    close_col = 'Close'
    todays_close = single_row[close_col].values[0]

    prev_close_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == last_trading_day]
    prev_close = prev_close_data[close_col].values[0] if not prev_close_data.empty else None

    five_days_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == five_trading_days_ago]
    five_days_close = five_days_data[close_col].values[0] if not five_days_data.empty else None

    one_month_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == one_month_ago]
    one_month_close = one_month_data[close_col].values[0] if not one_month_data.empty else None

    # Calculate percentage changes
    change_1d = ((todays_close - prev_close) / prev_close * 100) if prev_close else None
    change_5d = ((todays_close - five_days_close) / five_days_close * 100) if five_days_close else None
    change_1m = ((todays_close - one_month_close) / one_month_close * 100) if one_month_close else None

    # Display results
    print("\n" + "="*60)
    print("TEST RESULTS")
    print("="*60)
    print(f"\nStock: {symbol_name} ({yahoo_symbol})")
    print(f"Total download time: {total_time:.2f} seconds")
    print(f"\nPrice Data:")
    print(f"  Current Close: Rs.{todays_close:.2f}")
    print(f"  Previous Close: Rs.{prev_close:.2f}" if prev_close else "  Previous Close: N/A")
    print(f"\nReturns:")
    print(f"  1D: {change_1d:.2f}%" if change_1d is not None else "  1D: N/A")
    print(f"  5D: {change_5d:.2f}%" if change_5d is not None else "  5D: N/A")
    print(f"  1M: {change_1m:.2f}%" if change_1m is not None else "  1M: N/A")
    print("\n" + "="*60)

    if total_time < 120:
        print(f"[SUCCESS] Download completed within 2 minutes ({total_time:.2f}s)")
    else:
        print(f"[WARNING] Download took longer than 2 minutes ({total_time:.2f}s)")

    print("="*60)

    return True

if __name__ == "__main__":
    # Test with RELIANCE stock (one of the most liquid NSE stocks)
    print("\nTesting with RELIANCE.NS...")
    success = test_single_stock("RELIANCE")

    if success:
        print("\n[SUCCESS] Test completed successfully!")
        print("\nIf you want to test with a different stock, you can run:")
        print("  python test_single_download.py")
        print("\nAnd modify the symbol in the test_single_stock() call")
    else:
        print("\n[FAILED] Test failed!")
        print("\nPossible issues:")
        print("  1. Internet connection problem")
        print("  2. Yahoo Finance API is down or rate limiting")
        print("  3. Stock symbol is incorrect")
        print("  4. Market data not available for the requested dates")
