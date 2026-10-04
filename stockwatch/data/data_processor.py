"""Stock data download and processing functionality."""

import pandas as pd
import yfinance as yf
import time
from datetime import datetime
from urllib.parse import quote_plus
from pandas.tseries.offsets import BDay
from PyQt5.QtCore import QThread, pyqtSignal
import os
import warnings
import json
import logging
import traceback
from requests.exceptions import RequestException
# requests is already included via yfinance, no need to import separately
from tqdm import tqdm  # Import tqdm for progress bars
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler("stock_data_debug.log"), 
                              logging.StreamHandler()])
logger = logging.getLogger("StockDataProcessor")

warnings.filterwarnings('ignore')

def get_application_path():
    """Get the base path of the application, works both in development and when packaged with PyInstaller"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running as compiled executable (PyInstaller)
        return sys._MEIPASS
    else:
        # Running in normal Python environment
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class DataDownloadThread(QThread):
    """Thread for downloading stock data asynchronously."""
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(pd.DataFrame)
    error_signal = pyqtSignal(str)

    # Number of parallel download threads (5 is conservative to avoid rate limits)
    MAX_DOWNLOAD_THREADS = 5

    def __init__(self, symbols=None, date_to_use=None, parent=None):
        super().__init__(parent)
        self.symbols = symbols
        self.date_to_use = date_to_use
        
    def download_with_retry(self, symbol, start_date, end_date, max_retries=3, retry_delay=2):
        """Download data with retry mechanism to handle transient errors.

        Uses yf.Ticker().history() instead of yf.download() for thread-safety.
        yf.download() uses shared global state that causes race conditions
        when called from multiple threads.
        """
        for attempt in range(max_retries):
            try:
                # Use Ticker.history() instead of yf.download() for thread-safety
                # Each Ticker object is independent, avoiding shared state issues
                ticker = yf.Ticker(symbol)
                data = ticker.history(start=start_date, end=end_date)

                # history() returns columns: Open, High, Low, Close, Volume, Dividends, Stock Splits
                # Keep only the columns we need
                if not data.empty:
                    expected_cols = ['Open', 'High', 'Low', 'Close', 'Volume']
                    available_cols = [col for col in expected_cols if col in data.columns]
                    if available_cols:
                        data = data[available_cols]

                    # history() returns a tz-aware (Asia/Kolkata) index; drop the tz so
                    # dates stay plain YYYY-MM-DD downstream (Excel/UI Date column)
                    if data.index.tz is not None:
                        data.index = data.index.tz_localize(None)

                if data.empty:
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
                        continue

                return data
                
            except json.JSONDecodeError as e:
                error_msg = f"JSONDecodeError for {symbol}: {str(e)}"
                logger.error(error_msg)
                logger.error(f"Traceback: {traceback.format_exc()}")
                
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    logger.error(f"Failed to download {symbol} after {max_retries} attempts")
                    raise
                    
            except RequestException as e:
                error_msg = f"Network error for {symbol}: {str(e)}"
                logger.error(error_msg)
                
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    logger.error(f"Failed to download {symbol} after {max_retries} attempts")
                    raise
                    
            except Exception as e:
                error_msg = f"Unknown error for {symbol}: {str(e)}"
                logger.error(error_msg)
                logger.error(f"Traceback: {traceback.format_exc()}")
                
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    logger.error(f"Failed to download {symbol} after {max_retries} attempts")
                    raise
                
        return pd.DataFrame()  # Return empty dataframe if all retries failed

    def _download_single_stock(self, args):
        """Download a single stock - helper for thread pool.

        Args:
            args: Tuple of (symbol, start_date, end_date)

        Returns:
            Tuple of (symbol, data, error_message)
        """
        symbol, start_date, end_date = args
        try:
            data = self.download_with_retry(symbol, start_date, end_date)
            if not data.empty:
                return (symbol, data, None)
            return (symbol, None, "Empty data")
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Thread download error for {symbol}: {error_msg}")
            return (symbol, None, error_msg)

    def run(self):
        try:
            # Get the correct path to the EQUITY_L.csv file
            app_base_path = get_application_path()
            csv_path = os.path.join(app_base_path, "EQUITY_L.csv")
            logger.info(f"Loading stock data from: {csv_path}")
            
            if not os.path.exists(csv_path):
                error_msg = f"EQUITY_L.csv not found at {csv_path}"
                logger.error(error_msg)
                self.error_signal.emit(error_msg)
                return
                
            stocks_df = pd.read_csv(csv_path)
            stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'

            if self.symbols:
                yahoo_finance_symbols = [symbol for symbol in self.symbols]
            else:
                yahoo_finance_symbols = list(stocks_df['YahooEquiv'])

            error_companies = []

            pandas_today = pd.Timestamp(self.date_to_use)
            
            last_trading_day = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
            five_trading_days_ago = (pandas_today - BDay(5)).strftime('%Y-%m-%d')

            one_month_ago = pandas_today - pd.DateOffset(months=1)
            if one_month_ago.weekday() > 4:
                one_month_ago = one_month_ago + pd.DateOffset(days=(7 - one_month_ago.weekday()))
            one_month_ago = one_month_ago.strftime('%Y-%m-%d')

            stock_data = {}
            max_date = pd.Timestamp('2008-01-01')

            total_symbols = len(yahoo_finance_symbols)

            # Parallel download using ThreadPoolExecutor
            self.status_signal.emit(f"Downloading data for {total_symbols} symbols using {self.MAX_DOWNLOAD_THREADS} threads...")
            logger.info(f"Starting parallel download with {self.MAX_DOWNLOAD_THREADS} threads for {total_symbols} symbols")

            # Prepare arguments for thread pool
            end_date_str = (pd.to_datetime(self.date_to_use) + pd.Timedelta(days=1)).strftime('%Y-%m-%d')
            download_args = [(symbol, one_month_ago, end_date_str) for symbol in yahoo_finance_symbols]

            # Track progress
            completed_count = 0

            # Use ThreadPoolExecutor for parallel downloads
            with ThreadPoolExecutor(max_workers=self.MAX_DOWNLOAD_THREADS) as executor:
                # Submit all download tasks
                future_to_symbol = {
                    executor.submit(self._download_single_stock, args): args[0]
                    for args in download_args
                }

                # Create progress bar for console (only if not running as executable)
                if getattr(sys, 'frozen', False):
                    pbar = None
                else:
                    pbar = tqdm(total=total_symbols, desc="Downloading stock data", unit="symbol")

                # Process results as they complete
                for future in as_completed(future_to_symbol):
                    symbol = future_to_symbol[future]
                    completed_count += 1

                    try:
                        symbol, fetch_data, error = future.result()

                        if fetch_data is not None:
                            # Update max_date if needed (for tracking purposes only)
                            try:
                                if 'Date' in fetch_data.reset_index().columns:
                                    data_max_date = fetch_data.reset_index()['Date'].max()
                                    if data_max_date > max_date:
                                        max_date = data_max_date
                            except Exception as e:
                                # Non-critical error, just log at debug level
                                logger.debug(f"Error getting max date for {symbol}: {str(e)}")

                            stock_data[symbol] = fetch_data
                        else:
                            error_companies.append(symbol)
                            if error:
                                logger.warning(f"No data for {symbol}: {error}")

                    except Exception as e:
                        error_companies.append(symbol)
                        error_msg = f"Warning: Error downloading {symbol}: {str(e)}"
                        logger.error(error_msg)
                        logger.error(traceback.format_exc())

                    # Update progress for GUI
                    progress = int(completed_count / total_symbols * 100)
                    self.progress_signal.emit(progress)

                    # Update console progress bar
                    if pbar:
                        pbar.update(1)

                    # Emit status every 50 stocks to avoid overwhelming the UI
                    if completed_count % 50 == 0:
                        self.status_signal.emit(f"Downloaded {completed_count}/{total_symbols} symbols...")

                # Close progress bar
                if pbar:
                    pbar.close()

            self.status_signal.emit(f"Download completed. Processing data...")
            
            if error_companies:
                self.status_signal.emit(f"Failed to download data for {len(error_companies)} symbols.")
                logger.warning(f"Failed to download data for {len(error_companies)} symbols: {error_companies[:10]}{'...' if len(error_companies) > 10 else ''}")
            
            if not stock_data:
                error_msg = "No stock data was successfully downloaded. Check your internet connection or try again later."
                logger.error(error_msg)
                self.error_signal.emit(error_msg)
                return
                
            dates_available = set()
            for symbol, data in stock_data.items():
                if not data.empty:
                    dates_available.update(data.reset_index()['Date'].dt.strftime('%Y-%m-%d').tolist())
            
            if self.date_to_use not in dates_available:
                warning_msg = f"Warning: No data available for {self.date_to_use}. Using most recent data available."
                logger.warning(warning_msg)
                self.status_signal.emit(warning_msg)
                
            all_stock_data = pd.DataFrame()

            # Use tqdm for data processing loop as well (only if not running as executable)
            self.status_signal.emit(f"Processing data for {len(stock_data)} symbols...")
            # Disable tqdm in packaged executable to avoid stdout errors
            if getattr(sys, 'frozen', False):
                # Running as compiled executable - no tqdm
                processing_iterator = stock_data.items()
            else:
                # Running in development - use tqdm
                processing_iterator = tqdm(stock_data.items(), desc="Processing stock data", unit="symbol")

            for symbol, data in processing_iterator:
                if not data.empty:
                    data = data.reset_index()
                    data['SYMBOL'] = symbol
                    
                    if self.date_to_use in data['Date'].dt.strftime('%Y-%m-%d').tolist():
                        target_date = self.date_to_use
                    else:
                        available_dates = data['Date'].dt.strftime('%Y-%m-%d').tolist()
                        available_dates.sort(reverse=True)
                        if available_dates:
                            target_date = available_dates[0]
                            self.status_signal.emit(f"Using {target_date} data for {symbol} (requested date not available)")
                        else:
                            continue
                    
                    single_row = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == target_date]
                    if not single_row.empty:
                        close_col = 'Close'
                        # Safely access array values with bounds checking
                        if close_col in single_row.columns and len(single_row[close_col].values) > 0:
                            todays_close = single_row[close_col].values[0]
                        else:
                            continue  # Skip this symbol if no close price available

                        prev_close_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == last_trading_day]
                        prev_close = prev_close_data[close_col].values[0] if not prev_close_data.empty and len(prev_close_data[close_col].values) > 0 else None

                        five_days_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == five_trading_days_ago]
                        five_days_close = five_days_data[close_col].values[0] if not five_days_data.empty and len(five_days_data[close_col].values) > 0 else None

                        one_month_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == one_month_ago]
                        one_month_close = one_month_data[close_col].values[0] if not one_month_data.empty and len(one_month_data[close_col].values) > 0 else None
                        
                        single_row['Previous_Close'] = prev_close
                        single_row['1D'] = ((todays_close - prev_close) / prev_close * 100) if prev_close else None
                        single_row['5D'] = ((todays_close - five_days_close) / five_days_close * 100) if five_days_close else None
                        single_row['1M'] = ((todays_close - one_month_close) / one_month_close * 100) if one_month_close else None

                        all_stock_data = pd.concat([all_stock_data, single_row])

            if all_stock_data.empty:
                error_msg = "No valid stock data was found for the requested date."
                logger.error(error_msg)
                self.error_signal.emit(error_msg)
                return
                
            all_stock_data.reset_index(inplace=True, drop=True)
            all_stock_data['Date'] = all_stock_data['Date'].astype(str)
            
            try:
                all_stock_data.sort_values(by='1D', ascending=False, inplace=True)
                all_stock_data = all_stock_data.round(2)
            except KeyError as e:
                logger.error(f"Error sorting dataframe: {str(e)}")
                logger.error(f"Available columns: {all_stock_data.columns.tolist()}")

            try:
                cols = list(all_stock_data.columns)
                if 'SYMBOL' in cols:
                    cols.insert(0, cols.pop(cols.index('SYMBOL')))
                    all_stock_data = all_stock_data[cols]
                else:
                    logger.warning("'SYMBOL' column not found in DataFrame. Columns available: %s", cols)
            except Exception as e:
                logger.error(f"Error reordering columns: {str(e)}")

            all_stock_data.reset_index(inplace=True, drop=True)
            
            try:
                google_search_urls = 'https://www.google.com/search?q=' + all_stock_data['SYMBOL'].str.replace('.NS','').apply(quote_plus) + '+share+price'
                google_search_urls = google_search_urls.sort_index()
            except Exception as e:
                logger.error(f"Error creating Google search URLs: {str(e)}")
                google_search_urls = None

            self.status_signal.emit("Data processing completed.")

            if not self.symbols:
                os.makedirs("yahoo_finance_data", exist_ok=True)
                output_file_name = f'yahoo_finance_data/{self.date_to_use}_stock_market_data.xlsx'
                
                all_stock_data.columns = [str(col) if not isinstance(col, tuple) else col[0] for col in all_stock_data.columns]
                
                with pd.ExcelWriter(output_file_name, engine='xlsxwriter') as writer:
                    try:
                        all_stock_data.to_excel(writer, index=False, sheet_name='Sheet1')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Sheet1']

                        if google_search_urls is not None:
                            for i, url in enumerate(google_search_urls):
                                cell = f'A{i+2}'
                                worksheet.write_url(cell, url, string=all_stock_data.loc[i, 'SYMBOL'])

                        red_format = workbook.add_format({'bg_color': '#FFC7CE'})

                        if 'Close' in all_stock_data.columns:
                            close_col_idx = all_stock_data.columns.get_loc('Close')
                            close_col_letter = chr(ord('A') + close_col_idx)
                            worksheet.conditional_format(f'{close_col_letter}2:{close_col_letter}{len(all_stock_data) + 1}',
                                                        {'type': 'no_blanks', 'format': red_format})
                    except Exception as e:
                        logger.error(f"Error writing to Excel: {str(e)}")
                        logger.error(traceback.format_exc())
                        self.error_signal.emit(f"Error saving data: {str(e)}")
                        return

                self.status_signal.emit(f"Data saved to {output_file_name}")

            self.finished_signal.emit(all_stock_data)

        except Exception as e:
            error_msg = str(e)
            logger.error(f"Critical error in download thread: {error_msg}")
            logger.error(traceback.format_exc())
            self.error_signal.emit(error_msg)