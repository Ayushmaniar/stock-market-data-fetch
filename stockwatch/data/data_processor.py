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
import sys

# Set up logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler("stock_data_debug.log"),
                              logging.StreamHandler()])
logger = logging.getLogger("StockDataProcessor")

# Suppress yfinance error messages (delisted stocks, etc.)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

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

    def __init__(self, symbols=None, date_to_use=None, parent=None):
        super().__init__(parent)
        self.symbols = symbols
        self.date_to_use = date_to_use
        
    def download_with_retry(self, symbol, start_date, end_date, max_retries=3, retry_delay=2):
        """Download data with retry mechanism to handle transient errors"""
        for attempt in range(max_retries):
            try:
                data = yf.download(
                    symbol, 
                    start=start_date, 
                    end=end_date, 
                    progress=False
                )
                
                if isinstance(data.columns, pd.MultiIndex):
                    data.columns = [f"{col[0]}" for col in data.columns]

                # Ensure columns are in the expected order (yfinance may return them in any order)
                if not data.empty:
                    expected_cols = ['Open', 'High', 'Low', 'Close', 'Volume']
                    available_cols = [col for col in expected_cols if col in data.columns]
                    if available_cols:
                        data = data[available_cols]

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
            
            # Use tqdm for progress tracking in console (only if not running as executable)
            self.status_signal.emit(f"Downloading data for {total_symbols} symbols...")
            # Disable tqdm in packaged executable to avoid stdout errors
            if getattr(sys, 'frozen', False):
                # Running as compiled executable - no tqdm
                iterator = enumerate(yahoo_finance_symbols)
            else:
                # Running in development - use tqdm
                iterator = enumerate(tqdm(yahoo_finance_symbols, desc="Downloading stock data", unit="symbol"))

            for company_no, company in iterator:
                self.status_signal.emit(f"Downloading data for {company}")
                time.sleep(0.01)
                try:
                    fetch_data = self.download_with_retry(
                        company, 
                        start_date=one_month_ago, 
                        end_date=pd.to_datetime(self.date_to_use) + pd.Timedelta(days=1)
                    )
                    
                    try:
                        if not fetch_data.empty and 'Date' in fetch_data.reset_index().columns:
                            if fetch_data.reset_index()['Date'].max() > max_date:
                                max_date = fetch_data.reset_index()['Date'].max()
                    except Exception as e:
                        logger.error(f"Error getting max date for {company}: {str(e)}")

                    if not fetch_data.empty:
                        stock_data[company] = fetch_data
                    else:
                        error_companies.append(company)
                        error_msg = f"Warning: No data retrieved for {company}"
                        self.status_signal.emit(error_msg)
                        
                except json.JSONDecodeError as e:
                    error_companies.append(company)
                    error_msg = f"Warning: JSON decode error for {company}: {str(e)}"
                    logger.error(error_msg)
                    logger.error(traceback.format_exc())
                    self.status_signal.emit(error_msg)
                    
                except Exception as e:
                    error_companies.append(company)
                    error_msg = f"Warning: Error downloading {company}: {str(e)}"
                    logger.error(error_msg)
                    logger.error(traceback.format_exc())
                    self.status_signal.emit(error_msg)

                # Update progress for GUI
                progress = int((company_no + 1) / total_symbols * 100)
                self.progress_signal.emit(progress)

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

                        # Get previous close (1D ago)
                        prev_close_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == last_trading_day]
                        if not prev_close_data.empty:
                            prev_close = float(prev_close_data.iloc[0][close_col])
                        else:
                            prev_close = None

                        # Get 5 days ago close
                        five_days_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == five_trading_days_ago]
                        if not five_days_data.empty:
                            five_days_close = float(five_days_data.iloc[0][close_col])
                        else:
                            five_days_close = None

                        # For 1M, find the closest available date near one_month_ago
                        one_month_data = data.loc[data['Date'].dt.strftime('%Y-%m-%d') == one_month_ago]
                        if one_month_data.empty:
                            # If exact date not available, find the closest earlier date
                            available_dates_df = data[data['Date'].dt.strftime('%Y-%m-%d') <= one_month_ago]
                            if not available_dates_df.empty:
                                closest_date = available_dates_df['Date'].max()
                                one_month_data = data.loc[data['Date'] == closest_date]

                        if not one_month_data.empty:
                            one_month_close = float(one_month_data.iloc[0][close_col])
                        else:
                            one_month_close = None

                        single_row['Previous_Close'] = prev_close
                        single_row['1D'] = ((todays_close - prev_close) / prev_close * 100) if prev_close is not None else None
                        single_row['5D'] = ((todays_close - five_days_close) / five_days_close * 100) if five_days_close is not None else None
                        single_row['1M'] = ((todays_close - one_month_close) / one_month_close * 100) if one_month_close is not None else None

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