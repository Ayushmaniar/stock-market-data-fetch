import sys
import os
import pandas as pd
import yfinance as yf
from datetime import datetime
from tqdm import tqdm
from urllib.parse import quote_plus
import time
from pandas.tseries.offsets import BDay
import warnings
from PyQt5.QtCore import QThread, pyqtSignal

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QLabel, 
                             QFrame, QCompleter, QTextEdit, QProgressBar)
from PyQt5.QtCore import Qt, QSortFilterProxyModel, pyqtSlot
from PyQt5.QtGui import QColor, QPalette, QFont, QDesktopServices, QCursor
from PyQt5.QtCore import QUrl
from datetime import datetime

warnings.filterwarnings('ignore')

class ClickableLabel(QLabel):
    def __init__(self, text, url, parent=None):
        super().__init__(text, parent)
        self.url = url
        self.setStyleSheet("color: blue; text-decoration: underline;")
        self.setCursor(QCursor(Qt.PointingHandCursor))

        # Set a larger font size
        font = self.font()
        font.setPointSize(9)  # You can adjust this value as needed
        # font.setBold(True)     # Optional: make the font bold
        self.setFont(font)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            QDesktopServices.openUrl(QUrl(self.url))

class DataDownloadThread(QThread):
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(pd.DataFrame)
    error_signal = pyqtSignal(str)

    def __init__(self, symbols=None, parent=None):
        super().__init__(parent)
        self.symbols = symbols

    def run(self):
        try:
            # Read list of stocks from the CSV file
            stocks_df = pd.read_csv("../../EQUITY_L.csv")
            stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'

            stocks_df = stocks_df

            # If symbols are provided (for quick refresh), use only those
            if self.symbols:
                yahoo_finance_symbols = [symbol for symbol in self.symbols]
            else:
                yahoo_finance_symbols = list(stocks_df['YahooEquiv'])

            error_companies = []

            # Download data for the current day using yfinance
            today = datetime.today().strftime('%Y-%m-%d')
            pandas_today = pd.Timestamp(today)
            last_trading_day = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
            five_trading_days_ago = (pandas_today - BDay(5)).strftime('%Y-%m-%d')

            one_month_ago = pandas_today - pd.DateOffset(months=1)
            if one_month_ago.weekday() > 4:
                one_month_ago = one_month_ago + pd.DateOffset(days=(7 - one_month_ago.weekday()))
            one_month_ago = one_month_ago.strftime('%Y-%m-%d')

            stock_data = {}
            max_date = pd.Timestamp('2008-01-01')

            total_symbols = len(yahoo_finance_symbols)
            for company_no, company in enumerate(yahoo_finance_symbols):
                self.status_signal.emit(f"Downloading data for {company}")
                time.sleep(0.01)
                fetch_data = yf.download(company, start=one_month_ago, end=pd.to_datetime(today) + pd.Timedelta(days=1), progress=False)
                try:
                    if fetch_data.reset_index()['Date'].max() > max_date:
                        max_date = fetch_data.reset_index()['Date'].max()
                except:
                    pass

                if not fetch_data.empty:
                    stock_data[company] = fetch_data

                progress = int((company_no + 1) / total_symbols * 100)
                self.progress_signal.emit(progress)

            self.status_signal.emit(f"Download completed. Processing data...")

            if max_date.strftime('%Y-%m-%d') != today:
                self.error_signal.emit("No stock data was successfully downloaded. Today might be a holiday.")
                return

            all_stock_data = pd.DataFrame()

            for symbol, data in stock_data.items():
                if not data.empty:
                    data = data.reset_index()
                    data['SYMBOL'] = symbol
                    single_row = data.loc[data['Date'] == today]
                    if not single_row.empty:
                        todays_close = single_row['Close'].values[0]
                        prev_close = data.loc[data['Date'] == last_trading_day, 'Close'].values[0] if not data.loc[data['Date'] == last_trading_day].empty else None
                        five_days_close = data.loc[data['Date'] == five_trading_days_ago, 'Close'].values[0] if not data.loc[data['Date'] == five_trading_days_ago].empty else None
                        one_month_close = data.loc[data['Date'] == one_month_ago, 'Close'].values[0] if not data.loc[data['Date'] == one_month_ago].empty else None
                        
                        single_row['Previous_Close'] = prev_close
                        single_row['1D'] = ((todays_close - prev_close) / prev_close * 100) if prev_close else None
                        single_row['5D'] = ((todays_close - five_days_close) / five_days_close * 100) if five_days_close else None
                        single_row['1M'] = ((todays_close - one_month_close) / one_month_close * 100) if one_month_close else None

                        all_stock_data = pd.concat([all_stock_data, single_row])

            all_stock_data.reset_index(inplace=True, drop=True)
            all_stock_data['Date'] = all_stock_data['Date'].astype(str)
            all_stock_data.sort_values(by='1D', ascending=False, inplace=True)
            all_stock_data = all_stock_data.round(2)

            # Move 'SYMBOL' column to the first position
            cols = list(all_stock_data.columns)
            cols.insert(0, cols.pop(cols.index('SYMBOL')))
            all_stock_data = all_stock_data[cols]

            all_stock_data.reset_index(inplace=True, drop=True)

            # Create a new column containing the Google search URL for each symbol
            google_search_urls = 'https://www.google.com/search?q=' + all_stock_data['SYMBOL'].str.replace('.NS','').apply(quote_plus) + '+share+price'
            google_search_urls = google_search_urls.sort_index()

            self.status_signal.emit("Data processing completed.")

            # If it's a full refresh, save the data to a new Excel file
            if not self.symbols:
                output_file_name = f'../../yahoo_finance_data/{today}_stock_market_data.xlsx'
                with pd.ExcelWriter(output_file_name, engine='xlsxwriter') as writer:
                    # Write the DataFrame to the Excel file
                    all_stock_data.to_excel(writer, index=False, sheet_name='Sheet1')

                    # Get the workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']

                    for i, url in enumerate(google_search_urls):
                        cell = f'A{i+2}'  # Adjust the row index to match the data (indexing starts from 2 because of header)
                        worksheet.write_url(cell, url, string=all_stock_data.loc[i, 'SYMBOL'])

                        red_format = workbook.add_format({'bg_color': '#FFC7CE'})

                            # Apply the format to the 'Close' column
                        worksheet.conditional_format('F2:F' + str(len(all_stock_data) + 1), {'type': 'no_blanks', 'format': red_format})

                self.status_signal.emit(f"Data saved to {output_file_name}")
    

            self.finished_signal.emit(all_stock_data)

        except Exception as e:
            self.error_signal.emit(str(e))


class StockWatchlistApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Watchlist")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                selection-background-color: #e0e0e0;
            }
            QHeaderView::section {
                background-color: #4a90e2;
                color: white;
                padding: 5px;
                border: 1px solid #3a80d2;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 15px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #3a80d2;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 15px;
                font-size: 14px;
            }
        """)
        
        self.watchlist = []
        self.load_watchlist()
        
        self.sort_column = None
        self.sort_order = Qt.AscendingOrder

        self.last_full_refresh = None
        self.last_quick_refresh = None
        
        self.init_ui()
        self.df = self.load_data()
        self.init_ui()
        self.update_refresh_times()

        # self.installEventFilter(self)
        

    def ensure_data_folder_exists(self, data_folder):
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()

        # Add refresh 
        
        # Title
        title_label = QLabel("Stock Watchlist")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #333; margin: 10px 0;")
        layout.addWidget(title_label)
        
        # Add stock input
        input_layout = QHBoxLayout()
        input_layout.addStretch()
        self.stock_input = QLineEdit()
        self.stock_input.setFixedWidth(300)
        self.stock_input.setPlaceholderText("Enter stock symbol")
        self.stock_input.returnPressed.connect(self.add_stock_wrapper)
        
        # Load all stock symbols for auto-suggestion
        try:
            all_symbols_list = self.df['SYMBOL'].tolist()

        except:
            all_symbols_list = []
        self.all_symbols = all_symbols_list

        completer = QCompleter(self.all_symbols)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.stock_input.setCompleter(completer)
        
        add_button = QPushButton("Add Stock")
        add_button.clicked.connect(self.add_stock_wrapper)
        input_layout.addWidget(self.stock_input)
        input_layout.addWidget(add_button)
        input_layout.addStretch()
        layout.addLayout(input_layout)
        
        # Table for displaying stocks
        self.table = QTableWidget()
        self.table.setColumnCount(13)
        headers = ["SYMBOL", "Date", "Open", "High", "Low", "Close", "Adj Close", "Volume", "Previous_Close", "1D", "5D", "1M", "Action"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)
        
        # Refresh buttons
        refresh_layout = QHBoxLayout()
        self.refresh_all_button = QPushButton("DOWNLOAD ALL")
        self.refresh_all_button.clicked.connect(self.refresh_all_data)
        self.refresh_all_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                border-radius: 20px;
                font-weight: bold;
                font-size: 16px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)
        self.quick_refresh_button = QPushButton("QUICK REFRESH")
        self.quick_refresh_button.clicked.connect(self.quick_refresh_data)
        self.quick_refresh_button.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                border-radius: 20px;
                font-weight: bold;
                font-size: 16px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #f1c40f;
            }
        """)
        refresh_layout.addWidget(self.refresh_all_button)
        refresh_layout.addWidget(self.quick_refresh_button)
        layout.addLayout(refresh_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Status text bar
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFixedHeight(100)
        self.status_text.setVisible(False)
        layout.addWidget(self.status_text)
        
        # Save button
        save_button = QPushButton("Save Watchlist")
        save_button.clicked.connect(self.save_watchlist)
        save_button.setStyleSheet("""
            background-color: #3498db;
            font-weight: bold;
            font-size: 16px;
            padding: 10px 20px;
        """)
        save_button.setFixedWidth(200)
        save_layout = QHBoxLayout()
        save_layout.addStretch()
        save_layout.addWidget(save_button)
        save_layout.addStretch()
        layout.addLayout(save_layout)
        
        central_widget.setLayout(layout)
        
        try:
            self.update_table()
        except:
            pass

    def update_last_refresh_label(self):
        folder_path = '../../yahoo_finance_data'
        files = [f for f in os.listdir(folder_path) if f.endswith("_stock_market_data.xlsx")]
        if files:
            latest_file = max(files, key=lambda x: datetime.strptime(x[:10], "%Y-%m-%d"))
            self.last_refresh_label.setText(f"Last full refresh: {latest_file[:10]}")
        else:
            self.last_refresh_label.setText("Last full refresh: Never")
        
    def load_data(self):
        folder_path = "../../yahoo_finance_data"
        self.ensure_data_folder_exists(folder_path)
        files = [f for f in os.listdir(folder_path) if f.endswith("_stock_market_data.xlsx")]
        if not files:
            QMessageBox.warning(self, "No Data Found", "No stock data found. Please press the REFRESH ALL button to download data.")
            self.refresh_all_data()
            return None
        files = [f for f in os.listdir(folder_path) if f.endswith("_stock_market_data.xlsx")]
        latest_file = max(files, key=lambda x: datetime.strptime(x[:10], "%Y-%m-%d"))
        return pd.read_excel(os.path.join(folder_path, latest_file))
        
    def update_table(self):
        self.table.setRowCount(0)
        if self.df is not None:
            watchlist_df = self.df[self.df['SYMBOL'].isin(self.watchlist)]
            
            if self.sort_column is not None:
                watchlist_df = watchlist_df.sort_values(by=self.sort_column, ascending=(self.sort_order == Qt.AscendingOrder))
            
            for _, row in watchlist_df.iterrows():
                self.add_row_to_table(row)
        
    def add_row_to_table(self, row):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        # Create clickable label for stock symbol
        symbol = str(row['SYMBOL'])
        url = f"https://www.google.com/search?q={symbol.replace('.NS', '')}+share+price"
        clickable_label = ClickableLabel(symbol, url)
        self.table.setCellWidget(row_position, 0, clickable_label)


        for i, value in enumerate(row):
            item = QTableWidgetItem(str(value))
            if i in [9, 10, 11]:  # 1D, 5D, 1M columns
                value_float = float(value) if value and value != 'nan' else 0
                if value_float > 0:
                    item.setForeground(QColor('green'))
                elif value_float < 0:
                    item.setForeground(QColor('red'))
            if i == 0:
                pass
            else:
                self.table.setItem(row_position, i, item)
        
        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(lambda: self.delete_stock(row['SYMBOL']))
        delete_button.setStyleSheet("""
            background-color: #ff9999;
            color: black;
            border: none;
            padding: 3px 8px;
            border-radius: 10px;
            font-size: 12px;
        """)
        self.table.setCellWidget(row_position, 12, delete_button)
        
    def add_stock_wrapper(self):
        self.add_stock()
        self.stock_input.clear()

    def add_stock(self):
        symbol = self.stock_input.text().upper()
        if symbol and symbol not in self.watchlist:
            if self.df is not None and symbol in self.df['SYMBOL'].values:
                self.watchlist.append(symbol)
                self.quick_refresh_data()
                self.update_table()
            
            else:
                QMessageBox.warning(self, "Stock Not Found", "Stock unavailable in the data. Please try a different stock or correct its spelling.")
        elif symbol in self.watchlist:
            QMessageBox.information(self, "Duplicate Stock", "This stock is already in your watchlist.")
        
    def delete_stock(self, symbol):
        self.watchlist.remove(symbol)
        self.update_table()
        
    def save_watchlist(self):
        with open("../../watchlist.txt", "w") as f:
            for symbol in self.watchlist:
                f.write(f"{symbol}\n")
        QMessageBox.information(self, "Watchlist Saved", "Your watchlist has been saved successfully.")
        
    def load_watchlist(self):
        try:
            with open("../../watchlist.txt", "r") as f:
                self.watchlist = [line.strip() for line in f.readlines()]
        except FileNotFoundError:
            pass

    def on_header_clicked(self, logical_index):
        header_item = self.table.horizontalHeaderItem(logical_index)
        if header_item:
            self.sort_column = header_item.text()
            if self.sort_order == Qt.AscendingOrder:
                self.sort_order = Qt.DescendingOrder
            else:
                self.sort_order = Qt.AscendingOrder
            self.update_table()

    def refresh_all_data(self):
        self.start_data_download()
        # Update last refresh time after downloading
        self.update_last_refresh_label()

    def quick_refresh_data(self):
        if not self.watchlist:
            QMessageBox.warning(self, "No Stocks", "No stocks in the watchlist. Please add stocks before refreshing.")
            return
        self.start_data_download(self.watchlist)

    def start_data_download(self, symbols=None):
        self.download_thread = DataDownloadThread(symbols)
        self.download_thread.progress_signal.connect(self.update_progress)
        self.download_thread.status_signal.connect(self.update_status)
        self.download_thread.finished_signal.connect(self.update_data)
        self.download_thread.error_signal.connect(self.show_error)
        

        self.progress_bar.setVisible(True)
        self.status_text.setVisible(True)
        self.refresh_all_button.setEnabled(False)
        self.quick_refresh_button.setEnabled(False)
        
        # Set the current refresh type
        self.current_refresh_type = 'quick' if symbols else 'full'
        
        self.download_thread.start()

    @pyqtSlot(int)
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if self.current_refresh_type == 'quick':
            self.quick_refresh_button.setStyleSheet(f"""
                QPushButton {{
                    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #f39c12, stop:{value/100} #f39c12, stop:{value/100} transparent, stop:1 transparent);
                    border-radius: 20px;
                    font-weight: bold;
                    font-size: 16px;
                    padding: 10px 20px;
                }}
            """)
        else:
            self.refresh_all_button.setStyleSheet(f"""
                QPushButton {{
                    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #27ae60, stop:{value/100} #27ae60, stop:{value/100} transparent, stop:1 transparent);
                    border-radius: 20px;
                    font-weight: bold;
                    font-size: 16px;
                    padding: 10px 20px;
                }}
            """)

    @pyqtSlot(str)
    def update_status(self, status):
        self.status_text.append(status)

    @pyqtSlot(pd.DataFrame)
    def update_data(self, new_data):
        try:
            self.df = self.df.set_index('SYMBOL')
            new_data = new_data.set_index('SYMBOL')
            self.df.update(new_data.where(~new_data.isna()))
            self.df.reset_index(inplace=True)
        except:
            self.df = new_data

        # self.df = new_data
        self.update_table()
        self.progress_bar.setVisible(False)
        self.status_text.setVisible(False)
        self.refresh_all_button.setEnabled(True)
        self.quick_refresh_button.setEnabled(True)
        self.refresh_all_button.setStyleSheet(self.refresh_all_button.styleSheet().replace("transparent", "#27ae60"))
        self.quick_refresh_button.setStyleSheet(self.quick_refresh_button.styleSheet().replace("transparent", "#f39c12"))
        self.current_refresh_type = None  # Reset the refresh type
        # QMessageBox.information(self, "Data Updated", "Stock data has been successfully updated.")

    @pyqtSlot(str)
    def show_error(self, error_message):
        QMessageBox.critical(self, "Error", f"An error occurred: {error_message}")
        self.progress_bar.setVisible(False)
        self.status_text.setVisible(False)
        self.refresh_all_button.setEnabled(True)
        self.quick_refresh_button.setEnabled(True)
        self.refresh_all_button.setStyleSheet(self.refresh_all_button.styleSheet().replace("transparent", "#27ae60"))
        self.quick_refresh_button.setStyleSheet(self.quick_refresh_button.styleSheet().replace("transparent", "#f39c12"))
        self.current_refresh_type = None  # Reset the refresh type

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StockWatchlistApp()
    window.show()
    sys.exit(app.exec_())