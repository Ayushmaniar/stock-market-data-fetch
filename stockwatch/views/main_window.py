"""Main application window for the StockWatch application."""

import sys
import os
import pandas as pd
from datetime import datetime

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, 
                             QLabel, QTextEdit, QProgressBar, QCompleter)
from PyQt5.QtCore import Qt, pyqtSlot, QUrl
from PyQt5.QtGui import QColor

from stockwatch.views.ui_components import ClickableLabel
from stockwatch.data.data_processor import DataDownloadThread

class StockWatchlistApp(QMainWindow):
    """Main application window for the StockWatch application."""
    
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
        self.current_refresh_type = None  # Initialize to prevent AttributeError
        self.df = None  # Initialize dataframe to None before UI setup

        self.init_ui()
        self.df = self.load_data()

        # Update UI components after data is loaded
        if self.df is not None:
            self.all_symbols = self.df['SYMBOL'].tolist() if 'SYMBOL' in self.df.columns else []
            self.stock_input.completer().model().setStringList(self.all_symbols)
            self.update_table()

    def ensure_data_folder_exists(self, data_folder):
        """Ensure the data folder exists, creating it if necessary."""
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        
    def init_ui(self):
        """Initialize the user interface."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()

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
        if self.df is not None:
            try:
                all_symbols_list = self.df['SYMBOL'].tolist()
            except (KeyError, AttributeError) as e:
                all_symbols_list = []
        else:
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

        # Only update table if data is loaded
        if self.df is not None:
            try:
                self.update_table()
            except (KeyError, AttributeError, ValueError) as e:
                # Silently handle table update errors during initialization
                pass
        
    def load_data(self):
        """Load stock data from the most recent file."""
        folder_path = "yahoo_finance_data"
        self.ensure_data_folder_exists(folder_path)
        files = [f for f in os.listdir(folder_path) if f.endswith("_stock_market_data.xlsx")]
        if not files:
            QMessageBox.warning(self, "No Data Found", "No stock data found. Please press the REFRESH ALL button to download data.")
            self.refresh_all_data()
            return None
        latest_file = max(files, key=lambda x: datetime.strptime(x[:10], "%Y-%m-%d"))
        return pd.read_excel(os.path.join(folder_path, latest_file))
        
    def update_table(self):
        """Update the table with current watchlist data."""
        self.table.setRowCount(0)
        if self.df is not None:
            watchlist_df = self.df[self.df['SYMBOL'].isin(self.watchlist)]
            
            if self.sort_column is not None:
                watchlist_df = watchlist_df.sort_values(by=self.sort_column, ascending=(self.sort_order == Qt.AscendingOrder))
            
            for _, row in watchlist_df.iterrows():
                self.add_row_to_table(row)
        
    def add_row_to_table(self, row):
        """Add a row to the table with stock data."""
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
        """Wrapper for add_stock that clears the input field."""
        self.add_stock()
        self.stock_input.clear()

    def add_stock(self):
        """Add a stock to the watchlist."""
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
        """Remove a stock from the watchlist."""
        self.watchlist.remove(symbol)
        self.update_table()
        
    def save_watchlist(self):
        """Save the current watchlist to a file."""
        with open("watchlist.txt", "w") as f:
            for symbol in self.watchlist:
                f.write(f"{symbol}\n")
        QMessageBox.information(self, "Watchlist Saved", "Your watchlist has been saved successfully.")
        
    def load_watchlist(self):
        """Load the watchlist from a file."""
        try:
            with open("watchlist.txt", "r") as f:
                self.watchlist = [line.strip() for line in f.readlines()]
        except FileNotFoundError:
            pass

    def on_header_clicked(self, logical_index):
        """Handle header clicks for sorting."""
        header_item = self.table.horizontalHeaderItem(logical_index)
        if header_item:
            self.sort_column = header_item.text()
            if self.sort_order == Qt.AscendingOrder:
                self.sort_order = Qt.DescendingOrder
            else:
                self.sort_order = Qt.AscendingOrder
            self.update_table()

    def refresh_all_data(self):
        """Refresh all stock data."""
        today = datetime.today().strftime('%Y-%m-%d')
        pandas_today = pd.Timestamp(today)

        # Check if today is a weekend
        if datetime.today().weekday() >= 5:  # 5 = Saturday, 6 = Sunday
            last_weekday = (pandas_today - pd.offsets.BDay(1)).strftime("%Y-%m-%d")
            retry = QMessageBox.question(
                self, "Weekend download", f"Today {today} is a Saturday/Sunday, would you like to download data for the last Friday {last_weekday} instead?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if retry == QMessageBox.Yes:
                date_to_use = last_weekday
            else:
                date_to_use = today
        else:
            date_to_use = today

        self.start_data_download(date_to_use=date_to_use)

    def quick_refresh_data(self):
        """Refresh only watchlist stock data."""
        if not self.watchlist:
            QMessageBox.warning(self, "No Stocks", "No stocks in the watchlist. Please add stocks before refreshing.")
            return
        
        today = datetime.today().strftime('%Y-%m-%d')
        pandas_today = pd.Timestamp(today)

        # Check if today is a weekend
        if datetime.today().weekday() >= 5:  # 5 = Saturday, 6 = Sunday
            last_weekday = (pandas_today - pd.offsets.BDay(1)).strftime("%Y-%m-%d")
            retry = QMessageBox.question(
                self, "Weekend download", f"Today {today} is a Saturday/Sunday, would you like to download data for the last Friday {last_weekday} instead?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if retry == QMessageBox.Yes:
                date_to_use = last_weekday
            else:
                date_to_use = today
        else:
            date_to_use = today

        self.start_data_download(symbols=self.watchlist, date_to_use=date_to_use)

    def start_data_download(self, symbols=None, date_to_use=None):
        """Start the data download thread."""
        self.download_thread = DataDownloadThread(symbols, date_to_use)
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
        """Update the progress bar."""
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
        """Update the status text."""
        self.status_text.append(status)

    @pyqtSlot(pd.DataFrame)
    def update_data(self, new_data):
        """Update the data after download is complete."""
        try:
            self.df = self.df.set_index('SYMBOL')
            new_data = new_data.set_index('SYMBOL')
            self.df.update(new_data.where(~new_data.isna()))
            self.df.reset_index(inplace=True)
        except:
            self.df = new_data

        self.update_table()
        self.progress_bar.setVisible(False)
        self.status_text.setVisible(False)
        self.refresh_all_button.setEnabled(True)
        self.quick_refresh_button.setEnabled(True)
        self.refresh_all_button.setStyleSheet(self.refresh_all_button.styleSheet().replace("transparent", "#27ae60"))
        self.quick_refresh_button.setStyleSheet(self.quick_refresh_button.styleSheet().replace("transparent", "#f39c12"))
        self.current_refresh_type = None  # Reset the refresh type

    @pyqtSlot(str)
    def show_error(self, error_message):
        """Show an error message."""
        QMessageBox.critical(self, "Error", f"An error occurred: {error_message}")
        self.progress_bar.setVisible(False)
        self.status_text.setVisible(False)
        self.refresh_all_button.setEnabled(True)
        self.quick_refresh_button.setEnabled(True)
        self.refresh_all_button.setStyleSheet(self.refresh_all_button.styleSheet().replace("transparent", "#27ae60"))
        self.quick_refresh_button.setStyleSheet(self.quick_refresh_button.styleSheet().replace("transparent", "#f39c12"))
        self.current_refresh_type = None  # Reset the refresh type

    def closeEvent(self, event):
        """Handle application close event - cleanup threads before exit."""
        # Check if download thread is running
        if hasattr(self, 'download_thread') and self.download_thread.isRunning():
            reply = QMessageBox.question(
                self,
                'Download in Progress',
                'Data download is still in progress. Do you want to stop it and exit?',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                # Terminate the thread and exit
                self.download_thread.terminate()
                self.download_thread.wait(2000)  # Wait up to 2 seconds for cleanup
                event.accept()
            else:
                # Don't close the application
                event.ignore()
        else:
            # No download running, safe to close
            event.accept()