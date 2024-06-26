import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QLabel, QFrame
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPalette, QFont
from datetime import datetime

class StockWatchlistApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stock Watchlist")
        self.setGeometry(100, 100, 1000, 600)
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
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #3a80d2;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
        """)
        
        self.watchlist = []
        self.load_watchlist()
        
        self.sort_column = None
        self.sort_order = Qt.AscendingOrder
        
        self.init_ui()
        
    def init_ui(self):
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
        self.stock_input = QLineEdit()
        self.stock_input.setPlaceholderText("Enter stock symbol")
        self.stock_input.returnPressed.connect(self.add_stock)
        add_button = QPushButton("Add Stock")
        add_button.clicked.connect(self.add_stock)
        input_layout.addWidget(self.stock_input)
        input_layout.addWidget(add_button)
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
        
        # Save button
        save_button = QPushButton("Save Watchlist")
        save_button.clicked.connect(self.save_watchlist)
        save_button.setStyleSheet("""
            background-color: #27ae60;
            font-weight: bold;
        """)
        layout.addWidget(save_button)
        
        central_widget.setLayout(layout)
        
        self.load_data()
        
    def load_data(self):
        folder_path = "../yahoo_finance_data"
        files = [f for f in os.listdir(folder_path) if f.endswith("_stock_market_data.xlsx")]
        latest_file = max(files, key=lambda x: datetime.strptime(x[:10], "%Y-%m-%d"))
        
        self.df = pd.read_excel(os.path.join(folder_path, latest_file))
        self.update_table()
        
    def update_table(self):
        self.table.setRowCount(0)
        watchlist_df = self.df[self.df['SYMBOL'].isin(self.watchlist)]
        
        if self.sort_column is not None:
            watchlist_df = watchlist_df.sort_values(by=self.sort_column, ascending=(self.sort_order == Qt.AscendingOrder))
        
        for _, row in watchlist_df.iterrows():
            self.add_row_to_table(row)
        
    def add_row_to_table(self, row):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for i, value in enumerate(row):
            item = QTableWidgetItem(str(value))
            if i in [9, 10, 11]:  # 1D, 5D, 1M columns
                value_float = float(value)
                if value_float > 0:
                    item.setForeground(QColor('green'))
                elif value_float < 0:
                    item.setForeground(QColor('red'))
            self.table.setItem(row_position, i, item)
        
        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(lambda: self.delete_stock(row['SYMBOL']))
        delete_button.setStyleSheet("""
            background-color: #e74c3c;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 3px;
        """)
        self.table.setCellWidget(row_position, 12, delete_button)
        
    def add_stock(self):
        symbol = self.stock_input.text().upper()
        if symbol and symbol not in self.watchlist:
            if symbol in self.df['SYMBOL'].values:
                self.watchlist.append(symbol)
                self.update_table()
                self.stock_input.clear()
            else:
                QMessageBox.warning(self, "Stock Not Found", "Stock unavailable in the data. Please try a different stock or correct its spelling.")
        elif symbol in self.watchlist:
            QMessageBox.information(self, "Duplicate Stock", "This stock is already in your watchlist.")
        
    def delete_stock(self, symbol):
        self.watchlist.remove(symbol)
        self.update_table()
        
    def save_watchlist(self):
        with open("watchlist.txt", "w") as f:
            for symbol in self.watchlist:
                f.write(f"{symbol}\n")
        QMessageBox.information(self, "Watchlist Saved", "Your watchlist has been saved successfully.")
        
    def load_watchlist(self):
        try:
            with open("watchlist.txt", "r") as f:
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StockWatchlistApp()
    window.show()
    sys.exit(app.exec_())