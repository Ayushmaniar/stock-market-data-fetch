# StockWatch - Indian Stock Market Tracker

A PyQt5-based desktop application for tracking Indian stock market data. Download stock data from Yahoo Finance for NSE-listed stocks, calculate performance metrics (1D, 5D, 1M returns), and monitor your watchlist.

![StockWatch App](https://github.com/Ayushmaniar/stock-market-data-fetch/assets/35158611/55b86663-67aa-4370-ba6a-8eae31ce4f6c)

## Features

- 📊 Track NSE-listed stocks with real-time data from Yahoo Finance
- 📈 Calculate 1-day, 5-day, and 1-month performance metrics
- ⭐ Create and manage a custom watchlist
- 💾 Export data to Excel format
- 🔍 Auto-complete search for stock symbols
- 🔗 Clickable stock symbols linking to Google search
- 🎨 Clean, modern UI with progress visualization

## Installation

### Prerequisites

- Python 3.8 or higher
- Windows / macOS / Linux

### Install Dependencies

```bash
pip install pandas yfinance PyQt5 tqdm openpyxl xlsxwriter
```

### For Building Executables (Optional)

```bash
pip install PyInstaller
```

## Usage

### Running in Development Mode

```bash
python main.py
```

### Running with Tests

Run all tests first, then launch the application only if tests pass:

```bash
python run_with_tests.py
```

### Running Tests Only

```bash
# Run all tests
python -m unittest discover stockwatch/tests

# Run specific test files
python -m unittest stockwatch.tests.test_data_processor
python -m unittest stockwatch.tests.test_ui_components
python -m unittest stockwatch.tests.test_main_window
```

## Building the Executable

Build a standalone executable (takes 3-5 minutes):

```bash
python build_exe.py
```

The executable will be created in `dist/StockMarketApp/StockMarketApp.exe`

The build script will:
- ✅ Automatically clean old build artifacts
- ✅ Exclude unnecessary packages (PyTorch, TensorFlow) for faster builds
- ✅ Bundle required files (EQUITY_L.csv, watchlist.txt)
- ✅ Show detailed progress with timing information

## First-Time Setup

When you run the application for the first time:

1. The app will prompt you to download data for all stocks (**10-15 minutes minimum**)
2. Data is saved to the `yahoo_finance_data/` folder as an Excel file
3. After the initial download completes, **close and reopen the app**
4. Then you can add stocks to your watchlist and use the quick refresh feature

## Project Structure

```
stock-market-data-fetch/
├── stockwatch/              # Main application package
│   ├── views/              # UI components and main window
│   ├── data/               # Data download and processing
│   ├── models/             # Data models (placeholder)
│   ├── controllers/        # Business logic (placeholder)
│   ├── utils/              # Utility functions (placeholder)
│   └── tests/              # Unit tests
├── main.py                 # Application entry point
├── build_exe.py            # Build script for creating executable
├── run_with_tests.py       # Test runner + launcher
├── EQUITY_L.csv            # NSE stock symbols list
└── watchlist.txt           # User's watchlist (auto-generated)
```

## Application Features

### Download Modes

- **DOWNLOAD ALL**: Downloads data for all NSE-listed stocks (~1500+ stocks)
- **QUICK REFRESH**: Downloads data only for stocks in your watchlist

### Data Storage

- Downloaded data is stored in `yahoo_finance_data/` as Excel files
- Format: `YYYY-MM-DD_stock_market_data.xlsx`
- Includes: Open, High, Low, Close, Volume, and calculated metrics

### Watchlist Management

- Add stocks using the search bar with auto-complete
- Remove stocks with the delete button
- Save your watchlist to persist between sessions
- Watchlist stored in `watchlist.txt` (one symbol per line)

### Weekend/Holiday Handling

- Automatically detects weekends (Saturday/Sunday)
- Prompts to download data for the last trading day (Friday)
- Uses business day calculations for accurate metrics

## Troubleshooting

### Build Fails with Permission Error

If the build fails with "Could not remove dist/", ensure:
1. Close any running instances of `StockMarketApp.exe`
2. Close File Explorer windows showing the `dist` folder
3. Check Task Manager for background processes

The build script will automatically retry with delays.

### App Runs as Background Process

The app now includes proper cleanup on exit. If a download is in progress when you close the window, you'll be prompted to stop it.

### Missing Stock Data

Some stocks may be delisted or have no recent trading data. These are logged in `stock_data_debug.log` and skipped during processing.

## Contributing

Feel free to open issues or submit pull requests for improvements!

## License

This project is open source and available under the MIT License.
