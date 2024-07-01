## Usage

```
pip install pandas yfinance PyQt5 tqdm openpyxl xlsxwriter
pip install PyInstaller
```

### Once you have installed all these libraries, you would need to actually build the app with this command.
Note : This command might take 3-4 minutes to run

```
python -m PyInstaller final_app_build_script.py
```
### After this you would be able to locate the app inside the dist/final_app_build_script folder

### Note : When you would run the app for the first time, the app would download data for all stocks 
This might take around 10-15 minutes atleast, once the data has been downloaded, close the app.
You would be able to view the data in an excel sheet inside the "yahoo_finance_data" folder, post this you can reopen the app and start adding stocks to your watchlist



![image](https://github.com/Ayushmaniar/stock-market-data-fetch/assets/35158611/55b86663-67aa-4370-ba6a-8eae31ce4f6c)
