import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from tqdm import tqdm
from urllib.parse import quote_plus
import time
from pandas.tseries.offsets import BDay
import warnings
warnings.filterwarnings('ignore')

# Read list of stocks from the CSV file
stocks_df = pd.read_csv("EQUITY_L.csv")
stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'

# List of NIFTY 50 stocks
yahoo_finance_symbols = list(stocks_df['YahooEquiv'])
error_companies = []

progress_bar = tqdm(total=len(yahoo_finance_symbols))

# Download data for the current day using yfinance for all the stocks mentioned in the CSV file
# Download data for the current day using yfinance for all the stocks mentioned in the CSV file
today = datetime.today().strftime('%Y-%m-%d')
pandas_today = pd.Timestamp(today)
last_trading_day = (pandas_today - BDay(1)).strftime('%Y-%m-%d')
five_trading_days_ago = (pandas_today - BDay(5)).strftime('%Y-%m-%d')

one_month_ago = pandas_today - pd.DateOffset(months=1)
if one_month_ago.weekday()> 4:
    one_month_ago = one_month_ago + pd.DateOffset(days=(7 - one_month_ago.weekday()))
one_month_ago = one_month_ago.strftime('%Y-%m-%d')


stock_data = {}
max_date = pd.Timestamp('2008-01-01')

for company_no, company in enumerate(yahoo_finance_symbols):
    time.sleep(0.01)
    fetch_data = yf.download(company, start=one_month_ago, end=pd.to_datetime(today) + pd.Timedelta(days=1), progress=False)
    try:
        if fetch_data.reset_index()['Date'].max() > max_date:
            max_date = fetch_data.reset_index()['Date'].max()
    except:
        pass

    if not fetch_data.empty:
        stock_data[company] = fetch_data

    progress_bar.set_description("Processing item {}".format(company, company_no, len(yahoo_finance_symbols)))
    progress_bar.update(1)


progress_bar.close()

print(max_date)

if max_date.strftime('%Y-%m-%d') != today:
    raise ValueError('''\n\n
                     No stock data was successfully downloaded, most likely today is a holiday. \n\n
                     Please try downloading again later once there is a working stock market day \n\n''')


all_stock_data = pd.DataFrame()

for symbol, data in stock_data.items():
    if not data.empty:
        data = data.reset_index()
        data['SYMBOL'] = symbol
        single_row = data.loc[data['Date'] == today]
        todays_close = single_row['Close']
        prev_close = data.loc[data['Date'] == last_trading_day, 'Close'].values[0]
        five_days_close = data.loc[data['Date'] == five_trading_days_ago, 'Close'].values[0]
        one_month_close = data.loc[data['Date'] == one_month_ago, 'Close'].values[0]
        single_row['Previous_Close'] = prev_close
        single_row['1D'] = (todays_close - prev_close)/(prev_close) * 100
        single_row['5D'] = (todays_close - five_days_close)/(five_days_close) * 100
        single_row['1M'] = (todays_close - one_month_close)/(one_month_close) * 100

    all_stock_data = pd.concat([all_stock_data, single_row])

all_stock_data.reset_index(inplace=True,drop=True)

all_stock_data['Date'] = all_stock_data['Date'].astype(str)

all_stock_data.sort_values(by='1D', ascending=False, inplace=True)

all_stock_data = all_stock_data.round(2)

# Move 'SYMBOL' column to the first position
cols = list(all_stock_data.columns)
cols.insert(0, cols.pop(cols.index('SYMBOL')))
all_stock_data = all_stock_data[cols]

all_stock_data.reset_index(inplace= True, drop = True)

# Create a new column containing the Google search URL for each symbol
google_search_urls = 'https://www.google.com/search?q=' + all_stock_data['SYMBOL'].str.replace('.NS','').apply(quote_plus) + '+share+price'
google_search_urls = google_search_urls.sort_index()

output_file_name = f'yahoo_finance_data/{today}_stock_market_data.xlsx'

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

    for col_num, col in enumerate(all_stock_data.columns):
        max_length = max(all_stock_data[col].astype(str).apply(len).max(), len(col)) + 2  # +2 for some padding
        worksheet.set_column(col_num, col_num, max_length)

print(f"Data saved to {output_file_name}")