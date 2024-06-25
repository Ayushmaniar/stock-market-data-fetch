import pandas as pd
import yfinance as yf
from datetime import datetime
from tqdm import tqdm
from urllib.parse import quote_plus

# Read list of stocks from the CSV file
stocks_df = pd.read_csv("EQUITY_L.csv")
stocks_df['YahooEquiv'] = stocks_df['SYMBOL'] + '.NS'

# List of NIFTY 50 stocks
yahoo_finance_symbols = list(stocks_df['YahooEquiv'])
error_companies = []

progress_bar = tqdm(total=len(yahoo_finance_symbols))
import time

# Download data for the current day using yfinance for all the stocks mentioned in the CSV file
today = datetime.today().strftime('%Y-%m-%d')


stock_data = {}

for company_no , company in enumerate(yahoo_finance_symbols):
    time.sleep(0.01)
    data = yf.download(company, start=today, end=pd.to_datetime(today) + pd.Timedelta(days=1), progress = False)
    if not data.empty:
        stock_data[company] = data

    progress_bar.set_description("Processing item {}".format(company, company_no , len(yahoo_finance_symbols)))
    progress_bar.update(1)

    # if company_no == 10:
    #     break



progress_bar.close()

all_stock_data = pd.DataFrame()

for symbol, data in stock_data.items():
    if not data.empty:
        data['SYMBOL'] = symbol
        all_stock_data = pd.concat([all_stock_data, data])

all_stock_data.reset_index(inplace=True)

all_stock_data['percentage_change'] = (all_stock_data['Close'] - all_stock_data['Open']) / all_stock_data['Open'] * 100
all_stock_data.sort_values(by='percentage_change', ascending=False, inplace=True)

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

print(f"Data saved to {output_file_name}")