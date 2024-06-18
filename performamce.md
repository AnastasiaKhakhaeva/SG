# installing packages
#python -m pip install yahoo_fin
#python -m pip install investpy
#python -m pip install tradingeconomics
#python -m pip install pandasdmx pandas openpyxl
#python -m pip install fsspec

import pandas as pd
import pip
import yfinance as yf
import openpyxl as ox
import yahoo_fin
import investpy

start_date="2020-01-01"
end_date="2024-06-05"

tickers_df = pd.read_excel('tickers_list.xlsx') 
tickers = tickers_df['Ticker'].tolist() 

# prices
downloaded_data = yf.download(tickers, start=start_date, end=end_date, repair=True)
df_prices = pd.DataFrame(downloaded_data)

# currency matching
def get_currency(tickers):
    currencies = {}
    for ticker in tickers:
        try:
            security = yf.Ticker(ticker)
            currency = security.info['currency']
            currencies[ticker] = currency
        except:
            currencies[ticker] = "Error: Unable to retrieve currency"
    return currencies

currencies = get_currency(tickers)

df_FX_by_Ticker = pd.DataFrame(list(currencies.items()), columns=['Ticker', 'Currency'])


# currency
tickers_df = pd.read_excel('tickers_list.xlsx') 
tickers = tickers_df['Ticker'].tolist() 

FX_df = pd.read_excel('FX_list.xlsx') 
FX = FX_df['Currency'].tolist() 
downloaded_FX = yf.download(FX, start=start_date, end=end_date, repair=True)
df_FX = pd.DataFrame(downloaded_FX)

# divs
dividends = yf.download(tickers, start=start_date, end=end_date, actions='dividends')

# getting the file
with pd.ExcelWriter('Market_data.xlsx') as writer:
    df_prices.to_excel(writer, sheet_name='Close price')
    df_FX_by_Ticker.to_excel(writer, sheet_name='FX matching', index=False)
    df_FX.to_excel(writer,sheet_name='FX')
    dividends.to_excel(writer, sheet_name='divs')



