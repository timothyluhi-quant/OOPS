import numpy as np
import pandas as pd
import requests
import math
from scipy import stats
import xlsxwriter
import yfinance as yf
import io
from statistics import mean
from scipy.stats import percentileofscore as score

# --- STEP 1: SCRAPE WIKIPEDIA ---
url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
header = {
    'User-Agent': 'David/5.0 (Windows NT 10.0; Win64; X64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

response = requests.get(url, headers=header)
payload = pd.read_html(io.StringIO(response.text))

sp500_table = None
for table in payload:
    if 'Symbol' in table.columns:
        sp500_table = table
        break

if sp500_table is None:
    print("Error: Could not find the S&P 500 table on Wikipedia.")
    fresh_tickers = []
else:
    fresh_tickers = sp500_table['Symbol'].str.replace('.', '-', regex=False).tolist()
    print(f'Successfully loaded {len(fresh_tickers)} live tickers from Wikipedia.')


# --- STEP 2: PORTFOLIO INPUT ---
def portfolio_input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio:')

    try:
        float(portfolio_size)
    except ValueError:
        print('That is not a number! \nPlease try again:')
        portfolio_size = input('Enter the size of your portfolio:')


portfolio_input()

# --- STEP 3: BUILD HQM STRATEGY ---
hqm_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'One-Year Price Return',
    'One-Year Return Percentile',
    'Six-Month Price Return',
    'Six-Month Return Percentile',
    'Three-Month Price Return',
    'Three-Month Return Percentile',
    'One-Month Price Return',
    'One-Month Return Percentile',
    'HQM Score'
]
hqm_rows = []

for symbol in fresh_tickers:
    try:
        # Download data with a buffer to ensure end dates are included
        df = yf.download(symbol, period='2y', progress=False)

        if df.empty: continue

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        last_trading_date = df.index[
            -1]  # this is the key step to a dynamic strategy, which is to lock the date dynamically
        current_price = float(df['Close'].iloc[-1])  # df.index[-1] is the last day in the data

        date_1y_target = last_trading_date - pd.DateOffset(years=1)  # automatically trace back the start date
        date_6m_target = last_trading_date - pd.DateOffset(
            months=6)  # and DateOffset will handle all kinds of unusual contions
        date_3m_target = last_trading_date - pd.DateOffset(months=3)
        date_1m_target = last_trading_date - pd.DateOffset(months=1)

        try:
            price_1y = float(df[:date_1y_target]['Close'].iloc[-1])  # starts from 2 year ago, ends in date_1y_target
            price_6m = float(
                df[:date_6m_target]['Close'].iloc[-1])  # collects the nearest close price, which is always correct
            price_3m = float(df[:date_3m_target]['Close'].iloc[-1])
            price_1m = float(df[:date_1m_target]['Close'].iloc[-1])

        except IndexError:
            continue

        hqm_rows.append([
            symbol,
            current_price,
            'N/A',
            (current_price - price_1y) / price_1y,
            'N/A',
            (current_price - price_6m) / price_6m,
            'N/A',
            (current_price - price_3m) / price_3m,
            'N/A',
            (current_price - price_1m) / price_1m,
            'N/A',
            'N/A'
        ])

    except Exception:
        pass

hqm_df = pd.DataFrame(hqm_rows, columns=hqm_columns)
print(f'Total stocks processed successfully: {len(hqm_df)}')

# --- STEP 4: CALCULATE PERCENTILES ---
time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']

for row in hqm_df.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_df.loc[row, percentile_col] = score(hqm_df[change_col], hqm_df.loc[row, change_col]) / 100

# --- STEP 5: CALCULATE HQM SCORE ---
for row in hqm_df.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(float(hqm_df.loc[row, f'{time_period} Return Percentile']))
    hqm_df.loc[row, 'HQM Score'] = mean(momentum_percentiles)

# --- STEP 6: SELECT TOP 50 ---
hqm_df.sort_values('HQM Score', ascending=False, inplace=True)
hqm_df = hqm_df[:50]
hqm_df.reset_index(drop=True, inplace=True)


# --- STEP 6.5: VOLATILITY FILTER (Remove top 10% highest volatility) ---
print("Applying Volatility Filter: removing highest 10% volatility stocks...")
vol_list = []
for symbol in hqm_df['Ticker']:
    try:
        df = yf.download(symbol, period='1y', progress=False)
        if df.empty: continue
        daily_ret = df['Close'].pct_change().dropna()
        vol = daily_ret.std() * np.sqrt(252)
        vol_list.append([symbol, vol])
    except Exception:
        continue

vol_df = pd.DataFrame(vol_list, columns=['Ticker', 'Volatility'])
hqm_df = hqm_df.merge(vol_df, on='Ticker', how='left')

# Remove missing volatility and filter top 10%
before_count = len(hqm_df)
hqm_df = hqm_df.dropna(subset=['Volatility'])
cut = hqm_df['Volatility'].quantile(0.90)
hqm_df = hqm_df[hqm_df['Volatility'] <= cut]
hqm_df.reset_index(drop=True, inplace=True)
after_count = len(hqm_df)

print(f"Volatility filter removed {before_count - after_count} stocks. New count: {after_count}")


# --- STEP 7: CALCULATE SHARES TO BUY ---
position_size = float(portfolio_size) / len(hqm_df.index)
for i in hqm_df.index:
    hqm_df.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / hqm_df.loc[i, 'Price'])

# --- STEP 8: EXCEL FORMATTING ---
writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
hqm_df.to_excel(writer, sheet_name='Momentum Strategy', index=False)

background_color = '#0a0a23'
font_color = '#ffffff'
workbook = writer.book

string_template = workbook.add_format({'font_color': font_color, 'bg_color': background_color, 'border': 1})
dollar_template = workbook.add_format(
    {'num_format': '$0.00', 'font_color': font_color, 'bg_color': background_color, 'border': 1})
integer_template = workbook.add_format(
    {'num_format': '0', 'font_color': font_color, 'bg_color': background_color, 'border': 1})
percent_template = workbook.add_format(
    {'num_format': '0.0%', 'font_color': font_color, 'bg_color': background_color, 'border': 1})

# Corrected Column Mapping (Removed 'Index' so it aligns with index=False)
column_formats = {
    'A': ['Ticker', string_template],
    'B': ['Price', dollar_template],
    'C': ['Number of Shares to Buy', integer_template],
    'D': ['One-Year Price Return', percent_template],
    'E': ['One-Year Return Percentile', percent_template],
    'F': ['Six-Month Price Return', percent_template],
    'G': ['Six-Month Return Percentile', percent_template],
    'H': ['Three-Month Price Return', percent_template],
    'I': ['Three-Month Return Percentile', percent_template],
    'J': ['One-Month Price Return', percent_template],
    'K': ['One-Month Return Percentile', percent_template],
    'L': ['HQM Score', percent_template]
}

worksheet = writer.sheets['Momentum Strategy']

for column in column_formats.keys():
    worksheet.set_column(f'{column}:{column}', 25, column_formats[column][1])
    worksheet.write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.close()
print(" Momentum Strategy Completed and Saved!")
