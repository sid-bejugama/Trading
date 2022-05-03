from typing import final
from charset_normalizer import api
import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from secrets import IEX_CLOUD_API_TOKEN

# algorithm that weights S&P500 stocks equally regardless of market capitalization 

stocks = pd.read_csv('sp_500_stocks.csv')
symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
price = data['latestPrice']
marketCap = data['marketCap']

columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']

final_dataframe = pd.DataFrame(columns = columns)


def list_fragment(list, fragment_size):
    for i in range(0, len(list), fragment_size):
        yield list[i: i + fragment_size]
    
symbol_groups = list(list_fragment(stocks['Ticker'], 100))
symbol_strings = []
for symbol_group in symbol_groups:
    symbol_strings.append(','.join(symbol_group))
final_dataframe = pd.DataFrame(columns = columns)
for symbol_string in symbol_strings:
    batch_api_call__url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call__url).json()
    for symbol in symbol_string.split(','):
        if symbol in ['HFC', 'VIAC', 'WLTW']: continue
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol, 
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['marketCap'],
                    'N/A'

                ], 
                index = columns
            ),
            ignore_index = True
        )
portfolio_size = input('Enter the value of your portfolio:')

try:
    val = float(portfolio_size)
except ValueError:
    print("That's not a number! \nPlease try again.")
    portfolio_size = input('Enter the value of your portfolio:')
    val = float(portfolio_size)

position_size = val/len(final_dataframe.index)
for i in range(len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1,
    }
)
dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1,
    }
)
integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1,
    }
)
column_formats = {
    'A':['Ticker', string_format],
    'B':['Stock Price', dollar_format],
    'C':['Market Capitalization', dollar_format],
    'D':['Number of Shares to Buy', integer_format]
}
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()


