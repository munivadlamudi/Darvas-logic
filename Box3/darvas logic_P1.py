import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

def darvas_box_strategy(stock_symbol, start_date, end_date):
    # Fetch historical data for the stock
    data = yf.download(stock_symbol, start=start_date, end=end_date)
    data['High_box'] = data['High'].rolling(window=20).max()  # High box level (20-day rolling max)
    data['Low_box'] = data['Low'].rolling(window=20).min()    # Low box level (20-day rolling min)
    
    # Identify Buy and Sell signals
    data['Buy_signal'] = (data['Close'] > data['High_box'].shift(1)) & (data['Volume'] > data['Volume'].mean())
    data['Sell_signal'] = (data['Close'] < data['Low_box'].shift(1))
    
    return data

# Read stock symbols from CSV
stocks_df = pd.read_csv('stock_symbols.csv')  # Adjust the filename to your actual file
stock_symbols = stocks_df['stock_symbol'].tolist()  # Assuming the column name is 'stock_symbol'

# Set date range for fetching data 
start_date = '2018-01-01'
end_date = '2024-09-08'

# Prepare an empty DataFrame to store results
report_data = []

# Iterate over each stock symbol
for stock in stock_symbols:
    print(f"Processing {stock}...")
    darvas_data = darvas_box_strategy(stock, start_date, end_date)
    
    # Extract Buy and Sell signals with necessary information
    buy_signals = darvas_data[darvas_data['Buy_signal']]
    sell_signals = darvas_data[darvas_data['Sell_signal']]
    
    # Add Buy signals to the report
    for index, row in buy_signals.iterrows():
        report_data.append([stock, index, row['Close'], 'Buy'])
    
    # Add Sell signals to the report
    for index, row in sell_signals.iterrows():
        report_data.append([stock, index, row['Close'], 'Sell'])

# Convert the report data to a DataFrame
report_df = pd.DataFrame(report_data, columns=['Stock Symbol', 'Date', 'Price', 'Signal'])

# Save the report to Excel
report_df.to_excel('Reports/Darvas_Box_Trade_Report.xlsx', index=False)
print("Report saved as 'Darvas_Box_Trade_Report.xlsx'")