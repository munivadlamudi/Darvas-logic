import pandas as pd
import yfinance as yf

def darvas_box_strategy(stock_symbol, start_date, end_date):
    # Fetch historical data for the stock
    data = yf.download(stock_symbol, start=start_date, end=end_date)
    data['High_box'] = data['High'].rolling(window=60).max()  # High box level (60-day rolling max)
    data['Low_box'] = data['Low'].rolling(window=60).min()    # Low box level (60-day rolling min)
    
    # Identify Buy and Sell signals
    data['Buy_signal'] = (data['Close'] > data['High_box'].shift(1)) & (data['Volume'] > data['Volume'].mean())
    data['Sell_signal'] = (data['Close'] < data['Low_box'].shift(1))
    
    return data

def get_latest_price(stock_symbol):
    # Fetch the latest price for the stock
    ticker = yf.Ticker(stock_symbol)
    latest_data = ticker.history(period="1d")
    if not latest_data.empty:
        ltp = latest_data['Close'].iloc[-1]  # Get the latest closing price
        return ltp
    return None

# Read stock symbols from CSV
stocks_df = pd.read_csv('stock_symbols.csv')  # Adjust the filename to your actual file
stock_symbols = stocks_df['stock_symbol'].tolist()  # Assuming the column name is 'stock_symbol'

# Set date range for fetching data
start_date = '2018-01-01'
end_date = '2024-09-08'

# Prepare an empty list to store results
report_data = []

# Iterate over each stock symbol
for stock in stock_symbols:
    print(f"Processing {stock}...")
    darvas_data = darvas_box_strategy(stock, start_date, end_date)
    
    # Extract Buy and Sell signals with necessary information
    buy_signals = darvas_data[darvas_data['Buy_signal']].reset_index()
    sell_signals = darvas_data[darvas_data['Sell_signal']].reset_index()

    # Initialize variables to track the carrying position
    carrying_position = False
    entry_date = entry_price = exit_date = exit_price = high_box = low_box = ltp = None

    # Iterate through Buy signals and match with Sell signals
    for i, buy_row in buy_signals.iterrows():
        if not carrying_position:  # Only consider new buys if no position is carried
            # Set carrying position to True when a buy is initiated
            entry_date = buy_row['Date']
            entry_price = buy_row['Close']
            high_box = buy_row['High_box']
            low_box = buy_row['Low_box']
            carrying_position = True

            # Find the first Sell signal that comes after the Buy signal
            sell_signal = sell_signals[sell_signals['Date'] > entry_date].head(1)
            if not sell_signal.empty:
                # Update exit details
                exit_date = sell_signal.iloc[0]['Date']
                exit_price = sell_signal.iloc[0]['Close']
                
                # Calculate ROI using the exit price
                roi = ((exit_price - entry_price) / entry_price) * 100
                
                # Append combined Buy and Sell information
                report_data.append([stock, entry_date, entry_price, high_box, low_box, 'Buy', exit_date, exit_price, 'Sell', None, roi])
                
                # Reset carrying position as it was sold
                carrying_position = False
            else:
                # If no exit signal is found, fetch the latest price (LTP)
                ltp = get_latest_price(stock)
                
                # Calculate ROI using the LTP
                if ltp is not None:
                    roi = ((ltp - entry_price) / entry_price) * 100
                else:
                    roi = None  # If LTP couldn't be fetched, set ROI as None
                
                # Append the data with the LTP as the exit price for an open position
                report_data.append([stock, entry_date, entry_price, high_box, low_box, 'Buy', None, None, None, ltp, roi])
                # Continue to carry the position without resetting

# Convert the report data to a DataFrame
report_df = pd.DataFrame(report_data, columns=['Stock Symbol', 'Entry Date', 'Entry Price', 'High Box', 'Low Box', 'Entry Signal', 'Exit Date', 'Exit Price', 'Exit Signal', 'LTP', 'ROI (%)'])

# Sort the DataFrame by 'Stock Symbol' and 'Entry Date' in descending order
report_df = report_df.sort_values(by=['Stock Symbol', 'Entry Date'], ascending=[True, False])

# Save the report to Excel
report_df.to_excel('Reports/Darvas_Box_Trade_Combined_Report.xlsx', index=False)
print("Report saved as 'Darvas_Box_Trade_Combined_Report.xlsx'")
