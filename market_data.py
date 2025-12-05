import yfinance as yf
import pandas as pd
import os
from functools import reduce
import numpy as np

# ---------------------------
# 1. Set up output folder
# ---------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(script_dir, "Top5_Indices.xlsx")

# ---------------------------
# 2. Define indices and names (Top 5)
# ---------------------------
indices = ["^GSPC", "^DJI", "^IXIC", "^RUT", "^NDX"]
index_names = {
    "^GSPC": "S&P 500",
    "^DJI": "Dow Jones",
    "^IXIC": "Nasdaq Composite",
    "^RUT": "Russell 2000",
    "^NDX": "Nasdaq 100"
}

# ---------------------------
# 3. Load cached Trading Days if available
# ---------------------------
cached_trading = None
last_date = None

if os.path.exists(excel_path):
    try:
        cached_trading = pd.read_excel(excel_path, sheet_name="Trading Days")
        cached_trading["Date"] = pd.to_datetime(cached_trading["Date"])
        last_date = cached_trading["Date"].max()
        print(f"Cached data found. Last available date: {last_date.date()}")
    except:
        print("Excel exists but sheet 'Trading Days' not found. Will download full history.")
else:
    print("No cached Excel found. Downloading full history since 2000-01-01.")

# ---------------------------
# 4. Download data (incremental)
# ---------------------------
data_dict = {}
for ticker in indices:
    print(f"Downloading {ticker} ({index_names[ticker]})...")
    index = yf.Ticker(ticker)
    if last_date is not None:
        df = index.history(start=last_date)[['Close']]
    else:
        df = index.history(start="2000-01-01")[['Close']]

    if df.empty:
        print(f"No new data for {ticker}.")
        continue

    df.rename(columns={'Close': index_names[ticker]}, inplace=True)
    df.reset_index(inplace=True)
    df['Date'] = df['Date'].dt.tz_localize(None)
    data_dict[ticker] = df

# ---------------------------
# 5. Merge new data
# ---------------------------
if data_dict:
    dfs = list(data_dict.values())
    trading_df = reduce(lambda left, right: pd.merge(left, right, on='Date', how='outer'), dfs)
    if cached_trading is not None:
        trading_df = pd.concat([cached_trading, trading_df], ignore_index=True)
else:
    trading_df = cached_trading.copy() if cached_trading is not None else None

trading_df.drop_duplicates(subset=['Date'], inplace=True)
trading_df.sort_values(by="Date", inplace=True)

# ---------------------------
# 6. Add numeric date column for regression
# ---------------------------
trading_df['Numeric Date'] = (trading_df['Date'] - trading_df['Date'].min()).dt.days + 1

# ---------------------------
# 7. Add returns (percent and log)
# ---------------------------
for col in index_names.values():
    trading_df[f"{col} % Return"] = trading_df[col].pct_change().fillna(0)
    trading_df[f"{col} Log Return"] = np.log(trading_df[col] / trading_df[col].shift(1)).fillna(0)

# ---------------------------
# 8. Reorder columns: Date, Numeric Date, then for each index: Close | % Return | Log Return
# ---------------------------
cols_order = ['Date', 'Numeric Date']
for col in index_names.values():
    cols_order += [col, f"{col} % Return", f"{col} Log Return"]

trading_df = trading_df[cols_order]

# ---------------------------
# 9. Save Excel file (overwrite only Trading Days sheet)
# ---------------------------
if os.path.exists(excel_path):
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        trading_df.to_excel(writer, index=False, sheet_name="Trading Days")
else:
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        trading_df.to_excel(writer, index=False, sheet_name="Trading Days")

print("Excel updated. Trading Days sheet reordered and returns added.")
