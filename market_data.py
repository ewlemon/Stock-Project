import yfinance as yf
import pandas as pd
import os
from functools import reduce

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
    new_trading_df = reduce(lambda left, right: pd.merge(left, right, on='Date', how='outer'), dfs)
    if cached_trading is not None:
        trading_df = pd.concat([cached_trading, new_trading_df], ignore_index=True)
    else:
        trading_df = new_trading_df
else:
    trading_df = cached_trading.copy() if cached_trading is not None else None

trading_df.drop_duplicates(subset=['Date'], inplace=True)
trading_df.sort_values(by="Date", inplace=True)

# ---------------------------
# 6. Create continuous sheet
# ---------------------------
continuous_df = trading_df.copy()
continuous_df.set_index('Date', inplace=True)
continuous_df = continuous_df.asfreq('D')  # all calendar days
continuous_df.ffill(inplace=True)          # forward-fill missing
continuous_df.reset_index(inplace=True)

# ---------------------------
# 7. Add DayNumber column for regression
# ---------------------------
for df in [trading_df, continuous_df]:
    df['DayNumber'] = (df['Date'] - df['Date'].min()).dt.days + 1

# ---------------------------
# 8. Save Excel file (preserve other sheets)
# ---------------------------
if os.path.exists(excel_path):
    # Open workbook in append mode, overwrite only target sheets
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        continuous_df.to_excel(writer, index=False, sheet_name="Continuous")
        trading_df.to_excel(writer, index=False, sheet_name="Trading Days")
else:
    # Workbook doesn't exist; create new
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        continuous_df.to_excel(writer, index=False, sheet_name="Continuous")
        trading_df.to_excel(writer, index=False, sheet_name="Trading Days")

print("Excel updated with Continuous and Trading Days sheets. Other sheets are preserved.")
