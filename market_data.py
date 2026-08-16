import yfinance as yf
import pandas as pd
import os
import sys
import time
from functools import reduce
from yfinance.exceptions import YFRateLimitError

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
    except Exception:
        print("Excel exists but sheet 'Trading Days' not found. Will download full history.")
else:
    print("No cached Excel found. Downloading full history since 2000-01-01.")


# ---------------------------
# 4. Download data (incremental) with retry/backoff
# ---------------------------
def download_with_retry(ticker_symbol, start, max_retries=4, base_delay=5):
    """Fetch history for one ticker, retrying with exponential backoff on rate limits."""
    ticker_obj = yf.Ticker(ticker_symbol)
    for attempt in range(max_retries):
        try:
            df = ticker_obj.history(start=start)[['Close']]
            return df
        except YFRateLimitError:
            if attempt < max_retries - 1:
                wait = base_delay * (2 ** attempt)  # 5s, 10s, 20s, 40s
                print(f"  Rate limited on {ticker_symbol}. Waiting {wait}s "
                      f"(attempt {attempt + 1}/{max_retries})...")
                time.sleep(wait)
            else:
                print(f"  Giving up on {ticker_symbol} after {max_retries} attempts.")
                return pd.DataFrame()
        except Exception as e:
            # Any other transient error (network blip, empty response, etc.)
            print(f"  Error downloading {ticker_symbol}: {e}")
            if attempt < max_retries - 1:
                wait = base_delay * (2 ** attempt)
                time.sleep(wait)
            else:
                return pd.DataFrame()
    return pd.DataFrame()


data_dict = {}
start_date = last_date if last_date is not None else "2000-01-01"

for ticker in indices:
    print(f"Downloading {ticker} ({index_names[ticker]})...")
    df = download_with_retry(ticker, start_date)

    if df.empty:
        print(f"No new data for {ticker}.")
        time.sleep(3)  # still pace out requests even on failure/empty result
        continue

    df.rename(columns={'Close': index_names[ticker]}, inplace=True)
    df.reset_index(inplace=True)
    df['Date'] = df['Date'].dt.tz_localize(None)
    data_dict[ticker] = df

    time.sleep(3)  # space out requests between tickers to avoid tripping the rate limiter

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

# If every ticker failed AND there's no cached data to fall back on, bail out cleanly
# instead of crashing on a None dataframe.
if trading_df is None or trading_df.empty:
    print("No data available (all downloads failed and no cached data exists). Exiting without update.")
    sys.exit(0)

trading_df.drop_duplicates(subset=['Date'], inplace=True)
trading_df.sort_values(by="Date", inplace=True)

# ---------------------------
# 6. Add numeric date column for regression
# ---------------------------
trading_df['Numeric Date'] = (trading_df['Date'] - trading_df['Date'].min()).dt.days + 1

# ---------------------------
# 7. Add returns (percent)
# ---------------------------
for col in index_names.values():
    if col in trading_df.columns:
        trading_df[f"{col} % Return"] = trading_df[col].pct_change().fillna(0)

# ---------------------------
# 8. Reorder columns: Date, Numeric Date, then for each index: Close | % Return
# ---------------------------
cols_order = ['Date', 'Numeric Date']
for col in index_names.values():
    if col in trading_df.columns:
        cols_order += [col, f"{col} % Return"]

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