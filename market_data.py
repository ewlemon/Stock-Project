import yfinance as yf
import pandas as pd
import os

# ---------------------------
# 1. Set up output folder
# ---------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))  # folder where script is located

excel_path = os.path.join(script_dir, "Major_Markets_Closing.xlsx")
csv_path = os.path.join(script_dir, "Major_Markets_Closing.csv")

# ---------------------------
# 2. Define indices and names
# ---------------------------
indices = ["^GSPC", "^DJI", "^IXIC", "^RUT", "^NDX", "^FTSE", "^GDAXI"]  # example: added FTSE and DAX

index_names = {
    "^GSPC": "S&P 500",
    "^DJI": "Dow Jones",
    "^IXIC": "Nasdaq Composite",
    "^RUT": "Russell 2000",
    "^NDX": "Nasdaq 100",
    "^FTSE": "FTSE 100",
    "^GDAXI": "DAX"
}

# ---------------------------
# 3. Download data
# ---------------------------
data_dict = {}

for ticker in indices:
    index = yf.Ticker(ticker)
    df = index.history(period="20y")[['Close']]  # only closing prices
    df.rename(columns={'Close': index_names[ticker]}, inplace=True)  # rename column to common name
    df.reset_index(inplace=True)
    df['Date'] = df['Date'].dt.tz_localize(None)  # remove timezone
    data_dict[ticker] = df
    print(f"Downloaded {ticker} ({index_names[ticker]})")

# ---------------------------
# 4. Merge into one wide DataFrame
# ---------------------------
from functools import reduce

# Start merging on 'Date'
dfs = list(data_dict.values())
combined_df = reduce(lambda left, right: pd.merge(left, right, on='Date', how='outer'), dfs)

# ---------------------------
# 5. Save files
# ---------------------------
combined_df.to_excel(excel_path, index=False)
combined_df.to_csv(csv_path, index=False)

print(f"All major market closing prices (last 20 years) saved.")
print(f"Excel: {excel_path}")
print(f"CSV:   {csv_path}")
