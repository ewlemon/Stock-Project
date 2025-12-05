import yfinance as yf
import pandas as pd
import os
from functools import reduce
from openpyxl import load_workbook

# ---------------------------
# 1. Setup output
# ---------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(script_dir, "Top5_Indices.xlsx")

# ---------------------------
# 2. Define indices
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
# 3. Load cached data if exists
# ---------------------------
cached_df = None
last_date = None

if os.path.exists(excel_path):
    try:
        cached_df = pd.read_excel(excel_path, sheet_name="Historical")
        cached_df["Date"] = pd.to_datetime(cached_df["Date"])
        last_date = cached_df["Date"].max()
        print(f"Cached data found. Last available date: {last_date.date()}")
    except:
        cached_df = None
        print("Could not read Historical sheet. Will download full history.")
else:
    print("No cached data found. Downloading full history since 2000-01-01.")

# ---------------------------
# 4. Download/update data
# ---------------------------
data_dict = {}
for ticker in indices:
    print(f"Downloading {ticker} ({index_names[ticker]})...")
    index = yf.Ticker(ticker)
    
    start_date = last_date if last_date is not None else "2000-01-01"
    df = index.history(start=start_date)[['Close']]
    
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
    new_data_df = reduce(lambda left, right: pd.merge(left, right, on='Date', how='outer'), dfs)
    
    if cached_df is not None:
        combined_df = pd.concat([cached_df, new_data_df], ignore_index=True)
    else:
        combined_df = new_data_df
else:
    combined_df = cached_df.copy()

# Remove duplicates and sort
combined_df.drop_duplicates(subset=['Date'], inplace=True)
combined_df.sort_values(by="Date", inplace=True)

# ---------------------------
# 6. Create continuous daily series (forward-fill)
# ---------------------------
continuous_df = combined_df.copy()
full_date_range = pd.date_range(start=continuous_df['Date'].min(),
                                end=continuous_df['Date'].max(),
                                freq='B')  # business days

continuous_df.set_index('Date', inplace=True)
continuous_df = continuous_df.reindex(full_date_range)
continuous_df.fillna(method='ffill', inplace=True)
continuous_df.reset_index(inplace=True)
continuous_df.rename(columns={'index': 'Date'}, inplace=True)

# ---------------------------
# 7. Save/update Excel without overwriting Forecast sheet
# ---------------------------
if os.path.exists(excel_path):
    # Load existing workbook
    book = load_workbook(excel_path)
else:
    book = None

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    if book:
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
    
    # Update Historical and Continuous sheets
    combined_df.to_excel(writer, index=False, sheet_name="Historical")
    continuous_df.to_excel(writer, index=False, sheet_name="Continuous")
    
    # Forecast sheet remains untouched
print(f"Excel file updated: {excel_path}")
