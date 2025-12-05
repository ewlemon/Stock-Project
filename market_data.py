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
    except Exception as e:
        print(f"Could not read Historical sheet. Downloading full history. Error: {e}")
        cached_df = None

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
    combined_df = cached_df.copy() if cached_df is not None else pd.DataFrame()

# Remove duplicates and sort
combined_df.drop_duplicates(subset=['Date'], inplace=True)
combined_df.sort_values(by="Date", inplace=True)

# ---------------------------
# 6. Create continuous daily series (forward-fill)
# ---------------------------
if not combined_df.empty:
    continuous_df = combined_df.copy()
    full_date_range = pd.date_range(start=continuous_df['Date'].min(),
                                    end=continuous_df['Date'].max(),
                                    freq='B')  # business days

    continuous_df.set_index('Date', inplace=True)
    continuous_df = continuous_df.reindex(full_date_range)
    continuous_df.ffill(inplace=True)  # forward-fill missing days
    continuous_df.reset_index(inplace=True)
    continuous_df.rename(columns={'index': 'Date'}, inplace=True)
else:
    continuous_df = pd.DataFrame()

# ---------------------------
# 7. Save/update Excel safely
# ---------------------------
if not os.path.exists(excel_path):
    # Create new workbook
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Historical")
        continuous_df.to_excel(writer, index=False, sheet_name="Continuous")
    print(f"Created new Excel file: {excel_path}")
else:
    # Update existing workbook without touching Forecast sheet
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Historical")
        continuous_df.to_excel(writer, index=False, sheet_name="Continuous")
    print(f"Updated Excel file: {excel_path}")

print("Script completed successfully.")
