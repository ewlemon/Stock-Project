import yfinance as yf
import pandas as pd
import os

# ---------------------------
# 1. Set up output folder
# ---------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))  # folder where the script is located

excel_path = os.path.join(script_dir, "Major_Markets_Closing_10y.xlsx")
csv_path = os.path.join(script_dir, "Major_Markets_Closing_10y.csv")

# ---------------------------
# 2. Define indices and names
# ---------------------------
indices = ["^GSPC", "^DJI", "^IXIC", "^RUT", "^NDX"]

index_names = {
    "^GSPC": "S&P 500",
    "^DJI": "Dow Jones Industrial Average",
    "^IXIC": "Nasdaq Composite",
    "^RUT": "Russell 2000",
    "^NDX": "Nasdaq 100"
}

# ---------------------------
# 3. Download data
# ---------------------------
data_dict = {}

for ticker in indices:
    index = yf.Ticker(ticker)
    df = index.history(period="10y")
    df = df[['Close']]
    df.reset_index(inplace=True)
    df['Index'] = ticker
    df['Name'] = index_names[ticker]
    data_dict[ticker] = df
    print(f"Downloaded {ticker} ({index_names[ticker]})")

# ---------------------------
# 4. Combine into one DataFrame
# ---------------------------
combined_df = pd.concat(data_dict.values(), ignore_index=True)

# Remove timezone info to prevent Excel errors
combined_df['Date'] = combined_df['Date'].dt.tz_localize(None)

# ---------------------------
# 5. Save files
# ---------------------------
combined_df.to_excel(excel_path, index=False)
combined_df.to_csv(csv_path, index=False)

print(f"All major market closing prices (last 10 years) saved.")
print(f"Excel: {excel_path}")
print(f"CSV:   {csv_path}")
