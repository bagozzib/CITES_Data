import pandas as pd
import numpy as np
import requests
import time

API_KEY     = "451c65d84143f62db9d0a6ab9d41a6cb"
INPUT_PATH  = r"C:\Users\emuru\Downloads\CITES Extracted Data V2 (2).xlsx"

START_ROW   = 14002
END_ROW     = 18688

URL_TEMPLATE = "https://v2.namsor.com/NamSorAPIv2/api2/json/gender/{first}/{last}"
HEADERS      = {"X-API-KEY": API_KEY}

df = pd.read_excel(INPUT_PATH, engine="openpyxl")
total = len(df)

col_name = "gender guess column"
if col_name not in df.columns:
    df[col_name] = np.nan

start_idx = max(0, START_ROW - 1)
end_idx   = min(total - 1, END_ROW - 1)

for i in range(start_idx, end_idx + 1):
    full_name = str(df.at[i, "Formatted Person Name"]) if pd.notna(df.at[i, "Formatted Person Name"]) else ""
    parts = full_name.strip().split()
    first = parts[0] if parts else ""
    last  = parts[-1] if len(parts) > 1 else ""
    code = np.nan
    gender = None

    if first:
        try:
            resp = requests.get(
                URL_TEMPLATE.format(first=first, last=last),
                headers=HEADERS,
                timeout=5
            )
            resp.raise_for_status()
            data = resp.json()
            gender = data.get("likelyGender")
            if gender == "male":
                code = 0
            elif gender == "female":
                code = 1
        except Exception as e:
            # uncomment to debug: print(f"Row {i+1} error: {e}")
            code = np.nan

    df.at[i, col_name] = code
    print(f"Row {i+1}/{total} → {gender}", flush=True)
    time.sleep(0.1)  # polite pause between requests

df.to_excel(INPUT_PATH, index=False, engine="openpyxl")
print(f"\nDone! Updated rows {START_ROW} to {END_ROW} and saved to {INPUT_PATH}")
