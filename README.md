# PYTHON - 1. PROJECT NAPS Clean settlement file
import pandas as pd
import numpy as np
import os
import time
from datetime import datetime

start = time.time()

# === Parameters ===
input_file    = r"C:\Users\OMODELEC\Downloads\2020mar-apr.csv"   # ← update path
output_folder = r"C:\Users\OMODELEC\OneDrive - Access Bank PLC\Documents\NAPSS\Python outputs"
start_date    = "2020-04-01"
end_date      = "2020-04-30"

# === 1. Load and preprocess ===
df = pd.read_csv(input_file, dtype=str, low_memory=False)
df['TRANSACTIONAMOUNT'] = pd.to_numeric(df['TRANSACTIONAMOUNT'], errors='coerce')
df['TRANSACTIONAMOUNT'] *= np.where(df['DRCR'] == 'D', -1, 1)

# === 2. Parse datetime and filter ===
raw_dt = pd.to_datetime(df['SETTLEMENTSESSION'], format='%y%m%d%H%M', errors='coerce')
mask = (raw_dt >= pd.to_datetime(start_date)) & (raw_dt <= pd.to_datetime(end_date))
df = df.loc[mask].copy()
df['Transaction_Date'] = raw_dt.dt.strftime('%m/%d/%Y')
df['Session_Time']     = raw_dt.dt.strftime('%H:%M')

# === 3. Sort descending ===
df.sort_values(['Transaction_Date','Session_Time','TRANSACTIONAMOUNT'],
               ascending=[False,False,False], inplace=True)

# === 4. Pivot summary ===
summary = df.pivot_table(
    index='TRANSACTIONREF', 
    columns='DRCR',
    values='TRANSACTIONAMOUNT', 
    aggfunc='sum', 
    fill_value=0
).reset_index()

# Net as D + C
summary['NET'] = summary.get('D', 0) + summary.get('C', 0)

# === 5. Sanitize control chars ===
df.replace(r'[\x00-\x1F]', '', regex=True, inplace=True)
summary.replace(r'[\x00-\x1F]', '', regex=True, inplace=True)

# === 6. Write one workbook, Summary first ===
ts = datetime.now().strftime('%Y%m%d_%H%M%S')
months = raw_dt.dt.strftime('%b')[mask].unique()
month_str = "-".join(sorted(months))
filename = f"{ts}_NAPs Settlement_{month_str}.xlsx"
out_path = os.path.join(output_folder, filename)

with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    summary.to_excel(writer, sheet_name='Summary', index=False)
    df.to_excel(writer, sheet_name='Detail', index=False)

elapsed = time.time() - start
print(f"✅ Completed in {elapsed:.2f}s — saved to {out_path}")

