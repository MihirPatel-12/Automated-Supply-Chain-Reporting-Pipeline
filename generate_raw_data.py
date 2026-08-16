"""
generate_raw_data.py

Creates 5 realistic-but-messy "planner export" files, each with a different
real-world quirk, to simulate what a supply chain analyst actually receives
every week from different regional teams before any cleanup happens.

This script is NOT part of the pipeline itself -- it just builds the sample
raw inputs so the pipeline has something realistic to consolidate. Run once
before running pipeline.py.
"""
import random
from datetime import date, timedelta
import pandas as pd
import openpyxl
from openpyxl.styles import Font

random.seed(7)

RAW_DIR = "data/raw_exports"

skus = ["SKU001", "SKU002", "SKU003", "SKU004", "SKU005", "SKU006", "SKU007", "SKU008"]
start_date = date(2026, 8, 3)   # a Monday
n_days = 14                     # two weeks of data per region

def daily_rows(region_name):
    rows = []
    for d in range(n_days):
        the_date = start_date + timedelta(days=d)
        for sku in skus:
            base = {"SKU001": 12, "SKU002": 8, "SKU003": 20, "SKU004": 5,
                    "SKU005": 15, "SKU006": 9, "SKU007": 25, "SKU008": 6}[sku]
            units = max(0, round(random.gauss(base, base * 0.3)))
            rows.append({"SKU": sku, "Region": region_name, "Date": the_date, "Units Sold": units})
    return rows

# ---------------------------------------------------------------
# File 1: North -- clean .xlsx, this is what a "correct" export looks like
# ---------------------------------------------------------------
rows = daily_rows("North")
df = pd.DataFrame(rows)
df.to_excel(f"{RAW_DIR}/planner_export_north_week32.xlsx", index=False)

# ---------------------------------------------------------------
# File 2: South -- .csv with messy header names and US-style date strings
# ---------------------------------------------------------------
rows = daily_rows("South")
df = pd.DataFrame(rows)
df["Date"] = df["Date"].apply(lambda d: d.strftime("%m/%d/%Y"))  # MM/DD/YYYY text
df = df.rename(columns={"SKU": " SKU ", "Region": "region ", "Date": " Date", "Units Sold": "Units_Sold"})
df.to_csv(f"{RAW_DIR}/planner_export_south_week32.csv", index=False)

# ---------------------------------------------------------------
# File 3: East -- .csv with 2 metadata rows above the real header, plus
# one accidental duplicate row (someone exported twice into the same file)
# ---------------------------------------------------------------
rows = daily_rows("East")
df = pd.DataFrame(rows)
df["Date"] = df["Date"].astype(str)
dup_row = df.iloc[10:11]
df = pd.concat([df, dup_row], ignore_index=True)  # duplicate row injected
with open(f"{RAW_DIR}/planner_export_east_week32.csv", "w") as f:
    f.write("Weekly Planner Export\n")
    f.write("Generated: 2026-08-17 by E. Ramirez\n")
    df.to_csv(f, index=False)

# ---------------------------------------------------------------
# File 4: West -- .xlsx with some missing Units Sold values, one fully
# blank row, and Date stored as text instead of real dates
# ---------------------------------------------------------------
rows = daily_rows("West")
df = pd.DataFrame(rows)
df["Date"] = df["Date"].astype(str)
missing_idx = random.sample(range(len(df)), 5)
for idx in missing_idx:
    df.loc[idx, "Units Sold"] = None
wb = openpyxl.Workbook()
ws = wb.active
ws.append(list(df.columns))
for cell in ws[1]:
    cell.font = Font(bold=True)
blank_row_after = 6
for i, row in df.iterrows():
    ws.append(list(row))
    if i == blank_row_after:
        ws.append([None, None, None, None])  # stray blank row
wb.save(f"{RAW_DIR}/planner_export_west_week32.xlsx")

# ---------------------------------------------------------------
# File 5: Central -- .csv with columns in a different order, inconsistent
# region-name casing, and one data-entry error (negative units)
# ---------------------------------------------------------------
rows = daily_rows("Central")
df = pd.DataFrame(rows)
df["Region"] = df["Region"].apply(lambda r: r.lower() if random.random() < 0.4 else r)
df.loc[3, "Units Sold"] = -4  # data entry error
df = df[["Date", "SKU", "Region", "Units Sold"]]  # reordered columns
df.to_csv(f"{RAW_DIR}/planner_export_central_week32.csv", index=False)

print("5 raw planner export files created in", RAW_DIR)
