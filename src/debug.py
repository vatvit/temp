#!/usr/bin/env python3
"""Debug: Compare tovar:268 vs база ЗІП Samsung:238"""
import pandas as pd
from pathlib import Path

BASE_PATH = Path("/app")

# Load tovar row 268
df_tovar = pd.read_excel(BASE_PATH / "tovar2025_PRN.xlsx",
                         sheet_name="Январь продажи", header=None)
row_tovar = df_tovar.iloc[267]

# Load база ЗІП Samsung row 238
df_база = pd.read_excel(BASE_PATH / "База товара_PRN.xls",
                        sheet_name="ЗИП Samsung", header=None)
row_база = df_база.iloc[237]

print("=== tovar:268 vs база ЗІП Samsung:238 ===\n")
print(f"{'Col':<4} {'tovar:268':<35} {'база:238':<35} {'Match'}")
print("-" * 80)

cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
mismatches = []
for i in range(8):
    t = row_tovar[i]
    b = row_база[i]
    match = "✓" if t == b or (pd.isna(t) and pd.isna(b)) else "✗"
    t_str = str(t)[:33] if not pd.isna(t) else ""
    b_str = str(b)[:33] if not pd.isna(b) else ""
    print(f"{cols[i]:<4} {t_str:<35} {b_str:<35} {match}")
    if match == "✗":
        mismatches.append(cols[i])

print("-" * 80)
print(f"V    {str(row_tovar[21]):<35} {str(row_база[21]):<35} {'✓' if row_tovar[21] == row_база[21] else '✗'}")

print(f"\nMismatches in A-H: {mismatches}")
print(f"Current tolerance allows: 1 mismatch")
print(f"This pair has: {len(mismatches)} mismatches -> NO MATCH")
