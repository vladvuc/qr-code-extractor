#!/usr/bin/env python3
"""Quick verification script to check QR code analysis results"""

import pandas as pd

# Read the analyzed file
df = pd.read_excel("Sheet3_QR_Analyzed.xlsx")

print("=" * 70)
print("VERIFICATION REPORT - Sheet3_QR_Analyzed.xlsx")
print("=" * 70)

print(f"\nTotal rows: {len(df)}")
print(f"Columns: {', '.join(df.columns.tolist())}")

print("\n" + "=" * 70)
print("QR_CODE COLUMN ANALYSIS")
print("=" * 70)

# Count different types of results
qr_found = df[df['QR_CODE'].str.contains('http', na=False, case=False) |
              df['QR_CODE'].str.match(r'^[^N]', na=False)]
not_found = df[df['QR_CODE'] == 'NOT_FOUND']
errors = df[df['QR_CODE'].str.contains('ERROR', na=False)]
multiple_qr = df[df['QR_CODE'].str.contains('QR codes found', na=False)]

print(f"\nQR codes found (with data): {len(qr_found)}")
print(f"NOT_FOUND: {len(not_found)}")
print(f"Errors: {len(errors)}")
print(f"Multiple QR codes: {len(multiple_qr)}")

print("\n" + "=" * 70)
print("SAMPLE QR CODES FOUND (First 10)")
print("=" * 70)

found_samples = df[df['QR_CODE'].str.contains('http', na=False, case=False) |
                   (df['QR_CODE'].notna() & (df['QR_CODE'] != 'NOT_FOUND') &
                    ~df['QR_CODE'].str.contains('ERROR', na=False))].head(10)

if len(found_samples) > 0:
    for idx, row in found_samples.iterrows():
        qr_data = row['QR_CODE']
        if len(str(qr_data)) > 80:
            qr_data = str(qr_data)[:77] + "..."
        print(f"Row {idx+2}: {qr_data}")
else:
    print("No QR codes found in the dataset")

print("\n" + "=" * 70)
print("SAMPLE NOT_FOUND ENTRIES (First 5)")
print("=" * 70)

not_found_samples = df[df['QR_CODE'] == 'NOT_FOUND'].head(5)
for idx, row in not_found_samples.iterrows():
    print(f"Row {idx+2}: NOT_FOUND (url: {row['url'][:60]}...)")

print("\n" + "=" * 70)
print("DATA INTEGRITY CHECK")
print("=" * 70)

# Check if all rows have a QR_CODE value
null_qr = df['QR_CODE'].isna().sum()
empty_qr = (df['QR_CODE'] == '').sum()

print(f"Null QR_CODE values: {null_qr}")
print(f"Empty QR_CODE values: {empty_qr}")
print(f"All rows processed: {'YES' if null_qr == 0 and empty_qr == 0 else 'NO'}")

# Verify original columns are preserved
original_cols = ['ref', 'url', 'Photo Okret Sticker', 'QR_CODE']
preserved = all(col in df.columns for col in original_cols)
print(f"Original columns preserved: {'YES' if preserved else 'NO'}")

print("\n" + "=" * 70)
