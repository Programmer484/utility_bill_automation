#!/usr/bin/env python3
"""
Debug script to check Excel data types and formatting issues.
"""

import sys
import os
import pandas as pd

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
# Add parent directory to path to import config from root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from config import get_excel_path, get_excel_data_sheet

def debug_excel_data():
    print("=== EXCEL DATA DEBUG ===")
    excel_path = get_excel_path()
    data_sheet = get_excel_data_sheet()
    print(f"Reading from: {excel_path}")
    print(f"Sheet: {data_sheet}")
    print()
    
    try:
        df = pd.read_excel(excel_path, sheet_name=data_sheet)
        
        print("RAW DATA INFO:")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print()
        
        print("DATA TYPES:")
        for col in df.columns:
            print(f"  {col}: {df[col].dtype}")
        print()
        
        print("SAMPLE DATA (first 3 rows):")
        print(df.head(3))
        print()
        
        # Check bill_date column specifically
        if 'bill_date' in df.columns:
            print("BILL_DATE ANALYSIS:")
            print(f"  Type: {df['bill_date'].dtype}")
            print(f"  Sample values:")
            for i, val in enumerate(df['bill_date'].head(5)):
                print(f"    Row {i+1}: {repr(val)} (type: {type(val)})")
            print()
        
        # Check vendor column
        if 'vendor' in df.columns:
            print("VENDOR ANALYSIS:")
            print(f"  Type: {df['vendor'].dtype}")
            print(f"  Unique values: {df['vendor'].unique()}")
            print(f"  Value counts:")
            print(df['vendor'].value_counts())
            print()
        
        # Check house_number column
        if 'house_number' in df.columns:
            print("HOUSE_NUMBER ANALYSIS:")
            print(f"  Type: {df['house_number'].dtype}")
            print(f"  Unique values: {df['house_number'].unique()}")
            print()
            
        # Try the month calculation that email_drafts uses
        print("MONTH CALCULATION TEST:")
        df_test = df.copy()
        df_test.columns = [c.strip().lower() for c in df_test.columns]
        
        if 'bill_date' in df_test.columns:
            print("Converting bill_date to datetime...")
            df_test["bill_date"] = pd.to_datetime(df_test["bill_date"], errors="coerce")
            print(f"After conversion - null dates: {df_test['bill_date'].isna().sum()}")
            
            if not df_test['bill_date'].isna().all():
                from config import bill_date_to_month_end
                print("Testing month-end conversion...")
                for i, date_val in enumerate(df_test['bill_date'].head(3)):
                    if pd.notna(date_val):
                        iso_str = date_val.strftime("%Y-%m-%d")
                        month_end = bill_date_to_month_end(iso_str)
                        print(f"  {date_val} -> {iso_str} -> {month_end}")
        
    except Exception as e:
        print(f"ERROR reading Excel: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_excel_data()
