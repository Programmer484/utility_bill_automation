import os
import pandas as pd
import sys
from pathlib import Path
from typing import Optional, Iterable, List, Dict

# Add parent directory to path to import config
sys.path.append(str(Path(__file__).parent.parent))
from config import bill_date_to_month_end, get_tenant_data as get_config_tenant_data

# Tenant data is now loaded from Excel config - see config.py

DEFAULT_TENANT_DATA = {
    "tenant_name": "Tenant",
    "email": "tenant@example.com", 
    "base_rent": 1200,
    "utility_share_percent": 60
}

def get_tenant_data(house_number: str) -> Dict:
    """Get tenant data for a house number, with defaults if not found."""
    try:
        return get_config_tenant_data(str(house_number))
    except KeyError:
        # Fallback to default if house number not found in config
        return {
            **DEFAULT_TENANT_DATA,
            "tenant_name": f"Tenant {house_number}"
        }

def append_rows_to_excel(excel_path: str, rows: List[Dict], sheet_name: str) -> None:
    """
    Append rows to an Excel sheet; create file/sheet if needed.
    Keeps consistent column order for known fields and appends any extras at the end.
    Automatically adds tenant data (name, base rent, utility share) based on house_number.
    Skips duplicate rows based on filename.
    """
    if not rows:
        return

    # Enrich rows with tenant data (only tenant_name)
    enriched_rows = []
    for row in rows:
        enriched_row = row.copy()
        if "house_number" in row:
            tenant_data = get_tenant_data(str(row["house_number"]))
            enriched_row.update({
                "tenant_name": tenant_data["tenant_name"]
            })
        enriched_rows.append(enriched_row)

    # Only include the 6 specified columns
    core_columns = ["file", "house_number", "tenant_name", "bill_amount", "bill_date", "vendor"]
    df_new = pd.DataFrame(enriched_rows)

    extra_columns = [c for c in df_new.columns if c not in core_columns]
    df_new = df_new[core_columns + extra_columns] if extra_columns else df_new[core_columns]

    if os.path.exists(excel_path):
        try:
            existing = pd.read_excel(excel_path, sheet_name=sheet_name)
            existing.columns = [c.strip().lower() for c in existing.columns]
            df_new.columns = [c.strip().lower() for c in df_new.columns]
            
            # Ensure consistent date formatting as datetime objects
            if 'bill_date' in existing.columns:
                existing['bill_date'] = pd.to_datetime(existing['bill_date'], errors='coerce')
            if 'bill_date' in df_new.columns:
                df_new['bill_date'] = pd.to_datetime(df_new['bill_date'], errors='coerce')
            
            # Filter out duplicates based on filename
            if not existing.empty and 'file' in existing.columns:
                existing_files = set(existing['file'].astype(str))
                new_rows_mask = ~df_new['file'].astype(str).isin(existing_files)
                df_new_filtered = df_new[new_rows_mask]
                
                if df_new_filtered.empty:
                    print("No new rows to add - all files already exist in Excel")
                    return
                else:
                    skipped_count = len(df_new) - len(df_new_filtered)
                    if skipped_count > 0:
                        print(f"Skipped {skipped_count} duplicate files, adding {len(df_new_filtered)} new rows")
                
                combined = pd.concat([existing, df_new_filtered], ignore_index=True)
            else:
                combined = pd.concat([existing, df_new], ignore_index=True)
        except ValueError:
            combined = df_new
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_new.to_excel(writer, index=False, sheet_name=sheet_name)


def latest_month_totals(
    excel_path: str,
    sheet_name: str,
    houses: Optional[Iterable] = None
) -> pd.DataFrame:
    """
    Returns one row per house with:
      house_number | latest_month | total | ENMAX | ATCO | (other vendors if present)

    'latest_month' is a real date (month-end). Totals sum all vendors that fall in that month.
    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"file", "house_number", "bill_amount", "bill_date", "vendor"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")

    # Clean and normalize
    df["bill_amount"] = pd.to_numeric(df["bill_amount"], errors="coerce")
    df["bill_date"]   = pd.to_datetime(df["bill_date"], errors="coerce")
    df["vendor"]      = df["vendor"].astype(str).str.strip().str.upper()

    if houses:
        df = df[df["house_number"].isin(houses)]

    # Compute month-end for each bill_date
    df["month"] = df["bill_date"].apply(bill_date_to_month_end)
    df["month"] = pd.to_datetime(df["month"])

    # Find latest month per house
    last_month = df.groupby("house_number")["month"].max().rename("latest_month")
    df_latest = df.merge(last_month, on="house_number", how="inner")
    df_latest = df_latest[df_latest["month"] == df_latest["latest_month"]]

    # Vendor split
    vendor_split = (
        df_latest.pivot_table(
            index=["house_number", "latest_month"],
            columns="vendor",
            values="bill_amount",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    # Month totals
    totals = (
        df_latest.groupby(["house_number", "latest_month"])["bill_amount"]
        .sum()
        .rename("total")
        .reset_index()
    )

    result = vendor_split.merge(totals, on=["house_number", "latest_month"], how="left")
    return result.sort_values(["house_number", "latest_month"]).reset_index(drop=True)


