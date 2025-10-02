"""
Configuration management - reads settings from Excel Config sheet.
"""

import os
import pandas as pd
from pathlib import Path
from typing import Dict, Any, List

# Get the project root directory (rental folder)
PROJECT_ROOT = Path(__file__).parent
EXCEL_PATH = str(PROJECT_ROOT / "utility_bills.xlsx")

# Excel is the only source of truth for configuration

def _convert_value(key: str, value: Any) -> Any:
    """Convert Excel values to proper types."""
    if pd.isna(value):
        return None
    
    if key == 'image_bottom_crop_px':
        return int(value)
    elif key in ('move_processed_files', 'test_email_drafts'):
        if isinstance(value, bool):
            return value
        return str(value).lower() in ('true', '1', 'yes')
    
    return value

def load_config() -> Dict[str, Any]:
    """Load configuration from Excel Config sheet - Excel is the only source of truth."""
    try:
        config_df = pd.read_excel(EXCEL_PATH, sheet_name="Config")
        config = {}
        
        for _, row in config_df.iterrows():
            key = row['key']
            value = _convert_value(key, row['value'])
            if value is not None:
                config[key] = value
                
        return config
        
    except Exception as e:
        raise Exception(f"Could not load config from Excel Config sheet: {e}. Excel must be properly configured.")

def load_tenant_data() -> Dict[str, Dict[str, Any]]:
    """Load tenant data from Excel Tenants sheet - Excel is the only source of truth."""
    try:
        tenants_df = pd.read_excel(EXCEL_PATH, sheet_name="Tenants")
        tenant_dict = {}
        
        for _, row in tenants_df.iterrows():
            house_num = str(row['house_number'])
            tenant_dict[house_num] = {
                'tenant_name': row['tenant_name'],
                'email': row['email'],
                'base_rent': float(row['base_rent']),
                'utility_share_percent': int(row['utility_share_percent'])
            }
        
        return tenant_dict
        
    except Exception as e:
        raise Exception(f"Could not load tenant data from Excel Tenants sheet: {e}. Excel must be properly configured.")

def get_config(key: str) -> Any:
    """Get a single configuration value."""
    config = load_config()
    if key not in config:
        raise KeyError(f"Configuration key '{key}' not found")
    return config[key]

# Convenience functions for commonly used values
def get_excel_path() -> str:
    return EXCEL_PATH

def get_excel_data_sheet() -> str:
    return get_config('excel_data_sheet')

def get_raw_bills_folder() -> str:
    return get_config('raw_bills_folder')

def get_processed_bills_folder() -> str:
    return get_config('processed_bills_folder')

def get_images_folder() -> str:
    return get_config('images_folder')

def get_image_bottom_crop_px() -> int:
    return get_config('image_bottom_crop_px')

def get_atco_indicator() -> str:
    return get_config('atco_indicator')

def get_house_numbers() -> List[str]:
    """Get house numbers from the first column of the Tenants sheet."""
    try:
        tenants_df = pd.read_excel(EXCEL_PATH, sheet_name="Tenants")
        # Convert to string and return as list
        return [str(house) for house in tenants_df['house_number'].tolist()]
    except Exception as e:
        print(f"Error reading house numbers from Tenants sheet: {e}")
        return []

def get_move_processed_files() -> bool:
    return get_config('move_processed_files')

# Tenant data functions
def get_tenant_data(house_number: str = None) -> Dict[str, Any]:
    """Get tenant data for a specific house number, or all tenant data if no house_number provided."""
    tenant_data = load_tenant_data()
    if house_number:
        house_str = str(house_number)
        if house_str not in tenant_data:
            raise KeyError(f"No tenant data found for house number '{house_str}'")
        return tenant_data[house_str]
    return tenant_data

def get_all_house_numbers_with_tenants() -> List[str]:
    """Get list of all house numbers that have tenant data."""
    return list(load_tenant_data().keys())

# Month-end calculation utility
def bill_date_to_month_end(bill_date) -> str:
    """Convert bill date (datetime or string) to month-end date string."""
    import calendar
    from datetime import datetime
    
    if isinstance(bill_date, str):
        dt = datetime.strptime(bill_date, "%Y-%m-%d")
    else:
        dt = bill_date
    
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last_day).strftime("%Y-%m-%d")
