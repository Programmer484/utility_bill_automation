"""
Configuration management - reads settings from Excel Config sheet.
"""

import os
import pandas as pd
from pathlib import Path
from typing import Dict, Any, List

# Get the project root directory (rental folder)
# This file is now in the root directory
PROJECT_ROOT = Path(__file__).parent

# Excel file path (still hardcoded as it's needed to read config)
EXCEL_PATH = str(PROJECT_ROOT / "utility_bills.xlsx")

# Cache for config values
_config_cache: Dict[str, Any] = {}
_tenant_cache: Dict[str, Dict[str, Any]] = {}

def load_config() -> Dict[str, Any]:
    """Load configuration from Excel Config sheet."""
    global _config_cache
    
    if _config_cache:
        return _config_cache
    
    try:
        config_df = pd.read_excel(EXCEL_PATH, sheet_name="Config")
        
        # Convert to dictionary
        config_dict = {}
        for _, row in config_df.iterrows():
            key = row['key']
            value = row['value']
            
            # Type conversion based on key
            if key == 'image_bottom_crop_px':
                value = int(value)
            elif key == 'move_processed_files':
                # Handle both boolean objects and string values
                if isinstance(value, bool):
                    value = value  # Keep boolean as-is
                else:
                    value = str(value).lower() in ('true', '1', 'yes')
            elif key == 'house_numbers':
                value = [h.strip() for h in str(value).split(',')]
            
            config_dict[key] = value
        
        _config_cache = config_dict
        return config_dict
        
    except Exception as e:
        print(f"Warning: Could not load config from Excel: {e}")
        print("Using fallback hardcoded values...")
        return _get_fallback_config()

def _get_fallback_config() -> Dict[str, Any]:
    """Fallback configuration if Excel config cannot be loaded."""
    return {
        'excel_data_sheet': 'Data',
        'raw_bills_folder': str(PROJECT_ROOT / 'bills'),
        'processed_bills_folder': str(PROJECT_ROOT / 'bills_processed'),
        'images_folder': str(PROJECT_ROOT / 'bill_images'),
        'image_bottom_crop_px': 450,
        'atco_indicator': 'statements',
        'house_numbers': ['819', '1705', '1707', '1712'],
        'move_processed_files': False,
    }

def get_config(key: str) -> Any:
    """Get a single configuration value."""
    config = load_config()
    if key not in config:
        raise KeyError(f"Configuration key '{key}' not found")
    return config[key]

def load_tenant_data() -> Dict[str, Dict[str, Any]]:
    """Load tenant data from Excel Tenants sheet."""
    global _tenant_cache
    
    if _tenant_cache:
        return _tenant_cache
    
    try:
        tenants_df = pd.read_excel(EXCEL_PATH, sheet_name="Tenants")
        
        # Convert to dictionary with house_number as key
        tenant_dict = {}
        for _, row in tenants_df.iterrows():
            house_num = str(row['house_number'])  # Ensure string key
            tenant_dict[house_num] = {
                'tenant_name': row['tenant_name'],
                'email': row['email'],
                'base_rent': float(row['base_rent']),
                'utility_share_percent': int(row['utility_share_percent'])
            }
        
        _tenant_cache = tenant_dict
        return tenant_dict
        
    except Exception as e:
        print(f"Warning: Could not load tenant data from Excel: {e}")
        print("Using fallback hardcoded tenant data...")
        return _get_fallback_tenant_data()

def _get_fallback_tenant_data() -> Dict[str, Dict[str, Any]]:
    """Fallback tenant data if Excel tenants sheet cannot be loaded."""
    return {
        "819": {
            "tenant_name": "Sarah Johnson",
            "email": "sarah.johnson@email.com",
            "base_rent": 1200.0,
            "utility_share_percent": 60
        },
        "1705": {
            "tenant_name": "Mike Chen", 
            "email": "mike.chen@email.com",
            "base_rent": 1150.0,
            "utility_share_percent": 60
        },
        "1707": {
            "tenant_name": "Emma Rodriguez",
            "email": "emma.rodriguez@email.com", 
            "base_rent": 1200.0,
            "utility_share_percent": 60
        },
        "1712": {
            "tenant_name": "James Wilson",
            "email": "james.wilson@email.com",
            "base_rent": 1100.0,
            "utility_share_percent": 60
        }
    }

def reload_config():
    """Clear cache and reload configuration."""
    global _config_cache, _tenant_cache
    _config_cache = {}
    _tenant_cache = {}
    return load_config()

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
    return get_config('house_numbers')

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

# Month-end calculation (kept here as it's a utility function)
def bill_date_to_month_end(bill_date) -> str:
    """Convert bill date (datetime or string) to month-end date string."""
    import calendar
    from datetime import datetime
    
    # Handle both datetime objects and strings
    if isinstance(bill_date, str):
        dt = datetime.strptime(bill_date, "%Y-%m-%d")
    else:
        dt = bill_date
    
    last_day = calendar.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last_day).strftime("%Y-%m-%d")
