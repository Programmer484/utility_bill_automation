# Utility Bill Processing Workflow

This document describes the complete workflow for processing utility bills, from PDF extraction to email draft generation.

## Overview

The system processes utility bills (ENMAX and ATCO) for multiple rental properties, extracts key information, stores data in Excel, and generates email drafts for tenants with calculated rent amounts.

## 1. Initial Setup and Configuration

### Directory Creation
Creates three folders in this project folder:
- **bills** (put PDFs here to process)
- **bills_processed** (processed PDFs move here)
- **bill_images** (bill images saved here)

Folder paths come from your Excel Config sheet. Use relative paths ("bills") to create folders in this project folder, or absolute paths ("C:/Users/username/my_bills").

Creates folders automatically if missing. Skips if they already exist.

### Configuration Loading
Configuration is loaded from the Excel file's "Config" sheet, including:
- Folder paths for all operations
- House numbers to recognize in bills
- Vendor identification keywords
- Image processing settings
- File movement preferences

If the Excel config cannot be read, the system falls back to hardcoded default values.

### Tenant Data Loading
Tenant information is loaded from the Excel file's "Tenants" sheet, containing:
- Tenant names and email addresses
- Base rent amounts for each property
- Utility share percentages (what portion of utilities tenants pay)

## 2. PDF Discovery and Processing

### File Discovery
The system scans the raw bills folder for PDF files, ignoring:
- Hidden files (starting with dots)
- Non-PDF files
- Files in subdirectories

### Vendor Detection and Routing
For each PDF file found:
- PDF content is analyzed to determine vendor using text patterns:
  - **ENMAX**: Detected by presence of "ENMAX.COM" in the PDF text
  - **ATCO**: Detected by presence of "STATEMENT DATE:" in the PDF text
- If neither pattern is found, an error is thrown (no default assumption)
- Appropriate extraction function is selected based on detected vendor

## 3. Data Extraction from PDFs

### ATCO Bill Extraction
From ATCO utility statements:
- **House Number**: Searches for configured house numbers at the start of address lines
- **Bill Amount**: Looks for "TOTAL AMOUNT DUE" or "Amount Due" followed by dollar amounts
- **Bill Date**: Extracts from "Statement Date: [Month] [Day], [Year]" format
- **Vendor**: Always set to "ATCO"

### ENMAX Bill Extraction  
From ENMAX utility bills:
- **House Number**: Searches in "SERVICE ADDRESS" sections for configured house numbers
- **Bill Amount**: Extracts from "PreAuthorizedAmount" or "TotalCurrentCharges" fields
- **Bill Date**: Parses from "CurrentBillDate: [YYYY][MonthName][Day]" format
- **Vendor**: Always set to "ENMAX"

### Data Cleanup
All extracted information is cleaned up and standardized:
- Company names are made consistent (all caps, extra spaces removed)
- Bill amounts are converted to numbers for calculations
- Dates are put in a standard format (YYYY-MM-DD, like 2025-08-26)
- House numbers are converted to whole numbers when possible

## 4. File Processing

### Image Creation
- PDF first page → PNG image (bottom cropped)
- Naming: `[house]_[date]_[vendor].png`

### File Organization
- **Parameter**: `rename_files = TRUE/FALSE` in Excel config
- **If TRUE**: Move PDFs to processed folder and are renamed with format `[house] [Month Day Year] [vendor].pdf`
- **If FALSE**: PDFs remain in original folder

## 5. Email Generation

### Core Logic
- Groups bills by house and month (bill dates → month-end keys)
- Multiple vendors per house are summed together
- **Example**: ENMAX ($180) + ATCO ($218) = $398 total for house 1705

### Rent Calculation
**Final amount = Base rent + (Utility share % × Total utilities)**

### Test Mode
- **Excel config**: `test_email_drafts = TRUE` → prints email details
- **Production**: `test_email_drafts = FALSE` → saves drafts to Yahoo

## 6. Data Storage
Bills saved to Excel after email generation for record-keeping.

## Key Rules
- **House number + bill date required** (bills skipped if missing)
- **Errors don't stop processing** (failed files logged, others continue)
- **No PDFs = No emails** (only generates from current run)


