# Custom Bill Email Generator

This script processes utility bills from the `custom_bill/` folder and generates email drafts with a custom month.

## Purpose

Use this script when you need to:
- Process bills that are stored separately in the `custom_bill/` folder
- Specify a custom month for the email (different from the bill dates)
- Always use the "dual" email template (showing ENMAX + ATCO breakdown)

## Usage

### Run the script
```bash
python3 custom_bill_email.py
```

The script will prompt you for:
1. **Month number (1-12)**: Enter the month for the email
   - 1 = January
   - 2 = February
   - ...
   - 11 = November
   - 12 = December

2. **Test mode (y/n)**: Choose whether to preview or send
   - `y` = Test mode (preview email without sending)
   - `n` = Live mode (create draft in Yahoo Mail)

## Example Session

```
Custom Bill Email Generator
==================================================

Enter month number (1-12): 11
Run in test mode? (y/n): y

Starting email generation [TEST MODE] for November...
==================================================
```

## How It Works

1. **Reads PDFs**: Scans all PDF files in the `custom_bill/` folder
2. **Extracts Data**: Extracts house number, vendor, amount, and date from each PDF
3. **Creates Images**: Crops and saves images from the first page of each PDF to `bill_images/`
4. **Groups by House**: Groups bills by house number
5. **Generates Email**: Creates email draft with:
   - Subject: "{Month} utility bill" (using the custom month you specify)
   - Body: Uses the "dual" template showing ENMAX and ATCO breakdown
   - Attachments: Cropped images from the PDFs
   - Rent Date: First day of the month after the custom month

## Requirements

- PDFs must be in the `custom_bill/` folder
- PDFs must be valid ENMAX or ATCO utility bills
- Email credentials must be set in environment variables (unless using `--test` mode):
  - `YAHOO_USER`
  - `YAHOO_APP_PASSWORD`

## Output

### Test Mode
Prints email details to console:
- Subject line
- Recipient
- List of attachments
- Email body preview

### Live Mode
Creates draft email in Yahoo Mail drafts folder.

## Notes

- The script always uses the "dual" email template (ENMAX + ATCO breakdown)
- Images are cropped to remove footer information (configurable in Excel Config sheet)
- The custom month parameter controls the email subject and rent date calculation
- Original bill dates are preserved in the image filenames
- Reuses existing functions from the main pipeline for consistency and maintainability

