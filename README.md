# Expense Reports Sync 📝

Automate the process of importing data from Excel files to Google Sheets for tracking and integration with AppSheet.

---

## Features

- Automatically reads Excel files.
- Generates unique IDs for each entry.
- Uploads entries to a Google Sheet, merging with existing data.
- Skips temporary Excel files (e.g., `~$` files).
- Supports multiple employees with separate folders.

---

## Prerequisites

- Python 3.10+
- Google Sheets API enabled and credentials JSON (`credentials.json`)
- Installed Python packages (see `requirements.txt`):

  pip install -r requirements.txt

---

## Setup

1. Clone the repository:

   git clone https://github.com/DiegoCAntunes/excel-sheets-integration
   cd expense-reports-sync

2. Create a `.env` file in the project root with your settings:

   PERSON_NAME="PersonName for column"
   GOOGLE_SHEET_NAME="SheetName"
   GOOGLE_TAB_NAME="tabName"
   EXCEL_FOLDER="path\\to\\excel\files"

> Use double backslashes `\\` in Windows paths.

---

## Usage

Run the script:

    python expense-reports.py

The script will:

- Fetch existing data from the Google Sheet.
- Process all Excel files in the folder.
- Merge new entries with existing data.
- Upload the combined dataset back to Google Sheets.

---

## Notes

- Excel files must follow the standard format for this automation to work.
- `.env` keeps sensitive data out of the codebase.
- Skip temporary Excel files starting with `~$`.
- Adjust header row in the script if your Excel files have a different starting row.
