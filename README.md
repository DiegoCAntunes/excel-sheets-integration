# Expense Reports Sync 📝

Automate the process of uploading employee expense reports from Excel files to Google Sheets for tracking and integration with AppSheet.

---

## Features

- Automatically reads Excel expense reports from a local folder.
- Processes mileage, meals, lodging, other expenses, and totals.
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

   git clone https://github.com/yourusername/expense-reports-sync.git
   cd expense-reports-sync

2. Create a `.env` file in the project root with your settings:

   PERSON_NAME="Diego Cazetta Antunes"
   GOOGLE_SHEET_NAME="Copy of ExpenseDB_2025-08-08"
   GOOGLE_TAB_NAME="ExpenseEntries"
   EXCEL_FOLDER="G:\\Documents\\DITANU INC\\Projects\\250001 - Appsheet\\Expense Reports\\Diego Cazetta Antunes"

> Use double backslashes `\\` in Windows paths.

3. Ensure your Excel files start their table headers on the correct row (default: row 9).

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
