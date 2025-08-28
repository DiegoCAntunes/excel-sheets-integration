import glob
import pandas as pd
import uuid
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dotenv import load_dotenv
import os

load_dotenv()

# ---------------- CONFIG ----------------
PERSON_NAME = os.getenv("PERSON_NAME")
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_EXPENSE_NAME")
GOOGLE_TAB_NAME = os.getenv("GOOGLE_TAB_EXPENSE_NAME")
EXCEL_FOLDER = os.getenv("EXCEL_FOLDER_EXPENSE")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ---------------- GOOGLE AUTH ----------------
creds = None
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
        token.write(creds.to_json())

gc = gspread.authorize(creds)
sheet = gc.open(GOOGLE_SHEET_NAME).worksheet(GOOGLE_TAB_NAME)

# ---------------- FETCH EXISTING DATA ----------------
print("Fetching existing data from Google Sheet...")
existing = pd.DataFrame(sheet.get_all_records())

# ---------------- PROCESS EXCEL FILES ----------------
all_entries = []
excel_files = [f for f in glob.glob(os.path.join(EXCEL_FOLDER, "*.xlsx")) 
               if not os.path.basename(f).startswith("~$")]  # skip temp files
print(f"Found {len(excel_files)} Excel file(s) to process...")

for file in tqdm(excel_files, desc="Processing files"):
    # Read Excel with headers starting at row 9 (zero-indexed: 8)
    df = pd.read_excel(file, sheet_name="Sheet1", header=8)
    df.columns = df.columns.str.strip().str.upper()

    # Remove $ and commas from relevant columns before numeric conversion
    for col in ["MILEAGE TOTAL", "MEALS", "LODGING", "OTHER", "TOTAL"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"[$,]", "", regex=True)

    # Convert numeric columns to numeric and fill NaN
    df[["MILEAGE", "MILEAGE TOTAL", "MEALS", "LODGING", "OTHER", "TOTAL"]] = \
        df[["MILEAGE", "MILEAGE TOTAL", "MEALS", "LODGING", "OTHER", "TOTAL"]].apply(pd.to_numeric, errors="coerce").fillna(0)

    # Keep only rows where sum of numeric columns (excluding TOTAL) > 0
    df = df[df[["MILEAGE", "MILEAGE TOTAL", "MEALS", "LODGING", "OTHER"]].sum(axis=1) > 0]
    if df.empty:
        print(f"⚠️ No valid expense entries in {file}")
        continue

    # Build Google Sheet DataFrame
    df_master = pd.DataFrame({
        "ItemID": [uuid.uuid4().hex[:8] for _ in range(len(df))],
        "Date": pd.to_datetime(df["DATE"]).dt.strftime("%Y-%m-%d"),
        "JobNumber": df["JOB NUMBER"],
        "Description": df["DESCRIPTION"],
        "Name": PERSON_NAME,
        "Mileage": df["MILEAGE"],
        "Meals": df["MEALS"],
        "Lodging": df["LODGING"],
        "Other": df["OTHER"],
        "Total": df["TOTAL"],
        "ReceiptURL": ["" for _ in range(len(df))]
    })

    all_entries.append(df_master)
    print(f"✔ Processed {len(df_master)} entries from {os.path.basename(file)}")

# ---------------- UPLOAD TO GOOGLE SHEET ----------------
if all_entries:
    combined = pd.concat([existing] + all_entries, ignore_index=True)
    print("Clearing Google Sheet and uploading data...")
    sheet.clear()
    set_with_dataframe(sheet, combined)
    print(f"✅ Upload complete! Total rows in sheet: {len(combined)}")
else:
    print("⚠️ No valid expense entries found.")
