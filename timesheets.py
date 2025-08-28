import os
import glob
import pandas as pd
import uuid
import numpy as np
import warnings
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Ignore openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

load_dotenv()

# ---------------- CONFIG ----------------
PERSON_NAME = os.getenv("PERSON_NAME")
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_TIMESHEET_NAME")
GOOGLE_TAB_NAME = os.getenv("GOOGLE_TAB_TIMESHEET_NAME")
EXCEL_FOLDER = os.getenv("EXCEL_FOLDER_TIMESHEET")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ---------------- HELPER FUNCTIONS ----------------
def excel_duration_to_str(x):
    if pd.isna(x):
        return "0:00:00"
    if isinstance(x, pd.Timedelta):
        total_seconds = int(x.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours}:{minutes:02}:{seconds:02}"
    if isinstance(x, pd.Timestamp):
        return f"{x.hour}:{x.minute:02}:{x.second:02}"
    return str(x)

def sanitize_numeric(x):
    if pd.isna(x) or np.isinf(x):
        return 0
    return x

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
               if not os.path.basename(f).startswith("~$")]
print(f"Found {len(excel_files)} Excel file(s) to process...")

for file in tqdm(excel_files, desc="Processing files"):
    df = pd.read_excel(file, sheet_name="Pay Period")

    # Keep only rows with valid dates
    df = df[df["Date"].apply(lambda x: pd.to_datetime(x, errors="coerce")).notnull()]

    # Fill missing Mileage
    if "Mileage" not in df.columns:
        df["Mileage"] = 0

    # Build master DataFrame
    df_master = pd.DataFrame({
        "ID": [uuid.uuid4().hex[:8] for _ in range(len(df))],
        "Date": pd.to_datetime(df["Date"]).dt.strftime("%m/%d/%Y"),
        "Name": PERSON_NAME,
        "JobNumber": df["Project"],
        "Type": df["Type"],
        "Assembly": df["Assembly"],
        "Code": df["Code"],
        "Description": df["Description"],
        "Time In": df["Time In"].apply(excel_duration_to_str),
        "Time Out": df["Time Out"].apply(excel_duration_to_str),
        "Total": df["Total"].apply(excel_duration_to_str),
        "Departure Type": ["" for _ in range(len(df))],
        "Departure Custom": ["" for _ in range(len(df))],
        "Destination Type": ["" for _ in range(len(df))],
        "Destination Custom": ["" for _ in range(len(df))],
        "Driver": df["Mileage"].apply(lambda x: False if pd.isna(x) or str(x).strip()=="" else True),
        "Distance": df["Mileage"].apply(sanitize_numeric)
    })

    all_entries.append(df_master)
    print(f"✔ Processed {len(df_master)} rows from {os.path.basename(file)}")

# ---------------- UPLOAD TO GOOGLE SHEET ----------------
if all_entries:
    print("Merging new entries with existing Google Sheet data...")
    new_entries = pd.concat(all_entries, ignore_index=True)
    combined = pd.concat([existing, new_entries], ignore_index=True)

    # Sanitize before upload
    combined = combined.replace([np.inf, -np.inf], 0)

    print("Clearing Google Sheet and uploading data...")
    sheet.clear()
    set_with_dataframe(sheet, combined)

    print(f"✅ Upload complete! Total rows in sheet: {len(combined)}")
else:
    print("⚠️ No valid Excel files found.")
