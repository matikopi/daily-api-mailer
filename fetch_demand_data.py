import datetime
import json
from typing import List, Dict

import pandas as pd
import requests

# -------- Configuration -------- #
API_KEY = "65d9cbf274e3409494004763c8a023a2"

# Try both domains - .net and .co.il
BASE_URLS = [
    "https://apim-api.noga-iso.co.il/DEMAND/DEMANDAPI/v1"
]

DATE_FMT = "%d-%m-%Y"
OUTPUT_EXCEL = "demand_data.xlsx"

# -------------------------------- #

def _call_api(from_date: str, to_date: str) -> List[Dict]:
    """Send a POST request to the DEMAND API and return the parsed JSON response."""

    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "Content-Type": "application/json",
    }

    payload = {"fromDate": from_date, "toDate": to_date}

    # Try both base URLs
    for base_url in BASE_URLS:
        try:
            print(f"Trying {base_url}...")
            response = requests.post(base_url, headers=headers, json=payload, timeout=60)
            response.raise_for_status()
            print(f"Success with {base_url}")
            return response.json()
        except Exception as e:
            print(f"Failed with {base_url}: {e}")
            continue
    
    raise Exception("All API endpoints failed")


def _flatten_response(raw: List[Dict]) -> pd.DataFrame:
    """Transform the nested API response into a flat pandas DataFrame."""

    # Handle both possible response structures
    data_list: List[Dict] = []
    if isinstance(raw, dict) and "energy" in raw:
        data_list = raw["energy"] or []
    elif isinstance(raw, dict) and "demand" in raw:
        data_list = raw["demand"] or []
    elif isinstance(raw, list):
        data_list = raw
    else:
        raise ValueError("Unexpected API response structure")

    records: List[Dict] = []
    for day_entry in data_list:
        date_str = day_entry.get("date")
        # Handle different possible data field names
        data_field = day_entry.get("demandData") or day_entry.get("data") or []
        
        for record in data_field:
            record_flat = {"date": date_str}
            record_flat.update(record)
            records.append(record_flat)

    return pd.DataFrame(records)


def fetch_demand_data(from_date: datetime.date, to_date: datetime.date) -> pd.DataFrame:
    """Fetch DEMAND data between two dates and return as DataFrame."""

    raw_data = _call_api(from_date.strftime(DATE_FMT), to_date.strftime(DATE_FMT))
    return _flatten_response(raw_data)


def main() -> None:
    """Entry point for the script."""

    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    # 1. Fetch most recent full day (yesterday)
    print(f"Fetching DEMAND data for {yesterday.strftime(DATE_FMT)} …")
    recent_df = fetch_demand_data(yesterday, yesterday)

    # 2. Fetch data from 01-Jan-2021 to yesterday
    start_date = datetime.date(2021, 1, 1)
    print(f"Fetching DEMAND history from {start_date.strftime(DATE_FMT)} to {yesterday.strftime(DATE_FMT)} …")
    history_df = fetch_demand_data(start_date, yesterday)

    # 3. Write both DataFrames to an Excel workbook
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        recent_df.to_excel(writer, sheet_name="Most Recent Day", index=False)
        history_df.to_excel(writer, sheet_name="Since 2021-01-01", index=False)

    print(f"Done! DEMAND data written to {OUTPUT_EXCEL}.")


if __name__ == "__main__":
    main() 