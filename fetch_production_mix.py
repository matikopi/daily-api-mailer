import datetime
import json
from typing import List, Dict

import pandas as pd
import requests

# -------- Configuration -------- #
# Replace this with your actual subscription key
API_KEY = "4851c5dc2cad4fdc996da5a347965c57"

# Base endpoint for the Noga ISO API gateway
BASE_URL = "https://apim-api.noga-iso.co.il/productionmix/PRODMIXAPI/v1"

# Date format expected by the API (example given: "26-05-2024")
DATE_FMT = "%d-%m-%Y"

# Output Excel file name
OUTPUT_EXCEL = "production_mix.xlsx"

# -------------------------------- #

def _call_api(from_date: str, to_date: str) -> List[Dict]:
    """Send a POST request to the Production Mix API and return the parsed JSON response.

    Parameters
    ----------
    from_date : str
        Date string in dd-mm-YYYY format (inclusive).
    to_date : str
        Date string in dd-mm-YYYY format (inclusive).

    Returns
    -------
    List[Dict]
        Parsed JSON response containing production mix data.
    """

    headers = {
        "Ocp-Apim-Subscription-Key": API_KEY,
        "Content-Type": "application/json",
    }

    payload = {"fromDate": from_date, "toDate": to_date}

    response = requests.post(BASE_URL, headers=headers, json=payload, timeout=60)
    response.raise_for_status()
    return response.json()


def _flatten_response(raw: List[Dict]) -> pd.DataFrame:
    """Transform the nested API response into a flat pandas DataFrame."""

    # The API returns an object:\n{"energy": [ { "date": "...", "productionMixData": [...] }, ... ]}\n    data_list: List[Dict] = []
    if isinstance(raw, dict) and "energy" in raw:
        data_list = raw["energy"] or []
    elif isinstance(raw, list):
        data_list = raw  # Fallback for old behaviour if any
    else:
        raise ValueError("Unexpected API response structure")

    records: List[Dict] = []
    for day_entry in data_list:
        date_str = day_entry.get("date")
        for record in day_entry.get("productionMixData", []):
            record_flat = {"date": date_str}
            record_flat.update(record)
            records.append(record_flat)

    return pd.DataFrame(records)


def fetch_production_mix(from_date: datetime.date, to_date: datetime.date) -> pd.DataFrame:
    """Fetch production mix data between two dates and return as DataFrame."""

    raw_data = _call_api(from_date.strftime(DATE_FMT), to_date.strftime(DATE_FMT))
    return _flatten_response(raw_data)


def main() -> None:
    """Entry point for the script."""

    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    # 1. Fetch most recent full day (yesterday)
    print(f"Fetching data for {yesterday.strftime(DATE_FMT)} …")
    recent_df = fetch_production_mix(yesterday, yesterday)

    # 2. Fetch data from 01-Jan-2021 to yesterday
    start_date = datetime.date(2023, 2, 22)
    print(f"Fetching full history from {start_date.strftime(DATE_FMT)} to {yesterday.strftime(DATE_FMT)} …")
    history_df = fetch_production_mix(start_date, yesterday)

    # 3. Write both DataFrames to an Excel workbook
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        recent_df.to_excel(writer, sheet_name="Most Recent Day", index=False)
        history_df.to_excel(writer, sheet_name="Since 2023-02-22", index=False)

    print(f"Done! Data written to {OUTPUT_EXCEL}.")


if __name__ == "__main__":
    main() 