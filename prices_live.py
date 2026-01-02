import requests
import openpyxl
from datetime import datetime
import os

# ================= CONFIG =================
EXCEL_PATH = "Macro_Data.xlsx"
SHEET_NAME = "Live_Prices"

FINNHUB_API_KEY = os.getenv("FINNHUB_API_KEY")

# ================= FETCH PRICE =================
def get_price(symbol):
    url = (
        f"https://finnhub.io/api/v1/quote"
        f"?symbol={symbol}&token={FINNHUB_API_KEY}"
    )
    r = requests.get(url)
    r.raise_for_status()
    return r.json()["c"]

# ================= MAIN =================
wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb[SHEET_NAME]

timestamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

# Start from row 2 (skip headers)
for row in range(2, ws.max_row + 1):
    symbol = ws[f"A{row}"].value

    # Skip empty rows
    if not symbol:
        continue

    price = get_price(symbol)

    ws[f"B{row}"] = price
    ws[f"C{row}"] = timestamp

wb.save(EXCEL_PATH)

print("Live prices updated successfully")
