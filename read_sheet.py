import gspread, json, sys
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

sheet = spreadsheet.worksheet("스코어 집계 (입력)")
all_vals = sheet.get_all_values()

print("총 행수:", len(all_vals))
print("헤더(2행):", all_vals[1])
print()
print("데이터 처음 5행:")
for row in all_vals[4:9]:
    print(" ", row[:18])
print()
print("마지막 5행:")
for row in all_vals[-5:]:
    print(" ", row[:18])
