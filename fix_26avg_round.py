import gspread, sys, time
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_ROW = 60


def make_avg26(row):
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(COUNTA(N{row}:Q{row})=0,\"\","
        f"IFERROR(ROUND(AVERAGE(N{row}:Q{row}),1),\"\")))"
    )


def main():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet("스코어 집계 (입력)")

    header = sheet.row_values(2)
    t_col = header.index("26평균 스코어") + 1  # 1-based

    data = [[make_avg26(r)] for r in range(5, MAX_ROW + 1)]
    sheet.update(range_name=f"{chr(64+t_col)}5:{chr(64+t_col)}{MAX_ROW}",
                 values=data, value_input_option="USER_ENTERED")
    print(f"[OK] {chr(64+t_col)}열 ROUND(...,1) 적용 완료")
    time.sleep(2)

    # 확인
    result = sheet.get(f"B5:U15", value_render_option="FORMATTED_VALUE")
    print(f"\n{'이름':<12} {'26평균':<10} {'26랭킹'}")
    print("-" * 30)
    for row in result:
        name = row[0].strip() if row else ""
        avg  = row[t_col - 2] if len(row) > t_col - 2 else ""
        rank = row[t_col - 1] if len(row) > t_col - 1 else ""
        if name:
            print(f"  {name:<12} {avg:<10} {rank}")


if __name__ == "__main__":
    main()
