"""
기준스코어 수식 수정:
  변경 전: (누적평균 + 최근라운딩) × 0.5
  변경 후: (누적평균 + 26평균) × 0.5
  - 26평균 없으면 누적평균 그대로
  - 소수점 1자리
"""
import gspread, sys, time
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_ROW = 60


def col_letter(n):
    if n <= 26:
        return chr(64 + n)
    return chr(64 + (n - 1) // 26) + chr(64 + ((n - 1) % 26) + 1)


def make_ref_score(row, avg_col, avg26_col):
    A = f"{avg_col}{row}"    # 누적평균
    B = f"{avg26_col}{row}"  # 26평균
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(NOT(ISNUMBER({A})),\"\","
        f"IF(NOT(ISNUMBER({B})),"
        f"ROUND({A},1),"
        f"ROUND(({A}+{B})*0.5,1))))"
    )


def main():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    sheet = gspread.authorize(creds).open_by_key(SPREADSHEET_ID).worksheet("스코어 집계 (입력)")

    header = sheet.row_values(2)
    avg_col    = col_letter(header.index("누적평균") + 1)
    avg26_col  = col_letter(next(i + 1 for i, h in enumerate(header) if h == "26평균 스코어"))
    ref_col    = col_letter(header.index("기준스코어") + 1)

    print(f"누적평균: {avg_col}열, 26평균: {avg26_col}열, 기준스코어: {ref_col}열")

    formulas = [[make_ref_score(r, avg_col, avg26_col)] for r in range(5, MAX_ROW + 1)]
    sheet.update(range_name=f"{ref_col}5:{ref_col}{MAX_ROW}",
                 values=formulas, value_input_option="USER_ENTERED")
    print(f"[OK] {ref_col}열 수식 업데이트 완료")
    time.sleep(2)

    # 결과 확인
    result   = sheet.get("B5:B50", value_render_option="FORMATTED_VALUE")
    avg_v    = sheet.get(f"{avg_col}5:{avg_col}50",   value_render_option="FORMATTED_VALUE")
    avg26_v  = sheet.get(f"{avg26_col}5:{avg26_col}50", value_render_option="FORMATTED_VALUE")
    ref_v    = sheet.get(f"{ref_col}5:{ref_col}50",   value_render_option="FORMATTED_VALUE")

    print()
    print(f"{'이름':<12} {'누적평균':<10} {'26평균':<10} {'기준스코어'}")
    print("-" * 45)
    for i, row in enumerate(result):
        name  = row[0].strip() if row else ""
        avg   = avg_v[i][0]   if i < len(avg_v)   and avg_v[i]   else ""
        avg26 = avg26_v[i][0] if i < len(avg26_v)  and avg26_v[i] else ""
        ref   = ref_v[i][0]   if i < len(ref_v)   and ref_v[i]   else ""
        if name:
            print(f"  {name:<12} {avg:<10} {avg26:<10} {ref}")


if __name__ == "__main__":
    main()
