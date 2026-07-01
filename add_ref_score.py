"""
U열(26평균 스코어 랭킹) 바로 오른쪽에 기준스코어 열 추가
기준스코어 = (누적평균 + 최근라운딩) × 0.5  [가중평균]
- 최근라운딩 없으면 누적평균 그대로
- 소수점 1자리
"""
import gspread, sys, time
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_ROW = 60


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def col_letter(n):
    """1-based → 열 문자 (AA, AB 지원)"""
    if n <= 26:
        return chr(64 + n)
    return chr(64 + (n - 1) // 26) + chr(64 + ((n - 1) % 26) + 1)


def make_ref_score(row, avg_col, recent_col):
    """
    기준스코어 수식:
    - 이름 없으면 ''
    - 누적평균 없으면 ''
    - 최근라운딩 없으면 누적평균 그대로 표시
    - 둘 다 있으면 (avg + recent) * 0.5, 소수점 1자리
    """
    A = f"{avg_col}{row}"    # 누적평균 셀
    B = f"{recent_col}{row}" # 최근라운딩 셀
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(NOT(ISNUMBER({A})),\"\","
        f"IF(NOT(ISNUMBER({B})),"
        f"ROUND({A},1),"
        f"ROUND(({A}+{B})*0.5,1))))"
    )


def main():
    print("연결 중...")
    spreadsheet = connect()
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")

    header = sheet.row_values(2)

    # 열 위치 파악
    u_col    = next(i + 1 for i, h in enumerate(header) if "26평균" in h and "랭킹" in h)
    avg_col  = next(i + 1 for i, h in enumerate(header) if h == "누적평균")
    recent_col_1based = next(
        i + 1 for i, h in enumerate(header) if "최근라운딩" in h
    )

    insert_at    = u_col + 1           # U 바로 오른쪽
    avg_letter    = col_letter(avg_col)
    recent_letter = col_letter(recent_col_1based)
    new_col_letter = col_letter(insert_at)

    print(f"  누적평균: {avg_letter}열, 최근라운딩: {recent_letter}열")
    print(f"  기준스코어 삽입 위치: {new_col_letter}열 (U열 오른쪽)")

    if "기준스코어" in header:
        print("[!!] '기준스코어' 열이 이미 존재합니다.")
        return

    # ── 1. 열 삽입 ──
    sheet.insert_cols([[]], col=insert_at)
    print(f"[OK] {new_col_letter}열 삽입")
    time.sleep(2)

    # 삽입 후 최근라운딩 열이 한 칸 밀림 → 재계산
    recent_letter = col_letter(recent_col_1based + 1)
    new_col_letter = col_letter(insert_at)
    print(f"  삽입 후 최근라운딩: {recent_letter}열")

    # ── 2. 헤더 입력 ──
    updates = [(2, insert_at, "기준스코어")]
    data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in updates]
    sheet.batch_update(data, value_input_option="USER_ENTERED")
    time.sleep(1)

    # ── 3. 수식 입력 (행 5~60) ──
    formulas = [[make_ref_score(r, avg_letter, recent_letter)] for r in range(5, MAX_ROW + 1)]
    sheet.update(
        range_name=f"{new_col_letter}5:{new_col_letter}{MAX_ROW}",
        values=formulas,
        value_input_option="USER_ENTERED"
    )
    print(f"[OK] 기준스코어 수식 입력 완료 (행5~{MAX_ROW})")
    time.sleep(2)

    # ── 4. 결과 확인 ──
    result = sheet.get("B5:B50", value_render_option="FORMATTED_VALUE")
    ref_vals = sheet.get(
        f"{new_col_letter}5:{new_col_letter}50",
        value_render_option="FORMATTED_VALUE"
    )
    avg_vals = sheet.get(f"{avg_letter}5:{avg_letter}50", value_render_option="FORMATTED_VALUE")
    recent_vals = sheet.get(f"{recent_letter}5:{recent_letter}50", value_render_option="FORMATTED_VALUE")

    print()
    print(f"{'이름':<12} {'누적평균':<10} {'최근라운딩':<12} {'기준스코어'}")
    print("-" * 45)
    for i, row in enumerate(result):
        name   = row[0].strip() if row else ""
        avg    = avg_vals[i][0] if i < len(avg_vals) and avg_vals[i] else ""
        recent = recent_vals[i][0] if i < len(recent_vals) and recent_vals[i] else ""
        ref    = ref_vals[i][0] if i < len(ref_vals) and ref_vals[i] else ""
        if name:
            print(f"  {name:<12} {avg:<10} {recent:<12} {ref}")

    print("\n[완료]")


if __name__ == "__main__":
    main()
