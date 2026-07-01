"""
S열(누적평균 랭킹) 바로 오른쪽에 두 열 추가:
  새 T열: 26평균 스코어  (N:Q 범위 AVERAGE, 빈 버퍼 Q 포함 → 신규 대회 자동 확장)
  새 U열: 26평균 스코어 랭킹
"""
import gspread, sys, time
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES  = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_ROW = 60


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def batch_write(sheet, updates):
    if not updates:
        return
    data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in updates]
    sheet.batch_update(data, value_input_option="USER_ENTERED")


def make_avg26(row):
    """26년 평균 스코어 수식 (N~Q, 빈 버퍼 Q는 AVERAGE가 무시)"""
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(COUNTA(N{row}:Q{row})=0,\"\","
        f"IFERROR(AVERAGE(N{row}:Q{row}),\"\")))"
    )


def make_avg26_rank(row):
    """26평균 스코어 랭킹 수식 (T열 기준, S열 랭킹과 동일 구조)"""
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(NOT(ISNUMBER(T{row})),\"\","
        f"IF(T{row}=0,\"\","
        f"$T$3+1-(RANK.EQ(T{row},T$5:T${MAX_ROW})"
        f"+COUNTIFS(T$5:T${MAX_ROW},T{row},$B$5:$B${MAX_ROW},\"<\"&$B{row})))))"
    )


def main():
    print("연결 중...")
    spreadsheet = connect()
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")

    # ── 현재 헤더 확인 ──
    header = sheet.row_values(2)
    s_col_1based = header.index("누적평균 랭킹") + 1  # S열 1-based
    insert_at = s_col_1based + 1                       # S 바로 뒤
    t_col_letter = chr(64 + insert_at)                 # 새 T열 문자
    u_col_letter = chr(64 + insert_at + 1)             # 새 U열 문자

    print(f"누적평균 랭킹: {chr(64+s_col_1based)}열 → "
          f"새 열 삽입 위치: {t_col_letter} ({t_col_letter}=26평균, {u_col_letter}=26평균랭킹)")

    # 이미 추가된 경우 중단
    if "26평균 스코어" in header:
        print("[!!] '26평균 스코어' 열이 이미 존재합니다.")
        return

    # ── 1. 2열 삽입 ──
    sheet.insert_cols([[], []], col=insert_at)
    print(f"[OK] {t_col_letter}, {u_col_letter}열 삽입 완료")
    time.sleep(2)

    # ── 2. 헤더 / 카운트 수식 입력 ──
    updates = [
        (2, insert_at,     "26평균 스코어"),
        (2, insert_at + 1, "26평균\n스코어 랭킹"),
        # T3: 26평균이 있는 멤버 수 (U열 랭킹 수식의 $T$3 참조)
        (3, insert_at, f"=COUNTIFS({t_col_letter}5:{t_col_letter}{MAX_ROW},\">0\")"),
    ]
    batch_write(sheet, updates)
    print("[OK] 헤더 / 카운트 수식 입력 완료")
    time.sleep(1)

    # ── 3. 데이터 수식 (행 5~60) ──
    data_updates = []
    for row in range(5, MAX_ROW + 1):
        data_updates.append((row, insert_at,     make_avg26(row)))
        data_updates.append((row, insert_at + 1, make_avg26_rank(row)))

    # 180개씩 배치
    for i in range(0, len(data_updates), 180):
        chunk = data_updates[i:i + 180]
        data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in chunk]
        sheet.batch_update(data, value_input_option="USER_ENTERED")
        time.sleep(1)

    print(f"[OK] 데이터 수식 입력 완료 (행5~{MAX_ROW})")
    time.sleep(2)

    # ── 4. 결과 확인 ──
    result = sheet.get(f"B5:U{MAX_ROW}", value_render_option="FORMATTED_VALUE")
    # T, U는 삽입 후 새 위치 — 1-based 인덱스로 B=0, T=insert_at-2, U=insert_at-1
    t_idx = insert_at - 1   # B기준 0-based
    u_idx = insert_at        # B기준 0-based

    print()
    print(f"{'이름':<12} {'26평균':<8} {'26랭킹'}")
    print("-" * 30)
    for row in result:
        name   = row[0].strip() if row else ""
        avg26  = row[t_idx] if len(row) > t_idx else ""
        rank26 = row[u_idx] if len(row) > u_idx else ""
        if name:
            print(f"  {name:<12} {avg26:<8} {rank26}")

    print("\n[완료]")


if __name__ == "__main__":
    main()
