import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import time

# 설정
SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 26년 4월 26일 대회 스코어 (12명)
MARCH_2026_SCORES = {
    "오지호":  82,
    "박병우":  84,
    "김산":    85,
    "황인호":  94,
    "이건희":  98,
    "김준기":  102,
    "최원재":  103,
    "김현섭":  103,
    "권기표":  105,
    "최성호":  110,
    "백근호":  113,
    "주홍석":  115,
}

# 26년 3월 23일 기준 티어 (전체 60명)
TIER_MAP = {
    # 티어 0 (6명)
    "김산":       0, "이형석":  0, "이현우":  0, "윤석원":  0, "임승언":  0,
    "오지호":     0,
    # 티어 1 (11명)
    "김수민":     1, "나한영":  1, "이정승":  1, "박계홍":  1, "박병우":  1,
    "임익준":     1, "박세환":  1, "신태범":  1, "노용현":  1,
    "홍경택":     1, "김대우":  1,
    # 티어 2 (12명)
    "이민재":     2, "권기표":  2, "권한해":  2, "윤준원":  2, "류성준":  2,
    "윤현호":     2, "김성환":  2, "황인호":  2, "손원우":  2, "고영승":  2,
    "김준기":     2, "조성태":  2,
    # 티어 3 (12명)
    "최성호":     3, "김성환(92)": 3, "조현민": 3, "강충현":  3, "이준범":  3,
    "황호연":     3, "박헌우":  3, "박준상":  3, "한상윤":  3, "최정훈":  3,
    "손정모":     3, "이건희":  3,
    # 티어 4 (11명)
    "김현섭":     4, "김지수":  4, "최원재":  4, "박상용":  4, "김서래":  4,
    "이도영":     4, "박문석":  4, "백근호":  4, "최창규":  4, "노민일":  4,
    "이호준":     4,
    # 티어 5 (9명)
    "신재민":     5, "옥지엽":  5, "서정권":  5, "김진호":  5, "강승원":  5,
    "이종화":     5, "박승현":  5, "강명신":  5, "주홍석":  5,
}

NEW_MEMBERS = ["주홍석"]
EVENT_COL_NAME = "26' 2차 대회"
EVENT_DATE = "4/26/2026"


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def batch_update_cells(sheet, updates):
    """updates: list of (row, col, value) tuples (1-based row/col)"""
    if not updates:
        return
    data = [
        {"range": rowcol_to_a1(r, c), "values": [[v]]}
        for r, c, v in updates
    ]
    sheet.batch_update(data, value_input_option="USER_ENTERED")


# ─────────────────────────────────────────────
# 스코어 집계 (입력) 시트
# ─────────────────────────────────────────────
def update_score_sheet(spreadsheet):
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")
    all_values = sheet.get_all_values()

    header_row = all_values[1]  # 헤더는 2행 (index 1)
    name_col_idx = header_row.index("이름")

    # ── 새 멤버 추가 ──
    existing_names = set(
        row[name_col_idx].strip()
        for row in all_values[4:]
        if len(row) > name_col_idx and row[name_col_idx].strip()
    )
    for new_name in NEW_MEMBERS:
        if new_name not in existing_names:
            next_no = len(existing_names) + 1
            new_row = [""] * len(header_row)
            new_row[0] = str(next_no)
            new_row[name_col_idx] = new_name
            sheet.append_row(new_row, value_input_option="USER_ENTERED")
            print(f"  새 멤버 추가: {new_name}")
            existing_names.add(new_name)
            time.sleep(1)

    # ── 컬럼 삽입 ──
    all_values = sheet.get_all_values()
    header_row = all_values[1]

    if EVENT_COL_NAME not in header_row:
        avg_col_1based = header_row.index("누적평균") + 1  # 0-based → 1-based
        sheet.insert_cols([[]], col=avg_col_1based)
        print(f"  열 {avg_col_1based} 위치에 컬럼 삽입 완료")
        time.sleep(2)
        all_values = sheet.get_all_values()
        header_row = all_values[1]
        score_col_1based = header_row.index("누적평균")  # 방금 삽입된 빈 컬럼 위치 (0-based = avg-1)
        score_col_1based = avg_col_1based  # 삽입 위치 = 새 컬럼 위치 (1-based)
        # 헤더·날짜 배치 입력
        batch_update_cells(sheet, [
            (2, score_col_1based, EVENT_COL_NAME),
            (3, score_col_1based, EVENT_DATE),
        ])
        print(f"  헤더/날짜 입력 완료")
        time.sleep(1)
        all_values = sheet.get_all_values()
    else:
        print(f"  '{EVENT_COL_NAME}' 컬럼 이미 존재")

    # ── 스코어 + 티어 배치 입력 ──
    all_values = sheet.get_all_values()
    header_row = all_values[1]
    name_col_idx = header_row.index("이름")
    score_col_1based = header_row.index(EVENT_COL_NAME) + 1  # 0-based → 1-based
    tier_col_1based = header_row.index("현재 티어") + 1 if "현재 티어" in header_row else None

    score_updates = []
    tier_updates = []
    score_done, tier_done = [], []
    score_missing, tier_missing = [], []

    for row_idx, row in enumerate(all_values[4:], start=5):
        if len(row) <= name_col_idx:
            continue
        name = row[name_col_idx].strip()
        if not name:
            continue

        if name in MARCH_2026_SCORES:
            score_updates.append((row_idx, score_col_1based, MARCH_2026_SCORES[name]))
            score_done.append(name)

        if tier_col_1based and name in TIER_MAP:
            tier_updates.append((row_idx, tier_col_1based, TIER_MAP[name]))
            tier_done.append(name)

    # 스코어 배치 전송
    batch_update_cells(sheet, score_updates)
    print(f"  [OK] 스코어 입력: {len(score_done)}명")

    score_missing = [n for n in MARCH_2026_SCORES if n not in score_done]
    if score_missing:
        print(f"  [!!] 스코어 미입력 (시트에 없음): {score_missing}")

    time.sleep(2)

    # 티어 배치 전송
    if tier_updates:
        batch_update_cells(sheet, tier_updates)
        print(f"  [OK] 티어 업데이트: {len(tier_done)}명")
    else:
        print("  [!!] 티어 컬럼을 찾을 수 없거나 업데이트 대상 없음")

    tier_missing = [n for n in TIER_MAP if n not in tier_done]
    if tier_missing:
        print(f"  [!!] 티어 미입력 (시트에 없음): {tier_missing}")


# ─────────────────────────────────────────────
# 멤버명단 (자동) 시트
# ─────────────────────────────────────────────
def update_member_sheet(spreadsheet):
    sheet = spreadsheet.worksheet("멤버명단 (자동)")
    all_values = sheet.get_all_values()

    header_row = all_values[0]  # 헤더는 1행 (index 0)
    name_col_idx = header_row.index("이름") if "이름" in header_row else 1

    # ── 새 멤버 추가 ──
    existing_names = set(
        row[name_col_idx].strip()
        for row in all_values[2:]
        if len(row) > name_col_idx and row[name_col_idx].strip()
    )
    for new_name in NEW_MEMBERS:
        if new_name not in existing_names:
            next_no = len(existing_names) + 1
            new_row = [""] * len(header_row)
            new_row[0] = str(next_no)
            new_row[name_col_idx] = new_name
            sheet.append_row(new_row, value_input_option="USER_ENTERED")
            print(f"  멤버명단 새 멤버 추가: {new_name}")
            existing_names.add(new_name)
            time.sleep(1)

    # ── 컬럼 추가 ──
    all_values = sheet.get_all_values()
    header_row = all_values[0]

    if EVENT_COL_NAME not in header_row:
        insert_col = len(header_row) + 1
        batch_update_cells(sheet, [
            (1, insert_col, EVENT_COL_NAME),
            (2, insert_col, EVENT_DATE),
        ])
        print(f"  멤버명단 헤더/날짜 입력 (열 {insert_col})")
        time.sleep(1)
        all_values = sheet.get_all_values()
    else:
        print(f"  멤버명단 '{EVENT_COL_NAME}' 컬럼 이미 존재")

    # ── YES/NO 배치 입력 ──
    header_row = all_values[0]
    name_col_idx = header_row.index("이름") if "이름" in header_row else 1
    event_col_1based = header_row.index(EVENT_COL_NAME) + 1

    yn_updates = []
    for row_idx, row in enumerate(all_values[2:], start=3):
        if len(row) <= name_col_idx:
            continue
        name = row[name_col_idx].strip()
        if not name:
            continue
        value = "YES" if name in MARCH_2026_SCORES else "NO"
        yn_updates.append((row_idx, event_col_1based, value))

    batch_update_cells(sheet, yn_updates)
    print(f"  [OK] 멤버명단 YES/NO 입력: {len(yn_updates)}명")


def main():
    print("Google Sheets 연결 중...")
    try:
        spreadsheet = connect()
        print("[OK] 연결 성공!\n")
    except Exception as e:
        print(f"[ERR] 연결 실패: {e}")
        return

    print("=== 스코어 집계 (입력) 시트 ===")
    update_score_sheet(spreadsheet)

    time.sleep(3)

    print("\n=== 멤버명단 (자동) 시트 ===")
    update_member_sheet(spreadsheet)

    print("\n[완료] 전체 업데이트 완료!")


if __name__ == "__main__":
    main()
