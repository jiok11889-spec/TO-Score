import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import time
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 26년 5월 17일 대회 스코어 (18명)
MAY_2026_SCORES = {
    "김산":        73,
    "나한영":      75,
    "홍경택":      77,
    "김수민":      79,
    "최창환":      82,
    "임익준":      85,
    "이은광":      85,
    "강현수":      89,
    "김성환":      81,
    "권기표":      91,
    "김준기":      96,
    "조현민":      93,
    "김성환(92)":  95,
    "강충현":      100,
    "박헌우":      104,
    "백근호":      91,
    "김현섭":      98,
    "이도영":      103,
}

# 3차 대회 이후 최신 티어
# 상승: 나한영·홍경택(T1→T0), 김성환(T2→T1), 조현민·김성환(92)(T3→T2), 백근호(T4→T3)
# 하락: 이은광·강현수(T1→T2), 김준기(T2→T3), 강충현·박헌우(T3→T4), 이도영(T4→T5)
TIER_MAP = {
    # 티어 0 (8명)
    "김산":        0, "이형석":    0, "이현우":     0, "윤석원":    0,
    "임승언":      0, "오지호":    0, "나한영":     0, "홍경택":    0,
    # 티어 1 (11명)
    "김수민":      1, "이정승":    1, "박계홍":     1, "박병우":    1,
    "임익준":      1, "박세환":    1, "신태범":     1, "노용현":    1,
    "김대우":      1, "최창환":    1, "김성환":     1,
    # 티어 2 (14명)
    "이민재":      2, "권기표":    2, "권한해":     2, "윤준원":    2,
    "류성준":      2, "윤현호":    2, "황인호":     2, "손원우":    2,
    "고영승":      2, "조성태":    2, "이은광":     2, "강현수":    2,
    "조현민":      2, "김성환(92)": 2,
    # 티어 3 (10명)
    "최성호":      3, "이준범":    3, "황호연":     3, "박준상":    3,
    "한상윤":      3, "최정훈":    3, "손정모":     3, "이건희":    3,
    "김준기":      3, "백근호":    3,
    # 티어 4 (11명)
    "김현섭":      4, "김지수":    4, "최원재":     4, "박상용":    4,
    "김서래":      4, "박문석":    4, "최창규":     4, "노민일":    4,
    "이호준":      4, "강충현":    4, "박헌우":     4,
    # 티어 5 (10명)
    "신재민":      5, "옥지엽":    5, "서정권":     5, "김진호":    5,
    "강승원":      5, "이종화":    5, "박승현":     5, "강명신":    5,
    "주홍석":      5, "이도영":    5,
}

NEW_MEMBERS = []
EVENT_COL_NAME = "26' 3차 대회"
EVENT_DATE = "5/17/2026"
TIER_ROUND_COL = "6회"
TIER_ROUND_DATE = "5/17/2026"


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def batch_update_cells(sheet, updates):
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
    header_row = all_values[1]
    name_col_idx = header_row.index("이름")

    # ── 컬럼 삽입 ──
    if EVENT_COL_NAME not in header_row:
        avg_col_1based = header_row.index("누적평균") + 1
        sheet.insert_cols([[]], col=avg_col_1based)
        print(f"  열 {avg_col_1based} 위치에 컬럼 삽입 완료")
        time.sleep(2)
        all_values = sheet.get_all_values()
        header_row = all_values[1]
        score_col_1based = avg_col_1based
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
    score_col_1based = header_row.index(EVENT_COL_NAME) + 1
    tier_col_1based = header_row.index("현재 티어") + 1 if "현재 티어" in header_row else None

    score_updates, tier_updates = [], []
    score_done, tier_done = [], []

    for row_idx, row in enumerate(all_values[4:], start=5):
        if len(row) <= name_col_idx:
            continue
        name = row[name_col_idx].strip()
        if not name:
            continue

        if name in MAY_2026_SCORES:
            score_updates.append((row_idx, score_col_1based, MAY_2026_SCORES[name]))
            score_done.append(name)

        if tier_col_1based and name in TIER_MAP:
            tier_updates.append((row_idx, tier_col_1based, TIER_MAP[name]))
            tier_done.append(name)

    batch_update_cells(sheet, score_updates)
    print(f"  [OK] 스코어 입력: {len(score_done)}명")
    score_missing = [n for n in MAY_2026_SCORES if n not in score_done]
    if score_missing:
        print(f"  [!!] 스코어 미입력 (시트에 없음): {score_missing}")

    time.sleep(2)

    if tier_updates:
        batch_update_cells(sheet, tier_updates)
        print(f"  [OK] 티어 업데이트: {len(tier_done)}명")
    tier_missing = [n for n in TIER_MAP if n not in tier_done]
    if tier_missing:
        print(f"  [!!] 티어 미입력 (시트에 없음): {tier_missing}")


# ─────────────────────────────────────────────
# 멤버명단 (자동) 시트
# ─────────────────────────────────────────────
def update_member_sheet(spreadsheet):
    sheet = spreadsheet.worksheet("멤버명단 (자동)")
    all_values = sheet.get_all_values()
    header_row = all_values[0]
    name_col_idx = header_row.index("이름") if "이름" in header_row else 1

    # ── 컬럼 추가 ──
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
        value = "YES" if name in MAY_2026_SCORES else "NO"
        yn_updates.append((row_idx, event_col_1based, value))

    batch_update_cells(sheet, yn_updates)
    print(f"  [OK] 멤버명단 YES/NO 입력: {len(yn_updates)}명")


# ─────────────────────────────────────────────
# 티어정리_신규 시트 (6회 컬럼 추가)
# ─────────────────────────────────────────────
TIER_HISTORY_GID = 1117395163  # 티어정리_신규 (작업용)

def update_tier_history_sheet(spreadsheet):
    sheet = spreadsheet.get_worksheet_by_id(TIER_HISTORY_GID)
    all_values = sheet.get_all_values()

    hi = next((i for i, row in enumerate(all_values) if "이름" in row), None)
    if hi is None:
        print("  [!!] 이름 컬럼 없음")
        return

    header = all_values[hi]
    name_col = header.index("이름")

    if TIER_ROUND_COL in header:
        print(f"  '{TIER_ROUND_COL}' 컬럼 이미 존재")
    else:
        insert_col = len(header) + 1
        batch_update_cells(sheet, [
            (hi + 1, insert_col, TIER_ROUND_COL),
            (hi + 2, insert_col, TIER_ROUND_DATE),
        ])
        print(f"  [{TIER_ROUND_COL}] 헤더/날짜 입력 (열 {insert_col})")
        time.sleep(1)
        all_values = sheet.get_all_values()
        header = all_values[hi]

    round_col_1based = header.index(TIER_ROUND_COL) + 1

    updates = []
    done = []
    for i, row in enumerate(all_values[hi + 2:], start=hi + 3):
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if not name or name not in TIER_MAP:
            continue
        updates.append((i, round_col_1based, TIER_MAP[name]))
        done.append(name)

    if updates:
        batch_update_cells(sheet, updates)
        print(f"  [OK] 티어 히스토리 입력: {len(done)}명")
    missing = [n for n in TIER_MAP if n not in done]
    if missing:
        print(f"  [!!] 티어 히스토리 미입력 (시트에 없음): {missing}")


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
    time.sleep(3)

    print("\n=== 티어정리_신규 시트 ===")
    update_tier_history_sheet(spreadsheet)

    print("\n[완료] 전체 업데이트 완료!")


if __name__ == "__main__":
    main()
