"""
이은광, 최창환, 강현수 → 현재 티어 1로 업데이트
- 스코어 집계 (입력) 시트: 현재 티어 컬럼
- 티어정리_신규 시트: 5회 컬럼 (26' 2차 대회 4/26/2026 기준)
"""
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import time
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 이번에 T1로 올릴 멤버
FIX_TIER_1 = {"이은광", "최창환", "강현수"}

# 김산: T4 → T0 복귀
FIX_TIER_CUSTOM = {"김산": 0}


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def batch_update(sheet, updates):
    if not updates:
        return
    data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in updates]
    sheet.batch_update(data, value_input_option="USER_ENTERED")


def fix_score_sheet(spreadsheet):
    print("=== 스코어 집계 (입력) 시트: 현재 티어 수정 ===")
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")
    all_values = sheet.get_all_values()

    header_row = all_values[1]  # 2행이 헤더
    name_col = header_row.index("이름")
    tier_col = header_row.index("현재 티어") if "현재 티어" in header_row else None

    if tier_col is None:
        print("  [!!] 현재 티어 컬럼 없음")
        return

    tier_col_1based = tier_col + 1
    updates = []

    for i, row in enumerate(all_values[4:], start=5):  # 5행부터 데이터
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if name in FIX_TIER_1:
            old = row[tier_col].strip() if len(row) > tier_col else ""
            updates.append((i, tier_col_1based, 1))
            print(f"  {name}: {old or '(없음)'} → 1")
        elif name in FIX_TIER_CUSTOM:
            new_tier = FIX_TIER_CUSTOM[name]
            old = row[tier_col].strip() if len(row) > tier_col else ""
            updates.append((i, tier_col_1based, new_tier))
            print(f"  {name}: {old or '(없음)'} → {new_tier}")

    if updates:
        batch_update(sheet, updates)
        print(f"  [OK] {len(updates)}명 티어 1로 수정")
    else:
        print("  [!!] 해당 멤버 행을 찾지 못함")


def fix_tier_new_sheet(spreadsheet):
    print("\n=== 티어정리_신규 시트: 5회 컬럼 수정 ===")
    sheet = spreadsheet.worksheet("티어정리_신규")
    all_values = sheet.get_all_values()

    # 헤더 행 찾기
    hi = next((i for i, row in enumerate(all_values) if "이름" in row), None)
    if hi is None:
        print("  [!!] 이름 컬럼 없음")
        return

    header = all_values[hi]
    name_col = header.index("이름")

    # 회차 컬럼 (No, 이름 제외한 것들)
    round_cols = [h for h in header if h and h not in ("No", "이름")]
    print(f"  현재 회차 컬럼: {round_cols}")

    # 5회 컬럼 위치 확인
    if "5회" not in header:
        print("  [!!] 5회 컬럼 없음 — 자동 추가하지 않음 (수동으로 열 추가 후 재실행 필요)")
        return

    col_5 = header.index("5회") + 1  # 1-based

    # 날짜 행 (헤더 바로 다음 행)
    date_row_idx = hi + 2  # hi는 0-based, 구글시트는 1-based → hi+1+1

    updates = []
    for i, row in enumerate(all_values[hi + 2:], start=hi + 3):  # 날짜행 스킵
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if not name:
            continue
        if name in FIX_TIER_1:
            old = row[col_5 - 1].strip() if len(row) >= col_5 else ""
            updates.append((i, col_5, 1))
            print(f"  {name}: {old or '(없음)'} → 1")
        elif name in FIX_TIER_CUSTOM:
            new_tier = FIX_TIER_CUSTOM[name]
            old = row[col_5 - 1].strip() if len(row) >= col_5 else ""
            updates.append((i, col_5, new_tier))
            print(f"  {name}: {old or '(없음)'} → {new_tier}")

    if updates:
        batch_update(sheet, updates)
        print(f"  [OK] {len(updates)}명 5회 티어 1로 수정")
    else:
        print("  [!!] 해당 멤버를 찾지 못함 (시트에 없거나 5회 컬럼 문제)")


def main():
    print("Google Sheets 연결 중...")
    try:
        spreadsheet = connect()
        print("[OK] 연결 성공!\n")
    except Exception as e:
        print(f"[ERR] 연결 실패: {e}")
        return

    fix_score_sheet(spreadsheet)
    time.sleep(2)
    fix_tier_new_sheet(spreadsheet)

    print("\n[완료]")


if __name__ == "__main__":
    main()
