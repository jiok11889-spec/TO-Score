# -*- coding: utf-8 -*-
"""
스프레드시트 종합 수정 스크립트
- 스코어 집계: 컬럼 위치 교정, 새 멤버 행 수정, 스코어/티어 입력, 헤더 색상
- 멤버명단: 새 컬럼 + 새 멤버
- 티어정리_신규/기존: 새 컬럼 + 색상 업데이트
"""
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from gspread.utils import rowcol_to_a1
import time

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ─── 데이터 ─────────────────────────────────
MARCH_2026_SCORES = {
    "김산": 84, "김수민": 85, "윤준원": 88, "박병우": 88,
    "나한영": 89, "홍경택": 89, "류성준": 90, "이준범": 91,
    "김지수": 91, "권한해": 92, "임익준": 93, "강충현": 97,
    "이민재": 98, "황호연": 98, "이건희": 99, "김준기": 99,
    "이호준": 100, "김성환": 100, "신재민": 101, "권기표": 102,
    "조현민": 104, "김현섭": 105, "서정권": 106, "옥지엽": 110,
}

TIER_MAP = {
    "김산": 0, "이형석": 0, "이현우": 0, "윤석원": 0, "임승언": 0,
    "김수민": 1, "나한영": 1, "이정승": 1, "박계홍": 1, "박병우": 1,
    "오지호": 1, "임익준": 1, "박세환": 1, "신태범": 1, "노용현": 1,
    "홍경택": 1, "김대우": 1,
    "이민재": 2, "권기표": 2, "권한해": 2, "윤준원": 2, "류성준": 2,
    "윤현호": 2, "김성환": 2, "황인호": 2, "손원우": 2, "고영승": 2,
    "최성호": 2, "조성태": 2,
    "김준기": 3, "김성환(92)": 3, "조현민": 3, "강충현": 3, "이준범": 3,
    "황호연": 3, "박헌우": 3, "박준상": 3, "한상윤": 3, "최정훈": 3,
    "손정모": 3, "최원재": 3,
    "김현섭": 4, "김지수": 4, "이건희": 4, "박상용": 4, "김서래": 4,
    "이도영": 4, "박문석": 4, "백근호": 4, "최창규": 4, "노민일": 4,
    "이호준": 4,
    "신재민": 5, "옥지엽": 5, "서정권": 5, "김진호": 5, "강승원": 5,
    "이종화": 5, "박승현": 5, "강명신": 5,
}

TIER_COUNTS = {0: 5, 1: 12, 2: 12, 3: 12, 4: 11, 5: 8}  # 총 60
EVENT_COL_NAME = "26' 1차 대회"
EVENT_DATE = "3/22/2026"
NEW_MEMBERS = ["윤준원", "김현섭"]

# 색상 상수
BLUE  = {"red": 0.812, "green": 0.886, "blue": 0.953}   # 차 대회 컬럼
GREEN = {"red": 0.886, "green": 0.937, "blue": 0.851}   # 티어정리 최신 컬럼
GRAY  = {"red": 0.953, "green": 0.953, "blue": 0.953}   # 티어정리 이전 컬럼


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    service = build("sheets", "v4", credentials=creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet, service


def batch_update_cells(sheet, updates):
    """updates: list of (row, col, value) — 1-based"""
    if not updates:
        return
    data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in updates]
    sheet.batch_update(data, value_input_option="USER_ENTERED")


def color_range_request(sheet_id, start_row, end_row, start_col, end_col, color):
    """색상 적용 request dict (0-based 인덱스)"""
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "cell": {
                "userEnteredFormat": {"backgroundColor": color}
            },
            "fields": "userEnteredFormat.backgroundColor",
        }
    }


def apply_formats(service, requests):
    if not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()


# ══════════════════════════════════════════════
# 1. 스코어 집계 (입력)
# ══════════════════════════════════════════════
def fix_score_sheet(spreadsheet, service):
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")
    sheet_id = sheet.id

    # ── (a) 잘못 추가된 하단 2행 삭제 (row 69, 70) ──
    all_vals = sheet.get_all_values()
    total_rows = len(all_vals)
    # row 69,70 = index 68,69 → No 열이 비어 있고 이름이 col 15~16에 있는 행
    rows_to_delete = []
    for i in range(total_rows - 1, -1, -1):
        row = all_vals[i]
        if len(row) > 16 and row[0] == "" and row[15] in ("63", "64"):
            rows_to_delete.append(i + 1)  # 1-based
    for r in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(r)
        print(f"  잘못된 행 삭제: row {r}")
        time.sleep(0.5)

    # ── (b) 새 멤버 이름 수정 (row 67, 68) ──
    all_vals = sheet.get_all_values()
    # 이름 없는 행 중 No=63,64 찾기
    name_fixes = []
    for i, row in enumerate(all_vals):
        if len(row) > 0 and row[0] in ("63", "64") and (len(row) < 2 or row[1] == ""):
            no = row[0]
            name = "윤준원" if no == "63" else "김현섭"
            name_fixes.append((i + 1, 2, name))  # 1-based, col B
    if name_fixes:
        batch_update_cells(sheet, name_fixes)
        print(f"  새 멤버 이름 수정: {len(name_fixes)}건")
        time.sleep(1)

    # ── (c) 컬럼 스왑: col 15('26'1차') → col 14(버퍼), col 15 비우기 ──
    all_vals = sheet.get_all_values()
    header = all_vals[1]
    event_col_0 = header.index(EVENT_COL_NAME) if EVENT_COL_NAME in header else None
    buffer_col_0 = None
    # 버퍼는 마지막 대회 컬럼 다음의 빈 컬럼
    for i in range(2, len(header)):
        if header[i] == "" and i > 0:
            if header[i-1] != "":  # 앞이 비어있지 않은 첫 번째 빈 칸
                buffer_col_0 = i
                break

    if event_col_0 is not None and buffer_col_0 is not None and event_col_0 == buffer_col_0 + 1:
        # event가 버퍼 다음에 있음 → 스왑 필요
        ev_col_1 = event_col_0 + 1   # 1-based
        buf_col_1 = buffer_col_0 + 1  # 1-based

        # col ev_col_1의 모든 값 읽기
        ev_values = sheet.col_values(ev_col_1)

        # col buf_col_1 에 쓰기
        write_updates = []
        clear_updates = []
        for row_1based, val in enumerate(ev_values, start=1):
            if val:
                write_updates.append((row_1based, buf_col_1, val))
                clear_updates.append((row_1based, ev_col_1, ""))

        batch_update_cells(sheet, write_updates)
        time.sleep(1)
        batch_update_cells(sheet, clear_updates)
        time.sleep(1)
        print(f"  컬럼 스왑: col{ev_col_1} -> col{buf_col_1}, col{ev_col_1} 비우기 완료")
    else:
        if event_col_0 is not None and buffer_col_0 is not None and event_col_0 == buffer_col_0 - 1:
            print(f"  컬럼 위치 이미 올바름 (col {event_col_0+1})")
        else:
            print(f"  [!!] 컬럼 위치 확인 필요: event_col={event_col_0}, buffer_col={buffer_col_0}")

    # ── (d) 스코어 + 티어 배치 입력 ──
    time.sleep(1)
    all_vals = sheet.get_all_values()
    header = all_vals[1]
    name_col = header.index("이름")
    event_col_1 = header.index(EVENT_COL_NAME) + 1  # 1-based
    tier_col_1  = (header.index("현재 티어") + 1) if "현재 티어" in header else None

    score_updates = []
    tier_updates = []
    for row_idx, row in enumerate(all_vals[4:], start=5):
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if not name:
            continue
        if name in MARCH_2026_SCORES:
            current_score = row[event_col_1 - 1] if len(row) > event_col_1 - 1 else ""
            if not current_score:  # 이미 입력된 경우 스킵
                score_updates.append((row_idx, event_col_1, MARCH_2026_SCORES[name]))
        if tier_col_1 and name in TIER_MAP:
            tier_updates.append((row_idx, tier_col_1, TIER_MAP[name]))

    batch_update_cells(sheet, score_updates)
    print(f"  스코어 입력: {len(score_updates)}건 (기입력 제외)")
    time.sleep(1.5)

    batch_update_cells(sheet, tier_updates)
    print(f"  티어 업데이트: {len(tier_updates)}명")
    time.sleep(1.5)

    # ── (e) 헤더 색상: event 컬럼에 파란색 적용 ──
    all_vals = sheet.get_all_values()
    header = all_vals[1]
    ev_col_0 = header.index(EVENT_COL_NAME)  # 0-based
    format_requests = [
        color_range_request(sheet_id, 1, 2, ev_col_0, ev_col_0 + 1, BLUE)  # 헤더 행만
    ]
    apply_formats(service, format_requests)
    print(f"  헤더 색상(파란) 적용: col {ev_col_0+1}")


# ══════════════════════════════════════════════
# 2. 멤버명단 (자동)
# ══════════════════════════════════════════════
def fix_member_sheet(spreadsheet):
    sheet = spreadsheet.worksheet("멤버명단 (자동)")
    all_vals = sheet.get_all_values()
    header = all_vals[0]
    name_col = header.index("이름") if "이름" in header else 1

    # ── 새 멤버 추가 ──
    existing = set(row[name_col].strip() for row in all_vals[2:]
                   if len(row) > name_col and row[name_col].strip())
    for new_name in NEW_MEMBERS:
        if new_name not in existing:
            next_no = len(existing) + 1
            new_row = [""] * len(header)
            new_row[0] = str(next_no)
            new_row[name_col] = new_name
            sheet.append_row(new_row, value_input_option="USER_ENTERED")
            print(f"  멤버명단 새 멤버: {new_name}")
            existing.add(new_name)
            time.sleep(1)
            all_vals = sheet.get_all_values()

    # ── 26' 1차 대회 컬럼 추가 ──
    header = all_vals[0]
    if EVENT_COL_NAME not in header:
        new_col_1 = len(header) + 1
        batch_update_cells(sheet, [
            (1, new_col_1, EVENT_COL_NAME),
            (2, new_col_1, EVENT_DATE),
        ])
        print(f"  멤버명단 헤더/날짜 입력 (col {new_col_1})")
        time.sleep(1)
        all_vals = sheet.get_all_values()
    else:
        print(f"  멤버명단 '{EVENT_COL_NAME}' 이미 존재")

    # ── YES/NO 입력 ──
    header = all_vals[0]
    name_col = header.index("이름") if "이름" in header else 1
    ev_col_1 = header.index(EVENT_COL_NAME) + 1
    yn_updates = []
    for row_idx, row in enumerate(all_vals[2:], start=3):
        if len(row) <= name_col:
            continue
        name = row[name_col].strip()
        if not name:
            continue
        val = "YES" if name in MARCH_2026_SCORES else "NO"
        yn_updates.append((row_idx, ev_col_1, val))

    batch_update_cells(sheet, yn_updates)
    print(f"  멤버명단 YES/NO: {len(yn_updates)}명")


# ══════════════════════════════════════════════
# 3. 티어정리 시트 업데이트
# ══════════════════════════════════════════════
def fix_tier_sheet(spreadsheet, service, sheet_title):
    sheet = spreadsheet.worksheet(sheet_title)
    sheet_id = sheet.id
    all_vals = sheet.get_all_values()

    # 헤더 행, 날짜 행, 개인 데이터 시작 행 파악
    # 구조: rows 0-5=티어요약, 6=합계, 7=헤더(No,이름,1회,2회,...), 8=날짜, 9+=개인
    header_row_idx = 7   # 0-based
    date_row_idx   = 8
    data_start_idx = 9

    header = all_vals[header_row_idx]  # [No, 이름, 1회, 2회, 3회, ...]
    dates  = all_vals[date_row_idx]

    # 마지막으로 채워진 데이터 컬럼 (1회부터 시작=index 2)
    last_filled_col_0 = 2  # at least col index 2 exists
    for col_i in range(2, len(header)):
        # 개인 데이터 행에서 해당 컬럼이 하나라도 숫자이면 filled
        has_data = any(
            all_vals[r][col_i].strip() not in ("", "-")
            for r in range(data_start_idx, len(all_vals))
            if len(all_vals[r]) > col_i
        )
        if has_data:
            last_filled_col_0 = col_i

    new_col_0 = last_filled_col_0 + 1  # 0-based
    new_col_1 = new_col_0 + 1          # 1-based

    print(f"  [{sheet_title}] 마지막 데이터 col: {last_filled_col_0}, 신규 col: {new_col_0} (1-based: {new_col_1})")

    updates = []

    # ── 날짜 행 ──
    updates.append((date_row_idx + 1, new_col_1, EVENT_DATE))  # 1-based row

    # ── 티어 요약 (rows 0-5) ──
    for tier_val in range(6):
        row_in_sheet = tier_val + 1  # 1-based
        count = TIER_COUNTS.get(tier_val, 0)
        updates.append((row_in_sheet, new_col_1, count))

    # ── 합계 (row 7 = index 6) ──
    updates.append((7, new_col_1, 60))

    # ── 개인 데이터 ──
    name_col_0 = 1  # 이름 컬럼 index (0-based)
    tier_updates_count = 0
    for row_idx, row in enumerate(all_vals[data_start_idx:], start=data_start_idx + 1):
        if len(row) <= name_col_0:
            continue
        name = row[name_col_0].strip()
        if not name:
            continue
        if name in TIER_MAP:
            updates.append((row_idx, new_col_1, TIER_MAP[name]))
            tier_updates_count += 1

    batch_update_cells(sheet, updates)
    print(f"  데이터 입력: 요약 6행 + 개인 {tier_updates_count}명")
    time.sleep(1.5)

    # ── 색상 업데이트 ──
    # 이전 최신 컬럼(last_filled_col_0): green → gray
    # 신규 컬럼(new_col_0): gray → green
    data_row_count = len(all_vals) - data_start_idx
    format_reqs = [
        # 이전 최신 컬럼 → 회색
        color_range_request(
            sheet_id,
            data_start_idx,  # 0-based start
            data_start_idx + data_row_count + 10,
            last_filled_col_0, last_filled_col_0 + 1,
            GRAY
        ),
        # 신규 컬럼 → 녹색
        color_range_request(
            sheet_id,
            data_start_idx,
            data_start_idx + data_row_count + 10,
            new_col_0, new_col_0 + 1,
            GREEN
        ),
    ]
    apply_formats(service, format_reqs)
    print(f"  색상: col{last_filled_col_0} 회색, col{new_col_0} 녹색")


# ══════════════════════════════════════════════
# main
# ══════════════════════════════════════════════
def main():
    print("연결 중...")
    spreadsheet, service = connect()
    print("[OK] 연결 성공\n")

    print("=== 1. 스코어 집계 (입력) ===")
    fix_score_sheet(spreadsheet, service)
    time.sleep(2)

    print("\n=== 2. 멤버명단 (자동) ===")
    fix_member_sheet(spreadsheet)
    time.sleep(2)

    print("\n=== 3. 티어정리_신규 (작업중) ===")
    fix_tier_sheet(spreadsheet, service, "티어정리_신규 (작업중)")
    time.sleep(2)

    print("\n=== 4. 티어정리 (작업중) ===")
    fix_tier_sheet(spreadsheet, service, "티어정리 (작업중)")

    print("\n[완료] 전체 업데이트 완료!")


if __name__ == "__main__":
    main()
