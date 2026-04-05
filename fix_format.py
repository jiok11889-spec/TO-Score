# -*- coding: utf-8 -*-
"""
새로 추가된 컬럼들의 서식을 인접 컬럼에서 복사하여 적용
"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    client  = gspread.authorize(creds)
    sp = client.open_by_key(SPREADSHEET_ID)
    sheet_ids = {ws.title: ws.id for ws in sp.worksheets()}
    return service, sp, sheet_ids

def copy_format(service, src_sheet_id, dst_sheet_id,
                src_col0, dst_col0, row_start0, row_end0):
    """
    서식만 복사 (PASTE_FORMAT).
    src_col0, dst_col0: 0-based 컬럼 인덱스 (단일 컬럼)
    row_start0, row_end0: 0-based 행 범위 (exclusive end)
    """
    req = {
        "copyPaste": {
            "source": {
                "sheetId": src_sheet_id,
                "startRowIndex": row_start0,
                "endRowIndex":   row_end0,
                "startColumnIndex": src_col0,
                "endColumnIndex":   src_col0 + 1,
            },
            "destination": {
                "sheetId": dst_sheet_id,
                "startRowIndex": row_start0,
                "endRowIndex":   row_end0,
                "startColumnIndex": dst_col0,
                "endColumnIndex":   dst_col0 + 1,
            },
            "pasteType": "PASTE_FORMAT",
        }
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [req]}
    ).execute()

def color_range(sheet_id, row_start0, row_end0, col0, color):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_start0, "endRowIndex": row_end0,
                "startColumnIndex": col0,    "endColumnIndex": col0 + 1,
            },
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    }

def apply(service, requests):
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests}
        ).execute()

def main():
    service, sp, sheet_ids = connect()

    # ── 각 시트 열기 ──
    score_sheet  = sp.worksheet("스코어 집계 (입력)")
    member_sheet = sp.worksheet("멤버명단 (자동)")
    tier_new     = sp.worksheet("티어정리_신규 (작업중)")
    tier_old     = sp.worksheet("티어정리 (작업중)")

    sid_score  = sheet_ids["스코어 집계 (입력)"]
    sid_member = sheet_ids["멤버명단 (자동)"]
    sid_tier_n = sheet_ids["티어정리_신규 (작업중)"]
    sid_tier_o = sheet_ids["티어정리 (작업중)"]

    BLUE  = {"red": 0.812, "green": 0.886, "blue": 0.953}
    GREEN = {"red": 0.886, "green": 0.937, "blue": 0.851}
    GRAY  = {"red": 0.953, "green": 0.953, "blue": 0.953}

    # ═══════════════════════════════════════
    # 1. 스코어 집계 (입력)
    # ═══════════════════════════════════════
    print("=== 스코어 집계 서식 ===")
    vals = score_sheet.get_all_values()
    header = vals[1]
    event_col0 = header.index("26' 1차 대회")  # 0-based
    # 기준: "25' 3차 대회" (파란색 차대회 컬럼)
    ref_col0 = header.index("25' 3차 대회")
    total_rows = len(vals)

    # 전체 행 서식 복사 (참조 컬럼 → 신규 컬럼)
    copy_format(service, sid_score, sid_score, ref_col0, event_col0, 0, total_rows)
    print(f"  col{ref_col0+1}('25'3차') 서식 → col{event_col0+1}('26'1차') 복사 완료")
    time.sleep(1)

    # 헤더 행(row index 1)만 파란색으로 덮어씌우기 (복사된 서식에 추가)
    apply(service, [color_range(sid_score, 1, 2, event_col0, BLUE)])
    print("  헤더 파란색 보정 완료")
    time.sleep(1)

    # ═══════════════════════════════════════
    # 2. 멤버명단 (자동)
    # ═══════════════════════════════════════
    print("\n=== 멤버명단 서식 ===")
    vals = member_sheet.get_all_values()
    header = vals[0]
    event_col0 = header.index("26' 1차 대회")
    ref_col0   = event_col0 - 1  # 바로 앞 컬럼 ("25' 4차 대회")
    total_rows = len(vals)

    copy_format(service, sid_member, sid_member, ref_col0, event_col0, 0, total_rows)
    print(f"  col{ref_col0+1} 서식 → col{event_col0+1} 복사 완료")
    time.sleep(1)

    # ═══════════════════════════════════════
    # 3. 티어정리_신규
    # ═══════════════════════════════════════
    print("\n=== 티어정리_신규 서식 ===")
    vals = tier_new.get_all_values()
    header      = vals[7]  # row index 7 = 헤더행
    data_start0 = 9        # 개인 데이터 시작 (0-based)
    total_rows  = len(vals)

    # 신규 컬럼 위치: "4회" (마지막 '-' 전의 첫 빈칸 또는 헤더 기준)
    # 데이터가 있는 마지막 컬럼 찾기
    last_filled_col0 = 2
    for col_i in range(2, len(header)):
        has_data = any(
            vals[r][col_i].strip() not in ("", "-")
            for r in range(data_start0, total_rows)
            if len(vals[r]) > col_i
        )
        if has_data:
            last_filled_col0 = col_i
    new_col0 = last_filled_col0 + 1

    # 전체 서식 복사 (이전 최신 컬럼 → 신규 컬럼)
    copy_format(service, sid_tier_n, sid_tier_n, last_filled_col0, new_col0, 0, total_rows)
    print(f"  col{last_filled_col0+1} 서식 → col{new_col0+1} 복사")
    time.sleep(1)

    # 색상: 이전 최신 → 회색, 신규 → 녹색
    apply(service, [
        color_range(sid_tier_n, data_start0, total_rows, last_filled_col0, GRAY),
        color_range(sid_tier_n, data_start0, total_rows, new_col0,         GREEN),
    ])
    print(f"  col{last_filled_col0+1} 회색, col{new_col0+1} 녹색 적용")
    time.sleep(1)

    # ═══════════════════════════════════════
    # 4. 티어정리 (작업중)
    # ═══════════════════════════════════════
    print("\n=== 티어정리 (기존) 서식 ===")
    vals = tier_old.get_all_values()
    data_start0 = 9
    total_rows  = len(vals)

    last_filled_col0 = 2
    for col_i in range(2, 12):
        has_data = any(
            vals[r][col_i].strip() not in ("", "-")
            for r in range(data_start0, total_rows)
            if len(vals[r]) > col_i
        )
        if has_data:
            last_filled_col0 = col_i
    new_col0 = last_filled_col0 + 1

    copy_format(service, sid_tier_o, sid_tier_o, last_filled_col0, new_col0, 0, total_rows)
    print(f"  col{last_filled_col0+1} 서식 → col{new_col0+1} 복사")
    time.sleep(1)

    apply(service, [
        color_range(sid_tier_o, data_start0, total_rows, last_filled_col0, GRAY),
        color_range(sid_tier_o, data_start0, total_rows, new_col0,         GREEN),
    ])
    print(f"  col{last_filled_col0+1} 회색, col{new_col0+1} 녹색 적용")

    print("\n[완료] 서식 업데이트 완료!")

if __name__ == "__main__":
    main()
