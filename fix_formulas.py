"""
두 가지 수식 문제 일괄 수정:
1. 스코어 집계 S~Z 열: $41 하드코딩 → $60, 행5~60으로 확장
2. 멤버명단 N열 이후: 하드코딩 YES/NO → XLOOKUP 수식
"""
import gspread
import re
import sys
import time
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SCORE_MAX_ROW = 60   # 스코어 집계 데이터 최대 행 (최대 56명)
MEMBER_MAX_ROW = 90  # 멤버명단 YES/NO 수식 최대 행


def col_to_letter(n):
    return chr(64 + n)


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def batch_update_cells(sheet, updates, chunk_size=180):
    """180개씩 끊어 API 요청 (Google Sheets 배치 제한 대응)"""
    for i in range(0, len(updates), chunk_size):
        chunk = updates[i:i + chunk_size]
        data = [{"range": rowcol_to_a1(r, c), "values": [[v]]} for r, c, v in chunk]
        sheet.batch_update(data, value_input_option="USER_ENTERED")
        if i + chunk_size < len(updates):
            time.sleep(1)


def extend_row_formula(template, from_row, to_row, old_range_max, new_range_max):
    """
    수식 템플릿의 행 참조를 변환:
    - 절대 범위 끝 ($old_range_max → $new_range_max)
    - 상대 행 참조 (from_row → to_row, $로 고정된 건 제외)
    """
    result = template.replace(f'${old_range_max}', f'${new_range_max}')
    # column_letter + from_row (앞에 $ 없는 것만) → to_row
    result = re.sub(rf'(?<!\$){re.escape(str(from_row))}\b', str(to_row), result)
    return result


# ─────────────────────────────────────────────
# 1. 스코어 집계: S~Z 수식 행 확장 + $41 → $60
# ─────────────────────────────────────────────
def fix_score_formulas(spreadsheet):
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")

    # 5행 템플릿 수식 읽기 (S~Z)
    raw = sheet.get('S5:Z5', value_render_option='FORMULA')
    if not raw:
        print("  [!!] 템플릿 수식을 읽을 수 없음")
        return

    col_letters = list('STUVWXYZ')
    templates = {col_letters[i]: raw[0][i] if i < len(raw[0]) else '' for i in range(8)}

    OLD_MAX = 41
    updates = []

    for col_letter, tmpl in templates.items():
        col_idx = ord(col_letter) - 64  # S=19, T=20, ...
        if not tmpl or not str(tmpl).startswith('='):
            print(f"  {col_letter}열: 수식 없음(값 셀), 스킵")
            continue
        for row_idx in range(5, SCORE_MAX_ROW + 1):
            formula = extend_row_formula(tmpl, 5, row_idx, OLD_MAX, SCORE_MAX_ROW)
            updates.append((row_idx, col_idx, formula))

    batch_update_cells(sheet, updates)
    print(f"  [OK] S~Z 수식 행5~{SCORE_MAX_ROW} 업데이트 ({len(updates)}셀)")


# ─────────────────────────────────────────────
# 2. 멤버명단: 하드코딩 YES/NO → XLOOKUP 수식
# ─────────────────────────────────────────────
def fix_member_formulas(spreadsheet):
    sheet = spreadsheet.worksheet("멤버명단 (자동)")
    formulas = sheet.get_all_values(value_render_option='FORMULA')
    header = formulas[0]

    # 3행 기준으로 수식이 아닌 이벤트 칼럼 찾기 (A, B 제외)
    cols_to_fix = []
    for i, h in enumerate(header):
        if i < 2 or not h:
            continue
        data_val = formulas[2][i] if i < len(formulas[2]) else ''
        if data_val and not str(data_val).startswith('='):
            col_letter = col_to_letter(i + 1)
            cols_to_fix.append((i + 1, col_letter))

    if not cols_to_fix:
        print("  이미 모든 이벤트 칼럼이 수식입니다.")
        return

    print(f"  수식 교체 대상: {[c[1] for c in cols_to_fix]} 열")
    updates = []

    for col_1based, col_letter in cols_to_fix:
        # 행1: 스코어 집계 헤더 참조
        updates.append((1, col_1based, f"='스코어 집계 (입력)'!{col_letter}2"))
        # 행2: 날짜 XLOOKUP
        updates.append((2, col_1based,
            f"=XLOOKUP({col_letter}1,'스코어 집계 (입력)'!2:2,'스코어 집계 (입력)'!3:3)"))
        # 행3~MEMBER_MAX_ROW: YES/NO 수식
        for row_idx in range(3, MEMBER_MAX_ROW + 1):
            formula = (
                f"=IF($B{row_idx}=\"\",\"\","
                f"IFERROR(IF(XLOOKUP($B{row_idx},'스코어 집계 (입력)'!$B:$B,"
                f"'스코어 집계 (입력)'!{col_letter}:{col_letter}),\"YES\",\"NO\"),\"NO\"))"
            )
            updates.append((row_idx, col_1based, formula))

    batch_update_cells(sheet, updates)
    print(f"  [OK] {len(cols_to_fix)}개 칼럼 행1~{MEMBER_MAX_ROW} 수식 교체 완료")


def main():
    print("연결 중...")
    spreadsheet = connect()
    print("[OK] 연결 성공!\n")

    print("=== 스코어 집계: S~Z 수식 행 확장 + $41→$60 ===")
    fix_score_formulas(spreadsheet)
    time.sleep(3)

    print("\n=== 멤버명단: 하드코딩 YES/NO → XLOOKUP 수식 ===")
    fix_member_formulas(spreadsheet)
    print("\n[완료]")


if __name__ == "__main__":
    main()
