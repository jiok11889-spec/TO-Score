import gspread
from sheets_auth import credentials as sa_credentials
from gspread.utils import rowcol_to_a1
import time
import sys
from config import PRO_PLAYERS

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 26년 10월 18일 7차 대회 스코어 — 포웰CC 안성 (참가 12명, 출처: 캡쳐/261018/ SMARTSCORE 출력물)
# 실타수(스코어 점수) 기준, 신페리오 net 아님. 12명 합 1,079 → 평균 89.9 = 출력물 Team Average 일치.
# ※ 김성환 동명이인: 출력물엔 "김성환"만 표기(89타) → 김성환(92) (사용자 확인 완료, 2026-10-18)
# ※ 강승원(95)은 탈회 27명 정리(manage_members.py) 때 빠진 전 멤버, 회비 명단에도 없음 → 게스트로 보고 제외.
#    복귀 멤버면 시트 행 추가 후 이 딕셔너리에 넣고 재실행.
# ※ 박승훈(85)은 26년 9월부터 회비 납부한 신규 멤버, 이번이 첫 대회 → 규칙 E대로 기준스코어 생긴 뒤 TIER_MAP 배정.
#    시트에 행이 없으면 [!!] 미입력으로 나온다 → 행 추가 후 재실행.
# ※ 김수민(88)은 10월 복귀(config.RETIRED_PLAYERS에서 제외됨). 시트 행·과거 스코어 보존돼 있음.
OCT_2026_SCORES = {
    "박진석":      79,
    "황인호":      82,
    "김대우":      82,
    "박승훈":      85,
    "김수민":      88,
    "김성환(92)":  89,
    "백근호":      92,
    "권기표":      94,
    "이건희":      94,
    "이도영":      97,
    "이호준":     102,
}

# [7차 대회 10/18 반영 후 할 일] 아래 TIER_MAP은 아직 6차(9/11) 기준. 1차 실행으로 스코어가 들어가면
# `python tier_calc.py --emit` 결과(박승훈 첫 기록·김수민 복귀 반영)로 이 블록을 교체할 것.
#
# 6차 대회(9/11) 반영 후 기준스코어(누적평균×0.5 + 26년평균×0.5) 기반 티어 재배분
# 박정수 첫 기록으로 N=32→33, 규칙 D: R=N-35=-2 → T5-1, T4-1
#   → T1=7 / T2=7 / T3=7 / T4=6 / T5=6 (합 33), 경계 타이 없음
# T0 = PRO(config.py) 고정. 탈퇴자(config.RETIRED_PLAYERS)는 랭킹 제외, 시트 행·과거 스코어 보존.
#
# 8/4 재배분(9회, 대회 없음) → 6차(10회) 변경분:
#   윤현호      T3→T2  (기준점수 93.0)
#   이준범      T2→T3  (기준점수 96.2)
#   이호준      T4→T3  (기준점수 96.7)
#   김성환(92)  T3→T4  (기준점수 97.9, 6차 108타)
#   이도영      T5→T4  (기준점수 99.2, 6차 92타)
#   이건희      T4→T5  (기준점수 99.8)
#   박정수      신규→T5 (기준점수 112.0, 첫 대회)
# 이 표는 `python tier_calc.py --emit` 출력을 붙여넣은 것이다. 직접 손대지 말 것.
TIER_MAP = {
    # 티어 0: PRO 5명 (config.py PRO_PLAYERS)
    "김산": 0, "윤석원": 0, "이현우": 0, "이형석": 0, "임승언": 0,
    # 티어 1 (7명)
    "이정승": 1, "이은광": 1, "박병우": 1, "최창환": 1,
    "박진석": 1, "홍경택": 1, "김성환(82)": 1,
    # 티어 2 (7명)
    "나한영": 2, "김대우": 2, "윤준원": 2, "강현수": 2,
    "황인호": 2, "윤현호": 2, "임익준": 2,
    # 티어 3 (7명)
    "조현민": 3, "류성준": 3, "조성태": 3, "강충현": 3,
    "이준범": 3, "권기표": 3, "이호준": 3,
    # 티어 4 (6명)
    "김현섭": 4, "김성환(92)": 4, "최원재": 4, "김준기": 4,
    "이도영": 4, "박헌우": 4,
    # 티어 5 (6명)
    "이건희": 5, "백근호": 5, "옥지엽": 5, "서정권": 5,
    "주홍석": 5, "박정수": 5,
}

NEW_MEMBERS = []
EVENT_COL_NAME = "26' 7차 대회"
EVENT_DATE = "10/18/2026"
TIER_ROUND_COL = "11회"  # 10회 = 6차(9/11), 9회 = 8/4 재배분(대회 없음)
TIER_ROUND_DATE = "10/18/2026"


def connect():
    creds = sa_credentials(SCOPES)
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
        # 누적평균 앞의 빈 버퍼 칼럼(Q) 위치에 삽입
        # → AVERAGE 범위 안쪽 삽입이므로 수식이 자동으로 한 칸 확장됨
        insert_col = avg_col_1based - 1
        sheet.insert_cols([[]], col=insert_col)
        print(f"  열 {insert_col} 위치에 컬럼 삽입 완료")
        time.sleep(2)
        all_values = sheet.get_all_values()
        header_row = all_values[1]
        score_col_1based = insert_col
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

        if name in OCT_2026_SCORES:
            score_updates.append((row_idx, score_col_1based, OCT_2026_SCORES[name]))
            score_done.append(name)

        if tier_col_1based and name in TIER_MAP:
            tier = 0 if name in PRO_PLAYERS else TIER_MAP[name]
            tier_updates.append((row_idx, tier_col_1based, tier))
            tier_done.append(name)

    batch_update_cells(sheet, score_updates)
    print(f"  [OK] 스코어 입력: {len(score_done)}명")
    score_missing = [n for n in OCT_2026_SCORES if n not in score_done]
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

    # ── 컬럼 추가 ──
    if EVENT_COL_NAME not in header_row:
        insert_col = len(header_row) + 1
        col_letter = chr(64 + insert_col)
        batch_update_cells(sheet, [
            (1, insert_col, f"='스코어 집계 (입력)'!{col_letter}2"),
            (2, insert_col, f"=XLOOKUP({col_letter}1,'스코어 집계 (입력)'!2:2,'스코어 집계 (입력)'!3:3)"),
        ])
        print(f"  멤버명단 헤더/날짜 수식 입력 ({col_letter}열)")
        time.sleep(1)
        all_values = sheet.get_all_values()
    else:
        print(f"  멤버명단 '{EVENT_COL_NAME}' 컬럼 이미 존재")

    # ── YES/NO 수식 입력 (스코어 집계에 점수가 있으면 YES, 없으면 NO) ──
    header_row = all_values[0]
    event_col_1based = header_row.index(EVENT_COL_NAME) + 1
    col_letter = chr(64 + event_col_1based)

    # 현재 마지막 데이터 행 + 여유 10행까지 수식 입력
    last_row = max(
        (i + 3 for i, row in enumerate(all_values[2:])
         if len(row) > 1 and row[1].strip()),
        default=10
    ) + 10

    yn_updates = []
    for row_idx in range(3, last_row + 1):
        formula = (
            f"=IF($B{row_idx}=\"\",\"\","
            f"IFERROR(IF(XLOOKUP($B{row_idx},'스코어 집계 (입력)'!$B:$B,"
            f"'스코어 집계 (입력)'!{col_letter}:{col_letter}),\"YES\",\"NO\"),\"NO\"))"
        )
        yn_updates.append((row_idx, event_col_1based, formula))

    batch_update_cells(sheet, yn_updates)
    print(f"  [OK] 멤버명단 YES/NO 수식 입력: 행3~{last_row}")


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
        # 가드: 이 회차 열이 다른 이벤트(예: 대회 없는 재배분)로 이미 채워져 있으면 중단.
        # 2026-09-12 사례 — 9회 열이 8/4 재배분에 쓰인 걸 모르고 6차를 9회로 준비했었음.
        col0 = header.index(TIER_ROUND_COL)
        existing_date = all_values[hi + 1][col0].strip() if len(all_values[hi + 1]) > col0 else ""
        has_values = any(len(r) > col0 and r[col0].strip() for r in all_values[hi + 2:])
        if has_values and existing_date != TIER_ROUND_DATE:
            print(f"  [중단] '{TIER_ROUND_COL}' 열에 이미 값이 있고 날짜({existing_date or '빈칸'})가 "
                  f"TIER_ROUND_DATE({TIER_ROUND_DATE})와 다름. 회차 표(CLAUDE.md '대회 히스토리')를 확인하고 "
                  f"TIER_ROUND_COL을 다음 빈 회차로 바꿀 것.")
            return
        print(f"  '{TIER_ROUND_COL}' 컬럼 이미 존재")
        if not existing_date:
            batch_update_cells(sheet, [(hi + 2, col0 + 1, TIER_ROUND_DATE)])
            print(f"  [{TIER_ROUND_COL}] 날짜 칸 비어 있어 {TIER_ROUND_DATE} 기록")
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
        tier = 0 if name in PRO_PLAYERS else TIER_MAP[name]
        updates.append((i, round_col_1based, tier))
        done.append(name)

    if updates:
        batch_update_cells(sheet, updates)
        print(f"  [OK] 티어 히스토리 입력: {len(done)}명")

    # 시트에 행이 없는 멤버는 새 행으로 자동 추가(마지막 데이터행 다음부터).
    # PRO는 항상 T0라 히스토리 미추적 관례 유지. 신규 멤버는 이번 회차만 채워지고,
    # 다음 회차부터 2개 이상 값이 쌓이면 티어 변동 카드에 정상 반영된다.
    missing = [n for n in TIER_MAP if n not in done and n not in PRO_PLAYERS]
    if missing:
        append_row = len(all_values) + 1
        add_updates, added = [], []
        for name in missing:
            add_updates.append((append_row, name_col + 1, name))
            add_updates.append((append_row, round_col_1based, TIER_MAP[name]))
            added.append(name)
            append_row += 1
        batch_update_cells(sheet, add_updates)
        print(f"  [OK] 티어 히스토리 신규 행 추가: {len(added)}명 {added}")


def main():
    unconfirmed = [n for n in OCT_2026_SCORES if "??" in n]
    if unconfirmed:
        print(f"[중단] 동명이인 미확인: {unconfirmed} — 키를 실제 시트 이름으로 바꾼 뒤 실행하세요.")
        return
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
