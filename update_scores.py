import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import time
import sys
from config import PRO_PLAYERS

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 26년 7월 19일 5차 대회 스코어 — 세현CC (참가 20명 중 게스트 2명 제외 → 18명 기록)
# 실타수(스코어 점수) 기준, 신페리오 net 아님. 김성환* = 김성환(92)
# ※ 게스트 김태경(72)·박태민(102) 제외 → 인원(규칙 D N)에 미반영.
# ※ 박진석(83)은 원래 멤버지만 TIER_MAP 미등록 → 시트에 행 있으면 스코어 기록됨.
#    행 없으면 [!!] 미입력 → 시트에 행 추가 후 재실행. 기준스코어 생기면 N=37로 티어 재배분.
JUL_2026_SCORES = {
    "박진석":      83,
    "김성환":      84,
    "윤준원":      85,
    "김현섭":      86,
    "강현수":      88,
    "홍경택":      88,
    "권기표":      89,
    "황인호":      89,
    "이호준":      90,
    "강충현":      91,
    "나한영":      93,
    "조현민":      94,
    "류성준":      94,
    "김성환(92)":  94,
    "최원재":      94,
    "임익준":      96,
    "이건희":     106,
    "이도영":     107,
}

# 5차 대회(7/19) 이후 기준스코어(누적평균×0.5 + 26년평균×0.5) 기반 티어 재정렬
# 박진석(원래 멤버, 46행 버퍼행 신규 등록) 합류로 N=36→37, 규칙 D 재계산: R=N-35=2
#   → T5+1, T4+1 순 적용이 원칙이나, 랭킹 21/22위 조성태·강충현 동점(96.0)이
#     T3/T4 경계에 걸려 규칙 C(경계 타이→상위 티어 배치)가 먼저 적용됨
#     → 실제 실현 인원: T1=7 / T2=7 / T3=8(타이 흡수) / T4=7 / T5=8 (합 37, 규칙 D가 흡수)
# T0 = PRO(config.py) + 기록없는 오지호 고정
# T1~T5 = 비PRO 기준스코어 오름차순 배분 (기록없는 멤버는 현 티어 유지)
#
# 4차→5차 변경분 (기준스코어 재계산 결과):
#   강충현  T4→T3  (기준점수 96.0, 랭킹 22위 — 21위 조성태와 동점, 규칙C로 T3 동반 배치)
#   김지수  T3→T4  (기준점수 96.5, 랭킹 23위)
#   이호준  T5→T4  (기준점수 96.7, 랭킹 24위)
#   김현섭  T5→T4  (기준점수 97.6, 랭킹 25위)
#   최원재  T5→T4  (기준점수 98.8, 랭킹 27위)
#   이건희  T4→T5  (기준점수 99.8, 랭킹 30위)
#   황호연  T4→T5  (기준점수 100.3, 랭킹 31위)
#   이도영  T4→T5  (기준점수 101.2, 랭킹 32위)
#   박진석  신규→T1  (기준점수 83.0, 랭킹 4위 — 이은광과 동점이나 T1 밴드 내부라 규칙C 무관)
#   나한영  T1→T2   (기준점수 85.4, 랭킹 8위 — 박진석 합류로 T1 밴드 밀려남)
#   이준범  T2→T3   (기준점수 92.5, 랭킹 15위 — 나한영 밀림에 연쇄, T2 밴드 밀려남)
TIER_MAP = {
    # 티어 0: PRO 5명 (config.py PRO_PLAYERS)
    "김산":        0, "이형석":    0, "이현우":     0, "윤석원":    0, "임승언":    0,
    # 티어 1 (7명): 기준스코어 랭킹 1~7위
    "이정승":      1, "최창환":    1, "이은광":     1, "박진석":    1,
    "김수민":      1, "홍경택":    1, "박병우":     1,
    # 티어 2 (7명): 기준스코어 랭킹 8~14위
    "나한영":      2, "김대우":    2, "윤준원":     2, "김성환":    2,
    "강현수":      2, "황인호":    2, "임익준":     2,
    # 티어 3 (8명): 기준스코어 랭킹 15~22위 (21/22위 동점 규칙C로 1명 추가 흡수)
    "이준범":      3, "윤현호":    3, "김성환(92)": 3, "조현민":    3,
    "류성준":      3, "권기표":    3, "조성태":     3, "강충현":    3,
    # 티어 4 (7명): 기준스코어 랭킹 23~29위
    "김지수":      4, "이호준":    4, "김현섭":     4, "박헌우":    4,
    "최원재":      4, "김준기":    4, "손원우":     4,
    # 티어 5 (8명): 기준스코어 랭킹 30~37위
    "이건희":      5, "황호연":    5, "이도영":     5, "백근호":    5,
    "김서래":      5, "옥지엽":    5, "서정권":     5, "주홍석":    5,
}

NEW_MEMBERS = []
EVENT_COL_NAME = "26' 5차 대회"
EVENT_DATE = "7/19/2026"
TIER_ROUND_COL = "8회"
TIER_ROUND_DATE = "7/19/2026"


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

        if name in JUL_2026_SCORES:
            score_updates.append((row_idx, score_col_1based, JUL_2026_SCORES[name]))
            score_done.append(name)

        if tier_col_1based and name in TIER_MAP:
            tier = 0 if name in PRO_PLAYERS else TIER_MAP[name]
            tier_updates.append((row_idx, tier_col_1based, tier))
            tier_done.append(name)

    batch_update_cells(sheet, score_updates)
    print(f"  [OK] 스코어 입력: {len(score_done)}명")
    score_missing = [n for n in JUL_2026_SCORES if n not in score_done]
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
