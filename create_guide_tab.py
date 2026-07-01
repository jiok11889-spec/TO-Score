"""
구글 시트에 '운영 가이드' 탭 생성 및 내용 기입
오늘 수정한 로직, 구조, 주의사항 기록
"""
import gspread, sys, time
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


GUIDE_ROWS = [
    # [셀값, ...] 순서로 행 구성. 비어있는 칸은 "" 처리
    ["🗂️ TO-Score 운영 가이드", "", "", ""],
    ["최종 정비일: 2026-05-24", "", "", ""],
    ["", "", "", ""],

    # ── 시트 구조 ──
    ["■ 시트 구조 개요", "", "", ""],
    ["시트명", "역할", "주요 참조", "비고"],
    ["스코어 집계 (입력)", "대회 스코어 원본 입력 + 통계 수식", "직접 입력 + 수식", "메인 데이터 시트"],
    ["멤버명단 (자동)", "대회별 참석 YES/NO 자동 도출", "스코어 집계 XLOOKUP", "수식 전용, 값 직접 입력 금지"],
    ["티어정리_신규", "회차별 티어 이력 관리", "직접 입력", "update_scores.py가 기록"],
    ["랭킹", "누적평균 랭킹 표시", "스코어 집계 참조", ""],
    ["매치플레이", "매치플레이 결과", "직접 입력", ""],
    ["", "", "", ""],

    # ── 스코어 집계 열 구조 ──
    ["■ 스코어 집계 (입력) 열 구조", "", "", ""],
    ["열", "내용", "수식 여부", "주의사항"],
    ["A", "No (순번)", "값", ""],
    ["B", "이름", "값", ""],
    ["C ~ M", "25년 대회 스코어 (1차 ~ 4차 + 제주/매치플레이)", "값", ""],
    ["N", "26' 1차 대회 (3/15)", "값", ""],
    ["O", "26' 2차 대회 (4/26)", "값", ""],
    ["P", "26' 3차 대회 (5/17)", "값", ""],
    ["Q", "★ 빈 버퍼 열 ★", "없음", "신규 대회 삽입 위치 기준점 — 절대 삭제/이동 금지"],
    ["R", "누적평균", "=IFERROR(AVERAGE(C:Q),)", "Q열 범위까지 자동 포함"],
    ["S", "누적평균 랭킹", "수식 (R3 기준)", ""],
    ["T", "현재 티어", "값 (update_scores.py 입력)", ""],
    ["U", "티어별 누적평균 랭킹", "수식", ""],
    ["V", "개인 베스트", "=MIN(C:N)", "⚠ 아직 N열까지만 — 추후 Q열로 확장 필요"],
    ["W", "라운드 횟수", "=COUNTIFS(C:N,>0)", "⚠ 아직 N열까지만"],
    ["X", "가장 최근 라운딩 스코어", "수식 (XLOOKUP 역방향)", "⚠ 아직 N열까지만"],
    ["Y", "매치플레이 횟수", "수식", ""],
    ["Z", "매치플레이 평균", "수식", ""],
    ["", "", "", ""],

    # ── 신규 대회 추가 절차 ──
    ["■ 신규 대회 스코어 추가 절차 (update_scores.py)", "", "", ""],
    ["단계", "작업", "파일/위치", "핵심 로직"],
    ["1", "update_scores.py 상단 상수 수정",
     "update_scores.py",
     "MAY_2026_SCORES, TIER_MAP, EVENT_COL_NAME, EVENT_DATE, TIER_ROUND_COL, TIER_ROUND_DATE 업데이트"],
    ["2", "스코어 집계 (입력) 열 삽입",
     "update_score_sheet()",
     "누적평균(R열) -1 위치(=빈 버퍼 Q 앞)에 열 삽입 → AVERAGE(C:Q) 범위 자동 확장"],
    ["3", "멤버명단 (자동) 열 추가",
     "update_member_sheet()",
     "마지막 열 뒤에 추가, 헤더/날짜/YES·NO 모두 XLOOKUP 수식으로 기입 (값 직접 입력 X)"],
    ["4", "티어정리_신규 열 추가",
     "update_tier_history_sheet()",
     "마지막 열 뒤에 회차 컬럼 추가 후 TIER_MAP 기준 입력"],
    ["", "", "", ""],

    # ── 열 삽입 위치 핵심 규칙 ──
    ["■ 신규 대회 열 삽입 위치 — 핵심 규칙", "", "", ""],
    ["항목", "규칙", "이유", ""],
    ["삽입 위치",
     "누적평균(R열) 바로 앞의 빈 버퍼(Q열) 위치에 삽입",
     "AVERAGE(C:Q) 범위 내부에 삽입해야 수식이 자동으로 한 칸 확장됨",
     ""],
    ["코드 기준",
     "avg_col_1based = header.index('누적평균') + 1  →  insert_col = avg_col_1based - 1",
     "누적평균 열 1-based 위치에서 -1 = 빈 버퍼 위치",
     ""],
    ["잘못된 방식",
     "누적평균 바로 앞(insert_col = avg_col_1based)에 삽입",
     "범위 밖 삽입 → AVERAGE 수식이 확장되지 않음",
     ""],
    ["", "", "", ""],

    # ── 수식 범위 기준 ──
    ["■ 수식 행 범위 기준", "", "", ""],
    ["수식", "현재 범위", "확장 시 수정 파일", "비고"],
    ["누적평균 (R열)", "R5:R60", "fix_formulas.py → SCORE_MAX_ROW", "데이터 행: 5행부터 시작"],
    ["누적평균 랭킹 (S열)", "S5:S60, $R$3=COUNTIFS(R5:R60)", "fix_s_formula.py → MAX_ROW", "R3도 같이 수정 필요"],
    ["티어별 랭킹 (U열)", "U5:U60 ($T$5:$T$60)", "fix_formulas.py → SCORE_MAX_ROW", ""],
    ["개인베스트/횟수 등 (V~Z)", "V5:Z60", "fix_formulas.py → SCORE_MAX_ROW", ""],
    ["멤버명단 YES/NO", "행3:행90", "fix_formulas.py → MEMBER_MAX_ROW", ""],
    ["", "", "", ""],

    # ── R3 카운트 수식 주의사항 ──
    ["■ R3 (랭킹 모수) 수식 — 특별 주의", "", "", ""],
    ["위치: 스코어 집계 R3셀", "", "", ""],
    ["현재 수식", "=COUNTIFS(R5:R60,\">0\")", "", ""],
    ["역할", "평균이 있는(>0) 멤버 수를 카운트 → S열 랭킹 계산의 분모(총원)",
     "", ""],
    ["과거 문제",
     "=COUNTIFS(R4:R41,\">0\")로 41행까지만 세어 신규 멤버(42행~)가 누락 → 랭킹 음수 발생",
     "", ""],
    ["향후 멤버가 60명 초과 시",
     "fix_s_formula.py의 MAX_ROW를 올리고 재실행",
     "", ""],
    ["", "", "", ""],

    # ── 멤버명단 수식 구조 ──
    ["■ 멤버명단 (자동) 수식 구조", "", "", ""],
    ["행", "내용", "수식", ""],
    ["1행", "대회명 헤더",
     "='스코어 집계 (입력)'!{열}2  (C열부터)",
     "스코어 집계 2행 참조"],
    ["2행", "대회 날짜",
     "=XLOOKUP({열}1,'스코어 집계 (입력)'!2:2,'스코어 집계 (입력)'!3:3)",
     ""],
    ["3행~", "YES/NO",
     "=IF($B{행}=\"\",\"\",IFERROR(IF(XLOOKUP($B{행},...,{열}:{열}),\"YES\",\"NO\"),\"NO\"))",
     "스코어가 있으면 YES, 없으면 NO / 빈 행은 공백"],
    ["B열", "이름 목록",
     "=SORT(UNIQUE('스코어 집계 (입력)'!B5:B))",
     "스코어 집계 멤버 자동 동기화"],
    ["", "", "", ""],

    # ── 오늘 수정 이력 ──
    ["■ 2026-05-24 수정 이력", "", "", ""],
    ["번호", "문제", "원인", "해결"],
    ["1",
     "26' 2차·3차 대회 열이 1차 대회 옆이 아닌 잘못된 위치에 삽입됨",
     "기존 1차~누적평균 사이 빈 열 때문에 2차·3차가 밀려서 삽입됨",
     "fix_col_order.py로 빈 버퍼 열을 3차 뒤로 이동"],
    ["2",
     "누적평균 수식이 2차·3차 스코어를 포함하지 않음",
     "=AVERAGE(C:N)으로 N열까지만 참조",
     "=IFERROR(AVERAGE(C:Q),)로 Q열까지 확장 (R5:R45 일괄)"],
    ["3",
     "멤버명단 N·O·P열(26년 대회)이 YES/NO 값으로 하드코딩됨",
     "update_scores.py가 수식 대신 값을 직접 기입",
     "fix_formulas.py로 XLOOKUP 수식 교체 + update_scores.py 수정"],
    ["4",
     "스코어 집계 S~Z 수식이 41행까지만 존재 (신규 멤버 누락)",
     "시트 초기 설정 시 41행까지만 수식 작성",
     "fix_formulas.py로 S~Z 행5~60으로 확장, $41 → $60 범위 수정"],
    ["5",
     "S열(누적평균 랭킹)에 '-', 음수, (괄호) 등 이상값 발생",
     "R3가 COUNTIFS(R4:R41)으로 신규 멤버 미포함 → 랭킹 모수 부족 → 음수",
     "fix_s_formula.py: R3를 COUNTIFS(R5:R60)으로 수정 + S수식에 빈행/비숫자 예외처리 추가"],
    ["6",
     "멤버명단 N열 서식 미적용 (O·P열 포스맷 다름)",
     "열 삽입 시 서식 미복사",
     "Sheets API copyPaste PASTE_FORMAT으로 N→O,P 서식 복사"],
    ["", "", "", ""],

    # ── 주요 스크립트 목록 ──
    ["■ 주요 Python 스크립트 목록", "", "", ""],
    ["파일명", "역할", "실행 타이밍", ""],
    ["update_scores.py", "신규 대회 스코어·티어 입력 메인 스크립트", "대회 후 스코어 확정 시", ""],
    ["fix_col_order.py", "스코어 집계 열 순서 교정 (1회성)", "열 순서 틀렸을 때", ""],
    ["fix_formulas.py", "수식 범위 일괄 교정 (S~Z행, 멤버명단)", "멤버 수 크게 늘었을 때", ""],
    ["fix_s_formula.py", "S열 랭킹 수식 + R3 카운트 수식 교정", "랭킹 이상값 발생 시", ""],
    ["read_sheet.py", "스코어 집계 헤더/데이터 확인용", "디버깅 시", ""],
    ["export_tier.py", "티어 현황 내보내기", "필요 시", ""],
    ["server.py", "대시보드 로컬 서버 (port 8000)", "대시보드 볼 때", "python server.py"],
    ["", "", "", ""],

    # ── PRO 선수 관리 ──
    ["■ PRO 선수 관리 (config.py)", "", "", ""],
    ["현재 PRO 목록: 윤석원, 김산, 이현우, 이형석, 임승언", "", "", ""],
    ["PRO는 성적 무관 티어 0 고정 (update_scores.py + server.py 모두 적용)", "", "", ""],
    ["PRO 추가/제거 시 config.py의 PRO_PLAYERS 집합만 수정하면 전체 반영", "", "", ""],
]


def create_guide_tab(spreadsheet):
    # 기존 탭 있으면 삭제 후 재생성
    try:
        existing = spreadsheet.worksheet("운영 가이드")
        spreadsheet.del_worksheet(existing)
        print("  기존 '운영 가이드' 탭 삭제")
    except gspread.exceptions.WorksheetNotFound:
        pass

    sheet = spreadsheet.add_worksheet(title="운영 가이드", rows=200, cols=10)
    print("  '운영 가이드' 탭 생성")
    time.sleep(1)

    # 데이터 기입 (최대 4열)
    padded = []
    for row in GUIDE_ROWS:
        r = list(row) + [""] * (4 - len(row))
        padded.append(r[:4])

    sheet.update(range_name="A1", values=padded, value_input_option="USER_ENTERED")
    print(f"  {len(padded)}행 기입 완료")
    time.sleep(2)

    # ── 서식 설정 ──
    sheet_id = sheet._properties["sheetId"]
    requests = []

    def find_rows(keyword):
        return [i + 1 for i, r in enumerate(padded) if r and r[0].startswith(keyword)]

    # 제목행 (1행): 굵게, 폰트 크게
    requests.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": 4},
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "fontSize": 14,
                           "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.18, "green": 0.35, "blue": 0.60},
        }},
        "fields": "userEnteredFormat(textFormat,backgroundColor)"
    }})

    # 섹션 헤더 (■ 로 시작하는 행): 굵게 + 연한 배경
    section_rows = find_rows("■")
    for r in section_rows:
        requests.append({"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.85, "green": 0.90, "blue": 0.97},
            }},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"
        }})

    # 표 헤더행 (섹션 헤더 바로 다음 행 중 내용 있는 것): 회색 배경
    table_header_rows = []
    for r in section_rows:
        next_r = r + 1
        if next_r <= len(padded) and padded[next_r - 1][0] in [
            "시트명", "열", "단계", "항목", "수식", "행", "번호", "파일명"
        ]:
            table_header_rows.append(next_r)

    for r in table_header_rows:
        requests.append({"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.85, "green": 0.85, "blue": 0.85},
            }},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"
        }})

    # 열 너비 설정
    col_widths = [180, 320, 280, 260]
    for i, w in enumerate(col_widths):
        requests.append({"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w},
            "fields": "pixelSize"
        }})

    # 전체 셀 줄 바꿈 + 상단 정렬
    requests.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": len(padded),
                  "startColumnIndex": 0, "endColumnIndex": 4},
        "cell": {"userEnteredFormat": {
            "wrapStrategy": "WRAP",
            "verticalAlignment": "TOP",
        }},
        "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)"
    }})

    # ★ 빈 버퍼 열 행 강조 (노란 배경)
    for i, row in enumerate(padded):
        if row and "★ 빈 버퍼" in str(row[1]):
            requests.append({"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                          "startColumnIndex": 0, "endColumnIndex": 4},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.6},
                }},
                "fields": "userEnteredFormat(backgroundColor)"
            }})

    spreadsheet.batch_update({"requests": requests})
    print("  서식 적용 완료")


def main():
    print("연결 중...")
    spreadsheet = connect()
    print("[OK] 연결 성공!\n")
    create_guide_tab(spreadsheet)
    print("\n[완료] '운영 가이드' 탭 생성 완료!")


if __name__ == "__main__":
    main()
