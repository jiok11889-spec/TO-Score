"""
멤버 명단 정리 스크립트
- 탈회 27명 삭제
- 신규 6명 추가 (이미 있는 경우 스킵)
대상 시트: 스코어 집계 (입력), 멤버명단 (자동)
"""
import gspread
from google.oauth2.service_account import Credentials
import time
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

WITHDRAW_MEMBERS = {
    "박세환", "김진호", "박상용", "이민재", "최정훈", "권한해", "고영승", "고태현",
    "박승현", "박준상", "한상윤", "손정모", "오지호", "이수빈", "이종화", "조승현",
    "최창규", "이대로", "노용현", "신태범", "신재민", "강승원", "강명신", "박문석",
    "최성호", "박계홍", "노민일",
}

NEW_MEMBERS = ["윤준원", "김현섭", "강현수", "이은광", "최창환", "주홍석"]

FINAL_MEMBERS_ORDER = [
    "조현민", "권기표", "옥지엽", "나한영", "류성준", "황인호", "김성환(92)",
    "최원재", "임익준", "김준기", "김산", "김성환", "김지수", "백근호", "손원우",
    "이현우", "이건희", "이도영", "이호준", "황호연", "조성태", "이준범", "서정권",
    "김수민", "이형석", "임승언", "김대우", "이정승", "윤석원", "박헌우", "박병우",
    "홍경택", "윤현호", "강충현", "김서래",
    "윤준원", "김현섭", "강현수", "이은광", "최창환", "주홍석",
]


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def delete_rows_for_members(sheet, name_col_idx, data_start_idx, withdraw_set):
    """탈회 멤버 행 삭제 (역순으로 삭제해서 인덱스 밀림 방지)"""
    all_values = sheet.get_all_values()
    rows_to_delete = []

    for i, row in enumerate(all_values[data_start_idx:], start=data_start_idx + 1):
        if len(row) <= name_col_idx:
            continue
        name = row[name_col_idx].strip()
        if name in withdraw_set:
            rows_to_delete.append((i, name))

    if not rows_to_delete:
        print("  삭제할 탈회 멤버 없음")
        return 0

    for row_idx, name in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx)
        print(f"  삭제: {name} (행 {row_idx})")
        time.sleep(0.5)

    print(f"  [OK] {len(rows_to_delete)}명 삭제 완료")
    return len(rows_to_delete)


def add_new_members(sheet, name_col_idx, data_start_idx, new_members):
    """신규 멤버 추가 (이미 있으면 스킵)"""
    all_values = sheet.get_all_values()
    existing = {
        row[name_col_idx].strip()
        for row in all_values[data_start_idx:]
        if len(row) > name_col_idx and row[name_col_idx].strip()
    }
    header_len = max(len(row) for row in all_values[:data_start_idx + 1] if row)

    added = []
    for name in new_members:
        if name in existing:
            print(f"  이미 존재: {name} (스킵)")
            continue
        next_no = len(existing) + 1
        new_row = [""] * header_len
        new_row[0] = str(next_no)
        new_row[name_col_idx] = name
        sheet.append_row(new_row, value_input_option="USER_ENTERED")
        print(f"  추가: {name}")
        existing.add(name)
        added.append(name)
        time.sleep(1)

    print(f"  [OK] {len(added)}명 추가 완료")
    return len(added)


def update_score_sheet(spreadsheet):
    print("=== 스코어 집계 (입력) 시트 ===")
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")
    all_values = sheet.get_all_values()

    # 헤더: 2행 (index 1), 데이터 시작: 5행 (index 4)
    header_row = all_values[1]
    name_col_idx = header_row.index("이름")

    print("  [탈회 멤버 삭제]")
    deleted = delete_rows_for_members(sheet, name_col_idx, 4, WITHDRAW_MEMBERS)

    time.sleep(2)

    print("  [신규 멤버 추가]")
    added = add_new_members(sheet, name_col_idx, 4, NEW_MEMBERS)

    return deleted, added


def update_member_sheet(spreadsheet):
    print("\n=== 멤버명단 (자동) 시트 ===")
    sheet = spreadsheet.worksheet("멤버명단 (자동)")
    all_values = sheet.get_all_values()

    # 헤더: 1행 (index 0), 날짜행: 2행 (index 1), 데이터 시작: 3행 (index 2)
    header_row = all_values[0]
    name_col_idx = header_row.index("이름") if "이름" in header_row else 1

    print("  [탈회 멤버 삭제]")
    deleted = delete_rows_for_members(sheet, name_col_idx, 2, WITHDRAW_MEMBERS)

    time.sleep(2)

    print("  [신규 멤버 추가]")
    added = add_new_members(sheet, name_col_idx, 2, NEW_MEMBERS)

    return deleted, added


def main():
    print("Google Sheets 연결 중...")
    try:
        spreadsheet = connect()
        print("[OK] 연결 성공!\n")
    except Exception as e:
        print(f"[ERR] 연결 실패: {e}")
        return

    s_del, s_add = update_score_sheet(spreadsheet)
    time.sleep(3)
    m_del, m_add = update_member_sheet(spreadsheet)

    print("\n============================")
    print(f"[완료] 멤버 명단 업데이트 완료!")
    print(f"  스코어 시트: -{s_del}명 삭제, +{s_add}명 추가")
    print(f"  멤버명단 시트: -{m_del}명 삭제, +{m_add}명 추가")
    print(f"  최종 인원: 41명")
    print("============================")


if __name__ == "__main__":
    main()
