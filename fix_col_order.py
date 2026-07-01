"""
스코어 집계 (입력) 시트의 칼럼 순서 수정
현재: 26' 1차 대회 | [빈칼럼] | 26' 2차 대회 | 26' 3차 대회 | 누적평균
목표: 26' 1차 대회 | 26' 2차 대회 | 26' 3차 대회 | [빈칼럼] | 누적평균
"""
import gspread
import sys
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def main():
    print("연결 중...")
    spreadsheet = connect()
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")
    sheet_id = sheet._properties["sheetId"]

    all_values = sheet.get_all_values()
    header = all_values[1]

    print("현재 헤더 확인:")
    for i, h in enumerate(header):
        if h or (10 <= i <= 22):
            print(f"  [{i}] {h!r}")

    one_idx = header.index("26' 1차 대회")
    two_idx = header.index("26' 2차 대회")
    three_idx = header.index("26' 3차 대회")

    print(f"\n26' 1차 위치(0-based): {one_idx}")
    print(f"26' 2차 위치(0-based): {two_idx}")
    print(f"26' 3차 위치(0-based): {three_idx}")

    # 1차 바로 다음 칼럼이 비어있는지 확인 (사이에 있는 경우)
    # 2차와 1차 사이에 빈 칼럼이 있는 경우: gap은 1차~2차 사이
    # 2차와 3차 사이에 빈 칼럼이 있는 경우: gap은 2차~3차 사이
    if two_idx == one_idx + 2 and header[one_idx + 1] == "":
        gap_idx = one_idx + 1
        print(f"\n빈 칼럼(0-based {gap_idx})이 1차~2차 사이에 있음 → 3차 뒤로 이동")
    elif three_idx == two_idx + 2 and header[two_idx + 1] == "":
        gap_idx = two_idx + 1
        print(f"\n빈 칼럼(0-based {gap_idx})이 2차~3차 사이에 있음 → 3차 뒤로 이동")
    elif two_idx == one_idx + 1 and three_idx == two_idx + 1:
        print("\n[OK] 이미 올바른 순서입니다: 1차 | 2차 | 3차")
        return
    else:
        print(f"\n[!!] 예상한 구조와 다릅니다.")
        print(f"  1차={one_idx}, 2차={two_idx}, 3차={three_idx}")
        return

    # moveDimension API: destinationIndex는 원본 레이아웃 기준, "이 인덱스 앞에 삽입"
    # source가 destination보다 앞인 경우: 이동 후 위치 = three_idx + 1 앞(=누적평균 앞)
    destination = three_idx + 1

    body = {
        "requests": [{
            "moveDimension": {
                "source": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": gap_idx,
                    "endIndex": gap_idx + 1,
                },
                "destinationIndex": destination,
            }
        }]
    }

    print(f"  moveDimension: source startIndex={gap_idx}, destinationIndex={destination} (원본 기준 {destination}번 앞에 삽입)")
    spreadsheet.batch_update(body)
    print("\n[OK] 칼럼 이동 완료!")

    # 결과 확인
    all_values = sheet.get_all_values()
    header = all_values[1]
    print("\n수정 후 헤더:")
    for i, h in enumerate(header[one_idx:one_idx + 6], start=one_idx):
        print(f"  [{i}] {h!r}")

if __name__ == "__main__":
    main()
