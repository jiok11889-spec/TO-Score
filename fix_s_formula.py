"""
S열(누적평균 랭킹) 수식 수정:
1. R3: COUNTIFS 범위 R41 → R60
2. S5:S60: 빈 행/빈 평균 처리 추가, 범위 R60까지 확장
"""
import gspread, sys, time
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

MAX_ROW = 60


def make_s(row):
    """S열 랭킹 수식 생성
    - B열 비어있으면 '' (빈 행)
    - R열이 숫자 아니면 '' (점수 없는 멤버)
    - R열 = 0 이면 '' (프로 등)
    - 그 외 랭킹 계산
    """
    return (
        f"=IF($B{row}=\"\",\"\","
        f"IF(NOT(ISNUMBER(R{row})),\"\","
        f"IF(R{row}=0,\"\","
        f"$R$3+1-(RANK.EQ(R{row},R$5:R${MAX_ROW})"
        f"+COUNTIFS(R$5:R${MAX_ROW},R{row},$B$5:$B${MAX_ROW},\"<\"&$B{row})))))"
    )


def main():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    sheet = spreadsheet.worksheet("스코어 집계 (입력)")

    # 1. R3 수정
    sheet.update(range_name="R3",
                 values=[[f'=COUNTIFS(R5:R{MAX_ROW},">0")']],
                 value_input_option="USER_ENTERED")
    print(f"[OK] R3: =COUNTIFS(R5:R{MAX_ROW},\">0\")")
    time.sleep(1)

    # 수식 샘플 출력 (확인용)
    print("S5 수식 미리보기:", make_s(5))
    print("S42 수식 미리보기:", make_s(42))
    print()

    # 2. S5:S60 일괄 업데이트
    data = [[make_s(r)] for r in range(5, MAX_ROW + 1)]
    data_180 = data[:180]
    sheet.update(range_name=f"S5:S{MAX_ROW}",
                 values=data,
                 value_input_option="USER_ENTERED")
    print(f"[OK] S5:S{MAX_ROW} 수식 업데이트 완료")
    time.sleep(3)

    # 결과 확인
    result = sheet.get("B5:S60", value_render_option="FORMATTED_VALUE")
    print()
    print(f"{'이름':<12} {'누적평균':<10} {'랭킹'}")
    print("-" * 35)
    for row in result:
        name = row[0].strip() if row else ""
        avg  = row[16] if len(row) > 16 else ""
        rank = row[17] if len(row) > 17 else ""
        if name:
            flag = " ← 이상" if rank and rank.startswith("-") else ""
            flag = " ← ERROR" if "#" in str(rank) else flag
            print(f"  {name:<12} {avg:<10} {rank}{flag}")


if __name__ == "__main__":
    main()
