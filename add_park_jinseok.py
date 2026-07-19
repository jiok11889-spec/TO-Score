"""
일회성 스크립트: 박진석(원래 멤버, TIER_MAP 미등록) 46행 버퍼 행 채우기
- A열(No), T열(누적평균) 수식 46~60행 보정 (버퍼 행에 formula 누락돼 있었음)
- B46=박진석, R46(26' 5차 대회 스코어)=83
"""
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def main():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SPREADSHEET_ID)
    sheet = sh.worksheet("스코어 집계 (입력)")

    data = []
    # 버퍼 행(46~60) A열/T열 수식 보정 (인접 정상 행 45번 패턴 복사)
    for r in range(46, 61):
        data.append({"range": rowcol_to_a1(r, 1), "values": [["=ROW()-4"]]})
        data.append({"range": rowcol_to_a1(r, 20), "values": [[f"=IFERROR(AVERAGE(C{r}:S{r}),)"]]})

    # 박진석 이름 + 5차 대회 스코어(83)
    data.append({"range": rowcol_to_a1(46, 2), "values": [["박진석"]]})
    data.append({"range": rowcol_to_a1(46, 18), "values": [[83]]})

    sheet.batch_update(data, value_input_option="USER_ENTERED")
    print(f"[OK] 업데이트 완료: {len(data)}건 (A/T열 수식 46~60행 보정 + 박진석 46행 입력)")


if __name__ == "__main__":
    main()
