"""
현재 스프레드시트의 셀 서식(배경색) 읽기
"""
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def main():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)

    # 시트 메타데이터 및 서식 읽기
    result = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        includeGridData=True,
        ranges=[
            "티어정리_신규 (작업중)!A1:L71",
            "스코어 집계 (입력)!A1:Y5",
        ]
    ).execute()

    output = {}
    for sheet in result.get("sheets", []):
        title = sheet["properties"]["title"]
        output[title] = []
        for grid in sheet.get("data", []):
            for row_idx, row_data in enumerate(grid.get("rowData", [])):
                row_info = []
                for col_idx, cell in enumerate(row_data.get("values", [])):
                    bg = cell.get("effectiveFormat", {}).get("backgroundColor", {})
                    val = cell.get("formattedValue", "")
                    r = bg.get("red", 1.0)
                    g = bg.get("green", 1.0)
                    b = bg.get("blue", 1.0)
                    # 흰색(1,1,1)이 아닌 셀만 기록
                    if not (r > 0.98 and g > 0.98 and b > 0.98):
                        row_info.append({
                            "row": row_idx, "col": col_idx,
                            "val": val,
                            "bg": {"r": round(r,3), "g": round(g,3), "b": round(b,3)}
                        })
                if row_info:
                    output[title].extend(row_info)

    with open("format_data.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("format_data.json 저장 완료")

if __name__ == "__main__":
    main()
