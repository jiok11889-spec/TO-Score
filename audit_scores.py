"""스코어 집계 시트 정합성 검증 — 표시값 vs 실제 스코어로 계산한 값.

    python audit_scores.py          # 검증만 (읽기 전용)
    python audit_scores.py --fix    # 통계 열 수식의 참조 범위를 고쳐 다시 씀

배경(2026-08-04): 개인베스트·라운드횟수·가장최근라운딩 수식이 `C:N`으로 굳어 있어
26년 2~5차 대회(O~R열)를 통계에서 통째로 빠뜨렸다. 신규 대회 열은 버퍼(S) 앞에
삽입되므로, 범위가 N에서 끝나면 새 대회가 들어와도 자동으로 늘지 않는다.
→ 범위를 버퍼 열까지(C:S) 잡아야 대회가 추가돼도 따라 늘어난다.
누적평균(C:S)·26평균(N:S)은 원래 맞아서 기준스코어·티어에는 영향이 없었다.

대회를 추가한 뒤에는 이 검증을 한 번 돌려 0건인지 확인할 것.
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
SHEET = "스코어 집계 (입력)"

FIRST_DATA_ROW = 4          # 1-based
COL_NAME = 1                # 0-based
C_AVG, C_BEST, C_CNT, C_LAST, C_MCNT, C_MAVG = 19, 26, 27, 28, 29, 30


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except (ValueError, TypeError):
        return None


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)


def score_columns(hdr):
    """대회 열(0-based)과 버퍼 열 문자를 돌려준다."""
    cols = [c for c in range(2, C_AVG) if str(hdr[c]).strip()]
    buffer_idx = C_AVG - 1          # 누적평균 바로 앞 = 빈 버퍼 열
    return cols, chr(ord("A") + buffer_idx)


def audit(vals, cols):
    issues = {"last": [], "best": [], "cnt": [], "avg": []}
    for i, r in enumerate(vals):
        if i < FIRST_DATA_ROW - 1 or len(r) <= C_LAST or not r[COL_NAME].strip():
            continue
        name = r[COL_NAME].strip()
        sc = [(c, num(r[c])) for c in cols if c < len(r) and num(r[c]) is not None]
        if not sc:
            continue
        real = {"best": min(v for _, v in sc), "cnt": len(sc), "last": sc[-1][1],
                "avg": round(sum(v for _, v in sc) / len(sc), 1)}
        shown = {"best": num(r[C_BEST]), "cnt": num(r[C_CNT]),
                 "last": num(r[C_LAST]), "avg": num(r[C_AVG])}
        for k in issues:
            s = shown[k]
            if s is None:
                continue
            if (abs(s - real[k]) > 0.05) if k == "avg" else (s != real[k]):
                issues[k].append((name, s, real[k]))
    return issues


def formulas(r, buf):
    return {
        C_BEST: f'=IF(COUNT(C{r}:{buf}{r})=0,"",MIN(C{r}:{buf}{r}))',
        C_CNT: f'=COUNTIFS(C{r}:{buf}{r},">0")',
        C_LAST: (f'=ARRAY_CONSTRAIN(ARRAYFORMULA(XLOOKUP(TRUE,ISNUMBER(C{r}:{buf}{r}),'
                 f'C{r}:{buf}{r},"",0,-1)), 1, 1)'),
        C_MCNT: f'=COUNTIFS($C$2:${buf}$2,"*매치플레이*",C{r}:{buf}{r},"<>")',
        C_MAVG: (f'=IFERROR(AVERAGE(FILTER(C{r}:{buf}{r},ISNUMBER(C{r}:{buf}{r}),'
                 f'REGEXMATCH($C$2:${buf}$2,"매치플레이"))),"")'),
    }


def main():
    ws = connect().worksheet(SHEET)
    vals = ws.get_all_values()
    hdr = vals[1]
    cols, buf = score_columns(hdr)

    print("대회 열: " + ", ".join(f"{hdr[c]}({chr(65+c)})" for c in cols))
    print(f"통계 수식이 잡아야 할 범위: C:{buf}\n")

    issues = audit(vals, cols)
    total = sum(len(v) for v in issues.values())
    for key, label in [("last", "가장최근라운딩"), ("best", "개인베스트"),
                       ("cnt", "라운드횟수"), ("avg", "누적평균")]:
        rows = issues[key]
        print(f"■ {label}: 불일치 {len(rows)}명")
        for n, s, real in rows:
            print(f"    {n:14} 표시 {s:g} → 실제 {real:g}")
    print(f"\n총 불일치 {total}건")

    if total == 0:
        print("정합성 이상 없음.")
        return
    if "--fix" not in sys.argv:
        print("\n(읽기 전용) --fix 로 통계 열 수식의 참조 범위를 고칩니다.")
        return

    cells = []
    for i, r in enumerate(vals):
        if i < FIRST_DATA_ROW - 1 or not r[COL_NAME].strip():
            continue
        for c, f in formulas(i + 1, buf).items():
            cells.append(gspread.Cell(i + 1, c + 1, f))
    ws.update_cells(cells, value_input_option="USER_ENTERED")
    print(f"\n수식 {len(cells)}칸 갱신 완료 — 다시 실행해 0건인지 확인하세요.")


if __name__ == "__main__":
    main()
