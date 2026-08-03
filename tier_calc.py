"""유지 멤버만으로 티어를 재계산한다.

    python tier_calc.py                # 계산 + 현재 티어와 비교만 (쓰기 없음)
    python tier_calc.py --write        # 스코어 시트의 '현재 티어' 열만 갱신
    python tier_calc.py --emit         # update_scores.py에 붙일 TIER_MAP 블록 출력

탈퇴 멤버(config.RETIRED_PLAYERS)는 시트 행과 과거 스코어를 그대로 둔 채
티어 계산·랭킹에서만 빠진다. 복귀하면 목록에서 빼기만 하면 된다.

규칙 C — 경계 동점자는 전원 상위 티어로 올린다.
규칙 D — 기본 각 티어 7명(합 35). R = N - 35 만큼 T5→T1 순으로 가감하고,
         규칙 C로 생긴 초과 인원은 아래 티어에서 흡수한다.
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import gspread
from google.oauth2.service_account import Credentials

from config import PRO_PLAYERS, RETIRED_PLAYERS

SPREADSHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
JSON_KEY_FILE = "kinetic-horizon-492311-s5-55bd3f137a39.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
SHEET = "스코어 집계 (입력)"
COL_NAME, COL_BASE, COL_TIER = 1, 23, 24   # 0-based
FIRST_ROW = 4                              # 1-based, 데이터 시작 행
TIERS = [1, 2, 3, 4, 5]


def connect():
    creds = Credentials.from_service_account_file(JSON_KEY_FILE, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)


def read_players(ws):
    """(행번호, 이름, 기준스코어, 현재티어) — 유지·비PRO·기준스코어 보유자만."""
    rows = ws.get_all_values()
    out, skipped = [], []
    for i, r in enumerate(rows):
        if i < FIRST_ROW - 1 or len(r) <= COL_TIER or not r[COL_NAME].strip():
            continue
        name = r[COL_NAME].strip()
        try:
            base = float(r[COL_BASE].strip().replace(",", ""))
        except ValueError:
            base = None
        rec = (i + 1, name, base, r[COL_TIER].strip())
        if name in RETIRED_PLAYERS:
            skipped.append((name, "탈퇴"))
        elif name in PRO_PLAYERS:
            skipped.append((name, "PRO"))
        elif base is None:
            skipped.append((name, "기준스코어 없음"))
        else:
            out.append(rec)
    return sorted(out, key=lambda x: x[2]), skipped


def quotas(n):
    """규칙 D — 기본 7명씩, R만큼 T5부터 역순으로 가감."""
    q = {t: 7 for t in TIERS}
    r = n - 35
    order = [5, 4, 3, 2, 1]
    i = 0
    while r != 0:
        t = order[i % 5]
        q[t] += 1 if r > 0 else -1
        r += -1 if r > 0 else 1
        i += 1
    return q


def assign(rank, q):
    """규칙 C(경계 동점 → 상위 티어) 적용 후 규칙 D 잔여 인원 흡수."""
    result, idx = {}, 0
    for pos, t in enumerate(TIERS):
        if idx >= len(rank):
            break
        if t == TIERS[-1]:
            take = len(rank) - idx
        else:
            take = min(q[t], len(rank) - idx)
            # 경계 동점자는 전원 상위 티어로
            while idx + take < len(rank) and take > 0 and \
                    rank[idx + take][2] == rank[idx + take - 1][2]:
                take += 1
            # 아래 티어에 최소 1명씩은 남긴다
            left = len(rank) - (idx + take)
            need = len(TIERS) - pos - 1
            if left < need:
                take -= (need - left)
        for p in rank[idx:idx + take]:
            result[p[1]] = t
        idx += take
    return result


def main():
    ss = connect()
    ws = ss.worksheet(SHEET)
    rank, skipped = read_players(ws)
    n = len(rank)
    q = quotas(n)

    print(f"랭킹 대상 N = {n}명  (탈퇴·PRO·기준스코어 없음 제외)")
    print(f"규칙 D 정원: " + " / ".join(f"T{t}={q[t]}" for t in TIERS) +
          f"  (합 {sum(q.values())})\n")

    new = assign(rank, q)
    counts = {t: sum(1 for v in new.values() if v == t) for t in TIERS}
    print("확정 인원: " + " / ".join(f"T{t}={counts[t]}" for t in TIERS) +
          f"  (합 {sum(counts.values())})\n")

    changes = []
    for row, name, base, cur in rank:
        t = new[name]
        mark = ""
        if cur.isdigit() and int(cur) != t:
            mark = f"   ← {cur}티어에서 {'상승' if t < int(cur) else '하락'}"
            changes.append((name, int(cur), t))
        print(f"  T{t}  {name:12} 기준 {base:6.2f}{mark}")

    print(f"\n변동 {len(changes)}명")
    for name, a, b in changes:
        print(f"  {name:12} {a} → {b}")

    print("\n계산 제외:")
    for name, why in skipped:
        print(f"  {name:12} ({why})")

    if "--emit" in sys.argv:
        print("\n" + "=" * 50)
        print("TIER_MAP = {")
        pros = ", ".join(f'"{p}": 0' for p in sorted(PRO_PLAYERS))
        print(f"    # 티어 0: PRO {len(PRO_PLAYERS)}명 (config.py PRO_PLAYERS)")
        print(f"    {pros},")
        for t in TIERS:
            names = [nm for nm, v in new.items() if v == t]
            print(f"    # 티어 {t} ({len(names)}명)")
            for j in range(0, len(names), 4):
                print("    " + " ".join(f'"{nm}": {t},' for nm in names[j:j + 4]))
        print("}")

    if "--write" in sys.argv:
        cells = []
        for row, name, base, cur in rank:
            cells.append(gspread.Cell(row, COL_TIER + 1, new[name]))
        ws.update_cells(cells)
        print(f"\n'현재 티어' 열 {len(cells)}칸 갱신 완료")
    else:
        print("\n(쓰기 없음) --write 로 시트의 '현재 티어' 열만 갱신")


if __name__ == "__main__":
    main()
