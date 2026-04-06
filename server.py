"""
티오방 스코어 대시보드 서버
실행: python server.py
접속: http://localhost:8000
"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.request
import json
import csv
import io
import os

SHEET_ID = "16Ay7f7lhccjdfKhb-Fe1U6DVicAVq0dqS3kEzusgXg4"
GID = {
    "score":     "1112772180",
    "ranking":   "1298037096",
    "matchplay": "585545993",
    "member":    "300187863",
    "tier_new":  "1117395163",
}

def fetch_csv_rows(gid):
    """CSV를 제대로 파싱해서 2D 배열로 반환"""
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15) as res:
        raw = res.read().decode("utf-8")
    reader = csv.reader(io.StringIO(raw))
    return [row for row in reader]

def find_header_row(rows, key="이름"):
    """'이름' 컬럼이 있는 행 인덱스 반환"""
    for i, row in enumerate(rows):
        if key in row:
            return i
    return 0

def rows_to_dicts(rows, header_idx, skip_extra=0):
    """헤더 행 이후 데이터를 dict 리스트로 변환"""
    headers = [h.strip() for h in rows[header_idx]]
    result = []
    for row in rows[header_idx + 1 + skip_extra:]:
        if not any(cell.strip() for cell in row):
            continue  # 완전히 빈 행 스킵
        d = {}
        for i, h in enumerate(headers):
            d[h] = row[i].strip() if i < len(row) else ""
        # 이름이 없는 행(날짜행, 서브헤더행) 스킵
        name = d.get("이름", "").strip()
        if not name:
            continue
        result.append(d)
    return result

def parse_score_sheet(rows):
    hi = find_header_row(rows)
    return rows_to_dicts(rows, hi)

def parse_ranking_sheet(rows):
    hi = find_header_row(rows)
    return rows_to_dicts(rows, hi)

def parse_matchplay_sheet(rows):
    hi = find_header_row(rows)
    return rows_to_dicts(rows, hi)

def parse_member_sheet(rows):
    hi = find_header_row(rows)
    return rows_to_dicts(rows, hi, skip_extra=1)  # 날짜 행 스킵

def parse_tier_history(rows):
    """티어정리_신규: 마지막 2개 회차 비교 → 티어 변동 목록 반환"""
    hi = find_header_row(rows)
    if hi < 0:
        return []
    header = rows[hi]
    members = rows_to_dicts(rows, hi, skip_extra=1)

    data_col_names = [h for h in header if h and h not in ("No", "이름")]

    result = []
    for m in members:
        name = m.get("이름", "").strip()
        if not name:
            continue
        tiers = [m.get(c, "").strip() for c in data_col_names
                 if m.get(c, "").strip() not in ("", "-")]
        if len(tiers) >= 2:
            result.append({
                "이름": name,
                "이전티어": tiers[-2],
                "현재티어": tiers[-1],
            })
    return result

def get_all_data():
    score_rows     = fetch_csv_rows(GID["score"])
    ranking_rows   = fetch_csv_rows(GID["ranking"])
    matchplay_rows = fetch_csv_rows(GID["matchplay"])
    member_rows    = fetch_csv_rows(GID["member"])
    tier_new_rows  = fetch_csv_rows(GID["tier_new"])

    return {
        "score":        parse_score_sheet(score_rows),
        "ranking":      parse_ranking_sheet(ranking_rows),
        "matchplay":    parse_matchplay_sheet(matchplay_rows),
        "member":       parse_member_sheet(member_rows),
        "tier_history": parse_tier_history(tier_new_rows),
    }

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path == "/api/data":
            try:
                data = get_all_data()
                body = json.dumps(data, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        path = self.path.split("?")[0]
        if path == "/" or path == "":
            path = "/dashboard.html"

        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), path.lstrip("/"))
        if os.path.isfile(file_path):
            ext = file_path.rsplit(".", 1)[-1]
            mime = {"html": "text/html", "js": "text/javascript",
                    "css": "text/css"}.get(ext, "text/plain")
            with open(file_path, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", f"{mime}; charset=utf-8")
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    print(f"Server started: http://localhost:{port}")
    HTTPServer(("", port), Handler).serve_forever()
