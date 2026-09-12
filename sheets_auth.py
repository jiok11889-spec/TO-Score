"""구글 시트 서비스 계정 자격증명을 한 곳에서 만든다.

우선순위:
  1. 환경변수 GOOGLE_SA_JSON      — 키 파일의 JSON 내용 전체 (클라우드/모바일 세션용 secret)
  2. 환경변수 GOOGLE_SA_JSON_FILE — 키 파일 경로
  3. 저장소 루트의 키 파일        — PC 로컬 (gitignore 되어 있어 클라우드엔 없음)

클라우드(claude.ai/code) 세션은 git에 있는 파일만 받으므로 키 파일이 없다.
그 환경에서는 1번 secret을 등록해야 update_scores.py / tier_calc.py / audit_scores.py가 돈다.
"""
import json
import os

from google.oauth2.service_account import Credentials

KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "kinetic-horizon-492311-s5-55bd3f137a39.json")


def credentials(scopes):
    raw = os.environ.get("GOOGLE_SA_JSON", "").strip()
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=scopes)
    path = os.environ.get("GOOGLE_SA_JSON_FILE") or KEY_FILE
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=scopes)
    raise SystemExit(
        "[ERR] 구글 시트 키를 찾을 수 없습니다.\n"
        "  PC: 저장소 루트에 kinetic-horizon-*.json 키 파일이 있어야 합니다.\n"
        "  클라우드/모바일 세션: claude.ai/code 환경 설정에서 환경변수 GOOGLE_SA_JSON에\n"
        "  키 파일 내용(JSON 전체)을 secret으로 등록하세요. (CLAUDE.md '모바일에서 끝까지 진행하기' 참고)"
    )
