"""구글 시트 서비스 계정 자격증명을 한 곳에서 만든다.

우선순위:
  1. 환경변수 GOOGLE_SA_JSON_B64  — 키 파일을 base64로 한 줄 인코딩한 값 (클라우드/모바일 세션용 secret, 권장)
  2. 환경변수 GOOGLE_SA_JSON      — 키 파일의 JSON 내용 전체 (따옴표·줄바꿈 때문에 입력칸에서 잘릴 수 있음)
  3. 환경변수 GOOGLE_SA_JSON_FILE — 키 파일 경로
  4. 저장소 루트의 키 파일        — PC 로컬 (gitignore 되어 있어 클라우드엔 없음)

base64 값 만들기 (PC, 저장소 루트에서):
  PowerShell: [Convert]::ToBase64String([IO.File]::ReadAllBytes("kinetic-horizon-492311-s5-55bd3f137a39.json")) | Set-Clipboard
  bash:       base64 -w0 kinetic-horizon-492311-s5-55bd3f137a39.json | pbcopy

클라우드(claude.ai/code) 세션은 git에 있는 파일만 받으므로 키 파일이 없다.
그 환경에서는 1번 secret을 등록해야 update_scores.py / tier_calc.py / audit_scores.py가 돈다.
"""
import base64
import binascii
import json
import os

from google.oauth2.service_account import Credentials

KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "kinetic-horizon-492311-s5-55bd3f137a39.json")


def _info_from_env():
    """환경변수에서 키 dict를 얻는다. 없으면 None, 있는데 깨졌으면 원인을 담아 SystemExit."""
    b64 = os.environ.get("GOOGLE_SA_JSON_B64", "").strip()
    if b64:
        try:
            raw = base64.b64decode(b64, validate=True).decode("utf-8")
            return json.loads(raw)
        except (binascii.Error, UnicodeDecodeError, json.JSONDecodeError) as e:
            raise SystemExit(
                f"[ERR] GOOGLE_SA_JSON_B64 값이 깨져 있습니다 ({type(e).__name__}). 길이 {len(b64)}자.\n"
                "  PC에서 다시 만들어 통째로 붙여넣으세요:\n"
                '  [Convert]::ToBase64String([IO.File]::ReadAllBytes("kinetic-horizon-492311-s5-55bd3f137a39.json")) | Set-Clipboard'
            )
    raw = os.environ.get("GOOGLE_SA_JSON", "").strip()
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            raise SystemExit(
                f"[ERR] GOOGLE_SA_JSON 값이 JSON으로 안 읽힙니다 (길이 {len(raw)}자, {e.msg} at {e.pos}).\n"
                "  보통 2,000자가 넘어야 정상 — 입력칸에서 따옴표/줄바꿈에 잘린 경우입니다.\n"
                "  대신 GOOGLE_SA_JSON_B64 에 base64 한 줄로 등록하세요 (sheets_auth.py 상단 참고)."
            )
    return None


def credentials(scopes):
    info = _info_from_env()
    if info is not None:
        return Credentials.from_service_account_info(info, scopes=scopes)
    path = os.environ.get("GOOGLE_SA_JSON_FILE") or KEY_FILE
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=scopes)
    raise SystemExit(
        "[ERR] 구글 시트 키를 찾을 수 없습니다.\n"
        "  PC: 저장소 루트에 kinetic-horizon-*.json 키 파일이 있어야 합니다.\n"
        "  클라우드/모바일 세션: claude.ai/code 환경 설정에서 환경변수 GOOGLE_SA_JSON_B64에\n"
        "  키 파일을 base64 한 줄로 등록하세요. (CLAUDE.md '모바일에서 끝까지 진행하기' 참고)"
    )
