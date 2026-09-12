@echo off
echo 서버 시작 중...
start /min "" python "%~dp0server.py"
timeout /t 2 /nobreak > nul
start http://localhost:8000
exit
