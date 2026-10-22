@echo off
cd /d "%~dp0"
echo 正在启动计划书助手...
start "" http://127.0.0.1:8000
python app.py
pause
