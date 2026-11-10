@echo off
chcp 65001 >nul
title 사업관리비 실적계획 자동작성
cd /d "%~dp0"
python "app\control.py" open siljeok
if errorlevel 1 pause
timeout /t 4 >nul
