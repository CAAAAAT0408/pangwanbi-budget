@echo off
chcp 65001 >nul
title 지출계획서 자동작성
cd /d "%~dp0"
python "app\control.py" open jichul
if errorlevel 1 pause
timeout /t 4 >nul
