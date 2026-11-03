@echo off
chcp 65001 >nul
title 판관비 예산안 자동작성
cd /d "%~dp0"
python "app\control.py" open
if errorlevel 1 pause
timeout /t 4 >nul
