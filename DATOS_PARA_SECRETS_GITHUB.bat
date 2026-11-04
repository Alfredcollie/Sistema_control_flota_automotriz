@echo off
chcp 65001 >nul
cd /d "%~dp0"
python DATOS_PARA_SECRETS_GITHUB.py
echo.
pause
