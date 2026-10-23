@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================================
echo   Iniciando prueba de login Microsoft 365 + LM Studio
echo ==========================================================
.\.venv\Scripts\python.exe test_auth.py
echo.
pause
