@echo off
chcp 65001 >nul
echo ========================================================
echo   ETF Performance Comparison Dashboard
echo ========================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py
pause
