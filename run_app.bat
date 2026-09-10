@echo off
echo ========================================================
echo   한국 및 미국 증시 ETF 수익률 비교 대시보드
echo ========================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py
pause
