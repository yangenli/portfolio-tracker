@echo off
REM Double-click this to refresh docs\results.json with the latest prices.
cd /d "%~dp0"
"C:/Users/yange/anaconda3/envs/simple/python.exe" compute_portfolios.py
echo.
echo Done. docs\results.json updated. (Commit/push it if hosting on GitHub Pages.)
pause
