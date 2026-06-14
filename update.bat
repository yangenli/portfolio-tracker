@echo off
REM Double-click to refresh prices AND publish to the live GitHub Pages site.
cd /d "%~dp0"

echo === Refreshing prices from Yahoo Finance ===
"C:/Users/yange/anaconda3/envs/simple/python.exe" compute_portfolios.py
if errorlevel 1 ( echo. & echo Script failed - nothing published. & pause & exit /b 1 )

echo.
echo === Publishing to GitHub Pages ===
git add docs/results.json
git diff --cached --quiet && (
  echo No price changes since last run - nothing to publish.
) || (
  git commit -m "Update results" && git push && echo Published.
)

echo.
echo Live site: https://yangenli.github.io/portfolio-tracker/
echo (the page may take ~1 minute to show the new numbers)
pause
