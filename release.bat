@echo off
REM ─────────────────────────────────────────────────────────────────────────
REM release.bat — one-click: commit + push + wait for CI + download all
REM 5 binaries into .\downloads\.
REM
REM Requires: git and gh on PATH. Run `gh auth login` once before first use.
REM ─────────────────────────────────────────────────────────────────────────

setlocal EnableDelayedExpansion
cd /d "%~dp0"

set /p MSG=Commit message (Enter for "Update"):
if "!MSG!"=="" set MSG=Update

echo.
echo === [1/4] Committing and pushing ===
git add .
git commit -m "!MSG!"
if errorlevel 1 (
    echo Nothing to commit ^(or commit failed^). Continuing to check for existing runs...
)
git push
if errorlevel 1 (
    echo Push failed.
    pause
    exit /b 1
)

echo.
echo === [2/4] Waiting for GitHub to register the run ===
REM Push -> workflow trigger has ~2-5s of lag; give it a beat.
timeout /t 6 /nobreak >nul

echo.
echo === [3/4] Watching build (5-15 minutes) ===
for /f %%i in ('gh run list --limit 1 --json databaseId --jq ".[0].databaseId"') do set RUN_ID=%%i
echo Run ID: !RUN_ID!
echo Web view: https://github.com/ismailkoya/batch_configurator/actions/runs/!RUN_ID!
gh run watch !RUN_ID! --exit-status --interval 10
if errorlevel 1 (
    echo.
    echo One or more jobs failed. Check the web view above.
    pause
    exit /b 1
)

echo.
echo === [4/4] Downloading artifacts to .\downloads\ ===
if exist downloads rmdir /s /q downloads
gh run download !RUN_ID! --dir downloads
if errorlevel 1 (
    echo Download failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done. All 5 binaries are in:
echo     %CD%\downloads
echo ============================================================
echo.
dir /b downloads
echo.
pause
endlocal
