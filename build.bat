@echo off
REM ─────────────────────────────────────────────────────────────────────────
REM build.bat — one-shot build of the single-file Windows executable.
REM
REM What it does:
REM   1. Installs PyInstaller + runtime deps + Pillow (build-time only).
REM   2. Generates batch_configurator.ico via make_icon.py.
REM   3. Bundles batch_configurator.py + its deps into ONE .exe with the icon.
REM
REM Output:
REM   dist\BatchConfigurator.exe   ← ship this single file to anyone on Windows
REM
REM Requirements: Python 3.9+ on PATH. The exe itself needs no Python on the
REM target machine — PyInstaller embeds the interpreter and all libraries.
REM ─────────────────────────────────────────────────────────────────────────

setlocal
cd /d "%~dp0"

echo.
echo === [1/3] Installing build + runtime dependencies ===
python -m pip install --upgrade pip
python -m pip install pyinstaller pyserial cryptography pillow openpyxl
if errorlevel 1 (
    echo.
    echo  ERROR: pip install failed. Make sure Python is installed and on PATH.
    pause
    exit /b 1
)

echo.
echo === [2/3] Generating batch_configurator.ico ===
python make_icon.py
if errorlevel 1 (
    echo  ERROR: icon generation failed.
    pause
    exit /b 1
)

echo.
echo === [3/3] Building single-file executable ===
REM Flags:
REM   --onefile      one self-extracting .exe (no DLL folder to ship)
REM   --windowed     no console window flashing behind the Tk UI
REM   --icon         use the icon we just generated
REM   --name         output basename (so we get BatchConfigurator.exe, not batch_configurator.exe)
REM   --clean        purge PyInstaller's build cache before this run — avoids
REM                  stale-spec issues if you've tweaked dependencies
REM   --noconfirm    overwrite dist/ without prompting
pyinstaller --onefile --windowed ^
            --icon=batch_configurator.ico ^
            --name=BatchConfigurator ^
            --clean --noconfirm ^
            batch_configurator.py
if errorlevel 1 (
    echo.
    echo  ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done.  Your single-file Windows executable is here:
echo.
echo     %CD%\dist\BatchConfigurator.exe
echo.
echo  Send that one file to anyone — no Python install needed
echo  on their machine.
echo ============================================================
echo.
pause
endlocal
