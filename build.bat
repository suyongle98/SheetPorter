@echo off
setlocal
cd /d "%~dp0"

echo [1/3] Creating virtual environment...
if not exist ".venv" python -m venv .venv
if errorlevel 1 goto fail
call ".venv\Scripts\activate.bat"
if errorlevel 1 goto fail

echo [2/3] Installing dependencies...
python -m pip install --upgrade pip >nul
python -m pip install pyinstaller -r requirements.txt
if errorlevel 1 goto fail

echo [3/3] Building SheetPorter.exe...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name SheetPorter --add-data "examples;examples" sheetporter.py
if errorlevel 1 goto fail

echo.
if exist "dist\SheetPorter.exe" (
  echo Done: dist\SheetPorter.exe
) else (
  echo Build failed, please check the log above.
)
pause
exit /b 0

:fail
echo Build failed. See errors above.
pause
exit /b 1
