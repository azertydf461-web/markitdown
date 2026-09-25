@echo off
chcp 65001 >nul
cd /d "%~dp0"
python --version >nul 2>&1 || (
  echo Python not found. Install Python 3.10+ from https://www.python.org and tick "Add python.exe to PATH".
  pause
  exit /b 1
)
python -m pip install --upgrade pip
python -m pip install --upgrade "markitdown[pdf,docx,pptx,xlsx,xls,outlook]" pymupdf pywin32
echo.
echo Done. Optional for scanned documents: run get_tessdata.bat
pause
