@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist tessdata mkdir tessdata
for %%L in (rus eng) do (
  echo Downloading %%L.traineddata ...
  curl -L --fail -o "tessdata\%%L.traineddata" "https://github.com/tesseract-ocr/tessdata_best/raw/main/%%L.traineddata" || (
    echo Download failed. See README.md, section OCR, for manual download.
    pause
    exit /b 1
  )
)
echo Done: %~dp0tessdata
pause
