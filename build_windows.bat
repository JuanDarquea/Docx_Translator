@echo off
setlocal

REM Build Windows EXE using venv
call .venv\Scripts\activate

pyinstaller --noconfirm --clean --windowed --onefile ^
  --add-data "settings.json;." ^
  --add-data "Files\quality_rules.json;Files" ^
  --add-data "Files\protected_words.json;Files" ^
  --add-data "translation_memory.json;." ^
  Docx_Translator.py

echo.
echo Build complete. Output in dist\Docx_Translator.exe
pause
