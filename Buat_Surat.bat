@echo off
cd /d "%~dp0"
python --version >nul 2>nul
if not errorlevel 1 (
    python "main.py"
    goto selesai
)
py -3 --version >nul 2>nul
if not errorlevel 1 (
    py -3 "main.py"
    goto selesai
)
echo Python belum terpasang. Install Python 3 dari https://python.org lalu coba lagi.
:selesai
pause