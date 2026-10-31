@echo off
setlocal EnableExtensions
title Pemasang Buat Surat

REM ============================================================
REM  Ganti REPO_BASE dengan URL repo GitHub Anda (tanpa .git)
REM  Contoh: https://github.com/username/buat-surat
REM ============================================================
set "REPO_BASE=https://github.com/zawawisss/SPPGLetterGenerator"
REM ============================================================

set "ZIP_URL=%REPO_BASE%/archive/refs/heads/main.zip"

echo ============================================
echo            PEMASANG BUAT SURAT
echo ============================================
echo.

:pilih_folder
set "APP_DIR=%USERPROFILE%\Buat Surat"
set "pilihan="
echo Simpan di: %APP_DIR%
echo (Enter = setuju, atau ketik folder lain)
set /p "pilihan=Folder: "
if defined pilihan set "APP_DIR=%pilihan%"
set "APP_DIR=%APP_DIR:"=%"
if "%APP_DIR%"=="" goto pilih_folder
for %%I in ("%APP_DIR%") do set "APP_DIR=%%~fI"
if "%APP_DIR%"=="" goto pilih_folder

if not exist "%APP_DIR%" mkdir "%APP_DIR%" 2>nul
if not exist "%APP_DIR%" (
    echo Folder tidak bisa dibuat. Coba lagi.
    echo.
    goto pilih_folder
)

echo tes > "%APP_DIR%\_cek.txt" 2>nul
if not exist "%APP_DIR%\_cek.txt" (
    echo Folder tidak bisa dipakai. Coba lagi.
    echo.
    goto pilih_folder
)
del "%APP_DIR%\_cek.txt" >nul 2>nul

if exist "%APP_DIR%\main.py" (
    echo Aplikasi sudah ada.
    goto cek_python
)

echo Mengunduh...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" ^
  "$zip=Join-Path $env:TEMP 'buatsurat.zip';" ^
  "$tmp=Join-Path $env:TEMP 'buatsurat_extract';" ^
  "if(Test-Path $tmp){Remove-Item $tmp -Recurse -Force};" ^
  "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile $zip -UseBasicParsing;" ^
  "Expand-Archive -Path $zip -DestinationPath $tmp -Force;" ^
  "$src=Get-ChildItem $tmp -Directory | Select-Object -First 1;" ^
  "Copy-Item -Path (Join-Path $src.FullName '*') -Destination '%APP_DIR%' -Recurse -Force;" ^
  "Remove-Item $zip -Force;" ^
  "Remove-Item $tmp -Recurse -Force" >nul 2>nul

:cek_hasil
if not exist "%APP_DIR%\main.py" (
    echo Gagal mengunduh. Cek internet, lalu coba lagi.
    pause
    exit /b 1
)

:cek_python
python --version >nul 2>nul
if not errorlevel 1 goto buat_shortcut
py -3 --version >nul 2>nul
if not errorlevel 1 goto buat_shortcut

echo Python belum ada. Mencoba pasang...
echo.
where winget >nul 2>nul
if errorlevel 1 goto python_manual
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements >nul 2>nul
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do set "PYDIR=%%D"
if defined PYDIR set "PATH=%PYDIR%;%PYDIR%\Scripts;%PATH%"
python --version >nul 2>nul
if not errorlevel 1 goto buat_shortcut
py -3 --version >nul 2>nul
if not errorlevel 1 goto buat_shortcut

:python_manual
echo Python belum terpasang.
echo Install dari https://www.python.org/downloads/
echo lalu jalankan lagi.
start "" "https://www.python.org/downloads/"
pause
exit /b 1

:buat_shortcut
powershell -NoProfile -Command "$d=[Environment]::GetFolderPath('Desktop'); $ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut((Join-Path $d 'Buat Surat.lnk')); $sc.TargetPath='%APP_DIR%\Buat_Surat.bat'; $sc.WorkingDirectory='%APP_DIR%'; $sc.Save()" >nul 2>nul

echo.
echo Selesai. Shortcut ada di Desktop.
echo.
echo Membuka aplikasi...
cd /d "%APP_DIR%"
call "Buat_Surat.bat"
exit /b 0
