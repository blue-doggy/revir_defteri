@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Okul Revir Defteri - Tek Tikla Surum ve Guncelleme .EXE Ureticisi

cd /d "%~dp0"

echo ============================================================================
echo   OKUL REVIR DEFTERI - YENI SURUM VE GUNCELLEME .EXE DERLEYICISI
echo ============================================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python bulunamadi! Lutfen Python 3.10+ kurun ve PATH'e ekleyin.
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

echo [KONTROL] Gerekli kutuphaneler kontrol ediliyor (pywebview, pyinstaller, psutil)...
set "MISSING_PKGS="
python -c "import webview" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! pywebview"
python -c "import PyInstaller" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! pyinstaller"
python -c "import psutil" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! psutil"

if defined MISSING_PKGS (
    echo [*] Eksik paketler kuruluyor:!MISSING_PKGS!
    python -m pip install --upgrade !MISSING_PKGS!
    if errorlevel 1 (
        echo [HATA] Gerekli paketler kurulamadi!
        if /I not "%REVIR_NO_PAUSE%"=="1" pause
        exit /b 1
    )
)

echo.
python build_release.py %*
if errorlevel 1 (
    echo.
    echo [HATA] Derleme islemi sirasinda bir hata olustu!
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

echo.
echo ============================================================================
echo   TUM PAKETLER BASARIYLA OLUSTURULDU.
echo ============================================================================
if /I not "%REVIR_NO_PAUSE%"=="1" pause
