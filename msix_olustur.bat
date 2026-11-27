@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Okul Revir Defteri - Microsoft Store MSIX Paket Ureticisi

cd /d "%~dp0"

echo ============================================================================
echo   OKUL REVIR DEFTERI - MICROSOFT STORE (.MSIX) PAKET DERLEYICISI
echo ============================================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python bulunamadi! Lutfen Python 3.10+ kurun ve PATH'e ekleyin.
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

python bump_version.py --sync
if errorlevel 1 (
    echo [HATA] Surum senkronizasyonu basarisiz oldu!
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

python build_msix.py %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo [HATA] MSIX olusturma sirasinda hata meydana geldi ^(Kod: %EXIT_CODE%^).
) else (
    echo [TAMAMLANDI] MSIX hazirlik islemi basariyla bitti.
)

if /I not "%REVIR_NO_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
