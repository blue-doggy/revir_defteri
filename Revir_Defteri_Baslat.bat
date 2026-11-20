@echo off
chcp 65001 >nul
title Revir Defteri - Okul Saglik Yonetim Sistemi
cd /d "%~dp0"

echo ====================================================================
echo                   REVIR DEFTERI BASLATILIYOR...
echo ====================================================================
echo.

:: 1. Hizli Internet ve Guncelleme Kontrolu (1 saniye)
echo [*] Guncellemeler denetleniyor...
ping -n 1 -w 1000 8.8.8.8 >nul 2>&1
if %errorlevel% equ 0 (
    git remote >nul 2>&1
    if %errorlevel% equ 0 (
        git fetch origin main >nul 2>&1
        git pull --ff-only origin main >nul 2>&1
    )
) else (
    echo [i] Cevrimdisi mod: Internet baglantisi yok, mevcut surum aciliyor.
)

:: 2. Python Masaustu Yoneticisi ile Uygulamayi Baslat
echo [*] Masaustu uygulamasi aciliyor, lutfen bekleyin...
where pythonw >nul 2>&1
if %errorlevel% equ 0 (
    start "" pythonw "desktop_app.py"
) else (
    python "desktop_app.py"
)
