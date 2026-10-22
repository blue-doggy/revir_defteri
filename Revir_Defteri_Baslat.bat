@echo off
chcp 65001 >nul
title Revir Defteri - Okul Saglik Yonetim Sistemi
cls
echo ====================================================================
echo                   REVIR DEFTERI BASLATILIYOR...
echo ====================================================================
echo.

cd /d "%USERPROFILE%\OneDrive\Masaüstü\kod\ai persantil"

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

:: 2. Tarayiciyi ve Streamlit Uygulamasini Baslat
echo [*] Program aciliyor, lutfen bekleyin...
start "" "http://localhost:8501"
"C:\Users\OKUL\AppData\Local\Python\pythoncore-3.14-64\python.exe" -m streamlit run "persantil.py" --server.headless false --server.port 8501
pause
