@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Okul Revir Defteri - Masaustu Paketi ve Kurulum Sihirbazi Ureticisi

cd /d "%~dp0"

echo ============================================================================
echo   OKUL REVIR DEFTERI - EXE VE WINDOWS KURULUM SIHIRBAZI DERLEYICISI
echo ============================================================================
echo.

:: 1. Python kontrolu
python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python bulunamadi! Lutfen Python 3.10+ kurun ve PATH'e ekleyin.
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

:: 2. Gerekli paketleri kontrol et ve eksikleri kur (pywebview, pyinstaller, psutil)
echo [1/4] Gerekli Python kutuphaneleri kontrol ediliyor (pywebview, pyinstaller, psutil)...
set "MISSING_PKGS="

python -c "import webview" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! pywebview"
python -c "import PyInstaller" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! pyinstaller"
python -c "import psutil" >nul 2>&1 || set "MISSING_PKGS=!MISSING_PKGS! psutil"

if defined MISSING_PKGS (
    echo [*] Eksik paketler kuruluyor:!MISSING_PKGS!
    python -m pip install --upgrade !MISSING_PKGS!
    if errorlevel 1 (
        echo [HATA] Gerekli paketler kurulamadi! Internet baglantinizi veya pip ortamini kontrol edin.
        if /I not "%REVIR_NO_PAUSE%"=="1" pause
        exit /b 1
    )
) else (
    echo [OK] pywebview, pyinstaller ve psutil kurulu.
)
echo.

:: 3. Opsiyonel surum yukseltme (PyInstaller ve Inno Setup oncesi calistirilir)
echo [2/4] Surum kontrol ediliyor...
if /I "%~1"=="bump" (
    set "BUMP_KIND=%~2"
    if "!BUMP_KIND!"=="" set "BUMP_KIND=patch"
    python bump_version.py !BUMP_KIND!
    if errorlevel 1 (
        echo [HATA] Surum yukseltme basarisiz oldu!
        if /I not "%REVIR_NO_PAUSE%"=="1" pause
        exit /b 1
    )
) else (
    python bump_version.py --sync
    if errorlevel 1 (
        echo [HATA] Surum senkronizasyonu basarisiz oldu!
        if /I not "%REVIR_NO_PAUSE%"=="1" pause
        exit /b 1
    )
)

set "APP_VER=1.1.1"
if exist "VERSION" (
    for /f "usebackq tokens=* delims=" %%V in ("VERSION") do set "APP_VER=%%V"
)
echo [BILGI] Paketlenecek Uygulama Surumu: v!APP_VER!
echo.

:: 4. PyInstaller ile dist\RevirDefteri derlemesi
echo [3/4] PyInstaller ile RevirDefteri.spec derleniyor (bu islem 1-3 dakika surebilir)...
python -m PyInstaller RevirDefteri.spec --clean -y

if errorlevel 1 (
    echo [HATA] PyInstaller derlemesi basarisiz oldu!
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)

if not exist "dist\RevirDefteri\RevirDefteri.exe" (
    echo [HATA] Beklenen EXE dosyasi uretilemedi: dist\RevirDefteri\RevirDefteri.exe
    if /I not "%REVIR_NO_PAUSE%"=="1" pause
    exit /b 1
)
echo [OK] Masaustu EXE paketi hazir: dist\RevirDefteri\RevirDefteri.exe
echo.

:: 5. Inno Setup Compiler (ISCC.exe - Surum 7/6/5) varsa Setup.exe uret
echo [4/4] Inno Setup Compiler (ISCC.exe) araniyor...
set "ISCC_EXE="
for /f "delims=" %%P in ('where ISCC.exe 2^>nul') do if not defined ISCC_EXE set "ISCC_EXE=%%P"
if not defined ISCC_EXE if exist "C:\Program Files\Inno Setup 7\ISCC.exe" set "ISCC_EXE=C:\Program Files\Inno Setup 7\ISCC.exe"
if not defined ISCC_EXE if exist "C:\Program Files (x86)\Inno Setup 7\ISCC.exe" set "ISCC_EXE=C:\Program Files (x86)\Inno Setup 7\ISCC.exe"
if not defined ISCC_EXE if exist "%LocalAppData%\Programs\Inno Setup 7\ISCC.exe" set "ISCC_EXE=%LocalAppData%\Programs\Inno Setup 7\ISCC.exe"
if not defined ISCC_EXE if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC_EXE=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC_EXE=C:\Program Files\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set "ISCC_EXE=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC_EXE if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" set "ISCC_EXE=C:\Program Files (x86)\Inno Setup 5\ISCC.exe"

if not exist "installer" mkdir "installer"

if defined ISCC_EXE (
    echo [*] Inno Setup bulundu: "%ISCC_EXE%"
    "%ISCC_EXE%" "RevirDefteri_Setup.iss"
    if errorlevel 1 (
        echo [HATA] Inno Setup kurulum sihirbazi derlenirken hata olustu!
        if /I not "%REVIR_NO_PAUSE%"=="1" pause
        exit /b 1
    )
    set "SETUP_NAME=RevirDefteri_Kurulum_v!APP_VER!.exe"
    if exist "installer\!SETUP_NAME!" (
        del /Q "RevirDefteri_Kurulum_v*.exe" >nul 2>&1
        for %%F in ("installer\RevirDefteri_Kurulum_v*.exe") do if /I not "%%~nxF"=="!SETUP_NAME!" del /Q "%%F" >nul 2>&1
        copy /Y "installer\!SETUP_NAME!" "!SETUP_NAME!" >nul
        echo [BASARILI] Kurulum dosyasi olusturuldu ^(v!APP_VER!^):
        echo     - %cd%\!SETUP_NAME!
        echo     - %cd%\installer\!SETUP_NAME!
    )
) else (
    echo [UYARI] Sistemde Inno Setup 6 ^(ISCC.exe^) kurulu degil.
    echo         Tek klasor masaustu uygulamasi hazir: dist\RevirDefteri\RevirDefteri.exe
    echo         Kurulum sihirbazi uretmek icin Inno Setup 6 kurup ISCC.exe ile RevirDefteri_Setup.iss dosyasini derleyin.
)

echo.
echo ============================================================================
echo   ISLEM TAMAMLANDI.
echo ============================================================================
if /I not "%REVIR_NO_PAUSE%"=="1" pause
