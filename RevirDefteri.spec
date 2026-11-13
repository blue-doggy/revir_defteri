# -*- mode: python ; coding: utf-8 -*-
"""
Okul Revir Defteri ve Sağlık Takip Sistemi - PyInstaller Paketleme (.spec)

Derleme Komutu:
    python -m PyInstaller RevirDefteri.spec --clean -y

Çıktı Klasörü (Tek Klasör / onedir):
    dist/RevirDefteri/RevirDefteri.exe
"""

from pathlib import Path
from PyInstaller.utils.hooks import (
    collect_all,
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
    copy_metadata,
)

PROJECT_ROOT = Path(SPECPATH).resolve()

datas = [
    (str(PROJECT_ROOT / "persantil.py"), "."),
    (str(PROJECT_ROOT / "VERSION"), "."),
    (str(PROJECT_ROOT / "app_icon.png"), "."),
    (str(PROJECT_ROOT / "app_icon.ico"), "."),
    (str(PROJECT_ROOT / "turkiye-il-ilce-mahalle-koy.json"), "."),
    (str(PROJECT_ROOT / "persantil veriler"), "persantil veriler"),
]
if (PROJECT_ROOT / ".streamlit").exists():
    datas.append((str(PROJECT_ROOT / ".streamlit"), ".streamlit"))

binaries = []
hiddenimports = [
    "sqlite3",
    "pandas",
    "numpy",
    "altair",
    "openpyxl",
    "docx",
    "pymupdf",
    "fitz",
    "PIL",
    "PIL.Image",
    "PIL.ImageOps",
    "PIL.ImageDraw",
    "webview",
    "webview.platforms.edgechromium",
    "webview.platforms.winforms",
    "clr",
    "pythonnet",
    "clr_loader",
    "truststore",
    "psutil",
]

# Streamlit tüm veri dosyaları, alt modüller ve ikili eklentileri
st_datas, st_bins, st_hidden = collect_all("streamlit")
datas += st_datas
binaries += st_bins
hiddenimports += st_hidden

# Altair, python-docx, openpyxl, pymupdf, pywebview ve pythonnet bağımlılıkları
for pkg_name in ("altair", "docx", "openpyxl", "pymupdf", "webview", "clr_loader", "pythonnet"):
    try:
        datas += collect_data_files(pkg_name)
        binaries += collect_dynamic_libs(pkg_name)
        hiddenimports += collect_submodules(pkg_name)
    except Exception:
        pass

# importlib.metadata gereksinimleri (Streamlit ve bileşenleri için zorunlu)
for meta_pkg in (
    "streamlit",
    "altair",
    "openpyxl",
    "python-docx",
    "pymupdf",
    "pywebview",
    "pandas",
    "pillow",
    "packaging",
):
    try:
        datas += copy_metadata(meta_pkg)
    except Exception:
        pass

# Kullanılmayan ağır GUI/ML paketlerini hariç tut (taralı PDF OCR gerekiyorsa easyocr/torch çıkarılabilir)
excludes = [
    "tkinter",
    "_tkinter",
    "PyQt6",
    "PyQt5",
    "PySide6",
    "IPython",
    "jupyter",
    "notebook",
    "pytest",
    "torch",
    "torchvision",
    "torchaudio",
    "easyocr",
    "cv2",
    "scipy",
    "matplotlib",
]

a = Analysis(
    [str(PROJECT_ROOT / "desktop_app.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=list(dict.fromkeys(hiddenimports)),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="RevirDefteri",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(PROJECT_ROOT / "app_icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="RevirDefteri",
)
