# -*- mode: python ; coding: utf-8 -*-
"""
Okul Revir Defteri - PyInstaller Paketleme (.spec) Dosyası

Çalıştırma Komutu (Spec ile önerilen):
    python -m PyInstaller build_exe.spec --noconfirm --clean

Çıktı Klasörü:
    dist/RevirDefteri/RevirDefteri.exe
"""

from pathlib import Path
from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
    copy_metadata,
)

PROJECT_ROOT = Path(SPECPATH).resolve()

datas = [
    (str(PROJECT_ROOT / "persantil.py"), "."),
    (str(PROJECT_ROOT / "revir_ui.py"), "."),
    (str(PROJECT_ROOT / "gemini_assistant.py"), "."),
    (str(PROJECT_ROOT / "insulin_diabetes.py"), "."),
    (str(PROJECT_ROOT / "VERSION"), "."),
    (str(PROJECT_ROOT / "app_icon.ico"), "."),
    (str(PROJECT_ROOT / "app_icon.png"), "."),
    (str(PROJECT_ROOT / "turkiye-il-ilce-mahalle-koy.json"), "."),
    (str(PROJECT_ROOT / "persantil veriler"), "persantil veriler"),
]

binaries = []
hiddenimports = [
    "revir_ui",
    "persantil",
    "gemini_assistant",
    "insulin_diabetes",
    "clr_loader",
    "pythonnet",
    "truststore",
    "sqlite3",
    "pandas",
    "numpy",
    "PIL",
    "PIL.Image",
    "PIL.ImageOps",
    "PIL.ImageDraw",
    "docx",
    "openpyxl",
    "pymupdf",
    "fitz",
    "altair",
    "webview",
    "webview.platforms.edgechromium",
    "webview.platforms.winforms",
    "psutil",
]

for pkg_name in ("altair", "docx", "openpyxl", "pymupdf", "webview"):
    datas += collect_data_files(pkg_name)
    binaries += collect_dynamic_libs(pkg_name)
    hiddenimports += collect_submodules(pkg_name)

for extra_sub_pkg in ("clr_loader", "pythonnet"):
    try:
        datas += collect_data_files(extra_sub_pkg)
        binaries += collect_dynamic_libs(extra_sub_pkg)
        hiddenimports += collect_submodules(extra_sub_pkg)
    except Exception:
        pass

for meta_pkg in (
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

excludes = [
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
    [str(PROJECT_ROOT / "desktop_launcher.py")],
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
