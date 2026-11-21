# -*- mode: python ; coding: utf-8 -*-
"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Çapraz Platform PyInstaller Paketleme (.spec)

Derleme Komutu:
    python -m PyInstaller RevirDefteri.spec --clean -y

Çıktılar:
    Windows : dist/RevirDefteri/RevirDefteri.exe
    macOS   : dist/RevirDefteri.app
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
    copy_metadata,
)

PROJECT_ROOT = Path(SPECPATH).resolve()
IS_MACOS = sys.platform == "darwin"
IS_WINDOWS = sys.platform == "win32"

_version_file = PROJECT_ROOT / "VERSION"
APP_VERSION = (
    _version_file.read_text(encoding="utf-8").strip()
    if _version_file.exists()
    else "1.13.0"
)

datas = [
    (str(PROJECT_ROOT / "persantil.py"), "."),
    (str(PROJECT_ROOT / "revir_ui.py"), "."),
    (str(PROJECT_ROOT / "gemini_assistant.py"), "."),
    (str(PROJECT_ROOT / "insulin_diabetes.py"), "."),
    (str(PROJECT_ROOT / "document_studio.py"), "."),
    (str(PROJECT_ROOT / "gemini_api_anahtari_kilavuzu.pdf"), "."),
    (str(PROJECT_ROOT / "VERSION"), "."),
    (str(PROJECT_ROOT / "app_icon.png"), "."),
    (str(PROJECT_ROOT / "app_icon.ico"), "."),
    (str(PROJECT_ROOT / "turkiye-il-ilce-mahalle-koy.json"), "."),
    (str(PROJECT_ROOT / "persantil veriler"), "persantil veriler"),
]
if (PROJECT_ROOT / "app_icon.icns").exists():
    datas.append((str(PROJECT_ROOT / "app_icon.icns"), "."))

binaries = []
hiddenimports = [
    "revir_ui",
    "persantil",
    "gemini_assistant",
    "insulin_diabetes",
    "document_studio",
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
    "truststore",
    "psutil",
    "sounddevice",
    "_sounddevice_data",
]

collect_packages = [
    "altair",
    "docx",
    "openpyxl",
    "pymupdf",
    "webview",
    "sounddevice",
    "_sounddevice_data",
]

if IS_WINDOWS:
    hiddenimports += [
        "webview.platforms.edgechromium",
        "webview.platforms.winforms",
        "clr",
        "pythonnet",
        "clr_loader",
    ]
    collect_packages += ["clr_loader", "pythonnet"]
elif IS_MACOS:
    hiddenimports += [
        "webview.platforms.cocoa",
        "objc",
        "Foundation",
        "AppKit",
        "WebKit",
        "PyObjCTools",
    ]

for pkg_name in collect_packages:
    try:
        datas += collect_data_files(pkg_name)
        binaries += collect_dynamic_libs(pkg_name)
        hiddenimports += collect_submodules(pkg_name)
    except Exception:
        pass

# importlib.metadata gereksinimleri
for meta_pkg in (
    "altair",
    "openpyxl",
    "python-docx",
    "pymupdf",
    "pywebview",
    "pandas",
    "pillow",
    "packaging",
    "sounddevice",
    "truststore",
):
    try:
        datas += copy_metadata(meta_pkg)
    except Exception:
        pass

# Kullanılmayan ağır GUI/ML paketlerini hariç tut
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

if IS_MACOS:
    _icns = PROJECT_ROOT / "app_icon.icns"
    APP_ICON = str(_icns if _icns.exists() else (PROJECT_ROOT / "app_icon.png"))
else:
    APP_ICON = str(PROJECT_ROOT / "app_icon.ico")

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
    icon=APP_ICON,
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

if IS_MACOS:
    app = BUNDLE(
        coll,
        name="RevirDefteri.app",
        icon=APP_ICON,
        bundle_identifier="com.okulrevir.revirdefteri",
        info_plist={
            "CFBundleName": "RevirDefteri",
            "CFBundleDisplayName": "Okul Revir Defteri",
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_VERSION,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
            "NSMicrophoneUsageDescription": (
                "Okul Revir Defteri sesli komut ve Gemini AI sesli sohbet "
                "özellikleri için mikrofon erişimi gereklidir."
            ),
        },
    )
