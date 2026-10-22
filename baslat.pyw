"""
Revir Defteri - Bagimsiz Masaustu Uygulama Baslaticisi (.pyw)
Tarayiciya ihtiyac duymadan, konsol / siyah ekran cikarmadan
bagimsiz bir masaustu penceresi olarak baslatir.
"""
import os
import sys
import time
import socket
import subprocess
import urllib.request
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
PYTHONW_EXE = Path(sys.executable).parent / "pythonw.exe"
if not PYTHONW_EXE.exists():
    PYTHONW_EXE = Path(sys.executable)

CREATE_NO_WINDOW = 0x08000000

def is_server_running(port: int = 8501) -> bool:
    try:
        req = urllib.request.urlopen(f"http://localhost:{port}/", timeout=1.5)
        return req.status == 200
    except Exception:
        return False

def get_app_browser() -> str | None:
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None

def main():
    # 1. Hizli Internet ve Guncelleme Kontrolu (1 saniye)
    try:
        s = socket.create_connection(("8.8.8.8", 53), timeout=1.0)
        s.close()
        r = subprocess.run(["git", "remote", "get-url", "origin"], cwd=APP_DIR, capture_output=True, text=True, timeout=2, creationflags=CREATE_NO_WINDOW)
        if r.returncode != 0:
            subprocess.run(["git", "remote", "add", "origin", "https://github.com/blue-doggy/revir_defteri.git"], cwd=APP_DIR, capture_output=True, timeout=2, creationflags=CREATE_NO_WINDOW)
        subprocess.run(
            ["git", "pull", "--ff-only", "origin", "main"],
            cwd=APP_DIR,
            capture_output=True,
            timeout=5,
            creationflags=CREATE_NO_WINDOW
        )
    except Exception:
        pass

    # 2. Streamlit Arka Plan Servisini Baslat (Eger calismiyorsa)
    st_proc = None
    if not is_server_running(8501):
        st_cmd = [
            str(PYTHONW_EXE), "-m", "streamlit", "run", "persantil.py",
            "--server.headless", "true",
            "--server.port", "8501",
            "--global.developmentMode", "false"
        ]
        st_proc = subprocess.Popen(
            st_cmd,
            cwd=APP_DIR,
            creationflags=CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )

        # Sunucunun hazir olmasini bekle (maksimum 12 saniye)
        for _ in range(24):
            if is_server_running(8501):
                break
            time.sleep(0.5)

    # 3. Bagimsiz Masaustu Uygulama Penceresini Ac (Adres cubugu ve sekmeler olmadan)
    browser_exe = get_app_browser()
    profile_dir = APP_DIR / ".app_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    if browser_exe:
        app_args = [
            browser_exe,
            f"--app=http://localhost:8501",
            f"--user-data-dir={profile_dir}",
            "--window-size=1380,880",
            "--window-position=60,40",
            "--disable-extensions",
            "--disable-plugins",
        ]
        app_proc = subprocess.Popen(app_args, creationflags=CREATE_NO_WINDOW)
        # Pencere kapanana kadar bekle
        app_proc.wait()
    else:
        import webbrowser
        webbrowser.open("http://localhost:8501")

    # 4. Pencere kapandiginda arka plandaki Streamlit servisini temiz kapat
    if st_proc is not None:
        try:
            st_proc.terminate()
            st_proc.wait(timeout=3)
        except Exception:
            try:
                st_proc.kill()
            except Exception:
                pass

if __name__ == "__main__":
    main()

