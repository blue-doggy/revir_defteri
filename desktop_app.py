"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Masaüstü Başlatıcı (desktop_app.py)

Özellikler:
- Windows Mutex ile Tek Örnek (Single Instance) Koruması ve mevcut pencereyi öne getirme
- 127.0.0.1 üzerinde dinamik boş port bulma ve konsol penceresiz (CREATE_NO_WINDOW) Streamlit sunucusu
- Sunucu hazır olana kadar Türkçe Splash Screen (Açılış Ekranı) gösterimi
- pywebview ALLOW_DOWNLOADS=True ile Word/Excel/JSON/DB indirme desteği
- Pencere kapatıldığında (closed) arka plan Streamlit/Python süreç ağacını kalıntısız sonlandırma
"""
from __future__ import annotations

import argparse
import atexit
import ctypes
import html as html_lib
import os
import signal
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from ctypes import wintypes
from pathlib import Path

try:
	import psutil
except ImportError:
	psutil = None


def _ensure_standard_streams() -> None:
	"""PyInstaller --noconsole / pythonw.exe altında None olabilen stdout/stderr akışlarını korur."""
	if sys.stdout is None:
		sys.stdout = open(os.devnull, "w", encoding="utf-8")
	if sys.stderr is None:
		sys.stderr = open(os.devnull, "w", encoding="utf-8")


_ensure_standard_streams()

_DEFAULT_APP_VERSION = "1.2.0"


def _load_app_version() -> str:
	base_script = Path(__file__).resolve().parent
	exe_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else base_script
	meipass = Path(getattr(sys, "_MEIPASS", exe_dir))
	for folder in (exe_dir, meipass, base_script):
		v_file = folder / "VERSION"
		if v_file.exists():
			try:
				txt = v_file.read_text(encoding="utf-8").strip()
				if txt:
					return txt
			except Exception:
				pass
	return _DEFAULT_APP_VERSION


APP_VERSION = _load_app_version()
APP_TITLE = f"Okul Revir Defteri ve Sağlık Takip Sistemi v{APP_VERSION}"
MUTEX_NAME = "Local\\OkulRevirDefteri_SingleInstance_Mutex_v1"
SETUP_MUTEX_NAME = "Local\\OkulRevirDefteri_Setup_Mutex_v1"
DEFAULT_PORT = 8501
WINDOW_WIDTH = 1440
WINDOW_HEIGHT = 900
MIN_WINDOW_WIDTH = 1024
MIN_WINDOW_HEIGHT = 700
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

_SINGLE_INSTANCE_HANDLE = None

SPLASH_HTML = """<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Okul Revir Defteri Başlatılıyor</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html, body {
    height: 100%;
    width: 100%;
    font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
    background: linear-gradient(135deg, #0b2d48 0%, #124e78 52%, #1d70a2 100%);
    color: #ffffff;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    user-select: none;
  }
  .card {
    background: rgba(255, 255, 255, 0.11);
    border: 1px solid rgba(255, 255, 255, 0.24);
    border-radius: 20px;
    padding: 44px 52px;
    text-align: center;
    box-shadow: 0 24px 64px rgba(0, 0, 0, 0.35);
    backdrop-filter: blur(12px);
    max-width: 620px;
    width: 90%;
  }
  .badge {
    width: 74px;
    height: 74px;
    margin: 0 auto 22px auto;
    border-radius: 18px;
    background: #ffffff;
    color: #c62828;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 42px;
    font-weight: 800;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.22);
  }
  h1 {
    font-size: 22px;
    font-weight: 700;
    line-height: 1.45;
    margin-bottom: 10px;
    letter-spacing: 0.2px;
  }
  p {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.88);
    margin-bottom: 28px;
  }
  .spinner {
    width: 44px;
    height: 44px;
    margin: 0 auto 14px auto;
    border: 4px solid rgba(255, 255, 255, 0.28);
    border-top-color: #ffffff;
    border-radius: 50%;
    animation: spin 0.9s linear infinite;
  }
  .sub {
    font-size: 12px;
    color: rgba(255, 255, 255, 0.72);
  }
  @keyframes spin {
    to { transform: rotate(360deg); }
  }
</style>
</head>
<body>
  <div class="card">
    <div class="badge">✚</div>
    <h1>✚ Okul Revir Defteri Hazırlanıyor...</h1>
    <p>Yerel sağlık veritabanı ve klinik modüller başlatılıyor, lütfen bekleyin.</p>
    <div class="spinner"></div>
    <div class="sub">v__APP_VERSION__ · Güvenli Çevrimdışı Masaüstü Oturumu (127.0.0.1)</div>
  </div>
</body>
</html>
""".replace("__APP_VERSION__", APP_VERSION)


def _is_dir_writable(path: Path) -> bool:
	try:
		path.mkdir(parents=True, exist_ok=True)
		probe = path / f".revir_write_test_{os.getpid()}.tmp"
		probe.write_text("ok", encoding="utf-8")
		probe.unlink(missing_ok=True)
		return True
	except Exception:
		return False


def resolve_runtime_paths() -> dict[str, Path]:
	"""Geliştirme (.py) ve PyInstaller (.exe) modları için çalışma ve paket dizinlerini çözümler."""
	is_frozen = getattr(sys, "frozen", False)
	if is_frozen:
		exe_dir = Path(sys.executable).resolve().parent
		resource_dir = Path(getattr(sys, "_MEIPASS", exe_dir)).resolve()
	else:
		exe_dir = Path(__file__).resolve().parent
		resource_dir = exe_dir

	env_data = os.environ.get("REVIR_DATA_DIR", "").strip()
	if env_data:
		data_dir = Path(env_data).resolve()
		data_dir.mkdir(parents=True, exist_ok=True)
	elif _is_dir_writable(exe_dir):
		data_dir = exe_dir
	else:
		local_base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
		data_dir = local_base / "RevirDefteri"
		data_dir.mkdir(parents=True, exist_ok=True)

	script_candidates = (
		exe_dir / "persantil.py",
		resource_dir / "persantil.py",
	)
	script_path = next((p for p in script_candidates if p.exists()), resource_dir / "persantil.py")

	icon_candidates = (
		exe_dir / "app_icon.ico",
		resource_dir / "app_icon.ico",
	)
	icon_path = next((p for p in icon_candidates if p.exists()), resource_dir / "app_icon.ico")

	return {
		"exe_dir": exe_dir,
		"resource_dir": resource_dir,
		"data_dir": data_dir,
		"script_path": script_path,
		"icon_path": icon_path,
	}


def _bring_existing_window_to_front(title_keyword: str = "Okul Revir Defteri") -> bool:
	"""Halihazırda açık olan Revir Defteri penceresini bulur, küçültülmüşse açar ve öne getirir."""
	if sys.platform != "win32":
		return False
	user32 = ctypes.windll.user32
	hwnd = user32.FindWindowW(None, APP_TITLE)
	if not hwnd:
		matched: list[int] = []
		enum_proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

		def _enum_cb(h: int, _lparam: int) -> bool:
			if not user32.IsWindowVisible(h):
				return True
			length = user32.GetWindowTextLengthW(h)
			if length > 0:
				buf = ctypes.create_unicode_buffer(length + 1)
				user32.GetWindowTextW(h, buf, length + 1)
				if title_keyword.lower() in buf.value.lower():
					matched.append(h)
					return False
			return True

		user32.EnumWindows(enum_proc_type(_enum_cb), 0)
		if matched:
			hwnd = matched[0]

	if hwnd:
		SW_RESTORE = 9
		SW_SHOW = 5
		if user32.IsIconic(hwnd):
			user32.ShowWindow(hwnd, SW_RESTORE)
		else:
			user32.ShowWindow(hwnd, SW_SHOW)
		user32.SetForegroundWindow(hwnd)
		return True
	return False


def _is_setup_wizard_running() -> bool:
	"""Kurulum veya güncelleme sihirbazı (Inno Setup) aktifse True döndürür."""
	if sys.platform != "win32":
		return False
	SYNCHRONIZE = 0x00100000
	kernel32 = ctypes.windll.kernel32
	for m_name in (SETUP_MUTEX_NAME, "OkulRevirDefteri_Setup_Mutex_v1"):
		handle = kernel32.OpenMutexW(SYNCHRONIZE, False, m_name)
		if handle:
			try:
				kernel32.CloseHandle(handle)
			except Exception:
				pass
			return True
	return False


def acquire_single_instance_lock() -> bool:
	"""Windows Mutex ile ikinci uygulamanın veya aktif kurulum sihirbazının çakışmasını önler."""
	global _SINGLE_INSTANCE_HANDLE
	if sys.platform != "win32":
		return True
	if _is_setup_wizard_running():
		try:
			ctypes.windll.user32.MessageBoxW(
				0,
				"Okul Revir Defteri kurulum veya güncelleme sihirbazı şu anda çalışıyor.\n"
				"Güncelleme tamamlandıktan sonra programı yeniden açabilirsiniz.",
				APP_TITLE,
				0x00000040,
			)
		except Exception:
			pass
		return False
	kernel32 = ctypes.windll.kernel32
	kernel32.SetLastError(0)
	handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
	last_err = kernel32.GetLastError()
	ERROR_ALREADY_EXISTS = 183
	if handle and last_err == ERROR_ALREADY_EXISTS:
		try:
			kernel32.CloseHandle(handle)
		except Exception:
			pass
		_bring_existing_window_to_front()
		return False
	_SINGLE_INSTANCE_HANDLE = handle
	return True


def release_single_instance_lock() -> None:
	global _SINGLE_INSTANCE_HANDLE
	if sys.platform == "win32" and _SINGLE_INSTANCE_HANDLE:
		try:
			ctypes.windll.kernel32.CloseHandle(_SINGLE_INSTANCE_HANDLE)
		except Exception:
			pass
		_SINGLE_INSTANCE_HANDLE = None


def _ensure_streamlit_credentials() -> None:
	"""İlk kez kurulan bilgisayarlarda Streamlit e-posta istemini önler."""
	try:
		st_home = Path.home() / ".streamlit"
		cred_file = st_home / "credentials.toml"
		if not cred_file.exists():
			st_home.mkdir(parents=True, exist_ok=True)
			cred_file.write_text('[general]\nemail = ""\n', encoding="utf-8")
	except Exception:
		pass


def _is_port_in_use(port: int, host: str = "127.0.0.1") -> bool:
	with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
		s.settimeout(0.35)
		return s.connect_ex((host, int(port))) == 0


def find_free_port(preferred_port: int = DEFAULT_PORT, host: str = "127.0.0.1") -> int:
	"""127.0.0.1 üzerinde tercih edilen portu dener; meşgulse işletim sisteminden boş dinamik port alır."""
	if preferred_port > 0 and not _is_port_in_use(preferred_port, host):
		try:
			with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
				if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
					s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
				s.bind((host, int(preferred_port)))
				return int(preferred_port)
		except OSError:
			pass

	with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
		s.bind((host, 0))
		return int(s.getsockname()[1])


def _get_python_executable_for_dev() -> Path:
	py_dir = Path(sys.executable).resolve().parent
	pythonw = py_dir / "pythonw.exe"
	if pythonw.exists():
		return pythonw
	return Path(sys.executable).resolve()


def _run_streamlit_worker(port: int, script_path: Path) -> int:
	"""Alt süreç içinde Streamlit HTTP sunucusunu başlıksız ürün modunda doğrudan çalıştırır."""
	_ensure_standard_streams()
	_ensure_streamlit_credentials()

	script_dir = str(script_path.parent)
	if script_dir not in sys.path:
		sys.path.insert(0, script_dir)

	os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
	os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
	os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
	os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

	from streamlit.web import cli as stcli

	sys.argv = [
		"streamlit",
		"run",
		str(script_path),
		"--server.headless=true",
		"--server.fileWatcherType=none",
		f"--server.port={int(port)}",
		"--server.address=127.0.0.1",
		"--browser.gatherUsageStats=false",
		"--global.developmentMode=false",
		"--server.enableCORS=false",
		"--server.enableXsrfProtection=true",
		"--server.maxUploadSize=50",
		"--client.toolbarMode=minimal",
		"--client.showErrorDetails=false",
	]
	try:
		stcli.main()
	except SystemExit as exc:
		return int(exc.code) if isinstance(exc.code, int) else 0
	return 0


def start_streamlit_server(
	port: int,
	paths: dict[str, Path],
) -> tuple[subprocess.Popen, object]:
	"""Streamlit sunucusunu konsol penceresi açmadan (CREATE_NO_WINDOW) arka planda başlatır."""
	data_dir = paths["data_dir"]
	resource_dir = paths["resource_dir"]
	script_path = paths["script_path"]

	if not script_path.exists():
		raise FileNotFoundError(f"Uygulama dosyası bulunamadı: {script_path}")

	env = os.environ.copy()
	env["REVIR_DATA_DIR"] = str(data_dir)
	env["REVIR_RESOURCE_DIR"] = str(resource_dir)
	env["REVIR_WORKER_PORT"] = str(port)
	env["REVIR_WORKER_SCRIPT"] = str(script_path)
	env["STREAMLIT_SERVER_HEADLESS"] = "true"
	env["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"
	env["STREAMLIT_CLIENT_TOOLBAR_MODE"] = "minimal"
	env["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
	env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
	env["PYTHONIOENCODING"] = "utf-8"

	if getattr(sys, "frozen", False):
		cmd = [
			sys.executable,
			"--streamlit-worker",
			"--port",
			str(port),
			"--script",
			str(script_path),
		]
	else:
		py_exe = _get_python_executable_for_dev()
		cmd = [
			str(py_exe),
			str(Path(__file__).resolve()),
			"--streamlit-worker",
			"--port",
			str(port),
			"--script",
			str(script_path),
		]

	log_path = data_dir / ".revir_server.log"
	try:
		log_fp = open(log_path, "w", encoding="utf-8", errors="replace")
	except Exception:
		log_fp = open(os.devnull, "w", encoding="utf-8")

	popen_kwargs: dict[str, object] = {
		"cwd": str(data_dir),
		"env": env,
		"stdin": subprocess.DEVNULL,
		"stdout": log_fp,
		"stderr": log_fp,
	}
	if os.name == "nt":
		popen_kwargs["creationflags"] = CREATE_NO_WINDOW

	proc = subprocess.Popen(cmd, **popen_kwargs)
	return proc, log_fp


def wait_for_server(
	port: int,
	process: subprocess.Popen,
	timeout_seconds: float = 45.0,
	stop_event: threading.Event | None = None,
) -> bool:
	"""0.1s aralıklı hızlı soket kontrolü ve sağlık doğrulaması ile Streamlit sunucusunu bekler."""
	deadline = time.monotonic() + float(timeout_seconds)
	health_url = f"http://127.0.0.1:{port}/_stcore/health"
	root_url = f"http://127.0.0.1:{port}/"
	while time.monotonic() < deadline:
		if stop_event is not None and stop_event.is_set():
			return False
		if process.poll() is not None:
			return False
		port_open = False
		try:
			with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
				sock.settimeout(0.1)
				port_open = sock.connect_ex(("127.0.0.1", int(port))) == 0
		except OSError:
			port_open = False

		if port_open:
			for url in (health_url, root_url):
				try:
					with urllib.request.urlopen(url, timeout=0.6) as resp:
						if resp.status == 200:
							return True
				except (urllib.error.URLError, TimeoutError, OSError, ValueError):
					pass
		time.sleep(0.1)
	return False


def stop_streamlit_server(
	process: subprocess.Popen | None,
	log_fp: object | None = None,
) -> None:
	"""Arka plandaki Streamlit ve tüm alt süreçleri kalıntı (zombie) bırakmadan anında sonlandırır."""
	if process is not None and process.poll() is None:
		pid = process.pid
		if psutil is not None:
			try:
				parent = psutil.Process(pid)
				children = parent.children(recursive=True)
				for ch in children:
					try:
						ch.terminate()
					except psutil.Error:
						pass
				try:
					parent.terminate()
				except psutil.Error:
					pass
				_, alive = psutil.wait_procs([*children, parent], timeout=2.0)
				for proc_alive in alive:
					try:
						proc_alive.kill()
					except psutil.Error:
						pass
			except psutil.Error:
				pass

		if process.poll() is None:
			try:
				process.terminate()
				process.wait(timeout=2.0)
			except Exception:
				pass

		if process.poll() is None:
			try:
				if os.name == "nt":
					subprocess.run(
						["taskkill", "/F", "/T", "/PID", str(pid)],
						capture_output=True,
						timeout=5,
						creationflags=CREATE_NO_WINDOW,
					)
				else:
					os.kill(pid, signal.SIGKILL)
			except Exception:
				pass

		if process.poll() is None:
			try:
				process.kill()
				process.wait(timeout=1.5)
			except Exception:
				pass

	if log_fp is not None:
		try:
			log_fp.close()
		except Exception:
			pass


def _build_error_html(detail_text: str) -> str:
	safe_detail = html_lib.escape(detail_text or "Bilinmeyen başlatma hatası.")
	return f"""<!DOCTYPE html>
<html lang="tr">
<head><meta charset="utf-8" /><title>Başlatma Hatası</title>
<style>
  body {{ font-family: "Segoe UI", sans-serif; background: #f8fafc; color: #0f172a; display:flex; align-items:center; justify-content:center; height:100vh; margin:0; }}
  .box {{ background:#fff; border:1px solid #fecaca; border-radius:16px; padding:32px 36px; max-width:680px; width:90%; box-shadow:0 12px 32px rgba(0,0,0,0.08); }}
  h2 {{ color:#b91c1c; margin-top:0; font-size:20px; }}
  pre {{ background:#f1f5f9; padding:14px; border-radius:8px; overflow:auto; max-height:240px; font-size:12px; white-space:pre-wrap; }}
</style>
</head>
<body>
  <div class="box">
    <h2>Okul Revir Defteri Başlatılamadı</h2>
    <p>Yerel uygulama sunucusu yanıt vermedi. Pencereyi kapatıp programı yeniden başlatabilirsiniz.</p>
    <pre>{safe_detail}</pre>
  </div>
</body>
</html>"""


def launch_desktop_app() -> int:
	if not acquire_single_instance_lock():
		return 0

	_ensure_streamlit_credentials()
	paths = resolve_runtime_paths()
	data_dir = paths["data_dir"]
	icon_path = paths["icon_path"]

	webview_cache_dir = data_dir / ".webview2_cache"
	webview_cache_dir.mkdir(parents=True, exist_ok=True)
	os.environ["WEBVIEW2_USER_DATA_FOLDER"] = str(webview_cache_dir)
	os.environ["REVIR_DATA_DIR"] = str(data_dir)
	os.environ["REVIR_RESOURCE_DIR"] = str(paths["resource_dir"])

	import webview

	# Streamlit st.download_button (Word .docx, Excel .xlsx, .json) indirmelerinde Farklı Kaydet penceresini etkinleştir
	webview.settings["ALLOW_DOWNLOADS"] = True

	window = webview.create_window(
		title=APP_TITLE,
		html=SPLASH_HTML,
		width=WINDOW_WIDTH,
		height=WINDOW_HEIGHT,
		min_size=(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
		resizable=True,
		maximized=True,
	)

	port = find_free_port(DEFAULT_PORT)
	server_proc, log_fp = start_streamlit_server(port, paths)

	stop_event = threading.Event()
	cleanup_lock = threading.Lock()
	cleaned = {"done": False}

	def _cleanup() -> None:
		stop_event.set()
		with cleanup_lock:
			if cleaned["done"]:
				return
			cleaned["done"] = True
		stop_streamlit_server(server_proc, log_fp)
		release_single_instance_lock()

	atexit.register(_cleanup)

	def _on_window_closed(*_args) -> None:
		_cleanup()

	window.events.closed += _on_window_closed

	def _wait_and_navigate(win) -> None:
		ready = wait_for_server(port, server_proc, timeout_seconds=45.0, stop_event=stop_event)
		if stop_event.is_set():
			return
		if ready:
			try:
				win.load_url(f"http://127.0.0.1:{port}")
			except Exception:
				pass
			return

		log_tail = ""
		log_file = data_dir / ".revir_server.log"
		if log_file.exists():
			try:
				lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
				log_tail = "\n".join(lines[-15:])
			except Exception:
				pass
		try:
			win.load_html(_build_error_html(log_tail or "Sunucu zaman aşımına uğradı."))
		except Exception:
			pass

	icon_arg = str(icon_path) if icon_path.exists() else None
	try:
		try:
			webview.start(
				func=_wait_and_navigate,
				args=(window,),
				icon=icon_arg,
				gui="edgechromium",
				private_mode=False,
				storage_path=str(webview_cache_dir),
			)
		except Exception:
			webview.start(
				func=_wait_and_navigate,
				args=(window,),
				icon=icon_arg,
				private_mode=False,
				storage_path=str(webview_cache_dir),
			)
	finally:
		_cleanup()

	if os.environ.get("REVIR_NO_HARD_EXIT") != "1":
		os._exit(0)
	return 0


def main() -> int:
	if "--streamlit-worker" in sys.argv:
		parser = argparse.ArgumentParser(add_help=False)
		parser.add_argument("--streamlit-worker", action="store_true")
		parser.add_argument("--port", type=int, default=int(os.environ.get("REVIR_WORKER_PORT", DEFAULT_PORT)))
		parser.add_argument("--script", type=str, default=os.environ.get("REVIR_WORKER_SCRIPT", ""))
		args, _ = parser.parse_known_args()
		paths = resolve_runtime_paths()
		script_p = Path(args.script) if args.script else paths["script_path"]
		return _run_streamlit_worker(args.port, script_p)

	return launch_desktop_app()


if __name__ == "__main__":
	raise SystemExit(main())
