"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Bağımsız Masaüstü Başlatıcısı

pywebview (Microsoft Edge WebView2) penceresi içinde Streamlit sunucusunu
arka planda konsol açmadan başlatır, dosya indirmelerini etkinleştirir
ve pencere kapatıldığında arka plan sürecini tamamen sonlandırır.
"""
from __future__ import annotations

import argparse
import atexit
import os
import signal
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path


def _ensure_standard_streams() -> None:
	"""PyInstaller --noconsole / pythonw.exe altında None olabilen stdout/stderr akışlarını korur."""
	if sys.stdout is None:
		sys.stdout = open(os.devnull, "w", encoding="utf-8")
	if sys.stderr is None:
		sys.stderr = open(os.devnull, "w", encoding="utf-8")


_ensure_standard_streams()

APP_TITLE = "Okul Revir Defteri ve Sağlık Takip Sistemi"
DEFAULT_PORT = 8501
WINDOW_WIDTH = 1440
WINDOW_HEIGHT = 900
MIN_WINDOW_WIDTH = 1024
MIN_WINDOW_HEIGHT = 700
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)


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
	"""Geliştirme (.py) ve paketlenmiş (.exe) modlar için uygulama ve veri dizinlerini çözümler."""
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


def _ensure_streamlit_credentials() -> None:
	"""İlk kez çalıştırılan bilgisayarlarda Streamlit e-posta istemini önler."""
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
	"""Önce tercih edilen portu dener, doluysa işletim sisteminden boş dinamik bir TCP port alır."""
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
	"""Streamlit sunucusunu alt süreç içinde (başlıksız/ürün modunda) doğrudan çalıştırır."""
	_ensure_standard_streams()
	_ensure_streamlit_credentials()

	# PyInstaller içinde persantil.py yolunu ve bağımlılıkları garanti et
	script_dir = str(script_path.parent)
	if script_dir not in sys.path:
		sys.path.insert(0, script_dir)

	os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
	os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
	os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

	from streamlit.web import cli as stcli

	sys.argv = [
		"streamlit",
		"run",
		str(script_path),
		"--server.headless=true",
		"--global.developmentMode=false",
		"--browser.gatherUsageStats=false",
		"--server.enableCORS=false",
		"--server.enableXsrfProtection=true",
		"--server.maxUploadSize=50",
		"--client.toolbarMode=minimal",
		"--client.showErrorDetails=false",
		"--server.address=127.0.0.1",
		f"--server.port={int(port)}",
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
	"""Streamlit alt sürecini konsol penceresi açmadan başlatır."""
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
) -> bool:
	"""Streamlit yerel HTTP sunucusu yanıt verene kadar bekler."""
	deadline = time.monotonic() + float(timeout_seconds)
	urls = (
		f"http://127.0.0.1:{port}/_stcore/health",
		f"http://127.0.0.1:{port}/",
	)
	while time.monotonic() < deadline:
		if process.poll() is not None:
			return False
		for url in urls:
			try:
				with urllib.request.urlopen(url, timeout=1.0) as resp:
					if resp.status == 200:
						return True
			except (urllib.error.URLError, TimeoutError, OSError, ValueError):
				pass
		time.sleep(0.25)
	return False


def stop_streamlit_server(
	process: subprocess.Popen | None,
	log_fp: object | None = None,
) -> None:
	"""Arka plandaki Streamlit sürecini ve alt süreç ağacını kalıntı bırakmadan sonlandırır."""
	if process is not None and process.poll() is None:
		try:
			process.terminate()
			process.wait(timeout=3.0)
		except Exception:
			pass

		if process.poll() is None:
			try:
				if os.name == "nt":
					subprocess.run(
						["taskkill", "/F", "/T", "/PID", str(process.pid)],
						capture_output=True,
						timeout=5,
						creationflags=CREATE_NO_WINDOW,
					)
				else:
					os.kill(process.pid, signal.SIGKILL)
			except Exception:
				pass

		if process.poll() is None:
			try:
				process.kill()
				process.wait(timeout=2.0)
			except Exception:
				pass

	if log_fp is not None:
		try:
			log_fp.close()
		except Exception:
			pass


def _show_error_dialog(message: str) -> None:
	try:
		import ctypes
		ctypes.windll.user32.MessageBoxW(0, message, APP_TITLE, 0x10)
	except Exception:
		sys.stderr.write(f"{APP_TITLE}: {message}\n")


def launch_desktop_app() -> int:
	_ensure_streamlit_credentials()
	paths = resolve_runtime_paths()
	data_dir = paths["data_dir"]
	icon_path = paths["icon_path"]

	webview_cache_dir = data_dir / ".webview2_cache"
	webview_cache_dir.mkdir(parents=True, exist_ok=True)
	os.environ["WEBVIEW2_USER_DATA_FOLDER"] = str(webview_cache_dir)
	os.environ["REVIR_DATA_DIR"] = str(data_dir)
	os.environ["REVIR_RESOURCE_DIR"] = str(paths["resource_dir"])

	port = find_free_port(DEFAULT_PORT)
	server_proc, log_fp = start_streamlit_server(port, paths)

	cleanup_lock = threading.Lock()
	cleaned = {"done": False}

	def _cleanup() -> None:
		with cleanup_lock:
			if cleaned["done"]:
				return
			cleaned["done"] = True
		stop_streamlit_server(server_proc, log_fp)

	atexit.register(_cleanup)

	if not wait_for_server(port, server_proc, timeout_seconds=45.0):
		_cleanup()
		log_tail = ""
		log_file = data_dir / ".revir_server.log"
		if log_file.exists():
			try:
				lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
				log_tail = "\n".join(lines[-10:])
			except Exception:
				pass
		_show_error_dialog(
			"Revir Defteri arka plan sunucusu başlatılamadı.\n\n"
			+ (f"Ayrıntı:\n{log_tail}" if log_tail else "Lütfen log dosyasını kontrol edin.")
		)
		return 1

	import webview

	# Streamlit st.download_button (.docx, .xlsx, .json, .db) indirmelerini etkinleştir
	webview.settings["ALLOW_DOWNLOADS"] = True

	window = webview.create_window(
		title=APP_TITLE,
		url=f"http://127.0.0.1:{port}",
		width=WINDOW_WIDTH,
		height=WINDOW_HEIGHT,
		min_size=(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
		resizable=True,
	)

	def _on_window_closed(*_args) -> None:
		_cleanup()

	window.events.closed += _on_window_closed

	icon_arg = str(icon_path) if icon_path.exists() else None
	try:
		webview.start(
			icon=icon_arg,
			gui="edgechromium",
			private_mode=False,
			storage_path=str(webview_cache_dir),
		)
	except Exception:
		webview.start(
			icon=icon_arg,
			private_mode=False,
			storage_path=str(webview_cache_dir),
		)
	finally:
		_cleanup()

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
