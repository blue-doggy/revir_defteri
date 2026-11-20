"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Saf Python Masaüstü Yöneticisi (desktop_app.py)

Özellikler:
- Harici web çerçevesi kullanmadan doğrudan Python süreç-içi yerel arayüz motoru (revir_ui)
- Çapraz platform (macOS Cocoa WebKit / Windows Edge WebView2) yerel masaüstü penceresi (pywebview)
- Yerel işletim sistemi "Farklı Kaydet" penceresi ile Word/Excel/HTML/JSON/DB indirme köprüsü
- Tek Örnek (Single Instance) koruması (Windows Mutex / macOS-Linux PID kilidi)
- Pencere kapatıldığında kalıntısız ve anında kapanış
"""
from __future__ import annotations

import atexit
import os
import signal
import sys
import threading
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

_DEFAULT_APP_VERSION = "1.5.1"


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
WINDOW_WIDTH = 1440
WINDOW_HEIGHT = 900
MIN_WINDOW_WIDTH = 1024
MIN_WINDOW_HEIGHT = 700

_SINGLE_INSTANCE_HANDLE = None
_LOCK_FILE_PATH: Path | None = None


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
		if sys.platform == "win32":
			local_base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
		elif sys.platform == "darwin":
			local_base = Path.home() / "Library" / "Application Support"
		else:
			local_base = Path.home() / ".local" / "share"
		data_dir = local_base / "RevirDefteri"
		data_dir.mkdir(parents=True, exist_ok=True)

	icon_name = "app_icon.ico" if sys.platform == "win32" else "app_icon.png"
	icon_candidates = (exe_dir / icon_name, resource_dir / icon_name)
	icon_path = next((p for p in icon_candidates if p.exists()), resource_dir / icon_name)

	return {
		"exe_dir": exe_dir,
		"resource_dir": resource_dir,
		"data_dir": data_dir,
		"icon_path": icon_path,
	}


def _bring_existing_window_to_front(title_keyword: str = "Okul Revir Defteri") -> bool:
	"""Windows üzerinde halihazırda açık olan pencereyi öne getirir."""
	if sys.platform != "win32":
		return False
	try:
		import ctypes
		from ctypes import wintypes

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
			user32.ShowWindow(hwnd, SW_RESTORE if user32.IsIconic(hwnd) else SW_SHOW)
			user32.SetForegroundWindow(hwnd)
			return True
	except Exception:
		pass
	return False


def acquire_single_instance_lock(data_dir: Path) -> bool:
	"""Uygulamanın aynı anda tek bir masaüstü örneği olarak çalışmasını sağlar."""
	global _SINGLE_INSTANCE_HANDLE, _LOCK_FILE_PATH
	if sys.platform == "win32":
		try:
			import ctypes

			kernel32 = ctypes.windll.kernel32
			ERROR_ALREADY_EXISTS = 183
			handle = kernel32.CreateMutexW(None, True, MUTEX_NAME)
			last_err = kernel32.GetLastError()
			if not handle:
				return True
			if last_err == ERROR_ALREADY_EXISTS:
				kernel32.CloseHandle(handle)
				_bring_existing_window_to_front("Okul Revir Defteri")
				return False
			_SINGLE_INSTANCE_HANDLE = handle
		except Exception:
			pass
		return True

	# macOS / Linux PID kilit dosyası
	lock_path = data_dir / ".revir_desktop.pid"
	if lock_path.exists():
		try:
			old_pid = int(lock_path.read_text(encoding="utf-8").strip())
			if old_pid != os.getpid():
				os.kill(old_pid, 0)
				# Süreç hâlâ yaşıyor -> yine de temiz başlatmaya izin ver ya da öne getir
				if psutil is not None:
					try:
						proc = psutil.Process(old_pid)
						cmdline = " ".join(proc.cmdline()).lower()
						if "desktop_app" in cmdline or "persantil" in cmdline or "revirdefteri" in cmdline:
							return False
					except Exception:
						pass
		except (ValueError, ProcessLookupError, PermissionError, OSError):
			pass
	try:
		lock_path.write_text(str(os.getpid()), encoding="utf-8")
		_LOCK_FILE_PATH = lock_path
	except Exception:
		pass
	return True


def release_single_instance_lock() -> None:
	global _SINGLE_INSTANCE_HANDLE, _LOCK_FILE_PATH
	if sys.platform == "win32" and _SINGLE_INSTANCE_HANDLE is not None:
		try:
			import ctypes
			kernel32 = ctypes.windll.kernel32
			kernel32.ReleaseMutex(_SINGLE_INSTANCE_HANDLE)
			kernel32.CloseHandle(_SINGLE_INSTANCE_HANDLE)
		except Exception:
			pass
		_SINGLE_INSTANCE_HANDLE = None

	if _LOCK_FILE_PATH is not None:
		try:
			if _LOCK_FILE_PATH.exists():
				_LOCK_FILE_PATH.unlink(missing_ok=True)
		except Exception:
			pass
		_LOCK_FILE_PATH = None


class _RevirDesktopBridge:
	"""pywebview penceresi ile Python yerel dosya sistemi arasında güvenli köprü."""
	def __init__(self, server, window_holder: dict) -> None:
		self._server = server
		self._window_holder = window_holder

	def save_download(self, token: str, suggested_name: str = "dosya.bin") -> dict:
		item = self._server.get_download_item(str(token))
		if item is None:
			return {"saved": False, "error": "İndirme verisi bulunamadı."}
		raw_bytes, file_name, _mime = item
		target_name = suggested_name or file_name or "dosya.bin"
		win = self._window_holder.get("window")
		if win is None:
			return {"saved": False, "error": "Pencere hazır değil."}
		try:
			import webview
			dlg_type = getattr(getattr(webview, "FileDialog", None), "SAVE", None) or getattr(webview, "SAVE_DIALOG", 30)
			dest = win.create_file_dialog(
				dlg_type,
				directory=str(Path.home() / "Downloads"),
				save_filename=target_name,
			)
			if not dest:
				return {"saved": False, "cancelled": True}
			out_path = Path(dest[0] if isinstance(dest, (list, tuple)) else dest)
			out_path.write_bytes(raw_bytes)
			return {"saved": True, "path": str(out_path)}
		except Exception as exc:
			return {"saved": False, "error": str(exc)}


def launch_desktop_app() -> int:
	paths = resolve_runtime_paths()
	data_dir = paths["data_dir"]
	resource_dir = paths["resource_dir"]
	icon_path = paths["icon_path"]

	if not acquire_single_instance_lock(data_dir):
		return 0

	os.environ["REVIR_DATA_DIR"] = str(data_dir)
	os.environ["REVIR_RESOURCE_DIR"] = str(resource_dir)

	if str(resource_dir) not in sys.path:
		sys.path.insert(0, str(resource_dir))

	import revir_ui
	import persantil

	server = revir_ui.LocalUIServer(persantil.main, host="127.0.0.1", port=0)
	app_url = server.start()

	cleanup_lock = threading.Lock()
	cleaned = {"done": False}

	def _cleanup(*_args) -> None:
		with cleanup_lock:
			if cleaned["done"]:
				return
			cleaned["done"] = True
		server.stop()
		release_single_instance_lock()

	atexit.register(_cleanup)
	for sig_name in ("SIGINT", "SIGTERM"):
		sig_val = getattr(signal, sig_name, None)
		if sig_val is not None:
			try:
				signal.signal(sig_val, lambda *_a: (_cleanup(), os._exit(0)))
			except Exception:
				pass

	import webview

	webview.settings["ALLOW_DOWNLOADS"] = True

	window_holder: dict = {"window": None}
	bridge = _RevirDesktopBridge(server, window_holder)

	create_kwargs = dict(
		title=APP_TITLE,
		url=app_url,
		js_api=bridge,
		width=WINDOW_WIDTH,
		height=WINDOW_HEIGHT,
		min_size=(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
		resizable=True,
	)
	if sys.platform == "win32":
		create_kwargs["maximized"] = True

	window = webview.create_window(**create_kwargs)
	window_holder["window"] = window
	window.events.closed += _cleanup

	start_kwargs: dict = {"private_mode": False}
	if sys.platform == "win32":
		webview_cache_dir = data_dir / ".webview2_cache"
		webview_cache_dir.mkdir(parents=True, exist_ok=True)
		os.environ["WEBVIEW2_USER_DATA_FOLDER"] = str(webview_cache_dir)
		start_kwargs["storage_path"] = str(webview_cache_dir)
		start_kwargs["gui"] = "edgechromium"
		if icon_path.exists():
			start_kwargs["icon"] = str(icon_path)

	try:
		try:
			webview.start(**start_kwargs)
		except Exception:
			start_kwargs.pop("gui", None)
			start_kwargs.pop("icon", None)
			webview.start(**start_kwargs)
	finally:
		_cleanup()

	if os.environ.get("REVIR_NO_HARD_EXIT") != "1":
		os._exit(0)
	return 0


def main() -> int:
	return launch_desktop_app()


if __name__ == "__main__":
	raise SystemExit(main())
