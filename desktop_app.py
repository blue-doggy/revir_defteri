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
import ctypes
import json
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

_DEFAULT_APP_VERSION = "1.49.0"


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
MIN_WINDOW_WIDTH = 900
MIN_WINDOW_HEIGHT = 600

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
	elif sys.platform == "darwin" and is_frozen:
		# macOS .app paketi icinde veritabani tutmak guncellemede veri kaybina yol acar;
		# her zaman ~/Library/Application Support/RevirDefteri kullanilir.
		data_dir = Path.home() / "Library" / "Application Support" / "RevirDefteri"
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
	def __init__(self, server, window_holder: dict, taskbar_toolbar=None) -> None:
		self._server = server
		self._window_holder = window_holder
		self._taskbar_toolbar = taskbar_toolbar

	def sync_taskbar_media_state(self, is_playing: bool = False, title: str = "", visible: bool = True) -> dict:
		""" Görev çubuğu önizleme penceresindeki mini oynatıcı düğmelerini (▶/⏸ ve başlık) günceller. """
		tb = self._taskbar_toolbar
		if tb is not None:
			try:
				tb.sync_state(bool(is_playing), str(title or ""), bool(visible))
				return {"ok": True}
			except Exception as exc:
				return {"ok": False, "error": str(exc)}
		return {"ok": False}

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

	def select_music_folder(self, current_dir: str = "") -> dict:
		"""Yerel klasör seçici penceresini açar ve seçilen müzik klasörü yolunu döndürür."""
		win = self._window_holder.get("window")
		start_dir = str(current_dir or "").strip()
		if not start_dir or not os.path.isdir(start_dir):
			start_dir = str(Path.home())
		if win is not None:
			try:
				import webview
				dlg_type = getattr(getattr(webview, "FileDialog", None), "FOLDER", None) or getattr(webview, "FOLDER_DIALOG", 20)
				dest = win.create_file_dialog(dlg_type, directory=start_dir)
				if not dest:
					return {"selected": False, "cancelled": True}
				chosen = Path(dest[0] if isinstance(dest, (list, tuple)) else dest).resolve()
				if chosen.is_dir():
					return {"selected": True, "path": str(chosen)}
			except Exception:
				pass
		try:
			import media_player
			fallback_dir = media_player.pick_music_folder_native(start_dir)
			if fallback_dir:
				return {"selected": True, "path": fallback_dir}
		except Exception as exc:
			return {"selected": False, "error": str(exc)}
		return {"selected": False, "cancelled": True}


def _start_windows_media_key_listener(window_holder: dict, stop_event: threading.Event) -> threading.Thread | None:
	"""
	Windows'ta pywebview penceresi simge durumuna küçültüldüğünde veya arka plandayken bile
	fiziksel klavye multimedya tuşlarını (Play/Pause, Next, Previous, Stop) yakalayıp
	window.__revirMediaControl(...) köprüsüne iletir.
	"""
	if sys.platform != "win32":
		return None
	try:
		user32 = ctypes.windll.user32
	except Exception:
		return None

	vk_map = (
		(0xB3, "play_pause"),  # VK_MEDIA_PLAY_PAUSE
		(0xB0, "next"),        # VK_MEDIA_NEXT_TRACK
		(0xB1, "prev"),        # VK_MEDIA_PREV_TRACK
		(0xB2, "stop"),        # VK_MEDIA_STOP
	)

	def _worker() -> None:
		prev_down = {vk: False for vk, _ in vk_map}
		# Başlangıçtaki basılı durumları temizle
		for vk, _ in vk_map:
			try:
				prev_down[vk] = bool(user32.GetAsyncKeyState(vk) & 0x8000)
			except Exception:
				prev_down[vk] = False

		while not stop_event.wait(0.075):
			win = window_holder.get("window")
			if win is None:
				continue
			for vk, cmd in vk_map:
				try:
					is_down = bool(user32.GetAsyncKeyState(vk) & 0x8000)
				except Exception:
					is_down = False
				if is_down and not prev_down[vk]:
					try:
						win.evaluate_js(f"window.__revirMediaControl && window.__revirMediaControl('{cmd}')")
					except Exception:
						pass
				prev_down[vk] = is_down

	th = threading.Thread(target=_worker, daemon=True, name="RevirGlobalMediaKeyListener")
	th.start()
	return th


class _WindowsTaskbarMediaToolbar:
	"""
	Windows 10/11 Başlat / Görev Çubuğunda uygulama simgesinin üzerine fare getirildiğinde
	açılan canlı pencere önizlemesinin (Thumbnail Preview) altına ⏮ Önceki, ▶ Oynat / ⏸ Duraklat
	ve ⏭ Sonraki mini oynatıcı düğmelerini (ITaskbarList3::ThumbBarAddButtons) ekler.
	"""
	BTN_PREV = 1001
	BTN_PLAY_PAUSE = 1002
	BTN_NEXT = 1003
	SUBCLASS_ID = 0x5245564D  # 'REVM'
	WM_COMMAND = 0x0111
	THBN_CLICKED = 0x1800
	THB_ICON = 0x0002
	THB_TOOLTIP = 0x0004
	THB_FLAGS = 0x0008
	THBF_ENABLED = 0x0000
	THBF_HIDDEN = 0x0008

	def __init__(self, window_holder: dict, stop_event: threading.Event) -> None:
		self._window_holder = window_holder
		self._stop_event = stop_event
		self._lock = threading.RLock()
		self._is_playing: bool = False
		self._title: str = ""
		self._visible: bool = True
		self._hwnd: int = 0
		self._form = None
		self._taskbar_ptr = None
		self._vtable = None
		self._buttons_added: bool = False
		self._subclassed: bool = False
		self._subclass_proc_type = None
		self._subclass_proc_ref = None
		self._icons: dict[str, int] = {}
		self._wm_taskbar_created: int = 0

	@staticmethod
	def _create_media_hicon(kind: str, size: int = 0) -> int:
		if sys.platform != "win32":
			return 0
		from ctypes import wintypes
		user32 = ctypes.windll.user32
		gdi32 = ctypes.windll.gdi32
		if size <= 0:
			try:
				size = int(user32.GetSystemMetrics(49)) or 20  # SM_CXSMICON
			except Exception:
				size = 20
		size = max(16, min(48, int(size)))

		bgra: bytes
		try:
			from PIL import Image, ImageDraw
			ss = size * 4
			img = Image.new("RGBA", (ss, ss), (0, 0, 0, 0))
			drw = ImageDraw.Draw(img)
			pad = max(2, int(ss * 0.05))
			bg_col = (13, 148, 136, 242) if kind in ("play", "pause") else (30, 41, 59, 232)
			drw.ellipse([pad, pad, ss - pad, ss - pad], fill=bg_col)
			fg = (255, 255, 255, 255)
			if kind == "play":
				drw.polygon([(ss * 0.38, ss * 0.26), (ss * 0.38, ss * 0.74), (ss * 0.75, ss * 0.50)], fill=fg)
			elif kind == "pause":
				r_rad = max(2, int(ss * 0.04))
				drw.rounded_rectangle([ss * 0.31, ss * 0.27, ss * 0.45, ss * 0.73], radius=r_rad, fill=fg)
				drw.rounded_rectangle([ss * 0.55, ss * 0.27, ss * 0.69, ss * 0.73], radius=r_rad, fill=fg)
			elif kind == "prev":
				drw.rounded_rectangle([ss * 0.25, ss * 0.28, ss * 0.34, ss * 0.72], radius=2, fill=fg)
				drw.polygon([(ss * 0.71, ss * 0.28), (ss * 0.71, ss * 0.72), (ss * 0.37, ss * 0.50)], fill=fg)
			elif kind == "next":
				drw.polygon([(ss * 0.29, ss * 0.28), (ss * 0.29, ss * 0.72), (ss * 0.63, ss * 0.50)], fill=fg)
				drw.rounded_rectangle([ss * 0.66, ss * 0.28, ss * 0.75, ss * 0.72], radius=2, fill=fg)
			img = img.resize((size, size), Image.Resampling.LANCZOS)
			r, g, b, a = img.split()
			bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
		except Exception:
			buf = bytearray(size * size * 4)
			cx = cy = (size - 1) / 2.0
			rad = size * 0.44
			for y in range(size):
				for x in range(size):
					if (x - cx) ** 2 + (y - cy) ** 2 <= rad ** 2:
						off = (y * size + x) * 4
						buf[off : off + 4] = bytes((136, 148, 13, 240))
			bgra = bytes(buf)

		class _BITMAPINFOHEADER(ctypes.Structure):
			_fields_ = [
				("biSize", ctypes.c_uint32),
				("biWidth", ctypes.c_int32),
				("biHeight", ctypes.c_int32),
				("biPlanes", ctypes.c_uint16),
				("biBitCount", ctypes.c_uint16),
				("biCompression", ctypes.c_uint32),
				("biSizeImage", ctypes.c_uint32),
				("biXPelsPerMeter", ctypes.c_int32),
				("biYPelsPerMeter", ctypes.c_int32),
				("biClrUsed", ctypes.c_uint32),
				("biClrImportant", ctypes.c_uint32),
			]

		class _ICONINFO(ctypes.Structure):
			_fields_ = [
				("fIcon", wintypes.BOOL),
				("xHotspot", wintypes.DWORD),
				("yHotspot", wintypes.DWORD),
				("hbmMask", wintypes.HBITMAP),
				("hbmColor", wintypes.HBITMAP),
			]

		user32.GetDC.argtypes = [wintypes.HWND]
		user32.GetDC.restype = wintypes.HDC
		user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
		user32.ReleaseDC.restype = ctypes.c_int
		gdi32.CreateDIBSection.argtypes = [
			wintypes.HDC, ctypes.POINTER(_BITMAPINFOHEADER), ctypes.c_uint,
			ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, ctypes.c_uint32,
		]
		gdi32.CreateDIBSection.restype = wintypes.HBITMAP
		gdi32.CreateBitmap.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p]
		gdi32.CreateBitmap.restype = wintypes.HBITMAP
		gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
		gdi32.DeleteObject.restype = wintypes.BOOL
		user32.CreateIconIndirect.argtypes = [ctypes.POINTER(_ICONINFO)]
		user32.CreateIconIndirect.restype = wintypes.HICON
		user32.DestroyIcon.argtypes = [wintypes.HICON]
		user32.DestroyIcon.restype = wintypes.BOOL

		bmi = _BITMAPINFOHEADER()
		bmi.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
		bmi.biWidth = size
		bmi.biHeight = -size
		bmi.biPlanes = 1
		bmi.biBitCount = 32
		bmi.biCompression = 0

		bits = ctypes.c_void_p()
		hdc = user32.GetDC(None)
		hbm_color = gdi32.CreateDIBSection(hdc, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
		user32.ReleaseDC(None, hdc)
		if not hbm_color or not bits.value:
			return 0
		ctypes.memmove(bits, bgra, len(bgra))
		hbm_mask = gdi32.CreateBitmap(size, size, 1, 1, None)

		ii = _ICONINFO(True, 0, 0, hbm_mask, hbm_color)
		hicon = user32.CreateIconIndirect(ctypes.byref(ii))
		gdi32.DeleteObject(hbm_color)
		if hbm_mask:
			gdi32.DeleteObject(hbm_mask)
		return int(hicon or 0)

	def _dispatch_media_cmd(self, cmd: str) -> None:
		def _send() -> None:
			win = self._window_holder.get("window")
			if win is None:
				return
			try:
				win.evaluate_js(f"window.__revirMediaControl && window.__revirMediaControl('{cmd}', true)")
			except Exception:
				pass

		threading.Thread(target=_send, daemon=True).start()

	def install_subclass_on_hwnd(self, hwnd: int) -> bool:
		"""Belirtilen HWND üzerine görev çubuğu buton tıklamalarını dinleyen alt sınıf kancasını kurar."""
		if sys.platform != "win32" or not hwnd:
			return False
		from ctypes import wintypes
		user32 = ctypes.windll.user32
		comctl32 = ctypes.windll.comctl32

		LRESULT = ctypes.c_ssize_t
		WPARAM = ctypes.c_size_t
		LPARAM = ctypes.c_ssize_t
		UINT_PTR = ctypes.c_size_t
		DWORD_PTR = ctypes.c_size_t
		SUBCLASSPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR)

		comctl32.SetWindowSubclass.argtypes = [wintypes.HWND, SUBCLASSPROC, UINT_PTR, DWORD_PTR]
		comctl32.SetWindowSubclass.restype = wintypes.BOOL
		comctl32.RemoveWindowSubclass.argtypes = [wintypes.HWND, SUBCLASSPROC, UINT_PTR]
		comctl32.RemoveWindowSubclass.restype = wintypes.BOOL
		comctl32.DefSubclassProc.argtypes = [wintypes.HWND, wintypes.UINT, WPARAM, LPARAM]
		comctl32.DefSubclassProc.restype = LRESULT
		user32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]
		user32.RegisterWindowMessageW.restype = wintypes.UINT

		try:
			self._wm_taskbar_created = int(user32.RegisterWindowMessageW("TaskbarButtonCreated") or 0)
			if hasattr(user32, "ChangeWindowMessageFilterEx"):
				user32.ChangeWindowMessageFilterEx.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.DWORD, ctypes.c_void_p]
				user32.ChangeWindowMessageFilterEx.restype = wintypes.BOOL
				if self._wm_taskbar_created:
					user32.ChangeWindowMessageFilterEx(wintypes.HWND(hwnd), self._wm_taskbar_created, 1, None)
				user32.ChangeWindowMessageFilterEx(wintypes.HWND(hwnd), self.WM_COMMAND, 1, None)
		except Exception:
			pass

		toolbar_ref = self

		def _wnd_subclass(h_wnd, u_msg, w_param, l_param, _uid, _ref):
			if toolbar_ref._wm_taskbar_created and u_msg == toolbar_ref._wm_taskbar_created:
				toolbar_ref._buttons_added = False
				toolbar_ref._apply_thumbbar_on_ui_thread()
			elif u_msg == toolbar_ref.WM_COMMAND and (((w_param >> 16) & 0xFFFF) == toolbar_ref.THBN_CLICKED):
				btn_id = int(w_param & 0xFFFF)
				if btn_id == toolbar_ref.BTN_PREV:
					toolbar_ref._dispatch_media_cmd("prev")
					return 0
				if btn_id == toolbar_ref.BTN_PLAY_PAUSE:
					toolbar_ref._dispatch_media_cmd("play_pause")
					return 0
				if btn_id == toolbar_ref.BTN_NEXT:
					toolbar_ref._dispatch_media_cmd("next")
					return 0
			return comctl32.DefSubclassProc(h_wnd, u_msg, w_param, l_param)

		self._subclass_proc_type = SUBCLASSPROC
		self._subclass_proc_ref = SUBCLASSPROC(_wnd_subclass)
		ok = bool(comctl32.SetWindowSubclass(wintypes.HWND(hwnd), self._subclass_proc_ref, self.SUBCLASS_ID, 0))
		if ok:
			self._hwnd = int(hwnd)
			self._subclassed = True
		return ok

	def _ensure_com_and_icons_on_ui_thread(self) -> bool:
		if sys.platform != "win32":
			return False
		if not self._icons:
			for k in ("prev", "play", "pause", "next"):
				hi = self._create_media_hicon(k)
				if hi:
					self._icons[k] = hi
		if self._taskbar_ptr is not None and self._vtable is not None:
			return True

		class _GUID(ctypes.Structure):
			_fields_ = [
				("Data1", ctypes.c_uint32),
				("Data2", ctypes.c_uint16),
				("Data3", ctypes.c_uint16),
				("Data4", ctypes.c_ubyte * 8),
			]

		clsid = _GUID(0x56FDF344, 0xFD6D, 0x11D0, (ctypes.c_ubyte * 8)(0x95, 0x8A, 0x00, 0x60, 0x97, 0xC9, 0xA0, 0x90))
		iid = _GUID(0xEA1AFB91, 0x9E28, 0x4B86, (ctypes.c_ubyte * 8)(0x90, 0xE9, 0x9E, 0x9F, 0x8A, 0x5E, 0xEF, 0xAF))
		ole32 = ctypes.windll.ole32
		ole32.CoInitialize.argtypes = [ctypes.c_void_p]
		ole32.CoInitialize.restype = ctypes.c_long
		ole32.CoCreateInstance.argtypes = [
			ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32,
			ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p),
		]
		ole32.CoCreateInstance.restype = ctypes.c_long
		try:
			ole32.CoInitialize(None)
		except Exception:
			pass
		ptr = ctypes.c_void_p()
		hr = ole32.CoCreateInstance(ctypes.byref(clsid), None, 23, ctypes.byref(iid), ctypes.byref(ptr))
		if hr != 0 or not ptr.value:
			return False
		vt = ctypes.cast(ctypes.cast(ptr, ctypes.POINTER(ctypes.c_void_p))[0], ctypes.POINTER(ctypes.c_void_p))
		hr_init = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)(vt[3])(ptr)
		if hr_init != 0:
			ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vt[2])(ptr)
			return False
		self._taskbar_ptr = ptr
		self._vtable = vt
		return True

	def _build_thumb_buttons_array(self):
		class _THUMBBUTTON(ctypes.Structure):
			_fields_ = [
				("dwMask", ctypes.c_uint32),
				("iId", ctypes.c_uint32),
				("iBitmap", ctypes.c_uint32),
				("hIcon", ctypes.c_void_p),
				("szTip", ctypes.c_wchar * 260),
				("dwFlags", ctypes.c_uint32),
			]

		with self._lock:
			playing = self._is_playing
			vis = self._visible
		flags = self.THBF_ENABLED if vis else self.THBF_HIDDEN
		mask = self.THB_ICON | self.THB_TOOLTIP | self.THB_FLAGS

		arr = (_THUMBBUTTON * 3)()
		arr[0].dwMask = mask
		arr[0].iId = self.BTN_PREV
		arr[0].hIcon = self._icons.get("prev", 0)
		arr[0].szTip = "Önceki Kanal / Parça"
		arr[0].dwFlags = flags

		arr[1].dwMask = mask
		arr[1].iId = self.BTN_PLAY_PAUSE
		arr[1].hIcon = self._icons.get("pause" if playing else "play", 0)
		arr[1].szTip = "Duraklat" if playing else "Oynat"
		arr[1].dwFlags = flags

		arr[2].dwMask = mask
		arr[2].iId = self.BTN_NEXT
		arr[2].hIcon = self._icons.get("next", 0)
		arr[2].szTip = "Sonraki Kanal / Parça"
		arr[2].dwFlags = flags
		return arr

	def _apply_thumbbar_on_ui_thread(self) -> None:
		if not self._hwnd or not self._ensure_com_and_icons_on_ui_thread():
			return
		from ctypes import wintypes
		arr = self._build_thumb_buttons_array()
		try:
			if not self._buttons_added:
				add_fn = ctypes.WINFUNCTYPE(
					ctypes.c_long, ctypes.c_void_p, wintypes.HWND, ctypes.c_uint, ctypes.c_void_p
				)(self._vtable[15])
				hr = add_fn(self._taskbar_ptr, wintypes.HWND(self._hwnd), 3, ctypes.byref(arr))
				if hr == 0:
					self._buttons_added = True
			else:
				upd_fn = ctypes.WINFUNCTYPE(
					ctypes.c_long, ctypes.c_void_p, wintypes.HWND, ctypes.c_uint, ctypes.c_void_p
				)(self._vtable[16])
				upd_fn(self._taskbar_ptr, wintypes.HWND(self._hwnd), 3, ctypes.byref(arr))

			with self._lock:
				tip = f"{APP_TITLE} · {'▶ ' if self._is_playing else ''}{self._title}".strip(" ·") if self._title else APP_TITLE
			tip_fn = ctypes.WINFUNCTYPE(
				ctypes.c_long, ctypes.c_void_p, wintypes.HWND, wintypes.LPCWSTR
			)(self._vtable[19])
			tip_fn(self._taskbar_ptr, wintypes.HWND(self._hwnd), tip[:255])
		except Exception:
			pass

	def _invoke_on_form_ui(self, fn) -> None:
		form = self._form
		if form is None:
			return
		try:
			from System import Action  # type: ignore[import-not-found]
			if getattr(form, "IsDisposed", False) or not getattr(form, "IsHandleCreated", False):
				return
			if getattr(form, "InvokeRequired", False):
				form.BeginInvoke(Action(fn))
			else:
				fn()
		except Exception:
			pass

	def start_attaching(self) -> None:
		if sys.platform != "win32":
			return

		def _wait_and_attach() -> None:
			for _ in range(120):
				if self._stop_event.wait(0.15):
					return
				win = self._window_holder.get("window")
				form = getattr(win, "native", None) if win is not None else None
				if form is not None and getattr(form, "IsHandleCreated", False):
					self._form = form

					def _init_ui() -> None:
						try:
							hwnd = int(form.Handle.ToInt64())
							if hwnd and not self._subclassed:
								self.install_subclass_on_hwnd(hwnd)
							self._apply_thumbbar_on_ui_thread()
						except Exception:
							pass

					self._invoke_on_form_ui(_init_ui)
					# Pencere görev çubuğuna tam yerleştikten hemen sonra bir kez daha doğrula
					if not self._stop_event.wait(0.6):
						self._invoke_on_form_ui(self._apply_thumbbar_on_ui_thread)
					return

		threading.Thread(target=_wait_and_attach, daemon=True, name="RevirTaskbarThumbAttach").start()

	def sync_state(self, is_playing: bool, title: str = "", visible: bool = True) -> None:
		with self._lock:
			self._is_playing = bool(is_playing)
			self._title = str(title or "").strip()
			self._visible = bool(visible)
		if sys.platform == "win32" and self._form is not None:
			self._invoke_on_form_ui(self._apply_thumbbar_on_ui_thread)

	def close(self) -> None:
		if sys.platform != "win32":
			return

		def _teardown_ui() -> None:
			from ctypes import wintypes
			if self._subclassed and self._hwnd and self._subclass_proc_ref is not None:
				try:
					comctl32 = ctypes.windll.comctl32
					if self._subclass_proc_type is not None:
						comctl32.RemoveWindowSubclass.argtypes = [
							wintypes.HWND, self._subclass_proc_type, ctypes.c_size_t
						]
						comctl32.RemoveWindowSubclass.restype = wintypes.BOOL
					comctl32.RemoveWindowSubclass(wintypes.HWND(self._hwnd), self._subclass_proc_ref, self.SUBCLASS_ID)
				except Exception:
					pass
				self._subclassed = False
			if self._taskbar_ptr is not None and self._vtable is not None:
				try:
					ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(self._vtable[2])(self._taskbar_ptr)
				except Exception:
					pass
				self._taskbar_ptr = None
				self._vtable = None
			user32 = ctypes.windll.user32
			user32.DestroyIcon.argtypes = [wintypes.HICON]
			user32.DestroyIcon.restype = wintypes.BOOL
			for hi in list(self._icons.values()):
				try:
					if hi:
						user32.DestroyIcon(wintypes.HICON(hi))
				except Exception:
					pass
			self._icons.clear()

		if self._form is not None:
			self._invoke_on_form_ui(_teardown_ui)
		else:
			_teardown_ui()


def _load_window_state(data_dir: Path) -> dict[str, int]:
	"""Son kullanılan pencere boyutunu ve konumunu yerel dosyadan yükler."""
	st_path = data_dir / ".window_state.json"
	if not st_path.exists():
		return {}
	try:
		raw = json.loads(st_path.read_text(encoding="utf-8"))
		if not isinstance(raw, dict):
			return {}
		out: dict[str, int] = {}
		w = int(raw.get("width", 0))
		h = int(raw.get("height", 0))
		if MIN_WINDOW_WIDTH <= w <= 3840:
			out["width"] = w
		if MIN_WINDOW_HEIGHT <= h <= 2160:
			out["height"] = h
		if "x" in raw and "y" in raw:
			x = int(raw["x"])
			y = int(raw["y"])
			if -100 <= x <= 3200 and 0 <= y <= 1800:
				out["x"] = x
				out["y"] = y
		return out
	except Exception:
		return {}


def _save_window_state(data_dir: Path, win) -> None:
	"""Geçerli pencere boyutunu ve konumunu yerel dosyaya kaydeder."""
	if win is None:
		return
	try:
		w = int(getattr(win, "width", 0) or 0)
		h = int(getattr(win, "height", 0) or 0)
		if w < MIN_WINDOW_WIDTH or h < MIN_WINDOW_HEIGHT:
			return
		payload: dict[str, int] = {"width": w, "height": h}
		x = getattr(win, "x", None)
		y = getattr(win, "y", None)
		if x is not None and y is not None:
			payload["x"] = int(x)
			payload["y"] = int(y)
		(data_dir / ".window_state.json").write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
	except Exception:
		pass


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

	window_holder: dict = {"window": None}
	media_key_stop = threading.Event()
	taskbar_toolbar = _WindowsTaskbarMediaToolbar(window_holder, media_key_stop)
	cleanup_lock = threading.Lock()
	cleaned = {"done": False}

	def _cleanup(*_args) -> None:
		with cleanup_lock:
			if cleaned["done"]:
				return
			cleaned["done"] = True
		media_key_stop.set()
		try:
			taskbar_toolbar.close()
		except Exception:
			pass
		_save_window_state(data_dir, window_holder.get("window"))
		try:
			import media_player
			if getattr(media_player, "_GLOBAL_MEDIA_SERVER", None) is not None:
				media_player._GLOBAL_MEDIA_SERVER.stop()
		except Exception:
			pass
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

	bridge = _RevirDesktopBridge(server, window_holder, taskbar_toolbar=taskbar_toolbar)
	saved_win = _load_window_state(data_dir)

	create_kwargs = dict(
		title=APP_TITLE,
		url=app_url,
		js_api=bridge,
		width=saved_win.get("width", WINDOW_WIDTH),
		height=saved_win.get("height", WINDOW_HEIGHT),
		min_size=(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT),
		resizable=True,
		text_select=True,
	)
	if "x" in saved_win and "y" in saved_win:
		create_kwargs["x"] = saved_win["x"]
		create_kwargs["y"] = saved_win["y"]
	if sys.platform == "win32" and not saved_win:
		create_kwargs["maximized"] = True

	window = webview.create_window(**create_kwargs)
	window_holder["window"] = window
	_start_windows_media_key_listener(window_holder, media_key_stop)
	taskbar_toolbar.start_attaching()
	try:
		window.events.shown += lambda: taskbar_toolbar.start_attaching()
	except Exception:
		pass
	window.events.closing += lambda: _save_window_state(data_dir, window)
	window.events.closed += _cleanup
	try:
		window.events.resized += lambda _w, _h: _save_window_state(data_dir, window)
		window.events.moved += lambda _x, _y: _save_window_state(data_dir, window)
	except Exception:
		pass

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
