"""
Yan Menü (Sidebar) MEB Uyumlu Radyo & Yerel Klasör MP3 Medya Sunucusu Modülü.

Özellikler:
1. MEB / Fatih Ağı Uyumluluğu:
   - Yalnızca standart HTTPS (443 portu) üzerinden yayın yapan istasyonları kabul eder;
     şifresiz HTTP ve standart dışı portları (:8000, :8080, :9999 vb.) reddeder.
   - Çevrimdışı öncelikli (Offline-First) yerleşik TRT ve ulusal HTTPS radyo havuzu içerir.
   - İsteğe bağlı Radio Browser dizin sorgusu 2.5 sn zaman aşımı ve sistem/truststore
     SSL bağlamı ile çalışır; erişim engeli varsa beklemeden yerleşik listeye döner.
   - Masaüstü Chromium/WebView2 ortamlarında doğrudan <audio> ile açılamayan resmi
     HTTPS .m3u8 (HLS) akışları için yerel 127.0.0.1 AAC/ADTS akış köprüsü barındırır.
2. Yerel Klasör Seçimi ve HTTP Range Destekli Ses Sunucusu:
   - Seçilen müzik klasörünü (ve alt klasörlerini) .mp3, .wav, .ogg, .m4a, .flac için tarar.
   - 127.0.0.1 üzerinde dinamik boş portta çalışan ThreadingHTTPServer ile RFC 7233
     HTTP Range (206 Partial Content) başlıklarını destekler; büyük dosyaları RAM'e
     yüklemeden 64 KB parçalar halinde sunar ve anında ileri/geri sarmaya izin verir.
"""

from __future__ import annotations

import hashlib
import hmac
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import logging
import mimetypes
import os
from pathlib import Path
import re
import secrets
import ssl
import subprocess
import sys
import threading
import time
from typing import Any
import urllib.error
import urllib.parse
import urllib.request

LOGGER = logging.getLogger(__name__)

SUPPORTED_AUDIO_EXTENSIONS: frozenset[str] = frozenset({
	".mp3",
	".wav",
	".ogg",
	".m4a",
	".flac",
})

AUDIO_MIME_BY_EXT: dict[str, str] = {
	".mp3": "audio/mpeg",
	".wav": "audio/wav",
	".ogg": "audio/ogg",
	".m4a": "audio/mp4",
	".flac": "audio/flac",
}

MAX_SCANNED_TRACKS = 5000
STREAM_CHUNK_BYTES = 65536
RADIO_BROWSER_TR_API = (
	"https://de1.api.radio-browser.info/json/stations/bycountrycodeexact/TR"
	"?hidebroken=true&order=clickcount&reverse=true&limit=60"
)

# MEB-Güvenli Yerleşik İstasyon Havuzu (Tamamı standart HTTPS / 443 portu)
BUILTIN_MEB_SAFE_STATIONS: list[dict[str, Any]] = [
	{
		"id": "trt_fm",
		"name": "TRT FM",
		"category": "TRT / Ulusal",
		"url": "https://radio-trtfm.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtfm.medya.trt.com.tr/master.m3u8",
			"https://trtcanlifm-lh.akamaihd.net/i/TRTFM_1@181846/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_radyo1",
		"name": "TRT Radyo 1",
		"category": "TRT / Söz & Kültür",
		"url": "https://radio-trtradyo1.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtradyo1.medya.trt.com.tr/master.m3u8",
			"https://trtcanlifm-lh.akamaihd.net/i/TRTRADYO1_1@181945/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_radyo3",
		"name": "TRT Radyo 3 (Klasik & Caz)",
		"category": "TRT / Klasik & Caz",
		"url": "https://radio-trtradyo3.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtradyo3.medya.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_nagme",
		"name": "TRT Nağme (Türk Sanat Müziği)",
		"category": "TRT / Sanat Müziği",
		"url": "https://radio-trtnagme.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtnagme.medya.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_turku",
		"name": "TRT Türkü (Türk Halk Müziği)",
		"category": "TRT / Halk Müziği",
		"url": "https://radio-trtturku.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtturku.medya.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_haber",
		"name": "TRT Radyo Haber",
		"category": "TRT / Haber",
		"url": "https://radio-trtradyohaber.live.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtradyohaber.medya.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "ntv_radyo",
		"name": "NTV Radyo",
		"category": "Ulusal / Haber & Müzik",
		"url": "https://ntvradyo.radyotvonline.net/ntvradyo",
		"backup_urls": [],
		"is_hls": False,
	},
	{
		"id": "joyturk",
		"name": "JoyTürk",
		"category": "Ulusal / Türkçe Slow",
		"url": "https://playerservices.streamtheworld.com/api/livestream-redirect/JOY_TURK_SC",
		"backup_urls": [
			"https://playerservices.streamtheworld.com/api/livestream-redirect/JOY_TURKAAC.aac",
		],
		"is_hls": False,
	},
	{
		"id": "joyturk_akustik",
		"name": "JoyTürk Akustik",
		"category": "Ulusal / Akustik",
		"url": "https://playerservices.streamtheworld.com/api/livestream-redirect/JOYTURK_AKUSTIK_SC",
		"backup_urls": [
			"https://playerservices.streamtheworld.com/api/livestream-redirect/JOYTURK_AKUSTIKAAC.aac",
		],
		"is_hls": False,
	},
	{
		"id": "joy_fm",
		"name": "Joy FM (Dinlendirici / Lounge)",
		"category": "Ulusal / Lounge",
		"url": "https://playerservices.streamtheworld.com/api/livestream-redirect/JOY_FM_SC",
		"backup_urls": [
			"https://playerservices.streamtheworld.com/api/livestream-redirect/JOY_FMAAC.aac",
		],
		"is_hls": False,
	},
	{
		"id": "super_fm",
		"name": "Süper FM",
		"category": "Ulusal / Pop",
		"url": "https://playerservices.streamtheworld.com/api/livestream-redirect/SUPER_FM_SC",
		"backup_urls": [
			"https://playerservices.streamtheworld.com/api/livestream-redirect/SUPER_FMAAC.aac",
		],
		"is_hls": False,
	},
	{
		"id": "metro_fm",
		"name": "Metro FM",
		"category": "Ulusal / Yabancı Pop",
		"url": "https://playerservices.streamtheworld.com/api/livestream-redirect/METRO_FM_SC",
		"backup_urls": [
			"https://playerservices.streamtheworld.com/api/livestream-redirect/METRO_FMAAC.aac",
		],
		"is_hls": False,
	},
	{
		"id": "kral_fm",
		"name": "Kral FM",
		"category": "Ulusal / Müzik",
		"url": "https://kralwmp.radyotvonline.com/",
		"backup_urls": [],
		"is_hls": False,
	},
	{
		"id": "kral_pop",
		"name": "Kral Pop Radyo",
		"category": "Ulusal / Türkçe Pop",
		"url": "https://kralpopwmp.radyotvonline.com/",
		"backup_urls": [],
		"is_hls": False,
	},
]


def build_system_ssl_context() -> ssl.SSLContext:
	"""Windows/macOS sistem sertifika deposunu (MEB kök sertifikası dahil) tanıyan SSL bağlamı üretir."""
	try:
		import truststore  # type: ignore[import-not-found]
		return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
	except Exception:
		return ssl.create_default_context()


def is_meb_safe_stream_url(url: str) -> bool:
	"""
	Yalnızca standart https:// şeması ve 443 portu kullanan akış URL'lerini kabul eder.
	Şifresiz http:// adreslerini ve standart dışı portları (:8000, :8080, :9999 vb.) reddeder.
	"""
	raw = str(url or "").strip()
	if not raw or len(raw) > 1024:
		return False
	try:
		parsed = urllib.parse.urlparse(raw)
	except Exception:
		return False
	if (parsed.scheme or "").lower() != "https":
		return False
	host = (parsed.hostname or "").strip().lower()
	if not host or "." not in host:
		return False
	if host in {"localhost", "127.0.0.1", "0.0.0.0", "::1"}:
		return False
	try:
		port = parsed.port
	except ValueError:
		return False
	if port not in (None, 443):
		return False
	return True


def parse_custom_stations_json(raw_json: str) -> list[dict[str, Any]]:
	"""Kullanıcının eklediği özel HTTPS (443) radyo istasyonlarını doğrular ve döndürür."""
	if not raw_json or not str(raw_json).strip():
		return []
	try:
		items = json.loads(str(raw_json))
	except Exception:
		return []
	if not isinstance(items, list):
		return []
	valid: list[dict[str, Any]] = []
	for idx, item in enumerate(items):
		if not isinstance(item, dict):
			continue
		name = str(item.get("name") or "").strip()
		url = str(item.get("url") or "").strip()
		if not name or not is_meb_safe_stream_url(url):
			continue
		sid = str(item.get("id") or f"custom_{idx}_{hashlib.sha1(url.encode('utf-8')).hexdigest()[:8]}")
		valid.append({
			"id": sid,
			"name": name[:80],
			"category": str(item.get("category") or "Özel İstasyon")[:40],
			"url": url,
			"backup_urls": [
				str(u).strip()
				for u in (item.get("backup_urls") or [])
				if is_meb_safe_stream_url(str(u).strip())
			],
			"is_hls": ".m3u8" in url.lower(),
			"custom": True,
		})
	return valid


def fetch_online_meb_safe_stations(timeout: float = 2.5) -> list[dict[str, Any]]:
	"""
	Radio Browser API üzerinden Türkiye istasyonlarını 2.5 sn zaman aşımı ile çeker
	ve yalnızca standart HTTPS (443) uyumlu olanları filtreler.
	MEB ağında API engellenirse hata fırlatmadan boş liste döndürür.
	"""
	req = urllib.request.Request(
		RADIO_BROWSER_TR_API,
		headers={
			"User-Agent": "RevirDefteri-MediaPlayer/1.0",
			"Accept": "application/json",
		},
	)
	ctx = build_system_ssl_context()
	try:
		with urllib.request.urlopen(req, timeout=max(0.5, float(timeout)), context=ctx) as resp:
			raw = resp.read(512 * 1024)
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as exc:
		LOGGER.debug("Radio Browser dizin sorgusu atlandı (çevrimdışı/MEB filtresi): %s", exc)
		return []

	if not isinstance(data, list):
		return []

	results: list[dict[str, Any]] = []
	seen_urls: set[str] = set()
	for row in data:
		if not isinstance(row, dict):
			continue
		stream_url = str(row.get("url_resolved") or row.get("url") or "").strip()
		if not is_meb_safe_stream_url(stream_url):
			continue
		if stream_url in seen_urls:
			continue
		name = str(row.get("name") or "").strip()
		if not name:
			continue
		codec = str(row.get("codec") or "").upper()
		if codec and codec not in {"MP3", "AAC", "AAC+", "OGG", "FLAC", "UNKNOWN", ""}:
			if ".m3u8" not in stream_url.lower():
				continue
		seen_urls.add(stream_url)
		sid = "rb_" + hashlib.sha1(stream_url.encode("utf-8")).hexdigest()[:10]
		results.append({
			"id": sid,
			"name": name[:75],
			"category": "Çevrimiçi (HTTPS)",
			"url": stream_url,
			"backup_urls": [],
			"is_hls": ".m3u8" in stream_url.lower(),
		})
		if len(results) >= 25:
			break
	return results


def get_meb_safe_radio_stations(
	custom_stations_json: str = "",
	include_online: bool = False,
	timeout: float = 2.5,
	hls_proxy_base_url: str = "",
) -> list[dict[str, Any]]:
	"""
	Yerleşik MEB-güvenli istasyonları, kullanıcı özel HTTPS istasyonlarını ve
	(istenirse) çevrimiçi HTTPS istasyonları tek listede birleştirir.
	"""
	merged: list[dict[str, Any]] = []
	seen_urls: set[str] = set()

	custom_list = parse_custom_stations_json(custom_stations_json)
	online_list = fetch_online_meb_safe_stations(timeout=timeout) if include_online else []

	for src in (custom_list, BUILTIN_MEB_SAFE_STATIONS, online_list):
		for st in src:
			url = str(st.get("url") or "").strip()
			if not is_meb_safe_stream_url(url) or url in seen_urls:
				continue
			seen_urls.add(url)
			entry = dict(st)
			backups = [u for u in (entry.get("backup_urls") or []) if is_meb_safe_stream_url(u)]
			if hls_proxy_base_url and (entry.get("is_hls") or ".m3u8" in url.lower()):
				q_url = urllib.parse.quote(url, safe="")
				proxy_u = f"{hls_proxy_base_url.rstrip('/')}/hls_live?url={q_url}"
				entry["hls_proxy_url"] = proxy_u
			entry["backup_urls"] = backups
			merged.append(entry)
	return merged


def resolve_music_folder(custom_folder_setting: str, default_base_dir: Path) -> Path:
	"""
	Ayarlardaki custom_music_folder geçerli bir klasörse onu döndürür;
	boşsa veya bulunamıyorsa default_base_dir / 'muzikler' klasörünü oluşturup döndürür.
	"""
	raw = str(custom_folder_setting or "").strip()
	if raw:
		cand = Path(raw).expanduser()
		try:
			if cand.exists() and cand.is_dir():
				return cand.resolve()
		except Exception:
			pass
	fallback = (Path(default_base_dir) / "muzikler").resolve()
	try:
		fallback.mkdir(parents=True, exist_ok=True)
	except Exception:
		pass
	return fallback


def _clean_track_display_title(file_path: Path) -> str:
	stem = file_path.stem.replace("_", " ").strip()
	stem = re.sub(r"\s+", " ", stem)
	return stem or file_path.name


def _extract_adts_from_mpegts(ts_bytes: bytes) -> bytes:
	"""
	188 baytlık MPEG-TS paketlerinden ses PES akışını (ADTS AAC / MP3 çerçeveleri) ayıklar.
	Tarayıcı <audio> etiketi doğrudan .m3u8 oynatamadığında (örn. Windows Chromium)
	HLS segmentlerini kesintisiz ses akışına dönüştürmek için kullanılır.
	"""
	n = len(ts_bytes)
	if n < 188:
		return ts_bytes
	# İlk senkron baytını (0x47) bul
	sync_off = -1
	for i in range(min(376, n - 188 + 1)):
		if ts_bytes[i] == 0x47 and (i + 188 >= n or ts_bytes[i + 188] == 0x47):
			sync_off = i
			break
	if sync_off < 0:
		# Zaten ham AAC/ADTS veya MP3 segmenti
		return ts_bytes

	pmt_pids: set[int] = set()
	audio_pids: set[int] = set()
	pes_by_pid: dict[int, bytearray] = {}

	pos = sync_off
	while pos + 188 <= n:
		pkt = ts_bytes[pos : pos + 188]
		pos += 188
		if pkt[0] != 0x47:
			continue
		pusi = bool(pkt[1] & 0x40)
		pid = ((pkt[1] & 0x1F) << 8) | pkt[2]
		afc = (pkt[3] >> 4) & 0x03
		if afc in (0, 2):
			continue
		idx = 4
		if afc == 3:
			adapt_len = pkt[4]
			idx = 5 + adapt_len
		if idx >= 188:
			continue
		payload = pkt[idx:]

		if pid == 0:  # PAT
			if pusi and payload:
				ptr = payload[0]
				sec = payload[1 + ptr :]
			else:
				sec = payload
			if len(sec) >= 12 and sec[0] == 0x00:
				sec_len = ((sec[1] & 0x0F) << 8) | sec[2]
				prog_data = sec[8 : max(8, min(len(sec), 3 + sec_len - 4))]
				for j in range(0, len(prog_data) - 3, 4):
					prog_num = (prog_data[j] << 8) | prog_data[j + 1]
					p_pid = ((prog_data[j + 2] & 0x1F) << 8) | prog_data[j + 3]
					if prog_num != 0:
						pmt_pids.add(p_pid)
			continue

		if pid in pmt_pids:  # PMT
			if pusi and payload:
				ptr = payload[0]
				sec = payload[1 + ptr :]
			else:
				sec = payload
			if len(sec) >= 16 and sec[0] == 0x02:
				sec_len = ((sec[1] & 0x0F) << 8) | sec[2]
				prog_info_len = ((sec[10] & 0x0F) << 8) | sec[11]
				es_pos = 12 + prog_info_len
				es_end = min(len(sec), 3 + sec_len - 4)
				while es_pos + 5 <= es_end:
					stype = sec[es_pos]
					e_pid = ((sec[es_pos + 1] & 0x1F) << 8) | sec[es_pos + 2]
					es_info_len = ((sec[es_pos + 3] & 0x0F) << 8) | sec[es_pos + 4]
					# 0x0F: ADTS AAC, 0x03/0x04: MPEG Audio, 0x11: LATM AAC
					if stype in (0x0F, 0x03, 0x04, 0x11):
						audio_pids.add(e_pid)
					es_pos += 5 + es_info_len
			continue

		if pid in (0x1FFF, 0x0001, 0x0011):
			continue
		if audio_pids and pid not in audio_pids:
			continue

		if pusi and len(payload) >= 9 and payload[0] == 0x00 and payload[1] == 0x00 and payload[2] == 0x01:
			stream_id = payload[3]
			# Yalnızca ses akışları (0xC0..0xDF) veya özel ses akışı
			if not (0xC0 <= stream_id <= 0xDF):
				continue
			pes_hdr_len = payload[8]
			es_start = 9 + pes_hdr_len
			if es_start < len(payload):
				pes_by_pid.setdefault(pid, bytearray()).extend(payload[es_start:])
		elif pid in pes_by_pid:
			pes_by_pid[pid].extend(payload)

	if not pes_by_pid:
		return b""
	# En büyük ses PES akışını döndür
	best_pid = max(pes_by_pid, key=lambda k: len(pes_by_pid[k]))
	return bytes(pes_by_pid[best_pid])


def parse_byte_range_header(range_header: str | None, file_size: int) -> tuple[int, int] | None:
	"""
	RFC 7233 HTTP Range başlığını ('bytes=start-end', 'bytes=start-', 'bytes=-suffix') çözümler.
	Range başlığı yoksa None döndürür; geçersizse ValueError fırlatır.
	"""
	if not range_header:
		return None
	raw = range_header.strip()
	if not raw.lower().startswith("bytes="):
		return None
	spec = raw[6:].split(",", 1)[0].strip()
	if "-" not in spec:
		raise ValueError("Invalid Range format")
	start_s, end_s = spec.split("-", 1)
	start_s = start_s.strip()
	end_s = end_s.strip()

	if file_size <= 0:
		raise ValueError("Empty file")

	if start_s == "":
		# Son N bayt: bytes=-500
		if not end_s.isdigit():
			raise ValueError("Invalid suffix range")
		suffix = int(end_s)
		if suffix <= 0:
			raise ValueError("Zero suffix range")
		start = max(0, file_size - suffix)
		end = file_size - 1
		return (start, end)

	if not start_s.isdigit():
		raise ValueError("Invalid range start")
	start = int(start_s)
	if start >= file_size:
		raise ValueError("Range start out of bounds")

	if end_s == "":
		end = file_size - 1
	else:
		if not end_s.isdigit():
			raise ValueError("Invalid range end")
		end = min(file_size - 1, int(end_s))

	if end < start:
		raise ValueError("Range end before start")
	return (start, end)


class LocalMediaStreamServer:
	"""
	127.0.0.1 üzerinde dinamik portta çalışan, HTTP Range (206 Partial Content) destekli
	yerel ses akış sunucusu ve HTTPS (443) HLS ses köprüsü.
	"""

	def __init__(self, host: str = "127.0.0.1", port: int = 0) -> None:
		self.host = host
		self._lock = threading.RLock()
		self._secret = secrets.token_bytes(24)
		self._root_folder: Path | None = None
		self._tracks_by_id: dict[str, Path] = {}
		self._cached_tracks: list[dict[str, Any]] = []
		self._cached_folder_key: str = ""
		self._httpd = ThreadingHTTPServer((host, int(port)), self._build_handler())
		self._httpd.daemon_threads = True
		self.port = int(self._httpd.server_address[1])
		self.base_url = f"http://{self.host}:{self.port}"
		self._thread: threading.Thread | None = None

	def start(self) -> str:
		with self._lock:
			if self._thread is None or not self._thread.is_alive():
				self._thread = threading.Thread(
					target=self._httpd.serve_forever,
					kwargs={"poll_interval": 0.25},
					daemon=True,
					name="RevirLocalMediaStreamServer",
				)
				self._thread.start()
		return self.base_url

	def stop(self) -> None:
		try:
			self._httpd.shutdown()
			self._httpd.server_close()
		except Exception:
			pass

	def _make_track_id(self, resolved_file: Path) -> str:
		digest = hmac.new(self._secret, str(resolved_file).encode("utf-8", errors="ignore"), hashlib.sha256).hexdigest()
		return digest[:20]

	def scan_folder(self, folder: Path, force: bool = False) -> list[dict[str, Any]]:
		"""Seçilen klasörü ve alt klasörlerini ses dosyaları için tarar ve çalma listesini döndürür."""
		self.start()
		try:
			resolved_root = Path(folder).expanduser().resolve()
		except Exception:
			return []
		if not resolved_root.exists() or not resolved_root.is_dir():
			with self._lock:
				self._root_folder = resolved_root
				self._tracks_by_id.clear()
				self._cached_tracks = []
				self._cached_folder_key = ""
			return []

		try:
			root_mtime = resolved_root.stat().st_mtime_ns
		except Exception:
			root_mtime = 0
		cache_key = f"{resolved_root}::{root_mtime}"

		with self._lock:
			if not force and cache_key == self._cached_folder_key and self._root_folder == resolved_root:
				return list(self._cached_tracks)

		discovered: list[tuple[str, Path, int]] = []
		try:
			for dirpath, dirnames, filenames in os.walk(resolved_root):
				dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
				for fname in sorted(filenames):
					if fname.startswith("."):
						continue
					ext = os.path.splitext(fname)[1].lower()
					if ext not in SUPPORTED_AUDIO_EXTENSIONS:
						continue
					fpath = Path(dirpath) / fname
					try:
						resolved_f = fpath.resolve()
						if not resolved_f.is_relative_to(resolved_root):
							continue
						st = resolved_f.stat()
						if st.st_size <= 0:
							continue
						rel_str = str(resolved_f.relative_to(resolved_root)).replace("\\", "/")
						discovered.append((rel_str, resolved_f, int(st.st_size)))
					except Exception:
						continue
					if len(discovered) >= MAX_SCANNED_TRACKS:
						break
				if len(discovered) >= MAX_SCANNED_TRACKS:
					break
		except Exception as exc:
			LOGGER.warning("Müzik klasörü taranırken hata oluştu (%s): %s", resolved_root, exc)

		discovered.sort(key=lambda item: item[0].lower())
		new_map: dict[str, Path] = {}
		playlist: list[dict[str, Any]] = []
		for rel_str, resolved_f, size_bytes in discovered:
			tid = self._make_track_id(resolved_f)
			new_map[tid] = resolved_f
			sub_dir = str(Path(rel_str).parent).replace("\\", "/")
			folder_hint = "" if sub_dir in ("", ".") else sub_dir
			playlist.append({
				"id": tid,
				"title": _clean_track_display_title(resolved_f),
				"folder": folder_hint,
				"rel_path": rel_str,
				"ext": resolved_f.suffix.lower().lstrip("."),
				"size_bytes": size_bytes,
				"url": f"{self.base_url}/stream/{tid}",
			})

		with self._lock:
			self._root_folder = resolved_root
			self._tracks_by_id = new_map
			self._cached_tracks = playlist
			self._cached_folder_key = cache_key
		return list(playlist)

	def resolve_track_path(self, track_id: str) -> Path | None:
		clean_id = re.sub(r"[^a-f0-9]", "", str(track_id or "").lower())
		if not clean_id:
			return None
		with self._lock:
			root = self._root_folder
			fpath = self._tracks_by_id.get(clean_id)
		if root is None or fpath is None:
			return None
		try:
			resolved = fpath.resolve()
			if not resolved.is_relative_to(root) or not resolved.is_file():
				return None
			return resolved
		except Exception:
			return None

	def _build_handler(self):
		server_ref = self

		class _AudioRangeHandler(BaseHTTPRequestHandler):
			protocol_version = "HTTP/1.1"

			def log_message(self, format: str, *args) -> None:
				return  # Sessiz arka plan ses sunucusu

			def do_OPTIONS(self) -> None:
				self.send_response(HTTPStatus.NO_CONTENT)
				self.send_header("Access-Control-Allow-Origin", "*")
				self.send_header("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS")
				self.send_header("Access-Control-Allow-Headers", "Range, Content-Type")
				self.send_header("Content-Length", "0")
				self.end_headers()

			def do_HEAD(self) -> None:
				self._handle_request(send_body=False)

			def do_GET(self) -> None:
				self._handle_request(send_body=True)

			def _send_simple(self, status: int, text_msg: str, extra: dict[str, str] | None = None) -> None:
				raw = text_msg.encode("utf-8")
				self.send_response(status)
				self.send_header("Content-Type", "text/plain; charset=utf-8")
				self.send_header("Content-Length", str(len(raw)))
				self.send_header("Access-Control-Allow-Origin", "*")
				if extra:
					for k, v in extra.items():
						self.send_header(k, v)
				self.end_headers()
				try:
					self.wfile.write(raw)
				except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
					pass

			def _handle_request(self, send_body: bool) -> None:
				parsed = urllib.parse.urlparse(self.path)
				path = parsed.path

				if path == "/health":
					self._send_simple(HTTPStatus.OK, "ok")
					return

				if path == "/hls_live":
					qs = urllib.parse.parse_qs(parsed.query)
					m3u8_url = str((qs.get("url") or [""])[0]).strip()
					self._stream_hls_as_audio(m3u8_url, send_body=send_body)
					return

				if path.startswith("/stream/"):
					track_id = urllib.parse.unquote(path[len("/stream/") :]).split("/", 1)[0]
					fpath = server_ref.resolve_track_path(track_id)
					if fpath is None:
						self._send_simple(HTTPStatus.NOT_FOUND, "Track Not Found")
						return
					self._serve_audio_file(fpath, send_body=send_body)
					return

				self._send_simple(HTTPStatus.NOT_FOUND, "Not Found")

			def _serve_audio_file(self, fpath: Path, send_body: bool) -> None:
				try:
					st = fpath.stat()
					file_size = int(st.st_size)
				except Exception:
					self._send_simple(HTTPStatus.NOT_FOUND, "File Unavailable")
					return

				ext = fpath.suffix.lower()
				mime = AUDIO_MIME_BY_EXT.get(ext) or mimetypes.guess_type(fpath.name)[0] or "application/octet-stream"
				range_hdr = self.headers.get("Range")

				try:
					byte_range = parse_byte_range_header(range_hdr, file_size)
				except ValueError:
					self._send_simple(
						HTTPStatus.REQUESTED_RANGE_NOT_SATISFIABLE,
						"Requested Range Not Satisfiable",
						{"Content-Range": f"bytes */{file_size}", "Accept-Ranges": "bytes"},
					)
					return

				if byte_range is not None:
					start, end = byte_range
					length = end - start + 1
					self.send_response(HTTPStatus.PARTIAL_CONTENT)
					self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
				else:
					start, end = 0, max(0, file_size - 1)
					length = file_size
					self.send_response(HTTPStatus.OK)

				self.send_header("Content-Type", mime)
				self.send_header("Accept-Ranges", "bytes")
				self.send_header("Content-Length", str(length))
				self.send_header("Cache-Control", "no-cache")
				self.send_header("Access-Control-Allow-Origin", "*")
				self.send_header("Access-Control-Expose-Headers", "Content-Range, Accept-Ranges, Content-Length")
				self.end_headers()

				if not send_body or length <= 0:
					return

				try:
					with fpath.open("rb") as fh:
						if start > 0:
							fh.seek(start)
						remaining = length
						while remaining > 0:
							chunk = fh.read(min(STREAM_CHUNK_BYTES, remaining))
							if not chunk:
								break
							self.wfile.write(chunk)
							remaining -= len(chunk)
				except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
					return

			def _stream_hls_as_audio(self, m3u8_url: str, send_body: bool) -> None:
				"""
				Standart HTTPS (443) .m3u8 canlı radyo yayınını sistem SSL deposu ile okuyup
				tarayıcı <audio> etiketine kesintisiz ADTS AAC / MP3 akışı olarak aktarır.
				"""
				if not is_meb_safe_stream_url(m3u8_url):
					self._send_simple(HTTPStatus.FORBIDDEN, "Only standard HTTPS (443) streams are allowed")
					return

				ssl_ctx = build_system_ssl_context()
				headers = {"User-Agent": "Mozilla/5.0 RevirDefteri/1.0"}

				def _fetch_bytes(target_url: str, timeout_s: float = 4.0) -> tuple[bytes, str]:
					if not is_meb_safe_stream_url(target_url):
						raise ValueError("Unsafe redirect/segment URL")
					req = urllib.request.Request(target_url, headers=headers)
					with urllib.request.urlopen(req, timeout=timeout_s, context=ssl_ctx) as r:
						final_u = str(r.geturl() or target_url)
						if not is_meb_safe_stream_url(final_u):
							raise ValueError("Redirected to non-HTTPS/non-443 URL")
						return r.read(2 * 1024 * 1024), final_u

				try:
					raw_pl, base_u = _fetch_bytes(m3u8_url, timeout_s=3.5)
				except Exception as exc:
					self._send_simple(HTTPStatus.BAD_GATEWAY, f"HLS baglantisi kurulamadi: {exc}")
					return

				self.send_response(HTTPStatus.OK)
				self.send_header("Content-Type", "audio/aac")
				self.send_header("Cache-Control", "no-store")
				self.send_header("Access-Control-Allow-Origin", "*")
				self.send_header("Connection", "close")
				self.end_headers()

				if not send_body:
					return

				# Master playlist ise alt medya .m3u8 adresini seç
				text_pl = raw_pl.decode("utf-8", errors="replace")
				media_url = base_u
				if "#EXT-X-STREAM-INF" in text_pl:
					for line in text_pl.splitlines():
						s = line.strip()
						if s and not s.startswith("#"):
							cand_u = urllib.parse.urljoin(base_u, s)
							if is_meb_safe_stream_url(cand_u):
								media_url = cand_u
								break

				seen_segs: list[str] = []
				max_duration_sec = 4 * 3600  # Tek oturumda 4 saat üst sınırı
				started = time.monotonic()

				while (time.monotonic() - started) < max_duration_sec:
					try:
						pl_bytes, cur_base = _fetch_bytes(media_url, timeout_s=3.5)
						pl_txt = pl_bytes.decode("utf-8", errors="replace")
					except Exception:
						break

					target_dur = 2.0
					seg_urls: list[str] = []
					for line in pl_txt.splitlines():
						s = line.strip()
						if not s:
							continue
						if s.startswith("#EXT-X-TARGETDURATION:"):
							try:
								target_dur = max(1.0, min(8.0, float(s.split(":", 1)[1].strip())))
							except Exception:
								pass
						elif not s.startswith("#"):
							su = urllib.parse.urljoin(cur_base, s)
							if is_meb_safe_stream_url(su):
								seg_urls.append(su)

					if not seen_segs and len(seg_urls) > 2:
						# Canlı yayında düşük gecikme için son 2 segmentten başla
						for old_u in seg_urls[:-2]:
							seen_segs.append(old_u)
						seg_urls = seg_urls[-2:]

					sent_any = False
					for su in seg_urls:
						if su in seen_segs:
							continue
						seen_segs.append(su)
						if len(seen_segs) > 80:
							seen_segs = seen_segs[-50:]
						try:
							seg_raw, _ = _fetch_bytes(su, timeout_s=4.5)
							audio_bytes = _extract_adts_from_mpegts(seg_raw)
							if audio_bytes:
								self.wfile.write(audio_bytes)
								self.wfile.flush()
								sent_any = True
						except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
							return
						except Exception:
							continue

					time.sleep(0.8 if sent_any else min(2.0, target_dur * 0.5))

		return _AudioRangeHandler


_GLOBAL_MEDIA_SERVER_LOCK = threading.Lock()
_GLOBAL_MEDIA_SERVER: LocalMediaStreamServer | None = None


def get_local_media_server() -> LocalMediaStreamServer:
	"""Uygulama genelinde paylaşılan tekil LocalMediaStreamServer örneğini döndürür."""
	global _GLOBAL_MEDIA_SERVER
	with _GLOBAL_MEDIA_SERVER_LOCK:
		if _GLOBAL_MEDIA_SERVER is None:
			_GLOBAL_MEDIA_SERVER = LocalMediaStreamServer(host="127.0.0.1", port=0)
		_GLOBAL_MEDIA_SERVER.start()
		return _GLOBAL_MEDIA_SERVER


def pick_music_folder_native(initial_dir: str = "") -> str | None:
	"""
	Masaüstü penceresi (pywebview) dışında çalışırken veya yedek olarak yerel klasör seçici açar.
	Önce tkinter.filedialog dener; PyInstaller ile tkinter dışlanmışsa Windows'ta yerleşik
	System.Windows.Forms.FolderBrowserDialog penceresini kullanır.
	"""
	start_dir = str(initial_dir or "").strip()
	if not start_dir or not os.path.isdir(start_dir):
		start_dir = str(Path.home())

	try:
		import tkinter as tk
		from tkinter import filedialog
		root = tk.Tk()
		root.withdraw()
		try:
			root.attributes("-topmost", True)
		except Exception:
			pass
		chosen = filedialog.askdirectory(initialdir=start_dir, title="Müzik Klasörü Seçin")
		root.destroy()
		if chosen:
			return str(Path(chosen).resolve())
		return None
	except Exception:
		pass

	if sys.platform == "win32":
		ps_script = (
			"Add-Type -AssemblyName System.Windows.Forms; "
			"$dlg = New-Object System.Windows.Forms.FolderBrowserDialog; "
			"$dlg.Description = 'Müzik / MP3 Klasörü Seçin'; "
			f"$dlg.SelectedPath = '{start_dir.replace(chr(39), chr(39)*2)}'; "
			"if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { "
			"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
			"Write-Output $dlg.SelectedPath }"
		)
		try:
			creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
			res = subprocess.run(
				["powershell", "-NoProfile", "-STA", "-Command", ps_script],
				capture_output=True,
				text=True,
				encoding="utf-8",
				errors="replace",
				timeout=120,
				creationflags=creationflags,
			)
			out = (res.stdout or "").strip()
			if out and os.path.isdir(out):
				return str(Path(out).resolve())
		except Exception as exc:
			LOGGER.warning("Yerel klasör seçici açılamadı: %s", exc)
	return None


def build_sidebar_media_player_payload(
	app_cfg: dict[str, str],
	default_base_dir: Path,
	*,
	include_online_stations: bool = False,
	force_rescan: bool = False,
) -> dict[str, Any]:
	"""
	revir_ui.py yan menü oynatıcısına gönderilecek yapılandırma, MEB-güvenli radyo istasyonları
	ve yerel MP3 çalma listesini hazırlar.
	"""
	visible = str(app_cfg.get("show_sidebar_media_player", "1")).strip() != "0"
	srv = get_local_media_server()
	music_dir = resolve_music_folder(app_cfg.get("custom_music_folder", ""), default_base_dir)
	tracks = srv.scan_folder(music_dir, force=force_rescan) if visible else []
	stations = (
		get_meb_safe_radio_stations(
			custom_stations_json=app_cfg.get("custom_radio_stations_json", "[]"),
			include_online=include_online_stations,
			timeout=2.5,
			hls_proxy_base_url=srv.base_url,
		)
		if visible
		else []
	)
	return {
		"visible": visible,
		"stream_base_url": srv.base_url,
		"music_folder": str(music_dir),
		"music_folder_name": music_dir.name or str(music_dir),
		"is_custom_folder": bool(str(app_cfg.get("custom_music_folder", "")).strip()),
		"tracks": tracks,
		"track_count": len(tracks),
		"stations": stations,
		"station_count": len(stations),
	}
