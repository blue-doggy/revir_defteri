"""
Yan Menü (Sidebar) MEB Uyumlu Radyo, Çevrimdışı Ortam Sesleri & Kapaklı Yerel MP3 Medya Sunucusu Modülü.

Özellikler:
1. MEB / Fatih Ağı Uyumluluğu ve Akıllı Yerel Ses Köprüsü:
   - Yalnızca standart HTTPS (443 portu) üzerinden yayın yapan istasyonları kabul eder;
     şifresiz HTTP ve standart dışı portları (:8000, :8080, :9999 vb.) reddeder.
   - Hem .m3u8 (HLS) hem doğrudan HTTPS (.mp3/.aac) akışları (tüm yedek adresler dahil)
     127.0.0.1 yerel Python ses köprüsü (/hls_live ve /radio_proxy) üzerinden geçirir.
   - Alt oynatma listesi ve ilk ses çerçevesini doğrulamadan HTTP 200 göndermez; MEB/FortiGuard
     HTML engel sayfalarını anında algılayıp yedek adrese geçer, TRT .aac ID3 bloklarını temizler.
   - Sistem sertifika deposu (truststore) + özel MEB sertifikası (.cer/.crt/.pem, DER->PEM otomatik)
     ve tek tıkla MEB Radyo Ağ & Sertifika Tanılama (diagnose_radio_connectivity) desteği içerir.
2. Çevrimdışı Dinlendirici Revir Ortam Sesleri (/ambient/<id>):
   - Dış radyo akışlarının tamamen kapalı olduğu ağlarda sıfır internetle çalışan yerleşik WAV üreteci.
3. Albüm Kapaklı & HTTP Range Destekli Yerel Klasör Müzik Sunucusu (/stream/<id>, /cover/<id>):
   - MP3 (ID3v2 APIC + TIT2/TPE1), M4A (covr) ve FLAC (PICTURE) gömülü albüm kapaklarını ve
     klasör içi kapak görsellerini saf Python ile okur; yoksa estetik SVG kapak üretir.
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
import hmac
import html
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import logging
import math
import mimetypes
import os
from pathlib import Path
import random
import re
import secrets
import ssl
import struct
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

COVER_FILE_CANDIDATES: tuple[str, ...] = (
	"cover.jpg",
	"cover.jpeg",
	"cover.png",
	"folder.jpg",
	"folder.jpeg",
	"folder.png",
	"album.jpg",
	"album.png",
	"front.jpg",
	"front.png",
)

MAX_SCANNED_TRACKS = 5000
STREAM_CHUNK_BYTES = 65536
RADIO_BROWSER_TR_API = (
	"https://de1.api.radio-browser.info/json/stations/bycountrycodeexact/TR"
	"?hidebroken=true&order=clickcount&reverse=true&limit=60"
)

_SSL_OPTIONS_LOCK = threading.RLock()
_CUSTOM_MEB_CERT_PATH: str = ""
_ALLOW_RELAXED_SSL: bool = False

# Canlı doğrulanmış MEB-Güvenli Yerleşik İstasyon Havuzu (Tamamı standart HTTPS / 443 portu ve aktif akış)
BUILTIN_MEB_SAFE_STATIONS: list[dict[str, Any]] = [
	{
		"id": "trt_fm",
		"name": "TRT FM",
		"category": "TRT / Ulusal",
		"url": "https://rd-trtfm.medya.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtfm.medya.trt.com.tr/master_128.m3u8",
			"https://radio-trtfm.live.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_radyo1",
		"name": "TRT Radyo 1",
		"category": "TRT / Söz & Kültür",
		"url": "https://rd-trtradyo1.medya.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtradyo1.medya.trt.com.tr/master_128.m3u8",
			"https://radio-trtradyo1.live.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_radyo3",
		"name": "TRT Radyo 3 (Klasik & Caz)",
		"category": "TRT / Klasik & Caz",
		"url": "https://rd-trtradyo3.medya.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtradyo3.medya.trt.com.tr/master_128.m3u8",
			"https://radio-trtradyo3.live.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_nagme",
		"name": "TRT Nağme (Türk Sanat Müziği)",
		"category": "TRT / Sanat Müziği",
		"url": "https://rd-trtnagme.medya.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtnagme.medya.trt.com.tr/master_128.m3u8",
			"https://radio-trtnagme.live.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "trt_turku",
		"name": "TRT Türkü (Türk Halk Müziği)",
		"category": "TRT / Halk Müziği",
		"url": "https://rd-trtturku.medya.trt.com.tr/master.m3u8",
		"backup_urls": [
			"https://rd-trtturku.medya.trt.com.tr/master_128.m3u8",
			"https://radio-trtturku.live.trt.com.tr/master.m3u8",
		],
		"is_hls": True,
	},
	{
		"id": "diyanet_radyo",
		"name": "Diyanet Radyo (Huzur & Söz)",
		"category": "Kamu / Kültür",
		"url": "https://eustr73.mediatriple.net/videoonlylive/mtikoimxnztxlive/broadcast_5e3bf95a47e07.smil/playlist.m3u8",
		"backup_urls": [],
		"is_hls": True,
	},
	{
		"id": "radyo_turkuvaz",
		"name": "Radyo Turkuvaz",
		"category": "Ulusal / Türkçe Müzik",
		"url": "https://trkvz-radyolar.ercdn.net/radyoturkuvaz/playlist.m3u8",
		"backup_urls": [],
		"is_hls": True,
	},
	{
		"id": "lofi_code_radio",
		"name": "Code Radio (Lo-Fi & Odaklanma)",
		"category": "Lo-Fi / Odaklanma",
		"url": "https://coderadio-admin-v2.freecodecamp.org/listen/coderadio/radio.mp3",
		"backup_urls": [
			"https://coderadio-admin-v2.freecodecamp.org/listen/coderadio/low.mp3",
		],
		"is_hls": False,
	},
	{
		"id": "lofi_box_radio",
		"name": "Box Radio (Lo-Fi Chill Beats)",
		"category": "Lo-Fi / Chillhop",
		"url": "https://play.streamafrica.net/lofiradio",
		"backup_urls": [
			"https://radio.plaza.one/mp3",
			"https://radio.plaza.one/mp3_low",
		],
		"is_hls": False,
	},
	{
		"id": "ambient_sleeping_pill",
		"name": "Ambient Sleeping Pill (Saf Ambiyans)",
		"category": "Ambiyans / Dinlenme",
		"url": "https://radio.stereoscenic.com/asp-h",
		"backup_urls": [
			"https://radio.stereoscenic.com/asp-s",
			"https://radio.stereoscenic.com/ama-h",
			"https://radio.stereoscenic.com/ama-s",
		],
		"is_hls": False,
	},
]

# MEB ağında tüm canlı akışlar engellense dahi %100 çevrimdışı çalışan revir ortam sesleri
BUILTIN_AMBIENT_PRESETS: list[dict[str, str]] = [
	{
		"id": "amb_rain",
		"name": "Hafif Yağmur & Dinlenme",
		"category": "Çevrimdışı Revir Ambiyansı",
		"icon": "🌧️",
	},
	{
		"id": "amb_waves",
		"name": "Sakin Deniz Dalgaları",
		"category": "Çevrimdışı Revir Ambiyansı",
		"icon": "🌊",
	},
	{
		"id": "amb_breeze",
		"name": "Yumuşak Doğa Esintisi",
		"category": "Çevrimdışı Revir Ambiyansı",
		"icon": "🍃",
	},
	{
		"id": "amb_calm432",
		"name": "Huzurlu Akustik Ton (432 Hz)",
		"category": "Çevrimdışı Revir Ambiyansı",
		"icon": "🎹",
	},
]


def configure_ssl_options(*, custom_cert_path: str = "", allow_relaxed_ssl: bool = False) -> None:
	global _CUSTOM_MEB_CERT_PATH, _ALLOW_RELAXED_SSL
	with _SSL_OPTIONS_LOCK:
		_CUSTOM_MEB_CERT_PATH = str(custom_cert_path or "").strip()
		_ALLOW_RELAXED_SSL = bool(allow_relaxed_ssl)


def _load_cert_file_as_pem(cert_path_str: str) -> str | None:
	if not cert_path_str:
		return None
	p = Path(cert_path_str).expanduser()
	if not p.exists() or not p.is_file():
		return None
	raw = p.read_bytes()
	if not raw:
		return None
	if b"-----BEGIN CERTIFICATE-----" in raw:
		return raw.decode("utf-8", errors="replace")
	return ssl.DER_cert_to_PEM_cert(raw)


def build_system_ssl_context(
	*,
	custom_cert_path: str | None = None,
	allow_relaxed_ssl: bool | None = None,
) -> ssl.SSLContext:
	"""Windows/macOS sistem sertifika deposunu ve varsa özel MEB sertifikasını (.cer/.pem) tanıyan SSL bağlamı üretir."""
	with _SSL_OPTIONS_LOCK:
		c_path = _CUSTOM_MEB_CERT_PATH if custom_cert_path is None else str(custom_cert_path or "").strip()
		relaxed = _ALLOW_RELAXED_SSL if allow_relaxed_ssl is None else bool(allow_relaxed_ssl)

	if relaxed:
		ctx_r = ssl.create_default_context()
		ctx_r.check_hostname = False
		ctx_r.verify_mode = ssl.CERT_NONE
		return ctx_r

	pem_data = None
	try:
		pem_data = _load_cert_file_as_pem(c_path)
	except Exception as exc:
		LOGGER.warning("Özel MEB sertifikası okunamadı (%s): %s", c_path, exc)

	if pem_data:
		ctx_custom = ssl.create_default_context()
		try:
			ctx_custom.load_verify_locations(cadata=pem_data)
		except Exception as exc:
			LOGGER.warning("Özel MEB sertifikası SSL bağlamına yüklenemedi: %s", exc)
		return ctx_custom

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


def is_meb_block_or_html_payload(content_type: str, sample_bytes: bytes) -> bool:
	"""MEB / FortiGuard güvenlik duvarının ses/m3u8 yerine döndürdüğü HTML engel sayfasını tespit eder."""
	ct_low = str(content_type or "").lower()
	if "text/html" in ct_low or "application/xhtml" in ct_low:
		return True
	head = (sample_bytes or b"")[:600].lstrip().lower()
	if head.startswith((b"<!doctype html", b"<html", b"<head", b"<script", b"<!--")):
		return True
	for marker in (b"fortiguard", b"meb.gov.tr", b"fatihprojesi", b"erisim engellen", b"access blocked"):
		if marker in head:
			return True
	return False


def _strip_leading_id3(data: bytes) -> bytes:
	"""AAC/MP3 segment başındaki ID3v2 metadata etiketlerini ayıklayıp saf ses çerçevesini bırakır."""
	buf = data
	while len(buf) >= 10 and buf[:3] == b"ID3":
		flags = buf[5]
		tag_size = (
			((buf[6] & 0x7F) << 21)
			| ((buf[7] & 0x7F) << 14)
			| ((buf[8] & 0x7F) << 7)
			| (buf[9] & 0x7F)
		)
		skip = 10 + tag_size + (10 if (flags & 0x10) else 0)
		if skip <= 0 or skip >= len(buf):
			break
		buf = buf[skip:]
	return buf


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
			ct = resp.headers.get("Content-Type", "")
			raw = resp.read(512 * 1024)
			if is_meb_block_or_html_payload(ct, raw):
				return []
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
		if len(results) >= 20:
			break

	if not results:
		return []

	# Yalnızca mevcut ağda gerçek ses/HLS akışı veren (503/404/MEB engeli olmayan) kanalları tut
	verified: list[dict[str, Any]] = []
	with ThreadPoolExecutor(max_workers=min(6, len(results))) as pool:
		fut_to_idx = {pool.submit(diagnose_single_station, st, min(2.2, max(0.8, float(timeout)))): i for i, st in enumerate(results)}
		ok_indices: list[int] = []
		for fut in as_completed(fut_to_idx):
			idx = fut_to_idx[fut]
			try:
				if fut.result().get("ok"):
					ok_indices.append(idx)
			except Exception:
				pass
	for idx in sorted(ok_indices):
		verified.append(results[idx])
	return verified


def _wrap_stream_via_local_proxy(stream_url: str, proxy_base_url: str, force_hls: bool = False) -> str:
	base = proxy_base_url.rstrip("/")
	q_url = urllib.parse.quote(stream_url, safe="")
	if force_hls or ".m3u8" in stream_url.lower():
		return f"{base}/hls_live?url={q_url}"
	return f"{base}/radio_proxy?url={q_url}"


def get_meb_safe_radio_stations(
	custom_stations_json: str = "",
	include_online: bool = False,
	timeout: float = 2.5,
	hls_proxy_base_url: str = "",
) -> list[dict[str, Any]]:
	"""
	Yerleşik MEB-güvenli istasyonları, kullanıcı özel HTTPS istasyonlarını ve
	(istenirse) çevrimiçi HTTPS istasyonları yerel ses köprüsü zinciriyle birleştirir.
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
			is_hls = bool(entry.get("is_hls") or ".m3u8" in url.lower())
			backups = [u for u in (entry.get("backup_urls") or []) if is_meb_safe_stream_url(u)]
			entry["backup_urls"] = backups
			entry["is_hls"] = is_hls

			if hls_proxy_base_url:
				if is_hls:
					entry["hls_proxy_url"] = _wrap_stream_via_local_proxy(url, hls_proxy_base_url, force_hls=True)
				proxy_chain: list[str] = []
				for cand in [url, *backups]:
					cand_hls = is_hls or (".m3u8" in cand.lower())
					pu = _wrap_stream_via_local_proxy(cand, hls_proxy_base_url, force_hls=cand_hls)
					if pu not in proxy_chain:
						proxy_chain.append(pu)
				entry["proxy_chain"] = proxy_chain
			merged.append(entry)
	return merged


def diagnose_single_station(station: dict[str, Any], timeout: float = 3.5) -> dict[str, Any]:
	"""Tek bir radyo istasyonunu MEB engeli, SSL sertifikası ve canlı ses akışı açısından test eder."""
	ctx = build_system_ssl_context()
	candidates = [str(station.get("url") or "").strip(), *[str(u).strip() for u in (station.get("backup_urls") or [])]]
	candidates = [c for c in candidates if is_meb_safe_stream_url(c)]
	last_status = "dns_or_timeout"
	last_detail = "Geçerli HTTPS adresi yok"

	for cand in candidates:
		try:
			req = urllib.request.Request(
				cand,
				headers={"User-Agent": "Mozilla/5.0 RevirDefteri/1.0", "Range": "bytes=0-2048", "Icy-MetaData": "0"},
			)
			with urllib.request.urlopen(req, timeout=max(1.0, float(timeout)), context=ctx) as resp:
				final_u = str(getattr(resp, "geturl", lambda: cand)() or cand)
				if not is_meb_safe_stream_url(final_u):
					last_status = "meb_blocked"
					last_detail = f"Güvensiz porta/HTTP'ye yönlendirildi ({final_u[:60]})"
					continue
				ct = resp.headers.get("Content-Type", "")
				sample = resp.read(2048)
				if is_meb_block_or_html_payload(ct, sample):
					last_status = "meb_blocked"
					last_detail = "MEB / Güvenlik Duvarı HTML engel sayfası döndürdü"
					continue

				# Master .m3u8 ise alt listeyi de doğrula
				if b"#EXTM3U" in sample[:256] and b"#EXT-X-STREAM-INF" in sample:
					txt = sample.decode("utf-8", errors="replace")
					sub_u = next(
						(urllib.parse.urljoin(final_u, ln.strip()) for ln in txt.splitlines() if ln.strip() and not ln.strip().startswith("#")),
						None,
					)
					if sub_u and is_meb_safe_stream_url(sub_u):
						sub_req = urllib.request.Request(sub_u, headers={"User-Agent": "Mozilla/5.0 RevirDefteri/1.0"})
						with urllib.request.urlopen(sub_req, timeout=max(1.0, float(timeout)), context=ctx) as sub_r:
							sub_ct = sub_r.headers.get("Content-Type", "")
							sub_b = sub_r.read(1024)
							if is_meb_block_or_html_payload(sub_ct, sub_b) or b"#EXTM3U" not in sub_b[:256]:
								last_status = "meb_blocked"
								last_detail = "Alt HLS akışı engellendi veya geçersiz"
								continue

				if len(sample) > 0:
					return {
						"id": station.get("id", ""),
						"name": station.get("name", ""),
						"category": station.get("category", ""),
						"ok": True,
						"status_code": "ok",
						"status_label": "✅ Çalışıyor (Canlı Akış Hazır)",
						"tested_url": cand,
						"detail": f"HTTP {getattr(resp, 'status', 200)} · {ct or 'audio'}",
					}
		except ssl.SSLError as ssl_exc:
			last_status = "ssl_error"
			last_detail = f"SSL Sertifika Hatası: {ssl_exc}"
		except urllib.error.HTTPError as http_exc:
			last_status = "http_error"
			last_detail = f"HTTP {http_exc.code}: {http_exc.reason}"
		except urllib.error.URLError as url_exc:
			reason_s = str(url_exc.reason)
			if "SSL" in reason_s.upper() or "CERTIFICATE" in reason_s.upper():
				last_status = "ssl_error"
				last_detail = f"MEB / Sistem SSL Sertifika Hatası: {reason_s}"
			else:
				last_status = "dns_or_timeout"
				last_detail = f"Ağ / DNS Erişim Hatası: {reason_s}"
		except Exception as exc:
			last_status = "dns_or_timeout"
			last_detail = str(exc)

	label_map = {
		"meb_blocked": "⛔ MEB Filtresi / Ağ Engeli",
		"ssl_error": "🔒 MEB SSL Sertifika Hatası",
		"http_error": "⚠️ Yayın Sunucusu Hatası",
		"dns_or_timeout": "🌐 DNS / Zaman Aşımı Engeli",
	}
	return {
		"id": station.get("id", ""),
		"name": station.get("name", ""),
		"category": station.get("category", ""),
		"ok": False,
		"status_code": last_status,
		"status_label": label_map.get(last_status, "⚠️ Erişilemedi"),
		"tested_url": candidates[0] if candidates else "",
		"detail": last_detail,
	}


def diagnose_radio_connectivity(
	custom_stations_json: str = "",
	timeout: float = 3.2,
	max_stations: int = 12,
) -> dict[str, Any]:
	"""Yerleşik ve özel radyo istasyonlarını paralel test ederek MEB ağ teşhis raporu üretir."""
	stations = get_meb_safe_radio_stations(custom_stations_json=custom_stations_json, include_online=False)[:max_stations]
	results: list[dict[str, Any]] = []
	if not stations:
		return {"results": [], "ok_count": 0, "blocked_count": 0, "ssl_error_count": 0, "recommendation": ""}

	with ThreadPoolExecutor(max_workers=min(6, len(stations))) as pool:
		fut_map = {pool.submit(diagnose_single_station, st, timeout): idx for idx, st in enumerate(stations)}
		ordered: list[tuple[int, dict[str, Any]]] = []
		for fut in as_completed(fut_map):
			idx = fut_map[fut]
			try:
				ordered.append((idx, fut.result()))
			except Exception as exc:
				st_f = stations[idx]
				ordered.append((idx, {
					"id": st_f.get("id", ""),
					"name": st_f.get("name", ""),
					"category": st_f.get("category", ""),
					"ok": False,
					"status_code": "dns_or_timeout",
					"status_label": "🌐 Zaman Aşımı",
					"tested_url": st_f.get("url", ""),
					"detail": str(exc),
				}))
	ordered.sort(key=lambda x: x[0])
	results = [r for _, r in ordered]

	ok_cnt = sum(1 for r in results if r["ok"])
	blk_cnt = sum(1 for r in results if r["status_code"] == "meb_blocked")
	ssl_cnt = sum(1 for r in results if r["status_code"] == "ssl_error")

	if ssl_cnt > 0 and ok_cnt == 0:
		rec = "🔒 Bilgisayarda MEB kök sertifikası doğrulama hatası algılandı. Aşağıdan MEB_SERTIFIKASI.cer dosyasını seçin veya 'MEB Esnek SSL Modu'nu etkinleştirin."
	elif ok_cnt == 0:
		rec = "⛔ Bu ağda dış canlı radyo akışları tamamen engellenmiş görünüyor. Yan menüdeki '🌊 Ortam' (Çevrimdışı Sesler) veya '🎵 Klasörüm' (Yerel MP3) modunu kesintisiz kullanabilirsiniz."
	elif blk_cnt > 0:
		rec = f"ℹ️ Ticari radyoların bir kısmı ({blk_cnt} kanal) ağ filtresine takıldı ancak {ok_cnt} kanal (ör. TRT yayınları) aktif ve dinlenebilir durumda."
	else:
		rec = f"✅ Radyo bağlantısı sağlıklı ({ok_cnt}/{len(results)} istasyon aktif)."

	return {
		"results": results,
		"ok_count": ok_cnt,
		"blocked_count": blk_cnt,
		"ssl_error_count": ssl_cnt,
		"total_tested": len(results),
		"recommendation": rec,
	}


# ==============================================================================
# SAF PYTHON MP3 / M4A / FLAC KAPAK FOTOĞRAFI VE METADATA OKUYUCU
# ==============================================================================

def _decode_id3_text(payload: bytes) -> str:
	if not payload:
		return ""
	enc = payload[0]
	body = payload[1:]
	try:
		if enc == 0:
			txt = body.decode("latin-1", errors="replace")
		elif enc == 1:
			txt = body.decode("utf-16", errors="replace")
		elif enc == 2:
			txt = body.decode("utf-16-be", errors="replace")
		else:
			txt = body.decode("utf-8", errors="replace")
	except Exception:
		txt = ""
	return txt.split("\x00", 1)[0].strip()


def _parse_apic_frame(frame_data: bytes, id3_ver: int) -> tuple[bytes, str] | None:
	if len(frame_data) < 8:
		return None
	enc = frame_data[0]
	pos = 1
	if id3_ver == 2:
		fmt = frame_data[pos : pos + 3].decode("latin-1", errors="ignore").lower()
		mime = "image/png" if fmt == "png" else "image/jpeg"
		pos += 3
	else:
		end_mime = frame_data.find(b"\x00", pos)
		if end_mime < 0:
			return None
		mime = frame_data[pos:end_mime].decode("latin-1", errors="ignore").strip().lower() or "image/jpeg"
		pos = end_mime + 1

	if pos >= len(frame_data):
		return None
	pos += 1  # picture type byte

	# Açıklama (description) sonlandırıcısını atla
	if enc in (1, 2):
		while pos + 1 < len(frame_data):
			if frame_data[pos] == 0 and frame_data[pos + 1] == 0:
				pos += 2
				break
			pos += 2
	else:
		end_desc = frame_data.find(b"\x00", pos)
		if end_desc < 0:
			return None
		pos = end_desc + 1

	img_bytes = frame_data[pos:]
	# Bazı etiketlerde başta 0x00 kayması kalabilir; gerçek JPEG/PNG imzasını hizala
	if not img_bytes.startswith((b"\xff\xd8", b"\x89PNG", b"RIFF")):
		jpeg_i = img_bytes.find(b"\xff\xd8\xff")
		png_i = img_bytes.find(b"\x89PNG")
		if 0 <= jpeg_i <= 16:
			img_bytes = img_bytes[jpeg_i:]
			mime = "image/jpeg"
		elif 0 <= png_i <= 16:
			img_bytes = img_bytes[png_i:]
			mime = "image/png"

	if len(img_bytes) < 32:
		return None
	if "png" in mime or img_bytes.startswith(b"\x89PNG"):
		mime = "image/png"
	else:
		mime = "image/jpeg"
	return img_bytes, mime


def _extract_mp3_tags_and_cover(file_path: Path, want_cover: bool = True) -> dict[str, Any]:
	out: dict[str, Any] = {"title": "", "artist": "", "has_cover": False, "cover_bytes": None, "cover_mime": ""}
	try:
		with file_path.open("rb") as fh:
			hdr = fh.read(10)
			if len(hdr) < 10 or hdr[:3] != b"ID3":
				return out
			ver = hdr[3]
			flags = hdr[5]
			tag_size = (
				((hdr[6] & 0x7F) << 21)
				| ((hdr[7] & 0x7F) << 14)
				| ((hdr[8] & 0x7F) << 7)
				| (hdr[9] & 0x7F)
			)
			if tag_size <= 0:
				return out
			read_limit = min(tag_size, 4 * 1024 * 1024 if want_cover else 256 * 1024)
			buf = fh.read(read_limit)
	except Exception:
		return out

	pos = 0
	if (flags & 0x40) and len(buf) >= 4:
		if ver == 4:
			ext_sz = ((buf[0] & 0x7F) << 21) | ((buf[1] & 0x7F) << 14) | ((buf[2] & 0x7F) << 7) | (buf[3] & 0x7F)
			pos += max(4, ext_sz)
		else:
			ext_sz = struct.unpack(">I", buf[:4])[0]
			pos += 4 + ext_sz

	while pos < len(buf):
		if ver == 2:
			if pos + 6 > len(buf):
				break
			fid = buf[pos : pos + 3].decode("latin-1", errors="ignore")
			if not fid.strip("\x00"):
				break
			fsz = (buf[pos + 3] << 16) | (buf[pos + 4] << 8) | buf[pos + 5]
			f_flags = 0
			pos += 6
		else:
			if pos + 10 > len(buf):
				break
			fid = buf[pos : pos + 4].decode("latin-1", errors="ignore")
			if not fid.strip("\x00") or not re.match(r"^[A-Z0-9]{3,4}$", fid):
				break
			if ver == 4:
				fsz = ((buf[pos + 4] & 0x7F) << 21) | ((buf[pos + 5] & 0x7F) << 14) | ((buf[pos + 6] & 0x7F) << 7) | (buf[pos + 7] & 0x7F)
			else:
				fsz = struct.unpack(">I", buf[pos + 4 : pos + 8])[0]
			f_flags = (buf[pos + 8] << 8) | buf[pos + 9]
			pos += 10

		if fsz <= 0 or pos + fsz > len(buf):
			# Kapak çerçevesi read_limit sınırını aştıysa tam boyutuyla tekrar oku
			if fid in ("APIC", "PIC") and pos < len(buf):
				out["has_cover"] = True
			break

		fdata = buf[pos : pos + fsz]
		pos += fsz

		# ID3v2.4 data length indicator & unsynchronisation
		if ver == 4 and (f_flags & 0x0001) and len(fdata) >= 4:
			fdata = fdata[4:]
		if (ver == 4 and (f_flags & 0x0002)) or (flags & 0x80):
			fdata = fdata.replace(b"\xff\x00", b"\xff")

		if fid in ("TIT2", "TT2") and not out["title"]:
			out["title"] = _decode_id3_text(fdata)
		elif fid in ("TPE1", "TP1") and not out["artist"]:
			out["artist"] = _decode_id3_text(fdata)
		elif fid in ("APIC", "PIC") and not out["has_cover"]:
			out["has_cover"] = True
			if want_cover:
				parsed_pic = _parse_apic_frame(fdata, ver)
				if parsed_pic:
					out["cover_bytes"], out["cover_mime"] = parsed_pic

		if out["title"] and out["artist"] and out["has_cover"] and (not want_cover or out["cover_bytes"]):
			break
	return out


def _extract_flac_cover(file_path: Path, want_cover: bool = True) -> dict[str, Any]:
	out: dict[str, Any] = {"title": "", "artist": "", "has_cover": False, "cover_bytes": None, "cover_mime": ""}
	try:
		with file_path.open("rb") as fh:
			if fh.read(4) != b"fLaC":
				return out
			is_last = False
			while not is_last:
				bh = fh.read(4)
				if len(bh) < 4:
					break
				is_last = bool(bh[0] & 0x80)
				btype = bh[0] & 0x7F
				blen = (bh[1] << 16) | (bh[2] << 8) | bh[3]
				if btype == 6 and blen > 32:  # PICTURE
					out["has_cover"] = True
					if want_cover and blen <= 6 * 1024 * 1024:
						pb = fh.read(blen)
						off = 4
						if off + 4 <= len(pb):
							m_len = struct.unpack(">I", pb[off : off + 4])[0]
							off += 4
							mime = pb[off : off + m_len].decode("latin-1", errors="ignore") or "image/jpeg"
							off += m_len
							if off + 4 <= len(pb):
								d_len = struct.unpack(">I", pb[off : off + 4])[0]
								off += 4 + d_len + 16
								if off + 4 <= len(pb):
									img_len = struct.unpack(">I", pb[off : off + 4])[0]
									off += 4
									if img_len > 32 and off + img_len <= len(pb):
										out["cover_bytes"] = pb[off : off + img_len]
										out["cover_mime"] = mime
						break
				elif btype == 4 and blen <= 256 * 1024:  # VORBIS_COMMENT
					vb = fh.read(blen)
					if len(vb) >= 8:
						vlen = struct.unpack("<I", vb[:4])[0]
						off = 4 + vlen
						if off + 4 <= len(vb):
							ccnt = struct.unpack("<I", vb[off : off + 4])[0]
							off += 4
							for _ in range(min(ccnt, 40)):
								if off + 4 > len(vb):
									break
								clen = struct.unpack("<I", vb[off : off + 4])[0]
								off += 4
								if off + clen > len(vb):
									break
								kv = vb[off : off + clen].decode("utf-8", errors="ignore")
								off += clen
								if "=" in kv:
									k, v = kv.split("=", 1)
									ku = k.strip().upper()
									if ku == "TITLE" and not out["title"]:
										out["title"] = v.strip()
									elif ku == "ARTIST" and not out["artist"]:
										out["artist"] = v.strip()
				else:
					fh.seek(blen, os.SEEK_CUR)
	except Exception:
		pass
	return out


def _extract_m4a_cover(file_path: Path, want_cover: bool = True) -> dict[str, Any]:
	out: dict[str, Any] = {"title": "", "artist": "", "has_cover": False, "cover_bytes": None, "cover_mime": ""}
	try:
		with file_path.open("rb") as fh:
			raw = fh.read(3 * 1024 * 1024 if want_cover else 512 * 1024)
		idx = raw.find(b"covr")
		if idx >= 4:
			out["has_cover"] = True
			if want_cover:
				d_idx = raw.find(b"data", idx, min(len(raw), idx + 64))
				if d_idx >= 4:
					atom_sz = struct.unpack(">I", raw[d_idx - 4 : d_idx])[0]
					img_start = d_idx + 12
					img_end = d_idx - 4 + atom_sz
					if 32 < (img_end - img_start) and img_end <= len(raw):
						ib = raw[img_start:img_end]
						mime = "image/png" if ib.startswith(b"\x89PNG") else "image/jpeg"
						out["cover_bytes"] = ib
						out["cover_mime"] = mime
	except Exception:
		pass
	return out


def _find_folder_cover_image(audio_file: Path) -> Path | None:
	parent = audio_file.parent
	for ext in (".jpg", ".jpeg", ".png", ".webp"):
		same_stem = parent / f"{audio_file.stem}{ext}"
		if same_stem.is_file():
			return same_stem
	for cname in COVER_FILE_CANDIDATES:
		cp = parent / cname
		if cp.is_file():
			return cp
	return None


def extract_audio_metadata_and_cover(file_path: Path, *, want_cover_bytes: bool = False) -> dict[str, Any]:
	"""Ses dosyasından başlık, sanatçı ve gömülü/klasör kapak fotoğrafını çıkarır."""
	ext = file_path.suffix.lower()
	info: dict[str, Any] = {"title": "", "artist": "", "has_cover": False, "cover_bytes": None, "cover_mime": ""}
	if ext == ".mp3":
		info = _extract_mp3_tags_and_cover(file_path, want_cover=want_cover_bytes)
	elif ext == ".flac":
		info = _extract_flac_cover(file_path, want_cover=want_cover_bytes)
	elif ext == ".m4a":
		info = _extract_m4a_cover(file_path, want_cover=want_cover_bytes)

	if not info.get("has_cover") or (want_cover_bytes and not info.get("cover_bytes")):
		sidecar = _find_folder_cover_image(file_path)
		if sidecar is not None:
			info["has_cover"] = True
			if want_cover_bytes:
				try:
					raw_c = sidecar.read_bytes()
					if len(raw_c) <= 8 * 1024 * 1024:
						info["cover_bytes"] = raw_c
						info["cover_mime"] = mimetypes.guess_type(sidecar.name)[0] or "image/jpeg"
				except Exception:
					pass
	return info


def build_fallback_cover_svg(title: str, subtitle: str = "") -> bytes:
	"""Gömülü kapağı olmayan parçalar için estetik renk geçişli SVG albüm kapağı üretir."""
	palettes = [
		("#0d9488", "#115e59"),
		("#2563eb", "#1e3a8a"),
		("#7c3aed", "#4c1d95"),
		("#db2777", "#831843"),
		("#d97706", "#78350f"),
		("#059669", "#064e3b"),
	]
	h_idx = int(hashlib.md5((title or "revir").encode("utf-8")).hexdigest()[:6], 16) % len(palettes)
	c1, c2 = palettes[h_idx]
	words = [w for w in re.split(r"\s+", (title or "").strip()) if w]
	initials = "".join(w[0].upper() for w in words[:2] if w[0].isalnum()) or "♫"
	esc_init = html.escape(initials[:2])
	svg = (
		'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120" width="120" height="120">'
		"<defs>"
		f'<linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="{c1}"/><stop offset="100%" stop-color="{c2}"/></linearGradient>'
		"</defs>"
		'<rect width="120" height="120" rx="20" fill="url(#g)"/>'
		'<circle cx="60" cy="60" r="40" fill="none" stroke="rgba(255,255,255,0.22)" stroke-width="3"/>'
		'<circle cx="60" cy="60" r="14" fill="rgba(255,255,255,0.16)"/>'
		f'<text x="60" y="68" text-anchor="middle" font-family="system-ui,-apple-system,sans-serif" font-size="26" font-weight="800" fill="#ffffff">{esc_init}</text>'
		'<text x="96" y="30" text-anchor="middle" font-size="18" fill="rgba(255,255,255,0.8)">♪</text>'
		"</svg>"
	)
	return svg.encode("utf-8")


# ==============================================================================
# ÇEVRİMDIŞI DİNLENDİRİCİ ORTAM SESİ WAV ÜRETECİ (SIFIR İNTERNET)
# ==============================================================================

_AMBIENT_WAV_CACHE: dict[str, bytes] = {}
_AMBIENT_WAV_LOCK = threading.Lock()


def generate_ambient_wav_bytes(preset_id: str, duration_sec: float = 8.0, sample_rate: int = 16000) -> bytes:
	"""Belirtilen dinlendirici ortam sesi için pürüzsüz döngülü 16-bit PCM WAV ses verisi üretir."""
	key = f"{preset_id}:{duration_sec}:{sample_rate}"
	with _AMBIENT_WAV_LOCK:
		if key in _AMBIENT_WAV_CACHE:
			return _AMBIENT_WAV_CACHE[key]

	total_samples = max(4000, int(duration_sec * sample_rate))
	rng = random.Random(preset_id)
	samples = bytearray(total_samples * 2)
	fade_n = min(int(sample_rate * 0.4), total_samples // 6)

	lp1 = 0.0
	lp2 = 0.0
	two_pi = 2.0 * math.pi

	for i in range(total_samples):
		t = i / float(sample_rate)
		if preset_id == "amb_calm432":
			# Yumuşak harmonik 432 Hz + 216 Hz + 648 Hz akustik pad
			mod = 0.72 + 0.28 * math.sin(two_pi * t / duration_sec)
			val = (
				0.48 * math.sin(two_pi * 216.0 * t)
				+ 0.34 * math.sin(two_pi * 432.0 * t)
				+ 0.14 * math.sin(two_pi * 540.0 * t)
			) * mod * 0.28
		elif preset_id == "amb_waves":
			# Yavaş yükselen ve çekilen deniz dalgası zarfı
			white = rng.uniform(-1.0, 1.0)
			lp1 = lp1 * 0.965 + white * 0.035
			swell = 0.30 + 0.70 * (0.5 + 0.5 * math.sin(two_pi * t / (duration_sec * 0.5) - math.pi / 2.0))
			val = lp1 * swell * 0.75
		elif preset_id == "amb_breeze":
			white = rng.uniform(-1.0, 1.0)
			lp1 = lp1 * 0.982 + white * 0.018
			lp2 = lp2 * 0.94 + lp1 * 0.06
			gust = 0.55 + 0.45 * math.sin(two_pi * t / duration_sec)
			val = lp2 * gust * 0.95
		else:
			# amb_rain: yumuşak filtrelenmiş yağmur dokusu
			white = rng.uniform(-1.0, 1.0)
			lp1 = lp1 * 0.86 + white * 0.14
			val = lp1 * 0.36

		# Döngü geçişinde çıtlamayı önleyen yumuşak giriş/çıkış zarfı
		if i < fade_n:
			val *= i / float(fade_n)
		elif i >= total_samples - fade_n:
			val *= (total_samples - 1 - i) / float(fade_n)

		pcm = max(-32767, min(32767, int(val * 32767.0)))
		struct.pack_into("<h", samples, i * 2, pcm)

	data_size = len(samples)
	wav_hdr = struct.pack(
		"<4sI4s4sIHHIIHH4sI",
		b"RIFF",
		36 + data_size,
		b"WAVE",
		b"fmt ",
		16,
		1,  # PCM
		1,  # Mono
		sample_rate,
		sample_rate * 2,
		2,
		16,
		b"data",
		data_size,
	)
	wav_bytes = wav_hdr + bytes(samples)
	with _AMBIENT_WAV_LOCK:
		_AMBIENT_WAV_CACHE[key] = wav_bytes
	return wav_bytes


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
	Zaten ham AAC/ADTS veya ID3 etiketli AAC segmenti ise baştaki ID3 başlığını temizleyip döndürür.
	"""
	n = len(ts_bytes)
	if n < 188:
		return _strip_leading_id3(ts_bytes)
	sync_off = -1
	for i in range(min(376, n - 188 + 1)):
		if ts_bytes[i] == 0x47 and (i + 188 >= n or ts_bytes[i + 188] == 0x47):
			sync_off = i
			break
	if sync_off < 0:
		return _strip_leading_id3(ts_bytes)

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
	best_pid = max(pes_by_pid, key=lambda k: len(pes_by_pid[k]))
	return _strip_leading_id3(bytes(pes_by_pid[best_pid]))


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
	127.0.0.1 üzerinde dinamik portta çalışan yerel medya sunucusu:
	- /stream/<id>  : HTTP Range (206) yerel MP3/WAV/OGG/M4A/FLAC akışı
	- /cover/<id>   : Albüm kapak fotoğrafı (gömülü ID3/M4A/FLAC, klasör resmi veya SVG kapak)
	- /hls_live     : MEB-uyumlu HTTPS .m3u8 canlı ses köprüsü (ön doğrulamalı + ID3 temizlemeli)
	- /radio_proxy  : MEB-uyumlu HTTPS doğrudan MP3/AAC akış köprüsü
	- /ambient/<id> : Sıfır internet çevrimdışı dinlendirici ortam sesi (WAV)
	"""

	def __init__(self, host: str = "127.0.0.1", port: int = 0) -> None:
		self.host = host
		self._lock = threading.RLock()
		self._secret = secrets.token_bytes(24)
		self._root_folder: Path | None = None
		self._tracks_by_id: dict[str, Path] = {}
		self._cached_tracks: list[dict[str, Any]] = []
		self._cached_folder_key: str = ""
		self._cover_cache: dict[str, tuple[bytes, str]] = {}
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
		"""Seçilen klasörü ve alt klasörlerini ses dosyaları ve kapak meta verileri için tarar."""
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
				self._cover_cache.clear()
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
			meta = extract_audio_metadata_and_cover(resolved_f, want_cover_bytes=False)
			clean_file_title = _clean_track_display_title(resolved_f)
			display_title = clean_file_title
			tag_title = str(meta.get("title") or "").strip()
			tag_artist = str(meta.get("artist") or "").strip()

			playlist.append({
				"id": tid,
				"title": display_title,
				"tag_title": tag_title,
				"artist": tag_artist,
				"folder": folder_hint,
				"rel_path": rel_str,
				"ext": resolved_f.suffix.lower().lstrip("."),
				"size_bytes": size_bytes,
				"has_cover": bool(meta.get("has_cover")),
				"cover_url": f"{self.base_url}/cover/{tid}",
				"url": f"{self.base_url}/stream/{tid}",
			})

		with self._lock:
			self._root_folder = resolved_root
			self._tracks_by_id = new_map
			self._cached_tracks = playlist
			self._cached_folder_key = cache_key
			self._cover_cache.clear()
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

	def get_track_cover_payload(self, track_id: str) -> tuple[bytes, str] | None:
		clean_id = re.sub(r"[^a-f0-9]", "", str(track_id or "").lower())
		if not clean_id:
			return None
		with self._lock:
			cached = self._cover_cache.get(clean_id)
		if cached is not None:
			return cached

		fpath = self.resolve_track_path(clean_id)
		if fpath is None:
			return None

		meta = extract_audio_metadata_and_cover(fpath, want_cover_bytes=True)
		cov_b = meta.get("cover_bytes")
		cov_m = str(meta.get("cover_mime") or "image/jpeg")
		if isinstance(cov_b, (bytes, bytearray)) and len(cov_b) > 32:
			res = (bytes(cov_b), cov_m)
		else:
			title_hint = str(meta.get("title") or _clean_track_display_title(fpath))
			res = (build_fallback_cover_svg(title_hint), "image/svg+xml")

		with self._lock:
			if len(self._cover_cache) > 120:
				self._cover_cache.clear()
			self._cover_cache[clean_id] = res
		return res

	def get_ambient_presets(self) -> list[dict[str, Any]]:
		self.start()
		out: list[dict[str, Any]] = []
		for p in BUILTIN_AMBIENT_PRESETS:
			pid = p["id"]
			out.append({
				"id": pid,
				"name": p["name"],
				"category": p["category"],
				"icon": p["icon"],
				"url": f"{self.base_url}/ambient/{pid}",
			})
		return out

	def _build_handler(self):
		server_ref = self

		class _AudioRangeHandler(BaseHTTPRequestHandler):
			protocol_version = "HTTP/1.1"

			def log_message(self, format: str, *args) -> None:
				return

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

				if path == "/radio_proxy":
					qs = urllib.parse.parse_qs(parsed.query)
					stream_url = str((qs.get("url") or [""])[0]).strip()
					self._stream_direct_radio_proxy(stream_url, send_body=send_body)
					return

				if path.startswith("/ambient/"):
					preset_id = urllib.parse.unquote(path[len("/ambient/") :]).split("/", 1)[0].strip()
					valid_ids = {p["id"] for p in BUILTIN_AMBIENT_PRESETS}
					if preset_id not in valid_ids:
						self._send_simple(HTTPStatus.NOT_FOUND, "Ambient Preset Not Found")
						return
					wav_data = generate_ambient_wav_bytes(preset_id)
					self._serve_bytes_with_range(wav_data, "audio/wav", send_body=send_body)
					return

				if path.startswith("/cover/"):
					track_id = urllib.parse.unquote(path[len("/cover/") :]).split("/", 1)[0]
					cov = server_ref.get_track_cover_payload(track_id)
					if cov is None:
						self._send_simple(HTTPStatus.NOT_FOUND, "Cover Not Found")
						return
					cov_bytes, cov_mime = cov
					self.send_response(HTTPStatus.OK)
					self.send_header("Content-Type", cov_mime)
					self.send_header("Content-Length", str(len(cov_bytes)))
					self.send_header("Cache-Control", "public, max-age=3600")
					self.send_header("Access-Control-Allow-Origin", "*")
					self.end_headers()
					if send_body:
						try:
							self.wfile.write(cov_bytes)
						except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
							pass
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

			def _serve_bytes_with_range(self, raw_data: bytes, mime: str, send_body: bool) -> None:
				file_size = len(raw_data)
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
				self.end_headers()
				if send_body and length > 0:
					try:
						self.wfile.write(raw_data[start : end + 1])
					except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
						pass

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

			def _stream_direct_radio_proxy(self, stream_url: str, send_body: bool) -> None:
				"""Doğrudan HTTPS (443) MP3/AAC radyo akışlarını sistem/MEB SSL deposu ile yerel <audio>'ya aktarır."""
				if not is_meb_safe_stream_url(stream_url):
					self._send_simple(HTTPStatus.FORBIDDEN, "Only standard HTTPS (443) streams are allowed")
					return

				ssl_ctx = build_system_ssl_context()
				req = urllib.request.Request(
					stream_url,
					headers={"User-Agent": "Mozilla/5.0 RevirDefteri/1.0", "Icy-MetaData": "0", "Accept": "*/*"},
				)
				try:
					resp = urllib.request.urlopen(req, timeout=6.0, context=ssl_ctx)
				except Exception as exc:
					self._send_simple(HTTPStatus.BAD_GATEWAY, f"Radyo akisina baglanilamadi: {exc}")
					return

				try:
					final_u = str(resp.geturl() or stream_url)
					if not is_meb_safe_stream_url(final_u):
						self._send_simple(HTTPStatus.BAD_GATEWAY, "Unsafe redirect target")
						return
					ct = resp.headers.get("Content-Type") or "audio/mpeg"
					first_chunk = resp.read(4096)
					if not first_chunk or is_meb_block_or_html_payload(ct, first_chunk):
						self._send_simple(HTTPStatus.BAD_GATEWAY, "Blocked by firewall HTML page")
						return

					clean_ct = ct.split(";", 1)[0].strip().lower()
					if clean_ct in ("audio/aacp", "application/octet-stream"):
						clean_ct = "audio/aac" if first_chunk[:2] in (b"\xff\xf1", b"\xff\xf9") else "audio/mpeg"

					self.send_response(HTTPStatus.OK)
					self.send_header("Content-Type", clean_ct)
					self.send_header("Cache-Control", "no-store")
					self.send_header("Access-Control-Allow-Origin", "*")
					self.send_header("Connection", "close")
					self.end_headers()

					if not send_body:
						return

					self.wfile.write(_strip_leading_id3(first_chunk))
					self.wfile.flush()
					started = time.monotonic()
					while (time.monotonic() - started) < 4 * 3600:
						chunk = resp.read(16384)
						if not chunk:
							break
						self.wfile.write(chunk)
						self.wfile.flush()
				except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
					return
				finally:
					try:
						resp.close()
					except Exception:
						pass

			def _stream_hls_as_audio(self, m3u8_url: str, send_body: bool) -> None:
				"""
				Standart HTTPS (443) .m3u8 canlı radyo yayınını alt liste ve MEB engel kontrolü yaparak
				doğrular, ardından tarayıcı <audio> etiketine saf ADTS AAC / MP3 akışı olarak aktarır.
				"""
				if not is_meb_safe_stream_url(m3u8_url):
					self._send_simple(HTTPStatus.FORBIDDEN, "Only standard HTTPS (443) streams are allowed")
					return

				ssl_ctx = build_system_ssl_context()
				headers = {"User-Agent": "Mozilla/5.0 RevirDefteri/1.0"}

				def _fetch_checked(target_url: str, timeout_s: float = 5.5) -> tuple[bytes, str, str]:
					if not is_meb_safe_stream_url(target_url):
						raise ValueError("Unsafe redirect/segment URL")
					req = urllib.request.Request(target_url, headers=headers)
					with urllib.request.urlopen(req, timeout=timeout_s, context=ssl_ctx) as r:
						final_u = str(r.geturl() or target_url)
						if not is_meb_safe_stream_url(final_u):
							raise ValueError("Redirected to non-HTTPS/non-443 URL")
						ct = r.headers.get("Content-Type", "")
						body = r.read(2 * 1024 * 1024)
						if is_meb_block_or_html_payload(ct, body):
							raise ValueError("MEB firewall HTML block page")
						return body, final_u, ct

				try:
					raw_pl, base_u, _ = _fetch_checked(m3u8_url, timeout_s=5.5)
					text_pl = raw_pl.decode("utf-8", errors="replace")
					if "#EXTM3U" not in text_pl:
						raise ValueError("Not a valid HLS M3U8 playlist")

					media_url = base_u
					first_pl_txt = text_pl
					first_pl_base = base_u

					# Master playlist ise alt medya listesini HTTP 200 göndermeden ÖNCE doğrula!
					if "#EXT-X-STREAM-INF" in text_pl:
						variants = [
							urllib.parse.urljoin(base_u, ln.strip())
							for ln in text_pl.splitlines()
							if ln.strip() and not ln.strip().startswith("#")
						]
						variants = [v for v in variants if is_meb_safe_stream_url(v)]
						sub_ok = False
						for cand_v in variants:
							try:
								sub_b, sub_base, _ = _fetch_checked(cand_v, timeout_s=5.5)
								sub_t = sub_b.decode("utf-8", errors="replace")
								if "#EXTM3U" in sub_t:
									media_url = cand_v
									first_pl_txt = sub_t
									first_pl_base = sub_base
									sub_ok = True
									break
							except Exception:
								continue
						if not sub_ok:
							raise ValueError("All HLS sub-variants failed or 404")
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

				seen_segs: list[str] = []
				max_duration_sec = 4 * 3600
				started = time.monotonic()
				cached_first: tuple[str, str] | None = (first_pl_txt, first_pl_base)

				while (time.monotonic() - started) < max_duration_sec:
					if cached_first is not None:
						pl_txt, cur_base = cached_first
						cached_first = None
					else:
						try:
							pl_bytes, cur_base, _ = _fetch_checked(media_url, timeout_s=5.5)
							pl_txt = pl_bytes.decode("utf-8", errors="replace")
						except Exception:
							break

					target_dur = 3.0
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
							seg_raw, _, _ = _fetch_checked(su, timeout_s=6.0)
							audio_bytes = _extract_adts_from_mpegts(seg_raw)
							if audio_bytes:
								self.wfile.write(audio_bytes)
								self.wfile.flush()
								sent_any = True
						except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
							return
						except Exception:
							continue

					time.sleep(0.8 if sent_any else min(2.0, target_dur * 0.45))

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
	revir_ui.py yan menü oynatıcısına gönderilecek yapılandırma, MEB-güvenli radyo istasyonları,
	çevrimdışı ortam sesleri ve albüm kapaklı yerel MP3 çalma listesini hazırlar.
	"""
	configure_ssl_options(
		custom_cert_path=str(app_cfg.get("meb_custom_cert_path", "")).strip(),
		allow_relaxed_ssl=(str(app_cfg.get("meb_ssl_relaxed", "0")).strip() == "1"),
	)
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
	ambient = srv.get_ambient_presets() if visible else []
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
		"ambient_presets": ambient,
	}
