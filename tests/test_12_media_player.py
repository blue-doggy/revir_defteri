"""
Yan Menü MEB Radyo, Çevrimdışı Ortam Sesleri & Kapaklı Yerel MP3 Çalar Test Paketi (test_12_media_player.py)
"""
from __future__ import annotations

import http.client
import json
from pathlib import Path
import ssl
import struct
import sys
import tempfile
import unittest
from unittest import mock
import urllib.error
import urllib.parse
import urllib.request

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

import media_player
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


def _build_id3v23_mp3_with_cover(title: str, artist: str, png_bytes: bytes) -> bytes:
	def _frame(fid: str, payload: bytes) -> bytes:
		return fid.encode("ascii") + struct.pack(">I", len(payload)) + b"\x00\x00" + payload

	tit2 = _frame("TIT2", b"\x03" + title.encode("utf-8") + b"\x00")
	tpe1 = _frame("TPE1", b"\x03" + artist.encode("utf-8") + b"\x00")
	apic_payload = b"\x00image/png\x00\x03Cover\x00" + png_bytes
	apic = _frame("APIC", apic_payload)
	frames = tit2 + tpe1 + apic
	sz = len(frames)
	syncsafe = bytes([
		(sz >> 21) & 0x7F,
		(sz >> 14) & 0x7F,
		(sz >> 7) & 0x7F,
		sz & 0x7F,
	])
	return b"ID3\x03\x00\x00" + syncsafe + frames + (b"\xff\xfb\x90\x64" * 32)


class TestMediaPlayerModule(unittest.TestCase):
	def test_meb_safe_stream_url_filtering(self) -> None:
		# Standart HTTPS (443) akışları kabul edilmeli
		self.assertTrue(media_player.is_meb_safe_stream_url("https://rd-trtfm.medya.trt.com.tr/master.m3u8"))
		self.assertTrue(media_player.is_meb_safe_stream_url("https://playerservices.streamtheworld.com:443/api/livestream-redirect/JOY_TURK_SC"))
		self.assertTrue(media_player.is_meb_safe_stream_url("https://dygedge2.radyotvonline.net/ntvradyo/playlist.m3u8"))

		# Şifresiz HTTP ve standart dışı portlar (:8000, :8080, :9999) reddedilmeli
		self.assertFalse(media_player.is_meb_safe_stream_url("http://rd-trtfm.medya.trt.com.tr/master.m3u8"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:8000/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:8080/live.aac"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:9999/;"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://127.0.0.1/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://localhost/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url(""))

		# Tüm yerleşik istasyonlar ve yedek URL'ler MEB güvenli olmalı
		self.assertGreaterEqual(len(media_player.BUILTIN_MEB_SAFE_STATIONS), 7)
		for st in media_player.BUILTIN_MEB_SAFE_STATIONS:
			self.assertTrue(media_player.is_meb_safe_stream_url(st["url"]), f"Güvensiz yerleşik istasyon: {st}")
			for b_url in st.get("backup_urls", []):
				self.assertTrue(media_player.is_meb_safe_stream_url(b_url), f"Güvensiz yedek URL: {b_url}")

	def test_radio_browser_offline_fallback_and_port_filtering(self) -> None:
		# 1. MEB güvenlik duvarı Radio Browser API'yi engellediğinde hata vermeden yerleşik listeye dönmeli
		with mock.patch("urllib.request.urlopen", side_effect=urllib.error.URLError("Connection timed out")):
			online = media_player.fetch_online_meb_safe_stations(timeout=0.6)
			self.assertEqual(online, [])
			merged = media_player.get_meb_safe_radio_stations(include_online=True, timeout=0.6)
			self.assertGreaterEqual(len(merged), len(media_player.BUILTIN_MEB_SAFE_STATIONS))
			self.assertTrue(any("TRT FM" in s["name"] for s in merged))

		# 2. API yanıt döndüğünde yalnızca standart HTTPS (443) istasyonları almalı
		fake_api_rows = [
			{"name": "Engelli HTTP Radyo", "url_resolved": "http://stream.example.com/live.mp3", "codec": "MP3"},
			{"name": "Engelli Port 8000", "url_resolved": "https://stream.example.com:8000/live.mp3", "codec": "MP3"},
			{"name": "Uyumlu HTTPS Radyo", "url_resolved": "https://clean.example.com/live.aac", "codec": "AAC"},
		]
		fake_bytes = json.dumps(fake_api_rows).encode("utf-8")

		class _FakeResp:
			headers = {"Content-Type": "application/json"}

			def __enter__(self):
				return self

			def __exit__(self, *args):
				return False

			def read(self, *_args):
				return fake_bytes

		with mock.patch("urllib.request.urlopen", return_value=_FakeResp()):
			filtered = media_player.fetch_online_meb_safe_stations(timeout=1.0)
			self.assertEqual(len(filtered), 1)
			self.assertEqual(filtered[0]["name"], "Uyumlu HTTPS Radyo")
			self.assertEqual(filtered[0]["url"], "https://clean.example.com/live.aac")

	def test_local_media_server_recursive_scan_and_http_range(self) -> None:
		with tempfile.TemporaryDirectory() as td:
			music_root = Path(td) / "okul_muzikleri"
			sub_dir = music_root / "klasik_eserler"
			sub_dir.mkdir(parents=True)

			mp3_payload = bytes((i * 7 + 13) % 256 for i in range(4096))
			wav_payload = b"RIFF" + b"\x00" * 512
			(music_root / "01_Sabah_Melodisi.mp3").write_bytes(mp3_payload)
			(sub_dir / "02_Dinleti.wav").write_bytes(wav_payload)
			(sub_dir / "03_Fonetik.flac").write_bytes(b"fLaC" + b"\x11" * 128)
			(music_root / "notlar.txt").write_text("bu bir ses dosyasi degildir", encoding="utf-8")

			srv = media_player.LocalMediaStreamServer(host="127.0.0.1", port=0)
			srv.start()
			try:
				tracks = srv.scan_folder(music_root, force=True)
				self.assertEqual(len(tracks), 3)
				titles = [t["title"] for t in tracks]
				self.assertIn("01 Sabah Melodisi", titles)
				self.assertIn("02 Dinleti", titles)

				mp3_track = next(t for t in tracks if t["ext"] == "mp3")
				stream_url = mp3_track["url"]

				# 1. Tam dosya isteği -> 200 OK + Accept-Ranges: bytes
				with urllib.request.urlopen(stream_url, timeout=3.0) as resp:
					self.assertEqual(resp.status, 200)
					self.assertEqual(resp.headers.get("Accept-Ranges"), "bytes")
					self.assertEqual(resp.headers.get("Content-Type"), "audio/mpeg")
					self.assertEqual(resp.read(), mp3_payload)

				# 2. HTTP Range isteği -> 206 Partial Content + tam bayt dilimi
				req_range = urllib.request.Request(stream_url, headers={"Range": "bytes=100-249"})
				with urllib.request.urlopen(req_range, timeout=3.0) as resp_part:
					self.assertEqual(resp_part.status, 206)
					self.assertEqual(resp_part.headers.get("Content-Range"), f"bytes 100-249/{len(mp3_payload)}")
					self.assertEqual(int(resp_part.headers.get("Content-Length") or 0), 150)
					self.assertEqual(resp_part.read(), mp3_payload[100:250])

				# 3. Sondan bayt aralığı (bytes=-64) -> 206 Partial Content
				req_suffix = urllib.request.Request(stream_url, headers={"Range": "bytes=-64"})
				with urllib.request.urlopen(req_suffix, timeout=3.0) as resp_suf:
					self.assertEqual(resp_suf.status, 206)
					self.assertEqual(resp_suf.read(), mp3_payload[-64:])

				# 4. Geçersiz Range -> 416 Range Not Satisfiable
				req_bad = urllib.request.Request(stream_url, headers={"Range": "bytes=999999-1000010"})
				with self.assertRaises(urllib.error.HTTPError) as exc_info:
					urllib.request.urlopen(req_bad, timeout=3.0)
				self.assertEqual(exc_info.exception.code, 416)

				# 5. Bilinmeyen/geçersiz parça kimliği -> 404 Not Found
				with self.assertRaises(urllib.error.HTTPError) as exc_404:
					urllib.request.urlopen(f"{srv.base_url}/stream/deadbeef0000", timeout=3.0)
				self.assertEqual(exc_404.exception.code, 404)
			finally:
				srv.stop()

	def test_mp3_id3v2_cover_and_metadata_extraction_and_cover_endpoint(self) -> None:
		with tempfile.TemporaryDirectory() as td:
			root = Path(td) / "kapakli_arsiv"
			album_dir = root / "album_kapakli"
			plain_dir = root / "kapaksiz"
			album_dir.mkdir(parents=True)
			plain_dir.mkdir(parents=True)

			fake_png = b"\x89PNG\r\n\x1a\n" + b"\x42" * 80
			mp3_bytes = _build_id3v23_mp3_with_cover("Bahar Senfonisi", "Revir Oda Orkestrası", fake_png)
			(root / "01_bahar.mp3").write_bytes(mp3_bytes)

			fake_jpg = b"\xff\xd8\xff\xe0" + b"\x77" * 80
			(album_dir / "cover.jpg").write_bytes(fake_jpg)
			(album_dir / "02_yan_flut.wav").write_bytes(b"RIFF" + b"\x00" * 256)

			(plain_dir / "03_yalniz_piyano.wav").write_bytes(b"RIFF" + b"\x00" * 256)

			srv = media_player.LocalMediaStreamServer(host="127.0.0.1", port=0)
			srv.start()
			try:
				tracks = srv.scan_folder(root, force=True)
				self.assertEqual(len(tracks), 3)

				tr_mp3 = next(t for t in tracks if t["rel_path"] == "01_bahar.mp3")
				self.assertTrue(tr_mp3["has_cover"])
				self.assertEqual(tr_mp3["tag_title"], "Bahar Senfonisi")
				self.assertEqual(tr_mp3["artist"], "Revir Oda Orkestrası")
				with urllib.request.urlopen(tr_mp3["cover_url"], timeout=3.0) as r1:
					self.assertEqual(r1.status, 200)
					self.assertEqual(r1.headers.get("Content-Type"), "image/png")
					self.assertEqual(r1.read(), fake_png)

				tr_sidecar = next(t for t in tracks if "02_yan_flut.wav" in t["rel_path"])
				self.assertTrue(tr_sidecar["has_cover"])
				with urllib.request.urlopen(tr_sidecar["cover_url"], timeout=3.0) as r2:
					self.assertEqual(r2.status, 200)
					self.assertEqual(r2.headers.get("Content-Type"), "image/jpeg")
					self.assertEqual(r2.read(), fake_jpg)

				tr_svg = next(t for t in tracks if "03_yalniz_piyano.wav" in t["rel_path"])
				self.assertFalse(tr_svg["has_cover"])
				with urllib.request.urlopen(tr_svg["cover_url"], timeout=3.0) as r3:
					self.assertEqual(r3.status, 200)
					self.assertEqual(r3.headers.get("Content-Type"), "image/svg+xml")
					svg_body = r3.read()
					self.assertIn(b"<svg", svg_body)
			finally:
				srv.stop()

	def test_offline_ambient_wav_generator_and_endpoint(self) -> None:
		wav_bytes = media_player.generate_ambient_wav_bytes("amb_calm432", duration_sec=1.0, sample_rate=8000)
		self.assertTrue(wav_bytes.startswith(b"RIFF"))
		self.assertIn(b"WAVEfmt ", wav_bytes[:24])
		self.assertGreater(len(wav_bytes), 16000)

		srv = media_player.LocalMediaStreamServer(host="127.0.0.1", port=0)
		srv.start()
		try:
			presets = srv.get_ambient_presets()
			self.assertGreaterEqual(len(presets), 4)
			rain_url = next(p["url"] for p in presets if p["id"] == "amb_rain")

			with urllib.request.urlopen(rain_url, timeout=4.0) as resp:
				self.assertEqual(resp.status, 200)
				self.assertEqual(resp.headers.get("Content-Type"), "audio/wav")
				data = resp.read(64)
				self.assertTrue(data.startswith(b"RIFF"))

			req_part = urllib.request.Request(rain_url, headers={"Range": "bytes=0-43"})
			with urllib.request.urlopen(req_part, timeout=4.0) as resp_part:
				self.assertEqual(resp_part.status, 206)
				self.assertEqual(len(resp_part.read()), 44)

			with self.assertRaises(urllib.error.HTTPError) as exc_404:
				urllib.request.urlopen(f"{srv.base_url}/ambient/olmayan_ses", timeout=3.0)
			self.assertEqual(exc_404.exception.code, 404)
		finally:
			srv.stop()

	def test_meb_firewall_block_detection_id3_strip_and_proxy_chain(self) -> None:
		# 1. MEB / FortiGuard HTML engel sayfası algılama
		self.assertTrue(media_player.is_meb_block_or_html_payload("text/html; charset=utf-8", b"<html><body>FortiGuard</body></html>"))
		self.assertTrue(media_player.is_meb_block_or_html_payload("audio/aac", b"<!DOCTYPE html><html>meb.gov.tr erisim engellenmistir</html>"))
		self.assertFalse(media_player.is_meb_block_or_html_payload("audio/mpeg", b"\xff\xfb\x90\x64\x00\x11\x22\x33"))

		# 2. TRT AAC segment başındaki ID3v2 etiketini sıyırma
		id3_hdr = b"ID3\x04\x00\x00\x00\x00\x00\x0a" + (b"X" * 10)
		adts_tail = b"\xff\xf1\x50\x80\x02\x1f\xfcAAC_DATA"
		self.assertEqual(media_player._strip_leading_id3(id3_hdr + adts_tail), adts_tail)

		# 3. Hem .m3u8 hem doğrudan HTTPS akışlar için proxy_chain üretimi
		stations = media_player.get_meb_safe_radio_stations(hls_proxy_base_url="http://127.0.0.1:19876")
		for st in stations:
			chain = st.get("proxy_chain") or []
			self.assertGreaterEqual(len(chain), 1)
			if st.get("is_hls"):
				self.assertTrue(chain[0].startswith("http://127.0.0.1:19876/hls_live?url="))
			else:
				self.assertTrue(chain[0].startswith("http://127.0.0.1:19876/radio_proxy?url="))

		# 4. /radio_proxy ve /hls_live HTML engel veya 404 alt listede HTTP 200 yerine anında 502 dönmeli
		class _BlockedHtmlResp:
			status = 200
			headers = {"Content-Type": "text/html; charset=utf-8"}

			def __enter__(self):
				return self

			def __exit__(self, *args):
				return False

			def geturl(self):
				return "https://blocked.example.com/live.mp3"

			def close(self):
				pass

			def read(self, *_args):
				return b"<!DOCTYPE html><html><body>FortiGuard Web Filtering - Access Blocked</body></html>"

		srv = media_player.LocalMediaStreamServer(host="127.0.0.1", port=0)
		srv.start()
		try:
			with mock.patch("urllib.request.urlopen", return_value=_BlockedHtmlResp()):
				conn = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=3.0)
				q_u = urllib.parse.quote("https://blocked.example.com/live.mp3", safe="")
				conn.request("GET", f"/radio_proxy?url={q_u}")
				r_proxy = conn.getresponse()
				r_proxy.read()
				conn.close()
				self.assertEqual(r_proxy.status, 502)

				conn2 = http.client.HTTPConnection("127.0.0.1", srv.port, timeout=3.0)
				q_m3u8 = urllib.parse.quote("https://blocked.example.com/master.m3u8", safe="")
				conn2.request("GET", f"/hls_live?url={q_m3u8}")
				r_hls = conn2.getresponse()
				r_hls.read()
				conn2.close()
				self.assertEqual(r_hls.status, 502)
		finally:
			srv.stop()

	def test_ssl_configuration_and_meb_connectivity_diagnostics(self) -> None:
		try:
			media_player.configure_ssl_options(custom_cert_path="", allow_relaxed_ssl=True)
			ctx_relaxed = media_player.build_system_ssl_context()
			self.assertEqual(ctx_relaxed.verify_mode, ssl.CERT_NONE)
		finally:
			media_player.configure_ssl_options(custom_cert_path="", allow_relaxed_ssl=False)

		st_sample = {"id": "test_st", "name": "Test Radyo", "category": "Ulusal", "url": "https://clean.example.com/live.mp3"}

		class _AudioOkResp:
			status = 200
			headers = {"Content-Type": "audio/mpeg"}

			def __enter__(self):
				return self

			def __exit__(self, *args):
				return False

			def geturl(self):
				return "https://clean.example.com/live.mp3"

			def read(self, *_args):
				return b"\xff\xfb\x90\x64" * 64

		with mock.patch("urllib.request.urlopen", return_value=_AudioOkResp()):
			d_ok = media_player.diagnose_single_station(st_sample, timeout=1.5)
			self.assertTrue(d_ok["ok"])
			self.assertEqual(d_ok["status_code"], "ok")

		with mock.patch("urllib.request.urlopen", side_effect=ssl.SSLError("CERTIFICATE_VERIFY_FAILED")):
			d_ssl = media_player.diagnose_single_station(st_sample, timeout=1.5)
			self.assertFalse(d_ssl["ok"])
			self.assertEqual(d_ssl["status_code"], "ssl_error")
			report = media_player.diagnose_radio_connectivity(timeout=1.0, max_stations=3)
			self.assertEqual(report["ok_count"], 0)
			self.assertEqual(report["ssl_error_count"], 3)
			self.assertIn("MEB", report["recommendation"])

	def test_mpegts_audio_pes_extraction(self) -> None:
		def _pad_pkt(header_and_payload: bytes) -> bytes:
			self.assertLessEqual(len(header_and_payload), 188)
			return header_and_payload + (b"\xFF" * (188 - len(header_and_payload)))

		pat_sec = bytes([0x00, 0xB0, 0x0D, 0x00, 0x01, 0xC1, 0x00, 0x00, 0x00, 0x01, 0xE1, 0x00, 0x00, 0x00, 0x00, 0x00])
		pat_pkt = _pad_pkt(bytes([0x47, 0x40, 0x00, 0x10, 0x00]) + pat_sec)

		pmt_sec = bytes([
			0x02, 0xB0, 0x12, 0x00, 0x01, 0xC1, 0x00, 0x00, 0xE1, 0x01, 0xF0, 0x00,
			0x0F, 0xE1, 0x01, 0xF0, 0x00,
			0x00, 0x00, 0x00, 0x00,
		])
		pmt_pkt = _pad_pkt(bytes([0x47, 0x41, 0x00, 0x10, 0x00]) + pmt_sec)

		sample_adts = b"\xFF\xF1\x50\x80\x02\x1F\xFCDEADBEEF_AUDIO_FRAME"
		pes_pkt_data = bytes([0x00, 0x00, 0x01, 0xC0, 0x00, 0x00, 0x80, 0x00, 0x00]) + sample_adts
		ts_hdr = bytes([0x47, 0x41, 0x01, 0x30, 188 - 5 - len(pes_pkt_data)]) + (b"\xFF" * (188 - 5 - len(pes_pkt_data)))
		aud_pkt = ts_hdr + pes_pkt_data

		extracted = media_player._extract_adts_from_mpegts(pat_pkt + pmt_pkt + aud_pkt)
		self.assertEqual(extracted, sample_adts)

	def test_persantil_sidebar_media_player_integration(self) -> None:
		revir_ui.reset_session()
		with isolated_real_db() as (tmp_p, _db_p, _connect):
			self.assertIn('id="revir-persistent-audio"', revir_ui._SHELL_HTML)
			self.assertIn("window.__revirMediaControl", revir_ui._SHELL_HTML)
			self.assertIn("navigator.mediaSession", revir_ui._SHELL_HTML)
			self.assertIn("revir-mp-cover", revir_ui._SHELL_HTML)
			self.assertIn("revir-mp-list-btn", revir_ui._SHELL_HTML)
			self.assertIn("revir-mp-playlist-panel", revir_ui._SHELL_HTML)
			self.assertIn('data-mp-mode="ambient"', revir_ui._SHELL_HTML)
			self.assertIn("meb_custom_cert_path", persantil.DEFAULT_APP_SETTINGS)
			self.assertIn("meb_ssl_relaxed", persantil.DEFAULT_APP_SETTINGS)

			res_init = revir_ui.run_action_headless(persantil.main, {"action": "init"})
			mp_cfg = res_init.get("media_player")
			self.assertIsInstance(mp_cfg, dict)
			self.assertTrue(mp_cfg["visible"])
			self.assertGreaterEqual(mp_cfg["station_count"], 7)
			self.assertGreaterEqual(len(mp_cfg.get("ambient_presets") or []), 4)
			self.assertTrue(any(s.get("hls_proxy_url") for s in mp_cfg["stations"] if s.get("is_hls")))
			self.assertTrue(all(len(s.get("proxy_chain") or []) >= 1 for s in mp_cfg["stations"]))

			custom_mp3_dir = tmp_p / "secili_mp3ler"
			custom_mp3_dir.mkdir()
			fake_png = b"\x89PNG\r\n\x1a\n" + b"\x55" * 64
			(custom_mp3_dir / "Tenefus_Zili.mp3").write_bytes(
				_build_id3v23_mp3_with_cover("Teneffüs Melodisi", "Okul Zili", fake_png)
			)

			act_res = persantil._handle_media_player_action({"action": "set_folder", "path": str(custom_mp3_dir)})
			self.assertTrue(act_res["ok"])
			self.assertEqual(act_res["media_player"]["track_count"], 1)
			tr0 = act_res["media_player"]["tracks"][0]
			self.assertEqual(tr0["title"], "Tenefus Zili")
			self.assertEqual(tr0["tag_title"], "Teneffüs Melodisi")
			self.assertEqual(tr0["artist"], "Okul Zili")
			self.assertTrue(tr0["has_cover"])
			self.assertIn("/cover/", tr0["cover_url"])
			self.assertEqual(persantil.get_app_settings()["custom_music_folder"], str(custom_mp3_dir.resolve()))

			persantil.save_app_settings({"show_sidebar_media_player": "0"})
			res_hidden = revir_ui.run_action_headless(persantil.main, {"action": "init"})
			self.assertFalse(res_hidden["media_player"]["visible"])

	def test_windows_taskbar_thumbnail_media_toolbar_and_bridge(self) -> None:
		import threading
		import desktop_app

		stop_ev = threading.Event()
		dispatched: list[str] = []

		class _FakeWin:
			def evaluate_js(self, code: str) -> None:
				dispatched.append(code)

		holder = {"window": _FakeWin()}
		tb = desktop_app._WindowsTaskbarMediaToolbar(holder, stop_ev)
		bridge = desktop_app._RevirDesktopBridge(None, holder, taskbar_toolbar=tb)

		res_sync = bridge.sync_taskbar_media_state(True, "TRT FM", True)
		self.assertTrue(res_sync["ok"])
		self.assertTrue(tb._is_playing)
		self.assertEqual(tb._title, "TRT FM")

		arr_playing = tb._build_thumb_buttons_array()
		self.assertEqual(len(arr_playing), 3)
		self.assertEqual(arr_playing[0].iId, tb.BTN_PREV)
		self.assertEqual(arr_playing[1].iId, tb.BTN_PLAY_PAUSE)
		self.assertEqual(arr_playing[1].szTip, "Duraklat")
		self.assertEqual(arr_playing[2].iId, tb.BTN_NEXT)

		bridge.sync_taskbar_media_state(False, "TRT FM", True)
		arr_paused = tb._build_thumb_buttons_array()
		self.assertEqual(arr_paused[1].szTip, "Oynat")

		if sys.platform == "win32":
			import ctypes
			from ctypes import wintypes
			import time
			import clr  # type: ignore[import-not-found]
			clr.AddReference("System.Windows.Forms")
			from System.Windows.Forms import Form  # type: ignore[import-not-found]

			user32 = ctypes.windll.user32
			user32.DestroyIcon.argtypes = [wintypes.HICON]
			user32.DestroyIcon.restype = wintypes.BOOL
			for kind in ("prev", "play", "pause", "next"):
				hi = tb._create_media_hicon(kind, 20)
				self.assertNotEqual(hi, 0)
				user32.DestroyIcon(wintypes.HICON(hi))

			frm = Form()
			try:
				hwnd = int(frm.Handle.ToInt64())
				self.assertTrue(tb.install_subclass_on_hwnd(hwnd))
				user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t]
				user32.SendMessageW.restype = ctypes.c_ssize_t
				user32.SendMessageW(hwnd, tb.WM_COMMAND, (tb.THBN_CLICKED << 16) | tb.BTN_PREV, 0)
				user32.SendMessageW(hwnd, tb.WM_COMMAND, (tb.THBN_CLICKED << 16) | tb.BTN_PLAY_PAUSE, 0)
				user32.SendMessageW(hwnd, tb.WM_COMMAND, (tb.THBN_CLICKED << 16) | tb.BTN_NEXT, 0)
				for _ in range(20):
					if len(dispatched) >= 3:
						break
					time.sleep(0.02)
				joined = "\n".join(dispatched)
				self.assertIn("'prev'", joined)
				self.assertIn("'play_pause'", joined)
				self.assertIn("'next'", joined)
			finally:
				tb.close()
				frm.Dispose()
		stop_ev.set()


if __name__ == "__main__":
	unittest.main()
