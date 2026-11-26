"""
Yan Menü Radyo & Yerel MP3 Çalar Modülü Otomatik Test Paketi (test_12_media_player.py)
"""
from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
import urllib.error
import urllib.request

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

import media_player
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestMediaPlayerModule(unittest.TestCase):
	def test_meb_safe_stream_url_filtering(self) -> None:
		# Standart HTTPS (443) akışları kabul edilmeli
		self.assertTrue(media_player.is_meb_safe_stream_url("https://radio-trtfm.live.trt.com.tr/master.m3u8"))
		self.assertTrue(media_player.is_meb_safe_stream_url("https://playerservices.streamtheworld.com:443/api/livestream-redirect/JOY_TURK_SC"))
		self.assertTrue(media_player.is_meb_safe_stream_url("https://ntvradyo.radyotvonline.net/ntvradyo"))

		# Şifresiz HTTP ve standart dışı portlar (:8000, :8080, :9999) reddedilmeli
		self.assertFalse(media_player.is_meb_safe_stream_url("http://radio-trtfm.live.trt.com.tr/master.m3u8"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:8000/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:8080/live.aac"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://ornek-radyo.com:9999/;"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://127.0.0.1/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url("https://localhost/stream.mp3"))
		self.assertFalse(media_player.is_meb_safe_stream_url(""))

		# Tüm yerleşik istasyonlar ve yedek URL'ler MEB güvenli olmalı
		self.assertGreaterEqual(len(media_player.BUILTIN_MEB_SAFE_STATIONS), 10)
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

			res_init = revir_ui.run_action_headless(persantil.main, {"action": "init"})
			mp_cfg = res_init.get("media_player")
			self.assertIsInstance(mp_cfg, dict)
			self.assertTrue(mp_cfg["visible"])
			self.assertGreaterEqual(mp_cfg["station_count"], 10)
			self.assertTrue(any(s.get("hls_proxy_url") for s in mp_cfg["stations"] if s.get("is_hls")))

			custom_mp3_dir = tmp_p / "secili_mp3ler"
			custom_mp3_dir.mkdir()
			(custom_mp3_dir / "Tenefus_Zili.mp3").write_bytes(b"ID3" + b"\x00" * 256)

			act_res = persantil._handle_media_player_action({"action": "set_folder", "path": str(custom_mp3_dir)})
			self.assertTrue(act_res["ok"])
			self.assertEqual(act_res["media_player"]["track_count"], 1)
			self.assertEqual(act_res["media_player"]["tracks"][0]["title"], "Tenefus Zili")
			self.assertEqual(persantil.get_app_settings()["custom_music_folder"], str(custom_mp3_dir.resolve()))

			persantil.save_app_settings({"show_sidebar_media_player": "0"})
			res_hidden = revir_ui.run_action_headless(persantil.main, {"action": "init"})
			self.assertFalse(res_hidden["media_player"]["visible"])


if __name__ == "__main__":
	unittest.main()
