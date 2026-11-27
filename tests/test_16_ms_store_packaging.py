"""Microsoft Store MSIX paketleme, ikon üretimi ve çalışma zamanı uyumluluk testleri."""
from __future__ import annotations

import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest import mock

from PIL import Image

import build_msix
import desktop_app
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestMicrosoftStorePackaging(unittest.TestCase):
	def test_quad_version_conversion(self) -> None:
		self.assertEqual(build_msix.to_msix_quad_version("1.51.1"), "1.51.1.0")
		self.assertEqual(build_msix.to_msix_quad_version("v2.4.9"), "2.4.9.0")
		self.assertEqual(build_msix.to_msix_quad_version("1.0.0.7"), "1.0.0.0")
		with self.assertRaises(ValueError):
			build_msix.to_msix_quad_version("1.5")

	def test_manifest_generation_and_schema(self) -> None:
		cfg = build_msix.load_msix_config()
		xml_str = build_msix.render_appx_manifest(cfg, "1.52.0.0")
		root = ET.fromstring(xml_str.encode("utf-8"))
		ns = {
			"m": "http://schemas.microsoft.com/appx/manifest/foundation/windows10",
			"uap": "http://schemas.microsoft.com/appx/manifest/uap/windows10",
			"rescap": "http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities",
		}
		ident = root.find("m:Identity", ns)
		self.assertIsNotNone(ident)
		self.assertEqual(ident.attrib["Version"], "1.52.0.0")
		self.assertEqual(ident.attrib["ProcessorArchitecture"], "x64")

		app_el = root.find("m:Applications/m:Application", ns)
		self.assertIsNotNone(app_el)
		self.assertEqual(app_el.attrib["Executable"], "RevirDefteri.exe")
		self.assertEqual(app_el.attrib["EntryPoint"], "Windows.FullTrustApplication")

		full_trust = root.find("m:Capabilities/rescap:Capability[@Name='runFullTrust']", ns)
		mic_cap = root.find("m:Capabilities/m:DeviceCapability[@Name='microphone']", ns)
		self.assertIsNotNone(full_trust)
		self.assertIsNotNone(mic_cap)

	def test_store_assets_and_staging_hygiene(self) -> None:
		with tempfile.TemporaryDirectory() as tmp:
			tmp_p = Path(tmp)
			fake_dist = tmp_p / "dist" / "RevirDefteri"
			(fake_dist / "_internal").mkdir(parents=True)
			(fake_dist / "kisi_verileri").mkdir(parents=True)
			(fake_dist / "RevirDefteri.exe").write_bytes(b"MZ-fake-exe")
			(fake_dist / "revir_defteri.db").write_bytes(b"secret-user-db")
			(fake_dist / "RevirDefteri_RootCA.cer").write_bytes(b"self-signed-root")
			(fake_dist / "kisi_verileri" / "p1.json").write_text("{}", encoding="utf-8")

			staging = tmp_p / "staging"
			sub_assets = tmp_p / "submission"
			cfg = build_msix.load_msix_config()
			manifest_path = build_msix.stage_msix_directory(
				dist_dir=fake_dist,
				staging_dir=staging,
				icon_path=build_msix.PROJECT_ROOT / "app_icon.png",
				cfg=cfg,
				quad_version="1.52.0.0",
				submission_assets_dir=sub_assets,
			)
			self.assertTrue(manifest_path.exists())
			self.assertTrue((staging / "RevirDefteri.exe").exists())
			self.assertTrue((staging / "MS_STORE_BUILD").exists())
			self.assertTrue((staging / "_internal" / "MS_STORE_BUILD").exists())

			# Kullanıcı veritabanı ve kök sertifika kesinlikle MSIX içine girmemeli
			self.assertFalse((staging / "revir_defteri.db").exists())
			self.assertFalse((staging / "RevirDefteri_RootCA.cer").exists())
			self.assertFalse((staging / "kisi_verileri").exists())

			# Zorunlu MSIX ikon boyutlarını doğrula
			with Image.open(staging / "Assets" / "Square44x44Logo.png") as im44:
				self.assertEqual(im44.size, (44, 44))
			with Image.open(staging / "Assets" / "Square150x150Logo.png") as im150:
				self.assertEqual(im150.size, (150, 150))
			with Image.open(staging / "Assets" / "Wide310x150Logo.png") as im_wide:
				self.assertEqual(im_wide.size, (310, 150))
			with Image.open(staging / "Assets" / "StoreLogo.png") as im_store:
				self.assertEqual(im_store.size, (50, 50))
			self.assertTrue((sub_assets / "StoreListing_Logo_300x300.png").exists())
			self.assertTrue((sub_assets / "StoreListing_Hero_1920x1080.png").exists())

			# makeappx.exe ile gerçek .msix paketlemesini doğrula
			makeappx = build_msix.ensure_makeappx_exe(auto_download=True)
			self.assertIsNotNone(makeappx)
			out_msix = tmp_p / "RevirDefteri_Test.msix"
			build_msix.pack_msix(makeappx, staging, out_msix)
			self.assertTrue(out_msix.exists())
			self.assertGreater(out_msix.stat().st_size, 4096)

	def test_store_runtime_detection_and_self_updater_guard(self) -> None:
		with tempfile.TemporaryDirectory() as tmp:
			tmp_p = Path(tmp)
			(tmp_p / "MS_STORE_BUILD").write_text("microsoft-store\n", encoding="utf-8")
			self.assertTrue(desktop_app.is_ms_store_runtime(exe_dir=tmp_p))

		with mock.patch.dict("os.environ", {"REVIR_MS_STORE": "1"}, clear=False):
			self.assertTrue(persantil.is_microsoft_store_build())
			ok_dl, path_dl, msg_dl = persantil.download_github_release_asset(
				"https://github.com/blue-doggy/revir_defteri/releases/download/v1.0.0/RevirDefteri_Kurulum_v1.0.0.exe",
				"RevirDefteri_Kurulum_v1.0.0.exe",
			)
			self.assertFalse(ok_dl)
			self.assertIsNone(path_dl)
			self.assertIn("Microsoft Store", msg_dl)

			revir_ui.reset_session()
			with isolated_real_db():
				revir_ui.session_state["main_menu_page"] = "⚙️ Ayarlar"
				revir_ui.session_state["settings_vertical_nav"] = "💾 Veri & Sistem"
				out = revir_ui.run_action_headless(persantil.main, {"action": "init"})
				html_out = out.get("main_html") or ""
				self.assertNotIn("Çalışma Zamanı Hatası", html_out)
				self.assertNotIn("btn_chk_updates", html_out)
				self.assertIn("Microsoft Store", html_out)
				self.assertIn("Tıbbi Sorumluluk Beyanı", html_out)


if __name__ == "__main__":
	unittest.main()
