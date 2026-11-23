"""Unit tests for bump_version.py, build_release.py, desktop_app.py, and desktop_launcher.py."""
from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import bump_version
import build_release
import desktop_app


class TestBumpVersion(unittest.TestCase):
    def test_read_current_version_happy_and_missing_and_invalid(self):
        v = bump_version.read_current_version()
        self.assertRegex(v, r"^\d+\.\d+\.\d+$")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_p = Path(tmp)
            with mock.patch.object(bump_version, "VERSION_FILE", tmp_p / "VERSION"):
                self.assertEqual(bump_version.read_current_version(), "1.1.0")
                (tmp_p / "VERSION").write_text("bad_ver\n", encoding="utf-8")
                with self.assertRaises(ValueError):
                    bump_version.read_current_version()
                (tmp_p / "VERSION").write_text("2.3.4\n", encoding="utf-8")
                self.assertEqual(bump_version.read_current_version(), "2.3.4")

    def test_compute_next_version_all_modes_and_errors(self):
        self.assertEqual(bump_version.compute_next_version("1.2.3", "patch"), "1.2.4")
        self.assertEqual(bump_version.compute_next_version("1.2.3", "minor"), "1.3.0")
        self.assertEqual(bump_version.compute_next_version("1.2.3", "major"), "2.0.0")
        self.assertEqual(bump_version.compute_next_version("1.2.3", "5.6.7"), "5.6.7")
        self.assertEqual(bump_version.compute_next_version(" 0.0.0 ", " PATCH "), "0.0.1")
        with self.assertRaises(ValueError):
            bump_version.compute_next_version("1.2", "patch")
        with self.assertRaises(ValueError):
            bump_version.compute_next_version("1.2.3", "invalid_mode")

    def test_sync_version_references_isolated(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_p = Path(tmp)
            (tmp_p / "VERSION").write_text("1.0.0\n", encoding="utf-8")
            (tmp_p / "persantil.py").write_text('_DEFAULT_APP_VERSION = "1.0.0"\n', encoding="utf-8")
            (tmp_p / "desktop_app.py").write_text('_DEFAULT_APP_VERSION = "1.0.0"\n', encoding="utf-8")
            (tmp_p / "RevirDefteri_Setup.iss").write_text(
                '#define MyAppVersion "1.0.0"\n; Çıktı: installer\\RevirDefteri_Kurulum_v1.0.0.exe\n',
                encoding="utf-8",
            )
            (tmp_p / "kurulum_olustur.iss").write_text('#define MyAppVersion "1.0.0"\n', encoding="utf-8")

            with mock.patch.object(bump_version, "ROOT_DIR", tmp_p), \
                 mock.patch.object(bump_version, "VERSION_FILE", tmp_p / "VERSION"):
                changed = bump_version.sync_version_references("1.0.1")
                self.assertIn("VERSION", changed)
                self.assertIn("persantil.py", changed)
                self.assertIn("desktop_app.py", changed)
                self.assertEqual(bump_version.sync_version_references("1.0.1"), [])

    def test_main_modes_isolated(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_p = Path(tmp)
            (tmp_p / "VERSION").write_text("1.0.0\n", encoding="utf-8")
            with mock.patch.object(bump_version, "ROOT_DIR", tmp_p), \
                 mock.patch.object(bump_version, "VERSION_FILE", tmp_p / "VERSION"):
                self.assertEqual(bump_version.main(["--current"]), 0)
                self.assertEqual(bump_version.main(["--sync"]), 0)
                self.assertEqual(bump_version.main(["patch"]), 0)
                self.assertEqual((tmp_p / "VERSION").read_text(encoding="utf-8").strip(), "1.0.1")


class TestBuildRelease(unittest.TestCase):
    def test_find_iscc_executable(self):
        with mock.patch("shutil.which", return_value="/mock/ISCC.exe"):
            self.assertEqual(build_release.find_iscc_executable(), Path("/mock/ISCC.exe"))
        with mock.patch("shutil.which", return_value=None), \
             mock.patch.object(Path, "exists", return_value=False):
            self.assertIsNone(build_release.find_iscc_executable())

    def test_prompt_target_version(self):
        with mock.patch("builtins.input", return_value=""):
            self.assertEqual(build_release._prompt_target_version("1.2.3", "patch"), "1.2.4")
        with mock.patch("builtins.input", return_value="minor"):
            self.assertEqual(build_release._prompt_target_version("1.2.3", "patch"), "1.3.0")

    def test_clean_old_installers(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            keep = d / "RevirDefteri_Kurulum_v1.2.0.exe"
            old = d / "RevirDefteri_Kurulum_v1.1.0.exe"
            other = d / "notes.txt"
            for f in (keep, old, other):
                f.write_text("x", encoding="utf-8")
            build_release._clean_old_installers(d, keep.name)
            self.assertTrue(keep.exists())
            self.assertFalse(old.exists())
            self.assertTrue(other.exists())
            build_release._clean_old_installers(d / "nonexistent", keep.name)

    def test_parse_args_and_main_and_build_release_mocked(self):
        ns = build_release.parse_args(["--version", "minor", "--yes"])
        self.assertEqual(ns.version_arg, "minor")
        self.assertTrue(ns.yes)

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            spec = root / "RevirDefteri.spec"
            iss = root / "RevirDefteri_Setup.iss"
            dist_exe = root / "dist" / "RevirDefteri" / "RevirDefteri.exe"
            inst_dir = root / "installer"
            pkg_dir = root / "Kurulum_Paketleri"
            spec.write_text("# spec", encoding="utf-8")
            iss.write_text("# iss", encoding="utf-8")
            dist_exe.parent.mkdir(parents=True, exist_ok=True)
            dist_exe.write_bytes(b"MZ")
            inst_dir.mkdir(parents=True, exist_ok=True)
            (inst_dir / "RevirDefteri_Kurulum_v1.2.4.exe").write_bytes(b"SETUP")

            with mock.patch.object(build_release, "PROJECT_ROOT", root), \
                 mock.patch.object(build_release, "SPEC_FILE", spec), \
                 mock.patch.object(build_release, "ISS_FILE", iss), \
                 mock.patch.object(build_release, "DIST_EXE", dist_exe), \
                 mock.patch.object(build_release, "INSTALLER_DIR", inst_dir), \
                 mock.patch.object(build_release, "PACKAGES_DIR", pkg_dir), \
                 mock.patch.object(build_release, "sync_version_references", return_value=["VERSION"]), \
                 mock.patch.object(build_release, "find_iscc_executable", return_value=Path("/mock/ISCC.exe")), \
                 mock.patch.object(build_release, "read_current_version", return_value="1.2.3"), \
                 mock.patch("subprocess.run"):
                self.assertEqual(build_release.main(["--yes"]), 0)
                self.assertTrue((root / "RevirDefteri_Kurulum_v1.2.4.exe").exists())
                self.assertTrue((pkg_dir / "RevirDefteri_Kurulum_v1.2.4.exe").exists())


class TestDesktopApp(unittest.TestCase):
    def test_ensure_standard_streams_and_load_version(self):
        desktop_app._ensure_standard_streams()
        v = desktop_app._load_app_version()
        self.assertTrue(len(v) > 0)

    def test_is_dir_writable_and_resolve_runtime_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp) / "sub"
            self.assertTrue(desktop_app._is_dir_writable(d))
            with mock.patch.dict(os.environ, {"REVIR_DATA_DIR": str(d)}):
                paths = desktop_app.resolve_runtime_paths()
                self.assertEqual(paths["data_dir"], d.resolve())

    def test_window_state_save_and_load(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            self.assertEqual(desktop_app._load_window_state(d), {})
            (d / ".window_state.json").write_text("{corrupt", encoding="utf-8")
            self.assertEqual(desktop_app._load_window_state(d), {})
            (d / ".window_state.json").write_text("[]", encoding="utf-8")
            self.assertEqual(desktop_app._load_window_state(d), {})

            fake_win = mock.Mock(width=1280, height=800, x=50, y=60)
            desktop_app._save_window_state(d, fake_win)
            loaded = desktop_app._load_window_state(d)
            self.assertEqual(loaded, {"width": 1280, "height": 800, "x": 50, "y": 60})
            desktop_app._save_window_state(d, None)

    def test_single_instance_lock_and_release(self):
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(
            desktop_app, "MUTEX_NAME", "Local\\OkulRevirDefteri_SingleInstance_Mutex_UnitTest"
        ):
            d = Path(tmp)
            ok = desktop_app.acquire_single_instance_lock(d)
            self.assertTrue(ok)
            desktop_app.release_single_instance_lock()

    def test_bridge_save_download(self):
        srv = mock.Mock()
        srv.get_download_item.return_value = None
        bridge = desktop_app._RevirDesktopBridge(srv, {"window": None})
        res = bridge.save_download("missing")
        self.assertFalse(res["saved"])

        srv.get_download_item.return_value = (b"hello", "rapor.txt", "text/plain")
        res2 = bridge.save_download("tok1")
        self.assertFalse(res2["saved"])

        with tempfile.TemporaryDirectory() as tmp:
            out_file = Path(tmp) / "saved.txt"
            fake_win = mock.Mock()
            fake_win.create_file_dialog.return_value = [str(out_file)]
            bridge2 = desktop_app._RevirDesktopBridge(srv, {"window": fake_win})
            with mock.patch.dict(sys.modules, {"webview": mock.Mock(SAVE_DIALOG=30)}):
                res3 = bridge2.save_download("tok1", "rapor.txt")
                self.assertTrue(res3["saved"])
                self.assertEqual(out_file.read_bytes(), b"hello")

    def test_bring_existing_window_to_front_non_win(self):
        if sys.platform != "win32":
            self.assertFalse(desktop_app._bring_existing_window_to_front())


if __name__ == "__main__":
    unittest.main(verbosity=2)
