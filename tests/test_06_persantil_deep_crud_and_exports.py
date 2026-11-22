"""Deep CRUD, exports, main() render, and edge-case tests for persantil.py."""
from __future__ import annotations

import sys
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest import mock

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestPersantilDeepCRUDAndExports(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        revir_ui.reset_session()

    def tearDown(self):
        self.ctx.__exit__(None, None, None)

    def test_main_headless_render(self):
        with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
            "İmsak": "06:00", "Güneş": "07:30", "Öğle": "12:50", "İkindi": "15:20", "Akşam": "17:45", "Yatsı": "19:05"
        }):
            out = revir_ui.run_action_headless(persantil.main, {"action": "render"})
        self.assertIn("main_html", out)
        self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])

    def test_docx_and_excel_export_builders_with_turkish_unicode(self):
        docx_bytes = persantil.build_ai_custom_report_docx(
            title="Özel Sağlık Raporu ĞÜŞİÖÇ",
            subtitle="Alt Başlık",
            raw_markdown="## Bulgular\n- Öğrenci durumu normal.\n- **Takip** önerilir.",
        )
        self.assertTrue(isinstance(docx_bytes, (bytes, bytearray)) and len(docx_bytes) > 100)

        xlsx_bytes = persantil.build_ai_custom_report_excel(
            title="Tarama Listesi ĞÜŞİÖÇ",
            df=pd.DataFrame([{"Ad Soyad": "Ömer Şahin", "Sınıf": "4-A", "Boy": 140.5}]),
            raw_markdown="Not 1: Normal",
        )
        self.assertTrue(isinstance(xlsx_bytes, (bytes, bytearray)) and len(xlsx_bytes) > 100)

    def test_dialogs_render_headless(self):
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('4-A')")
            c.execute(
                "INSERT INTO students(class_id, full_name, student_no, sex, birth_date, person_type) VALUES (1, 'Ömer Şahin', '101', 'E', '2016-05-01', 'Öğrenci')"
            )
            c.execute(
                "INSERT INTO visits(student_id, visit_date, visit_time, complaint) VALUES (1, '2026-11-22', '10:00', 'Baş ağrısı')"
            )
            c.commit()
        persantil.set_active_person(1)
        revir_ui.session_state["active_person_id"] = 1

        for dlg_call in [
            lambda: persantil.open_class_manager_dialog(),
            lambda: persantil.open_new_profile_dialog("Öğrenci"),
            lambda: persantil.open_edit_profile_dialog(1),
            lambda: persantil.open_add_medication_dialog(1),
            lambda: persantil.open_screening_dialog(1),
            lambda: persantil.open_person_dossier_dialog(1),
            lambda: persantil.open_gemini_api_guide_dialog(),
        ]:
            with self.subTest(dlg=dlg_call):
                revir_ui.reset_session()
                revir_ui.session_state["active_person_id"] = 1
                out = revir_ui.run_action_headless(dlg_call, {"action": "render"})
                self.assertIn("main_html", out)
                self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])

    def test_edge_cases_empty_and_invalid_inputs(self):
        # Empty / None / extreme values on core helpers
        self.assertEqual(persantil.format_file_size(0), "0 B")
        self.assertIsInstance(persantil.format_file_size(-5), str)
        self.assertEqual(persantil._slugify_person_name("Şükran Öğütçü!"), "sukran_ogutcu")
        self.assertIsNone(persantil.age_in_months(None, date(2026, 11, 22)))
        self.assertIsNone(persantil.age_in_months("", date(2026, 11, 22)))
        self.assertIsNone(persantil.evaluate_measurement(None, "Boy (cm)", "E", 120))
        self.assertIsNone(persantil.evaluate_measurement(-10.0, "Boy (cm)", "E", 120))
        self.assertIsNone(persantil.get_person_by_id(999999))

    def test_gemini_chat_pdf_create_profiles_classes_and_teachers(self):
        import base64
        import json
        import gemini_assistant

        fake_pdf = b"%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF"
        ai_payload = {
            "action": "create_profiles",
            "classes": [{"name": "4-B", "teacher_name": "Fatma Şahin"}],
            "students": [
                {
                    "full_name": "Zeynep Öztürk",
                    "class_name": "4-B",
                    "student_no": "214",
                    "sex": "K",
                    "birth_date": "2016-03-11",
                    "guardian_name": "Mehmet Öztürk",
                    "guardian_phone": "0532 000 11 22",
                    "allergies": "Polen",
                    "blood_group": "A Rh+",
                }
            ],
            "staff": [
                {
                    "full_name": "Fatma Şahin",
                    "role": "Öğretmen",
                    "class_name": "4-B",
                    "sex": "K",
                    "phone": "0505 111 22 33",
                }
            ],
        }
        fake_reply = f"PDF incelendi, sınıf ve profiller oluşturuluyor.\n[[REVIR_EXPORT:{json.dumps(ai_payload, ensure_ascii=False)}]]"
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, fake_reply, "gemini-3-flash-preview")):
            ok, clean_txt, _m, _tr, export_spec, _w = gemini_assistant.chat_with_gemini_assistant(
                api_key="test",
                model="gemini-3-flash-preview",
                user_message="Ekli PDF'den sınıfları, öğretmenleri ve öğrencileri oluştur",
                attachments=[{"name": "sinif_listesi.pdf", "type": "application/pdf", "data_b64": base64.b64encode(fake_pdf).decode("ascii")}],
            )
        self.assertTrue(ok)
        self.assertIsInstance(export_spec, dict)
        self.assertEqual(export_spec["action"], "create_profiles")

        dls, note = persantil._execute_gemini_chat_export(export_spec, None, clean_txt)
        self.assertEqual(dls, [])
        self.assertIn("1 sınıf", note)
        self.assertIn("1 öğrenci profili", note)
        self.assertIn("1 öğretmen/personel profili", note)

        with self.connect() as c:
            cls_row = c.execute("SELECT name, teacher_name FROM classes WHERE name='4-B'").fetchone()
            self.assertIsNotNone(cls_row)
            self.assertEqual(cls_row["teacher_name"], "Fatma Şahin")
            stu_row = c.execute("SELECT id, full_name, student_no, person_type FROM students WHERE full_name='Zeynep Öztürk'").fetchone()
            self.assertIsNotNone(stu_row)
            self.assertEqual(stu_row["person_type"], "Öğrenci")
            hp_row = c.execute("SELECT allergies, blood_group FROM health_profiles WHERE student_id=?", (stu_row["id"],)).fetchone()
            self.assertEqual(hp_row["allergies"], "Polen")
            stf_row = c.execute("SELECT full_name, person_type FROM students WHERE full_name='Fatma Şahin'").fetchone()
            self.assertIsNotNone(stf_row)
            self.assertEqual(stf_row["person_type"], "Öğretmen")

    def test_gemini_chat_requires_confirmation_before_write_or_delete_and_preserves_classes(self):
        import json
        import gemini_assistant

        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('Mevcut-1A')")
            c.execute("INSERT INTO classes(name) VALUES ('Silinecek-Bos')")
            c.commit()

        # 1. Simulate chat create_profiles -> MUST return write_proposal and NOT mutate DB yet
        create_payload = {
            "action": "create_profiles",
            "classes": [{"name": "Yeni-2A", "teacher_name": "Kemal Demir"}],
            "students": [{"full_name": "Ece Yılmaz", "class_name": "Yeni-2A", "student_no": "501", "sex": "K"}],
        }
        fake_reply = f"Yeni-2A sınıfı ve Ece Yılmaz eklenecek. Onaylıyor musunuz?\n[[REVIR_EXPORT:{json.dumps(create_payload, ensure_ascii=False)}]]"
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
             mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Yeni-2A sınıfı eklenecek.", "gemini-3-flash-preview", "", create_payload, None)):
            res = persantil._handle_floating_gemini_chat({"message": "Yeni-2A ekle"}, [])

        self.assertTrue(res["ok"])
        self.assertIsNotNone(res.get("write_proposal"))
        self.assertEqual(res["write_proposal"]["mode"], "batch_action")
        self.assertFalse(res["write_proposal"]["is_destructive"])
        with self.connect() as c:
            self.assertIsNone(c.execute("SELECT id FROM classes WHERE name='Yeni-2A'").fetchone())

        # Confirm -> now DB mutates, AND Mevcut-1A is still preserved
        apply_res = persantil._handle_floating_gemini_apply_write({"token": res["write_proposal"]["token"]})
        self.assertTrue(apply_res["ok"])
        with self.connect() as c:
            self.assertIsNotNone(c.execute("SELECT id FROM classes WHERE name='Yeni-2A'").fetchone())
            self.assertIsNotNone(c.execute("SELECT id FROM classes WHERE name='Mevcut-1A'").fetchone())
            self.assertIsNotNone(c.execute("SELECT id FROM students WHERE full_name='Ece Yılmaz'").fetchone())

        # 2. Simulate chat delete_records -> MUST return destructive write_proposal and NOT delete yet
        del_payload = {"action": "delete_records", "classes": ["Silinecek-Bos"]}
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
             mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Silinecek-Bos silinecek. Onaylıyor musunuz?", "gemini-3-flash-preview", "", del_payload, None)):
            res_del = persantil._handle_floating_gemini_chat({"message": "Silinecek-Bos şubesini sil"}, [])

        self.assertTrue(res_del["write_proposal"]["is_destructive"])
        with self.connect() as c:
            self.assertIsNotNone(c.execute("SELECT id FROM classes WHERE name='Silinecek-Bos'").fetchone())

        apply_del = persantil._handle_floating_gemini_apply_write({"token": res_del["write_proposal"]["token"]})
        self.assertTrue(apply_del["ok"])
        with self.connect() as c:
            self.assertIsNone(c.execute("SELECT id FROM classes WHERE name='Silinecek-Bos'").fetchone())
            self.assertIsNotNone(c.execute("SELECT id FROM classes WHERE name='Mevcut-1A'").fetchone())
            self.assertIsNotNone(c.execute("SELECT id FROM classes WHERE name='Yeni-2A'").fetchone())

    def test_text_selection_enabled_in_desktop_and_css(self):
        desktop_src = (PROJECT_ROOT / "desktop_app.py").read_text(encoding="utf-8")
        self.assertIn("text_select=True", desktop_src)
        css = persantil.build_app_css(
            "'DM Sans', sans-serif", "", "15px", "1.4rem",
            persantil.THEME_PRESETS["Klinik Turkuaz (Varsayılan)"],
            "Klinik Turkuaz (Varsayılan)",
            "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)",
        )
        self.assertIn("user-select: text", css)

    def test_gemini_rename_classes_and_sql_update_coercion_with_merge_and_confirm(self):
        import gemini_assistant

        with self.connect() as c:
            cur_old = c.execute("INSERT INTO classes(name, teacher_name) VALUES ('Özel Eğitim Ana Sınıfı / A Şubesi', 'BETÜL CANSU')")
            old_cid = cur_old.lastrowid
            c.execute("INSERT INTO classes(name) VALUES ('A Şubesi')")  # pre-existing empty collision branch
            c.execute("INSERT INTO classes(name, teacher_name) VALUES ('Özel Eğitim Ana Sınıfı / B Şubesi', 'CENNET SÖNMEZ')")
            c.execute("INSERT INTO students(class_id, full_name, sex, person_type) VALUES (?, 'Ali Can', 'E', 'Öğrenci')", (old_cid,))
            c.commit()

        # Simulate Gemini mistakenly emitting UPDATE classes inside sql_export
        bad_sql_spec = {
            "action": "sql_export",
            "sql": "UPDATE classes SET name = REPLACE(name, 'Özel Eğitim Ana Sınıfı / ', '') WHERE name LIKE 'Özel Eğitim Ana Sınıfı / %'",
        }
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
             mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Sınıf adları sadeleştirilecek.", "gemini-3-flash-preview", "", bad_sql_spec, None)):
            res = persantil._handle_floating_gemini_chat({"message": "sınıf adlarından özel eğitim ana sınıfı / kısmını sil"}, [])

        self.assertTrue(res["ok"])
        self.assertNotIn("Güvenlik gereği", res["reply_text"])
        self.assertIsNotNone(res.get("write_proposal"))
        self.assertEqual(res["write_proposal"]["mode"], "batch_action")
        self.assertTrue(any("A Şubesi" in it for it in res["write_proposal"]["items"]))

        # Confirm -> merges A Şubesi (preserving Ali Can + BETÜL CANSU) and renames B Şubesi cleanly
        apply_res = persantil._handle_floating_gemini_apply_write({"token": res["write_proposal"]["token"]})
        self.assertTrue(apply_res["ok"])
        with self.connect() as c:
            cls_names = sorted([r["name"] for r in c.execute("SELECT name FROM classes").fetchall()])
            self.assertNotIn("Özel Eğitim Ana Sınıfı / A Şubesi", cls_names)
            self.assertNotIn("Özel Eğitim Ana Sınıfı / B Şubesi", cls_names)
            self.assertIn("A Şubesi", cls_names)
            self.assertIn("B Şubesi", cls_names)
            row_a = c.execute("SELECT id, teacher_name FROM classes WHERE name='A Şubesi'").fetchone()
            self.assertEqual(row_a["teacher_name"], "BETÜL CANSU")
            stu = c.execute("SELECT class_id FROM students WHERE full_name='Ali Can'").fetchone()
            self.assertEqual(stu["class_id"], row_a["id"])

    def test_topbar_ai_pill_toggles_ai_mode_and_redirects_without_key(self):
        import gemini_assistant
        import revir_ui

        persantil.save_app_settings({"gemini_enabled": "1", "gemini_floating_chat": "1"})
        persantil.clear_all_data_caches()

        # 1. With API key configured: clicking topbar AI pill toggles AI OFF then back ON
        with mock.patch.object(gemini_assistant, "get_api_key", return_value="AIzaFakeKey12345"):
            res_init = revir_ui.run_action_headless(persantil.main, {"action": "init"})
            self.assertIn("🤖 AI ·", res_init["topbar_html"])
            self.assertIn("window.__revirTopbarAction('toggle_ai', '')", res_init["topbar_html"])
            self.assertTrue(res_init["floating_chat_enabled"])

            # Click pill -> AI OFF
            res_off = revir_ui.run_action_headless(
                persantil.main,
                {"action": "topbar_action", "key": "toggle_ai", "value": ""},
            )
            self.assertIn("⚪ AI Kapalı", res_off["topbar_html"])
            self.assertFalse(res_off["floating_chat_enabled"])
            self.assertEqual(persantil.get_app_settings().get("gemini_enabled"), "0")
            self.assertFalse(revir_ui.session_state.get("cfg_gemini_enabled_toggle"))

            # Click pill again -> AI ON
            res_on = revir_ui.run_action_headless(
                persantil.main,
                {"action": "topbar_action", "key": "toggle_ai", "value": ""},
            )
            self.assertIn("🤖 AI ·", res_on["topbar_html"])
            self.assertTrue(res_on["floating_chat_enabled"])
            self.assertEqual(persantil.get_app_settings().get("gemini_enabled"), "1")
            self.assertTrue(revir_ui.session_state.get("cfg_gemini_enabled_toggle"))

        # 2. Without API key configured: clicking topbar AI pill navigates to Gemini AI settings with warning
        revir_ui.reset_session()
        persantil.save_app_settings({"gemini_enabled": "0", "gemini_api_key": ""})
        persantil.clear_all_data_caches()
        with mock.patch.object(gemini_assistant, "get_api_key", return_value=""):
            revir_ui.run_action_headless(persantil.main, {"action": "init"})
            res_nokey = revir_ui.run_action_headless(
                persantil.main,
                {"action": "topbar_action", "key": "toggle_ai", "value": ""},
            )
            self.assertEqual(res_nokey["active_page"], "⚙️ Ayarlar")
            self.assertIn("API anahtarı", res_nokey["main_html"])

    def test_github_release_update_check_download_and_installer_open(self):
        import io as _io
        import urllib.error
        import revir_ui

        # 1. Version parsing and comparison boundaries
        self.assertEqual(persantil._parse_version_tuple("RevirDefteri_Kurulum_v1.16.16.exe"), (1, 16, 16))
        self.assertEqual(persantil._parse_version_tuple("RevirDefteri_macOS_arm64_v1.17.2.dmg"), (1, 17, 2))
        self.assertEqual(persantil._compare_versions("v1.16.9", "1.16.16"), -1)
        self.assertEqual(persantil._compare_versions("1.16", "1.16.0"), 0)
        self.assertEqual(persantil._compare_versions("v2.0.0", "1.99.99"), 1)

        sample_releases = [
            {
                "tag_name": "v1.99.0",
                "name": "Okul Revir Defteri v1.99.0",
                "draft": False,
                "prerelease": False,
                "html_url": "https://github.com/blue-doggy/revir_defteri/releases/tag/v1.99.0",
                "published_at": "2026-11-22T11:00:00Z",
                "body": "- Yeni ozellikler\n- Hata duzeltmeleri",
                "assets": [
                    {
                        "name": "RevirDefteri_Kurulum_v1.99.0.exe",
                        "size": 115000000,
                        "browser_download_url": "https://github.com/blue-doggy/revir_defteri/releases/download/v1.99.0/RevirDefteri_Kurulum_v1.99.0.exe",
                        "content_type": "application/x-msdownload",
                        "updated_at": "2026-11-22T11:05:00Z",
                    }
                    # Note: .dmg not yet built for v1.99.0!
                ],
            },
            {
                "tag_name": "v1.98.0",
                "name": "Okul Revir Defteri v1.98.0",
                "draft": False,
                "prerelease": False,
                "html_url": "https://github.com/blue-doggy/revir_defteri/releases/tag/v1.98.0",
                "published_at": "2026-11-21T11:00:00Z",
                "body": "- macOS dmg dahil tam surum",
                "assets": [
                    {
                        "name": "RevirDefteri_Kurulum_v1.98.0.exe",
                        "size": 114000000,
                        "browser_download_url": "https://github.com/blue-doggy/revir_defteri/releases/download/v1.98.0/RevirDefteri_Kurulum_v1.98.0.exe",
                    },
                    {
                        "name": "RevirDefteri_macOS_arm64_v1.98.0.dmg",
                        "size": 148000000,
                        "browser_download_url": "https://github.com/blue-doggy/revir_defteri/releases/download/v1.98.0/RevirDefteri_macOS_arm64_v1.98.0.dmg",
                    },
                ],
            },
        ]

        class _FakeResp:
            def __init__(self, payload_bytes: bytes):
                self._buf = _io.BytesIO(payload_bytes)

            def read(self, n: int = -1) -> bytes:
                return self._buf.read(n)

            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

        import json as _json
        api_bytes = _json.dumps(sample_releases).encode("utf-8")

        # 2. Windows detects v1.99.0 .exe, macOS selects v1.98.0 .dmg (and notes v1.99.0 pending)
        with mock.patch("urllib.request.urlopen", side_effect=lambda req, timeout=12: _FakeResp(api_bytes)):
            win_res = persantil.check_for_github_release_updates(platform_override="win32", current_version="1.16.16")
        self.assertTrue(win_res["success"])
        self.assertTrue(win_res["has_updates"])
        self.assertEqual(win_res["latest_version"], "1.99.0")
        self.assertEqual(win_res["asset"]["name"], "RevirDefteri_Kurulum_v1.99.0.exe")
        self.assertEqual(win_res["asset"]["size_bytes"], 115000000)
        self.assertIn("MB", win_res["asset"]["size_str"])

        with mock.patch("urllib.request.urlopen", side_effect=lambda req, timeout=12: _FakeResp(api_bytes)):
            mac_res = persantil.check_for_github_release_updates(platform_override="darwin", current_version="1.16.16")
        self.assertTrue(mac_res["success"])
        self.assertTrue(mac_res["has_updates"])
        self.assertEqual(mac_res["latest_version"], "1.98.0")
        self.assertEqual(mac_res["asset"]["name"], "RevirDefteri_macOS_arm64_v1.98.0.dmg")
        self.assertIn("v1.99.0", mac_res["pending_build_note"])

        # Already up to date against macOS .dmg v1.98.0
        with mock.patch("urllib.request.urlopen", side_effect=lambda req, timeout=12: _FakeResp(api_bytes)):
            mac_cur = persantil.check_for_github_release_updates(platform_override="darwin", current_version="1.98.0")
        self.assertTrue(mac_cur["success"])
        self.assertFalse(mac_cur["has_updates"])

        # Network failure graceful handling
        with mock.patch("urllib.request.urlopen", side_effect=urllib.error.URLError("offline")):
            err_res = persantil.check_for_github_release_updates()
        self.assertFalse(err_res["success"])
        self.assertFalse(err_res["has_updates"])
        self.assertIn("bağlanılamadı", err_res["message"])

        # 3. Streaming installer download & launch helpers in tmp dir
        dl_dir = self.tmp_p / "downloads"
        fake_dmg = b"DMG_PAYLOAD_" * 500
        with mock.patch("urllib.request.urlopen", side_effect=lambda req, timeout=120: _FakeResp(fake_dmg)):
            ok_dl, dl_path, msg_dl = persantil.download_github_release_asset(
                "https://github.com/blue-doggy/revir_defteri/releases/download/v1.98.0/RevirDefteri_macOS_arm64_v1.98.0.dmg",
                "RevirDefteri_macOS_arm64_v1.98.0.dmg",
                expected_size=len(fake_dmg),
                target_dir=dl_dir,
            )
        self.assertTrue(ok_dl, msg_dl)
        self.assertIsNotNone(dl_path)
        self.assertEqual(dl_path.read_bytes(), fake_dmg)

        # Invalid URL rejected
        ok_bad, _, _ = persantil.download_github_release_asset("http://insecure/file.exe", "file.exe", target_dir=dl_dir)
        self.assertFalse(ok_bad)

        # Installer open & reveal commands (mocked Popen / os.startfile)
        import os as _os
        import subprocess as _sp
        with mock.patch.object(_sp, "Popen") as p_mock:
            ok_op, _ = persantil.open_downloaded_installer(dl_path, platform_override="darwin", reveal_only=False)
            self.assertTrue(ok_op)
            p_mock.assert_called_with(["open", str(dl_path)], close_fds=True)

            ok_rv, _ = persantil.open_downloaded_installer(dl_path, platform_override="darwin", reveal_only=True)
            self.assertTrue(ok_rv)
            p_mock.assert_called_with(["open", "-R", str(dl_path)], close_fds=True)

            # Windows launcher & reveal
            with mock.patch.object(_os, "startfile", create=True) as sf_mock:
                ok_win, _ = persantil.open_downloaded_installer(dl_path, platform_override="win32", reveal_only=False)
                self.assertTrue(ok_win)
                sf_mock.assert_called_with(str(dl_path))
            ok_win_rv, _ = persantil.open_downloaded_installer(dl_path, platform_override="win32", reveal_only=True)
            self.assertTrue(ok_win_rv)
            p_mock.assert_called_with(["explorer", f"/select,{dl_path}"], close_fds=True)

        # 4. Headless UI Settings -> Sürüm & Güncelleme displays asset info and download button
        revir_ui.reset_session()
        revir_ui.session_state["main_menu_page"] = "⚙️ Ayarlar"
        revir_ui.session_state["settings_vertical_nav"] = "💾 Veri, İçe Aktar & Sistem"
        revir_ui.session_state["last_update_check_result"] = mac_res
        out_ui = revir_ui.run_action_headless(persantil.main, {"action": "init"})
        html_ui = out_ui.get("main_html") or ""
        self.assertIn("RevirDefteri_macOS_arm64_v1.98.0.dmg", html_ui)
        self.assertIn(mac_res["asset"]["size_str"], html_ui)
        self.assertIn("Güncelleme Paketini İndir", html_ui)


if __name__ == "__main__":
    unittest.main(verbosity=2)
