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


if __name__ == "__main__":
    unittest.main(verbosity=2)
