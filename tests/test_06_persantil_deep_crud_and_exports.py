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


if __name__ == "__main__":
    unittest.main(verbosity=2)
