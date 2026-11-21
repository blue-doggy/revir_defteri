"""Unit tests for document_studio.py and gemini_assistant.py."""
from __future__ import annotations

import json
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import document_studio
import gemini_assistant


def _connect_row(db_path: Path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


class TestDocumentStudio(unittest.TestCase):
    def test_tr_fold_and_path_helpers(self):
        self.assertEqual(document_studio.tr_fold("Öğrenci Çağı ŞİğÜ"), "ogrenci cagi sigu")
        self.assertEqual(document_studio.tr_fold(""), "")
        self.assertEqual(document_studio.tr_fold(None), "")
        self.assertEqual(document_studio._safe_stem("Adı Soyadı / 2026!", "fallback"), "Adı Soyadı _ 2026!")
        self.assertEqual(document_studio._safe_stem("///", "fallback"), "___")
        self.assertEqual(document_studio._safe_stem("", "fallback"), "fallback")

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            p1 = document_studio._unique_target_path(d, "rapor", ".txt")
            p1.write_text("1", encoding="utf-8")
            p2 = document_studio._unique_target_path(d, "rapor", ".txt")
            self.assertNotEqual(p1, p2)
            self.assertTrue(p2.name.startswith("rapor_"))

        self.assertTrue(document_studio._rects_overlap_strongly((0, 0, 100, 20), (10, 0, 110, 20)))
        self.assertFalse(document_studio._rects_overlap_strongly((0, 0, 10, 10), (50, 50, 60, 60)))
        self.assertIsNotNone(document_studio.find_unicode_ttf_font())

    def test_html_shells_and_office_roundtrips(self):
        shell = document_studio.build_a4_html_shell("Başlık", "<p>İçerik ğüşiöç</p>")
        self.assertIn("Başlık", shell)
        txt_html = document_studio.render_text_a4_html(Path("not.txt"), "Satır 1\nSatır 2")
        self.assertIn("Satır 1", txt_html)

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            app_cfg = {"school_name": "Test İlkokulu", "nurse_name": "Hemşire Ayşe"}

            docx_p = document_studio.create_new_office_document_in_depot(d, "docx", "Veli İzin", app_cfg)
            self.assertTrue(docx_p.exists())
            paras, tables = document_studio._load_docx_structure(docx_p)
            self.assertIsInstance(paras, list)
            docx_out = d / "out.docx"
            document_studio._save_docx_structure(docx_p, docx_out, {0: "Güncel Başlık"}, ["Ek paragraf"], {})
            self.assertTrue(docx_out.exists())
            self.assertIn("Ek paragraf", document_studio.render_docx_a4_html(docx_out))

            xlsx_p = document_studio.create_new_office_document_in_depot(d, "xlsx", "Tarama Liste", app_cfg)
            self.assertTrue(xlsx_p.exists())
            import openpyxl
            wb = openpyxl.load_workbook(xlsx_p, read_only=True)
            cur_sheet = wb.sheetnames[0]
            wb.close()
            df = document_studio._load_xlsx_sheet_df(xlsx_p, cur_sheet)
            self.assertIsInstance(df, pd.DataFrame)
            xlsx_out = d / "out.xlsx"
            document_studio._save_xlsx_sheet_df(xlsx_p, xlsx_out, cur_sheet, df, "YeniSayfa")
            self.assertTrue(xlsx_out.exists())
            self.assertIn("<table", document_studio.render_xlsx_a4_html(xlsx_out, cur_sheet).lower())

            import fitz
            pdf_p = d / "form.pdf"
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 100), "Adı Soyadı: ........................")
            page.insert_text((50, 130), "Tarih: ____ / ____ / ______")
            doc.save(pdf_p)
            doc.close()

            fields = document_studio.scan_pdf_writable_fields(pdf_p, 0)
            self.assertIsInstance(fields, list)
            test_fields = fields + [{"id": "f1", "label": "Adı Soyadı", "x0": 120, "y0": 90, "x1": 300, "y1": 108, "value": ""}]
            matched_cnt = document_studio.auto_match_pdf_fields_locally(
                test_fields,
                app_cfg,
                {"full_name": "Ali Veli", "class_name": "4-A", "tc_no": "12345678901"},
                {"chronic_diseases": "Astım", "allergies": "Yok"},
            )
            self.assertGreaterEqual(matched_cnt, 1)
            self.assertTrue(any(f.get("value") == "Ali Veli" for f in test_fields))
            png = document_studio.render_pdf_page_preview_png(pdf_p, 0, test_fields, show_guides=True)
            self.assertTrue(png.startswith(b"\x89PNG"))
            filled_p = d / "filled.pdf"
            document_studio.save_filled_pdf(pdf_p, {0: test_fields}, filled_p)
            self.assertTrue(filled_p.exists())

    def test_rank_and_apply_clinical_record(self):
        from tests.db_test_helper import isolated_real_db
        with isolated_real_db() as (_tmp_p, _db_p, connect_fn):
            with connect_fn() as conn:
                conn.execute("INSERT INTO classes(name) VALUES ('3-B')")
                conn.execute("INSERT INTO students(class_id, full_name, student_no, birth_date, sex, person_type) VALUES (1, 'Zeynep Kaya', '405', '2017-01-10', 'K', 'Öğrenci')")
                conn.execute("INSERT INTO health_profiles(student_id, updated_at) VALUES (1, '2026-11-22')")
                conn.commit()

            self.assertEqual(document_studio._append_unique_csv("Astım", "Alerji"), "Astım, Alerji")
            self.assertEqual(document_studio._append_unique_csv("Astım", "astım"), "Astım")

            ranked = document_studio.rank_matching_people(connect_fn, {"full_name": "Zeynep Kaya", "tc_no": "11122233344"})
            self.assertEqual(ranked[0]["id"], 1)

            res = document_studio.apply_extracted_clinical_record(
                connect_fn,
                1,
                {
                    "document_type": "laboratuvar_sonucu",
                    "document_date": "2026-11-20",
                    "facility": "Devlet Hastanesi",
                    "summary": "Demir eksikliği",
                    "chronic_diseases": ["Anemi"],
                    "allergies": ["Polen"],
                    "blood_group": "A Rh+",
                    "lab_results": [{"test_name": "HGB", "value": "11.2", "unit": "g/dL", "ref_range": "11.5-15.5"}],
                    "measurements": {"weight_kg": 28.5, "height_cm": 130.0},
                    "medications": [{"med_name": "Ferro Sanol", "dosage_text": "1x1"}],
                },
            )
            self.assertIsInstance(res, list)
            self.assertGreaterEqual(len(res), 1)


class TestGeminiAssistant(unittest.TestCase):
    def test_key_and_anonymization_and_parsers(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            gemini_assistant.save_local_api_key("AIzaTestKey123456", d)
            self.assertEqual(gemini_assistant.get_api_key({}, d), "AIzaTestKey123456")
            self.assertTrue(gemini_assistant.is_ai_enabled({"gemini_enabled": "1"}, d))
            self.assertFalse(gemini_assistant.is_ai_enabled({"gemini_enabled": "0"}, d))

            anon = gemini_assistant.anonymize_free_text(
                "Ahmet Yılmaz TC 12345678901 Tel 0532 111 22 33",
                sensitive_names=["Ahmet Yılmaz"],
            )
            self.assertNotIn("12345678901", anon)
            self.assertNotIn("Ahmet Yılmaz", anon)

            dis = gemini_assistant.format_disability_profile_for_ai(
                {"chronic_diseases": "Tip 1 Diyabet", "disability_status": "Var"},
                {"full_name": "Ahmet Yılmaz"},
                sensitive_names=["Ahmet Yılmaz"],
            )
            self.assertNotIn("Ahmet Yılmaz", dis)

            ctx = gemini_assistant.build_anonymous_clinical_context(
                {"full_name": "Ahmet Yılmaz", "sex": "E"},
                {"chronic_diseases": "Astım"},
                {"weight_kg": 30, "height_cm": 135, "bmi": 16.5},
                [{"med_name": "Ventolin"}],
                age_months=120.0,
            )
            self.assertIn("Yaş", ctx)
            self.assertNotIn("Ahmet Yılmaz", ctx)

            parsed = gemini_assistant._parse_json_response('```json\n{"ok": true, "n": 2}\n```')
            self.assertEqual(parsed, {"ok": True, "n": 2})
            self.assertEqual(gemini_assistant._parse_json_response("not json"), {})
            self.assertIn("a", gemini_assistant._stringify_ai_field(["a", "b"]))

    def test_ai_wrappers_with_mocked_call_gemini(self):
        json_reply = json.dumps({
            "triage_level": "Orta",
            "summary": "Takip önerilir",
            "events": [{"event_date": "2026-11-25", "title": "Seminer", "category": "Etkinlik"}],
            "items": [{"name": "Çorba", "carbs_g": 20}],
            "classes": [{"name": "4-A", "students": ["Ali Veli"]}],
            "fields": [{"id": "f1", "value": "Test"}],
            "days": [{"date": "2026-11-22", "menus": {"Öğle": [{"name": "Pilav", "carbs_g": 30}]}}],
            "lab_results": [{"test_name": "HGB", "value": "12.0", "unit": "g/dL", "ref_range": "11-15"}],
            "hadith": "Temizlik imandandır.",
            "quote": "Sağlık en büyük hazinedir.",
            "sql": "SELECT COUNT(*) AS cnt FROM students",
            "answer": "1 öğrenci var.",
        }, ensure_ascii=False)

        with tempfile.TemporaryDirectory() as tmp:
            db_p = Path(tmp) / "nlq.db"
            conn = sqlite3.connect(db_p)
            conn.execute("CREATE TABLE students (id INTEGER PRIMARY KEY, full_name TEXT)")
            conn.execute("INSERT INTO students VALUES (1, 'Ali')")
            conn.commit()
            try:
                with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, json_reply, "gemini-2.5-flash")):
                    self.assertIsNotNone(gemini_assistant.test_gemini_connection("key", "gemini-2.5-flash"))
                    self.assertIsNotNone(gemini_assistant.analyze_visit_triage(
                        api_key="k", model="m", anon_context="ctx", complaint="Baş ağrısı",
                        vitals_summary="36.6", symptoms_draft="Yok", observation_draft="Normal", action_draft="İstirahat",
                    ))
                    self.assertIsNotNone(gemini_assistant.parse_lab_report_with_gemini(
                        api_key="k", model="m", file_bytes=b"HGB 12.0", file_name="lab.txt", anon_context="ctx",
                    ))
                    tmp_txt = Path(tmp) / "doc.txt"
                    tmp_txt.write_text("Örnek belge", encoding="utf-8")
                    self.assertIsNotNone(gemini_assistant.summarize_document_with_gemini(
                        api_key="k", model="m", file_path=tmp_txt, extracted_text="Metin",
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_nanda_care_plan(
                        api_key="k", model="m", anon_context="ctx", compiled_facts="Astım",
                    ))
                    self.assertIsNotNone(gemini_assistant.interpret_growth_trajectory(
                        api_key="k", model="m", anon_context="ctx",
                        measurements_df=pd.DataFrame([{"date": "2026-10-01", "bmi": 17}]),
                    ))
                    self.assertIsNotNone(gemini_assistant.check_medication_safety(
                        api_key="k", model="m", anon_context="ctx", medication_name="Parol",
                        route="Oral", dosage_pattern="2x1", dose_unit="500 mg", meal_relation="Tok",
                    ))
                    self.assertIsNotNone(gemini_assistant.analyze_outbreak_trends(
                        api_key="k", model="m", anon_visit_summary=[{"complaint": "Ateş", "count": 5}],
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_meb_activity_narrative(
                        api_key="k", model="m", stats={"visits": 10}, start_date_str="2026-10-01",
                        end_date_str="2026-10-31", school_name="Test Okulu",
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_student_report_ai_analysis(
                        api_key="k", model="m", anon_context="ctx", clinical_snapshot={}, missing_data_items=[],
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_screening_batch_ai_analysis(
                        api_key="k", model="m", cohort_summary={"total": 50}, missing_audit=[],
                    ))
                    self.assertIsNotNone(gemini_assistant.revise_report_ai_text(
                        api_key="k", model="m", current_text="Eski metin", instruction="Kısalt",
                    ))
                    self.assertIsNotNone(gemini_assistant.natural_language_db_query(
                        api_key="k", model="m", question="Kaç öğrenci var?", connection=conn,
                    ))
                    self.assertIsNotNone(gemini_assistant.chat_with_gemini_assistant(
                        api_key="k", model="m", user_message="Merhaba", history=[],
                    ))
                    self.assertIsNotNone(gemini_assistant.parse_monthly_school_menu_with_gemini(
                        api_key="k", model="m", year_month="2026-11", raw_text="1 Kasım Çorba",
                    ))
                    self.assertIsNotNone(gemini_assistant.estimate_extra_foods_carbs_with_gemini(
                        api_key="k", model="m", food_description="1 elma",
                    ))
                    self.assertIsNotNone(gemini_assistant.analyze_insulin_dose_with_gemini(
                        api_key="k", model="m", anon_student_context="ctx", growth_info={},
                        diabetes_profile={}, calc_result={}, meal_details={},
                    ))
                    self.assertIsNotNone(gemini_assistant.analyze_glycemic_trend_with_gemini(
                        api_key="k", model="m", anon_student_context="ctx", growth_info={},
                        tir_stats={}, logs_summary=[],
                    ))
                    self.assertIsNotNone(gemini_assistant.parse_classes_and_students_with_gemini(
                        api_key="k", model="m", instruction_text="4-A Ali Veli",
                    ))
                    self.assertIsNotNone(gemini_assistant.extract_clinical_record_from_document_with_gemini(
                        api_key="k", model="m", file_bytes=b"HGB 12", file_name="lab.txt", extracted_text="HGB 12",
                    ))
                    self.assertIsNotNone(gemini_assistant.infer_insights_from_student_reports_with_gemini(
                        api_key="k", model="m", anon_context="ctx", clinical_snapshot={}, missing_data_items=[],
                    ))
                    self.assertIsNotNone(gemini_assistant.scan_pdf_form_fields_with_gemini(
                        api_key="k", model="m", page_png_bytes=b"\x89PNG",
                    ))
                    self.assertIsNotNone(gemini_assistant.assist_office_document_with_gemini(
                        api_key="k", model="m", doc_type="docx", doc_payload={"paragraphs": ["İçerik"]}, instruction="Özetle",
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_calendar_events_with_ai(
                        api_key="k", model="m", year=2026, month=11,
                    ))
                    self.assertIsNotNone(gemini_assistant.generate_daily_leaf_content_with_ai(
                        api_key="k", model="m", target_date="2026-11-22",
                    ))
            finally:
                conn.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
