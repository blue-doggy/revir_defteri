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

    def test_apply_vaccinations_and_screenings_in_clinical_record(self):
        from tests.db_test_helper import isolated_real_db
        with isolated_real_db() as (_tmp_p, _db_p, connect_fn):
            with connect_fn() as conn:
                conn.execute("INSERT INTO classes(name) VALUES ('4-A')")
                conn.execute("INSERT INTO students(class_id, full_name, student_no, birth_date, sex, person_type) VALUES (1, 'Ece Yıldız', '202', '2016-05-12', 'K', 'Öğrenci')")
                conn.execute("INSERT INTO health_profiles(student_id, updated_at) VALUES (1, '2026-11-22')")
                conn.commit()

            moh_mock = [
                {"code": "kkk_2", "vaccine": "KKK (Kızamık-Kızamıkçık-Kabakulak)", "dose": "2. Doz", "period_label": "İlkokul 1", "period_months": 48},
            ]
            logs = document_studio.apply_extracted_clinical_record(
                connect_fn,
                1,
                {
                    "document_type": "Aşı ve Tarama",
                    "document_date": "2026-11-21",
                    "vaccinations": [
                        {"vaccine_name": "KKK", "dose_name": "2. Doz", "is_done": True, "administered_date": "2026-11-10", "facility": "ASM"}
                    ],
                    "school_screenings": {
                        "screening_date": "2026-11-21",
                        "height_cm": 136.5,
                        "weight_kg": 31.2,
                        "vision_right": "10/10",
                        "notes": "Sağlıklı tarama",
                    },
                },
                apply_profile=False,
                apply_labs=False,
                apply_vitals=False,
                apply_measurements=False,
                apply_meds=False,
                apply_vaccinations=True,
                apply_screenings=True,
                moh_schedule=moh_mock,
            )
            self.assertTrue(any("aşı" in x.lower() for x in logs))
            self.assertTrue(any("tarama" in x.lower() for x in logs))
            with connect_fn() as conn:
                vac = conn.execute("SELECT vaccine_code, is_done FROM student_vaccinations WHERE student_id=1").fetchone()
                scr = conn.execute("SELECT height_cm, vision_right FROM school_screenings WHERE student_id=1").fetchone()
            self.assertEqual(vac["vaccine_code"], "kkk_2")
            self.assertEqual(vac["is_done"], 1)
            self.assertAlmostEqual(scr["height_cm"], 136.5)
            self.assertEqual(scr["vision_right"], "10/10")


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

    def test_phase1_phase2_phase3_gemini_enhancements(self):
        # 1. Schema hint completeness & security
        for tbl in ("student_vaccinations", "seizure_logs", "fridge_temp_logs", "calendar_events", "calendar_daily_notes", "academic_years", "health_forms"):
            self.assertIn(tbl, gemini_assistant._DB_SCHEMA_HINT)
        self.assertNotIn("app_settings(", gemini_assistant._DB_SCHEMA_HINT)

        self.assertTrue(gemini_assistant.is_safe_readonly_sql("SELECT COUNT(*) FROM students"))
        self.assertFalse(gemini_assistant.is_safe_readonly_sql("DELETE FROM students"))
        self.assertFalse(gemini_assistant.is_safe_readonly_sql("SELECT * FROM students; DROP TABLE classes"))
        self.assertFalse(gemini_assistant.is_safe_readonly_sql("SELECT value FROM app_settings WHERE key='gemini_api_key'"))

        # 2. Strict KVKK PDF extraction masks PII & omits raw PDF inline binary when text layer exists
        import pymupdf
        pdf_doc = pymupdf.open()
        pg = pdf_doc.new_page()
        pg.insert_text((50, 80), "Okul Saglik Raporu: Ahmet Yilmaz TC 12345678901 Tel 0532 111 22 33 Astim izlemi mevcuttur.")
        pdf_bytes = pdf_doc.tobytes()
        pdf_doc.close()

        inline_strict, txt_strict = gemini_assistant._extract_attachment_context(
            "rapor.pdf", "application/pdf", pdf_bytes, sensitive_names=["Ahmet Yilmaz"], strict_kvkk_pdf=True
        )
        self.assertIsNone(inline_strict)
        self.assertIn("Katı KVKK", txt_strict)
        self.assertNotIn("12345678901", txt_strict)
        self.assertNotIn("Ahmet Yilmaz", txt_strict)

        inline_std, _txt_std = gemini_assistant._extract_attachment_context(
            "rapor.pdf", "application/pdf", pdf_bytes, sensitive_names=["Ahmet Yilmaz"], strict_kvkk_pdf=False
        )
        self.assertIsNotNone(inline_std)

        # 3. SQL 1-step self-repair & security block on repaired mutating SQL
        with tempfile.TemporaryDirectory() as tmp:
            db_p = Path(tmp) / "repair.db"
            conn = sqlite3.connect(db_p)
            conn.execute("CREATE TABLE students (id INTEGER PRIMARY KEY, full_name TEXT)")
            conn.execute("INSERT INTO students VALUES (1, 'Ayşe Demir')")
            conn.commit()
            try:
                with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, json.dumps({"sql": "SELECT COUNT(*) AS adet FROM students"}), "m")):
                    ok_r, df_r, err_r, sql_r = gemini_assistant.execute_safe_readonly_sql_with_repair(
                        api_key="k", model="m", sql_q="SELECT nonexistent_col FROM students", connection=conn, question="Öğrenci sayısı"
                    )
                self.assertTrue(ok_r, err_r)
                self.assertEqual(int(df_r.iloc[0]["adet"]), 1)
                self.assertEqual(sql_r, "SELECT COUNT(*) AS adet FROM students")

                with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, json.dumps({"sql": "DELETE FROM students"}), "m")):
                    ok_bad, _df_bad, err_bad, _ = gemini_assistant.execute_safe_readonly_sql_with_repair(
                        api_key="k", model="m", sql_q="SELECT nonexistent_col FROM students", connection=conn, question="Silme"
                    )
                self.assertFalse(ok_bad)
                self.assertIn("SQL çalıştırma hatası", err_bad)
            finally:
                conn.close()

        # 4. New Faz 2 & Faz 3 Gemini helpers with KVKK check + mocked call_gemini
        captured_prompts: list[str] = []

        def _fake_call_gemini(prompt, **kwargs):
            captured_prompts.append(prompt)
            payload = {
                "summary": "2 aşı eşleşti",
                "matched_vaccines": [{"vaccine_code": "hep_b_1", "is_done": True, "administered_date": "2018-01-01", "facility": "ASM"}],
                "guardian_letter": "Sayın Veli [ÖĞRENCİ_ADI] aşısı...",
                "asm_referral_note": "ASM yönlendirme notu",
                "pattern_analysis": "Nöbet süresi kısa ve stabil",
                "teacher_emergency_card": "Öğretmen Acil Kartı [ÖĞRENCİ_ADI]",
                "guardian_followup_note": "Veli özet notu",
                "risk_assessment": "Dolap 4.2°C stabil",
                "immediate_actions": "Rutin kontrol",
                "official_report_text": "Resmi soğuk zincir tutanağı",
                "priority_summary": "2 ürün kritik eşikte",
                "disposal_warning": "SKT geçmiş kalem yok",
                "procurement_petition": "Okul müdürlüğüne istem yazısı",
                "class_teacher_notice": "Sınıf hijyen duyurusu",
                "hygiene_checklist": "1. Havalandır",
                "risk_summary": "Tansiyon takibi önerilir",
                "workplace_health_tips": "Ergonomi önerileri",
                "physician_referral_note": "Aile hekimi kontrolü",
                "mode": "column_mapping",
                "explanation": "Sütunlar eşlendi",
                "column_map": {"class_name": "Sube", "first_name": "Adi", "last_name": "Soyadi", "height_cm": "Boy", "weight_kg": "Kilo"},
                "height_unit": "m",
                "weight_unit": "g",
                "rows": [{"class_name": "3-A", "full_name": "Can Kaya", "height_cm": 130.0, "weight_kg": 28.0}],
            }
            if not kwargs.get("expect_json"):
                return True, "Yönetici Brifingi Özeti", "m"
            return True, json.dumps(payload, ensure_ascii=False), "m"

        moh_sched = [{"code": "hep_b_1", "vaccine": "Hepatit B", "dose": "1. Doz", "period_label": "0. Ay"}]
        with mock.patch.object(gemini_assistant, "call_gemini", side_effect=_fake_call_gemini):
            ok1, d1, _ = gemini_assistant.parse_vaccination_card_with_gemini(
                api_key="k", model="m", file_bytes=b"Hepatit B 1. doz 2018-01-01", file_name="asi.txt", moh_schedule=moh_sched
            )
            self.assertTrue(ok1)
            self.assertEqual(d1["matched_vaccines"][0]["vaccine_code"], "hep_b_1")

            ok2, d2, _ = gemini_assistant.generate_vaccination_notice_with_gemini(
                api_key="k", model="m", anon_context="Yaş: 7 yaş", overdue_list=["KKK 2. Doz"], up_to_date_pct=85.0
            )
            self.assertTrue(ok2)
            self.assertIn("guardian_letter", d2)

            ok3, d3, _ = gemini_assistant.analyze_seizures_and_teacher_card_with_gemini(
                api_key="k", model="m", anon_context="Yaş: 9 yaş", seizure_protocol="5 dk aşarsa 112"
            )
            self.assertTrue(ok3)
            self.assertIn("teacher_emergency_card", d3)

            ok4, d4, _ = gemini_assistant.generate_cold_chain_breach_report_with_gemini(
                api_key="k", model="m", year_month="2026-11", month_logs_summary=[{"dolap_c": 4.2}], breach_logs=[]
            )
            self.assertTrue(ok4)
            self.assertIn("official_report_text", d4)

            ok5, d5, _ = gemini_assistant.analyze_inventory_and_procurement_with_gemini(
                api_key="k", model="m", alerts={"critical_stock": [], "expired": [], "expiring_soon": []}, total_items=5
            )
            self.assertTrue(ok5)
            self.assertIn("procurement_petition", d5)

            ok6, d6, _ = gemini_assistant.generate_isolation_communications_with_gemini(
                api_key="k", model="m", disease_name="Suçiçeği", class_name="2-A", detected_date="2026-11-20",
                notes_masked="TC 12345678901 Tel 0532 111 22 33 evde takip"
            )
            self.assertTrue(ok6)
            self.assertIn("hygiene_checklist", d6)
            self.assertNotIn("12345678901", captured_prompts[-1])

            ok7, d7, _ = gemini_assistant.analyze_adult_preventive_health_with_gemini(
                api_key="k", model="m", anon_context="Yetişkin 38 yaş", vitals_summary=[{"systolic_mmhg": 130}]
            )
            self.assertTrue(ok7)
            self.assertIn("risk_summary", d7)

            ok8, d8, _ = gemini_assistant.map_screening_table_with_gemini(
                api_key="k", model="m", headers=["Sube", "Adi", "Soyadi", "Boy", "Kilo"],
                sample_rows=[["2-A", "Ali", "TC 12345678901", "1.35", "32000"]]
            )
            self.assertTrue(ok8)
            self.assertEqual(d8["height_unit"], "m")
            self.assertNotIn("12345678901", captured_prompts[-1])

            ok9, d9, _ = gemini_assistant.generate_executive_briefing_with_gemini(
                api_key="k", model="m", briefing_scope="Okul Taraması", stats_or_cohort={"toplam": 100}
            )
            self.assertTrue(ok9)
            self.assertIn("Brifing", d9)

    def test_evren_platform_integration_and_settings_ui(self):
        import io
        import urllib.error
        import persantil
        import revir_ui
        from tests.db_test_helper import isolated_real_db

        captured_reqs = []

        class _FakeResp:
            def __init__(self, payload_obj):
                self._raw = json.dumps(payload_obj, ensure_ascii=False).encode("utf-8")

            def read(self):
                return self._raw

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

        def _fake_urlopen_ok(req, timeout=0):
            body = json.loads(req.data.decode("utf-8")) if getattr(req, "data", None) else {}
            captured_reqs.append((req.full_url, dict(req.headers), body, req.get_method()))
            return _FakeResp({
                "choices": [
                    {
                        "message": {
                            "content": "<think>iç düşünce</think>{\"ok\": true, \"yanit\": \"EVREN hazır\"}"
                        }
                    }
                ]
            })

        # 1. call_gemini auto-routes evren_llm_* keys to EVREN /chat/completions and strips <think>
        with mock.patch("urllib.request.urlopen", side_effect=_fake_urlopen_ok):
            ok_ev, txt_ev, used_ev = gemini_assistant.call_gemini(
                "Merhaba EVREN",
                api_key="evren_llm_test_key_123",
                model="gemini-3-flash-preview",
                expect_json=True,
                temperature=0.1,
            )
        self.assertTrue(ok_ev)
        self.assertTrue(used_ev.startswith("evren:"))
        self.assertNotIn("<think>", txt_ev)
        self.assertEqual(json.loads(txt_ev)["yanit"], "EVREN hazır")
        self.assertEqual(captured_reqs[0][0], "https://evren-llmapi.ssyz.org.tr/v1/chat/completions")
        hdrs0 = {k.lower(): v for k, v in captured_reqs[0][1].items()}
        self.assertEqual(hdrs0.get("authorization"), "Bearer evren_llm_test_key_123")
        self.assertEqual(hdrs0.get("x-api-key"), "evren_llm_test_key_123")
        self.assertEqual(captured_reqs[0][2].get("response_format"), {"type": "json_object"})

        # 2. HTTP 400 on response_format automatically retries once without response_format
        rf_attempts = []

        def _fake_urlopen_400_then_ok(req, timeout=0):
            body = json.loads(req.data.decode("utf-8"))
            rf_attempts.append(body)
            if len(rf_attempts) == 1:
                raise urllib.error.HTTPError(
                    req.full_url,
                    400,
                    "Bad Request",
                    hdrs=None,
                    fp=io.BytesIO(b'{"error":{"message":"response_format json_object unsupported"}}'),
                )
            return _FakeResp({"choices": [{"message": {"content": '{"plans":[]}'}}]})

        with mock.patch("urllib.request.urlopen", side_effect=_fake_urlopen_400_then_ok):
            ok_rf, txt_rf, _ = gemini_assistant.call_gemini(
                "Test JSON",
                api_key="evren_llm_test_key_123",
                model="glm-5.3",
                expect_json=True,
                provider="evren",
            )
        self.assertTrue(ok_rf)
        self.assertEqual(len(rf_attempts), 2)
        self.assertIn("response_format", rf_attempts[0])
        self.assertNotIn("response_format", rf_attempts[1])
        self.assertEqual(txt_rf, '{"plans":[]}')

        # 3. HTTP 403 terms_not_accepted stops immediately with clear Turkish guidance
        def _fake_urlopen_terms(req, timeout=0):
            raise urllib.error.HTTPError(
                req.full_url,
                403,
                "Forbidden",
                hdrs=None,
                fp=io.BytesIO(b'{"error":{"message":"terms_not_accepted"}}'),
            )

        with mock.patch("urllib.request.urlopen", side_effect=_fake_urlopen_terms) as m_terms:
            ok_tm_err, msg_tm_err, _ = gemini_assistant.call_gemini(
                "Merhaba",
                api_key="evren_llm_test_key_123",
                model="gemma-4-31b",
                provider="evren",
            )
        self.assertFalse(ok_tm_err)
        self.assertEqual(m_terms.call_count, 1)
        self.assertIn("terms_not_accepted", msg_tm_err)
        self.assertIn("EVREN Kullanım Şartlarını Onayla", msg_tm_err)

        # 4. Multimodal image with text-only requested model routes to EVREN vision model + data URI
        img_reqs = []

        def _fake_urlopen_vision(req, timeout=0):
            body = json.loads(req.data.decode("utf-8"))
            img_reqs.append(body)
            return _FakeResp({"choices": [{"message": {"content": "Görsel okundu"}}]})

        with mock.patch("urllib.request.urlopen", side_effect=_fake_urlopen_vision):
            ok_img, txt_img, used_img = gemini_assistant.call_gemini(
                "Bu fişi oku",
                api_key="evren_llm_test_key_123",
                model="deepseek-v4-flash",
                inline_files=[(b"\xff\xd8\xff\xe0testimg", "image/jpeg")],
                provider="evren",
            )
        self.assertTrue(ok_img)
        self.assertEqual(used_img, "evren:gemma-4-31b")
        self.assertEqual(txt_img, "Görsel okundu")
        user_content = img_reqs[0]["messages"][0]["content"]
        self.assertIsInstance(user_content, list)
        self.assertEqual(user_content[1]["type"], "image_url")
        self.assertTrue(user_content[1]["image_url"]["url"].startswith("data:image/jpeg;base64,"))

        # 5. list_evren_models & accept_evren_terms parse API endpoints properly
        def _fake_urlopen_mgmt(req, timeout=0):
            if req.full_url.endswith("/models"):
                return _FakeResp({"data": [{"id": "qwen3-embedding-8b"}, {"id": "deepseek-v4-flash"}, {"id": "gemma-4-31b"}]})
            if req.full_url.endswith("/terms/status"):
                return _FakeResp({"accepted": False, "current_version": 2})
            if req.full_url.endswith("/terms/text"):
                return _FakeResp({"version": 2, "text": "Kullanım şartları metni"})
            if req.full_url.endswith("/terms/accept"):
                return _FakeResp({"accepted": True, "version": 2})
            return _FakeResp({})

        with mock.patch("urllib.request.urlopen", side_effect=_fake_urlopen_mgmt):
            ok_ls, ids_ls, _ = gemini_assistant.list_evren_models("evren_llm_test_key_123")
            ok_ac, msg_ac = gemini_assistant.accept_evren_terms("evren_llm_test_key_123")
        self.assertTrue(ok_ls)
        self.assertEqual(ids_ls, ["gemma-4-31b", "deepseek-v4-flash", "qwen3-embedding-8b"])
        self.assertTrue(ok_ac)
        self.assertIn("v2", msg_ac)

        # 6. persantil._get_active_gemini_config + settings UI headless render for both Gemini & EVREN
        with isolated_real_db():
            persantil.save_app_settings({
                "ai_provider": "evren",
                "gemini_enabled": "1",
                "evren_api_key": "evren_llm_cfg_999",
                "evren_model": "qwen3-vl-30b",
                "evren_base_url": "https://evren-llmapi.ssyz.org.tr/v1",
            })
            persantil.clear_all_data_caches()
            is_on_ev, key_ev, mod_ev = persantil._get_active_gemini_config()
            self.assertTrue(is_on_ev)
            self.assertEqual(key_ev, "evren_llm_cfg_999")
            self.assertEqual(mod_ev, "qwen3-vl-30b")

            revir_ui.reset_session()
            out_ev_sync = revir_ui.run_action_headless(
                lambda: persantil._sync_floating_chat_ui(),
                {"action": "render"},
            )
            self.assertTrue(out_ev_sync["floating_chat_enabled"])
            self.assertEqual(out_ev_sync["floating_chat_title"], "🇹🇷 EVREN Klinik Asistan")
            self.assertIn("EVREN", out_ev_sync["floating_chat_label"])

            tb_ev_html = persantil._build_topbar_html("📊 Genel Bakış", None, False, True, "qwen3-vl-30b", has_ai_key=True)
            self.assertIn("🇹🇷 EVREN · qwen3-vl-30b", tb_ev_html)

            # Topbar toggle_ai must work cleanly with EVREN-only key without redirecting to Settings
            revir_ui.session_state.pop("target_page", None)
            try:
                persantil._handle_topbar_action("toggle_ai", "")
            except revir_ui._RerunException:
                pass
            self.assertIsNone(revir_ui.session_state.get("target_page"))
            self.assertEqual(persantil.get_app_settings().get("gemini_enabled"), "0")
            out_off_sync = revir_ui.run_action_headless(lambda: persantil._sync_floating_chat_ui(), {"action": "render"})
            self.assertFalse(out_off_sync["floating_chat_enabled"])

            try:
                persantil._handle_topbar_action("toggle_ai", "")
            except revir_ui._RerunException:
                pass
            self.assertEqual(persantil.get_app_settings().get("gemini_enabled"), "1")
            out_on_sync = revir_ui.run_action_headless(lambda: persantil._sync_floating_chat_ui(), {"action": "render"})
            self.assertTrue(out_on_sync["floating_chat_enabled"])

            revir_ui.reset_session()
            out_ev_ui = revir_ui.run_action_headless(
                lambda: persantil.render_gemini_ai_settings_panel(persantil.get_app_settings()),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_ev_ui["main_html"])
            self.assertIn("cfg_evren_api_key_inp", out_ev_ui["main_html"])
            self.assertIn("btn_accept_evren_terms", out_ev_ui["main_html"])
            self.assertIn("evren.ssyz.org.tr", out_ev_ui["main_html"])
            self.assertIn("__revirOpenChatBubble", out_ev_ui["main_html"])
            self.assertIn("Sohbet Penceresini Şimdi Aç", out_ev_ui["main_html"])

            out_gm_switch = revir_ui.run_action_headless(
                lambda: persantil.render_gemini_ai_settings_panel(persantil.get_app_settings()),
                {"action": "change", "key": "cfg_ai_provider_sel", "value": 0},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_gm_switch["main_html"])
            self.assertEqual(persantil.get_app_settings().get("ai_provider"), "gemini")

            revir_ui.reset_session()
            cfg_gm = persantil.get_app_settings()
            out_gm_ui = revir_ui.run_action_headless(
                lambda: persantil.render_gemini_ai_settings_panel(cfg_gm),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_gm_ui["main_html"])
            self.assertIn("cfg_gemini_api_key_inp", out_gm_ui["main_html"])
            self.assertIn("btn_save_gemini_cfg", out_gm_ui["main_html"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
