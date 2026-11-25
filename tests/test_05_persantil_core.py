"""Comprehensive unit & integration tests for persantil.py (data, math, DB CRUD, exports, and headless UI)."""
from __future__ import annotations

import sys
import json
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


class TestPersantilCoreAndUI(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        revir_ui.reset_session()

    def tearDown(self):
        self.ctx.__exit__(None, None, None)

    def test_01_settings_icons_assets_and_templates(self):
        self.assertTrue(len(persantil._load_app_version()) > 0)
        self.assertIsNotNone(persantil.get_app_icon())
        self.assertIsInstance(persantil.is_dark_mode_active({"dark_mode_adaptation": "Açık Mod (Her Zaman Aydınlık)"}), bool)
        theme = persantil.THEME_PRESETS["Klinik Turkuaz (Varsayılan)"]
        css = persantil.build_app_css(
            "'DM Sans', sans-serif",
            "",
            "15px",
            "1.4rem",
            theme,
            "Klinik Turkuaz (Varsayılan)",
            "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)",
        )
        self.assertIn("<style", css)

        cfg = persantil.get_app_settings()
        self.assertIn("graduation_grade_level", cfg)
        persantil.save_app_settings({"school_name": "Deneme İlkokulu ĞÜŞİÖÇ", "graduation_grade_level": "8"})
        cfg2 = persantil.get_app_settings()
        self.assertEqual(cfg2["school_name"], "Deneme İlkokulu ĞÜŞİÖÇ")
        self.assertEqual(cfg2["graduation_grade_level"], "8")

        tpl = persantil.get_care_plan_templates()
        self.assertIsInstance(tpl, dict)
        persantil.save_care_plan_templates(tpl)
        persantil.reset_care_plan_templates()
        sec_tpl = persantil.get_care_plan_section_templates("Astım")
        self.assertIsInstance(sec_tpl, dict)

        provs = persantil.get_tr_provinces()
        self.assertIsInstance(provs, list)

    def test_widget_layout_profile_persistence_and_reset(self):
        default = {
            "stats": {"id": "stats", "x_pct": 0, "y_px": 0, "w_pct": 100, "h_px": 120, "visible": True, "z": 1, "accent": "#0d9488"},
        }
        saved = {
            "stats": {"id": "stats", "x_pct": 12.5, "y_px": 40, "w_pct": 60, "h_px": 240, "visible": False, "z": 4, "accent": "#2563eb"},
        }

        persantil.save_widget_layout("layout_test", saved, "Kişisel Düzen")
        self.assertEqual(persantil.get_saved_widget_layout("layout_test", default), saved)
        bundle = json.loads(persantil.get_app_settings()["widget_layout:layout_test"])
        self.assertEqual(bundle["active_preset"], "Kişisel Düzen")
        self.assertEqual(bundle["presets"]["Kişisel Düzen"], saved)

        persantil.reset_widget_layout("layout_test")
        self.assertEqual(persantil.get_saved_widget_layout("layout_test", default), default)
        with self.assertRaises(ValueError):
            persantil._layout_json_from_sync("{invalid")

    def test_dossier_layout_renders_widget_markers_and_sync_engine(self):
        with self.connect() as connection:
            class_id = connection.execute("INSERT INTO classes(name) VALUES ('Düzen Testi')").lastrowid
            person_id = connection.execute(
                "INSERT INTO students(class_id, full_name, sex, birth_date, person_type) VALUES (?, 'Yerleşim Testi', 'K', '2018-01-01', 'Öğrenci')",
                (class_id,),
            ).lastrowid
        person = persantil.get_person_by_id(person_id)
        with self.connect() as connection:
            health_profile = connection.execute("SELECT * FROM health_profiles WHERE student_id=?", (person_id,)).fetchone()
            visit_stats = connection.execute("SELECT COUNT(*) cnt, MAX(visit_date) last_date FROM visits WHERE student_id=?", (person_id,)).fetchone()

        output = revir_ui.run_action_headless(
            lambda: persantil._render_dossier_tab_overview(
                person,
                person_id,
                health_profile,
                visit_stats,
                persantil.THEME_PRESETS["Klinik Turkuaz (Varsayılan)"],
                "#475569",
            ),
            {"action": "render"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", output["main_html"])
        self.assertIn("data-layout-widget-marker=\"vision\"", output["main_html"])
        self.assertIn("layout_dossier_cards", output["scripts_html"])

    def test_report_settings_have_one_entry_point(self):
        output = revir_ui.run_action_headless(persantil.render_reports, {"action": "render"})
        main_html = str(output.get("main_html") or "")
        self.assertIn("Rapor Ayarları", main_html)
        self.assertIn("Tarama ve Persantil Şablonları", main_html)
        self.assertIn("Kişi Dosyası Rapor Ayarları", main_html)
        self.assertNotIn("Kişi Raporu Ayarları", main_html)

    def test_screening_report_wizard_gates_preview_and_export(self):
        sample_df = pd.DataFrame([{"Ad Soyad": "Örnek Öğrenci", "Durum": "Normal"}])
        sample_meta = {"risk_summary": [], "selected_classes": [], "selected_dates": ["2026-11-22"]}
        with (
            mock.patch.object(
                persantil,
                "get_available_screening_and_measurement_dates",
                return_value=[{"date": "2026-11-22", "label": "22.11.2026"}],
            ),
            mock.patch.object(
                persantil,
                "build_custom_screening_report_data",
                return_value=(sample_df, sample_meta),
            ) as build_report,
            mock.patch.object(persantil, "export_custom_screening_report_docx", return_value=b"docx") as export_docx,
            mock.patch.object(persantil, "export_custom_screening_report_excel", return_value=b"xlsx") as export_xlsx,
            mock.patch.object(persantil, "export_custom_screening_report_html", return_value="<html></html>") as export_html,
        ):
            scope_output = revir_ui.run_action_headless(
                persantil.render_dynamic_screening_reports_tab,
                {"action": "render"},
            )
            self.assertIn("Adım 1/3", scope_output["main_html"])
            self.assertIn("Önizlemeye Geç", scope_output["main_html"])
            build_report.assert_not_called()
            export_docx.assert_not_called()
            export_xlsx.assert_not_called()
            export_html.assert_not_called()

            preview_output = revir_ui.run_action_headless(
                persantil.render_dynamic_screening_reports_tab,
                {"action": "click", "key": "dyn_rep_wizard_next_preview"},
            )
            self.assertIn("Adım 2/3", preview_output["main_html"])
            self.assertIn("Dinamik Rapor Önizlemesi", preview_output["main_html"])
            build_report.assert_called_once()
            self.assertEqual(build_report.call_args.kwargs["selected_dates"], ["2026-11-22"])
            export_docx.assert_not_called()
            export_xlsx.assert_not_called()
            export_html.assert_not_called()

            export_output = revir_ui.run_action_headless(
                persantil.render_dynamic_screening_reports_tab,
                {"action": "click", "key": "dyn_rep_wizard_next_export"},
            )
            self.assertIn("Adım 3/3", export_output["main_html"])
            self.assertIn("Son Kontrol", export_output["main_html"])
            self.assertIn("Rapor şablonu", export_output["main_html"])
            self.assertIn("Word (.docx) İndir", export_output["main_html"])
            self.assertEqual(export_docx.call_count, 1)
            self.assertEqual(export_xlsx.call_count, 1)
            self.assertEqual(export_html.call_count, 1)

    def test_meb_report_wizard_defers_word_generation_until_export(self):
        stats = {
            "students_total": 12,
            "screened_cohort_cnt": 9,
            "bmi_eval_total": 8,
            "vax_students_cnt": 7,
            "fridge_logs_cnt": 2,
            "special_needs_cnt": 1,
            "care_plans_total": 1,
            "visits_total": 14,
            "referral_cnt": 2,
        }
        with (
            mock.patch.object(
                persantil,
                "get_active_academic_year",
                return_value={"start_date": "2026-09-01", "end_date": "2027-06-30"},
            ),
            mock.patch.object(persantil, "collect_meb_annual_activity_stats", return_value=stats),
            mock.patch.object(persantil, "_get_active_gemini_config", return_value=(False, "", "test-model")),
            mock.patch.object(persantil, "build_meb_annual_activity_report_docx", return_value=b"docx") as build_docx,
        ):
            scope_output = revir_ui.run_action_headless(
                persantil._render_meb_activity_report_tab,
                {"action": "render"},
            )
            self.assertIn("Adım 1/3", scope_output["main_html"])
            self.assertIn("MEB rapor önizlemesine geç", scope_output["main_html"])
            build_docx.assert_not_called()

            preview_output = revir_ui.run_action_headless(
                persantil._render_meb_activity_report_tab,
                {"action": "click", "key": "meb_rep_wizard_next_preview"},
            )
            self.assertIn("Adım 2/3", preview_output["main_html"])
            self.assertIn("Kayıtlı Öğrenci", preview_output["main_html"])
            build_docx.assert_not_called()

            export_output = revir_ui.run_action_headless(
                persantil._render_meb_activity_report_tab,
                {"action": "click", "key": "meb_rep_wizard_next_export"},
            )
            self.assertIn("Adım 3/3", export_output["main_html"])
            self.assertIn("Son Kontrol", export_output["main_html"])
            self.assertIn("Rapor dönemi", export_output["main_html"])
            self.assertIn("MEB Dönem / Yıl Sonu Resmi Faaliyet Raporunu İndir", export_output["main_html"])
            build_docx.assert_called_once()

    def test_classic_report_wizard_gates_person_and_class_exports(self):
        with self.connect() as connection:
            class_id = connection.execute("INSERT INTO classes(name) VALUES ('Wizard 1-A')").lastrowid
            person_id = connection.execute(
                "INSERT INTO students(class_id, full_name, sex, birth_date, person_type) VALUES (?, 'Wizard Öğrenci', 'K', '2016-02-03', 'Öğrenci')",
                (class_id,),
            ).lastrowid
        report_cfg = persantil.get_report_settings()

        with (
            mock.patch.object(persantil, "build_student_report_section_html", return_value="<section></section>"),
            mock.patch.object(persantil, "build_class_batch_report_docx", return_value=b"class-docx") as build_class_docx,
            mock.patch.object(persantil, "build_student_report_docx", return_value=b"person-docx") as build_person_docx,
            mock.patch.object(persantil, "_render_student_ai_report_editor", return_value=(False, "")),
        ):
            revir_ui.session_state["report_scope_mode"] = "🏫 Tüm Sınıf Raporu (Toplu)"
            class_scope = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "render"},
            )
            self.assertIn("Adım 1/3", class_scope["main_html"])
            self.assertIn("Sınıf raporu önizlemesine geç", class_scope["main_html"])
            build_class_docx.assert_not_called()

            class_preview = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "click", "key": "classic_rep_wizard_next_preview"},
            )
            self.assertIn("Adım 2/3", class_preview["main_html"])
            self.assertIn("Yazdırılabilir Sınıf Raporu Önizleme", class_preview["main_html"])
            self.assertNotIn("Tüm Sınıfın Word Raporunu İndir", class_preview["main_html"])
            build_class_docx.assert_not_called()

            class_export = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "click", "key": "classic_rep_wizard_next_export_class"},
            )
            self.assertIn("Adım 3/3", class_export["main_html"])
            self.assertIn("Son Kontrol", class_export["main_html"])
            self.assertIn("Öğrenci sayısı", class_export["main_html"])
            self.assertIn("Tüm Sınıfın Word Raporunu İndir", class_export["main_html"])
            build_class_docx.assert_not_called()

            revir_ui.reset_session()
            revir_ui.session_state["active_person_id"] = person_id
            person_scope = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "render"},
            )
            self.assertIn("Kişi raporu önizlemesine geç", person_scope["main_html"])

            person_preview = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "click", "key": "classic_rep_wizard_next_preview"},
            )
            self.assertIn("Adım 2/3", person_preview["main_html"])
            self.assertIn("Yazdırılabilir Kişi Raporu Önizleme", person_preview["main_html"])
            self.assertNotIn("Kişi Sağlık Raporunu İndir", person_preview["main_html"])
            build_person_docx.assert_not_called()

            person_export = revir_ui.run_action_headless(
                lambda: persantil._render_classic_reports_content(["Wizard 1-A"], report_cfg),
                {"action": "click", "key": "classic_rep_wizard_next_export_person"},
            )
            self.assertIn("Adım 3/3", person_export["main_html"])
            self.assertIn("Son Kontrol", person_export["main_html"])
            self.assertIn("Rapor türü", person_export["main_html"])
            self.assertIn("Kişi Sağlık Raporunu İndir", person_export["main_html"])
            build_person_docx.assert_not_called()

    def test_dashboard_distinguishes_failed_checks_from_clear_status(self):
        failures = (
            mock.patch.object(persantil, "get_scheduled_doses_for_date", side_effect=RuntimeError("test")),
            mock.patch.object(persantil, "get_pending_followup_visits", side_effect=RuntimeError("test")),
            mock.patch.object(persantil, "get_infirmary_inventory_alerts", side_effect=RuntimeError("test")),
            mock.patch.object(persantil, "get_fridge_temp_logs", side_effect=RuntimeError("test")),
        )
        for failure in failures:
            failure.start()
        try:
            output = revir_ui.run_action_headless(persantil.render_overview, {"action": "render"})
        finally:
            for failure in reversed(failures):
                failure.stop()
        main_html = str(output.get("main_html") or "")
        self.assertIn("Bazı sağlık kontrollerinin durumu alınamadı", main_html)
        self.assertNotIn("Bekleyen ilaç dozu, açık sevk/izlem veya kritik stok uyarısı bulunmuyor", main_html)

    def test_navigation_pages_all_have_renderers(self):
        unmapped_pages = [
            page.split(" ", 1)[-1]
            for page in persantil.PAGES
            if page.split(" ", 1)[-1] not in persantil.PAGE_RENDERERS
        ]
        self.assertEqual(unmapped_pages, [])

    def test_02_academic_years_and_backups(self):
        years = persantil.get_all_academic_years()
        self.assertGreaterEqual(len(years), 1)
        active = persantil.get_active_academic_year()
        self.assertIsNotNone(active)

        # Seed classes & students into master DB for promote test
        with self.connect() as c:
            for cname in ["Anasınıfı", "1-A", "3-A", "4-A", "5-A"]:
                cur = c.execute("INSERT INTO classes(name) VALUES (?)", (cname,))
                cid = cur.lastrowid
                c.execute(
                    "INSERT INTO students(class_id, full_name, sex, birth_date, person_type) VALUES (?, ?, 'E', '2016-05-01', 'Öğrenci')",
                    (cid, f"Öğrenci {cname}"),
                )
            c.commit()

        y_id = persantil.create_academic_year("2028-2029", "2028-09-01", "2029-06-15", student_policy="promote", graduation_grade=4)
        self.assertGreater(y_id, 0)
        persantil.update_academic_year(y_id, "2028-2029 Güncel", "2028-09-05", "2029-06-20")
        persantil.set_active_academic_year(y_id)
        y_db = self.tmp_p / "egitim_yillari" / "2028_2029" / "revir_defteri.db"
        import contextlib
        import sqlite3
        with contextlib.closing(sqlite3.connect(y_db)) as yc:
            names = sorted([r[0] for r in yc.execute("SELECT name FROM classes").fetchall()])
            self.assertEqual(names, ["2-A", "4-A", "Anasınıfı"])

        # Test backups
        self.assertIn("KB", persantil.format_file_size(2048))
        b_res = persantil.run_daily_auto_backup(force=True)
        self.assertIsNotNone(b_res)
        blist = persantil.list_auto_backups()
        self.assertGreaterEqual(len(blist), 1)

        # Switch back & delete test year
        persantil.set_active_academic_year(years[0]["id"])
        persantil.delete_academic_year(y_id)

    def test_03_who_math_and_growth_evaluation(self):
        age_m = persantil.age_in_months("2016-05-01", date(2026, 5, 1))
        self.assertAlmostEqual(age_m, 120.0, delta=0.5)
        self.assertIsNone(persantil.age_in_months("invalid-date", date(2026, 5, 1)))

        # Test WHO LMS / SDS / Percentile helpers if present
        p50 = persantil.percentile(0.0)
        self.assertAlmostEqual(p50, 50.0, delta=0.5)
        self.assertLess(persantil.percentile(-3.0), 1.0)
        self.assertGreater(persantil.percentile(3.0), 99.0)
        self.assertIsInstance(persantil.get_percentile_band(50.0), str)

        ev = persantil.evaluate_measurement(138.0, "Boy (cm)", "E", 120)
        self.assertTrue(ev is None or isinstance(ev, dict))
        curves = persantil.generate_who_percentile_curves("Boy (cm)", "E", 0, 72)
        self.assertIsInstance(curves, pd.DataFrame)

    def test_04_kvkk_and_person_crud_and_clinical_records(self):
        revir_ui.session_state["kvkk_privacy_mode"] = True
        masked = persantil.mask_kvkk_value("Ahmet Yılmaz", "name")
        self.assertNotEqual(masked, "Ahmet Yılmaz")
        self.assertEqual(persantil.mask_kvkk_value("", "name"), "")

        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('2-B')")
            c.execute(
                "INSERT INTO students(class_id, full_name, student_no, sex, birth_date, person_type) VALUES (1, 'Elif Demir', '102', 'K', '2018-03-15', 'Öğrenci')"
            )
            c.commit()

        p = persantil.get_person_by_id(1)
        self.assertEqual(p["full_name"], "Elif Demir")
        persantil.set_active_person(1)
        latest = persantil.get_person_latest_measurements(1)
        self.assertIsInstance(latest, dict)
        dis = persantil.get_person_disability_profile(1, person=p)
        self.assertIsInstance(dis, dict)
        self.assertIsInstance(persantil.format_disability_badge_text(dis), str)

    def test_05_headless_render_all_major_pages(self):
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('1-A')")
            c.execute(
                "INSERT INTO students(class_id, full_name, student_no, sex, birth_date, person_type) VALUES (1, 'Can Öztürk', '201', 'E', '2019-09-01', 'Öğrenci')"
            )
            c.commit()
        persantil.set_active_person(1)

        pages = [
            persantil.render_overview,
            persantil.render_students,
            persantil.render_civilians,
            persantil.render_health_profile_and_plan,
            persantil.render_visit,
            persantil.render_medication_tracker,
            persantil.render_vital_tracking,
            persantil.render_measurements,
            persantil.render_school_screening,
            persantil.render_reports,
            persantil.render_academic_year_settings,
            persantil.render_documents_depot,
            persantil.render_settings,
            persantil.render_insulin_diabetes,
        ]
        for page_fn in pages:
            with self.subTest(page=page_fn.__name__):
                revir_ui.reset_session()
                revir_ui.session_state["active_person_id"] = 1
                with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
                    "İmsak": "06:00", "Güneş": "07:30", "Öğle": "12:50", "İkindi": "15:20", "Akşam": "17:45", "Yatsı": "19:05"
                }):
                    out = revir_ui.run_action_headless(page_fn, {"action": "render"})
                self.assertIn("main_html", out)
                self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"], f"{page_fn.__name__} raised runtime error!")

    def test_06_students_table_columns_and_name_click(self):
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('4-A')")
            c.execute(
                "INSERT INTO students(class_id, full_name, student_no, sex, birth_date, person_type) VALUES (1, 'Ayşe Yılmaz', '305', 'K', '2016-11-22', 'Öğrenci')"
            )
            c.commit()
        revir_ui.reset_session()
        out = revir_ui.run_action_headless(persantil.render_students, {"action": "render"})
        html = out["main_html"]
        self.assertIn("Şube Adı", html)
        self.assertIn("Okul No", html)
        self.assertIn("Ad Soyad", html)
        self.assertIn("Cinsiyet", html)
        self.assertIn("10 yıl", html)
        for hidden_col in ("Boy (cm)", "Kilo (kg)", "Doğum Tarihi", "VKİ (BMI)", "📂 Dosya"):
            self.assertNotIn(hidden_col, html)

        # Simulate clicking the "Ad Soyad" cell on row 0 -> opens dossier dialog
        out_click = revir_ui.run_action_headless(
            persantil.render_students,
            {"action": "df_select", "key": "profile_roster_table_0", "value": {"rows": [0], "cells": [[0, "Ad Soyad"]]}},
        )
        self.assertIsNotNone(out_click.get("dialog_html"))
        self.assertIn("Ayşe Yılmaz", out_click.get("dialog_html") or "")

    def test_07_screening_word_docx_import(self):
        import io
        import docx

        # 1. Word (.docx) with class heading + screening table
        d1 = docx.Document()
        d1.add_heading("3-A Şubesi", level=2)
        tbl = d1.add_table(rows=3, cols=5)
        headers = ["Ad Soyad", "Cinsiyet", "Doğum Tarihi", "Boy (cm)", "Kilo (kg)"]
        for idx, h in enumerate(headers):
            tbl.cell(0, idx).text = h
        row1 = ["Mehmet Öztürk", "E", "12.05.2017", "134.5", "30.2"]
        row2 = ["Elif Şahin", "K", "2017-08-20", "131", "28.4"]
        for c_i, val in enumerate(row1):
            tbl.cell(1, c_i).text = val
        for c_i, val in enumerate(row2):
            tbl.cell(2, c_i).text = val
        buf1 = io.BytesIO()
        d1.save(buf1)

        recs = persantil.parse_screening_file_or_image(buf1.getvalue(), "tarama_3a.docx")
        self.assertEqual(len(recs), 2)
        self.assertEqual(recs[0]["class_name"], "3-A Şubesi")
        self.assertEqual(recs[0]["full_name"], "Mehmet Öztürk")
        self.assertEqual(recs[0]["sex"], "Erkek")
        self.assertEqual(recs[0]["birth_date"], "2017-05-12")
        self.assertAlmostEqual(recs[0]["height_cm"], 134.5, places=1)
        self.assertAlmostEqual(recs[0]["weight_kg"], 30.2, places=1)

        # 2. Word (.docx) with delimited paragraph lines (no table)
        d2 = docx.Document()
        d2.add_paragraph("Sınıf;Ad Soyad;Cinsiyet;Doğum Tarihi;Boy (cm);Kilo (kg)")
        d2.add_paragraph("2-B;Zeynep Kaya;K;2018-03-10;126;25.5")
        buf2 = io.BytesIO()
        d2.save(buf2)
        recs2 = persantil.parse_screening_file_or_image(buf2.getvalue(), "satirlar.docx")
        self.assertEqual(len(recs2), 1)
        self.assertEqual(recs2[0]["class_name"], "2-B")
        self.assertEqual(recs2[0]["full_name"], "Zeynep Kaya")

        # 3. Commit imported Word rows & verify screening UI uploader accepts .docx
        upd, scr, add = persantil.commit_screening_import(recs, screening_date="2026-11-22", default_class="3-A Şubesi")
        self.assertEqual(scr, 2)
        self.assertEqual(add, 2)
        out_scr = revir_ui.run_action_headless(persantil.render_screening_import_section, {"action": "render"})
        self.assertIn(".docx", out_scr["main_html"])

    def test_08_academic_year_breaks_ai_fill_and_single_adult_emoji(self):
        import school_calendar

        # 1. Verify single man-woman emoji (👫) is used for Yetişkinler and no double ZWJ (🧑‍💼) remains
        self.assertIn("👫 Yetişkinler", persantil.PAGES)
        self.assertNotIn("🧑‍💼 Yetişkinler", persantil.PAGES)

        # 2. Render render_academic_year_settings and click AI fill button (btn_ai_fill_ay_calendar)
        revir_ui.reset_session()
        out_ay = revir_ui.run_action_headless(persantil.render_academic_year_settings, {"action": "render"})
        self.assertIn("Ara Tatiller ve Yarıyıl (Sömestr / 15 Tatil) Tarihleri", out_ay["main_html"])
        self.assertIn("Yapay Zekâ ile Bilgileri Otomatik Doldur", out_ay["main_html"])

        out_ai_fill = revir_ui.run_action_headless(
            persantil.render_academic_year_settings,
            {"action": "click", "key": "btn_ai_fill_ay_calendar"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", out_ai_fill["main_html"])

        with self.connect() as c:
            cnt_all = c.execute(
                "SELECT COUNT(*) FROM calendar_events WHERE source = 'meb_akademik_takvim'"
            ).fetchone()[0]
            cnt_b1 = c.execute(
                "SELECT COUNT(*) FROM calendar_events WHERE source = 'meb_akademik_takvim' AND title LIKE '%1. Dönem Ara Tatili%'"
            ).fetchone()[0]
            cnt_sem = c.execute(
                "SELECT COUNT(*) FROM calendar_events WHERE source = 'meb_akademik_takvim' AND title LIKE '%Yarıyıl Tatili / Sömestr%'"
            ).fetchone()[0]
            cnt_b2 = c.execute(
                "SELECT COUNT(*) FROM calendar_events WHERE source = 'meb_akademik_takvim' AND title LIKE '%2. Dönem Ara Tatili%'"
            ).fetchone()[0]

        self.assertGreaterEqual(cnt_b1, 5)
        self.assertGreaterEqual(cnt_sem, 10)
        self.assertGreaterEqual(cnt_b2, 5)
        self.assertGreaterEqual(cnt_all, 22)

    def test_09_consolidated_settings_and_box_layout_excludes_calendar(self):
        expected_groups = [
            "🏫 Kurum & Şube",
            "🎨 Görünüm & Düzen",
            "📅 Takvim & Yemekhane",
            "🎛️ Görünürlük & Rapor",
            "🩺 Klinik & WHO",
            "🤖 Yapay Zekâ & KVKK",
            "💾 Veri & Sistem",
        ]
        self.assertEqual([g["key"] for g in persantil.SETTINGS_GROUPS_META], expected_groups)

        # 1. Verify all 7 groups render cleanly & contain expected consolidated tabs
        revir_ui.reset_session()
        revir_ui.session_state["main_menu_page"] = "⚙️ Ayarlar"
        for gk in expected_groups:
            revir_ui.session_state["settings_vertical_nav"] = gk
            out = revir_ui.run_action_headless(persantil.main, {"action": "render"})
            self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])
            if gk == "📅 Takvim & Yemekhane":
                self.assertIn("Takvim Ayarları", out["main_html"])
                self.assertIn("Yemekhane Ayarları", out["main_html"])
                self.assertIn("Haftada Kaç Gün Yemek Çıkıyor?", out["main_html"])
            elif gk == "🎨 Görünüm & Düzen":
                self.assertIn("Kutucuk &amp; Pencere Düzeni", out["main_html"])
            elif gk == "🎛️ Görünürlük & Rapor":
                self.assertIn("Rapor Çıktı Ayarları", out["main_html"])

        # 2. Verify legacy settings navigation group mapping
        revir_ui.session_state["settings_vertical_nav"] = "🤖 Yapay Zekâ (Gemini) & KVKK"
        revir_ui.run_action_headless(persantil.main, {"action": "render"})
        self.assertEqual(revir_ui.session_state["settings_vertical_nav"], "🤖 Yapay Zekâ & KVKK")

        # 3. Verify saving box_layout_enabled toggle from Görünüm & Düzen
        revir_ui.session_state["settings_vertical_nav"] = "🎨 Görünüm & Düzen"
        revir_ui.session_state["cfg_box_layout_enabled_toggle"] = False
        revir_ui.run_action_headless(persantil.main, {"action": "click", "key": "btn_save_box_layout_cfg"})
        self.assertEqual(persantil.get_app_settings().get("box_layout_enabled"), "0")
        revir_ui.session_state["cfg_box_layout_enabled_toggle"] = True
        revir_ui.run_action_headless(persantil.main, {"action": "click", "key": "btn_save_box_layout_cfg"})
        self.assertEqual(persantil.get_app_settings().get("box_layout_enabled"), "1")

        # 4. Verify topbar box layout button is disabled on Takvim & Planlama and enabled elsewhere
        revir_ui.session_state["main_menu_page"] = "📅 Takvim & Planlama"
        out_cal = revir_ui.run_action_headless(persantil.main, {"action": "render"})
        tb_cal = out_cal.get("topbar_html") or ""
        cal_btn_tag = tb_cal.split('id="revir-tb-layout-btn"')[1].split(">")[0]
        self.assertIn("disabled", cal_btn_tag)

        revir_ui.session_state["main_menu_page"] = "📊 Genel Bakış"
        out_ov = revir_ui.run_action_headless(persantil.main, {"action": "render"})
        tb_ov = out_ov.get("topbar_html") or ""
        ov_btn_tag = tb_ov.split('id="revir-tb-layout-btn"')[1].split(">")[0]
        self.assertNotIn("disabled", ov_btn_tag)


    def test_10_adults_table_columns_and_role_status_options(self):
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('Sivil / Diğer')")
            c.execute(
                """INSERT INTO students(class_id, full_name, sex, birth_date, guardian_phone, notes, person_type)
                   VALUES (1, 'Mehmet Demir', 'E', '1985-04-10', '05321234567', 'Müdür Yardımcısı', 'Sivil / diğer kişi')"""
            )
            c.execute(
                """INSERT INTO students(class_id, full_name, sex, birth_date, guardian_phone, notes, person_type)
                   VALUES (1, 'Zeynep Kaya', 'K', '1992-09-12', '05339876543', 'Matematik Öğretmeni', 'Sivil / diğer kişi')"""
            )
            c.commit()

        # 1. Adults table shows ONLY Ad Soyad and Görev / Statü columns
        revir_ui.reset_session()
        out = revir_ui.run_action_headless(persantil.render_civilians, {"action": "render"})
        html = out["main_html"]
        self.assertIn("Ad Soyad", html)
        self.assertIn("Görev / Statü", html)
        self.assertIn("Mehmet Demir", html)
        self.assertIn("Müdür Yardımcısı", html)
        self.assertIn("Matematik Öğretmeni", html)
        for hidden_col in ("Boy (cm)", "Kilo (kg)", "Doğum Tarihi", "Cinsiyet", "VKİ (BMI)", "📂 Dosya", "05321234567"):
            self.assertNotIn(hidden_col, html)

        # 2. Clicking Ad Soyad on row opens person dossier dialog
        out_click = revir_ui.run_action_headless(
            persantil.render_civilians,
            {"action": "df_select", "key": "civilian_roster_table_0", "value": {"rows": [0], "cells": [[0, "Ad Soyad"]]}},
        )
        self.assertIsNotNone(out_click.get("dialog_html"))
        self.assertIn("Mehmet Demir", out_click.get("dialog_html") or "")

        # 3. Role / status options include school & dorm management, memur types, teacher branches, İŞKUR, sözleşmeli, hizmetli, stajyer
        required_roles = (
            "Müdür",
            "Müdür Yardımcısı",
            "Yurt Müdürü",
            "Yurt Müdür Yardımcısı",
            "Yurt Yönetim Memuru",
            "Memur",
            "VHKİ Memuru",
            "Mutemet",
            "Sınıf Öğretmeni",
            "Türkçe Öğretmeni",
            "Matematik Öğretmeni",
            "İngilizce Öğretmeni",
            "Rehberlik / PDR Öğretmeni",
            "İŞKUR Personeli",
            "Sözleşmeli Personel",
            "Hizmetli",
            "Stajyer",
        )
        for role_lbl in required_roles:
            self.assertIn(role_lbl, persantil.ADULT_ROLE_STATUS_OPTIONS)

        # 4. New profile dialog & express bar render new roles cleanly
        revir_ui.reset_session()
        out_new = revir_ui.run_action_headless(lambda: persantil.open_new_profile_dialog("Sivil / diğer kişi"), {"action": "render"})
        new_html = (out_new.get("dialog_html") or "") + (out_new.get("main_html") or "")
        self.assertIn("Yurt Müdürü", new_html)
        self.assertIn("İŞKUR Personeli", new_html)
        self.assertIn("Stajyer", new_html)

        revir_ui.reset_session()
        out_exp = revir_ui.run_action_headless(persantil._render_express_adult_registration, {"action": "render"})
        exp_html = out_exp.get("main_html") or ""
        self.assertIn("Hizmetli", exp_html)
        self.assertIn("Sözleşmeli Personel", exp_html)
        self.assertIn("Matematik Öğretmeni", exp_html)

    def test_11_smart_grid_projection_and_overview_widgets(self):
        # 1. Legacy %-based coordinates map onto 12-column grid & respect readability minimums
        legacy = {
            "overview_stats": {"id": "overview_stats", "x_pct": 0, "y_px": 0, "w_pct": 16, "h_px": 100, "visible": True, "z": 1, "accent": "#0d9488"},
            "overview_outbreak": {"id": "overview_outbreak", "x_pct": 51, "y_px": 305, "w_pct": 49, "h_px": 520, "visible": True, "z": 1, "accent": "#f43f5e"},
        }
        projected = persantil._project_widget_layout_to_grid(legacy)
        self.assertEqual(projected["overview_stats"]["min_col_span"], persantil._WIDGET_MIN_GRID_BOUNDS["overview_stats"][0])
        self.assertEqual(projected["overview_stats"]["col_span"], 6)  # Clamped up from ~2 cols for readability
        self.assertEqual(projected["overview_outbreak"]["col"], 6)
        self.assertEqual(projected["overview_outbreak"]["col_span"], 6)

        # 2. _normalize_widget_layout rejects invalid coordinates but keeps valid 12-col fields
        grid_payload = {
            "overview_meds": {
                "id": "overview_meds", "x_pct": 0, "y_px": 0, "w_pct": 50, "h_px": 260,
                "col": 10, "row": 2, "col_span": 2, "row_span": 1, "min_col_span": 4, "min_row_span": 3,
                "visible": True, "z": 1, "accent": "#0d9488",
            }
        }
        norm = persantil._normalize_widget_layout(grid_payload)
        self.assertEqual(norm["overview_meds"]["col_span"], 4)
        self.assertEqual(norm["overview_meds"]["row_span"], 3)
        self.assertEqual(norm["overview_meds"]["col"], 8)  # 12 - 4 boundary clamp

        # 3. Overview page wraps every card cleanly inside smart grid & drops legacy resize script
        revir_ui.reset_session()
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "fake-key", "gemini-2.5-flash")):
            out_ov = revir_ui.run_action_headless(persantil.render_overview, {"action": "render"})
        ov_html = str(out_ov.get("main_html") or "")
        ov_scripts = str(out_ov.get("scripts_html") or "")
        for wid in (
            "overview_stats",
            "overview_quick_links",
            "overview_meds",
            "overview_diabetes",
            "overview_chart",
            "overview_calendar",
            "overview_gemini",
            "overview_pending",
            "overview_outbreak",
            "overview_recent",
        ):
            self.assertIn(f'data-layout-widget-marker="{wid}"', ov_html)
        # Günün Öncelikleri resides inside the overview_stats card
        stats_slice = ov_html.split('data-layout-widget-marker="overview_stats"')[1].split('data-layout-widget-marker="overview_quick_links"')[0]
        self.assertIn("Günün Öncelikleri", stats_slice)
        # Smart grid engine active & legacy conflicting resize system removed
        for js_marker in (
            "resolveGridCollisions",
            "compactGridUpward",
            "detectSwapTarget",
            "revir-layout-grid-overlay",
            "revir-swap-pill",
            "revir_smart_grid_v2:",
        ):
            self.assertIn(js_marker, ov_scripts)
        self.assertNotIn("revir-ov-resizable", ov_html + ov_scripts)
        self.assertNotIn("initOverviewResizableCards", ov_html + ov_scripts)


if __name__ == "__main__":
    unittest.main(verbosity=2)
