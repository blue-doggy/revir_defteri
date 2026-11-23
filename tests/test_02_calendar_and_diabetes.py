"""Unit tests for all functions in school_calendar.py and insulin_diabetes.py."""
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

import revir_ui
import school_calendar
import insulin_diabetes
from tests.db_test_helper import isolated_real_db


class TestSchoolCalendar(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()

    def tearDown(self):
        self.ctx.__exit__(None, None, None)

    def test_render_symbol_badge_all_variants(self):
        for sym in ("official", "holiday", "meb", "exam", "special", "unknown", "", None):
            html_full = school_calendar.render_symbol_badge(sym or "", compact=False)
            html_cmp = school_calendar.render_symbol_badge(sym or "", compact=True)
            self.assertIsInstance(html_full, str)
            self.assertIsInstance(html_cmp, str)

    def test_gregorian_to_hijri_str_boundaries(self):
        res_ok = school_calendar.gregorian_to_hijri_str(date(2026, 11, 22))
        self.assertTrue(isinstance(res_ok, str) and len(res_ok) > 0)
        res_leap = school_calendar.gregorian_to_hijri_str(date(2024, 2, 29))
        self.assertTrue(isinstance(res_leap, str))

    def test_offline_prayer_times_and_provider(self):
        pt = school_calendar._compute_offline_prayer_times("İstanbul", date(2026, 11, 22))
        self.assertIn("İmsak", pt)
        self.assertIn("Akşam", pt)
        pt_unknown = school_calendar._compute_offline_prayer_times("Bilinmeyenİl", date(2026, 6, 21))
        self.assertIn("Güneş", pt_unknown)
        with mock.patch("urllib.request.urlopen", side_effect=OSError("offline")):
            res = school_calendar.get_prayer_times_for_province("Ankara", date(2026, 11, 22), self.connect)
            self.assertIsInstance(res, dict)
            self.assertIn("Öğle", res)

    def test_meb_weeks_and_exam_date_helpers_and_quotes(self):
        weeks = school_calendar.get_active_meb_weeks_for_date(date(2026, 11, 10))
        self.assertIsInstance(weeks, list)
        d = school_calendar._nth_weekday_of_month(2026, 11, 0, 1)
        self.assertEqual(d.weekday(), 0)
        ex_date = school_calendar._resolve_exam_date_for_year(2026, {"month": 6, "weekday": 6, "nth": 1, "title": "LGS"})
        self.assertRegex(ex_date, r"^\d{4}-\d{2}-\d{2}$")
        q = school_calendar.get_builtin_daily_quote_and_hadith(date(2026, 11, 22))
        self.assertTrue(isinstance(q, dict) or isinstance(q, tuple))

    def test_seeding_and_ai_save_and_month_sync(self):
        added = school_calendar.ensure_year_builtin_holidays_seeded(2026, self.connect, include_special_days=True)
        self.assertGreater(added, 0)
        ex_added = school_calendar.seed_selected_exams_to_db(2026, ["LGS"], self.connect)
        self.assertGreaterEqual(ex_added, 0)
        saved = school_calendar.save_ai_generated_events_to_db(
            [{"event_date": "2026-11-25", "title": "Özel Gün ÇŞĞ", "category": "Etkinlik"}],
            self.connect,
            upsert=True,
        )
        self.assertGreaterEqual(saved, 1)
        self.assertEqual(school_calendar.save_ai_generated_events_to_db([], self.connect), 0)
        with self.connect() as c:
            cnt = c.execute("SELECT COUNT(*) FROM calendar_events").fetchone()[0]
            self.assertGreater(cnt, 0)

        revir_ui.reset_session()
        revir_ui.session_state["cal_selected_date"] = date(2026, 10, 15)
        school_calendar._sync_selected_date_when_month_changes(2026, 11, today=date(2026, 11, 22))
        self.assertEqual(str(revir_ui.session_state["cal_selected_date"])[:7], "2026-11")

    def test_calendar_ui_renders_headless(self):
        revir_ui.reset_session()
        settings = {
            "calendar_province": "Ankara",
            "calendar_show_prayer_times": "1",
            "calendar_quote_mode": "Karışık (Hadis-i Şerif + Özlü Söz / Günün Yazısı)",
            "calendar_auto_load_holidays": "1",
            "gemini_enabled": "0",
        }
        kwargs = dict(
            connect_db_fn=self.connect,
            get_app_settings_fn=lambda: settings,
            save_app_settings_fn=lambda d: settings.update(d),
            get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
            get_tr_provinces_fn=lambda: ["Ankara", "İstanbul", "İzmir"],
        )
        with mock.patch.object(school_calendar, "get_prayer_times_for_province", return_value={
            "İmsak": "06:00", "Güneş": "07:30", "Öğle": "12:50", "İkindi": "15:20", "Akşam": "17:45", "Yatsı": "19:05"
        }):
            def _app():
                school_calendar.render_calendar_and_planning_page(**kwargs)
                school_calendar.render_calendar_settings_tab(
                    connect_db_fn=self.connect,
                    current_cfg=settings,
                    save_app_settings_fn=lambda d: settings.update(d),
                    get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
                    get_tr_provinces_fn=lambda: ["Ankara", "İstanbul", "İzmir"],
                )
                school_calendar.render_overview_calendar_widget(
                    connect_db_fn=self.connect,
                    get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
                )
            out = revir_ui.run_action_headless(_app, {"action": "render"})
            self.assertIn("main_html", out)
            self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])

            out2 = revir_ui.run_action_headless(
                lambda: school_calendar.open_calendar_event_manager_dialog(
                    sel_iso="2026-11-22",
                    connect_db_fn=self.connect,
                    get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
                ),
                {"action": "render"},
            )
            self.assertIn("main_html", out2)
            self.assertNotIn("Çalışma Zamanı Hatası", out2["main_html"])

    def test_calendar_event_box_confirm_and_delete_buttons(self):
        today_iso = date.today().isoformat()
        with self.connect() as c:
            cur = c.execute(
                """INSERT INTO calendar_events(
                    event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
                    start_time, end_time, description, source, is_completed, created_at
                ) VALUES (?, '4-A Göz Taraması', 'revir_etkinlik', 'health', 0, 0, '09:30', '11:00', 'Test tarama', 'manual', 0, ?)""",
                (today_iso, datetime.now().isoformat(timespec="seconds")),
            )
            ev_id = int(cur.lastrowid)
            c.commit()

        settings = {
            "calendar_province": "Ankara",
            "calendar_show_prayer_times": "1",
            "calendar_quote_mode": "Karışık (Hadis-i Şerif + Özlü Söz / Günün Yazısı)",
            "calendar_auto_load_holidays": "0",
        }
        cal_kwargs = dict(
            connect_db_fn=self.connect,
            get_app_settings_fn=lambda: settings,
            save_app_settings_fn=lambda d: settings.update(d),
            get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
            get_tr_provinces_fn=lambda: ["Ankara", "İstanbul", "İzmir"],
        )

        with mock.patch.object(school_calendar, "get_prayer_times_for_province", return_value={
            "İmsak": "06:00", "Güneş": "07:30", "Öğle": "12:50", "İkindi": "15:20", "Akşam": "17:45", "Yatsı": "19:05"
        }):
            revir_ui.reset_session()
            revir_ui.session_state["cal_selected_date"] = today_iso
            out_cal = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_cal["main_html"])
            self.assertIn(f"btn_cal_ev_done_{ev_id}", out_cal["main_html"])
            self.assertIn(f"btn_cal_ev_del_{ev_id}", out_cal["main_html"])
            self.assertEqual(out_cal["main_html"].count(f"btn_cal_ev_done_{ev_id}"), 1)
            self.assertEqual(out_cal["main_html"].count(f"btn_cal_ev_del_{ev_id}"), 1)

            other_day_iso = date.today().replace(day=2 if date.today().day == 1 else 1).isoformat()
            revir_ui.session_state["cal_selected_date"] = other_day_iso
            out_other = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_other["main_html"])
            self.assertNotIn(f"btn_cal_ev_done_{ev_id}", out_other["main_html"])
            self.assertNotIn(f"btn_cal_ev_del_{ev_id}", out_other["main_html"])
            revir_ui.session_state["cal_selected_date"] = today_iso

            out_ov = revir_ui.run_action_headless(
                lambda: school_calendar.render_overview_calendar_widget(
                    connect_db_fn=self.connect,
                    get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
                ),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_ov["main_html"])
            self.assertIn(f"btn_cal_ev_done_{ev_id}", out_ov["main_html"])
            self.assertIn(f"btn_cal_ev_del_{ev_id}", out_ov["main_html"])

            # 1. Click ✓ (done) on calendar page -> is_completed == 1 and line-through / ✓ Yapıldı shown
            out_done = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "click", "key": f"btn_cal_ev_done_{ev_id}"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_done["main_html"])
            self.assertIn("✓ Yapıldı", out_done["main_html"])
            self.assertIn("line-through", out_done["main_html"])
            with self.connect() as c:
                self.assertEqual(
                    int(c.execute("SELECT is_completed FROM calendar_events WHERE id = ?", (ev_id,)).fetchone()[0]),
                    1,
                )

            # 2. Click ✓ again -> toggles back to is_completed == 0
            out_undo = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "click", "key": f"btn_cal_ev_done_{ev_id}"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_undo["main_html"])
            with self.connect() as c:
                self.assertEqual(
                    int(c.execute("SELECT is_completed FROM calendar_events WHERE id = ?", (ev_id,)).fetchone()[0]),
                    0,
                )

            # 3. Click ✕ (delete) -> removes row from DB
            out_del = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "click", "key": f"btn_cal_ev_del_{ev_id}"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_del["main_html"])
            with self.connect() as c:
                self.assertEqual(
                    int(c.execute("SELECT COUNT(*) FROM calendar_events WHERE id = ?", (ev_id,)).fetchone()[0]),
                    0,
                )


class TestInsulinDiabetes(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('4-A')")
            c.execute("INSERT INTO students(class_id, full_name, sex, birth_date) VALUES (1, 'Ömer Şahin', 'E', '2016-05-01')")
            c.execute("INSERT INTO health_profiles(student_id, chronic_conditions, updated_at) VALUES (1, '', '2026-11-22')")
            c.commit()

    def tearDown(self):
        self.ctx.__exit__(None, None, None)

    def test_profile_crud_and_sync(self):
        dp = insulin_diabetes.default_diabetes_profile(1)
        self.assertEqual(dp["student_id"], 1)
        prof = insulin_diabetes.get_diabetes_profile(self.connect, 1)
        self.assertEqual(prof["student_id"], 1)
        prof["is_diabetic"] = 1
        prof["icr_breakfast"] = 12.5
        prof["isf"] = 45.0
        insulin_diabetes.save_diabetes_profile(self.connect, 1, prof, sync_chronic=True)
        prof2 = insulin_diabetes.get_diabetes_profile(self.connect, 1)
        self.assertEqual(float(prof2["icr_breakfast"]), 12.5)
        self.assertEqual(float(prof2["isf"]), 45.0)
        with self.connect() as c:
            cd = c.execute("SELECT chronic_conditions FROM health_profiles WHERE student_id=1").fetchone()[0]
            self.assertIn("Diyabet", cd)

    def test_menus_crud(self):
        insulin_diabetes.upsert_school_meal_menu(
            self.connect, "2026-11-22", "Öğle",
            [{"name": "Mercimek Çorbası", "carbs_g": 20.0}, {"name": "Pilav", "carbs_g": 30.0}],
            source_note="Test menü",
        )
        menus = insulin_diabetes.get_school_menus_for_month(self.connect, "2026-11")
        self.assertEqual(len(menus), 1)
        got = insulin_diabetes.get_school_menu_for_date(self.connect, "2026-11-22", "Öğle")
        self.assertIsNotNone(got)
        insulin_diabetes.delete_school_menu_by_id(self.connect, menus[0]["id"])
        self.assertIsNone(insulin_diabetes.get_school_menu_for_date(self.connect, "2026-11-22", "Öğle"))

    def test_iob_site_round_bolus_and_tir(self):
        self.assertEqual(insulin_diabetes._round_to_increment(3.26, 0.5), 3.5)
        self.assertEqual(insulin_diabetes._round_to_increment(3.24, 0.5), 3.0)
        self.assertEqual(insulin_diabetes._round_to_increment(3.26, 0.0), 3.5)
        self.assertEqual(insulin_diabetes._round_to_increment(-1.0, 0.5), 0.0)

        prof = insulin_diabetes.get_diabetes_profile(self.connect, 1)
        rec = insulin_diabetes.calculate_bolus_recommendation(
            profile=prof,
            meal_type="Öğle Yemeği",
            bg_mg_dl=210.0,
            menu_carbs_raw=45.0,
            menu_consumed_pct=100.0,
            extra_carbs_g=0.0,
            iob_units=0.5,
        )
        self.assertIsInstance(rec, dict)
        self.assertIn("recommended_dose_units", rec)
        self.assertGreaterEqual(rec["recommended_dose_units"], 0.0)

        rec_low = insulin_diabetes.calculate_bolus_recommendation(
            profile=prof,
            meal_type="Öğle Yemeği",
            bg_mg_dl=55.0,
            menu_carbs_raw=15.0,
        )
        self.assertEqual(rec_low["recommended_dose_units"], 0.0)

        now_dt = datetime.now()
        with self.connect() as c:
            c.execute(
                """INSERT INTO insulin_bolus_logs(
                    student_id, log_date, log_time, meal_type, bg_mg_dl, total_carbs_g,
                    calculated_dose_units, applied_dose_units, injection_site, notes, created_at
                ) VALUES (1, ?, ?, 'Öğle Yemeği', 180, 40, 4.0, 4.0, 'Karın Sol', 'test', ?)""",
                (now_dt.strftime("%Y-%m-%d"), now_dt.strftime("%H:%M"), now_dt.isoformat(timespec="seconds")),
            )
            c.commit()

        iob, active = insulin_diabetes.calculate_active_iob(self.connect, 1, datetime.now())
        self.assertGreater(iob, 0.0)
        self.assertEqual(len(active), 1)
        self.assertEqual(insulin_diabetes.get_last_injection_site(self.connect, 1)[0], "Karın Sol")

        logs_df = insulin_diabetes.get_student_bolus_logs_df(self.connect, 1, days_back=30)
        self.assertEqual(len(logs_df), 1)
        tir = insulin_diabetes.compute_glycemic_tir_stats(logs_df)
        self.assertEqual(tir["total_logs"], 1)
        self.assertEqual(insulin_diabetes.compute_glycemic_tir_stats(pd.DataFrame())["total_logs"], 0)

        plan_df, plan_prof = insulin_diabetes.build_monthly_insulin_plan_df(self.connect, 1, datetime.now().strftime("%Y-%m"))
        self.assertIsInstance(plan_df, pd.DataFrame)
        self.assertIsInstance(plan_prof, dict)

    def test_diabetes_ui_renders_headless(self):
        import persantil
        revir_ui.reset_session()
        out = revir_ui.run_action_headless(persantil.render_insulin_diabetes, {"action": "render"})
        self.assertIn("main_html", out)
        self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
