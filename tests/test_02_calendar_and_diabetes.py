"""Unit tests for all functions in school_calendar.py and insulin_diabetes.py."""
from __future__ import annotations

import json
import re
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
            self.assertIn("revir-cal-split-anchor", out_cal["main_html"])
            self.assertIn(".revir-cal-split-anchor + .revir-cols", out_cal.get("styles_html", ""))
            self.assertIn(".revir-leaf-ev-card:hover .revir-leaf-ev-del", out_cal.get("styles_html", ""))
            self.assertIn(".revir-leaf-ev-card:hover .revir-leaf-ev-done", out_cal.get("styles_html", ""))
            self.assertIn('role="checkbox"', out_cal["main_html"])
            self.assertIn("revir-leaf-ev-card", out_cal["main_html"])
            self.assertIn("✓ Tamamlandı", out_cal["main_html"])
            self.assertIn(f"btn_cal_ev_done_{ev_id}", out_cal["main_html"])
            self.assertIn(f"btn_cal_ev_del_{ev_id}", out_cal["main_html"])
            self.assertEqual(out_cal["main_html"].count(f"btn_cal_ev_done_{ev_id}"), 1)
            self.assertEqual(out_cal["main_html"].count(f"btn_cal_ev_del_{ev_id}"), 1)
            self.assertIn("Tamamlanan Etkinlikler (0)", out_cal["main_html"])

            other_day_iso = date.today().replace(day=2 if date.today().day == 1 else 1).isoformat()
            revir_ui.session_state["cal_selected_date"] = other_day_iso
            out_other = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "render"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_other["main_html"])
            self.assertNotIn(f"btn_cal_ev_done_{ev_id}", out_other["main_html"])
            self.assertNotIn(f"btn_cal_ev_del_{ev_id}", out_other["main_html"])
            self.assertIn("4-A Göz Taraması", out_other["main_html"])
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

            # 1. Click ✓ Tamamlandı -> removed from leaf, moved to Tamamlanan Etkinlikler, still visible in monthly grid
            out_done = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "click", "key": f"btn_cal_ev_done_{ev_id}"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_done["main_html"])
            with self.connect() as c:
                self.assertEqual(
                    int(c.execute("SELECT is_completed FROM calendar_events WHERE id = ?", (ev_id,)).fetchone()[0]),
                    1,
                )
            grid_part = out_done["main_html"].split("🗓️ GÜNLÜK TAKVİM YAPRAĞI")[0]
            self.assertIn("✓ 4-A Göz Taraması", grid_part)
            leaf_part = out_done["main_html"].split("📌 GÜNÜN TATİL, SINAV VE ETKİNLİKLERİ")[1].split("EZAN VAKİTLERİ")[0]
            self.assertNotIn("4-A Göz Taraması", leaf_part)
            self.assertIn("Bu güne ait tüm etkinlikler tamamlandı", leaf_part)
            self.assertIn("Tamamlanan Etkinlikler (1)", out_done["main_html"])
            completed_part = out_done["main_html"].split("Tamamlanan Etkinlikler (1)")[1]
            self.assertIn("4-A Göz Taraması", completed_part)
            self.assertIn("↩️ Geri Al", completed_part)

            # 1b. Event management dialog clearly marks completed event as "✅ Tamamlandı"
            out_mgr_done = revir_ui.run_action_headless(
                lambda: school_calendar.open_calendar_event_manager_dialog(
                    sel_iso=today_iso,
                    connect_db_fn=self.connect,
                    get_active_gemini_config_fn=lambda: (False, "", "gemini-2.5-flash"),
                ),
                {"action": "render"},
            )
            mgr_html = (out_mgr_done.get("dialog_html") or "") + (out_mgr_done.get("main_html") or "")
            self.assertIn("✅ Tamamlandı", mgr_html)
            self.assertIn("4-A Göz Taraması", mgr_html)
            self.assertIn("↩️ Geri Al", mgr_html)

            # 2. Click ↩️ Geri Al -> toggles back to is_completed == 0 and returns to leaf
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
            self.assertIn("Tamamlanan Etkinlikler (0)", out_undo["main_html"])
            leaf_undo_part = out_undo["main_html"].split("📌 GÜNÜN TATİL, SINAV VE ETKİNLİKLERİ")[1].split("EZAN VAKİTLERİ")[0]
            self.assertIn("4-A Göz Taraması", leaf_undo_part)
            self.assertIn("✓ Tamamlandı", leaf_undo_part)

            # 3. Click ✕ (delete) -> removes row from DB and disappears everywhere (grid, leaf, completed)
            out_del = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "click", "key": f"btn_cal_ev_del_{ev_id}"},
            )
            self.assertNotIn("Çalışma Zamanı Hatası", out_del["main_html"])
            self.assertIn("takvimden silindi", out_del["main_html"])
            self.assertNotIn("4-A Göz Taraması", out_del["main_html"].split("Kayıtlı Gün/Etkinlik:")[1])
            out_after_del = revir_ui.run_action_headless(
                lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs),
                {"action": "render"},
            )
            self.assertNotIn("4-A Göz Taraması", out_after_del["main_html"])
            with self.connect() as c:
                self.assertEqual(
                    int(c.execute("SELECT COUNT(*) FROM calendar_events WHERE id = ?", (ev_id,)).fetchone()[0]),
                    0,
                )

    def test_calendar_leaf_semantic_dedup_ai_consolidation_and_meta_box_actions(self):
        # 1. Seed semantic duplicates on 2026-11-22 (Diş Hekimleri Günü + MEB Haftası)
        now_s = datetime.now().isoformat(timespec="seconds")
        with self.connect() as c:
            c.execute(
                """INSERT INTO calendar_events(
                    event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
                    profession_group, description, source, is_completed, created_at
                ) VALUES (?, ?, 'meslek_gunu', 'profession', 0, 0, ?, ?, 'manual', 0, ?)""",
                ("2026-11-22", "22 Kasım Diş Hekimleri Günü & Ağız ve Diş Sağlığı Haftası", "Diş Hekimleri ve Ağız-Diş Sağlığı Çalışanları", "Günün Meslek Grubu: Diş Hekimleri", now_s),
            )
            c.execute(
                """INSERT INTO calendar_events(
                    event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
                    profession_group, description, source, is_completed, created_at
                ) VALUES (?, ?, 'belirli_gun', 'profession', 0, 0, ?, ?, 'builtin', 0, ?)""",
                ("2026-11-22", "MEB - 22 Kasım Diş Hekimleri Günü & Ağız ve Diş Sağlığı Haftası (21-27 Kasım)", "Diş Hekimleri ve Ağız-Diş Sağlığı Çalışanları", "", now_s),
            )
            c.commit()

        # AI save rejects semantic duplicate on same date
        ins_dup = school_calendar.save_ai_generated_events_to_db(
            [{"event_date": "2026-11-22", "title": "Ağız ve Diş Sağlığı Haftası (21–27 Kasım)", "category": "belirli_gun"}],
            self.connect,
            upsert=False,
        )
        self.assertEqual(ins_dup, 0)

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
            get_tr_provinces_fn=lambda: ["Ankara"],
        )
        with mock.patch.object(school_calendar, "get_prayer_times_for_province", return_value={
            "İmsak": "06:00", "Güneş": "07:30", "Öğle": "12:50", "İkindi": "15:20", "Akşam": "17:45", "Yatsı": "19:05"
        }):
            revir_ui.reset_session()
            revir_ui.session_state["cal_view_year"] = 2026
            revir_ui.session_state["cal_view_month"] = 11
            revir_ui.session_state["cal_selected_date"] = "2026-11-22"
            out = revir_ui.run_action_headless(lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs), {"action": "render"})
            html_leaf = out["main_html"].split("🗓️ GÜNLÜK TAKVİM YAPRAĞI", 1)[1]
            # Semantic duplicates merged cleanly: 1 event card instead of 3+ overlapping cards, with merged Revir focus
            self.assertEqual(html_leaf.count("class=\"revir-leaf-ev-done\""), 2)  # 1 Profession card + 1 Unified Event card
            self.assertIn("Doğru diş fırçalama", html_leaf)
            # Profession card has hover X & Tamamlandı checkbox
            prof_slice = html_leaf.split("GÜNÜN MESLEK GRUBU & ANLAMI", 1)[1].split("GÜNÜN TATİL, SINAV VE ETKİNLİKLERİ", 1)[0]
            self.assertIn("btn_cal_meta_del_", prof_slice)
            self.assertIn("btn_cal_meta_done_", prof_slice)

            # Complete profession card -> removed from leaf, appears in Tamamlanan Etkinlikler, undo restores it
            m_prof_done = re.search(r"btn_cal_meta_done_(prof_[0-9_a-f]+)", prof_slice)
            self.assertIsNotNone(m_prof_done)
            prof_wid = m_prof_done.group(0)
            out_pdone = revir_ui.run_action_headless(lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs), {"action": "click", "key": prof_wid})
            leaf_after = out_pdone["main_html"].split("🗓️ GÜNLÜK TAKVİM YAPRAĞI", 1)[1]
            self.assertNotIn("GÜNÜN MESLEK GRUBU & ANLAMI", leaf_after.split("EZAN VAKİTLERİ", 1)[0])
            self.assertIn("Tamamlanan Etkinlikler (1)", leaf_after)
            m_undo = re.search(r"btn_cal_ev_done_(\d+)", leaf_after.split("Tamamlanan Etkinlikler (1)", 1)[1])
            self.assertIsNotNone(m_undo)
            out_pundo = revir_ui.run_action_headless(lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs), {"action": "click", "key": m_undo.group(0)})
            self.assertIn("GÜNÜN MESLEK GRUBU & ANLAMI", out_pundo["main_html"])

            # Delete MEB Week box on 2026-11-12 -> hides permanently on that date
            revir_ui.session_state["cal_selected_date"] = "2026-11-12"
            out_meb = revir_ui.run_action_headless(lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs), {"action": "render"})
            self.assertIn("Atatürk Haftası", out_meb["main_html"])
            m_meb_del = re.search(r"btn_cal_meta_del_(meb_[0-9_a-f]+)", out_meb["main_html"])
            self.assertIsNotNone(m_meb_del)
            out_meb_after = revir_ui.run_action_headless(lambda: school_calendar.render_calendar_and_planning_page(**cal_kwargs), {"action": "click", "key": m_meb_del.group(0)})
            self.assertNotIn("Atatürk Haftası", out_meb_after["main_html"].split("🗓️ GÜNLÜK TAKVİM YAPRAĞI", 1)[1])

        # 2. Gemini AI consolidation physically merges duplicate DB records down to 1
        with self.connect() as c:
            ids_22 = [int(r["id"]) for r in c.execute("SELECT id FROM calendar_events WHERE event_date='2026-11-22' ORDER BY id").fetchall()]
        self.assertEqual(len(ids_22), 2)
        ai_merge_resp = json.dumps({
            "summary": "Diş Hekimleri Günü ve Ağız-Diş Sağlığı Haftası tekilleştirildi.",
            "merged_groups": [{
                "keep_id": ids_22[0],
                "keep_title": "22 Kasım Diş Hekimleri Günü & Ağız ve Diş Sağlığı Haftası",
                "keep_description": "Doğru diş fırçalama ve ağız-diş sağlığı farkındalığı",
                "keep_profession_group": "Diş Hekimleri ve Ağız-Diş Sağlığı Çalışanları",
                "drop_ids": [ids_22[1]],
                "suppress_meta_keys": [],
            }],
        }, ensure_ascii=False)
        with mock.patch.object(school_calendar, "consolidate_calendar_items_with_ai", return_value=(True, json.loads(ai_merge_resp), "gemini-2.5-flash")):
            chg, msg = school_calendar.consolidate_calendar_day_with_ai_or_local(
                self.connect, target_date="2026-11-22", use_ai=True, api_key="fake-key", model="gemini-2.5-flash"
            )
        self.assertGreaterEqual(chg, 2)
        with self.connect() as c:
            rem_cnt = c.execute("SELECT COUNT(*) FROM calendar_events WHERE event_date='2026-11-22'").fetchone()[0]
        self.assertEqual(rem_cnt, 1)


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
        self.assertIn("✨ AI ile Bilgi Al", out["main_html"])

        # Yazılan atıştırmalık için "✨ AI ile Bilgi Al" tıklanınca Karb (g), kcal ve Gİ bilgisi gelmeli ve "➕ Ekle" ile listeye eklenmeli
        persantil.save_app_settings({"gemini_enabled": "1", "gemini_api_key": "test-fake-key"})
        revir_ui.session_state["ins_ex_nm_1"] = "1 paket çubuk kraker ve kutu ayran"
        fake_food_ai = {
            "items": [
                {"name": "Çubuk Kraker", "portion": "1 paket (40g)", "carbs_g": 28.0, "kcal": 165.0, "gi": "Yüksek"},
                {"name": "Ayran", "portion": "1 kutu (170ml)", "carbs_g": 4.5, "kcal": 65.0, "gi": "Düşük"},
            ],
            "total_carbs_g": 32.5,
            "total_kcal": 230.0,
            "glycemic_note": "Ayran krakerin glisemik yükseliş hızını dengeler.",
        }
        with mock.patch("gemini_assistant.estimate_extra_foods_carbs_with_gemini", return_value=(True, fake_food_ai, "gemini-3-flash-preview")):
            out_lk = revir_ui.run_action_headless(persantil.render_insulin_diabetes, {"action": "click", "key": "ins_ex_ai_lookup_1"})
        self.assertIn("AI Besin &amp; Karbonhidrat Bilgisi", out_lk["main_html"])
        self.assertAlmostEqual(float(revir_ui.session_state.get("ins_ex_cg_1") or 0.0), 32.5, places=1)

        out_add = revir_ui.run_action_headless(persantil.render_insulin_diabetes, {"action": "click", "key": "ins_ex_add_1"})
        extras = revir_ui.session_state.get("ins_extras_list_1") or []
        self.assertEqual(len(extras), 2)
        self.assertAlmostEqual(sum(x["carbs_g"] for x in extras), 32.5, places=1)
        self.assertIn("Çubuk Kraker", out_add["main_html"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
