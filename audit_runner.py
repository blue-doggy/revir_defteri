"""
Exhaustive Isolated Sandbox Audit Runner for Revir Defteri
Runs 10 automated integration, fuzzing, cross-module propagation, export, backup,
frozen-path, and headless UI verification suites inside a fresh TemporaryDirectory.
Never touches production DB or workspace folders.
"""
from __future__ import annotations

import contextlib
import io
import json
import math
import os
import sqlite3
import sys
import tempfile
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest import mock

import openpyxl
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import cafeteria as caf
import desktop_app
import document_studio as dstudio
import gemini_assistant as gai
import insulin_diabetes as idm
import persantil as app
import revir_ui
import school_calendar as scal

try:
    import docx  # noqa: F401
except Exception:
    docx = None


class AuditReport:
    def __init__(self) -> None:
        self.results: list[dict[str, str]] = []

    def check(self, suite: str, name: str, cond: bool, detail: str = "", severity: str = "P1") -> bool:
        status = "PASS" if cond else "FAIL"
        self.results.append({
            "suite": suite,
            "name": name,
            "status": status,
            "severity": severity if not cond else "OK",
            "detail": detail,
        })
        mark = "PASS" if cond else f"FAIL [{severity}]"
        msg = f"  [{mark}] {name}"
        if detail:
            msg += f" :: {detail}"
        print(msg)
        return cond

    def warn(self, suite: str, name: str, detail: str = "", severity: str = "P2") -> None:
        self.results.append({
            "suite": suite,
            "name": name,
            "status": "WARN",
            "severity": severity,
            "detail": detail,
        })
        print(f"  [WARN {severity}] {name} :: {detail}")

    def summary(self) -> tuple[int, int, int]:
        p = sum(1 for r in self.results if r["status"] == "PASS")
        w = sum(1 for r in self.results if r["status"] == "WARN")
        f = sum(1 for r in self.results if r["status"] == "FAIL")
        return p, w, f


@contextlib.contextmanager
def sandbox_env():
    """Creates a clean isolated TempDir with DB + folders matching production layout."""
    with tempfile.TemporaryDirectory(prefix="revir_audit_") as tmp:
        tmp_p = Path(tmp)
        db_p = tmp_p / "revir_defteri.db"
        kisi_p = tmp_p / "kisi_verileri"
        yedek_p = tmp_p / "oto_yedekler"
        egitim_p = tmp_p / "egitim_yillari"
        evrak_p = tmp_p / "revir_evraklari"
        for d in (kisi_p, yedek_p, egitim_p, evrak_p):
            d.mkdir(parents=True, exist_ok=True)

        patches = [
            mock.patch.object(app, "APP_DIR", tmp_p),
            mock.patch.object(app, "_RESOLVED_APP_DIR", tmp_p),
            mock.patch.object(app, "DB_PATH", db_p),
            mock.patch.object(app, "PERSON_DATA_DIR", kisi_p),
            mock.patch.object(app, "AUTO_BACKUP_DIR", yedek_p),
            mock.patch.object(app, "EGITIM_YILLARI_DIR", egitim_p),
            mock.patch.object(app, "EVRAKLAR_DIR", evrak_p),
        ]
        for pt in patches:
            pt.start()
        revir_ui.reset_session()
        app._DB_PATH_CACHE.clear()
        app._ACTIVE_YEAR_CACHE.clear()
        try:
            app.ensure_db_initialized.clear()
        except Exception:
            pass
        app.clear_all_data_caches()
        try:
            yield tmp_p
        finally:
            try:
                app.ensure_db_initialized.clear()
            except Exception:
                pass
            app._DB_PATH_CACHE.clear()
            app._ACTIVE_YEAR_CACHE.clear()
            app.clear_all_data_caches()
            revir_ui.reset_session()
            for pt in reversed(patches):
                pt.stop()


def _seed_class_and_student(
    class_name: str = "4-A",
    full_name: str = "Ali Yılmaz",
    sex: str = "E",
    birth_date: str = "2017-03-15",
    person_type: str = "Öğrenci",
    student_no: str = "101",
    guardian_name: str = "Fatma Yılmaz",
    guardian_phone: str = "0532 111 22 33",
) -> tuple[int, int]:
    with app.connect_db() as conn:
        c_row = conn.execute("SELECT id FROM classes WHERE name=?", (class_name,)).fetchone()
        if c_row:
            cid = int(c_row["id"])
        else:
            cid = int(conn.execute("INSERT INTO classes(name) VALUES (?)", (class_name,)).lastrowid)
        cur = conn.execute(
            """INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date, guardian_name, guardian_phone)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (cid, person_type, student_no, full_name, sex, birth_date, guardian_name, guardian_phone),
        )
        sid = int(cur.lastrowid)
    app.sync_after_person_mutation(sid)
    return cid, sid


def _seed_visit(student_id: int, v_date: str | None = None, v_time: str = "10:15", complaint: str = "Baş ağrısı", referral: str = "Sevk edilmedi", follow_up: str = "İzlem Gerekmiyor") -> int:
    vd = v_date or date.today().isoformat()
    with app.connect_db() as conn:
        fup_st = app._infer_visit_follow_up_status(referral, follow_up)
        cur = conn.execute(
            """INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms, observation, action, guardian_contacted, referral, follow_up, follow_up_status, incident_location, witness_teacher)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (student_id, vd, v_time, complaint, "Halsizlik", "Ateş normal", "Dinlendirildi", 0, referral, follow_up, fup_st, "Sınıf", "Öğretmen A"),
        )
        vid = int(cur.lastrowid)
    app.sync_after_person_mutation(student_id)
    return vid


def run_suite_1(rep: AuditReport) -> None:
    print("\n=== SUITE 1: Database Initialization, Migrations & Academic Year Isolation ===")
    s = "Suite 1"
    with sandbox_env() as tmp_p:
        app.initialize_db()
        rep.check(s, "initialize_db + ensure_db_initialized", app.ensure_db_initialized() is True, severity="P0")
        with app.connect_db() as conn:
            caf.ensure_cafeteria_tables(conn)
            idm.ensure_diabetes_tables(conn)
            fk_on = conn.execute("PRAGMA foreign_keys").fetchone()[0]
            integ = conn.execute("PRAGMA integrity_check").fetchone()[0]
            rep.check(s, "PRAGMA foreign_keys = ON", fk_on == 1, f"got {fk_on}", severity="P0")
            rep.check(s, "PRAGMA integrity_check = ok", str(integ).lower() == "ok", f"got {integ}", severity="P0")

        years0 = app.get_all_academic_years()
        rep.check(s, "get_all_academic_years initialized", len(years0) >= 1, f"len={len(years0)}")

        # Seed rich cross-module records in default active year before cloning/creating fresh year
        cid_1a, sid_1 = _seed_class_and_student("1-A", "Ayşe Kaya", "K", "2019-05-10", student_no="10")
        cid_4a, sid_2 = _seed_class_and_student("4-A", "Mehmet Demir", "E", "2016-05-10", student_no="20")
        vid_1 = _seed_visit(sid_1)
        app.save_school_screening(sid_1, date.today().isoformat(), height_cm=115.0, weight_kg=21.0)
        app.add_seizure_log(sid_1, date.today().isoformat(), "10:00", 45, "Fokal", triggers="Işık")
        app.add_isolation_record(sid_1, "Suçiçeği", date.today().isoformat())
        idm.save_diabetes_profile(lambda: app.connect_db(), sid_1, {"is_diabetic": True, "icr_breakfast": 10.0, "target_bg": 110.0, "correction_threshold": 150.0})
        caf.upsert_diner(lambda: app.connect_db(), student_id=sid_1, full_name="Ayşe Kaya", role_type="Öğrenci", class_or_dept="1-A")
        caf.upsert_diner(lambda: app.connect_db(), student_id=sid_2, full_name="Mehmet Demir", role_type="Öğrenci", class_or_dept="4-A")

        # 1a. create_academic_year(policy='fresh') MUST NOT leak students/seizures/isolations/diabetes/cafeteria_diners
        ay_fresh = app.create_academic_year("2027-2028 Fresh", "2027-09-01", "2028-06-20", student_policy="fresh")
        app.set_active_academic_year(ay_fresh)
        app.initialize_db()
        with app.connect_db() as conn:
            caf.ensure_cafeteria_tables(conn)
            idm.ensure_diabetes_tables(conn)
            rem_students = conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
            rem_seizure = conn.execute("SELECT COUNT(*) FROM seizure_logs").fetchone()[0]
            rem_iso = conn.execute("SELECT COUNT(*) FROM isolation_records").fetchone()[0]
            rem_diab = conn.execute("SELECT COUNT(*) FROM diabetes_profiles").fetchone()[0]
            rem_diners = conn.execute("SELECT COUNT(*) FROM cafeteria_diners").fetchone()[0]
            fk_chk = conn.execute("PRAGMA foreign_key_check").fetchall()
        rep.check(
            s,
            "create_academic_year('fresh') cleans all student/clinical/cafeteria tables",
            (rem_students == 0 and rem_seizure == 0 and rem_iso == 0 and rem_diab == 0 and rem_diners == 0 and len(fk_chk) == 0),
            f"stu={rem_students}, seiz={rem_seizure}, iso={rem_iso}, diab={rem_diab}, diners={rem_diners}, fk_err={len(fk_chk)}",
            severity="P0",
        )

        # 1b. Switch back & test 'clone' and 'promote'
        app.set_active_academic_year(years0[0]["id"])
        ay_clone = app.create_academic_year("2028-2029 Clone", "2028-09-01", "2029-06-20", student_policy="clone")
        app.set_active_academic_year(ay_clone)
        with app.connect_db() as conn:
            cnt_clone_stu = conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
            cnt_clone_vis = conn.execute("SELECT COUNT(*) FROM visits").fetchone()[0]
            cnt_clone_seiz = conn.execute("SELECT COUNT(*) FROM seizure_logs").fetchone()[0]
            cnt_clone_iso = conn.execute("SELECT COUNT(*) FROM isolation_records").fetchone()[0]
        rep.check(s, "create_academic_year('clone') preserves roster, clears episodic visits/seizures/isolations", cnt_clone_stu == 2 and cnt_clone_vis == 0 and cnt_clone_seiz == 0 and cnt_clone_iso == 0, f"stu={cnt_clone_stu}, vis={cnt_clone_vis}, seiz={cnt_clone_seiz}, iso={cnt_clone_iso}")

        app.set_active_academic_year(years0[0]["id"])
        ay_prom = app.create_academic_year("2029-2030 Promote", "2029-09-01", "2030-06-20", student_policy="promote", graduation_grade=4)
        app.set_active_academic_year(ay_prom)
        with app.connect_db() as conn:
            cnames = {r[0] for r in conn.execute("SELECT name FROM classes").fetchall()}
            prom_diners = conn.execute("SELECT COUNT(*) FROM cafeteria_diners").fetchone()[0]
            prom_fk = conn.execute("PRAGMA foreign_key_check").fetchall()
        rep.check(
            s,
            "create_academic_year('promote') promotes 1-A->2-A & graduates 4-A cleanly without FK or diner orphans",
            ("2-A" in cnames and "4-A" not in cnames and prom_diners == 1 and len(prom_fk) == 0),
            f"classes={cnames}, diners={prom_diners}, fk_err={len(prom_fk)}",
            severity="P0",
        )

        # Test update, transfer, delete year
        app.update_academic_year(ay_prom, "2029-2030 Rev", "2029-09-05", "2030-06-25")
        yrs_now = app.get_all_academic_years()
        rep.check(s, "update_academic_year updates metadata", any(y["name"] == "2029-2030 Rev" for y in yrs_now))
        t_cnt = app.transfer_students_from_year(years0[0]["id"], [sid_2])
        rep.check(s, "transfer_students_from_year transfers graduating/other-year student", t_cnt == 1, f"transferred={t_cnt}")

        app.set_active_academic_year(years0[0]["id"])
        ok_del_ay, _ = app.delete_academic_year(ay_prom)
        rep.check(s, "delete_academic_year removes secondary year", ok_del_ay)

        # Test delete cascades + FK integrity
        cid_del, sid_del = _seed_class_and_student("3-B", "Cascade Test", "K", "2018-01-01", student_no="99")
        v_id = _seed_visit(sid_del)
        m_res = app.record_linked_measurement(sid_del, date.today(), 125.0, 25.0, "2018-01-01", "K")
        vr_res = app.record_linked_vital(sid_del, datetime.now().isoformat(timespec="minutes"), 100, 65, 80, 36.5, 92)
        with app.connect_db() as conn:
            mid = conn.execute("SELECT id FROM measurements WHERE student_id=?", (sid_del,)).fetchone()[0]
            vid_vt = conn.execute("SELECT id FROM vital_readings WHERE student_id=?", (sid_del,)).fetchone()[0]
        app.delete_visit(v_id)
        app.delete_measurement(mid)
        app.delete_vital(vid_vt)
        app.delete_student(sid_del)
        with app.connect_db() as conn:
            fk_after = conn.execute("PRAGMA foreign_key_check").fetchall()
            integ_after = conn.execute("PRAGMA integrity_check").fetchone()[0]
        rep.check(
            s,
            "Cascade deletes (visit, measurement, vital, student) pass integrity_check & zero FK orphans",
            len(fk_after) == 0 and str(integ_after).lower() == "ok" and bool(m_res) and bool(vr_res),
            f"fk_after={fk_after}, integ={integ_after}",
            severity="P0",
        )

        # Test reset_current_year_students_and_records clears active year clinical/cafeteria/diabetes tables
        _, sid_rst = _seed_class_and_student("2-C", "Reset Temp", "K", "2018-04-04")
        app.add_seizure_log(sid_rst, date.today().isoformat(), "11:00", 20, "Fokal")
        app.add_isolation_record(sid_rst, "Kızamık", date.today().isoformat())
        idm.save_diabetes_profile(lambda: app.connect_db(), sid_rst, {"is_diabetic": True})
        caf.upsert_diner(lambda: app.connect_db(), student_id=sid_rst, full_name="Reset Temp", role_type="Öğrenci", class_or_dept="2-C")
        app.reset_current_year_students_and_records()
        with app.connect_db() as conn:
            yr_rem = {
                t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                for t in ("students", "seizure_logs", "isolation_records", "diabetes_profiles", "cafeteria_diners")
            }
        rep.check(
            s,
            "reset_current_year_students_and_records clears all active-year tables cleanly",
            all(v == 0 for v in yr_rem.values()),
            f"remaining={yr_rem}",
            severity="P0",
        )

        # Test perform_factory_reset clears ALL clinical + diabetes + cafeteria + supply/fridge/menu tables
        app.add_infirmary_inventory_item("Test Bandaj", "Pansuman", 5, "paket")
        app.add_fridge_temp_log(date.today().isoformat(), "08:00", 4.5, 22.0, 45.0, "Hemşire")
        idm.upsert_school_meal_menu(lambda: app.connect_db(), date.today().isoformat(), "Öğle Yemeği", [{"name": "Çorba", "carbs_g": 20.0}])
        ok_rst, msg_rst = app.perform_factory_reset()
        with app.connect_db() as conn:
            rem_all = {}
            for t in ("students", "seizure_logs", "isolation_records", "infirmary_inventory", "fridge_temp_logs", "diabetes_profiles", "insulin_bolus_logs", "school_meal_menus", "cafeteria_diners"):
                if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (t,)).fetchone():
                    rem_all[t] = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        rep.check(
            s,
            "perform_factory_reset clears ALL tables including seizures/isolation/inventory/fridge/diabetes/menus",
            ok_rst and all(v == 0 for v in rem_all.values()),
            f"remaining={ {k: v for k, v in rem_all.items() if v} }, msg={msg_rst}",
            severity="P0",
        )


def run_suite_2(rep: AuditReport) -> None:
    print("\n=== SUITE 2: Smart Clinical Input Parsers & Edge-Case Fuzzing ===")
    s = "Suite 2"
    nan_v = float("nan")
    cases = [
        ("normalize_temperature", app.normalize_temperature, [
            ("366", 36.6), ("3665", 36.7), ("36,6", 36.6), ("38.2 C", 38.2), ("999", None), (None, None), (nan_v, None), ("-5", None), ("abc", None)
        ]),
        ("normalize_height", app.normalize_height, [
            ("1055", 105.5), ("1.45", 145.0), ("1,42", 142.0), ("0", None), (None, None), (nan_v, None), ("abc", None)
        ]),
        ("normalize_weight (child)", lambda x: app.normalize_weight(x, is_adult=False), [
            ("185", 18.5), ("18,5 kg", 18.5), ("18500", 18.5), (None, None), (nan_v, None), ("0", None)
        ]),
        ("normalize_weight (adult)", lambda x: app.normalize_weight(x, is_adult=True), [
            ("755", 75.5), ("75,5", 75.5), (None, None)
        ]),
        ("normalize_head_circ", app.normalize_head_circ, [
            ("485", 48.5), ("48,5 cm", 48.5), (None, None), ("0", None)
        ]),
        ("normalize_bp", lambda args: app.normalize_bp(*args) if isinstance(args, tuple) else app.normalize_bp(args), [
            ("120/80", (120, 80)), ("12-8", (120, 80)), ("12080", (120, 80)), ((12, 8), (120, 80)), ((80, 120), (120, 80)), (None, (None, None))
        ]),
        ("normalize_spo2", app.normalize_spo2, [
            ("988", 98), ("%97", 97), ("99", 99), ("150", None), ("9999", None), (None, None)
        ]),
        ("normalize_pulse", app.normalize_pulse, [
            ("85", 85), ("85 bpm", 85), (None, None), ("999", None)
        ]),
        ("normalize_glucose", app.normalize_glucose, [
            ("95 mg/dL", 95.0), ("110", 110.0), (None, None), ("9999", None)
        ]),
        ("normalize_phone", app.normalize_phone, [
            ("5321112233", "05321112233"), ("+90 532 111 22 33", "05321112233"), ("0532-111-22-33", "05321112233"), (None, None)
        ]),
        ("normalize_schedule_times", app.normalize_schedule_times, [
            ("930 1330", "09:30, 13:30"), ("9:30, 14.00", "09:30, 14:00"), (None, "")
        ]),
        ("normalize_dosage_pattern", app.normalize_dosage_pattern, [
            ("3*1", "3x1"), ("2X1", "2x1"), (None, "")
        ]),
        ("parse_birth_date", app.parse_birth_date, [
            ("15032019", "2019-03-15"), ("15.03.2019", "2019-03-15"), ("2019-03-15", "2019-03-15"), ("31022020", None), ("99999999", None), (None, None)
        ]),
    ]

    for label, fn, tests in cases:
        mismatches = []
        for inp, exp in tests:
            try:
                got = fn(inp)
            except Exception as exc:
                mismatches.append((inp, f"EXC:{exc}", exp))
                continue
            if exp is None:
                ok = got is None
            elif isinstance(exp, float):
                ok = (got is not None and abs(float(got) - exp) < 0.05)
            else:
                ok = got == exp
            if not ok:
                mismatches.append((inp, got, exp))
        rep.check(s, label, len(mismatches) == 0, f"mismatches={mismatches}" if mismatches else f"{len(tests)}/{len(tests)} ok", severity="P1")


def run_suite_3(rep: AuditReport) -> None:
    print("\n=== SUITE 3: WHO LMS Anthropometry, Z-Scores, Percentiles & Ideal Weight Engine ===")
    s = "Suite 3"
    with sandbox_env():
        app.initialize_db()
        lms = app._ensure_lms_cache()
        rep.check(s, "_ensure_lms_cache populated", len(lms) > 200, f"rows={len(lms)}", severity="P0")
        row_36 = app.get_lms_row("weight", "E", 36)
        rep.check(s, "get_lms_row('weight','E',36)", row_36 is not None and row_36["m"] > 0, str(row_36))

        # z_score & Box-Cox limit + zero-safety
        z_mid = app.z_score(14.3, row_36["l"], row_36["m"], row_36["s"])
        z_l0 = app.z_score(15.0, 0.0, 15.0, 0.1)
        p_l0 = app.percentile(z_l0)
        rep.check(s, "z_score L!=0 & L==0 -> percentile 50%", math.isfinite(z_mid) and abs(z_l0) < 1e-6 and abs(p_l0 - 50.0) < 0.1)

        # Pediatric 0-60 months (with HC), 61-228 months (no HC), adult (>228 months)
        ev_child36 = app.evaluate_measurement(14.3, "weight", "E", 36)
        ev_hc36 = app.evaluate_measurement(49.5, "head_circ", "E", 36)
        ev_teen = app.evaluate_measurement(155.0, "height", "K", 156)
        ev_adult = app.evaluate_measurement(70.0, "weight", "E", 360)
        rep.check(
            s,
            "evaluate_measurement 36mo / head_circ / 156mo / adult",
            ev_child36 is not None and ev_hc36 is not None and ev_teen is not None and ev_adult is None,
            f"weight36={bool(ev_child36)}, hc36={bool(ev_hc36)}, teen={bool(ev_teen)}, adult_none={ev_adult is None}",
        )

        # Partial measurements + historical fallback & growth summary table
        iw = app.calculate_ideal_weight(120.0, "E", 84, 23.0)
        tbl_u5 = app.get_clinical_growth_summary_table(100.0, 16.0, 16.0, "E", 48, head_circ_cm=49.5)
        tbl_school = app.get_clinical_growth_summary_table(140.0, 34.0, 17.35, "K", 120, head_circ_cm=None)
        tbl_adult = app.get_clinical_growth_summary_table(175.0, 72.0, 23.51, "E", 360)
        rep.check(s, "calculate_ideal_weight + clinical growth tables (u5/school/adult)", bool(iw) and not tbl_u5.empty and not tbl_school.empty and not tbl_adult.empty)

        curves = app.generate_who_percentile_curves("weight", "E", 0, 60)
        cutoffs = app.get_who_reference_cutoffs_table("E", 48, 100.0, 16.0, 16.0, 49.5)
        rep.check(s, "generate_who_percentile_curves & cutoffs", not curves.empty and not cutoffs.empty)

        # Partial measurements enrichment (height-only day 1, weight-only day 2, complete day 3)
        df_m = pd.DataFrame([
            {"measured_on": "2026-09-10", "height_cm": 120.0, "weight_kg": None, "head_circ_cm": None},
            {"measured_on": "2026-09-15", "height_cm": None, "weight_kg": 24.0, "head_circ_cm": None},
            {"measured_on": "2026-09-20", "height_cm": 121.0, "weight_kg": 24.5, "head_circ_cm": None},
        ])
        df_enr = app.enrich_measurements_with_percentiles(df_m, "E", "2019-09-10")
        rep.check(
            s,
            "enrich_measurements_with_percentiles handles partial & full rows safely",
            (
                len(df_enr) == 3
                and pd.notna(df_enr.iloc[0]["Boy_Z"])
                and pd.isna(df_enr.iloc[0]["BMI"])
                and pd.notna(df_enr.iloc[1]["Kilo_Z"])
                and pd.notna(df_enr.iloc[2]["BMI"])
                and pd.notna(df_enr.iloc[2]["BMI_Z"])
            ),
            f"cols_ok={'BMI' in df_enr.columns}, bmi_cnt={int(df_enr['BMI'].notna().sum())}",
        )


def run_suite_4(rep: AuditReport) -> None:
    print("\n=== SUITE 4: Cross-Module Data Propagation & Bidirectional Sync ===")
    s = "Suite 4"
    with sandbox_env():
        app.initialize_db()
        _, sid = _seed_class_and_student("2-A", "Zeynep Ak", "K", "2019-03-10", guardian_name="Hasan Ak", guardian_phone="0533 222 33 44")
        today_s = date.today().isoformat()

        # 1. save_school_screening -> measurements + vital_readings + health_profiles, NO duplicate on re-save
        app.save_school_screening(sid, today_s, height_cm=118.0, weight_kg=22.0, temperature_c=36.7, pulse_bpm=88, spo2_pct=98, wears_glasses="Evet")
        app.save_school_screening(sid, today_s, height_cm=119.0, weight_kg=22.5, temperature_c=36.8, pulse_bpm=86, spo2_pct=99, wears_glasses="Evet")
        with app.connect_db() as conn:
            m_cnt = conn.execute("SELECT COUNT(*) FROM measurements WHERE student_id=?", (sid,)).fetchone()[0]
            m_row = conn.execute("SELECT height_cm, weight_kg, bmi, age_months FROM measurements WHERE student_id=?", (sid,)).fetchone()
            v_cnt = conn.execute("SELECT COUNT(*) FROM vital_readings WHERE student_id=?", (sid,)).fetchone()[0]
            hp = conn.execute("SELECT * FROM health_profiles WHERE student_id=?", (sid,)).fetchone()
        rep.check(
            s,
            "save_school_screening upserts measurements & vitals cleanly without duplicates",
            m_cnt == 1 and abs(float(m_row["height_cm"]) - 119.0) < 0.01 and float(m_row["bmi"]) > 0 and v_cnt == 1 and hp is not None and hp["wears_glasses"] == "Evet",
            f"m_cnt={m_cnt}, v_cnt={v_cnt}, bmi={m_row['bmi'] if m_row else None}",
            severity="P0",
        )

        # 2. record_linked_measurement partial completion & record_linked_vital same-slot upsert
        app.record_linked_measurement(sid, date.today() + timedelta(days=1), 120.0, None, "2019-03-10", "K")
        app.record_linked_measurement(sid, date.today() + timedelta(days=1), None, 23.0, "2019-03-10", "K")
        slot_ts = f"{today_s} 11:30"
        slot_norm = slot_ts.replace(" ", "T")
        app.record_linked_vital(sid, slot_ts, 100, 65, None, None, None)
        app.record_linked_vital(sid, slot_ts, None, None, 90, 36.6, None, spo2=95)
        with app.connect_db() as conn:
            m2 = conn.execute("SELECT height_cm, weight_kg, bmi FROM measurements WHERE student_id=? AND measured_on=?", (sid, (date.today() + timedelta(days=1)).isoformat())).fetchone()
            v2 = conn.execute("SELECT systolic_mmhg, diastolic_mmhg, pulse_bpm, temperature_c, spo2_pct FROM vital_readings WHERE student_id=? AND measured_at=?", (sid, slot_norm)).fetchone()
        rep.check(
            s,
            "record_linked_measurement & record_linked_vital complete partial entries on same slot",
            m2 is not None and m2["height_cm"] == 120.0 and m2["weight_kg"] == 23.0 and v2 is not None and v2["systolic_mmhg"] == 100 and v2["pulse_bpm"] == 90 and v2["spo2_pct"] == 95,
            f"m2={dict(m2) if m2 else None}, v2={dict(v2) if v2 else None}",
            severity="P0",
        )

        # 3. Bidirectional sync guardian fields
        app._upsert_health_profile_fields(sid, {"guardian_name": "Murat Ak", "guardian_phone": "0544 999 88 77"})
        with app.connect_db() as conn:
            stu_row = conn.execute("SELECT guardian_name, guardian_phone FROM students WHERE id=?", (sid,)).fetchone()
        rep.check(s, "Bidirectional sync health_profiles -> students", stu_row["guardian_name"] == "Murat Ak", str(dict(stu_row)))

        # 4. JSON file save/load/apply roundtrip including seizures & isolation
        _seed_visit(sid, today_s, "09:10", "Öksürük", "Aileye bilgi verildi", "İzlem")
        app.save_student_vaccination(sid, "KKK_1", True, "2020-04-01", "ASM", "LOT-1")
        app.add_seizure_log(sid, today_s, "12:00", 30, "Absans")
        app.add_isolation_record(sid, "İnfluenza", today_s)
        jpath = app.save_person_data_file(sid)
        jdata = app.load_person_data_file(sid)
        with app.connect_db() as conn:
            for t in ("visits", "measurements", "vital_readings", "school_screenings", "student_vaccinations", "seizure_logs", "isolation_records"):
                conn.execute(f"DELETE FROM {t} WHERE student_id=?", (sid,))
        app.apply_person_data_file_to_db(sid, jdata)
        with app.connect_db() as conn:
            counts = {t: conn.execute(f"SELECT COUNT(*) FROM {t} WHERE student_id=?", (sid,)).fetchone()[0]
                      for t in ("visits", "measurements", "vital_readings", "school_screenings", "student_vaccinations", "seizure_logs", "isolation_records")}
        rep.check(
            s,
            "JSON save -> clear -> apply_person_data_file_to_db round-trip fidelity",
            jpath is not None and all(v >= 1 for v in counts.values()),
            f"restored_counts={counts}",
            severity="P0",
        )


def run_suite_5(rep: AuditReport) -> None:
    print("\n=== SUITE 5: Medication Inventory, Dose Logging & Infirmary Supplies ===")
    s = "Suite 5"
    with sandbox_env():
        app.initialize_db()
        _, sid = _seed_class_and_student("3-A", "Efe Yıldız", "E", "2018-02-02")
        # Find next Monday & Saturday
        d_mon = date.today()
        while d_mon.weekday() != 0:
            d_mon += timedelta(days=1)
        d_sat = d_mon + timedelta(days=5)

        mid = app.add_student_medication(
            sid, "Parol Şurup", total_volume=100.0, dose_amount=5.0,
            scheduled_times="09:30, 14:00", start_date=(d_mon - timedelta(days=7)).isoformat(),
            duration_days=30, active_days="Sadece Okul Günleri (Pazartesi - Cuma)",
        )
        doses_mon = app.get_scheduled_doses_for_date(d_mon)
        doses_sat = app.get_scheduled_doses_for_date(d_sat)
        rep.check(s, "Weekday vs weekend schedule filtering", len(doses_mon) == 2 and len(doses_sat) == 0, f"mon={len(doses_mon)}, sat={len(doses_sat)}")

        app.log_medication_dose(mid, sid, d_mon.isoformat(), "09:30", "09:35", 5.0, "ml", status="Verildi")
        with app.connect_db() as conn:
            rem1 = conn.execute("SELECT remaining_volume FROM student_medications WHERE id=?", (mid,)).fetchone()[0]
        # Edge case: change log status from Verildi -> Atlandı (should restore 5.0!)
        app.log_medication_dose(mid, sid, d_mon.isoformat(), "09:30", "09:40", 5.0, "ml", status="Atlandı")
        with app.connect_db() as conn:
            rem_after_skip = conn.execute("SELECT remaining_volume FROM student_medications WHERE id=?", (mid,)).fetchone()[0]
        rep.check(s, "log_medication_dose decrements and handles Verildi->Atlandı volume restoration", rem1 == 95.0 and rem_after_skip == 100.0, f"after_give={rem1}, after_skip={rem_after_skip}", severity="P1")

        # Give twice & undo, test upper clamp <= total_volume
        app.log_medication_dose(mid, sid, d_mon.isoformat(), "14:00", "14:05", 5.0, "ml", status="Verildi")
        app.undo_medication_dose(mid, d_mon.isoformat(), "14:00")
        app.update_student_medication(med_id=mid, remaining_volume=98.0)
        app.log_medication_dose(mid, sid, d_mon.isoformat(), "14:00", "14:05", 5.0, "ml", status="Verildi")
        with app.connect_db() as conn:
            rem_b = conn.execute("SELECT remaining_volume FROM student_medications WHERE id=?", (mid,)).fetchone()[0]
        app.update_student_medication(med_id=mid, remaining_volume=99.0)
        app.undo_medication_dose(mid, d_mon.isoformat(), "14:00")
        with app.connect_db() as conn:
            rem_clamped = conn.execute("SELECT remaining_volume FROM student_medications WHERE id=?", (mid,)).fetchone()[0]
        rep.check(s, "undo_medication_dose restores without exceeding total_volume", rem_clamped <= 100.0 and rem_b >= 0, f"rem_b={rem_b}, clamped={rem_clamped}")

        app.delete_student_medication(mid)
        with app.connect_db() as conn:
            med_left = conn.execute("SELECT COUNT(*) FROM student_medications WHERE id=?", (mid,)).fetchone()[0]
        rep.check(s, "delete_student_medication removes record", med_left == 0)

        # Infirmary supply inventory & cold-chain fridge logs
        iid = app.add_infirmary_inventory_item("Steril Gazlı Bez", "Pansuman", 2.0, "paket", min_threshold=5.0, expiry_date=(date.today() + timedelta(days=5)).isoformat())
        inv_items = app.get_infirmary_inventory()
        alerts = app.get_infirmary_inventory_alerts()
        app.delete_infirmary_inventory_item(iid)
        fid_ok = app.add_fridge_temp_log(date.today().isoformat(), "08:30", 5.2, 22.0, 45.0, "Hemşire")
        fid_bad = app.add_fridge_temp_log(date.today().isoformat(), "15:30", 9.5, 24.0, 48.0, "Hemşire")
        flogs = app.get_fridge_temp_logs()
        rep.check(
            s,
            "Infirmary inventory alerts & fridge cold-chain temperature alerts",
            len(inv_items) >= 1 and len(alerts["critical_stock"]) >= 1 and len(alerts["expiring_soon"]) >= 1 and len(flogs) == 2 and not app.is_fridge_temp_alert(5.2) and app.is_fridge_temp_alert(9.5),
            f"alerts={ {k: len(v) for k, v in alerts.items()} }",
        )
        app.delete_fridge_temp_log(fid_ok)
        app.delete_fridge_temp_log(fid_bad)


def run_suite_6(rep: AuditReport) -> None:
    print("\n=== SUITE 6: Vaccination Schedule, Outbreak Radar, Seizure Logs & Follow-Up Queue ===")
    s = "Suite 6"
    with sandbox_env():
        app.initialize_db()
        _, sid1 = _seed_class_and_student("1-B", "Elif Çetin", "K", "2019-06-15")
        _, sid2 = _seed_class_and_student("1-B", "Can Öztürk", "E", "2019-08-20")
        _, sid3 = _seed_class_and_student("1-B", "Ece Koç", "K", "2019-09-01")

        filled = app.quick_fill_routine_vaccinations_for_student(sid1, "2019-06-15", up_to_age=True)
        v_sum = app.get_student_vaccination_summary(sid1, "2019-06-15")
        v_html = app.build_vaccination_matrix_html(sid1, "2019-06-15")
        rep.check(s, "Vaccinations quick-fill, summary & matrix HTML", filled > 0 and v_sum["completed_count"] >= filled and "<table" in v_html.lower(), f"v_sum={v_sum}")

        # Follow-up queue
        v_open = _seed_visit(sid1, complaint="Yüksek Ateş", referral="Hastaneye Sevk", follow_up="Veli aranacak")
        pending_fup = app.get_pending_followup_visits()
        app.close_visit_followup(v_open, "Veli teslim aldı, kontrol tamam.", "Tamamlandı")
        pending_after = app.get_pending_followup_visits()
        rep.check(s, "Pending follow-up queue & close_visit_followup", any(x["id"] == v_open for x in pending_fup) and not any(x["id"] == v_open for x in pending_after))

        # Outbreak detection clustering (3 fever/flu cases within window)
        d0 = date.today().isoformat()
        _seed_visit(sid1, d0, "09:00", "Ateş ve Öksürük")
        _seed_visit(sid2, d0, "10:00", "Ateş ve Boğaz Ağrısı")
        _seed_visit(sid3, d0, "11:00", "Yüksek Ateş / Grip")
        outbreaks = app.detect_school_outbreaks(days_window=5, min_cases=3)
        rep.check(s, "detect_school_outbreaks detects fever/respiratory cluster", len(outbreaks) >= 1, f"clusters={len(outbreaks)}")

        # Isolation & seizures
        iso_id = app.add_isolation_record(sid2, "Kızamıkçık", d0)
        app.close_isolation_record(iso_id, d0, clearance_received=True, notes_append="Rapor sunuldu")
        active_iso = app.get_isolation_records(active_only=True)
        closed_iso = app.get_isolation_records(active_only=False)
        sz_id = app.add_seizure_log(sid1, d0, "13:15", 90, "Tonik-Klonik", emergency_med_given="Rektal Diazepam", ambulance_called=True, guardian_notified=True)
        sz_logs = app.get_student_seizure_logs(sid1)
        rep.check(s, "Isolation lifecycle & seizure log CRUD", len(active_iso) == 0 and any(r["id"] == iso_id for r in closed_iso) and len(sz_logs) == 1 and sz_logs[0]["id"] == sz_id)


def run_suite_7(rep: AuditReport) -> None:
    print("\n=== SUITE 7: Document & Report Exporters (.docx, .xlsx, .html, WhatsApp) ===")
    s = "Suite 7"
    with sandbox_env():
        app.initialize_db()
        _, sid = _seed_class_and_student("4-B", "Deniz Aydın", "K", "2017-01-12", guardian_phone="0532 555 66 77")
        vid = _seed_visit(sid, complaint="Karın ağrısı", referral="Veli bilgilendirildi")
        app.quick_fill_routine_vaccinations_for_student(sid, "2017-01-12")
        app.save_school_screening(sid, date.today().isoformat(), height_cm=134.0, weight_kg=30.0)
        med_id = app.add_student_medication(sid, "Ventolin İnhaler", med_type="İnhaler / Sprey", dose_amount=2.0)
        app.add_seizure_log(sid, date.today().isoformat(), "09:45", 40, "Fokal")
        app.add_infirmary_inventory_item("Flaster", "Pansuman", 10, "adet")
        app.add_fridge_temp_log(date.today().isoformat(), "09:00", 4.5, 21.0, 40.0)

        docx_checks = [
            ("build_vaccination_card_docx", lambda: app.build_vaccination_card_docx(sid)),
            ("build_visit_official_document_docx('tutanak')", lambda: app.build_visit_official_document_docx(vid, "tutanak")),
            ("build_visit_official_document_docx('veli_fisi')", lambda: app.build_visit_official_document_docx(vid, "veli_fisi")),
            ("build_medication_consent_docx", lambda: app.build_medication_consent_docx(med_id=med_id, student_id=sid)),
            ("build_seizure_log_docx", lambda: app.build_seizure_log_docx(sid)),
            ("build_infirmary_inventory_docx", lambda: app.build_infirmary_inventory_docx()),
            ("build_fridge_temp_monthly_docx", lambda: app.build_fridge_temp_monthly_docx()),
        ]
        for label, fn in docx_checks:
            try:
                raw = fn()
                ok_zip = isinstance(raw, (bytes, bytearray)) and raw[:4] == b"PK\x03\x04"
                if ok_zip and docx is not None:
                    doc = docx.Document(io.BytesIO(raw))
                    ok_zip = len(doc.paragraphs) + len(doc.tables) > 0
                rep.check(s, label, ok_zip, f"bytes={len(raw) if raw else 0}", severity="P1")
            except Exception as exc:
                rep.check(s, label, False, f"EXC: {exc}", severity="P1")

        # Excel exports
        stus = app.get_students()
        with app.connect_db() as c:
            v_df = pd.read_sql_query("SELECT * FROM visits", c)
            m_df = pd.read_sql_query("SELECT * FROM measurements", c)
            vt_df = pd.read_sql_query("SELECT * FROM vital_readings", c)
        xl_checks = [
            ("workbook_download", lambda: app.workbook_download(stus, v_df, m_df, vt_df)),
            ("make_roster_template", app.make_roster_template),
            ("make_eokul_dogum_ve_yas_template", app.make_eokul_dogum_ve_yas_template),
            ("make_vital_reference_template", app.make_vital_reference_template),
        ]
        for label, fn in xl_checks:
            try:
                raw = fn()
                ok_xl = isinstance(raw, (bytes, bytearray)) and raw[:4] == b"PK\x03\x04"
                if ok_xl:
                    wb = openpyxl.load_workbook(io.BytesIO(raw))
                    ok_xl = len(wb.sheetnames) >= 1
                rep.check(s, label, ok_xl, f"bytes={len(raw) if raw else 0}", severity="P1")
            except Exception as exc:
                rep.check(s, label, False, f"EXC: {exc}", severity="P1")

        # Dynamic screening reports + HTML + WhatsApp URL
        df_dyn = app.build_dynamic_screening_report_df(selected_dates=[date.today().isoformat()])
        dyn_docx = app.build_dynamic_screening_docx(df_dyn)
        dyn_xl = app.build_dynamic_screening_excel(df_dyn)
        dyn_html = app.build_dynamic_screening_html(df_dyn)
        vcard_html = app.build_vaccination_card_html(sid)
        wa_url = app.generate_whatsapp_guardian_link(sid, vid)
        rep.check(s, "Dynamic screening builders (df/docx/xlsx/html)", not df_dyn.empty and dyn_docx[:4] == b"PK\x03\x04" and dyn_xl[:4] == b"PK\x03\x04" and "<html" in dyn_html.lower())
        rep.check(s, "build_vaccination_card_html & WhatsApp URL", "<html" in vcard_html.lower() and bool(wa_url and wa_url.startswith("https://wa.me/905325556677")))


def run_suite_8(rep: AuditReport) -> None:
    print("\n=== SUITE 8: Backup, Restore, KVKK Masking & Frozen .EXE Path Resolution ===")
    s = "Suite 8"
    with sandbox_env() as tmp_p:
        app.initialize_db()
        _, sid = _seed_class_and_student("4-C", "Canan Yıldırım", "K", "2017-04-04", guardian_name="Hasan Yıldırım", guardian_phone="0532 777 88 99")

        # Daily backups + pruning
        b1 = app.run_daily_auto_backup(max_backups=2, force=True, backup_date=date(2026, 9, 1))
        b2 = app.run_daily_auto_backup(max_backups=2, force=True, backup_date=date(2026, 9, 2))
        b3 = app.run_daily_auto_backup(max_backups=2, force=True, backup_date=date(2026, 9, 3))
        lst = app.list_auto_backups()
        rep.check(s, "run_daily_auto_backup + _prune_auto_backups(max=2)", b1 is not None and b2 is not None and b3 is not None and len(lst) == 2, f"backups={len(lst)}")

        # Corrupted restore rejection + valid restore recovery
        corrupt_p = tmp_p / "corrupt.db"
        corrupt_p.write_bytes(b"NOT_SQLITE_GARBAGE")
        rejected_corrupt = False
        try:
            app.restore_auto_backup_file(corrupt_p)
        except ValueError:
            rejected_corrupt = True
        app.delete_student(sid)
        with app.connect_db() as c:
            before_cnt = c.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        ok_restored = app.restore_auto_backup_file(b3)
        with app.connect_db() as c:
            after_cnt = c.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        rep.check(s, "restore_auto_backup_file rejects corrupt file & restores valid backup", rejected_corrupt and before_cnt == 0 and ok_restored and after_cnt == 1, severity="P0")

        # KVKK masking ON / OFF
        revir_ui.session_state["kvkk_privacy_mode"] = True
        masked_nm = app.mask_kvkk_value("Canan Yıldırım", "name")
        masked_ph = app.mask_kvkk_value("0532 777 88 99", "phone")
        df_ros = app.get_students_enriched()
        df_masked = app.apply_kvkk_mask_to_roster_df(df_ros)
        revir_ui.session_state["kvkk_privacy_mode"] = False
        unmasked_nm = app.mask_kvkk_value("Canan Yıldırım", "name")
        rep.check(
            s,
            "KVKK masking toggles cleanly ON/OFF",
            ("*" in masked_nm and "*" in masked_ph and unmasked_nm == "Canan Yıldırım" and "*" in str(df_masked.iloc[0]["Ad Soyad"])),
            f"masked={masked_nm}, ph={masked_ph}",
        )

        # Frozen .EXE & _MEIPASS simulation
        mock_meipass = tmp_p / "frozen_meipass"
        mock_meipass.mkdir(parents=True, exist_ok=True)
        (mock_meipass / "frozen_only_asset.txt").write_text("ok", encoding="utf-8")
        with mock.patch.object(sys, "frozen", True, create=True), mock.patch.object(sys, "_MEIPASS", str(mock_meipass), create=True):
            resolved_frozen = app._resolve_bundled_asset("frozen_only_asset.txt")
        rep.check(
            s,
            "_resolve_bundled_asset resolves runtime sys._MEIPASS dynamically when frozen=True",
            resolved_frozen.exists() and str(mock_meipass) in str(resolved_frozen),
            f"resolved={resolved_frozen}",
            severity="P1",
        )


def run_suite_9(rep: AuditReport) -> None:
    print("\n=== SUITE 9: Native revir_ui Headless Render Across ALL 15 PAGES & KVKK Modes ===")
    s = "Suite 9"
    # Verify architectural reality: streamlit is NOT imported
    uses_st = "streamlit" in sys.modules or "import streamlit" in (PROJECT_ROOT / "persantil.py").read_text(encoding="utf-8")
    rep.warn(s, "UI Engine Verification (Prompt mentioned Streamlit AppTest)", f"Codebase migrated from Streamlit to native revir_ui (streamlit_imported={uses_st}); executing revir_ui.run_action_headless across all {len(app.PAGES)} pages.", severity="P2")

    with sandbox_env():
        app.initialize_db()
        _, sid = _seed_class_and_student("4-A", "Kemal Şahin", "E", "2017-05-20", student_no="44")
        _seed_visit(sid, date.today().isoformat(), "10:00", "Baş ağrısı")
        app.save_school_screening(sid, date.today().isoformat(), height_cm=130.0, weight_kg=28.0)

        for active_val in (None, sid):
            for kvkk_on in (False, True):
                revir_ui.session_state["active_person_id"] = active_val
                revir_ui.session_state["kvkk_privacy_mode"] = kvkk_on
                for pg in app.PAGES:
                    revir_ui.session_state["main_menu_page"] = pg
                    try:
                        out = revir_ui.run_action_headless(app.main, {"action": "render"})
                        html_out = str(out.get("main_html") or "")
                        err = ("Çalışma Zamanı Hatası" in html_out) or ("Traceback" in html_out)
                        rep.check(
                            s,
                            f"Page render '{pg}' (active={active_val}, kvkk={kvkk_on})",
                            not err and len(html_out) > 100,
                            f"len={len(html_out)}" if not err else "RuntimeError in page HTML!",
                            severity="P0" if err else "P1",
                        )
                    except Exception as exc:
                        rep.check(s, f"Page render '{pg}' (active={active_val}, kvkk={kvkk_on})", False, f"EXC: {exc}", severity="P0")


def run_suite_10(rep: AuditReport) -> None:
    print("\n=== SUITE 10: Desktop App, Build/Version Scripts & Companion Modules ===")
    s = "Suite 10"
    import app_navigation
    import report_settings_ui
    import report_workflow_ui
    with sandbox_env() as tmp_p:
        ok_w = desktop_app._is_dir_writable(tmp_p)
        lock_ok = desktop_app.acquire_single_instance_lock(tmp_p)
        desktop_app.release_single_instance_lock()
        ver_str = desktop_app._load_app_version()
        rep.check(s, "desktop_app writable check, single-instance lock & version", ok_w and lock_ok and bool(ver_str), f"ver={ver_str}")
        rep.check(
            s,
            "v1.35 companion modules (app_navigation, report_settings_ui, report_workflow_ui)",
            len(app_navigation.PAGES) == len(app.PAGES) and callable(report_settings_ui.render_report_settings_panel) and callable(report_workflow_ui.render_confirmation_summary),
            f"pages={len(app_navigation.PAGES)}",
        )


def main() -> int:
    rep = AuditReport()
    for runner in (run_suite_1, run_suite_2, run_suite_3, run_suite_4, run_suite_5, run_suite_6, run_suite_7, run_suite_8, run_suite_9, run_suite_10):
        try:
            runner(rep)
        except Exception:
            traceback.print_exc()
            rep.check(runner.__name__, "unhandled suite crash", False, traceback.format_exc()[-300:], severity="P0")

    p, w, f = rep.summary()
    print(f"\n==================== AUDIT SUMMARY ====================")
    print(f"PASS: {p} | WARN: {w} | FAIL: {f}")
    for r in rep.results:
        if r["status"] in ("FAIL", "WARN"):
            print(f" - [{r['status']} {r['severity']}] {r['suite']} :: {r['name']} -> {r['detail']}")
    return 1 if f > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
