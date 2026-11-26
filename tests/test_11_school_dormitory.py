"""Comprehensive tests for Okul Yurdu / Pansiyon Yönetimi (school_dormitory.py + persantil integration)."""
from __future__ import annotations

import io
import json
import sys
import unittest
from datetime import date, timedelta
from pathlib import Path

import docx
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import insulin_diabetes as idm
import persantil
import revir_ui
import school_dormitory as sdorm
from tests.db_test_helper import isolated_real_db


def _docx_text(raw: bytes) -> str:
    d = docx.Document(io.BytesIO(raw))
    out = [p.text for p in d.paragraphs if p.text]
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                if c.text:
                    out.append(c.text)
    return "\n".join(out)


class TestSchoolDormitoryModule(unittest.TestCase):
    def setUp(self):
        revir_ui.reset_session()
        self._ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self._ctx.__enter__()
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('9-A')")
            cid = int(c.execute("SELECT id FROM classes WHERE name='9-A'").fetchone()["id"])
            c.execute(
                "INSERT INTO students(full_name, sex, student_no, class_id, person_type, guardian_name, guardian_phone) VALUES ('Ali Kaya', 'E', '101', ?, 'Öğrenci', 'Veli Kaya', '05321112233')",
                (cid,),
            )
            self.s1 = int(c.execute("SELECT id FROM students WHERE full_name='Ali Kaya'").fetchone()["id"])
            c.execute(
                "INSERT INTO students(full_name, sex, student_no, class_id, person_type) VALUES ('Mert Demir', 'E', '102', ?, 'Öğrenci')",
                (cid,),
            )
            self.s2 = int(c.execute("SELECT id FROM students WHERE full_name='Mert Demir'").fetchone()["id"])
            c.execute(
                "INSERT INTO health_profiles(student_id, allergies, chronic_conditions, seizure_emergency_protocol, updated_at) VALUES (?, 'Yer fıstığı', 'Epilepsi', 'Yan yatır, >5 dk 112 ara', '2026-11-01')",
                (self.s1,),
            )
            idm.ensure_diabetes_tables(c)
            c.execute(
                "INSERT INTO diabetes_profiles(student_id, is_diabetic, diabetes_type, emergency_protocol, updated_at) VALUES (?, 1, 'Tip 1 Diyabet', '15g meyve suyu ver', '2026-11-01')",
                (self.s1,),
            )
            c.commit()

    def tearDown(self):
        self._ctx.__exit__(None, None, None)

    def test_own_and_external_student_enrollment_room_map_and_contact_clustering(self):
        # 1. Enroll existing school students into Room 204
        cnt = sdorm.enroll_existing_school_students_to_dorm(
            self.connect,
            [self.s1, self.s2],
            block_name="A Blok",
            floor_no="2. Kat",
            room_no="204",
        )
        self.assertEqual(cnt, 2)

        # 2. Enroll external school student into adjacent Room 205 on same floor
        ext_rid = sdorm.enroll_external_school_student_to_dorm(
            self.connect,
            full_name="Zeynep Aksoy",
            origin_school_name="Cumhuriyet Fen Lisesi",
            sex="K",
            class_name="10-B",
            student_no="889",
            guardian_name="Ayşe Aksoy",
            guardian_phone="05440001122",
            home_location="Konya / Ereğli",
            block_name="A Blok",
            floor_no="2. Kat",
            room_no="205",
            bed_bunk_no="Alt-1",
            uses_devices="Gözlük",
            chronic_conditions="Astım",
        )
        self.assertGreater(ext_rid, 0)

        residents = sdorm.get_dorm_residents(self.connect)
        self.assertEqual(len(residents), 3)
        by_nm = {r["full_name"]: r for r in residents}
        self.assertEqual(by_nm["Zeynep Aksoy"]["resident_source"], "Başka Okul Öğrencisi")
        self.assertEqual(by_nm["Zeynep Aksoy"]["origin_school_name"], "Cumhuriyet Fen Lisesi")
        self.assertIsNotNone(by_nm["Zeynep Aksoy"]["student_id"])

        # 3. Contagious contact tracing: Ali Kaya (Room 204) -> Mert Demir (1st degree roommate), Zeynep Aksoy (2nd degree adjacent room 205)
        ali_rid = int(by_nm["Ali Kaya"]["id"])
        contacts = sdorm.get_room_and_neighbor_contacts(self.connect, ali_rid)
        self.assertEqual([r["full_name"] for r in contacts["roommates"]], ["Mert Demir"])
        self.assertIn("1. Derece", contacts["roommates"][0]["contact_degree"])
        self.assertEqual([r["full_name"] for r in contacts["floor_neighbors"]], ["Zeynep Aksoy"])
        self.assertIn("2. Derece Komşu Oda", contacts["floor_neighbors"][0]["contact_degree"])

        # Recent contagious visit auto-cluster detection
        today_s = date.today().isoformat()
        with self.connect() as c:
            c.execute(
                "INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms) VALUES (?, ?, '21:10', 'Kaşıntı, uyuz şüphesi', 'Parmak aralarında döküntü')",
                (self.s1, today_s),
            )
            c.commit()
        cl = sdorm.detect_recent_contagious_room_clusters(self.connect, within_days=7)
        self.assertEqual(len(cl), 1)
        self.assertEqual(cl[0]["roommate_count"], 1)
        self.assertEqual(cl[0]["neighbor_count"], 1)

    def test_observation_shift_handover_112_meds_night_cards_food_alarm_and_meb_inspection(self):
        sdorm.enroll_existing_school_students_to_dorm(self.connect, [self.s1, self.s2], block_name="A Blok", floor_no="2. Kat", room_no="204")
        ext_rid = sdorm.enroll_external_school_student_to_dorm(
            self.connect, full_name="Elif Yılmaz", origin_school_name="Atatürk Lisesi", sex="K",
            block_name="B Blok", floor_no="1. Kat", room_no="102",
        )
        residents = {r["full_name"]: int(r["id"]) for r in sdorm.get_dorm_residents(self.connect)}
        ali_rid, mert_rid = residents["Ali Kaya"], residents["Mert Demir"]
        today_s = date.today().isoformat()

        # 1. Observation bed admission + hourly vital + discharge
        stay_id = sdorm.admit_to_observation_bed(
            self.connect, resident_id=ali_rid, bed_label="Yatak 1",
            admitted_date=today_s, admitted_time="22:15", complaint="Bulantı, kusma ve ateş",
            admitted_by="Gece Hemşiresi",
        )
        sdorm.add_observation_vital_log(self.connect, stay_id=stay_id, recorded_at=f"{today_s} 23:00", temperature_c=38.2, pulse_bpm=96)

        # 2. Foodborne outbreak early red alarm (2nd student Mert also has vomiting/diarrhea within 2h!)
        with self.connect() as c:
            c.execute(
                "INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms) VALUES (?, ?, '23:20', 'Şiddetli kusma ve ishal', 'Karın ağrısı')",
                (self.s2, today_s),
            )
            c.commit()
        outbreak = sdorm.detect_foodborne_outbreak_clusters(self.connect, target_date=today_s)
        self.assertTrue(outbreak["alarm"])
        self.assertEqual(outbreak["case_count"], 2)

        # 3. 112 Emergency transfer with 3 rest days -> auto creates rest exemption & due alert check
        tr_id = sdorm.record_emergency_transfer(
            self.connect, resident_id=mert_rid, transfer_date=today_s, transfer_time="23:45",
            reason_complaint="Akut gastroenterit / dehidratasyon", transport_mode="112 Ambulans",
            hospital_name="Şehir Hastanesi Çocuk Acil", escort_staff_name="Belletmen Kemal",
            guardian_notified_time="23:50", return_time="03:10", epicrisis_summary="IV sıvı + 2 gün istirahat",
            rest_days_given=1,
        )
        tr_docx = _docx_text(sdorm.build_emergency_transfer_docx(self.connect, tr_id, school_name="Test Pansiyonu"))
        self.assertIn("Şehir Hastanesi Çocuk Acil", tr_docx)
        self.assertIn("Belletmen Kemal", tr_docx)

        rest_docx = _docx_text(sdorm.build_rest_and_exemption_list_docx(self.connect, ref_date=today_s))
        self.assertIn("Mert Demir", rest_docx)

        # 4. Auto night handover summary + docx
        auto_sum = sdorm.build_auto_night_handover_summary(self.connect, today_s)
        self.assertIn("Ali Kaya", auto_sum)
        self.assertIn("Mert Demir", auto_sum)
        ho_id = sdorm.save_shift_handover(
            self.connect, shift_date=today_s, shift_type="17:00 - 08:00 Gece Nöbeti",
            from_staff="Hemşire A", to_staff="Belletmen B", summary_text=auto_sum,
        )
        self.assertIn("Hemşire A", _docx_text(sdorm.build_shift_handover_docx(self.connect, ho_id)))

        sdorm.discharge_observation_stay(self.connect, stay_id, outcome_status="İyileşti — Odasına Döndü")

        # 5. Chronic Night Emergency Room Card (.docx) includes Ali Kaya's Type 1 Diabetes + Epilepsy room instructions
        cards = sdorm.get_chronic_night_risk_residents(self.connect)
        self.assertTrue(any(c["full_name"] == "Ali Kaya" and "Tip 1 Diyabet" in c["risk_tags"] and "Epilepsi" in c["risk_tags"] for c in cards))
        cards_docx = _docx_text(sdorm.build_night_emergency_room_cards_docx(self.connect, school_name="Test Okulu"))
        self.assertIn("Oda 204", cards_docx)
        self.assertIn("15g meyve suyu", cards_docx)

        # 6. Seed sanitation, food sample, staff portör, room med inspection & evaluate MEB Inspection score + Excel/Docx
        with self.connect() as c:
            c.execute("INSERT INTO dorm_room_med_inspections(inspection_date, block_name, room_no, created_at) VALUES (?, 'A Blok', '204', ?)", (today_s, today_s))
            c.execute("INSERT INTO dorm_food_samples(sample_date, meal_slot, sampled_dishes, retention_until_iso, created_at) VALUES (?, 'Akşam Yemeği', 'Çorba, Pilav', ?, ?)", (today_s, today_s, today_s))
            c.execute("INSERT INTO dorm_staff_health_checks(staff_name, exam_date, valid_until_date, created_at) VALUES ('Aşçı Usta', ?, ?, ?)", (today_s, (date.today() + timedelta(days=180)).isoformat(), today_s))
            c.execute("INSERT INTO dorm_facility_sanitation_logs(check_date, check_type, measured_value, is_compliant, created_at) VALUES (?, 'Su Deposu Klor', '0.35 ppm', 1, ?)", (today_s, today_s))
            c.execute("INSERT INTO dorm_tournament_kits(event_name, destination, coach_or_teacher, given_date, created_at) VALUES ('Voleybol İl Maçı', 'Spor Salonu', 'Antrenör Can', ?, ?)", (today_s, today_s))
            c.commit()

        insp = sdorm.evaluate_meb_dorm_inspection(self.connect)
        self.assertGreaterEqual(insp["total_score"], 75)
        insp_docx = _docx_text(sdorm.build_meb_dorm_inspection_docx(self.connect, school_name="Test Pansiyonu"))
        self.assertIn("MEB PANSİYON", insp_docx)

        xl_bytes = sdorm.build_dormitory_full_excel(self.connect)
        wb = openpyxl.load_workbook(io.BytesIO(xl_bytes), read_only=True)
        try:
            self.assertIn("Yurt_Ogrencileri", wb.sheetnames)
            self.assertIn("MEB_Denetim_Puani", wb.sheetnames)
        finally:
            wb.close()

        # 7. Navigation, Command Palette & Headless UI Render across all 6 Dormitory tabs + KVKK mode
        self.assertIn("🏨 Okul Yurdu", persantil.PAGES)
        cmd_hits = persantil._handle_cmd_search("204")
        self.assertTrue(any(h.get("value") == "🏨 Okul Yurdu" for h in cmd_hits))

        out_pg = revir_ui.run_action_headless(persantil.render_school_dormitory_page, {"action": "render"})
        main_h = str(out_pg.get("main_html") or "")
        self.assertNotIn("Çalışma Zamanı Hatası", main_h)
        self.assertIn("Okul Yurdu / Pansiyon", main_h)
        self.assertIn("TOPLU GIDA ZEHİRLENMESİ ALARMI", main_h)

        ctx_api = {
            "connect_db": persantil.connect_db,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": lambda: (False, "", ""),
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "open_person_dossier_dialog": persantil.open_person_dossier_dialog,
            "is_section_visible": persantil.is_section_visible,
        }
        all_res = sdorm.get_dorm_residents(self.connect)
        cfg_now = persantil.get_app_settings()
        for tab_fn in (
            lambda: sdorm._render_tab_rooms_and_contacts(ctx_api, all_res, cfg_now),
            lambda: sdorm._render_tab_observation_and_night_shift(ctx_api, all_res, cfg_now),
            lambda: sdorm._render_tab_meds_cabinet_and_night_cards(ctx_api, all_res, cfg_now),
            lambda: sdorm._render_tab_isolation_leave_rest_sports(ctx_api, all_res, cfg_now),
            lambda: sdorm._render_tab_food_safety_hygiene_dysmenorrhea_devices(ctx_api, all_res, cfg_now),
            lambda: sdorm._render_tab_facility_psych_parent_inspection(ctx_api, all_res, cfg_now),
        ):
            res_t = revir_ui.run_action_headless(tab_fn, {"action": "render"})
            self.assertNotIn("Çalışma Zamanı Hatası", str(res_t.get("main_html") or ""))
            self.assertNotIn("Traceback", json.dumps(res_t, ensure_ascii=False))
        _ = ext_rid


if __name__ == "__main__":
    unittest.main()
