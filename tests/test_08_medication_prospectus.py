# -*- coding: utf-8 -*-
"""İlaç Prospektüsü, Türkiye Piyasa İlaç Veri Bankası ve AI Sorgu Entegrasyon Testleri."""
from __future__ import annotations

import io
import json
import sys
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest import mock
import zipfile

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import gemini_assistant as gai
import medication_prospectus as mprosp
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestMedicationProspectus(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        revir_ui.reset_session()

    def tearDown(self):
        revir_ui.reset_session()
        self.ctx.__exit__(None, None, None)

    def _seed_student_and_med(
        self,
        *,
        med_name: str = "Augmentin-BID 400 mg Süspansiyon",
        allergies: str = "Penisilin alerjisi",
    ) -> tuple[int, int]:
        with self.connect() as conn:
            c_row = conn.execute("SELECT id FROM classes WHERE name='3-A'").fetchone()
            cid = int(c_row["id"]) if c_row else int(conn.execute("INSERT INTO classes(name) VALUES ('3-A')").lastrowid)
            sid = int(
                conn.execute(
                    "INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date) VALUES (?, 'Öğrenci', '301', 'Zeynep Kaya', 'K', '2018-04-10')",
                    (cid,),
                ).lastrowid
            )
            conn.execute(
                """INSERT INTO health_profiles(student_id, allergies, chronic_conditions, updated_at)
                   VALUES (?, ?, 'Astım', ?)""",
                (sid, allergies, datetime.now().isoformat(timespec="seconds")),
            )
        med_id = persantil.add_student_medication(
            student_id=sid,
            medication_name=med_name,
            route="Ağızdan (Oral)",
            dosage_pattern="2x1",
            dose_amount=1.0,
            dose_unit="ölçek",
            total_volume=100.0,
            scheduled_times="09:30, 14:30",
            active_days="Her Gün (Hafta Sonu Dahil)",
            start_date=date.today().isoformat(),
        )
        return sid, med_id

    def test_builtin_drug_bank_seeding_matching_crud_and_docx(self):
        persantil.initialize_db()
        stats = mprosp.get_prospectus_stats(persantil.connect_db)
        self.assertGreaterEqual(stats["builtin"], 55)
        self.assertEqual(stats["builtin"], len(mprosp.BUILTIN_TR_MEDICATION_BANK))

        # Re-seeding stays strictly idempotent
        with self.connect() as conn:
            added_again = mprosp.seed_default_prospectus_if_needed(conn)
        self.assertEqual(added_again, 0)

        # Smart Turkish matching for common clinical brand strings
        calpol = mprosp.match_medication_prospectus(persantil.connect_db, "Calpol 120 mg Şurup")
        self.assertIsNotNone(calpol)
        self.assertIn("Parasetamol", calpol["generic_name"])
        for req_field in (
            "general_purpose",
            "administration_and_meal",
            "max_daily_dose",
            "side_effects",
            "contraindicated_drugs",
        ):
            self.assertTrue(str(calpol.get(req_field) or "").strip(), f"Missing field {req_field}")

        aug = mprosp.match_medication_prospectus(persantil.connect_db, "Augmentin-BID 400 mg Süspansiyon")
        self.assertIsNotNone(aug)
        self.assertIn("Amoksisilin", aug["generic_name"])

        vent = mprosp.match_medication_prospectus(persantil.connect_db, "Ventolin Evohaler 100 mcg")
        self.assertIsNotNone(vent)
        self.assertEqual(vent["route_default"], "Solunum (İnhalasyon / Hava)")

        epi = mprosp.match_medication_prospectus(persantil.connect_db, "EpiPen Jr 0.15 mg")
        self.assertIsNotNone(epi)
        self.assertIn("Anafilaksi", epi["therapeutic_class"] + " " + epi["general_purpose"])

        self.assertIsNone(mprosp.match_medication_prospectus(persantil.connect_db, "BilinmeyenOzelMolekul 999 mg"))

        # Custom CRUD + Turkish-folded search + Word (.docx) export
        new_id = mprosp.upsert_prospectus_record(
            persantil.connect_db,
            {
                "medication_name": "Özel Piyasa Şurubu 100 mg",
                "generic_name": "ÖzelEtken",
                "brand_examples": "OzelSurup",
                "therapeutic_class": "Özel Solunum",
                "dosage_form": "Şurup / Süspansiyon",
                "route_default": "Ağızdan (Oral)",
                "meal_relation_default": "Tok karnına",
                "storage_condition_default": "Oda Sıcaklığı (15-25 °C)",
                "general_purpose": "Kuru öksürüğü yumuşatır.",
                "administration_and_meal": "Tok karnına ölçek ile içilir.",
                "max_daily_dose": "Çocuk max 15 ml/gün.",
                "side_effects": "Hafif mide hassasiyeti.",
                "contraindicated_drugs": "Sakinleştirici ilaçlarla alınmaz.",
            },
        )
        found = mprosp.list_prospectus_records(persantil.connect_db, search_text="ozel piyasa")
        self.assertTrue(any(int(x["id"]) == new_id for x in found))

        docx_bytes = mprosp.build_prospectus_docx(
            found[0],
            ensure_docx_fn=persantil._ensure_docx_loaded,
            app_cfg={"school_name": "Test Okulu", "nurse_name": "Hemşire A"},
        )
        self.assertTrue(docx_bytes.startswith(b"PK\x03\x04"))
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            doc_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        self.assertIn("Özel Piyasa Şurubu", doc_xml)
        self.assertIn("GÜNLÜK ALINABİLECEK MAKSİMUM DOZ", doc_xml)

        self.assertTrue(mprosp.delete_prospectus_record(persantil.connect_db, new_id))

        # Re-insert custom item and verify perform_factory_reset restores pristine built-in bank
        mprosp.upsert_prospectus_record(
            persantil.connect_db,
            {"medication_name": "Sıfırlanacak Özel İlaç", "general_purpose": "Test"},
        )
        eg_dir = self.tmp_p / "egitim_yillari"
        ev_dir = self.tmp_p / "revir_evraklari"
        eg_dir.mkdir(exist_ok=True)
        ev_dir.mkdir(exist_ok=True)
        with (
            mock.patch.object(persantil, "EGITIM_YILLARI_DIR", eg_dir),
            mock.patch.object(persantil, "EVRAKLAR_DIR", ev_dir),
        ):
            ok_rst, _ = persantil.perform_factory_reset()
        self.assertTrue(ok_rst)
        stats_after = mprosp.get_prospectus_stats(persantil.connect_db)
        self.assertEqual(stats_after["builtin"], len(mprosp.BUILTIN_TR_MEDICATION_BANK))
        self.assertEqual(stats_after["custom"], 0)

    def test_gemini_ai_prospectus_lookup_full_and_missing_only(self):
        fake_ai_json = {
            "medication_name": "Prospan Öksürük Şurubu",
            "generic_name": "Hedera helix (Duvar Sarmaşığı Ekstresi)",
            "brand_examples": "Prospan, Hedera",
            "therapeutic_class": "Bitkisel Ekspektoran",
            "dosage_form": "Şurup / Süspansiyon",
            "route_default": "Ağızdan",
            "meal_relation_default": "Yemeklerden sonra tok",
            "storage_condition_default": "Oda Sıcaklığı",
            "general_purpose": "Akut bronşit ve balgamlı öksürükte sekresyonu yumuşatır.",
            "administration_and_meal": "Aç veya tok karnına ölçek ile verilir.",
            "max_daily_dose": "6-11 yaş max 10 ml/gün; yetişkin max 22.5 ml/gün.",
            "side_effects": "Nadiren hafif ishal veya alerjik döküntü.",
            "contraindicated_drugs": "Bilinen kritik etkileşim azdır; fruktoz intoleransında kullanılmaz.",
            "pediatric_notes": "Okul çağı çocuklarda su alımı artırılmalıdır.",
            "full_prospectus_notes": "Şişe her doz öncesi çalkalanmalıdır.",
        }
        with mock.patch.object(gai, "call_gemini", return_value=(True, json.dumps(fake_ai_json, ensure_ascii=False), "gemini-3-flash-preview")):
            ok, res, used_m = gai.lookup_medication_prospectus_with_ai(
                api_key="test-key",
                model="gemini-3-flash-preview",
                medication_name="Prospan",
                existing_fields={"general_purpose": "Kendi özel kullanım amacım."},
                fill_mode="missing_only",
            )
        self.assertTrue(ok)
        self.assertEqual(used_m, "gemini-3-flash-preview")
        self.assertEqual(res["general_purpose"], "Kendi özel kullanım amacım.")
        self.assertEqual(res["route_default"], "Ağızdan (Oral)")
        self.assertEqual(res["meal_relation_default"], "Tok karnına")
        self.assertIn("22.5 ml", res["max_daily_dose"])
        self.assertIn("fruktoz", res["contraindicated_drugs"])

    def test_clickable_medication_in_schedule_cabinet_and_history_opens_prospectus(self):
        persantil.initialize_db()
        sid, med_id = self._seed_student_and_med(med_name="Augmentin-BID 400 mg Süspansiyon")
        today_iso = date.today().isoformat()
        persantil.log_medication_dose(
            medication_id=med_id,
            student_id=sid,
            log_date=today_iso,
            scheduled_time="09:30",
            given_time="09:35",
            dose_given=1.0,
            dose_unit="ölçek",
            status="Verildi",
            administered_by="Hemşire Test",
        )

        # 1. Tab 1 Timeline button click opens prospectus modal
        t1_btn_key = f"btn_dose_prosp_dose_{med_id}_09:30_{today_iso}"
        out_t1 = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "click", "key": t1_btn_key},
        )
        dlg1 = out_t1.get("dialog_html") or ""
        self.assertIn("Genel Kullanım Amacı", dlg1)
        self.assertIn("Günlük Alınabilecek Maksimum Doz", dlg1)
        self.assertIn("Olası Yan Etkiler", dlg1)
        self.assertIn("Beraber Kullanılmaması Gereken İlaçlar", dlg1)
        self.assertIn("Penisilin alerjisi", dlg1)

        # 2. Tab 1 Table View click opens prospectus modal
        revir_ui.reset_session()
        revir_ui.session_state["med_tb_view_mode"] = "Tablo Görünümü"
        out_tb = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "df_select", "key": f"med_tb_df_{today_iso}", "value": {"rows": [0], "cells": [[0, "İlaç Adı"]]}},
        )
        self.assertIn("Amoksisilin + Klavulanik Asit", out_tb.get("dialog_html") or "")

        # 3. Tab 2 Cabinet button click opens prospectus modal
        revir_ui.reset_session()
        revir_ui.session_state["_force_tab_label"] = "📦 İlaç Dolabı & Aktif Tedavi Planları (1)"
        out_t2 = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "click", "key": f"btn_cab_prosp_{med_id}"},
        )
        self.assertIn("YEMEĞİN BAŞINDA", out_t2.get("dialog_html") or "")

        # 4. Tab 3 History table row click opens prospectus modal
        revir_ui.reset_session()
        revir_ui.session_state["_force_tab_label"] = "📜 Revir İlaç Defteri & Geçmiş Kayıtlar"
        out_t3 = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "df_select", "key": "med_history_df", "value": {"rows": [0], "cells": [[0, "İlaç Adı"]]}},
        )
        self.assertIn("Augmentin-BID", out_t3.get("dialog_html") or "")

    def test_prospectus_tab_ai_search_custom_add_and_missing_modal_flow(self):
        persantil.initialize_db()
        fake_custom_ai = {
            "medication_name": "NadirTest Şurup 25 mg",
            "generic_name": "NadirTestium",
            "brand_examples": "NadirTest",
            "therapeutic_class": "Özel Tedavi",
            "dosage_form": "Şurup / Süspansiyon",
            "route_default": "Ağızdan (Oral)",
            "meal_relation_default": "Aç karnına",
            "storage_condition_default": "Buzdolabı (+2 / +8 °C)",
            "general_purpose": "Nadir Test sendromu semptomlarını hafifletir.",
            "administration_and_meal": "Sabah aç karnına uygulanır ve buzdolabında saklanır.",
            "max_daily_dose": "Günde en fazla 2x5 ml (250 mg).",
            "side_effects": "Hafif baş ağrısı veya uyuklama.",
            "contraindicated_drugs": "Makrolid grubu antibiyotiklerle birlikte verilmez.",
            "pediatric_notes": "Uygulama sonrası 15 dk dinlenme önerilir.",
            "full_prospectus_notes": "10 gün içinde tüketilmelidir.",
        }

        # Jump via header button to Prospectus Tab
        out_jump = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "click", "key": "btn_open_prosp_tab_top"},
        )
        self.assertIn("İlaç Prospektüsü Sorgulama &amp; Piyasa Veri Bankası", out_jump["main_html"])
        self.assertIn("Genel Kullanım Amacı", out_jump["main_html"])

        # Add a medication NOT in local DB to student and click it -> AI fetch & save from modal
        sid, unknown_med_id = self._seed_student_and_med(med_name="NadirTest Şurup 25 mg")
        with (
            mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "fake-key", "gemini-3-flash-preview")),
            mock.patch.object(gai, "call_gemini", return_value=(True, json.dumps(fake_custom_ai, ensure_ascii=False), "gemini-3-flash-preview")),
        ):
            revir_ui.reset_session()
            revir_ui.session_state["_force_tab_label"] = "📦 İlaç Dolabı & Aktif Tedavi Planları (1)"
            out_open_missing = revir_ui.run_action_headless(
                persantil.render_medication_tracker,
                {"action": "click", "key": f"btn_cab_prosp_{unknown_med_id}"},
            )
            self.assertIn("Gemini AI ile Prospektüsünü Sorgula &amp; Kaydet", out_open_missing.get("dialog_html") or "")

            out_fetched = revir_ui.run_action_headless(
                persantil.render_medication_tracker,
                {"action": "click", "key": "btn_modal_ai_fetch_new"},
            )
            dlg_fetched = out_fetched.get("dialog_html") or ""
            self.assertIn("Nadir Test sendromu semptomlarını", dlg_fetched)
            self.assertIn("Makrolid grubu antibiyotiklerle birlikte verilmez.", dlg_fetched)

        # Confirm saved permanently in SQLite
        matched_now = mprosp.match_medication_prospectus(persantil.connect_db, "NadirTest Şurup")
        self.assertIsNotNone(matched_now)
        self.assertEqual(matched_now["source_type"], "ai")
        self.assertEqual(int(matched_now["ai_enriched"]), 1)

        # Test custom editor dialog + AI missing-fields fill + save
        revir_ui.reset_session()
        revir_ui.session_state["_force_tab_label"] = "📖 İlaç Prospektüsü & Veri Bankası"
        with (
            mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "fake-key", "gemini-3-flash-preview")),
            mock.patch.object(gai, "call_gemini", return_value=(True, json.dumps({**fake_custom_ai, "medication_name": "Manuel Özel Damla", "route_default": "Göz Damlası"}, ensure_ascii=False), "gemini-3-flash-preview")),
        ):
            revir_ui.run_action_headless(persantil.render_medication_tracker, {"action": "render"})
            out_ed = revir_ui.run_action_headless(
                persantil.render_medication_tracker,
                {"action": "click", "key": "btn_prosp_add_new"},
            )
            self.assertIn("Yeni İlacı &amp; Prospektüsünü Veri Bankasına Kaydet", out_ed.get("dialog_html") or "")

            revir_ui.session_state["prosp_ed_new_medication_name"] = "Manuel Özel Damla"
            revir_ui.session_state["prosp_ed_new_general_purpose"] = "Manuel girilen göz nemlendirici amacı."
            out_ai_fill = revir_ui.run_action_headless(
                persantil.render_medication_tracker,
                {"action": "click", "key": "prosp_ed_new_btn_ai_missing"},
            )
            self.assertIn("Manuel girilen göz nemlendirici amacı.", out_ai_fill.get("dialog_html") or "")
            self.assertEqual(revir_ui.session_state.get("prosp_ed_new_route_default"), "Göz Damlası")

            revir_ui.run_action_headless(
                persantil.render_medication_tracker,
                {"action": "click", "key": "prosp_ed_new_btn_save"},
            )

        saved_custom = mprosp.match_medication_prospectus(persantil.connect_db, "Manuel Özel Damla")
        self.assertIsNotNone(saved_custom)
        self.assertEqual(saved_custom["general_purpose"], "Manuel girilen göz nemlendirici amacı.")
        self.assertEqual(saved_custom["route_default"], "Göz Damlası")

        # Verify Add Medication Dialog shows live match helper when adding new med
        revir_ui.reset_session()
        revir_ui.session_state["med_dlg_mode_radio"] = "➕ Yeni İlaç Ekle"
        revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "click", "key": "btn_open_add_med_top"},
        )
        out_dlg = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "change", "key": "dlg_new_name", "value": "Calpol 120 mg"},
        )
        self.assertIn("Prospektüs Eşleşti", out_dlg.get("dialog_html") or "")
        out_apply = revir_ui.run_action_headless(
            persantil.render_medication_tracker,
            {"action": "click", "key": "dlg_new_btn_fill_from_prosp"},
        )
        self.assertIn("Oda Sıcaklığı", out_apply.get("dialog_html") or "")
        self.assertEqual(revir_ui.session_state.get("dlg_new_route"), "Ağızdan (Oral)")


if __name__ == "__main__":
    unittest.main()
