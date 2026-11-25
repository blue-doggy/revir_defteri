# -*- coding: utf-8 -*-
"""Revir Sarf Malzemesi, İlk Yardım Çantaları, Cihaz Kalibrasyonu ve Stok Hareketleri Entegrasyon Testleri."""
from __future__ import annotations

from datetime import date, timedelta
import io
import sys
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import gemini_assistant
import infirmary_supply as invsup
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestInfirmarySupplyV2(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        revir_ui.reset_session()

    def tearDown(self):
        revir_ui.reset_session()
        self.ctx.__exit__(None, None, None)

    def _seed_student(self, full_name: str = "Emre Yıldız") -> int:
        with self.connect() as conn:
            c_row = conn.execute("SELECT id FROM classes WHERE name='4-B'").fetchone()
            cid = int(c_row["id"]) if c_row else int(conn.execute("INSERT INTO classes(name) VALUES ('4-B')").lastrowid)
            sid = int(
                conn.execute(
                    "INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date) VALUES (?, 'Öğrenci', '402', ?, 'E', '2017-03-14')",
                    (cid, full_name),
                ).lastrowid
            )
            return sid

    def test_01_schema_extended_catalog_and_movements_and_physical_count(self):
        added = persantil.load_default_infirmary_inventory()
        self.assertGreaterEqual(added, 20)
        items = persantil.get_infirmary_inventory()
        names = {it["item_name"] for it in items}
        self.assertIn("Steril Yanık Örtüsü (Jelli / Alüminyum)", names)
        self.assertIn("Suni Solunum Maskesi (CPR Pocket Mask)", names)

        # Açılış hareketleri kaydedilmiş olmalı
        moves = invsup.get_supply_movements(self.connect, limit=50)
        self.assertGreaterEqual(len(moves), 20)

        target_item = next(it for it in items if "Serum Fizyolojik" in it["item_name"])
        iid = int(target_item["id"])
        start_qty = float(target_item["quantity"])

        # Tüketim (-3) ve Tedarik (+5)
        res_use = invsup.record_supply_movement(
            self.connect,
            inventory_id=iid,
            movement_type="Sarf / Kullanım",
            delta_qty=-3.0,
            performed_by="Hemşire Ayşe",
            reason="Diz sıyrığı pansumanı",
        )
        self.assertAlmostEqual(res_use["new_qty"], start_qty - 3.0)

        res_add = invsup.record_supply_movement(
            self.connect,
            inventory_id=iid,
            movement_type="Giriş / Tedarik",
            delta_qty=5.0,
            performed_by="Hemşire Ayşe",
            reason="Yeni fatura girişi",
            lot_number="LOT-2026-88",
        )
        self.assertAlmostEqual(res_add["new_qty"], start_qty + 2.0)

        # Fiziki sayım farkı düzeltme testi
        df_cnt = pd.DataFrame([
            {
                "id": iid,
                "Malzeme Adı": target_item["item_name"],
                "Sistem Miktarı": start_qty + 2.0,
                "Sayılan Miktar": start_qty - 1.0,
            }
        ])
        cnt_res = invsup.apply_physical_inventory_count(self.connect, df_cnt, performed_by="Sayım Ekibi")
        self.assertEqual(cnt_res["adjusted"], 1)
        latest_item = next(it for it in persantil.get_infirmary_inventory() if int(it["id"]) == iid)
        self.assertAlmostEqual(float(latest_item["quantity"]), start_qty - 1.0)
        self.assertIsNotNone(latest_item["last_counted_at"])

    def test_02_visit_form_auto_deducts_infirmary_supplies(self):
        revir_ui.reset_session()
        sid = self._seed_student("Elif Naz Kaya")
        persantil.set_active_person(sid)
        iid = persantil.add_infirmary_inventory_item(
            "Steril Gazlı Bez Test",
            "Pansuman & Sarf",
            quantity=12.0,
            unit="Paket",
            min_threshold=3.0,
        )
        sup_lbl = f"Steril Gazlı Bez Test — Kalan: 12 Paket [#{iid}]"

        revir_ui.session_state[f"uv_stu_sym_{sid}"] = "Dizde sıyrık"
        revir_ui.session_state[f"uv_stu_act_{sid}"] = "Pansuman yapıldı"
        revir_ui.session_state[f"uv_stu_sup_sel_{sid}"] = [sup_lbl]
        revir_ui.session_state[f"uv_stu_sup_amt_{sid}"] = 2.0

        out_v = revir_ui.run_action_headless(
            persantil.render_visit,
            {"action": "click", "key": f"btn_save_uv_stu_{sid}"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", out_v["main_html"])

        # Stok 12 -> 10 düşmeli, visits observation içinde [Sarf: ...] ve supply_movements'ta visit_id olmalı
        updated_it = next(it for it in persantil.get_infirmary_inventory() if int(it["id"]) == iid)
        self.assertAlmostEqual(float(updated_it["quantity"]), 10.0)

        with self.connect() as conn:
            v_row = conn.execute("SELECT id, observation FROM visits WHERE student_id=? ORDER BY id DESC LIMIT 1", (sid,)).fetchone()
            self.assertIsNotNone(v_row)
            self.assertIn("[Sarf: Steril Gazlı Bez Test (2 Paket)]", v_row["observation"])
            mv_row = conn.execute("SELECT * FROM supply_movements WHERE inventory_id=? AND visit_id=?", (iid, v_row["id"])).fetchone()
            self.assertIsNotNone(mv_row)
            self.assertEqual(int(mv_row["student_id"]), sid)
            self.assertAlmostEqual(float(mv_row["delta_qty"]), -2.0)

    def test_03_first_aid_kits_replenishment_devices_and_factory_reset(self):
        persantil.load_default_infirmary_inventory()
        kn = invsup.seed_default_first_aid_kits(self.connect, custodian_default="Hemşire Zeynep")
        dn = invsup.seed_default_medical_devices(self.connect)
        self.assertEqual(kn, 6)
        self.assertEqual(dn, 6)

        kits = invsup.list_first_aid_kits(self.connect)
        sport_kit = next(k for k in kits if "Spor Salonu" in k["kit_name"])
        sport_items = sport_kit["items"]

        # İlk kalemde (Elastik Bandaj) 4 yerine 1 mevcut kalsın -> depodan +3 ikmal edilmeli
        before_bandaj = next(it for it in persantil.get_infirmary_inventory() if "Elastik Bandaj" in it["item_name"])
        before_qty = float(before_bandaj["quantity"])

        df_kit_edit = pd.DataFrame([
            {
                "id": int(ki["id"]),
                "Malzeme": ki["item_name"],
                "Gerekli": float(ki["required_qty"]),
                "Mevcut": 1.0 if "Elastik Bandaj" in ki["item_name"] else float(ki["current_qty"]),
                "Birim": ki["unit"],
                "SKT (YYYY-AA-GG)": "",
                "Durum": "Eksik" if "Elastik Bandaj" in ki["item_name"] else "Tamam",
                "Not": "",
            }
            for ki in sport_items
        ])
        res_rep = invsup.save_kit_items_and_inspect(
            self.connect,
            kit_id=int(sport_kit["id"]),
            edited_items_df=df_kit_edit,
            checked_by="Hemşire Zeynep",
            replenish_from_depot=True,
        )
        self.assertGreaterEqual(res_rep["replenished_count"], 1)
        self.assertEqual(res_rep["missing_count"], 0)
        self.assertEqual(res_rep["overall_status"], "Tamam (Hazır)")

        after_bandaj = next(it for it in persantil.get_infirmary_inventory() if int(it["id"]) == int(before_bandaj["id"]))
        self.assertAlmostEqual(float(after_bandaj["quantity"]), before_qty - 3.0)

        # Cihaz kalibrasyon / AED ped SKT uyarısı + Hazırlık Skoru
        past_iso = (date.today() - timedelta(days=10)).isoformat()
        invsup.upsert_medical_device(
            self.connect,
            device_name="Acil Test AED",
            device_type="OED / AED Defibrilatör",
            battery_status="Zayıf / Değişmeli",
            next_calibration_date=past_iso,
            pad_or_sensor_expiry=past_iso,
            status="Çalışıyor",
        )
        devs = invsup.list_medical_devices(self.connect)
        test_aed = next(d for d in devs if d["device_name"] == "Acil Test AED")
        self.assertTrue(test_aed["needs_attention"])
        self.assertIn("Kalibrasyon Gecikmiş", test_aed["alert_flags"])
        self.assertIn("Ped/Sensör SKT Geçmiş", test_aed["alert_flags"])

        summary = invsup.compute_infirmary_readiness_summary(self.connect)
        self.assertLess(summary["readiness_score"], 100)
        self.assertGreaterEqual(len(summary["devices_attention"]), 1)

        # Fabrika ayarlarına sıfırlama tüm yeni tabloları temizlemeli
        ok_rst, _ = persantil.perform_factory_reset()
        self.assertTrue(ok_rst)
        with self.connect() as conn:
            for tbl in invsup.SUPPLY_TABLES:
                cnt = conn.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
                self.assertEqual(cnt, 0, f"Temizlenmeyen tablo: {tbl}")

    def test_04_expiry_quarantine_procurement_exports_and_headless_ui(self):
        import docx
        import openpyxl
        import pymupdf

        past_d = (date.today() - timedelta(days=5)).isoformat()
        soon_d = (date.today() + timedelta(days=14)).isoformat()
        mid_d = (date.today() + timedelta(days=60)).isoformat()

        id_exp = persantil.add_infirmary_inventory_item("Miadı Geçmiş Turnike", "Acil Müdahale", 2.0, "Adet", 3.0, past_d, "Acil Çantası")
        persantil.add_infirmary_inventory_item("Yakın SKT Flaster", "Pansuman & Sarf", 3.0, "Rulo", 5.0, soon_d, "Dolap")
        persantil.add_infirmary_inventory_item("Orta Vade Bandaj", "İlk Yardım & Ortopedi", 12.0, "Rulo", 4.0, mid_d, "Dolap")
        with self.connect() as conn:
            conn.execute("UPDATE infirmary_inventory SET unit_cost=50.0, target_stock=10.0 WHERE id=?", (id_exp,))

        invsup.seed_default_first_aid_kits(self.connect)
        invsup.seed_default_medical_devices(self.connect)

        buckets = invsup.get_expiry_buckets(invsup.get_enriched_inventory(self.connect))
        self.assertEqual(len(buckets["expired"]), 1)
        self.assertEqual(len(buckets["expiring_30"]), 1)
        self.assertEqual(len(buckets["expiring_90"]), 1)

        # SKT geçmiş ürün imha edildiğinde karantina/imha listesine geçmeli ve aktif kritik alarmdan düşmeli
        invsup.record_supply_movement(
            self.connect,
            inventory_id=id_exp,
            movement_type="SKT İmha / Zayi",
            delta_qty=-2.0,
            performed_by="Hemşire",
            new_quarantine_status="İmha Edildi",
        )
        alerts_after = persantil.get_infirmary_inventory_alerts()
        self.assertEqual(len(alerts_after["expired"]), 0)
        buckets_after = invsup.get_expiry_buckets(invsup.get_enriched_inventory(self.connect))
        self.assertEqual(len(buckets_after["quarantined"]), 1)

        # Satın alma planı & bütçe
        proc = invsup.build_procurement_plan(self.connect)
        self.assertGreater(proc["item_count"], 0)

        # Word (.docx), Excel (.xlsx), A4 PDF (.pdf) çıktıları
        docx_b = persantil.build_infirmary_inventory_docx()
        dx = docx.Document(io.BytesIO(docx_b))
        full_dx = "\n".join(t.text for tbl in dx.tables for r in tbl.rows for t in r.cells)
        self.assertIn("İlk Yardım Hazırlık Skoru", full_dx)

        xlsx_b = invsup.build_supply_excel_bytes(self.connect)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_b))
        self.assertIn("Ana Envanter", wb.sheetnames)
        self.assertIn("Ilk Yardim Cantalari", wb.sheetnames)
        self.assertIn("Tibbi Cihazlar", wb.sheetnames)
        wb.close()

        pdf_b = invsup.build_supply_a4_pdf_bytes(self.connect, persantil.get_app_settings())
        pdf_doc = pymupdf.open(stream=pdf_b, filetype="pdf")
        self.assertGreaterEqual(len(pdf_doc), 1)
        pdf_norm = " ".join(pdf_doc[0].get_text("text").replace("\xa0", " ").split())
        self.assertIn("REVİR İLK YARDIM & SARF STOK DENETİM RAPORU", pdf_norm)
        pdf_doc.close()

        # Headless UI render (6 sekme + Gemini AI paneli)
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "m")):
            revir_ui.reset_session()
            out_ui = revir_ui.run_action_headless(persantil._render_med_tab_infirmary_inventory)
        html_ui = out_ui["main_html"]
        self.assertNotIn("Çalışma Zamanı Hatası", html_ui)
        self.assertIn("İlk Yardım Hazırlık", html_ui)
        self.assertIn("Ana Stok &amp; Hızlı İşlem", html_ui)
        self.assertIn("Hareketler &amp; Fiziki Sayım", html_ui)
        self.assertIn("Çantalar &amp; Cihazlar", html_ui)
        self.assertIn("Stok Önceliklendirme", html_ui)


if __name__ == "__main__":
    unittest.main(verbosity=2)
