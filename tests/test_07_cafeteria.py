"""Comprehensive tests for Yemekhane Yönetimi (cafeteria.py + AI + UI + exports)."""
from __future__ import annotations

import io
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

import docx
import openpyxl
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import cafeteria as caf
import gemini_assistant
import insulin_diabetes as idm
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


def _extract_docx_all_text(raw_bytes: bytes) -> str:
    doc = docx.Document(io.BytesIO(raw_bytes))
    chunks: list[str] = [p.text for p in doc.paragraphs if p.text]
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if cell.text:
                    chunks.append(cell.text)
    return "\n".join(chunks)


class TestCafeteriaModule(unittest.TestCase):
    def setUp(self):
        revir_ui.reset_session()
        self._ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self._ctx.__enter__()
        with self.connect() as c:
            c.execute("INSERT INTO classes(name) VALUES ('4-A')")
            cls_id = int(c.execute("SELECT id FROM classes WHERE name='4-A'").fetchone()["id"])
            c.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Personel')")
            pers_cls_id = int(c.execute("SELECT id FROM classes WHERE name='Personel'").fetchone()["id"])
            c.execute(
                """INSERT INTO students(full_name, sex, class_id, person_type)
                   VALUES ('Ali Kaya', 'E', ?, 'Öğrenci')""",
                (cls_id,),
            )
            self.stu_id = int(c.execute("SELECT id FROM students WHERE full_name='Ali Kaya'").fetchone()["id"])
            c.execute(
                """INSERT INTO health_profiles(student_id, allergies, updated_at)
                   VALUES (?, 'İnek sütü / laktoz ve glüten alerjisi', '2026-10-01T10:00')""",
                (self.stu_id,),
            )
            c.execute(
                """INSERT INTO students(full_name, sex, class_id, person_type)
                   VALUES ('Zeynep Öğretmen', 'K', ?, 'Öğretmen')""",
                (pers_cls_id,),
            )
            self.teacher_id = int(
                c.execute("SELECT id FROM students WHERE full_name='Zeynep Öğretmen'").fetchone()["id"]
            )
            c.commit()

    def tearDown(self):
        self._ctx.__exit__(None, None, None)

    def test_schema_dishes_scaling_and_allergen_crossmatch(self):
        expected_tables = {
            "cafeteria_dishes",
            "cafeteria_dish_ingredients",
            "cafeteria_menu_plans",
            "cafeteria_purchases",
            "cafeteria_purchase_items",
            "cafeteria_diners",
            "cafeteria_monthly_fees",
            "cafeteria_cash_ledger",
        }
        with self.connect() as c:
            existing = {
                r["name"]
                for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }
        self.assertTrue(expected_tables.issubset(existing))

        baseline_ings = [
            {
                "ingredient_name": "Kırmızı Mercimek",
                "amount": 50.0,
                "unit": "g",
                "kcal_portion": 170.0,
                "carbs_portion_g": 30.0,
                "allergen_tag": "",
            },
            {
                "ingredient_name": "Buğday Unu",
                "amount": 10.0,
                "unit": "g",
                "kcal_portion": 36.0,
                "carbs_portion_g": 7.5,
                "allergen_tag": "Gluten",
            },
            {
                "ingredient_name": "Tereyağı",
                "amount": 5.0,
                "unit": "g",
                "kcal_portion": 36.0,
                "carbs_portion_g": 0.0,
                "allergen_tag": "Süt",
            },
        ]

        # Simulate doubling Buğday Unu (10 -> 20 g) without manually changing kcal -> scales 2x;
        # keep Tereyağı kcal manually overridden (45.0) -> manual value preserved.
        edited_df = pd.DataFrame([
            {
                "Malzeme": "Kırmızı Mercimek",
                "Miktar": 50.0,
                "Birim": "g",
                "Kalori (kcal)": 170.0,
                "Karb (g)": 30.0,
                "Alerjen": "",
            },
            {
                "Malzeme": "Buğday Unu",
                "Miktar": 20.0,
                "Birim": "g",
                "Kalori (kcal)": 36.0,
                "Karb (g)": 7.5,
                "Alerjen": "Gluten",
            },
            {
                "Malzeme": "Tereyağı",
                "Miktar": 10.0,
                "Birim": "g",
                "Kalori (kcal)": 45.0,
                "Karb (g)": 0.2,
                "Alerjen": "Süt",
            },
        ])
        scaled = caf.scale_edited_ingredients(baseline_ings, edited_df, auto_scale_nutrition=True)
        self.assertEqual(len(scaled), 3)
        self.assertAlmostEqual(scaled[1]["kcal_portion"], 72.0, places=1)
        self.assertAlmostEqual(scaled[1]["carbs_portion_g"], 15.0, places=1)
        self.assertAlmostEqual(scaled[2]["kcal_portion"], 45.0, places=1)
        self.assertEqual(scaled[1]["allergen_tag"], "Gluten (Buğday/Arpa)")
        self.assertEqual(scaled[2]["allergen_tag"], "Süt ve Süt Ürünleri (Laktoz)")

        dish_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Sütlü Tarhana Çorbası",
            category="Çorba",
            portion_g=260.0,
            ingredients=scaled,
            protein_g=10.5,
            fat_g=6.0,
            recipe_notes="Okul öğle servisi",
            source="ai",
        )
        self.assertGreater(dish_id, 0)

        dishes = caf.get_all_dishes(self.connect)
        self.assertEqual(len(dishes), 1)
        d0 = dishes[0]
        self.assertAlmostEqual(d0["total_kcal"], 170.0 + 72.0 + 45.0, places=1)
        self.assertAlmostEqual(d0["carbs_g"], 45.2, places=1)
        self.assertIn("Gluten (Buğday/Arpa)", d0["allergens"])
        self.assertIn("Süt ve Süt Ürünleri (Laktoz)", d0["allergens"])

        # Add extra active diner with peanut allergy & verify cross-check returns both Ali Kaya and Ayşe Demir
        caf.upsert_diner(
            self.connect,
            full_name="Ayşe Demir",
            role_type="Misafir",
            allergies_override="Yer fıstığı anafilaksi",
        )
        hits = caf.match_allergens_against_school(
            self.connect,
            ["Gluten (Buğday/Arpa)", "Süt ve Süt Ürünleri (Laktoz)", "Yer Fıstığı"],
        )
        hit_names = {h["full_name"] for h in hits}
        self.assertIn("Ali Kaya", hit_names)
        self.assertIn("Ayşe Demir", hit_names)

        # Delete check
        caf.delete_dish_by_id(self.connect, dish_id)
        self.assertEqual(caf.get_all_dishes(self.connect), [])

    def test_workdays_holidays_menu_planning_and_diabetes_sync(self):
        # Test October 2026: Oct 29 (Cumhuriyet Bayramı) is official holiday -> not a workday.
        # Oct 28 is half-day (Cumhuriyet Bayramı Arifesi).
        days_half_work = caf.get_month_workdays_info(self.connect, "2026-10", treat_half_day_as_workday=True)
        by_iso = {d["date_iso"]: d for d in days_half_work}
        self.assertEqual(len(days_half_work), 31)
        self.assertFalse(by_iso["2026-10-03"]["is_workday"])  # Saturday
        self.assertFalse(by_iso["2026-10-04"]["is_workday"])  # Sunday
        self.assertFalse(by_iso["2026-10-29"]["is_workday"])  # Cumhuriyet Bayramı
        self.assertTrue(by_iso["2026-10-28"]["is_workday"])   # Half-day counted as workday

        days_no_half = caf.get_month_workdays_info(self.connect, "2026-10", treat_half_day_as_workday=False)
        by_iso_nh = {d["date_iso"]: d for d in days_no_half}
        self.assertFalse(by_iso_nh["2026-10-28"]["is_workday"])

        # Add 2 dishes and plan a meal + sync to diabetes
        d_soup = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Mercimek Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Mercimek", "amount": 50, "unit": "g", "kcal_portion": 180, "carbs_portion_g": 28, "allergen_tag": ""}],
        )
        d_main = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Peynirli Makarna",
            category="Yardımcı Yemek",
            ingredients=[
                {"ingredient_name": "Makarna", "amount": 80, "unit": "g", "kcal_portion": 290, "carbs_portion_g": 56, "allergen_tag": "Gluten"},
                {"ingredient_name": "Beyaz Peynir", "amount": 30, "unit": "g", "kcal_portion": 80, "carbs_portion_g": 1, "allergen_tag": "Süt"},
            ],
        )

        res_slot = caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2026-10-05",
            meal_slot="Öğle Yemeği",
            dish_ids=[d_soup, d_main],
            custom_items=[{"name": "Mevsim Salatası", "kcal": 45, "carbs_g": 5, "allergen": ""}],
            notes="Hafta başlangıcı",
            sync_to_diabetes=True,
        )
        self.assertAlmostEqual(res_slot["total_kcal"], 595.0, places=1)
        self.assertAlmostEqual(res_slot["total_carbs_g"], 90.0, places=1)
        self.assertIn("Gluten (Buğday/Arpa)", res_slot["allergens"])
        self.assertIn("Süt ve Süt Ürünleri (Laktoz)", res_slot["allergens"])

        plans = caf.get_menu_plans_for_month(self.connect, "2026-10")
        self.assertIn(("2026-10-05", "Öğle Yemeği"), plans)

        diab_menus = idm.get_school_menus_for_month(self.connect, "2026-10")
        self.assertTrue(any(m["menu_date"] == "2026-10-05" and m["meal_type"] == "Öğle Yemeği" for m in diab_menus))

        # Manual full-month sync
        caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2026-10-06",
            meal_slot="Sabah Kahvaltısı",
            dish_ids=[d_main],
            sync_to_diabetes=False,
        )
        synced = caf.sync_cafeteria_month_to_diabetes(self.connect, "2026-10")
        self.assertGreaterEqual(synced, 2)
        diab_menus_after = idm.get_school_menus_for_month(self.connect, "2026-10")
        self.assertTrue(any(m["menu_date"] == "2026-10-06" and m["meal_type"] == "Kahvaltı" for m in diab_menus_after))

        removed = caf.delete_menu_plan_slot(self.connect, "2026-10-06", "Sabah Kahvaltısı")
        self.assertEqual(removed, 1)

    def test_menu_a4_html_and_docx_monthly_and_weekly_with_signatures(self):
        d1 = caf.upsert_dish_with_ingredients(
            self.connect,
            name="İzmir Köfte",
            category="Ana Yemek",
            ingredients=[{"ingredient_name": "Ekmek İçi", "amount": 15, "unit": "g", "kcal_portion": 310, "carbs_portion_g": 18, "allergen_tag": "Gluten"}],
        )
        caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2026-10-05",
            meal_slot="Öğle Yemeği",
            dish_ids=[d1],
        )
        signers = [
            ("Hazırlayan", "Hemşire Elif Yılmaz"),
            ("Yemekhane Sorumlusu", "Ahmet Demir"),
            ("Okul Müdürü (Onay)", "Dr. Kemal Arslan"),
        ]
        meta = {"institution_top_header": "T.C. MEB", "school_name": "Örnek Anadolu Lisesi"}

        for scope in ("monthly", "weekly"):
            html_out = caf.build_menu_a4_html(
                self.connect,
                year_month="2026-10",
                scope=scope,
                week_index=1,
                active_slots=["Öğle Yemeği"],
                show_kcal=True,
                show_allergens=True,
                signers=signers,
                school_meta=meta,
            )
            self.assertIn("İzmir Köfte", html_out)
            self.assertIn("310 kcal", html_out)
            self.assertIn("Gluten", html_out)
            self.assertIn("Hemşire Elif Yılmaz", html_out)
            self.assertIn("Ahmet Demir", html_out)
            self.assertIn("Dr. Kemal Arslan", html_out)
            self.assertIn("İmza", html_out)

            docx_bytes = caf.build_menu_a4_docx(
                self.connect,
                year_month="2026-10",
                scope=scope,
                week_index=1,
                active_slots=["Öğle Yemeği"],
                show_kcal=True,
                show_allergens=True,
                signers=signers,
                school_meta=meta,
            )
            self.assertGreater(len(docx_bytes), 4000)
            txt = _extract_docx_all_text(docx_bytes)
            self.assertIn("İzmir Köfte", txt)
            self.assertIn("310 kcal", txt)
            self.assertIn("Gluten", txt)
            self.assertIn("Hemşire Elif Yılmaz", txt)
            self.assertIn("Ahmet Demir", txt)
            self.assertIn("Dr. Kemal Arslan", txt)
            self.assertIn("İmza", txt)

    def test_purchases_invoices_and_expense_statistics(self):
        pid1 = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-10-05",
            invoice_no="FTR-101",
            supplier_name="Özkan Bakliyat",
            category="Bakliyat & Kuru Gıda",
            items=[
                {"item_name": "Kırmızı Mercimek", "category": "Bakliyat & Kuru Gıda", "quantity": 25.0, "unit_type": "kg", "unit_price": 60.0, "vat_rate": 1.0},
                {"item_name": "Ayçiçek Yağı", "category": "Yağ & Baharat", "quantity": 10.0, "unit_type": "litre", "unit_price": 80.0, "vat_rate": 10.0},
            ],
        )
        # 25*60*1.01 = 1515.0 ; 10*80*1.10 = 880.0 -> total = 2395.0
        pid2 = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-09-20",
            invoice_no="FTR-099",
            supplier_name="Merkez Hal",
            category="Sebze & Meyve",
            items=[{"item_name": "Patates", "category": "Sebze & Meyve", "quantity": 50.0, "unit_type": "kg", "unit_price": 20.0, "vat_rate": 0.0}],
        )
        self.assertGreater(pid1, 0)
        self.assertGreater(pid2, 0)

        p_oct, it_oct = caf.get_purchases_and_items(self.connect, year_month="2026-10")
        self.assertEqual(len(p_oct), 1)
        self.assertEqual(len(it_oct), 2)

        p_all, it_all = caf.get_purchases_and_items(self.connect, year_month="Tümü", query_text="Patates")
        self.assertEqual(len(it_all), 1)
        self.assertEqual(it_all[0]["item_name"], "Patates")

        stats = caf.get_expense_summary_stats(self.connect, "2026-10")
        self.assertAlmostEqual(stats["month_total"], 2395.0, places=2)
        self.assertAlmostEqual(stats["year_total"], 3395.0, places=2)
        self.assertAlmostEqual(stats["all_time_total"], 3395.0, places=2)
        self.assertEqual(stats["month_invoice_count"], 1)

        caf.delete_purchase_by_id(self.connect, pid2)
        stats_after = caf.get_expense_summary_stats(self.connect, "2026-10")
        self.assertAlmostEqual(stats_after["all_time_total"], 2395.0, places=2)

    def test_diners_monthly_fees_cash_ledger_vault_and_documents(self):
        added = caf.import_school_people_as_diners(self.connect, [self.stu_id, self.teacher_id], daily_meals_count=1)
        self.assertEqual(added, 2)
        exempt_id = caf.upsert_diner(
            self.connect,
            full_name="Burak Burslu",
            role_type="Burslu / Ücretsiz",
            class_or_dept="5-B",
            daily_meals_count=1,
            fee_exempt=True,
        )
        self.assertGreater(exempt_id, 0)

        # In 2026-10 with half-day workday=True: workdays = 21 -> expected student fee: 21 * 1 * 120 = 2520; teacher: 21 * 1 * 150 = 3150
        refreshed = caf.generate_or_refresh_monthly_fees(
            self.connect,
            "2026-10",
            student_meal_price=120.0,
            staff_meal_price=150.0,
            workdays_override=20,
            reset_planned_days=True,
        )
        self.assertEqual(refreshed, 3)

        diners = {d["full_name"]: d for d in caf.get_diners(self.connect)}
        stu_did = int(diners["Ali Kaya"]["id"])
        tch_did = int(diners["Zeynep Öğretmen"]["id"])

        # Partial payment from parent Mehmet Kaya + full payment from Zeynep Öğretmen + school handover
        caf.record_cash_transaction(
            self.connect,
            tx_date="2026-10-08",
            tx_type="IN_PAYMENT",
            amount=1800.0,
            diner_id=stu_did,
            for_person_name="Ali Kaya",
            delivered_by_name="Veli Mehmet Kaya",
            received_by_name="Hemşire Elif",
            payment_method="Nakit",
            receipt_no="MKB-01",
        )
        caf.record_cash_transaction(
            self.connect,
            tx_date="2026-10-09",
            tx_type="IN_PAYMENT",
            amount=3000.0,
            diner_id=tch_did,
            for_person_name="Zeynep Öğretmen",
            delivered_by_name="Zeynep Öğretmen",
            payment_method="IBAN / Havale",
        )
        tx_hand = caf.record_cash_transaction(
            self.connect,
            tx_date="2026-10-15",
            tx_type="OUT_SCHOOL_HANDOVER",
            amount=2500.0,
            delivered_by_name="Hemşire Elif",
            received_by_name="Müdür Kemal Arslan",
            receipt_no="TES-01",
            notes="Ekim ilk taksit okul idaresine teslim",
        )
        self.assertGreater(tx_hand, 0)

        cash = caf.get_cash_summary_and_diner_balances(self.connect, "2026-10")
        bal_by_name = {b["full_name"]: b for b in cash["balances"]}
        self.assertAlmostEqual(bal_by_name["Ali Kaya"]["expected_amount"], 2400.0, places=2)
        self.assertAlmostEqual(bal_by_name["Ali Kaya"]["paid_amount"], 1800.0, places=2)
        self.assertAlmostEqual(bal_by_name["Ali Kaya"]["remaining_amount"], 600.0, places=2)
        self.assertEqual(bal_by_name["Ali Kaya"]["status_label"], "🟡 Kısmi Ödeme")
        self.assertEqual(bal_by_name["Zeynep Öğretmen"]["status_label"], "🟢 Ödendi")
        self.assertEqual(bal_by_name["Burak Burslu"]["status_label"], "Muaf / Burslu")

        # Update planned days on Ali Kaya fee row -> 15 days * 1 * 120 - 0 = 1800 -> becomes Ödendi
        caf.update_monthly_fee_row(
            self.connect,
            int(bal_by_name["Ali Kaya"]["id"]),
            planned_days=15,
            meals_per_day=1,
            unit_meal_price=120.0,
            discount_or_adjustment=0.0,
            note="5 gün raporlu düştü",
        )
        cash2 = caf.get_cash_summary_and_diner_balances(self.connect, "2026-10")
        bal2 = {b["full_name"]: b for b in cash2["balances"]}
        self.assertEqual(bal2["Ali Kaya"]["status_label"], "🟢 Ödendi")
        self.assertAlmostEqual(cash2["month_collected"], 4800.0, places=2)
        self.assertAlmostEqual(cash2["month_handover"], 2500.0, places=2)
        self.assertAlmostEqual(cash2["vault_balance"], 2300.0, places=2)

        # Verify delivered_by_name preserved in transactions
        self.assertTrue(any(t.get("delivered_by_name") == "Veli Mehmet Kaya" for t in cash2["month_transactions"]))

        # Handover Word document & Full 5-sheet Excel workbook
        ho_docx = caf.build_cash_handover_docx(
            self.connect,
            year_month="2026-10",
            school_meta={"school_name": "Örnek İlkokulu"},
            deliverer_name="Hemşire Elif",
            receiver_name="Müdür Kemal Arslan",
        )
        ho_txt = _extract_docx_all_text(ho_docx)
        self.assertIn("Veli Mehmet Kaya", ho_txt)
        self.assertIn("Müdür Kemal Arslan", ho_txt)

        wb_bytes = caf.build_cafeteria_full_workbook_excel(self.connect, "2026-10")
        wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), read_only=True)
        self.assertEqual(
            wb.sheetnames,
            ["Aylik_Menu", "Yemekler_Bankasi", "Gelen_Urun_Gider", "Abone_Aidat", "Kasa_Defteri"],
        )
        wb.close()

    def test_gemini_ai_cafeteria_wrappers_and_confirmation_gate(self):
        # 1. Section 14 wrappers with mocked call_gemini
        dish_json = json.dumps({
            "portion_g": 240,
            "protein_g": 9.5,
            "fat_g": 5.0,
            "recipe_notes": "Az tuzlu pişirin",
            "ingredients": [
                {"ingredient_name": "Nohut", "amount": 60, "unit": "g", "kcal_portion": 210, "carbs_portion_g": 36, "allergen_tag": ""},
                {"ingredient_name": "Un", "amount": 5, "unit": "g", "kcal_portion": 18, "carbs_portion_g": 3.8, "allergen_tag": "Gluten (Buğday/Arpa)"},
            ],
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, dish_json, "gemini-3-flash-preview")):
            ok, res, _ = gemini_assistant.generate_dish_ingredients_and_nutrition_with_gemini(
                api_key="k", model="gemini-3-flash-preview", dish_name="Etli Nohut"
            )
        self.assertTrue(ok)
        self.assertEqual(len(res["ingredients"]), 2)

        menu_json = json.dumps({
            "plans": [
                {"plan_date": "2026-10-12", "meal_slot": "Öğle Yemeği", "dishes": [{"name": "Etli Nohut", "kcal": 228, "carbs_g": 39.8, "allergen": "Gluten (Buğday/Arpa)"}]}
            ]
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, menu_json, "gemini-3-flash-preview")):
            ok_m, res_m, _ = gemini_assistant.generate_monthly_cafeteria_menu_with_gemini(
                api_key="k", model="gemini-3-flash-preview", year_month="2026-10",
                workdays_list=[{"date": "2026-10-12"}], meal_slots=["Öğle Yemeği"], available_dishes=[],
            )
        self.assertTrue(ok_m)
        self.assertEqual(res_m["plans"][0]["plan_date"], "2026-10-12")

        inv_json = json.dumps({
            "invoice_date": "2026-10-12",
            "invoice_no": "FTR-AI-1",
            "supplier_name": "Güneş Gıda",
            "category": "Bakliyat & Kuru Gıda",
            "items": [{"item_name": "Pirinç", "category": "Bakliyat & Kuru Gıda", "quantity": 15, "unit_type": "kg", "unit_price": 50, "vat_rate": 1, "expiry_date": ""}],
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, inv_json, "gemini-3-flash-preview")):
            ok_i, res_i, _ = gemini_assistant.parse_cafeteria_invoice_with_gemini(
                api_key="k", model="gemini-3-flash-preview", raw_text="15 kg pirinç 50 TL",
            )
        self.assertTrue(ok_i)
        self.assertEqual(res_i["supplier_name"], "Güneş Gıda")

        pay_json = json.dumps({
            "payments": [{"tx_date": "2026-10-12", "for_person_name": "Ali Kaya", "delivered_by_name": "Anne Fatma", "amount": 1200, "payment_method": "Nakit", "notes": "Ekim"}]
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, pay_json, "gemini-3-flash-preview")):
            ok_p, res_p, _ = gemini_assistant.parse_cafeteria_menu_or_payments_text_with_gemini(
                api_key="k", model="gemini-3-flash-preview", task_mode="payments", raw_text="Ali Kaya annesi Fatma 1200 TL",
            )
        self.assertTrue(ok_p)
        self.assertEqual(res_p["payments"][0]["delivered_by_name"], "Anne Fatma")

        # 2. Verify ALL 5 mutating chat export actions MUST go through confirmation card gate before DB mutation
        actions_sequence = [
            (
                {"action": "cafeteria_upsert_dishes", "dishes": [{"name": "AI Yayla Çorbası", "category": "Çorba", "portion_g": 240, "ingredients": [{"ingredient_name": "Yoğurt", "amount": 60, "unit": "g", "kcal_portion": 45, "carbs_portion_g": 4, "allergen_tag": "Süt ve Süt Ürünleri (Laktoz)"}]}]},
                False,
                lambda c: c.execute("SELECT COUNT(*) FROM cafeteria_dishes WHERE name='AI Yayla Çorbası'").fetchone()[0],
            ),
            (
                {"action": "cafeteria_plan_menu", "plans": [{"plan_date": "2026-10-12", "meal_slot": "Öğle Yemeği", "dishes": [{"name": "AI Yayla Çorbası"}]}]},
                False,
                lambda c: c.execute("SELECT COUNT(*) FROM cafeteria_menu_plans WHERE plan_date='2026-10-12'").fetchone()[0],
            ),
            (
                {"action": "cafeteria_record_purchase", "invoice_date": "2026-10-12", "supplier_name": "AI Market", "category": "Süt & Kahvaltılık", "items": [{"item_name": "Süzme Yoğurt", "quantity": 10, "unit_type": "kg", "unit_price": 75, "vat_rate": 1}]},
                False,
                lambda c: c.execute("SELECT COUNT(*) FROM cafeteria_purchases WHERE supplier_name='AI Market'").fetchone()[0],
            ),
            (
                {"action": "cafeteria_record_cash", "transactions": [{"tx_date": "2026-10-12", "tx_type": "IN_PAYMENT", "for_person_name": "Ali Kaya", "delivered_by_name": "Baba Veli", "amount": 950.0}]},
                False,
                lambda c: c.execute("SELECT COUNT(*) FROM cafeteria_cash_ledger WHERE delivered_by_name='Baba Veli'").fetchone()[0],
            ),
        ]

        for payload, exp_destr, count_fn in actions_sequence:
            with self.connect() as c:
                before_cnt = count_fn(c)
            with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
                 mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Öneri hazır.", "gemini-3-flash-preview", "", payload, None)):
                res_chat = persantil._handle_floating_gemini_chat({"message": "İşlemi yap"}, [])
            self.assertTrue(res_chat["ok"])
            self.assertIsNotNone(res_chat.get("write_proposal"), f"Expected write_proposal for {payload['action']}")
            self.assertEqual(bool(res_chat["write_proposal"]["is_destructive"]), exp_destr)
            with self.connect() as c:
                self.assertEqual(count_fn(c), before_cnt, f"{payload['action']} mutated DB before confirmation!")
            apply_r = persantil._handle_floating_gemini_apply_write({"token": res_chat["write_proposal"]["token"]})
            self.assertTrue(apply_r["ok"])
            with self.connect() as c:
                self.assertGreater(count_fn(c), before_cnt)

        # Destructive delete proposal
        del_payload = {"action": "cafeteria_delete_records", "dish_names": ["AI Yayla Çorbası"], "plan_dates": [{"plan_date": "2026-10-12"}]}
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
             mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Silinecek.", "gemini-3-flash-preview", "", del_payload, None)):
            res_del = persantil._handle_floating_gemini_chat({"message": "Sil"}, [])
        self.assertTrue(res_del["write_proposal"]["is_destructive"])
        with self.connect() as c:
            self.assertEqual(c.execute("SELECT COUNT(*) FROM cafeteria_dishes WHERE name='AI Yayla Çorbası'").fetchone()[0], 1)
        apply_del = persantil._handle_floating_gemini_apply_write({"token": res_del["write_proposal"]["token"]})
        self.assertTrue(apply_del["ok"])
        with self.connect() as c:
            self.assertEqual(c.execute("SELECT COUNT(*) FROM cafeteria_dishes WHERE name='AI Yayla Çorbası'").fetchone()[0], 0)
            self.assertEqual(c.execute("SELECT COUNT(*) FROM cafeteria_menu_plans WHERE plan_date='2026-10-12'").fetchone()[0], 0)

        # Non-mutating cafeteria_export via chat
        exp_payload = {"action": "cafeteria_export", "format": "both", "year_month": "2026-10"}
        with mock.patch.object(persantil, "_get_active_gemini_config", return_value=(True, "k", "gemini-3-flash-preview")), \
             mock.patch.object(gemini_assistant, "chat_with_gemini_assistant", return_value=(True, "Dosyalar hazır.", "gemini-3-flash-preview", "", exp_payload, None)):
            res_exp = persantil._handle_floating_gemini_chat({"message": "2026-10 yemekhane çıktısı ver"}, [])
        self.assertTrue(res_exp["ok"])
        self.assertIsNone(res_exp.get("write_proposal"))
        self.assertGreaterEqual(len(res_exp.get("downloads") or []), 2)

    def test_cafeteria_ui_headless_page_and_all_five_tabs(self):
        # Seed sample rows so tables, editors & charts inside all 5 tabs render non-empty states cleanly
        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Ezogelin Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Bulgur", "amount": 40, "unit": "g", "kcal_portion": 140, "carbs_portion_g": 28, "allergen_tag": "Gluten"}],
        )
        ym = "2026-11"
        caf.upsert_menu_plan_slot(self.connect, plan_date=f"{ym}-02", meal_slot="Öğle Yemeği", dish_ids=[d_id])
        caf.upsert_purchase_invoice(
            self.connect,
            invoice_date=f"{ym}-03",
            supplier_name="Test Tedarik",
            items=[{"item_name": "Bulgur", "quantity": 10, "unit_type": "kg", "unit_price": 35, "vat_rate": 1}],
        )
        caf.import_school_people_as_diners(self.connect, [self.stu_id])
        caf.generate_or_refresh_monthly_fees(self.connect, ym, student_meal_price=100.0)
        caf.record_cash_transaction(self.connect, tx_date=f"{ym}-04", tx_type="IN_PAYMENT", amount=500.0, for_person_name="Ali Kaya", delivered_by_name="Veli")

        # Full page render via persantil.render_cafeteria_page
        out_page = revir_ui.run_action_headless(persantil.render_cafeteria_page, {"action": "render"})
        main_html = str(out_page.get("main_html") or "")
        self.assertIn("Yemekhane, Menü Planlama, Gider", main_html)
        self.assertIn("Menü Planlayıcı", main_html)
        self.assertIn("Yemekler Bankası", main_html)

        ctx_api = {
            "connect_db": persantil.connect_db,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        cfg = persantil.get_app_settings()
        for tab_fn in (
            lambda: caf._render_tab_menu_planner(ctx_api, ym, cfg),
            lambda: caf._render_tab_dishes(ctx_api),
            lambda: caf._render_tab_purchases_and_expenses(ctx_api, ym, cfg),
            lambda: caf._render_tab_diners_and_cash(ctx_api, ym, cfg),
            lambda: caf._render_tab_settings_and_exports(ctx_api, ym, cfg),
        ):
            out_tab = revir_ui.run_action_headless(tab_fn, {"action": "render"})
            self.assertIsInstance(out_tab, dict)
            self.assertNotIn("Traceback", json.dumps(out_tab, ensure_ascii=False))

    def test_recalculate_menu_on_dish_edit_diabetes_cleanup_copy_and_archive_months(self):
        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Tarhana",
            category="Çorba",
            ingredients=[{"ingredient_name": "Tarhana Tozu", "amount": 30, "unit": "g", "kcal_portion": 110, "carbs_portion_g": 20, "allergen_tag": "Gluten"}],
        )
        caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2024-02-05",
            meal_slot="Öğle Yemeği",
            dish_ids=[d_id],
            sync_to_diabetes=True,
        )
        # Update dish nutrition -> menu plan MUST recalculate automatically
        caf.upsert_dish_with_ingredients(
            self.connect,
            dish_id=d_id,
            name="Tarhana",
            category="Çorba",
            ingredients=[{"ingredient_name": "Tarhana Tozu", "amount": 60, "unit": "g", "kcal_portion": 220, "carbs_portion_g": 40, "allergen_tag": "Gluten"}],
        )
        plans = caf.get_menu_plans_for_month(self.connect, "2024-02")
        self.assertAlmostEqual(plans[("2024-02-05", "Öğle Yemeği")]["total_kcal"], 220.0, places=1)
        self.assertAlmostEqual(plans[("2024-02-05", "Öğle Yemeği")]["total_carbs_g"], 40.0, places=1)

        # Historical month outside -4..+7 MUST appear in _month_options
        ym_opts, _ = caf._month_options(self.connect)
        self.assertIn("2024-02", ym_opts)

        # Copy menu slot to another date
        copied = caf.copy_menu_plans(
            self.connect,
            date_pairs=[("2024-02-05", "2024-02-12")],
            meal_slots=["Öğle Yemeği"],
            sync_to_diabetes=True,
        )
        self.assertEqual(copied, 1)
        plans_after = caf.get_menu_plans_for_month(self.connect, "2024-02")
        self.assertIn(("2024-02-12", "Öğle Yemeği"), plans_after)

        # Clearing slot deletes both menu plan and synced diabetes row
        caf.upsert_menu_plan_slot(self.connect, plan_date="2024-02-05", meal_slot="Öğle Yemeği", dish_ids=[], custom_items=[], sync_to_diabetes=True)
        diab = idm.get_school_menus_for_month(self.connect, "2024-02")
        self.assertFalse(any(m["menu_date"] == "2024-02-05" for m in diab))
        self.assertTrue(any(m["menu_date"] == "2024-02-12" for m in diab))

        # Bracket custom dish parser
        parsed_c = caf.parse_custom_dish_entries("Tam Buğday Ekmek [130 kcal, 24g karb, gluten], Mevsim Salata")
        self.assertEqual(len(parsed_c), 2)
        self.assertEqual(parsed_c[0]["kcal"], 130.0)
        self.assertEqual(parsed_c[0]["carbs_g"], 24.0)
        self.assertEqual(parsed_c[0]["allergen"], "Gluten (Buğday/Arpa)")

    def test_clinical_celiac_matching_and_compact_weekly_grid_a4_with_logo(self):
        import base64
        with self.connect() as c:
            cls_id = int(c.execute("SELECT id FROM classes WHERE name='4-A'").fetchone()["id"])
            c.execute("INSERT INTO students(full_name, sex, class_id, person_type) VALUES ('Ece Çölyak', 'K', ?, 'Öğrenci')", (cls_id,))
            eid = int(c.execute("SELECT id FROM students WHERE full_name='Ece Çölyak'").fetchone()["id"])
            c.execute(
                "INSERT INTO health_profiles(student_id, chronic_conditions, dietary_swallowing_notes, updated_at) VALUES (?, 'Çölyak hastalığı', 'Glütensiz diyet gerekir', '2026-10-01')",
                (eid,),
            )
            c.commit()
        hits = caf.match_allergens_against_school(self.connect, ["Gluten (Buğday/Arpa)"])
        self.assertIn("Ece Çölyak", {h["full_name"] for h in hits})

        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Kıymalı Börek",
            category="Pilav / Makarna / Börek",
            ingredients=[{"ingredient_name": "Yufka", "amount": 80, "unit": "g", "kcal_portion": 290, "carbs_portion_g": 45, "allergen_tag": "Gluten"}],
        )
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-05", meal_slot="Öğle Yemeği", dish_ids=[d_id])
        # Tiny valid 1x1 PNG base64
        tiny_png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        meta = {
            "school_name": "Atatürk Ortaokulu",
            "cafeteria_school_logo_b64": tiny_png,
            "cafeteria_board_note": "* Özel alerji notu: Revire başvurun.",
        }
        for lm in ("compact_one_page", "weekly_grid"):
            html_out = caf.build_menu_a4_html(self.connect, year_month="2026-10", layout_mode=lm, school_meta=meta)
            self.assertIn("Kıymalı Börek", html_out)
            self.assertIn("data:image/png;base64", html_out)
            self.assertIn("Özel alerji notu", html_out)
            d_bytes = caf.build_menu_a4_docx(self.connect, year_month="2026-10", layout_mode=lm, school_meta=meta)
            txt = _extract_docx_all_text(d_bytes)
            self.assertIn("Kıymalı Börek", txt)
            self.assertIn("Özel alerji notu", txt)
        _ = base64  # ensure import used cleanly

    def test_invoice_edit_skt_inflation_portion_cost_and_shopping_list(self):
        from datetime import date, timedelta
        soon = (date.today() + timedelta(days=5)).isoformat()
        pid = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-09-01",
            supplier_name="Tedarik A",
            category="Bakliyat & Kuru Gıda",
            payment_status="Vadeli / Bekliyor",
            items=[{"item_name": "Kırmızı Mercimek", "quantity": 10.0, "unit_type": "kg", "unit_price": 50.0, "vat_rate": 0.0, "expiry_date": soon}],
        )
        stats_p = caf.get_expense_summary_stats(self.connect, "2026-09")
        self.assertAlmostEqual(stats_p["month_pending_total"], 500.0, places=2)

        caf.update_purchase_payment_status(self.connect, pid, "Ödendi")
        loaded = caf.get_purchase_by_id(self.connect, pid)
        self.assertEqual(loaded["payment_status"], "Ödendi")

        # Second invoice with higher price -> inflation + cost estimate
        caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-10-01",
            supplier_name="Tedarik B",
            category="Bakliyat & Kuru Gıda",
            items=[{"item_name": "Kırmızı Mercimek", "quantity": 10.0, "unit_type": "kg", "unit_price": 75.0, "vat_rate": 0.0}],
        )
        inf = caf.get_ingredient_price_inflation(self.connect)
        self.assertTrue(any(x["item_name"] == "Kırmızı Mercimek" and x["pct_change"] == 50.0 for x in inf))

        exp_items = caf.get_expiring_purchase_items(self.connect, within_days=30)
        self.assertTrue(any(e["item_name"] == "Kırmızı Mercimek" for e in exp_items))

        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Mercimek Yemeği",
            ingredients=[{"ingredient_name": "Kırmızı Mercimek", "amount": 100.0, "unit": "g", "kcal_portion": 340.0, "carbs_portion_g": 60.0}],
        )
        cost_r = caf.estimate_dish_portion_cost(self.connect, d_id)
        # 100g = 0.1kg * 75 TL/kg = 7.50 TL
        self.assertAlmostEqual(cost_r["estimated_cost_tl"], 7.5, places=2)

        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-05", meal_slot="Öğle Yemeği", dish_ids=[d_id])
        shop = caf.build_menu_shopping_list(self.connect, "2026-10", diner_count=50)
        # 50 * 100g = 5 kg -> 5 * 75 = 375 TL
        self.assertAlmostEqual(shop["estimated_total_tl"], 375.0, places=2)
        xl_bytes = caf.build_shopping_list_excel(shop, "2026-10")
        self.assertGreater(len(xl_bytes), 2000)

    def test_diner_deactivate_carry_forward_bulk_sibling_receipt_and_budget(self):
        caf.import_school_people_as_diners(self.connect, [self.stu_id, self.teacher_id])
        caf.generate_or_refresh_monthly_fees(self.connect, "2026-10", student_meal_price=100.0, staff_meal_price=120.0, workdays_override=10)
        diners = {d["full_name"]: int(d["id"]) for d in caf.get_diners(self.connect)}
        stu_did, tch_did = diners["Ali Kaya"], diners["Zeynep Öğretmen"]

        # Ali expected = 1000 TL; pays 400 -> 600 remaining
        tx_ids = caf.record_bulk_sibling_payments(
            self.connect,
            tx_date="2026-10-10",
            delivered_by_name="Veli Kaya",
            allocations=[
                {"diner_id": stu_did, "amount": 400.0},
                {"diner_id": tch_did, "amount": 1200.0},
            ],
            receipt_no="MKB-BULK-1",
            year_month="2026-10",
        )
        self.assertEqual(len(tx_ids), 2)

        # Receipt docx
        rcpt_bytes = caf.build_payment_receipt_docx(self.connect, tx_id=tx_ids[0], school_meta={"school_name": "Test Okulu"})
        rcpt_txt = _extract_docx_all_text(rcpt_bytes)
        self.assertIn("Ali Kaya", rcpt_txt)
        self.assertIn("Veli Kaya", rcpt_txt)
        self.assertIn("600.00", rcpt_txt.replace(",", ""))

        # Deleting diner with fees MUST raise; deactivate MUST preserve fees
        with self.assertRaises(ValueError):
            caf.delete_diner_by_id(self.connect, stu_did, force=False)

        # Generate November fees & carry forward unpaid 600 TL from October
        caf.generate_or_refresh_monthly_fees(self.connect, "2026-11", student_meal_price=100.0, staff_meal_price=120.0, workdays_override=10)
        cf_n = caf.carry_forward_unpaid_balances(self.connect, "2026-10", "2026-11")
        self.assertGreaterEqual(cf_n, 1)
        cash_nov = caf.get_cash_summary_and_diner_balances(self.connect, "2026-11")
        ali_nov = next(b for b in cash_nov["balances"] if b["full_name"] == "Ali Kaya")
        self.assertAlmostEqual(ali_nov["carry_forward_debt"], 600.0, places=2)
        self.assertAlmostEqual(ali_nov["expected_amount"], 1600.0, places=2)

        caf.deactivate_diner_by_id(self.connect, tch_did, end_date="2026-10-31")
        self.assertEqual(len(caf.get_diners(self.connect, active_only=True)), 1)
        caf.reactivate_diner_by_id(self.connect, tch_did)
        self.assertEqual(len(caf.get_diners(self.connect, active_only=True)), 2)

        # Budget P&L check
        caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-10-11",
            supplier_name="Tedarik",
            items=[{"item_name": "Et", "quantity": 5, "unit_type": "kg", "unit_price": 200.0, "vat_rate": 0.0}],
        )
        pnl = caf.get_budget_and_pnl_summary(self.connect, "2026-10")
        self.assertAlmostEqual(pnl["month_collected"], 1600.0, places=2)
        self.assertAlmostEqual(pnl["month_expenses"], 1000.0, places=2)
        self.assertAlmostEqual(pnl["month_net_cash_vs_exp"], 600.0, places=2)

    def test_new_gemini_wrappers_and_cmd_search(self):
        audit_j = json.dumps({
            "balance_score": 90,
            "overall_summary": "Dengeli menü",
            "risk_findings": [{"plan_date": "2026-10-05", "meal_slot": "Öğle Yemeği", "risk_level": "Orta", "issue": "Gluten", "affected_groups": "Çölyak", "recommended_alternative": "Glütensiz çorba"}],
            "weekly_nutrition_notes": ["Baklagil dengesi iyi"],
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, audit_j, "gemini-3-flash-preview")):
            ok1, r1, _ = gemini_assistant.audit_cafeteria_menu_with_gemini(
                api_key="k", model="gemini-3-flash-preview", year_month="2026-10",
                planned_menu_summary=[], clinical_risk_summary={}, available_dishes=[],
            )
        self.assertTrue(ok1)
        self.assertEqual(r1["balance_score"], 90)

        batch_j = json.dumps({
            "enriched_dishes": [{"id": 1, "name": "Yayla Çorbası", "category": "Çorba", "portion_g": 250, "recipe_notes": "", "ingredients": []}]
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, batch_j, "gemini-3-flash-preview")):
            ok2, r2, _ = gemini_assistant.batch_enrich_dishes_with_gemini(api_key="k", model="gemini-3-flash-preview", dishes_to_enrich=[{"id": 1, "name": "Yayla Çorbası"}])
        self.assertTrue(ok2)
        self.assertEqual(len(r2["enriched_dishes"]), 1)

        fin_j = json.dumps({
            "executive_summary": "Gelir-gider dengeli",
            "inflation_alerts": ["Yağ +%15"],
            "cost_saving_actions": ["Mevsim sebzesi"],
            "parent_reminder_template": "Sayın Velimiz...",
        }, ensure_ascii=False)
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, fin_j, "gemini-3-flash-preview")):
            ok3, r3, _ = gemini_assistant.analyze_cafeteria_finances_and_inflation_with_gemini(
                api_key="k", model="gemini-3-flash-preview", year_month="2026-10",
                budget_summary={}, price_changes=[], unpaid_summary=[],
            )
        self.assertTrue(ok3)
        self.assertIn("Sayın Velimiz", r3["parent_reminder_template"])

        caf.upsert_dish_with_ingredients(self.connect, name="Hünkar Beğendi", ingredients=[])
        cmd_hits = persantil._handle_cmd_search("Hünkar")
        self.assertTrue(any(h.get("title") == "Hünkar Beğendi" and h.get("value") == "🍽️ Yemekhane Yönetimi" for h in cmd_hits))

    def test_dishes_table_no_cost_col_and_popup_dialog_and_schedule_skip_empty_a4(self):
        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Yayla Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Yoğurt", "amount": 60, "unit": "g", "kcal_portion": 45, "carbs_portion_g": 4, "allergen_tag": "Süt"}],
        )
        # 1. Dishes tab must NOT include Tahmini Maliyet column in Kayıtlı Yemekler Listesi
        ctx_api = {
            "connect_db": persantil.connect_db,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        out_dishes = revir_ui.run_action_headless(lambda: caf._render_tab_dishes(ctx_api), {"action": "render"})
        self.assertNotIn("Tahmini Maliyet (₺)", str(out_dishes.get("main_html") or ""))

        # 2. Pop-up editor dialog renders all 5 planned sections cleanly
        out_dlg = revir_ui.run_action_headless(lambda: caf.open_dish_editor_dialog(ctx_api, d_id), {"action": "render"})
        dlg_html = str(out_dlg.get("dialog_html") or "")
        for sec_marker in ("Temel Yemek Bilgileri", "Yapay Zekâ", "Malzeme Listesi", "Besin Özeti", "Tarif / Servis Notu"):
            self.assertIn(sec_marker, dlg_html)

        # 3. Settings: 6 days/week (Mon-Sat), 2 meals/day, custom meal hours -> reflected in workdays & A4 outputs
        persantil.save_app_settings({
            "cafeteria_service_days_per_week": "6",
            "cafeteria_service_weekdays": "0,1,2,3,4,5",
            "cafeteria_meals_per_day": "2",
            "cafeteria_active_meal_slots": "Sabah Kahvaltısı,Öğle Yemeği",
            "cafeteria_time_breakfast": "08:15 - 08:45",
            "cafeteria_time_lunch": "12:15 - 13:00",
        })
        persantil.clear_all_data_caches()

        days_oct = caf.get_month_workdays_info(self.connect, "2026-10")
        by_iso = {d["date_iso"]: d for d in days_oct}
        self.assertTrue(by_iso["2026-10-03"]["is_workday"])   # Saturday included when 6 days/week configured
        self.assertFalse(by_iso["2026-10-04"]["is_workday"])  # Sunday excluded

        # 4. Plan ONLY 2026-10-05; leave 2026-10-06..31 empty -> A4 MUST omit empty days & show schedule + hours
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-05", meal_slot="Öğle Yemeği", dish_ids=[d_id])
        cfg_now = persantil.get_app_settings()
        a4_html = caf.build_menu_a4_html(self.connect, year_month="2026-10", school_meta=cfg_now)
        self.assertIn("05.10.2026", a4_html)
        self.assertNotIn("06.10.2026", a4_html)
        self.assertNotIn("07.10.2026", a4_html)
        self.assertIn("6 Gün", a4_html)
        self.assertIn("2 Öğün", a4_html)
        self.assertIn("08:15 - 08:45", a4_html)
        self.assertIn("12:15 - 13:00", a4_html)

        a4_docx = caf.build_menu_a4_docx(self.connect, year_month="2026-10", school_meta=cfg_now)
        docx_txt = _extract_docx_all_text(a4_docx)
        self.assertIn("05.10.2026", docx_txt)
        self.assertNotIn("06.10.2026", docx_txt)
        self.assertNotIn("07.10.2026", docx_txt)
        self.assertIn("6 Gün", docx_txt)
        self.assertIn("08:15 - 08:45", docx_txt)
        self.assertIn("12:15 - 13:00", docx_txt)

        # 5. Completely empty month returns clean empty message without blank rows
        emp_html = caf.build_menu_a4_html(self.connect, year_month="2026-12", school_meta=cfg_now)
        self.assertIn("öğün girilmiş gün bulunamadı", emp_html)
        emp_docx = _extract_docx_all_text(caf.build_menu_a4_docx(self.connect, year_month="2026-12", school_meta=cfg_now))
        self.assertIn("öğün girilmiş gün bulunmamaktadır", emp_docx)


if __name__ == "__main__":
    unittest.main()
