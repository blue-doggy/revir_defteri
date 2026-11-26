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
        self.assertNotIn("Aylar Arası Planlama Merkezi", main_html)
        self.assertLess(main_html.index("Aktif Ay"), main_html.index("Düzenlenecek Haftayı Seçin"))

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
        revir_ui.session_state["caf_menu_week_month"] = "2026-10"
        revir_ui.session_state["caf_menu_wk_pick"] = "5. Hafta (26 – 30 Ekim)"
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
        self.assertEqual(revir_ui.session_state["caf_menu_week_month"], ym)
        self.assertTrue(revir_ui.session_state["caf_menu_wk_pick"].startswith("1. Hafta"))

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

    def test_ramazan_mode_province_iftar_sahur_in_menu_and_exports(self):
        # 1. Province normalization (ASCII Turkish tolerant + safe fallback)
        self.assertEqual(caf.resolve_province_name("istanbul"), "İstanbul")
        self.assertEqual(caf.resolve_province_name("Kahramanmaras"), "Kahramanmaraş")
        self.assertEqual(caf.resolve_province_name("BilinmeyenIl", fallback="Ankara"), "Ankara")

        # 2. Offline prayer calculation produces deterministic province-specific Sahur (İmsak) & İftar (Akşam)
        with mock.patch("urllib.request.urlopen", side_effect=OSError("offline")):
            van_t = caf.get_ramazan_times_for_dates(self.connect, "Van", ["2026-03-02"])
            ist_t = caf.get_ramazan_times_for_dates(self.connect, "İstanbul", ["2026-03-02"])

        self.assertIn("2026-03-02", van_t)
        self.assertIn("2026-03-02", ist_t)
        van_sahur, van_iftar = van_t["2026-03-02"]["sahur"], van_t["2026-03-02"]["iftar"]
        ist_sahur, ist_iftar = ist_t["2026-03-02"]["sahur"], ist_t["2026-03-02"]["iftar"]
        for tm in (van_sahur, van_iftar, ist_sahur, ist_iftar):
            self.assertRegex(tm, r"^\d{2}:\d{2}$")
        # Eastern Van always has earlier Sahur and İftar than western İstanbul
        self.assertLess(van_sahur, ist_sahur)
        self.assertLess(van_iftar, ist_iftar)

        # 3. Enable Ramazan Modu in app settings and verify schedule parser exposes İftar & Sahur slots
        persantil.save_app_settings({
            "school_name": "Örnek Anadolu Lisesi",
            "cafeteria_ramazan_mode": "1",
            "cafeteria_ramazan_province": "İstanbul",
            "cafeteria_service_days_per_week": "5",
            "cafeteria_service_weekdays": "0,1,2,3,4",
            "cafeteria_meals_per_day": "2",
            "cafeteria_active_meal_slots": "İftar,Sahur",
        })
        cfg_ram = persantil.get_app_settings()
        sched = caf.parse_service_schedule_from_cfg(cfg_ram)
        self.assertTrue(sched["ramazan_mode"])
        self.assertEqual(sched["ramazan_province"], "İstanbul")
        self.assertEqual(sched["available_slots"], caf.ALL_MEAL_SLOTS)
        self.assertEqual(sched["active_meal_slots"], ["İftar", "Sahur"])

        # Disabling Ramazan Modu hides İftar/Sahur cleanly
        sched_off = caf.parse_service_schedule_from_cfg({**cfg_ram, "cafeteria_ramazan_mode": "0"})
        self.assertFalse(sched_off["ramazan_mode"])
        self.assertEqual(sched_off["available_slots"], caf.MEAL_SLOTS)
        self.assertNotIn("İftar", sched_off["active_meal_slots"])

        # 4. Create dishes and plan İftar + Sahur on 2026-03-02
        d_iftar = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Ramazan Mercimek Çorbası",
            category="Çorba",
            portion_g=250.0,
            ingredients=[
                {"ingredient_name": "Mercimek", "amount": 60.0, "unit": "g", "kcal_portion": 200.0, "carbs_portion_g": 34.0, "allergen_tag": ""},
            ],
            source="manual",
        )
        d_sahur = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Sahur Kahvaltı Tabağı",
            category="Kahvaltı",
            portion_g=220.0,
            ingredients=[
                {"ingredient_name": "Beyaz Peynir", "amount": 50.0, "unit": "g", "kcal_portion": 145.0, "carbs_portion_g": 1.5, "allergen_tag": "Süt"},
            ],
            source="manual",
        )
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-03-02", meal_slot="İftar", dish_ids=[d_iftar])
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-03-02", meal_slot="Sahur", dish_ids=[d_sahur])

        # 5. A4 HTML (standard & weekly_grid), A4 Word (standard & weekly_grid), and Excel workbook include İftar/Sahur & times
        with mock.patch("urllib.request.urlopen", side_effect=OSError("offline")):
            a4_std_html = caf.build_menu_a4_html(self.connect, year_month="2026-03", layout_mode="standard", school_meta=cfg_ram)
            a4_wk_html = caf.build_menu_a4_html(self.connect, year_month="2026-03", layout_mode="weekly_grid", school_meta=cfg_ram)
            a4_std_docx = caf.build_menu_a4_docx(self.connect, year_month="2026-03", layout_mode="standard", school_meta=cfg_ram)
            a4_wk_docx = caf.build_menu_a4_docx(self.connect, year_month="2026-03", layout_mode="weekly_grid", school_meta=cfg_ram)
            wb_bytes = caf.build_cafeteria_full_workbook_excel(self.connect, "2026-03")

        for html_out in (a4_std_html, a4_wk_html):
            self.assertIn("Ramazan Modu", html_out)
            self.assertIn("İstanbul", html_out)
            self.assertIn(ist_sahur, html_out)
            self.assertIn(ist_iftar, html_out)
            self.assertIn("Ramazan Mercimek Çorbası", html_out)
            self.assertIn("Sahur Kahvaltı Tabağı", html_out)

        for raw_docx in (a4_std_docx, a4_wk_docx):
            d_txt = _extract_docx_all_text(raw_docx)
            self.assertIn("Ramazan Modu", d_txt)
            self.assertIn("İstanbul", d_txt)
            self.assertIn(ist_sahur, d_txt)
            self.assertIn(ist_iftar, d_txt)
            self.assertIn("Ramazan Mercimek Çorbası", d_txt)
            self.assertIn("Sahur Kahvaltı Tabağı", d_txt)

        wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=True)
        try:
            ws = wb["Aylik_Menu"]
            rows = list(ws.iter_rows(values_only=True))
            headers = list(rows[0])
            self.assertIn("Sahur (İmsak)", headers)
            self.assertIn("İftar", headers)
            sahur_idx = headers.index("Sahur (İmsak)")
            iftar_idx = headers.index("İftar")
            data_rows = [r for r in rows[1:] if r and r[0] == "2026-03-02"]
            self.assertTrue(data_rows)
            self.assertEqual(data_rows[0][sahur_idx], ist_sahur)
            self.assertEqual(data_rows[0][iftar_idx], ist_iftar)
        finally:
            wb.close()

        # 6. Headless UI rendering of Menu Planner & Settings tabs with Ramazan Modu active
        revir_ui.reset_session()
        revir_ui.session_state["caf_global_sel_ym"] = "2026-03"
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        with mock.patch("urllib.request.urlopen", side_effect=OSError("offline")):
            out_menu = revir_ui.run_action_headless(
                lambda: caf._render_tab_menu_planner(ctx_api, "2026-03", cfg_ram),
                {"action": "render"},
            )
            out_cfg = revir_ui.run_action_headless(
                lambda: caf._render_tab_settings_and_exports(ctx_api, "2026-03", cfg_ram),
                {"action": "render"},
            )
        self.assertNotIn("Çalışma Zamanı Hatası", out_menu["main_html"])
        self.assertIn("Ramazan Modu", out_menu["main_html"])
        self.assertIn(ist_sahur, out_menu["main_html"])
        self.assertIn(ist_iftar, out_menu["main_html"])
        self.assertNotIn("Çalışma Zamanı Hatası", out_cfg["main_html"])
        self.assertIn("Ramazan Modu", out_cfg["main_html"])

    def test_a4_weekly_colors_alternating_tones_and_inline_dishes_with_carbs(self):
        d1 = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Tarhana Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Tarhana", "amount": 40, "unit": "g", "kcal_portion": 150, "carbs_portion_g": 22, "allergen_tag": "Gluten"}],
        )
        d2 = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Yoğurtlu Semizotu",
            category="Salata / Meze",
            ingredients=[{"ingredient_name": "Yoğurt", "amount": 80, "unit": "g", "kcal_portion": 95, "carbs_portion_g": 8, "allergen_tag": "Süt"}],
        )
        d3 = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Izgara Tavuk",
            category="Ana Yemek",
            ingredients=[{"ingredient_name": "Tavuk Göğüs", "amount": 120, "unit": "g", "kcal_portion": 220, "carbs_portion_g": 0, "allergen_tag": ""}],
        )
        # Week 1: Oct 1 & 2; Week 2: Oct 5 & 6; Week 3: Oct 12
        for dt_iso in ("2026-10-01", "2026-10-02", "2026-10-05", "2026-10-06"):
            caf.upsert_menu_plan_slot(self.connect, plan_date=dt_iso, meal_slot="Öğle Yemeği", dish_ids=[d1, d2])
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-12", meal_slot="Öğle Yemeği", dish_ids=[d3])

        # 1. Style map uses much lighter default pastel tones + consecutive day light/dark alternation
        p_w1 = caf.get_week_palette_pair(0)
        p_w2 = caf.get_week_palette_pair(1)
        p_w3 = caf.get_week_palette_pair(2)
        styles, _ = caf._build_week_day_style_map([
            {"date_iso": "2026-10-01"},
            {"date_iso": "2026-10-02"},
            {"date_iso": "2026-10-05"},
            {"date_iso": "2026-10-06"},
            {"date_iso": "2026-10-12"},
        ])
        self.assertEqual(styles["2026-10-01"]["bg_hex"], p_w1["light"])  # Week 1 Light Yellow
        self.assertEqual(styles["2026-10-02"]["bg_hex"], p_w1["dark"])   # Week 1 Dark Yellow
        self.assertEqual(styles["2026-10-05"]["bg_hex"], p_w2["light"])  # Week 2 Light Orange
        self.assertEqual(styles["2026-10-06"]["bg_hex"], p_w2["dark"])   # Week 2 Dark Orange
        self.assertEqual(styles["2026-10-12"]["bg_hex"], p_w3["light"])  # Week 3 Light Green

        # 2. HTML outputs: no week-color legend, "Öğün" instead of "Menü Günü", inline dishes + carbs + lighter colors
        for lm in ("standard", "compact_one_page", "weekly_grid"):
            html_out = caf.build_menu_a4_html(self.connect, year_month="2026-10", layout_mode=lm)
            self.assertIn("print-color-adjust: exact", html_out)
            self.assertNotIn("Hafta Renkleri", html_out)
            self.assertNotIn("Menü Günü", html_out)
            self.assertIn("5 Öğün", html_out)
            self.assertIn("Tarhana Çorbası · Yoğurtlu Semizotu", html_out)
            self.assertIn("245 kcal", html_out)
            self.assertIn("30 g Karb.", html_out)
            self.assertIn("Alerjen: Yok", html_out)
            for hex_c in (p_w1["light"], p_w1["dark"], p_w2["light"], p_w2["dark"], p_w3["light"]):
                self.assertIn(hex_c, html_out)

        # 3. Word (.docx) outputs: no week-color legend, "Öğün" instead of "Menü Günü", inline dishes + XML shading
        for lm in ("standard", "compact_one_page", "weekly_grid"):
            raw_doc = caf.build_menu_a4_docx(self.connect, year_month="2026-10", layout_mode=lm)
            txt = _extract_docx_all_text(raw_doc)
            self.assertNotIn("Hafta Renkleri", txt)
            self.assertNotIn("Menü Günü", txt)
            self.assertIn("5 Öğün", txt)
            self.assertIn("Tarhana Çorbası · Yoğurtlu Semizotu", txt)
            self.assertIn("Tarhana Çorbası · Yoğurtlu Semizotu\nKalori: 245 kcal  |  Karb.: 30 g  |  Alerjen:", txt)
            self.assertIn("Izgara Tavuk\nKalori: 220 kcal  |  Karb.: 0 g  |  Alerjen: Yok", txt)
            d_obj = docx.Document(io.BytesIO(raw_doc))
            shd_fills = {
                el.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
                for tbl in d_obj.tables
                for row in tbl.rows
                for cell in row.cells
                for el in cell._tc.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd")
            }
            for w_hex in (p_w1["light"], p_w1["dark"], p_w2["light"], p_w2["dark"], p_w3["light"]):
                self.assertIn(w_hex.lstrip("#").upper(), shd_fills)

        # 4. Excel Aylik_Menu rows colored per week & alternating days + inline dishes
        wb_b = caf.build_cafeteria_full_workbook_excel(self.connect, "2026-10")
        wb = openpyxl.load_workbook(io.BytesIO(wb_b))
        try:
            ws = wb["Aylik_Menu"]
            fills_by_date = {}
            dishes_by_date = {}
            for row in ws.iter_rows(min_row=2):
                dt_v = str(row[0].value or "")[:10]
                fills_by_date[dt_v] = str(row[0].fill.fgColor.rgb or "")[-6:]
                dishes_by_date[dt_v] = str(row[2].value or "")
            self.assertEqual(fills_by_date["2026-10-01"], p_w1["light"].lstrip("#").upper())
            self.assertEqual(fills_by_date["2026-10-02"], p_w1["dark"].lstrip("#").upper())
            self.assertEqual(fills_by_date["2026-10-05"], p_w2["light"].lstrip("#").upper())
            self.assertEqual(fills_by_date["2026-10-06"], p_w2["dark"].lstrip("#").upper())
            self.assertEqual(fills_by_date["2026-10-12"], p_w3["light"].lstrip("#").upper())
            self.assertEqual(dishes_by_date["2026-10-01"], "Tarhana Çorbası · Yoğurtlu Semizotu")
        finally:
            wb.close()

        # 5. Headless UI menu planner exposes Karbonhidrat toggle & renders without error
        revir_ui.reset_session()
        revir_ui.session_state["caf_global_sel_ym"] = "2026-10"
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        out_m = revir_ui.run_action_headless(
            lambda: caf._render_tab_menu_planner(ctx_api, "2026-10", persantil.get_app_settings()),
            {"action": "render"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", out_m["main_html"])
        self.assertIn("Karbonhidrat (g) Göster", out_m["main_html"])

    def test_menu_planner_week_arrow_navigation_top_and_bottom(self):
        revir_ui.reset_session()
        revir_ui.session_state["caf_global_sel_ym"] = "2026-10"
        cfg_now = persantil.get_app_settings()
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        render_fn = lambda: caf._render_tab_menu_planner(ctx_api, "2026-10", cfg_now)

        # Initial render: Week 1 active, top & bottom left/right arrows rendered
        out1 = revir_ui.run_action_headless(render_fn, {"action": "render"})
        h1 = out1["main_html"]
        self.assertNotIn("Çalışma Zamanı Hatası", h1)
        self.assertEqual(h1.count("◀ Önceki Hafta"), 2)
        self.assertEqual(h1.count("Sonraki Hafta ▶"), 2)
        self.assertIn("caf_btn_prev_wk_top", h1)
        self.assertIn("caf_btn_next_wk_top", h1)
        self.assertIn("caf_btn_prev_wk_bot", h1)
        self.assertIn("caf_btn_next_wk_bot", h1)
        self.assertIn("01.10.2026", h1)
        self.assertTrue(str(revir_ui.session_state.get("caf_menu_wk_pick")).startswith("1. Hafta"))

        # Click TOP right arrow -> switches to Week 2
        out2 = revir_ui.run_action_headless(render_fn, {"action": "click", "key": "caf_btn_next_wk_top"})
        self.assertTrue(str(revir_ui.session_state.get("caf_menu_wk_pick")).startswith("2. Hafta"))
        self.assertIn("05.10.2026", out2["main_html"])

        # Click BOTTOM right arrow -> switches to Week 3
        out3 = revir_ui.run_action_headless(render_fn, {"action": "click", "key": "caf_btn_next_wk_bot"})
        self.assertTrue(str(revir_ui.session_state.get("caf_menu_wk_pick")).startswith("3. Hafta"))
        self.assertIn("12.10.2026", out3["main_html"])

        # Click BOTTOM left arrow -> returns to Week 2
        out4 = revir_ui.run_action_headless(render_fn, {"action": "click", "key": "caf_btn_prev_wk_bot"})
        self.assertTrue(str(revir_ui.session_state.get("caf_menu_wk_pick")).startswith("2. Hafta"))
        self.assertIn("05.10.2026", out4["main_html"])

        # Click TOP left arrow -> returns to Week 1
        out5 = revir_ui.run_action_headless(render_fn, {"action": "click", "key": "caf_btn_prev_wk_top"})
        self.assertTrue(str(revir_ui.session_state.get("caf_menu_wk_pick")).startswith("1. Hafta"))
        self.assertIn("01.10.2026", out5["main_html"])

    def test_invoice_click_popup_edit_add_remove_and_selectable_statistics(self):
        # 1. Create two multi-item invoices and verify stored items + summary on purchases list
        pid_oct = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-10-05",
            invoice_no="FTR-1001",
            supplier_name="Anadolu Toptan",
            category="Bakliyat & Kuru Gıda",
            items=[
                {"item_name": "Kırmızı Mercimek", "category": "Bakliyat & Kuru Gıda", "quantity": 10.0, "unit_type": "kg", "unit_price": 45.0, "vat_rate": 0.0},
            ],
        )
        pid_nov = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-11-10",
            invoice_no="FTR-1102",
            supplier_name="Güneş Gıda",
            category="Bakliyat & Kuru Gıda",
            items=[
                {"item_name": "Pirinç", "category": "Bakliyat & Kuru Gıda", "quantity": 10.0, "unit_type": "kg", "unit_price": 50.0, "vat_rate": 0.0},
                {"item_name": "Ayçiçek Yağı", "category": "Yağ & Baharat", "quantity": 5.0, "unit_type": "litre", "unit_price": 80.0, "vat_rate": 0.0},
            ],
        )
        self.assertGreater(pid_oct, 0)
        self.assertGreater(pid_nov, 0)

        p_all, _ = caf.get_purchases_and_items(self.connect, year_month="Tümü")
        by_pid = {int(p["id"]): p for p in p_all}
        self.assertEqual(by_pid[pid_nov]["item_count"], 2)
        self.assertIn("Pirinç", by_pid[pid_nov]["items_summary"])
        self.assertIn("50 ₺", by_pid[pid_nov]["items_summary"])

        # 2. Headless UI: clicking row 0 (pid_nov, newest) opens invoice pop-up modal
        revir_ui.reset_session()
        revir_ui.session_state["caf_purch_scope_flt"] = "Tüm Tarihler"
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        tab_fn = lambda: caf._render_tab_purchases_and_expenses(ctx_api, "2026-11", persantil.get_app_settings())
        out_init = revir_ui.run_action_headless(tab_fn, {"action": "render"})
        self.assertNotIn("Çalışma Zamanı Hatası", out_init["main_html"])
        self.assertIn("Kayıtlı Faturalar", out_init["main_html"])
        self.assertIn("FTR-1102", out_init["main_html"])

        out_open = revir_ui.run_action_headless(
            tab_fn,
            {"action": "df_select", "key": "caf_invoices_list_tbl", "value": {"rows": [0], "cells": [[0, "Fatura No"]]}},
        )
        dlg_html = str(out_open.get("dialog_html") or "")
        self.assertIn("Pirinç", dlg_html)
        self.assertIn("Ayçiçek Yağı", dlg_html)
        self.assertIn("caf_dlg_i_items_editor", dlg_html)

        # Test "+ Yeni Ürün Satırı Ekle" keeps dialog open cleanly
        out_add_row = revir_ui.run_action_headless(tab_fn, {"action": "click", "key": "caf_dlg_i_btn_add_row"})
        self.assertIn("caf_dlg_i_items_editor", str(out_add_row.get("dialog_html") or ""))

        # Re-open fresh dialog cleanly to apply name/price edit, deletion, and new item addition
        revir_ui.run_action_headless(tab_fn, {"action": "close_dialog", "all": True})
        revir_ui.run_action_headless(
            tab_fn,
            {"action": "df_select", "key": "caf_invoices_list_tbl", "value": {"rows": [0], "cells": [[0, "Fatura No"]]}},
        )
        # Rename row 0 -> Osmancık Pirinç @ 62.0, delete row 1 (Ayçiçek Yağı), add Tereyağı @ 300.0
        for mut in (
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Ürün Adı", "val": "Osmancık Pirinç"},
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Birim Fiyat (₺)", "val": 62.0},
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "KDV (%)", "val": 0.0},
            {"op": "delete_row", "row_kind": "base", "row_id": 1},
            {"op": "add_row"},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "Ürün Adı", "val": "Tereyağı"},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "Kategori", "val": "Süt & Kahvaltılık"},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "Miktar", "val": 2.0},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "Miktar Çeşidi", "val": "kg"},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "Birim Fiyat (₺)", "val": 300.0},
            {"op": "cell", "row_kind": "added", "row_id": 0, "col": "KDV (%)", "val": 0.0},
        ):
            revir_ui.run_action_headless(tab_fn, {"action": "data_editor", "key": "caf_dlg_i_items_editor", "value": mut})

        out_saved = revir_ui.run_action_headless(tab_fn, {"action": "click", "key": "caf_dlg_i_btn_save_inv"})
        self.assertIsNone(out_saved.get("dialog_html"))

        updated = caf.get_purchase_by_id(self.connect, pid_nov)
        self.assertIsNotNone(updated)
        u_items = {it["item_name"]: it for it in updated["items"]}
        self.assertEqual(set(u_items.keys()), {"Osmancık Pirinç", "Tereyağı"})
        self.assertAlmostEqual(u_items["Osmancık Pirinç"]["unit_price"], 62.0, places=2)
        self.assertAlmostEqual(u_items["Tereyağı"]["unit_price"], 300.0, places=2)
        # 10*62 + 2*300 = 1220.0 TL
        self.assertAlmostEqual(updated["total_amount"], 1220.0, places=2)

        # 3. Verify analytics backend + multi-sheet Excel + selectable statistics UI filtering
        caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-11-18",
            invoice_no="FTR-1109",
            supplier_name="Güneş Gıda",
            category="Bakliyat & Kuru Gıda",
            items=[
                {"item_name": "Osmancık Pirinç", "category": "Bakliyat & Kuru Gıda", "quantity": 10.0, "unit_type": "kg", "unit_price": 70.0, "vat_rate": 0.0},
            ],
        )
        an = caf.compute_purchase_analytics(self.connect, year_month="Tümü", product_name="Osmancık Pirinç")
        self.assertEqual(an["summary"]["invoice_count"], 3)
        self.assertEqual(an["summary"]["top_product_name"], "Osmancık Pirinç")
        self.assertAlmostEqual(an["summary"]["total_spend_tl"], 450.0 + 1220.0 + 700.0, places=2)
        self.assertTrue(any(p["Ürün Adı"] == "Osmancık Pirinç" and p["Toplam Miktar"] == 20.0 for p in an["product_summary"]))
        self.assertEqual(len(an["monthly_trend"]), 2)
        self.assertEqual(len(an["product_history"]), 2)
        self.assertTrue(any(c["Ürün"] == "Osmancık Pirinç" and c["Son Fiyat (₺)"] == 70.0 for c in an["price_changes"]))

        an_xlsx = caf.build_purchase_analytics_excel(an, "Tüm Tarihler")
        wb_an = openpyxl.load_workbook(io.BytesIO(an_xlsx), read_only=True)
        try:
            self.assertEqual(
                wb_an.sheetnames,
                ["Ozet_KPI", "Urun_Bazli", "Kategori_Bazli", "Tedarikci_Analizi", "Aylik_Trend", "Fiyat_Degisimi"],
            )
        finally:
            wb_an.close()

        # Select ONE view at a time via single-window selector box (caf_stat_view_pick) & verify others stay hidden
        import html as _html_m
        revir_ui.session_state["caf_stat_view_pick"] = caf.PURCHASE_STAT_VIEWS[1]
        out_cat_st = revir_ui.run_action_headless(tab_fn, {"action": "render"})
        cat_html = str(out_cat_st.get("main_html") or "")
        self.assertNotIn("Çalışma Zamanı Hatası", cat_html)
        self.assertIn(f"<h5 class='revir-h5'>{_html_m.escape(caf.PURCHASE_STAT_VIEWS[1])}</h5>", cat_html)
        self.assertNotIn(f"<h5 class='revir-h5'>{_html_m.escape(caf.PURCHASE_STAT_VIEWS[0])}</h5>", cat_html)

        revir_ui.session_state["caf_stat_view_pick"] = caf.PURCHASE_STAT_VIEWS[5]
        revir_ui.session_state["caf_stat_prod_pick"] = "Osmancık Pirinç"
        out_stats = revir_ui.run_action_headless(tab_fn, {"action": "render"})
        st_html = str(out_stats.get("main_html") or "")
        self.assertNotIn("Çalışma Zamanı Hatası", st_html)
        self.assertIn(f"<h5 class='revir-h5'>{_html_m.escape(caf.PURCHASE_STAT_VIEWS[5])}</h5>", st_html)
        self.assertIn("Osmancık Pirinç", st_html)
        self.assertNotIn(f"<h5 class='revir-h5'>{_html_m.escape(caf.PURCHASE_STAT_VIEWS[0])}</h5>", st_html)
        self.assertNotIn(f"<h5 class='revir-h5'>{_html_m.escape(caf.PURCHASE_STAT_VIEWS[2])}</h5>", st_html)

        # Month-by-month selector in stats filters cleanly to 2026-10 (450.00 TL only)
        revir_ui.session_state["caf_stat_scope"] = f"2026-10 — {caf._fmt_ym('2026-10')}"
        out_oct_st = revir_ui.run_action_headless(tab_fn, {"action": "render"})
        self.assertIn("450.00 ₺", str(out_oct_st.get("main_html") or ""))

    def test_color_tone_setting_and_right_signature_persistence_bugfix(self):
        # 1. Color tone preset configures background lightness across HTML, Word & Excel
        d_id = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Ezogelin Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Mercimek", "amount": 50, "unit": "g", "kcal_portion": 180, "carbs_portion_g": 28}],
        )
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-05", meal_slot="Öğle Yemeği", dish_ids=[d_id])
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-06", meal_slot="Öğle Yemeği", dish_ids=[d_id])

        pal_ultra = caf.get_week_palette_pair(0, color_tone="Ultra Açık Pastel")
        pal_bold = caf.get_week_palette_pair(0, color_tone="Orta Ton Pastel")
        self.assertNotEqual(pal_ultra["dark"], pal_bold["dark"])

        persantil.save_app_settings({"cafeteria_menu_color_tone": "Ultra Açık Pastel"})
        cfg_ultra = persantil.get_app_settings()
        h_ultra = caf.build_menu_a4_html(self.connect, year_month="2026-10", school_meta=cfg_ultra)
        self.assertIn(pal_ultra["light"], h_ultra)
        self.assertIn(pal_ultra["dark"], h_ultra)
        self.assertIn("2 Öğün", h_ultra)
        self.assertNotIn("Menü Günü", h_ultra)
        self.assertNotIn("Hafta Renkleri", h_ultra)

        # 2. Monochrome preset renders white / crisp light gray alternating cells
        persantil.save_app_settings({"cafeteria_menu_color_tone": "Beyaz / Renksiz"})
        h_mono = caf.build_menu_a4_html(self.connect, year_month="2026-10", school_meta=persantil.get_app_settings())
        self.assertIn("background:#FFFFFF", h_mono)
        self.assertIn("background:#F8FAFC", h_mono)

        # 3. Signature regression bugfix:
        # Nurse clears own name from Slot 1 (left) and moves own name to Slot 3 (right).
        # Must NOT re-inject nurse_name into Slot 1 on settings display or A4 HTML/Word export!
        persantil.save_app_settings({
            "nurse_name": "Hemşire Elif Kaya",
            "principal_name": "Müdür Kemal Arslan",
            "cafeteria_sig1_role": "Okul Müdürü",
            "cafeteria_sig1_name": "",
            "cafeteria_sig2_role": "",
            "cafeteria_sig2_name": "",
            "cafeteria_sig3_role": "Revir Hemşiresi",
            "cafeteria_sig3_name": "Hemşire Elif Kaya",
            "cafeteria_sigs_customized": "1",
        })
        cfg_sig = persantil.get_app_settings()
        slots = caf.resolve_menu_signer_slots(cfg_sig)
        self.assertEqual(slots[0], ("Okul Müdürü", ""))
        self.assertEqual(slots[1], ("", ""))
        self.assertEqual(slots[2], ("Revir Hemşiresi", "Hemşire Elif Kaya"))

        signers_p = caf.resolve_menu_signers_for_print(cfg_sig)
        h_sig = caf.build_menu_a4_html(self.connect, year_month="2026-10", signers=signers_p, school_meta=cfg_sig)
        sig_boxes = h_sig.split("class='sig-box'")[1:]
        self.assertEqual(len(sig_boxes), 2)
        self.assertNotIn("Hemşire Elif Kaya", sig_boxes[0])  # Left slot MUST NOT contain nurse's name!
        self.assertIn("Hemşire Elif Kaya", sig_boxes[1])     # Right slot MUST contain nurse's name!

        d_sig = _extract_docx_all_text(
            caf.build_menu_a4_docx(self.connect, year_month="2026-10", signers=signers_p, school_meta=cfg_sig)
        )
        self.assertEqual(d_sig.count("Hemşire Elif Kaya"), 1)
        self.assertLess(d_sig.find("Okul Müdürü"), d_sig.find("Hemşire Elif Kaya"))

        # Even when Slot 1 Role is ALSO cleared (only Slot 3 active), right-aligned styling is applied
        signers_right_only = [("", ""), ("", ""), ("Revir Hemşiresi", "Hemşire Elif Kaya")]
        h_r_only = caf.build_menu_a4_html(self.connect, year_month="2026-10", signers=signers_right_only, school_meta=cfg_sig)
        self.assertIn("margin-left:auto", h_r_only)
        d_r_only_bytes = caf.build_menu_a4_docx(self.connect, year_month="2026-10", signers=signers_right_only, school_meta=cfg_sig)
        d_r_doc = docx.Document(io.BytesIO(d_r_only_bytes))
        sig_tbl = d_r_doc.tables[-1]
        self.assertEqual(len(sig_tbl.columns), 2)
        self.assertEqual(sig_tbl.cell(0, 0).text.strip(), "")
        self.assertIn("Hemşire Elif Kaya", sig_tbl.cell(0, 1).text)

        # Settings UI displays empty Slot 1 name & nurse name in Slot 3 + color tone selector
        revir_ui.reset_session()
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        out_cfg = revir_ui.run_action_headless(
            lambda: caf._render_tab_settings_and_exports(ctx_api, "2026-10", cfg_sig),
            {"action": "render"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", out_cfg["main_html"])
        self.assertIn("Menü Çıktısı Renk Tonu", out_cfg["main_html"])
        self.assertEqual(revir_ui.session_state.get("caf_cfg_s1_n"), "")
        self.assertEqual(revir_ui.session_state.get("caf_cfg_s3_n"), "Hemşire Elif Kaya")

    def test_institutional_excel_templates_imports_and_invoice_unit_cost_revision(self):
        # 1. Menu CSV (Karaman Özel Eğitim Anaokulu 3-table format) -> import + 3-sheet xlsx round-trip
        menu_csv = (
            "KARAMAN ÖZEL EĞİTİM ANAOKULU;;;;;;;\n"
            "KASIM 2026 YEMEK LİSTESİ;;;;;;;\n"
            "Hafta;Tarih;Gün;1. Çeşit;2. Çeşit;3. Çeşit;Tahmini Enerji (kcal);Günlük Alerjen Notu\n"
            "1. Hafta;02.11.2026;Pazartesi;Kıymalı Ispanak;Makarna;Yoğurt;500;Gluten, Süt Ürünü\n"
            "1. Hafta;03.11.2026;Salı;Mercimek Çorbası;Pirinç Pilavı;Meyve;480;Gluten, Süt Ürünü\n"
            ";;;;;;;\n"
            "YEMEK KALORİ VE ALERJEN BİLGİ TABLOSU;;;;;;;\n"
            "Kategori;Yemek / Ürün Adı;Ort. Kalori (kcal);Kalori Notu;Alerjen Detayı;Gluten;Süt Ürünü;Yumurta\n"
            "Çorbalar;Mercimek Çorbası;150;~150 kcal;Gluten, Süt;Var;Var;Yok\n"
            "Ana Yemekler;Kıymalı Ispanak;180;~180 kcal;Yok;Yok;Yok;Yok\n"
            "Pilav & Makarna;Makarna;200;~200 kcal;Gluten;Var;Yok;Yok\n"
            "Pilav & Makarna;Pirinç Pilavı;260;~260 kcal;Süt Ürünü;Yok;Var;Yok\n"
            "Yan Ürün;Yoğurt;120;~120 kcal;Süt Ürünü;Yok;Var;Yok\n"
            "Yan Ürün;Meyve;70;~70 kcal;Yok;Yok;Yok;Yok\n"
        ).encode("utf-8")

        res_menu_imp = caf.import_school_menu_from_excel_or_csv(
            self.connect,
            menu_csv,
            "karaman_kasim_2026.csv",
            default_slot="Öğle Yemeği",
            sync_to_diabetes=True,
        )
        self.assertGreaterEqual(res_menu_imp["plans_saved"], 2)
        self.assertGreaterEqual(res_menu_imp["dishes_upserted"], 6)

        plans_nov = caf.get_menu_plans_for_month(self.connect, "2026-11")
        self.assertIn(("2026-11-02", "Öğle Yemeği"), plans_nov)
        self.assertEqual(plans_nov[("2026-11-02", "Öğle Yemeği")]["dish_names"], ["Kıymalı Ispanak", "Makarna", "Yoğurt"])
        self.assertAlmostEqual(plans_nov[("2026-11-02", "Öğle Yemeği")]["total_kcal"], 500.0, places=1)

        school_meta = {"school_name": "Karaman Özel Eğitim Anaokulu", "cafeteria_sig1_role": "Hazırlayan", "cafeteria_sig1_name": "Muhammet DOĞANÇAY"}
        wb_menu_bytes = caf.build_school_menu_workbook_excel(self.connect, "2026-11", school_meta=school_meta)
        wb_menu = openpyxl.load_workbook(io.BytesIO(wb_menu_bytes), data_only=False)
        self.assertEqual(wb_menu.sheetnames, ["Yemek_Listesi", "Kalori_ve_Alerjen_Tablosu", "Menu_Ozeti_ve_Analiz"])
        ws_ml = wb_menu["Yemek_Listesi"]
        self.assertIn("KARAMAN ÖZEL EĞİTİM ANAOKULU", str(ws_ml["A1"].value))
        hdr_ml = [ws_ml.cell(4, c).value for c in range(1, 9)]
        self.assertIn("1. Çeşit", hdr_ml)
        self.assertIn("2. Çeşit", hdr_ml)
        self.assertIn("3. Çeşit", hdr_ml)

        # 2. Invoice institutional Excel template + CSV/Excel import + unit_cost computation
        tmpl_bytes = caf.build_purchase_invoice_template_excel(self.connect, year_month="2026-11", school_meta=school_meta)
        wb_inv_t = openpyxl.load_workbook(io.BytesIO(tmpl_bytes), data_only=False)
        self.assertEqual(wb_inv_t.sheetnames, ["Fatura_Giris_Sablonu", "Urun_Birim_Maliyet_Referans", "Fatura_Ozeti_ve_Analiz"])
        ws_fg = wb_inv_t["Fatura_Giris_Sablonu"]
        hdr_fg = [ws_fg.cell(4, c).value for c in range(1, 14)]
        self.assertIn("Miktar Cinsi", hdr_fg)
        self.assertIn("Birim Maliyet (KDV Dahil ₺)", hdr_fg)
        self.assertTrue(str(ws_fg["K5"].value).startswith("=ROUND("))
        self.assertTrue(str(ws_fg["L5"].value).startswith("=IF("))

        inv_csv = (
            "Fatura Tarihi;Fatura No;Tedarikçi / Firma;Ödeme Durumu;Ürün Adı;Kategori;Miktar;Miktar Cinsi;Birim Fiyat (KDV Hariç ₺);KDV (%);Toplam Tutar (KDV Dahil ₺);SKT (Ops.)\n"
            "2026-11-12;FTR-2026-77;Özkan Toptan;Vadeli / Bekliyor;Mercimek;Bakliyat & Kuru Gıda;25;kg;50;0;;\n"
            "2026-11-12;FTR-2026-77;Özkan Toptan;Vadeli / Bekliyor;Salça;Bakliyat & Kuru Gıda;5;teneke;0;0;600;\n"
        ).encode("utf-8")
        parsed_invs = caf.parse_purchase_invoices_from_excel_or_csv(inv_csv, "faturalar.csv")
        self.assertEqual(len(parsed_invs), 1)
        self.assertEqual(parsed_invs[0]["payment_status"], "Vadeli / Bekliyor")
        self.assertEqual(len(parsed_invs[0]["items"]), 2)

        pid_xl = caf.upsert_purchase_invoice(
            self.connect,
            invoice_date=parsed_invs[0]["invoice_date"],
            invoice_no=parsed_invs[0]["invoice_no"],
            supplier_name=parsed_invs[0]["supplier_name"],
            category=parsed_invs[0]["category"],
            items=parsed_invs[0]["items"],
            payment_status=parsed_invs[0]["payment_status"],
        )
        loaded_p = caf.get_purchase_by_id(self.connect, pid_xl)
        self.assertIsNotNone(loaded_p)
        by_nm = {it["item_name"]: it for it in loaded_p["items"]}
        self.assertAlmostEqual(by_nm["Mercimek"]["unit_cost"], 50.0, places=2)
        # 5 teneke total 600 TL -> back-calculated unit_price 120.0 & unit_cost 120.0
        self.assertAlmostEqual(by_nm["Salça"]["unit_price"], 120.0, places=2)
        self.assertAlmostEqual(by_nm["Salça"]["unit_cost"], 120.0, places=2)
        self.assertAlmostEqual(loaded_p["total_amount"], 1850.0, places=2)

        # Also check exporting an existing invoice populates its sheets cleanly
        ex_p_xlsx = caf.build_purchase_invoice_template_excel(self.connect, purchase_id=pid_xl, school_meta=school_meta)
        wb_ex_p = openpyxl.load_workbook(io.BytesIO(ex_p_xlsx), data_only=False)
        self.assertEqual(wb_ex_p["Fatura_Giris_Sablonu"].cell(5, 5).value, "Mercimek")

        # Analytics exposes Ort. Birim Maliyet (₺), Son Birim Maliyet (₺) and Birim Maliyet (₺)
        an_res = caf.compute_purchase_analytics(self.connect, year_month="2026-11", product_name="Salça")
        self.assertTrue(any(p["Ürün Adı"] == "Salça" and abs(p["Ort. Birim Maliyet (₺)"] - 120.0) < 0.01 for p in an_res["product_summary"]))
        self.assertAlmostEqual(an_res["product_history"][0]["Birim Maliyet (₺)"], 120.0, places=2)

        # 3. Headless UI verification of revised Invoice tab, Dialog Unit Cost breakdown, total-price mode & payment filter
        revir_ui.reset_session()
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        purch_tab = lambda: caf._render_tab_purchases_and_expenses(ctx_api, "2026-11", school_meta)
        out_p = revir_ui.run_action_headless(purch_tab, {"action": "render"})
        self.assertNotIn("Çalışma Zamanı Hatası", out_p["main_html"])
        self.assertIn("Resmi Fatura Excel Şablonu İndir", out_p["main_html"])
        self.assertIn("FTR-2026-77", out_p["main_html"])
        self.assertIn("Birim:", out_p["main_html"])

        # Payment status filter ("Vadeli / Bekliyor" -> idx 2) keeps FTR-2026-77; ("Ödendi" -> idx 1) excludes it
        out_flt_paid = revir_ui.run_action_headless(
            purch_tab, {"action": "change", "key": "caf_purch_st_flt", "value": 1}
        )
        self.assertNotIn("FTR-2026-77", out_flt_paid["main_html"])

        revir_ui.run_action_headless(
            purch_tab, {"action": "change", "key": "caf_purch_st_flt", "value": 2}
        )
        out_dlg = revir_ui.run_action_headless(
            purch_tab,
            {"action": "df_select", "key": "caf_invoices_list_tbl", "value": {"rows": [0], "cells": [[0, "Fatura No"]]}},
        )
        dlg_html = str(out_dlg.get("dialog_html") or "")
        self.assertIn("Birim Maliyet (KDV Dahil ₺)", dlg_html)
        self.assertIn("Kalem Bazlı Birim Maliyet", dlg_html)
        self.assertIn("caf_dlg_i_btn_mark_paid", dlg_html)

        # Switch to Total Price entry mode (idx 1) in dialog & update Salça total -> 750 TL (5 teneke => 150 TL/teneke)
        revir_ui.run_action_headless(
            purch_tab,
            {"action": "change", "key": "caf_dlg_i_price_mode", "value": 1},
        )
        revir_ui.run_action_headless(
            purch_tab,
            {"action": "data_editor", "key": "caf_dlg_i_items_editor", "value": {"op": "cell", "row_kind": "base", "row_id": 1, "col": "Toplam Tutar (₺)", "val": 750.0}},
        )
        revir_ui.run_action_headless(purch_tab, {"action": "click", "key": "caf_dlg_i_btn_save_inv"})
        reloaded_p = caf.get_purchase_by_id(self.connect, pid_xl)
        reloaded_by_nm = {it["item_name"]: it for it in reloaded_p["items"]}
        self.assertAlmostEqual(reloaded_by_nm["Salça"]["unit_cost"], 150.0, places=2)
        self.assertAlmostEqual(reloaded_by_nm["Salça"]["line_total"], 750.0, places=2)

        # Re-open and click "Ödendi İşaretle" in dialog
        revir_ui.run_action_headless(
            purch_tab,
            {"action": "df_select", "key": "caf_invoices_list_tbl", "value": {"rows": [0], "cells": [[0, "Fatura No"]]}},
        )
        revir_ui.run_action_headless(purch_tab, {"action": "click", "key": "caf_dlg_i_btn_mark_paid"})
        self.assertEqual(caf.get_purchase_by_id(self.connect, pid_xl)["payment_status"], "Ödendi")

        # 4. Headless UI check for Menu Planner 3-sheet Excel download + importer controls
        out_m = revir_ui.run_action_headless(
            lambda: caf._render_tab_menu_planner(ctx_api, "2026-11", school_meta),
            {"action": "render"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", out_m["main_html"])
        self.assertIn("3 Sekmeli Menü Excel", out_m["main_html"])
        self.assertIn("Excel / CSV Şablonundan Menü", out_m["main_html"])

    def test_seven_week_invoice_excel_category_monthly_avg_and_tender_plan(self):
        # Full 7-date institutional grocery workbook (matching user's pasted structure totaling 200,763 TL)
        raw_csv = """Tüm Haftaların Toplam Gider Özeti,
,
Hafta (Tarih Aralığı),Toplam Harcama
11.09.2026,41092
25.09.2026,36701
08.10.2026,21425
15.10.2026,22700
26.10.2026,17815
30.10.2026,9562
16.11.2026,51468
,
genel toplam,200763
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
11.09.2026,Fiş-001,Zerdeçal,1,Kg,280,280
11.09.2026,Fiş-001,Sarımsak tozu,1,Kg,495,495
11.09.2026,Fiş-001,Susam,1,Kg,210,210
11.09.2026,Fiş-001,Sumak,1,Kg,585,585
11.09.2026,Fiş-001,Toz tarhana 500gr,2,Adet,100,200
11.09.2026,Fiş-001,15'li Vanilin,5,Adet,40,200
11.09.2026,Fiş-001,5'li Kabartma tozu,5,Adet,25,125
11.09.2026,Fiş-001,Ayçiçek yağı,8,Adet,560,4480
11.09.2026,Fiş-001,Zeytinyağı,8,Adet,1450,11600
11.09.2026,Fiş-001,Kakao 1kg,1,Kg,740,740
11.09.2026,Fiş-001,Osmancık Pirinç,5,Adet,720,3600
11.09.2026,Fiş-001,Pilavlık bulgur,5,Adet,285,1425
11.09.2026,Fiş-001,Penne Makarna 5kg,2,Adet,300,600
11.09.2026,Fiş-001,Konserve Bezelye,3,Adet,240,720
11.09.2026,Fiş-001,Konserve Garnitür,3,Adet,295,885
11.09.2026,Fiş-001,Közlenmiş biber,3,Adet,120,360
11.09.2026,Fiş-001,Calve Ketçap,5,Adet,135,675
11.09.2026,Fiş-001,Öncü domates Salçası,6,Adet,580,3480
11.09.2026,Fiş-001,Öncü biber Salçası,4,Adet,895,3580
11.09.2026,Fiş-001,Süt İçim 1lt,24,Adet,58,1392
11.09.2026,Fiş-001,Torku Tereyağ,3,Adet,830,2490
11.09.2026,Fiş-001,Yoğurt 2kg,10,Adet,160,1600
11.09.2026,Fiş-001,Toz şeker 5kg,2,Adet,245,490
11.09.2026,Fiş-001,Donuk Mısır 2.5kg,2,Adet,440,880
,,,,,HAFTALIK TOPLAM:,41092
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
25.09.2026,Fiş-001,Tavuk Pirzola / 10'luk Yoğurt,1,Adet,650,650
25.09.2026,Fiş-001,Kıyma,6.5,Kg,1200,7800
25.09.2026,Fiş-001,Ispanak,20,Kg,97,1940
25.09.2026,Fiş-001,Süzme Yoğurt,2,Adet,185,370
25.09.2026,Fiş-001,Tereyağ,3,Kg,750,2250
25.09.2026,Fiş-001,Kuşbaşı,3,Kg,1200,3600
25.09.2026,Fiş-001,Teremyağ,2,Adet,65,130
25.09.2026,Fiş-001,Burgu Makarna 5kg,2,Adet,300,600
25.09.2026,Fiş-001,Üzüm Sirkesi,1,Adet,250,250
25.09.2026,Fiş-001,Soya Sosu,2,Adet,183,366
25.09.2026,Fiş-001,Yumurta,2,Koli,150,300
25.09.2026,Fiş-001,Arpa Şehriye,2,Adet,150,300
25.09.2026,Fiş-001,Limon,1,Kg,80,80
25.09.2026,Fiş-001,Domates,3,Kg,45,135
25.09.2026,Fiş-001,Patates,40,Kg,30,1200
25.09.2026,Fiş-001,Soğan,5,Kg,30,150
25.09.2026,Fiş-001,Nane,1,Bağ,25,25
25.09.2026,Fiş-001,Maydanoz,2,Bağ,20,40
25.09.2026,Fiş-001,Marul,2,Adet,45,90
25.09.2026,Fiş-001,Kabak,2,Kg,45,90
25.09.2026,Fiş-001,Patlıcan,3,Kg,60,180
25.09.2026,Fiş-001,Biber,4,Kg,80,320
25.09.2026,Fiş-001,Karpuz,2,Adet,50,100
25.09.2026,Fiş-001,Elma,4,Kg,75,300
25.09.2026,Fiş-001,Muz,4,Kg,100,400
25.09.2026,Fiş-001,Havuç,26,Kg,25,650
25.09.2026,Fiş-001,Fındık,0.3,Kg,1300,390
25.09.2026,Fiş-002,Limon,1,Kg,85,85
25.09.2026,Fiş-002,Yumurta,2,Koli,140,280
25.09.2026,Fiş-002,Baget,8.5,Kg,170,1445
25.09.2026,Fiş-002,Cebel Süzme,2,Adet,185,370
25.09.2026,Fiş-002,Sütaş 10kg yoğurt,1,Adet,690,690
25.09.2026,Fiş-002,Toros Bonfile,8.5,Kg,450,3825
25.09.2026,Fiş-002,Sütaş 10kg yoğurt,1,Adet,690,690
25.09.2026,Fiş-002,Sütaş 10kg yoğurt,1,Adet,650,650
25.09.2026,Fiş-002,Cebel Süzme,2,Adet,185,370
25.09.2026,Fiş-002,Muz,3,Kg,100,300
25.09.2026,Fiş-002,Elma,4,Kg,70,280
25.09.2026,Fiş-002,Armut,5,Kg,90,450
25.09.2026,Fiş-002,Yeşil biber,3,Kg,70,210
25.09.2026,Fiş-002,Kapya,2,Kg,60,120
25.09.2026,Fiş-002,Domates,5,Kg,45,225
25.09.2026,Fiş-002,Salatalık,2,Kg,35,70
25.09.2026,Fiş-002,Marul,1,Adet,45,45
25.09.2026,Fiş-002,Mor Lahana,1,Adet,50,50
25.09.2026,Fiş-002,Havuç,2,Kg,40,80
25.09.2026,Fiş-002,Maydanoz,1,Bağ,20,20
25.09.2026,Fiş-002,Kabak,5,Kg,40,200
25.09.2026,Fiş-002,Kıyma,3,Kg,1100,3300
25.09.2026,Fiş-002,Bardak Su,2,Koli,120,240
25.09.2026,Fiş-002,Browni intense,64,Adet,0,0
,,,,,HAFTALIK TOPLAM:,36701
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
08.10.2026,Fiş-001,10kg Yoğurt,2,Adet,700,1400
08.10.2026,Fiş-001,Süzme Yoğurt,2,Adet,185,370
08.10.2026,Fiş-001,Makarna,3,Adet,35,105
08.10.2026,Fiş-001,Tavuk / Fener,5,Kg,265,1325
08.10.2026,Fiş-001,Yumurta,2,Koli,140,280
08.10.2026,Fiş-001,Kıyma,2,Kg,1200,2400
08.10.2026,Fiş-001,Soğan,5,Kg,35,175
08.10.2026,Fiş-001,Muz,2.2,Kg,125,275
08.10.2026,Fiş-001,Mandalina,3,Kg,35,105
08.10.2026,Fiş-001,Portakal,3,Kg,40,120
08.10.2026,Fiş-001,Kapya,3,Kg,50,150
08.10.2026,Fiş-001,Biber,4,Kg,85,340
08.10.2026,Fiş-001,Maydanoz,1,Bağ,20,20
08.10.2026,Fiş-001,Domates,5,Kg,45,225
08.10.2026,Fiş-001,Salatalık,2,Kg,40,80
08.10.2026,Fiş-001,Marul,1,Adet,45,45
08.10.2026,Fiş-001,Mor Lahana,1,Adet,50,50
08.10.2026,Fiş-001,Havuç,3,Kg,45,135
08.10.2026,Fiş-001,Dolma Biber,7,Kg,85,595
08.10.2026,Fiş-001,10kg Yoğurt,1,Adet,700,700
08.10.2026,Fiş-001,Baget,7.5,Kg,180,1350
08.10.2026,Fiş-001,Kuşbaşı,7,Kg,1200,8400
08.10.2026,Fiş-001,Kıyma,1.5,Kg,1200,1800
08.10.2026,Fiş-001,Dolma Biber,7,Kg,90,630
08.10.2026,Fiş-001,Acı Sos / Salça,1,Adet,350,350
,,,,,HAFTALIK TOPLAM:,21425
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
15.10.2026,Fiş-001,Kuşbaşı,4,Kg,1200,4800
15.10.2026,Fiş-001,Süzme Yoğurt,2,Adet,185,370
15.10.2026,Fiş-001,5kg Kır. Mercimek Duru,8,Adet,350,2800
15.10.2026,Fiş-001,5kg Pirinç Duru,8,Adet,700,5600
15.10.2026,Fiş-001,5kg Safir Tuz / Yağ,6,Adet,250,1500
15.10.2026,Fiş-001,Süt (24'lü),1,Koli,1480,1480
15.10.2026,Fiş-001,5kg Makarna,4,Adet,320,1280
15.10.2026,Fiş-001,Üzüm / Un,1,Adet,750,750
15.10.2026,Fiş-001,Kabak,2,Kg,50,100
15.10.2026,Fiş-001,Havuç,4,Kg,45,180
15.10.2026,Fiş-001,Mor Lahana,1,Adet,50,50
15.10.2026,Fiş-001,Maydanoz,1,Bağ,20,20
15.10.2026,Fiş-001,Marul,2,Adet,45,90
15.10.2026,Fiş-001,Domates,2,Kg,50,100
15.10.2026,Fiş-001,Salatalık,4,Kg,60,240
15.10.2026,Fiş-001,Patlıcan,6,Kg,50,300
15.10.2026,Fiş-001,Kapya,3,Kg,50,150
15.10.2026,Fiş-001,Biber,4,Kg,85,340
15.10.2026,Fiş-001,Muz,2.2,Kg,125,275
15.10.2026,Fiş-001,Mandalina,3,Kg,35,105
15.10.2026,Fiş-001,Elma,3.5,Kg,50,175
15.10.2026,Fiş-001,Soğan,10,Kg,35,350
15.10.2026,Fiş-001,Sarımsak,1,Kg,175,175
15.10.2026,Fiş-001,Patates (42kg),42,Kg,35,1470
,,,,,HAFTALIK TOPLAM:,22700
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
26.10.2026,Fiş-001,Yoğurt,2,Adet,750,1500
26.10.2026,Fiş-001,Tavuk Sote,6,Kg,260,1560
26.10.2026,Fiş-001,Et,2,Kg,1200,2400
26.10.2026,Fiş-001,Yeşil Biber,1,Kg,95,95
26.10.2026,Fiş-001,Havuç,1,Kg,40,40
26.10.2026,Fiş-001,Et,2,Kg,1200,2400
26.10.2026,Fiş-001,Kapya,2,Kg,50,100
26.10.2026,Fiş-001,Yeşil Biber,3,Kg,95,285
26.10.2026,Fiş-001,Havuç,2,Kg,40,80
26.10.2026,Fiş-001,Limon,2,Kg,45,90
26.10.2026,Fiş-001,Lahana,1,Adet,50,50
26.10.2026,Fiş-001,Marul,2,Adet,50,100
26.10.2026,Fiş-001,Salatalık,2,Kg,45,90
26.10.2026,Fiş-001,Maydanoz,1,Bağ,20,20
26.10.2026,Fiş-001,Muz,3,Kg,130,390
26.10.2026,Fiş-001,Mandalina,4,Kg,50,200
26.10.2026,Fiş-001,Yeşillik / Nane,3,Bağ,30,90
26.10.2026,Fiş-001,Süzme Yoğurt,2,Adet,185,370
26.10.2026,Fiş-001,Tereyağ / Peynir,3,Kg,850,2550
26.10.2026,Fiş-001,Kıyma / Et,2,Kg,1200,2400
26.10.2026,Fiş-001,Yoğurt,1,Adet,750,750
26.10.2026,Fiş-001,Tavuk,6.5,Kg,260,1690
26.10.2026,Fiş-001,Karnabahar,8.5,Kg,50,425
26.10.2026,Fiş-001,Yumurta,1,Koli,140,140
,,,,,HAFTALIK TOPLAM:,17815
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
30.10.2026,Fiş-001,Meyve Suyu (Nektar 200ml),81,Adet,17,1377
30.10.2026,Fiş-001,Yumurta,1,Koli,150,150
30.10.2026,Fiş-001,10kg Yoğurt,2,Adet,700,1400
30.10.2026,Fiş-001,Kırmızı Et,4,Kg,1200,4800
30.10.2026,Fiş-001,Kapya,3,Kg,55,165
30.10.2026,Fiş-001,Biber,4,Kg,95,380
30.10.2026,Fiş-001,Kabak,1.5,Kg,60,90
30.10.2026,Fiş-001,Patlıcan,3,Kg,55,165
30.10.2026,Fiş-001,Havuç,2,Kg,40,80
30.10.2026,Fiş-001,Mor Lahana,1,Adet,55,55
30.10.2026,Fiş-001,Marul,2,Adet,50,100
30.10.2026,Fiş-001,Salatalık,2,Kg,45,90
30.10.2026,Fiş-001,Maydanoz,2,Bağ,20,40
30.10.2026,Fiş-001,Muz,3,Kg,130,390
30.10.2026,Fiş-001,Elma,4,Kg,40,160
30.10.2026,Fiş-001,Mandalina,4,Kg,30,120
,,,,,HAFTALIK TOPLAM:,9562
Tarih,Fiş/Fatura No,Malın Cinsi,Miktar,Birim,Birim Fiyatı,Tutar
16.11.2026,Fiş-001,Kırmızı Et,2,Kg,1200,2400
16.11.2026,Fiş-001,10kg Yoğurt,2,Adet,750,1500
16.11.2026,Fiş-001,Sıvı Yağ (5 Lt),8,Adet,575,4600
16.11.2026,Fiş-001,Donuk Patates (Koli),3,Koli,1250,3750
16.11.2026,Fiş-001,Yufka,2,Paket,130,260
16.11.2026,Fiş-001,Kırmızı Et,3,Kg,1200,3600
16.11.2026,Fiş-001,Tavuk But / Pirzola,5,Kg,270,1350
16.11.2026,Fiş-001,Tereyağ (1kg),1,Kg,750,750
16.11.2026,Fiş-001,10kg Yoğurt,1,Adet,750,750
16.11.2026,Fiş-001,Süzme Yoğurt,2,Adet,185,370
16.11.2026,Fiş-001,Kırmızı Et,2,Kg,1200,2400
16.11.2026,Fiş-001,Kapya,3,Kg,125,375
16.11.2026,Fiş-001,Biber,3,Kg,95,285
16.11.2026,Fiş-001,Patlıcan,8,Kg,60,480
16.11.2026,Fiş-001,Kabak,2,Kg,60,120
16.11.2026,Fiş-001,Marul,1,Adet,30,30
16.11.2026,Fiş-001,Mor Lahana / Havuç,2,Adet,50,100
16.11.2026,Fiş-001,Salatalık,2,Kg,50,100
16.11.2026,Fiş-001,Maydanoz,2,Bağ,20,40
16.11.2026,Fiş-001,Soğan (15kg),15,Kg,25,375
16.11.2026,Fiş-001,Rende Kaşar,3,Kg,550,1650
16.11.2026,Fiş-001,Sucuk Torku,3,Kg,320,960
16.11.2026,Fiş-002,10kg Yoğurt,2,Adet,650,1300
16.11.2026,Fiş-002,Süzme Yoğurt,2,Adet,185,370
16.11.2026,Fiş-002,Yumurta,2,Koli,150,300
16.11.2026,Fiş-002,Donuk Patates,25,Kg,110,2750
16.11.2026,Fiş-002,Kıyma,2,Kg,1100,2200
16.11.2026,Fiş-002,Yeşil Biber,4,Kg,105,420
16.11.2026,Fiş-002,Kapya,3.06,Kg,50,153
16.11.2026,Fiş-002,Domates,2,Kg,45,90
16.11.2026,Fiş-002,Salatalık,2,Kg,40,80
16.11.2026,Fiş-002,Marul,1,Adet,45,45
16.11.2026,Fiş-002,Mor Lahana,1,Adet,50,50
16.11.2026,Fiş-002,Havuç,2,Kg,40,80
16.11.2026,Fiş-002,Krema,8,Adet,75,600
16.11.2026,Fiş-002,Dilimli Kaşar,6,Adet,165,990
16.11.2026,Fiş-002,Tereyağ / Peynir,3,Kg,420,1260
16.11.2026,Fiş-002,Köftelik / Sucuk,4,Kg,475,1900
16.11.2026,Fiş-002,Lor Peyniri,3,Kg,100,300
16.11.2026,Fiş-002,Küp Şeker / Konserve,1,Koli,300,300
16.11.2026,Fiş-002,Un (10kg),4,Adet,125,500
16.11.2026,Fiş-002,10kg Yoğurt,2,Adet,650,1300
16.11.2026,Fiş-002,Tavuk Göğüs / Külbastı,6,Kg,250,1500
16.11.2026,Fiş-002,Kuşbaşı,2,Kg,1200,2400
16.11.2026,Fiş-002,Sotelik Tavuk,7.5,Kg,250,1875
16.11.2026,Fiş-002,Süzme Yoğurt,2,Adet,185,370
16.11.2026,Fiş-002,Nar Ekşisi,1,Adet,190,190
16.11.2026,Fiş-002,Ekmek,1,Toplu,3900,3900
,,,,,HAFTALIK TOPLAM:,51468
""".encode("utf-8")

        # 1. Verify date-grouped parsing: exact 7 date-named invoices totaling 200,763.00 TL
        parsed_by_date = caf.parse_purchase_invoices_from_excel_or_csv(raw_csv, "giderler.csv", group_mode="date")
        self.assertEqual(len(parsed_by_date), 7)
        self.assertEqual([inv["invoice_name"] for inv in parsed_by_date], [
            "11.09.2026", "25.09.2026", "08.10.2026", "15.10.2026", "26.10.2026", "30.10.2026", "16.11.2026"
        ])
        for inv in parsed_by_date:
            caf.upsert_purchase_invoice(
                self.connect,
                invoice_date=inv["invoice_date"],
                invoice_name=inv["invoice_name"],
                invoice_no=inv["invoice_no"],
                supplier_name=inv["supplier_name"],
                category=inv["category"],
                items=inv["items"],
                payment_status=inv["payment_status"],
            )

        all_invs, all_items = caf.get_purchases_and_items(self.connect, year_month="Tümü")
        self.assertEqual(len(all_invs), 7)
        by_name_tot = {inv["invoice_name"]: inv["total_amount"] for inv in all_invs}
        self.assertAlmostEqual(by_name_tot["11.09.2026"], 41092.0, places=2)
        self.assertAlmostEqual(by_name_tot["25.09.2026"], 36701.0, places=2)
        self.assertAlmostEqual(by_name_tot["08.10.2026"], 21425.0, places=2)
        self.assertAlmostEqual(by_name_tot["15.10.2026"], 22700.0, places=2)
        self.assertAlmostEqual(by_name_tot["26.10.2026"], 17815.0, places=2)
        self.assertAlmostEqual(by_name_tot["30.10.2026"], 9562.0, places=2)
        self.assertAlmostEqual(by_name_tot["16.11.2026"], 51468.0, places=2)
        self.assertAlmostEqual(sum(inv["total_amount"] for inv in all_invs), 200763.0, places=2)

        # Verify unit 'toplu' normalized on Ekmek, pack conversion on Sıvı Yağ (5 Lt) -> 40 litre & Yoğurt
        ekmek_it = next(it for it in all_items if it["item_name"] == "Ekmek")
        self.assertEqual(ekmek_it["unit_type"], "toplu")
        self.assertEqual(ekmek_it["category"], "Unlu Mamul & Ekmek")
        sivi_yag_it = next(it for it in all_items if it["item_name"] == "Sıvı Yağ (5 Lt)")
        self.assertAlmostEqual(sivi_yag_it["net_qty"], 40.0, places=2)
        self.assertEqual(sivi_yag_it["net_unit"], "litre")
        self.assertAlmostEqual(sivi_yag_it["net_unit_cost"], 115.0, places=2)

        # Verify Sebze & Meyve item-level filtering returns produce items across mixed invoices
        _, veg_items = caf.get_purchases_and_items(self.connect, year_month="2026-10", category="Sebze & Meyve")
        self.assertGreater(len(veg_items), 15)
        self.assertTrue(all(it["category"] == "Sebze & Meyve" for it in veg_items))

        # 2. Verify analytics: 3 active months, canonical monthly average + tender planning (.xlsx & .docx)
        an = caf.compute_purchase_analytics(self.connect, year_month="Tümü")
        self.assertEqual(an["summary"]["active_months"], 3)
        self.assertAlmostEqual(an["summary"]["total_spend_tl"], 200763.0, places=2)
        self.assertEqual(len(an["month_category_matrix"]), 3)
        canon_by_k = {(c["Ana Gıda / İhale Kalemi"], c["Net Birim"]): c for c in an["canonical_summary"]}
        self.assertIn(("Kırmızı Et (Kıyma / Kuşbaşı)", "kg"), canon_by_k)
        red_meat = canon_by_k[("Kırmızı Et (Kıyma / Kuşbaşı)", "kg")]
        # Red meat total = 12.5 + 10.5 + 4.0 + 6.0 + 4.0 + 11.0 = 48.0 kg over 3 months -> 16.0 kg/ay
        self.assertAlmostEqual(red_meat["Toplam Net Miktar"], 48.0, places=2)
        self.assertAlmostEqual(red_meat["Aylık Ort. Miktar"], 16.0, places=2)

        # Verify receipt-split mode yields 9 separate receipt invoices with exact same 200,763 TL sum
        parsed_by_rec = caf.parse_purchase_invoices_from_excel_or_csv(raw_csv, "giderler.csv", group_mode="receipt")
        self.assertEqual(len(parsed_by_rec), 9)
        self.assertTrue(any("· Fiş-002" in inv["invoice_name"] for inv in parsed_by_rec))

        plan_12m = caf.build_procurement_and_tender_plan(an, period_months=12.0, qty_margin_pct=10.0)
        self.assertGreater(plan_12m["summary"]["total_budget_tl"], 700000.0)
        meat_plan = next(r for r in plan_12m["rows"] if r["Gıda Kalemi / Ürün"] == "Kırmızı Et (Kıyma / Kuşbaşı)")
        self.assertGreater(meat_plan["12 Aylık Yıllık (Marjlı)"], 200.0)

        xl_plan = caf.build_tender_and_order_plan_excel(plan_12m, "12 Aylık Tam Yıl İhale")
        wb_tp = openpyxl.load_workbook(io.BytesIO(xl_plan))
        self.assertEqual(wb_tp.sheetnames, ["Siparis_ve_Ihale_Cetveli", "Kategori_Butce_Ozeti"])
        docx_plan = caf.build_tender_summary_docx(plan_12m, "12 Aylık Tam Yıl İhale")
        self.assertIn("Kırmızı Et", _extract_docx_all_text(docx_plan))

        # 3. Verify manual entry UI: date becomes Fatura Adı, entering Qty + Bizden Alınan Tutar calculates Birim Fiyat
        revir_ui.reset_session()
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": persantil._get_active_gemini_config,
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        purch_tab = lambda: caf._render_tab_purchases_and_expenses(ctx_api, "2026-11", {})
        out_tab = revir_ui.run_action_headless(purch_tab, {"action": "render"})
        self.assertNotIn("Çalışma Zamanı Hatası", out_tab["main_html"])
        self.assertIn("Sipariş Planlayıcı &amp; Yıllık İhale", out_tab["main_html"])
        self.assertIn("Aylık Ortalama Kullanım", out_tab["main_html"])

        # Enter 3 kg Et with total 3600 TL without touching Birim Fiyat -> auto calculates 1200 TL/kg
        for mut in (
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Ürün Adı", "val": "Kırmızı Et"},
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Miktar", "val": 3.0},
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Miktar Çeşidi", "val": "kg"},
            {"op": "cell", "row_kind": "base", "row_id": 0, "col": "Toplam Tutar (₺)", "val": 3600.0},
        ):
            revir_ui.run_action_headless(purch_tab, {"action": "data_editor", "key": "caf_inv_items_editor", "value": mut})
        revir_ui.run_action_headless(purch_tab, {"action": "click", "key": "caf_inv_btn_save_inv"})
        all_after, _ = caf.get_purchases_and_items(self.connect, year_month="Tümü")
        newest = all_after[0]
        self.assertAlmostEqual(newest["total_amount"], 3600.0, places=2)
        self.assertAlmostEqual(newest["items"][0]["unit_price"], 1200.0, places=2)
        self.assertEqual(newest["items"][0]["category"], "Et / Tavuk / Balık")

    def test_text_invoice_parser_multi_invoice_headers_and_single_window_stats(self):
        # 1. Exact user-pasted Sipariş Fişi text with Turkish thousands dots (4.480, 11.600, 1.425, 2.490)
        pasted_sep = """Belge Türü: Sipariş Fişi
Tarih: 11.09.2026

Miktar | Malın Cinsi | Fiyat | Tutar
1 | Zerdeçal | 280 | 280
1 | Sarımsak Tozu | 195 | 195
1 | Susam | 210 | 210
1 | Sucuk | 585 | 585
2 | Toz Tarhana | 100 | 200
5 | 15'li Vanilin / Kabartma Tozu | 40 | 200
5 | 5'li Kabartma Tozu | 25 | 125
8 | Ayçiçek Yağı (Teneke) | 560 | 4.480
8 | Zeytinyağı (5 Lt) | 1.450 | 11.600
1 | Kakao | 740 | 740
5 | Susamlı Bisküvi / Pirinç | 720 | 3.600
5 | Bulgur | 285 | 1.425
2 | Şehriye / Makarna | 300 | 600
3 | Konserve Bezelye | 240 | 720
3 | Konserve Garnitür | 295 | 885
3 | Salça (Biber / Domates) | 120 | 360
5 | İnce Domates / Biber Salçası | 135 | 675
6 | Süt (Koli) | 580 | 3.480
4 | Şeker (5 kg) | 895 | 3.580
24 | Süt / İçecek (Adet) | 58 | 1.392
3 | Donuk Mısır (2.5 kg) | 830 | 2.490
10 | Okunamayan Kalem | 160 | 1.600
2 | Okunamayan Kalem | 245 | 490
2 | Okunamayan Kalem | 440 | 880"""

        pasted_oct = """Belge Türü: Sipariş Fişi
Tarih: 08.10.2026

Miktar | Malın Cinsi | Fiyat | Tutar
2 | 10kg Yoğurt | 700 | 1.400
7.5 | Baget | 180 | 1.350"""

        parsed_single = caf.parse_purchase_invoices_from_text(pasted_sep)
        self.assertEqual(len(parsed_single), 1)
        inv1 = parsed_single[0]
        self.assertEqual(inv1["invoice_date"], "2026-09-11")
        self.assertEqual(inv1["invoice_name"], "11.09.2026 · Sipariş Fişi")
        self.assertEqual(len(inv1["items"]), 24)
        by_nm1 = {it["item_name"]: it for it in inv1["items"] if it["item_name"] != "Okunamayan Kalem"}
        self.assertEqual(by_nm1["Ayçiçek Yağı (Teneke)"]["unit_type"], "teneke")
        self.assertAlmostEqual(by_nm1["Ayçiçek Yağı (Teneke)"]["line_total"], 4480.0, places=2)
        self.assertAlmostEqual(by_nm1["Zeytinyağı (5 Lt)"]["unit_price"], 1450.0, places=2)
        self.assertAlmostEqual(by_nm1["Zeytinyağı (5 Lt)"]["line_total"], 11600.0, places=2)
        self.assertAlmostEqual(by_nm1["Zeytinyağı (5 Lt)"]["net_qty"], 40.0, places=2)
        self.assertEqual(by_nm1["Zeytinyağı (5 Lt)"]["net_unit"], "litre")
        self.assertEqual(by_nm1["Süt (Koli)"]["unit_type"], "koli")
        self.assertAlmostEqual(by_nm1["Donuk Mısır (2.5 kg)"]["net_qty"], 7.5, places=2)
        expected_sep_total = 280+195+210+585+200+200+125+4480+11600+740+3600+1425+600+720+885+360+675+3480+3580+1392+2490+1600+490+880
        self.assertAlmostEqual(sum(it["line_total"] for it in inv1["items"]), float(expected_sep_total), places=2)

        # 2. Multi-invoice text splits into 2 separate headers
        parsed_multi = caf.parse_purchase_invoices_from_text(pasted_sep + "\n\n" + pasted_oct)
        self.assertEqual(len(parsed_multi), 2)
        self.assertEqual([x["invoice_date"] for x in parsed_multi], ["2026-09-11", "2026-10-08"])
        self.assertAlmostEqual(parsed_multi[1]["items"][1]["quantity"], 7.5, places=2)

        # 3. Scanned image PDF without text layer triggers Gemini fallback cleanly & keeps errors informative when AI off
        import pymupdf
        doc_img = pymupdf.open()
        doc_img.new_page()
        pdf_bytes = doc_img.tobytes()
        doc_img.close()
        invs_no_ai, errs_no_ai = caf._parse_uploaded_invoice_sources(
            raw_text="",
            uploaded_files=[("fis_25_09.pdf", pdf_bytes)],
            is_gemini_on=False,
        )
        self.assertEqual(invs_no_ai, [])
        self.assertEqual(len(errs_no_ai), 1)
        self.assertIn("Gemini API", errs_no_ai[0])

        def _fake_ai_pdf(**_kw):
            return True, {
                "invoice_date": "2026-09-25",
                "document_type": "Sipariş Fişi",
                "invoice_name": "25.09.2026 · Sipariş Fişi",
                "items": [{"item_name": "Tavuk Pirzola", "quantity": 7.5, "unit_type": "kg", "unit_price": 260, "line_total": 1950}],
            }, "gemini-mock"

        invs_with_ai, errs_with_ai = caf._parse_uploaded_invoice_sources(
            raw_text="",
            uploaded_files=[("fis_25_09.pdf", pdf_bytes)],
            is_gemini_on=True,
            api_key="test",
            model="gemini-mock",
            gemini_parse_fn=_fake_ai_pdf,
        )
        self.assertEqual(errs_with_ai, [])
        self.assertEqual(len(invs_with_ai), 1)
        self.assertAlmostEqual(invs_with_ai[0]["total_preview"], 1950.0, places=2)

        # 4. Headless UI end-to-end: paste 2 invoices -> parse -> save -> separate expanders + month/view box stats
        revir_ui.reset_session()
        ctx_api = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": lambda: (False, "", ""),
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        purch_tab = lambda: caf._render_tab_purchases_and_expenses(ctx_api, "2026-11", {})
        revir_ui.run_action_headless(purch_tab, {"action": "change", "key": "caf_ai_inv_txt", "value": pasted_sep + "\n\n" + pasted_oct})
        out_parse = revir_ui.run_action_headless(purch_tab, {"action": "click", "key": "caf_btn_ai_parse_inv"})
        self.assertIn("2 fatura", out_parse["main_html"])
        out_save_q = revir_ui.run_action_headless(purch_tab, {"action": "click", "key": "caf_btn_save_parsed_queue"})
        html_after = out_save_q["main_html"]
        self.assertIn("11.09.2026 · Sipariş Fişi", html_after)
        self.assertIn("08.10.2026 · Sipariş Fişi", html_after)
        self.assertIn("Zerdeçal", html_after)
        self.assertIn("Baget", html_after)

        # Test switching single-window stats scope from All -> 2026-09 month only
        revir_ui.session_state["caf_stat_scope"] = f"2026-09 — {caf._fmt_ym('2026-09')}"
        out_sep_st = revir_ui.run_action_headless(purch_tab, {"action": "render"})
        stats_slice = out_sep_st["main_html"].split("Fatura &amp; Tüketim İstatistikleri", 1)[1].split("Planlı Menüye Göre", 1)[0]
        self.assertIn("40,792.00 ₺", stats_slice)
        self.assertNotIn("2,750.00 ₺", stats_slice)
        self.assertNotIn("43,542.00 ₺", stats_slice)

        # Verify Gemini invoice normalizer handles Turkish thousands strings & unit_price*qty checks
        norm_g = gemini_assistant._normalize_gemini_invoice_dict({
            "invoice_date": "2026-10-08",
            "items": [
                {"item_name": "Kuşbaşı", "quantity": "7", "unit_type": "kg", "unit_price": "1.200", "line_total": "8.400"},
                {"item_name": "Baget", "quantity": "7.5", "unit_type": "kg", "unit_price": "180", "line_total": "1.350"},
            ],
        })
        self.assertAlmostEqual(norm_g["items"][0]["line_total"], 8400.0, places=2)
        self.assertAlmostEqual(norm_g["items"][1]["quantity"], 7.5, places=2)
        self.assertAlmostEqual(norm_g["items"][1]["line_total"], 1350.0, places=2)

        # Verify Sipariş Fişi 5-col table where MİKTAR is blank and quantity ('2', '4kg') is inside BİRİM column
        siparis_5col = """Belge Türü: Sipariş Fişi
Tarih: 30 / 10 / 2026

MİKTAR | BİRİM | MALIN CİNSİ | FİYAT | TUTAR
| 2 | 10kg Yoğurt | 700 | 1.400
| 4kg | Et Kırmızı | 1.200 | 4.800"""
        parsed_5col = caf.parse_purchase_invoices_from_text(siparis_5col)
        self.assertEqual(len(parsed_5col), 1)
        self.assertEqual(parsed_5col[0]["invoice_date"], "2026-10-30")
        self.assertAlmostEqual(parsed_5col[0]["items"][0]["quantity"], 2.0, places=2)
        self.assertAlmostEqual(parsed_5col[0]["items"][0]["line_total"], 1400.0, places=2)
        self.assertAlmostEqual(parsed_5col[0]["items"][1]["quantity"], 4.0, places=2)
        self.assertEqual(parsed_5col[0]["items"][1]["unit_type"], "kg")
        self.assertAlmostEqual(parsed_5col[0]["items"][1]["line_total"], 4800.0, places=2)

    def test_menu_json_recovery_docx_table_excel_and_retry(self):
        import io
        import docx

        # 1. JSON recovery: truncated MAX_TOKENS JSON, trailing commas, top-level array with prose, non-json
        trunc = (
            '{"plans": ['
            '{"plan_date": "2026-11-02", "meal_slot": "Öğle Yemeği", "dishes": [{"name": "Mercimek Çorbası"}]}, '
            '{"plan_date": "2026-11-03", "dishes": [{"name": "Piri'
        )
        rescued = gemini_assistant._parse_json_value(trunc)
        self.assertIsInstance(rescued, dict)
        self.assertEqual(len(rescued["plans"]), 1)
        self.assertEqual(rescued["plans"][0]["plan_date"], "2026-11-02")

        trailing = '{"plans": [{"plan_date": "2026-11-04", "dishes": ["Pilav",],}],}'
        self.assertEqual(gemini_assistant._parse_json_response(trailing)["plans"][0]["plan_date"], "2026-11-04")

        arr_prose = 'İşte menü:\n[{"date": "02.11.2026", "meal_type": "Öğle", "yemekler": ["Tarhana Çorbası", "Kuru Fasulye"]}]\nAfiyet olsun.'
        arr_val = gemini_assistant._parse_json_value(arr_prose)
        self.assertIsInstance(arr_val, list)
        self.assertEqual(gemini_assistant._parse_json_response("not json"), {})

        # 2. Word .docx with menu stored ONLY inside a table (empty paragraphs) including horizontally merged cells
        doc = docx.Document()
        tbl = doc.add_table(rows=2, cols=3)
        tbl.cell(0, 0).merge(tbl.cell(0, 2))
        tbl.cell(0, 0).text = "KASIM 2026 YEMEK LİSTESİ"
        tbl.cell(1, 0).text = "02.11.2026"
        tbl.cell(1, 1).text = "Öğle Yemeği"
        tbl.cell(1, 2).text = "Mercimek Çorbası, Etli Nohut, Pilav"
        bio_docx = io.BytesIO()
        doc.save(bio_docx)
        _, docx_txt = gemini_assistant._extract_attachment_context("kasim_menu.docx", "", bio_docx.getvalue())
        self.assertIn("Mercimek Çorbası", docx_txt)
        self.assertIn("Etli Nohut", docx_txt)
        self.assertEqual(docx_txt.count("KASIM 2026 YEMEK LİSTESİ"), 1)

        # 3. Excel .xlsx with title row and meal data across columns
        df_xl = pd.DataFrame([
            ["OKUL YEMEKHANESİ AYLIK MENÜ", None, None],
            ["03.11.2026", "Öğle", "Yayla Çorbası, İzmir Köfte"],
        ])
        bio_xl = io.BytesIO()
        with pd.ExcelWriter(bio_xl, engine="openpyxl") as writer:
            df_xl.to_excel(writer, index=False, header=False, sheet_name="Kasim")
        _, xl_txt = gemini_assistant._extract_attachment_context("kasim_menu.xlsx", "", bio_xl.getvalue())
        self.assertIn("Yayla Çorbası", xl_txt)
        self.assertIn("İzmir Köfte", xl_txt)

        # 4. End-to-end menu parser: truncated JSON, alternate keys/array, retry after garbage, and empty input
        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, trunc, "gemini-3-flash-preview")):
            ok_t, res_t, _ = gemini_assistant.parse_cafeteria_menu_or_payments_text_with_gemini(
                api_key="k",
                model="gemini-3-flash-preview",
                task_mode="menu",
                workdays_list=["2026-11-02", "2026-11-03"],
                meal_slots=["Sabah Kahvaltısı", "Öğle Yemeği"],
                file_bytes=bio_docx.getvalue(),
                file_name="kasim_menu.docx",
            )
        self.assertTrue(ok_t)
        self.assertEqual(len(res_t["plans"]), 1)
        self.assertEqual(res_t["plans"][0]["plan_date"], "2026-11-02")

        with mock.patch.object(gemini_assistant, "call_gemini", return_value=(True, arr_prose, "gemini-3-flash-preview")):
            ok_a, res_a, _ = gemini_assistant.parse_cafeteria_menu_or_payments_text_with_gemini(
                api_key="k",
                model="gemini-3-flash-preview",
                task_mode="menu",
                workdays_list=["2026-11-02"],
                meal_slots=["Sabah Kahvaltısı", "Öğle Yemeği"],
                raw_text="02.11.2026 Öğle: Tarhana Çorbası, Kuru Fasulye",
            )
        self.assertTrue(ok_a)
        self.assertEqual(res_a["plans"][0]["plan_date"], "2026-11-02")
        self.assertEqual(res_a["plans"][0]["meal_slot"], "Öğle Yemeği")
        self.assertEqual([d["name"] for d in res_a["plans"][0]["dishes"]], ["Tarhana Çorbası", "Kuru Fasulye"])

        retry_responses = [
            (True, "Üzgünüm, metin anlaşılamadı.", "gemini-3-flash-preview"),
            (True, '{"plans":[{"plan_date":"2026-11-05","dishes":[{"name":"Ezogelin Çorbası"}]}]}', "gemini-3-flash-preview"),
        ]
        with mock.patch.object(gemini_assistant, "call_gemini", side_effect=retry_responses) as m_call:
            ok_r, res_r, _ = gemini_assistant.parse_cafeteria_menu_or_payments_text_with_gemini(
                api_key="k",
                model="gemini-3-flash-preview",
                task_mode="menu",
                workdays_list=["2026-11-05"],
                raw_text="05.11.2026 Ezogelin Çorbası",
            )
        self.assertTrue(ok_r)
        self.assertEqual(m_call.call_count, 2)
        self.assertEqual(res_r["plans"][0]["dishes"][0]["name"], "Ezogelin Çorbası")

        with mock.patch.object(gemini_assistant, "call_gemini") as m_empty:
            ok_e, res_e, _ = gemini_assistant.parse_cafeteria_menu_or_payments_text_with_gemini(
                api_key="k", model="gemini-3-flash-preview", task_mode="menu", raw_text="   ",
            )
        self.assertFalse(ok_e)
        self.assertEqual(m_empty.call_count, 0)
        self.assertIn("belge", res_e["error"].lower())

    def test_fee_group_title_policy_filtering_and_cash_first_tab_flow(self):
        # 1. Seed diverse school roles + guardian in students table
        with self.connect() as c:
            c.execute("UPDATE students SET guardian_name='Veli Ayşe Kaya' WHERE id=?", (self.stu_id,))
            c.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sivil / Diğer')")
            civ_cid = c.execute("SELECT id FROM classes WHERE name='Sivil / Diğer'").fetchone()[0]
            md_id = c.execute(
                "INSERT INTO students(class_id, person_type, full_name, sex, notes) VALUES (?, 'Sivil / diğer kişi', 'Hakan MüdürYrd', 'E', 'Müdür Yardımcısı')",
                (civ_cid,),
            ).lastrowid
            hz_id = c.execute(
                "INSERT INTO students(class_id, person_type, full_name, sex, notes) VALUES (?, 'Sivil / diğer kişi', 'Kemal Hizmetli', 'E', 'Hizmetli')",
                (civ_cid,),
            ).lastrowid
            stj_id = c.execute(
                "INSERT INTO students(class_id, person_type, full_name, sex, notes) VALUES (?, 'Sivil / diğer kişi', 'Elif Stajyer', 'K', 'Stajyer')",
                (civ_cid,),
            ).lastrowid
            veli_id = c.execute(
                "INSERT INTO students(class_id, person_type, full_name, sex, notes) VALUES (?, 'Sivil / diğer kişi', 'Murat Veli', 'E', 'Veli')",
                (civ_cid,),
            ).lastrowid
            c.commit()

        # Import people -> verify real titles & fee groups resolved cleanly
        caf.import_school_people_as_diners(self.connect, [self.stu_id, self.teacher_id, md_id, hz_id, stj_id, veli_id])
        diners_by_nm = {d["full_name"]: d for d in caf.get_diners(self.connect)}
        self.assertEqual(diners_by_nm["Zeynep Öğretmen"]["effective_role_title"], "Öğretmen")
        self.assertEqual(diners_by_nm["Zeynep Öğretmen"]["fee_group"], "👩‍🏫 Öğretmenler")
        self.assertEqual(caf.resolve_diner_role_group_and_title(person_type="Sivil / diğer kişi", notes="Matematik Öğretmeni"), ("👩‍🏫 Öğretmenler", "Matematik Öğretmeni"))
        self.assertEqual(diners_by_nm["Hakan MüdürYrd"]["fee_group"], "🏫 Okul & Yurt Yönetimi")
        self.assertEqual(diners_by_nm["Kemal Hizmetli"]["fee_group"], "👷 Sözleşmeli, İŞKUR, Hizmetli & Stajyer")
        self.assertEqual(diners_by_nm["Murat Veli"]["fee_group"], "👨‍👩‍👧 Veliler & Misafirler")

        # 2. Restrict fee collection in settings: Öğrenciler + Öğretmenler + Hizmetli (exclude Stajyer/Yönetim/Veliler)
        restricted_cfg = {
            "cafeteria_fee_scope_configured": "1",
            "cafeteria_fee_groups": json.dumps(["🎓 Öğrenciler", "👩‍🏫 Öğretmenler"], ensure_ascii=False),
            "cafeteria_fee_included_titles": json.dumps(["Hizmetli"], ensure_ascii=False),
            "cafeteria_fee_excluded_titles": json.dumps(["Stajyer"], ensure_ascii=False),
            "cafeteria_include_student_guardians": "0",
        }
        cands = caf.get_fee_collection_candidates(self.connect, restricted_cfg)
        cand_names = {c["full_name"] for c in cands}
        self.assertEqual(cand_names, {"Ali Kaya", "Zeynep Öğretmen", "Kemal Hizmetli"})
        self.assertNotIn("Hakan MüdürYrd", cand_names)
        self.assertNotIn("Elif Stajyer", cand_names)
        self.assertNotIn("Murat Veli", cand_names)
        self.assertNotIn("Veli Ayşe Kaya", cand_names)

        # 3. Enable Veliler + student guardians -> standalone Veli AND student guardian appear
        parent_cfg = {
            "cafeteria_fee_scope_configured": "1",
            "cafeteria_fee_groups": json.dumps(["👨‍👩‍👧 Veliler & Misafirler"], ensure_ascii=False),
            "cafeteria_include_student_guardians": "1",
        }
        parent_names = {c["full_name"] for c in caf.get_fee_collection_candidates(self.connect, parent_cfg)}
        self.assertEqual(parent_names, {"Murat Veli", "Veli Ayşe Kaya"})

        # 4. Headless UI: first active tab of Yemek Yiyenler & Kasa MUST be Ücret Tahsilatı & Kasa
        revir_ui.reset_session()
        ctx_api = {
            "connect_db": persantil.connect_db,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": lambda: (False, "", ""),
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        persantil.save_app_settings(restricted_cfg)
        cur_cfg = persantil.get_app_settings()
        out_cash = revir_ui.run_action_headless(
            lambda: caf._render_tab_diners_and_cash(ctx_api, "2026-11", cur_cfg),
            {"action": "render"},
        )
        cash_html = str(out_cash.get("main_html") or "")
        # First active pane (active class on tab 0) contains Ücret Tahsilatı Al
        first_pane = cash_html.split('class="revir-tab-pane active"', 1)[1].split('class="revir-tab-pane"', 1)[0]
        self.assertIn("Ücret Tahsilatı Al", first_pane)
        # Payment dropdown contains ONLY eligible candidates, NOT excluded titles
        pay_block = first_pane.split('data-key="caf_pay_diner"', 1)[1].split('caf_form_in_pay', 1)[0]
        self.assertIn("Ali Kaya", pay_block)
        self.assertIn("Zeynep Öğretmen", pay_block)
        self.assertIn("Kemal Hizmetli", pay_block)
        self.assertNotIn("Hakan MüdürYrd", pay_block)
        self.assertNotIn("Elif Stajyer", pay_block)
        self.assertNotIn("Murat Veli", pay_block)

        # 5. Settings panel preset button updates DB fee groups cleanly
        revir_ui.reset_session()
        revir_ui.run_action_headless(
            lambda: caf.render_cafeteria_settings_panel(ctx_api, persantil.get_app_settings(), key_prefix="caf"),
            {"action": "click", "key": "caf_fee_preset_stu"},
        )
        saved_pol = caf.parse_fee_collection_policy(persantil.get_app_settings())
        self.assertTrue(saved_pol["configured"])
        self.assertEqual(saved_pol["eligible_groups"], ["🎓 Öğrenciler"])

    def test_custom_renamed_meals_a4_page_modes_ramazan_date_sub_and_smart_planner(self):
        # 1. Custom meal creation + existing meal renaming with DB plan migration
        d_soup = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Mercimek Çorbası",
            category="Çorba",
            ingredients=[{"ingredient_name": "Kırmızı Mercimek", "amount": 60.0, "unit": "g", "kcal_portion": 180.0, "carbs_portion_g": 28.0, "allergen_tag": ""}],
        )
        d_borek = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Peynirli Börek",
            category="Pilav / Makarna / Börek",
            ingredients=[
                {"ingredient_name": "Yufka", "amount": 80.0, "unit": "g", "kcal_portion": 260.0, "carbs_portion_g": 42.0, "allergen_tag": "Gluten"},
                {"ingredient_name": "Beyaz Peynir", "amount": 40.0, "unit": "g", "kcal_portion": 110.0, "carbs_portion_g": 1.0, "allergen_tag": "Süt"},
            ],
        )
        d_pilav = caf.upsert_dish_with_ingredients(
            self.connect,
            name="Pirinç Pilavı",
            category="Pilav / Makarna / Börek",
            ingredients=[{"ingredient_name": "Pirinç", "amount": 75.0, "unit": "g", "kcal_portion": 250.0, "carbs_portion_g": 52.0, "allergen_tag": ""}],
        )

        # Plan under default "Sabah Kahvaltısı", then rename to "Kuşluk Kahvaltısı" -> plan MUST migrate
        caf.upsert_menu_plan_slot(self.connect, plan_date="2026-10-05", meal_slot="Sabah Kahvaltısı", dish_ids=[d_borek])
        res_ren = caf.rename_meal_slot(
            self.connect,
            persantil.save_app_settings,
            persantil.get_app_settings(),
            "Sabah Kahvaltısı",
            "Kuşluk Kahvaltısı",
            new_time="09:30 - 10:00",
        )
        self.assertTrue(res_ren["ok"])
        self.assertEqual(res_ren["migrated_plans"], 1)

        # Add a brand-new user-defined custom meal slot
        res_add_c = caf.add_custom_meal_slot(
            self.connect,
            persantil.save_app_settings,
            persantil.get_app_settings(),
            "Gece Etüt Ara Öğünü",
            meal_time="20:30 - 21:00",
            diabetes_slot="Sahur / Gece",
            activate=True,
        )
        self.assertTrue(res_add_c["ok"])

        cfg_custom = persantil.get_app_settings()
        sched_c = caf.parse_service_schedule_from_cfg(cfg_custom)
        self.assertIn("Kuşluk Kahvaltısı", sched_c["available_slots"])
        self.assertIn("Gece Etüt Ara Öğünü", sched_c["available_slots"])
        self.assertEqual(sched_c["meal_times"]["Kuşluk Kahvaltısı"], "09:30 - 10:00")
        self.assertEqual(sched_c["meal_times"]["Gece Etüt Ara Öğünü"], "20:30 - 21:00")

        # Plan 3 distinct meals on 2026-10-05 (including allergy alt_menu_note) & sync to diabetes
        caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2026-10-05",
            meal_slot="Öğle Yemeği",
            dish_ids=[d_soup, d_pilav],
            alt_menu_note="Çölyaklı öğrenciye haşlama patates",
            sync_to_diabetes=True,
        )
        caf.upsert_menu_plan_slot(
            self.connect,
            plan_date="2026-10-05",
            meal_slot="Gece Etüt Ara Öğünü",
            dish_ids=[d_soup],
            sync_to_diabetes=True,
        )
        diab_oct = idm.get_school_menus_for_month(self.connect, "2026-10")
        self.assertTrue(any(m["menu_date"] == "2026-10-05" and m["meal_type"] == "Sahur / Gece" for m in diab_oct))

        active_3 = ["Kuşluk Kahvaltısı", "Öğle Yemeği", "Gece Etüt Ara Öğünü"]
        plans_oct = caf.get_menu_plans_for_month(self.connect, "2026-10")
        self.assertIn(("2026-10-05", "Kuşluk Kahvaltısı"), plans_oct)

        # 2. Each meal keeps its own calories, carbs & allergens strictly under its own column
        workdays_oct = [d for d in caf.get_month_workdays_info(self.connect, "2026-10") if d.get("is_workday")]
        col_rows = caf.build_planned_menu_columnar_rows(
            workdays_oct, plans_oct, active_3, meal_times=sched_c["meal_times"]
        )
        self.assertEqual(len(col_rows), 1)
        kusluk_col = col_rows[0]["Kuşluk Kahvaltısı (09:30 - 10:00)"]
        ogle_col = col_rows[0]["Öğle Yemeği (12:00 - 13:00)"]
        self.assertIn("Peynirli Börek", kusluk_col)
        self.assertIn("370 kcal", kusluk_col)
        self.assertIn("Gluten", kusluk_col)
        self.assertNotIn("Mercimek Çorbası", kusluk_col)
        self.assertIn("Mercimek Çorbası · Pirinç Pilavı", ogle_col)
        self.assertIn("430 kcal", ogle_col)
        self.assertIn("İkame: Çölyaklı öğrenciye haşlama patates", ogle_col)

        # 3. A4 Multi-Meal Page Modes: combined (1 sheet) vs per_meal_page (3 sheets) vs auto_2_per_page (2 sheets)
        h_comb = caf.build_menu_a4_html(
            self.connect,
            year_month="2026-10",
            active_slots=active_3,
            school_meta=cfg_custom,
            meal_page_mode="combined",
            show_per_dish_nutrition=True,
        )
        self.assertEqual(h_comb.count("<div class='a4-sheet'>"), 1)
        self.assertIn("180 kcal, 28g K", h_comb)  # per-dish nutrition detail
        self.assertNotIn("Günlük Toplam", h_comb)  # each meal's kcal/allergen stays under its own column

        h_per_meal = caf.build_menu_a4_html(
            self.connect,
            year_month="2026-10",
            active_slots=active_3,
            school_meta=cfg_custom,
            meal_page_mode="per_meal_page",
        )
        self.assertEqual(h_per_meal.count("<div class='a4-sheet'>"), 3)
        self.assertIn("3 A4 Sayfa", h_per_meal)
        sheets_split = h_per_meal.split("<div class='a4-sheet'>")[1:]
        self.assertIn("Kuşluk Kahvaltısı", sheets_split[0])
        self.assertNotIn("Gece Etüt Ara Öğünü", sheets_split[0])
        self.assertIn("Öğle Yemeği", sheets_split[1])
        self.assertIn("Gece Etüt Ara Öğünü", sheets_split[2])

        h_auto2 = caf.build_menu_a4_html(
            self.connect,
            year_month="2026-10",
            active_slots=active_3,
            school_meta=cfg_custom,
            meal_page_mode="auto_2_per_page",
        )
        self.assertEqual(h_auto2.count("<div class='a4-sheet'>"), 2)

        d_per_meal_bytes = caf.build_menu_a4_docx(
            self.connect,
            year_month="2026-10",
            active_slots=active_3,
            school_meta=cfg_custom,
            meal_page_mode="per_meal_page",
        )
        d_pm_doc = docx.Document(io.BytesIO(d_per_meal_bytes))
        # 3 meal tables + 3 signer tables = 6 tables across 3 separated pages
        self.assertGreaterEqual(len(d_pm_doc.tables), 6)
        d_pm_txt = _extract_docx_all_text(d_per_meal_bytes)
        self.assertIn("İkame: Çölyaklı öğrenciye haşlama patates", d_pm_txt)

        # Multi-slot institutional Excel puts each meal's dishes, kcal/carbs & allergens under its own column headers
        persantil.save_app_settings({"cafeteria_active_meal_slots": ",".join(active_3), "cafeteria_meals_per_day": "3"})
        wb_ms_bytes = caf.build_school_menu_workbook_excel(self.connect, "2026-10", school_meta=persantil.get_app_settings())
        wb_ms = openpyxl.load_workbook(io.BytesIO(wb_ms_bytes))
        try:
            ws_ms = wb_ms["Yemek_Listesi"]
            hdrs_ms = [ws_ms.cell(4, c).value for c in range(1, ws_ms.max_column + 1)]
            self.assertIn("Kuşluk Kahvaltısı — Yemekler", hdrs_ms)
            self.assertIn("Kuşluk Kahvaltısı — Kalori & Karb.", hdrs_ms)
            self.assertIn("Kuşluk Kahvaltısı — Alerjen", hdrs_ms)
            self.assertIn("Öğle Yemeği — Alerjen", hdrs_ms)
        finally:
            wb_ms.close()

        # 4. Ramazan Mode: Sahur & İftar times rendered directly UNDER the date in A4 HTML, Word & Columnar rows
        ram_cfg = {
            **persantil.get_app_settings(),
            "cafeteria_ramazan_mode": "1",
            "cafeteria_ramazan_province": "Ankara",
        }
        with mock.patch("urllib.request.urlopen", side_effect=OSError("offline")):
            ank_rt = caf.get_ramazan_times_for_dates(self.connect, "Ankara", ["2026-10-05"])["2026-10-05"]
            for lm_r in ("standard", "compact_one_page", "weekly_grid"):
                h_ram = caf.build_menu_a4_html(self.connect, year_month="2026-10", active_slots=active_3[:2], layout_mode=lm_r, school_meta=ram_cfg)
                self.assertIn("ram-date-sub", h_ram)
                self.assertIn(f"Sahur:", h_ram)
                self.assertIn(ank_rt["sahur"], h_ram)
                self.assertIn(ank_rt["iftar"], h_ram)
                d_ram_t = _extract_docx_all_text(
                    caf.build_menu_a4_docx(self.connect, year_month="2026-10", active_slots=active_3[:2], layout_mode=lm_r, school_meta=ram_cfg)
                )
                self.assertIn(f"Sahur: {ank_rt['sahur']}", d_ram_t)
                self.assertIn(f"İftar: {ank_rt['iftar']}", d_ram_t)

        # 5. Aligned cross-month copy, day swap, rotation templates, offline rule engine, stock match & allergy prep docx
        # Swap 2026-10-05 with empty 2026-10-06 -> moves plans to Oct 6 and clears Oct 5
        swapped_n = caf.swap_menu_plan_days(self.connect, "2026-10-05", "2026-10-06", meal_slots=active_3)
        self.assertEqual(swapped_n, 3)
        plans_after_sw = caf.get_menu_plans_for_month(self.connect, "2026-10")
        self.assertNotIn(("2026-10-05", "Öğle Yemeği"), plans_after_sw)
        self.assertIn(("2026-10-06", "Öğle Yemeği"), plans_after_sw)

        # Aligned copy from 2026-10 (Tuesday Oct 6) -> 2026-11 (fills all Tuesdays in November!)
        cp_res = caf.copy_month_menu_aligned(self.connect, "2026-10", "2026-11", meal_slots=["Öğle Yemeği"])
        self.assertGreaterEqual(cp_res["copied"], 4)

        # Save & apply rotation template
        st_res = caf.save_menu_rotation_template(
            self.connect, persantil.save_app_settings, persantil.get_app_settings(),
            template_name="Sonbahar Standart", source_ym="2026-11", meal_slots=["Öğle Yemeği"],
        )
        self.assertTrue(st_res["ok"])
        ap_res = caf.apply_menu_rotation_template(
            self.connect, persantil.get_app_settings(), template_name="Sonbahar Standart", target_ym="2026-12",
        )
        self.assertTrue(ap_res["ok"])
        self.assertGreater(ap_res["applied"], 0)

        # Offline nutrition rules detect weekly repetition & calorie band deviation
        eval_nov = caf.evaluate_menu_nutrition_rules(
            caf.get_menu_plans_for_month(self.connect, "2026-11"),
            [d for d in caf.get_month_workdays_info(self.connect, "2026-11") if d.get("is_workday")],
            active_slots=["Öğle Yemeği"],
            kcal_band=(650.0, 800.0),
        )
        self.assertTrue(any(w["Kural"] == "MEB Kalori Bandı" for w in eval_nov["warnings"]))

        # Expiring stock matcher + live portion cost map + allergy kitchen prep sheet (.docx)
        from datetime import date, timedelta
        exp_dt = (date.today() + timedelta(days=4)).isoformat()
        caf.upsert_purchase_invoice(
            self.connect,
            invoice_date="2026-10-01",
            supplier_name="Stok Market",
            items=[{"item_name": "Kırmızı Mercimek", "quantity": 5.0, "unit_type": "kg", "unit_price": 80.0, "expiry_date": exp_dt}],
        )
        cost_m = caf.estimate_all_dishes_cost_map(self.connect)
        self.assertAlmostEqual(cost_m[d_soup], 4.8, places=2)
        exp_sug = caf.match_expiring_stock_to_dishes(self.connect, within_days=14)
        self.assertTrue(any("Mercimek Çorbası" in s["suggested_dishes"] for s in exp_sug))

        prep_docx = _extract_docx_all_text(
            caf.build_allergy_kitchen_prep_docx(self.connect, "2026-10", school_meta=persantil.get_app_settings())
        )
        self.assertIn("ÖZEL DİYET", prep_docx)
        self.assertIn("Ali Kaya", prep_docx)

        # 6. Headless UI render of enhanced Menu Planner with 3 active custom/renamed meals + crowding hint + SKT banner
        revir_ui.reset_session()
        revir_ui.session_state["caf_global_sel_ym"] = "2026-10"
        ctx_ui = {
            "connect_db": self.connect,
            "get_app_settings": persantil.get_app_settings,
            "save_app_settings": persantil.save_app_settings,
            "_get_active_gemini_config": lambda: (False, "", ""),
            "mask_kvkk_value": persantil.mask_kvkk_value,
            "clear_all_data_caches": persantil.clear_all_data_caches,
            "is_section_visible": persantil.is_section_visible,
        }
        out_mp = revir_ui.run_action_headless(
            lambda: caf._render_tab_menu_planner(ctx_ui, "2026-10", persantil.get_app_settings()),
            {"action": "render"},
        )
        mp_html = str(out_mp.get("main_html") or "")
        self.assertNotIn("Çalışma Zamanı Hatası", mp_html)
        self.assertIn("✏️ Haftalık Menü Planla", mp_html)
        self.assertIn("🖨️ A4 Çıktı &amp; Canlı Önizleme", mp_html)
        self.assertIn("✨ AI, Şablon &amp; Araçlar", mp_html)
        self.assertIn("caf_btn_jump_a4_studio", mp_html)
        self.assertIn("Öğün Yönetimi: Öğün Adlarını Değiştir", mp_html)
        self.assertIn("Sayfa Sığdırma Önerisi", mp_html)
        self.assertIn("SKT / Sıfır Atık Stok Önerisi", mp_html)
        self.assertIn("Gece Etüt Ara Öğünü", mp_html)
        self.assertIn("setA4Zoom", h_comb)

        # Clicking "🖨️ A4 Önizle & Çıktı" shortcut activates the Split-Screen A4 Studio tab
        out_jump = revir_ui.run_action_headless(
            lambda: caf._render_tab_menu_planner(ctx_ui, "2026-10", persantil.get_app_settings()),
            {"action": "click", "key": "caf_btn_jump_a4_studio"},
        )
        self.assertNotIn("Çalışma Zamanı Hatası", str(out_jump.get("main_html") or ""))

        # Clean reset of custom slot & default renames
        caf.delete_custom_meal_slot(self.connect, persantil.save_app_settings, persantil.get_app_settings(), "Gece Etüt Ara Öğünü", delete_plans=True)
        caf.reset_default_meal_slot_names(self.connect, persantil.save_app_settings, persantil.get_app_settings())
        sched_restored = caf.parse_service_schedule_from_cfg(persantil.get_app_settings())
        self.assertIn("Sabah Kahvaltısı", sched_restored["available_slots"])
        self.assertNotIn("Gece Etüt Ara Öğünü", sched_restored["available_slots"])


if __name__ == "__main__":
    unittest.main()
