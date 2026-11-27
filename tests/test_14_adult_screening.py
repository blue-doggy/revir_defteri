"""
Okul Sağlık Taraması Paneli — Minimalist Yetişkin Taraması Sekmesi Test Paketi (test_14_adult_screening.py)
"""
from __future__ import annotations

from pathlib import Path
import sys
import unittest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestAdultScreeningTab(unittest.TestCase):
	def setUp(self) -> None:
		revir_ui.reset_session()

	def test_adult_screening_classification_rules(self) -> None:
		c_ok = persantil.classify_adult_screening_metrics(22.4, 118, 76, 92, "Açlık")
		self.assertEqual(c_ok["bmi"], ("Normal", "ok"))
		self.assertEqual(c_ok["bp"], ("Normal", "ok"))
		self.assertEqual(c_ok["glucose"], ("Normal", "ok"))

		c_high = persantil.classify_adult_screening_metrics(31.2, 148, 94, 134, "Açlık")
		self.assertEqual(c_high["bmi"], ("Obez", "danger"))
		self.assertEqual(c_high["bp"], ("Yüksek TA", "warn"))
		self.assertEqual(c_high["glucose"], ("Yüksek KŞ", "danger"))

		c_low = persantil.classify_adult_screening_metrics(17.6, 85, 55, 62, "Tokluk")
		self.assertEqual(c_low["bmi"], ("Zayıf", "warn"))
		self.assertEqual(c_low["bp"], ("Düşük TA", "warn"))
		self.assertEqual(c_low["glucose"], ("Düşük KŞ", "danger"))

	def test_save_adult_screening_syncs_measurements_vitals_and_high_adult_weight(self) -> None:
		with isolated_real_db() as (_tmp_p, _db_p, connect):
			with connect() as conn:
				conn.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sivil / Diğer')")
				civ_cid = int(conn.execute("SELECT id FROM classes WHERE name='Sivil / Diğer'").fetchone()["id"])
				adult_id = int(
					conn.execute(
						"""INSERT INTO students(class_id, person_type, full_name, sex, birth_date, notes)
						   VALUES (?, 'Sivil / diğer kişi', 'Selim Öğretmen', 'E', '1982-05-14', 'Matematik Öğretmeni')""",
						(civ_cid,),
					).lastrowid
				)

			# 135 kg yetişkin kilosu çocuk gibi /10'a bölünmeden 135.0 kg olarak kaydedilmeli
			scr_id = persantil.save_adult_screening(
				person_id=adult_id,
				screening_date="2026-11-27",
				height_cm="182",
				weight_kg="135",
				bp_sys="135/85",
				pulse_bpm="74",
				glucose_mg_dl="98",
				glucose_context="Açlık",
				vision_status="Gözlük / Lens Kullanıyor",
				notes="Düzenli yürüyüş önerildi",
				screened_by="Hemşire Ayşe",
			)
			self.assertGreater(scr_id, 0)

			rows = persantil._get_adult_screening_rows()
			self.assertEqual(len(rows), 1)
			r0 = rows[0]
			self.assertTrue(r0["is_screened"])
			self.assertEqual(r0["height_cm"], 182.0)
			self.assertEqual(r0["weight_kg"], 135.0)
			self.assertAlmostEqual(r0["bmi"], round(135.0 / (1.82 ** 2), 1), places=1)
			self.assertEqual(r0["sys_mmhg"], 135)
			self.assertEqual(r0["dia_mmhg"], 85)
			self.assertEqual(r0["pulse_bpm"], 74)
			self.assertEqual(r0["glucose_mg_dl"], 98.0)
			self.assertIn("Gözlük", r0["vision"])

	def test_school_screening_panel_renders_minimalist_adult_tab(self) -> None:
		with isolated_real_db() as (_tmp_p, _db_p, connect):
			with connect() as conn:
				cls_id = int(conn.execute("INSERT INTO classes(name) VALUES ('4-A')").lastrowid)
				conn.execute(
					"INSERT INTO students(class_id, person_type, full_name, sex) VALUES (?, 'Öğrenci', 'Ece Öğrenci', 'K')",
					(cls_id,),
				)
				conn.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sivil / Diğer')")
				civ_cid = int(conn.execute("SELECT id FROM classes WHERE name='Sivil / Diğer'").fetchone()["id"])
				conn.execute(
					"""INSERT INTO students(class_id, person_type, full_name, sex, notes)
					   VALUES (?, 'Sivil / diğer kişi', 'Zeynep Müdür Yardımcısı', 'K', 'Müdür Yardımcısı')""",
					(civ_cid,),
				)

			out_panel = revir_ui.run_action_headless(persantil.render_school_screening, {"action": "render"})
			self.assertIn("🧑 Yetişkin Taraması", out_panel["main_html"])

			out_adult = revir_ui.run_action_headless(persantil.render_adult_screening_tab, {"action": "render"})
			html_body = out_adult["main_html"]
			self.assertIn("Yetişkin &amp; Personel", html_body)
			self.assertIn("Zeynep Müdür Yardımcısı", html_body)
			self.assertIn("Tansiyon (mmHg)", html_body)
			self.assertIn("Kan Şekeri", html_body)
			# Öğrenciler yetişkin taraması listesine karışmamalı
			self.assertNotIn("Ece Öğrenci", html_body)


if __name__ == "__main__":
	unittest.main()
