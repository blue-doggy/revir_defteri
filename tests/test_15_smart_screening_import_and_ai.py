"""
Okul Sağlık Taraması — Akıllı Excel Şablonları, Hibrit Dosya/Görsel/Metin Okuyucu & Gemini AI Toplu Tarama Test Paketi
(test_15_smart_screening_import_and_ai.py)
"""
from __future__ import annotations

import base64
import io
from pathlib import Path
import sys
import unittest
from unittest import mock

import docx
import openpyxl
from PIL import Image, ImageDraw

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

import persantil
import revir_ui
import screening_import as simport
from tests.db_test_helper import isolated_real_db


USER_EXACT_ADULT_TABLE = """Ad Soyad\tBoy\tKilo\tTansiyon\tŞeker\tNabız
Müberra Görgü\t168\t77,6\t115-60\t96\t89
Şebnem Ülker\t159\t63,90\t100-60\t103\t68
Mesliha Şimşek\t162\t75,75\t117-70\t116\t94
Filiz Akdağ\t168\t53,6\t110-70\t90\t87
Havva Gültar\t158\t79,95\t140-80\t119\t97
Adile Akkaya\t162\t64,20\t100-60\t123\t88
Aycan Dilekçi\t170\t79,35\t110-60\t162\t187
Rabia Canıvar\t156\t65,75\t140-90\t179\t88
Fadime Kocataş\t157\t63,6\t140-90\t104\t84
Hatice Kozal\t147\t55,50\t130-90\t102\t87
Ayşegül Akdoğan\t163\t74,10\t121-77\t101\t75
Nazlı Alkan\t157\t63,70\t100-60\t104\t83
Seda Deveci\t161\t69,30\t124-82\t114\t84
Betül Cansu\t169\t83\t120-70\t-\t71
Esra Erboy\t180\t70\t120-60\t111\t
76"""


class TestSmartScreeningImportAndAI(unittest.TestCase):
	def setUp(self) -> None:
		revir_ui.reset_session()

	def test_01_user_exact_15_adult_paste_with_wrapped_line_and_dash_placeholder(self) -> None:
		recs, grp, _note = simport.parse_smart_screening_source(raw_text=USER_EXACT_ADULT_TABLE, target_group="auto")
		self.assertEqual(grp, "adult")
		self.assertEqual(len(recs), 15)

		by_name = {r["full_name"]: r for r in recs}
		muberra = by_name["Müberra Görgü"]
		self.assertEqual(muberra["height_cm"], 168.0)
		self.assertAlmostEqual(muberra["weight_kg"], 77.6, places=2)
		self.assertEqual((muberra["bp_sys"], muberra["bp_dia"]), (115, 60))
		self.assertEqual(muberra["glucose_mg_dl"], 96.0)
		self.assertEqual(muberra["pulse_bpm"], 89)

		# '-' şeker değeri None olmalı, nabız 71 korunmalı
		betul = by_name["Betül Cansu"]
		self.assertIsNone(betul["glucose_mg_dl"])
		self.assertEqual(betul["pulse_bpm"], 71)
		self.assertEqual((betul["bp_sys"], betul["bp_dia"]), (120, 70))

		# Alt satıra kaymış '76' nabız değeri Esra Erboy'un satırına birleştirilmeli!
		esra = by_name["Esra Erboy"]
		self.assertEqual(esra["height_cm"], 180.0)
		self.assertEqual(esra["weight_kg"], 70.0)
		self.assertEqual((esra["bp_sys"], esra["bp_dia"]), (120, 60))
		self.assertEqual(esra["glucose_mg_dl"], 111.0)
		self.assertEqual(esra["pulse_bpm"], 76)

		# Aycan Dilekçi'nin 187 nabzı kaybolmadan korunmalı
		self.assertEqual(by_name["Aycan Dilekçi"]["pulse_bpm"], 187)

		# Tek-transaction veritabanı kaydı doğrulaması
		with isolated_real_db() as (_tmp_p, _db_p, _connect):
			res = persantil._commit_smart_screening_records(recs, "2026-11-27", screened_by="Hemşire Ayşe")
			self.assertEqual(res["added_adults"], 15)
			self.assertEqual(res["screened_count"], 15)
			self.assertEqual(res["vitals_count"], 15)

			adult_rows = persantil._get_adult_screening_rows()
			self.assertEqual(len(adult_rows), 15)
			esra_db = next(a for a in adult_rows if a["full_name"] == "Esra Erboy")
			self.assertEqual(esra_db["pulse_bpm"], 76)
			self.assertEqual(esra_db["sys_mmhg"], 120)
			self.assertEqual(esra_db["dia_mmhg"], 60)
			self.assertEqual(esra_db["glucose_mg_dl"], 111.0)

	def test_02_student_and_adult_excel_templates_blank_and_roster_roundtrip(self) -> None:
		with isolated_real_db() as (_tmp_p, _db_p, connect):
			with connect() as conn:
				c_id = int(conn.execute("INSERT INTO classes(name) VALUES ('5-B')").lastrowid)
				conn.execute(
					"INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date) VALUES (?, 'Öğrenci', '204', 'Kerem Yıldız', 'E', '2015-03-18')",
					(c_id,),
				)
				conn.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sivil / Diğer')")
				civ_id = int(conn.execute("SELECT id FROM classes WHERE name='Sivil / Diğer'").fetchone()["id"])
				conn.execute(
					"INSERT INTO students(class_id, person_type, full_name, sex, notes) VALUES (?, 'Sivil / diğer kişi', 'Elif Öğretmen', 'K', 'Sınıf Öğretmeni')",
					(civ_id,),
				)

			stu_tpl = simport.build_student_screening_excel_template(connect, include_roster=True)
			wb_s = openpyxl.load_workbook(io.BytesIO(stu_tpl))
			ws_s = wb_s["Ogrenci_Taramasi"]
			self.assertIn("Bilgi_ve_Kurallar", wb_s.sheetnames)
			self.assertEqual(ws_s.cell(2, 3).value, "Kerem Yıldız")
			ws_s.cell(2, 6, 144.5)  # Boy
			ws_s.cell(2, 7, 38.2)   # Kilo
			ws_s.cell(2, 10, 82)    # Nabız
			out_s = io.BytesIO()
			wb_s.save(out_s)
			wb_s.close()

			recs_s, grp_s, _ = simport.parse_smart_screening_source(file_bytes=out_s.getvalue(), file_name="ogrenci.xlsx")
			self.assertEqual(grp_s, "student")
			self.assertEqual(len(recs_s), 1)
			self.assertEqual(recs_s[0]["full_name"], "Kerem Yıldız")
			self.assertEqual(recs_s[0]["height_cm"], 144.5)
			self.assertEqual(recs_s[0]["pulse_bpm"], 82)

			ad_tpl = simport.build_adult_screening_excel_template(connect, include_roster=True)
			wb_a = openpyxl.load_workbook(io.BytesIO(ad_tpl))
			ws_a = wb_a["Yetiskin_Taramasi"]
			self.assertEqual(ws_a.cell(2, 1).value, "Elif Öğretmen")
			ws_a.cell(2, 5, 166)
			ws_a.cell(2, 6, 62.4)
			ws_a.cell(2, 7, "120/78")
			ws_a.cell(2, 8, 94)
			ws_a.cell(2, 10, 74)
			out_a = io.BytesIO()
			wb_a.save(out_a)
			wb_a.close()

			recs_a, grp_a, _ = simport.parse_smart_screening_source(file_bytes=out_a.getvalue(), file_name="yetiskin.xlsx")
			self.assertEqual(grp_a, "adult")
			self.assertEqual(len(recs_a), 1)
			self.assertEqual((recs_a[0]["bp_sys"], recs_a[0]["bp_dia"]), (120, 78))

			res_upd = persantil._commit_smart_screening_records(recs_a, "2026-11-27")
			self.assertEqual(res_upd["added_adults"], 0)
			self.assertEqual(res_upd["updated_profiles"], 1)
			self.assertEqual(res_upd["screened_count"], 1)

	def test_03_gemini_chat_bulk_adult_screening_and_no_kvkk_name_corruption(self) -> None:
		with isolated_real_db() as (_tmp_p, _db_p, connect):
			# Aktif seçili öğrenci adı 'Esra Demir' olsa bile 'Esra Erboy' adı bozulmamalı!
			with connect() as conn:
				c_id = int(conn.execute("INSERT INTO classes(name) VALUES ('3-A')").lastrowid)
				esra_stu_id = int(
					conn.execute(
						"INSERT INTO students(class_id, person_type, full_name, sex) VALUES (?, 'Öğrenci', 'Esra Demir', 'K')",
						(c_id,),
					).lastrowid
				)
			revir_ui.session_state["active_person_id"] = esra_stu_id

			persantil.save_app_settings({"gemini_enabled": "1", "gemini_api_key": "test-fake-key"})

			# Gemini metin yanıtı döndürse ve JSON etiketi eksik kalsa dahi yerel hibrit ayrıştırıcı 15/15 kişiyi yakalamalı!
			with mock.patch("gemini_assistant.call_gemini", return_value=(True, "15 yetişkinin tarama verileri onayınıza hazırlandı.", "gemini-2.5-flash")) as m_call:
				chat_res = persantil._handle_floating_gemini_chat(
					{"message": USER_EXACT_ADULT_TABLE + "\n\nyetişkinlere profil oluşturup tarama verisini kaydet"},
					[],
				)
				sent_prompt = m_call.call_args[0][0]
				self.assertIn("Esra Erboy", sent_prompt)
				self.assertNotIn("[KVKK-İSİM] Erboy", sent_prompt)

			self.assertTrue(chat_res["ok"])
			prop = chat_res.get("write_proposal")
			self.assertIsNotNone(prop)
			self.assertIn("15 Kişi", prop["title"])

			apply_res = persantil._handle_floating_gemini_apply_write({"token": prop["token"]})
			self.assertTrue(apply_res["ok"])

			adults = persantil._get_adult_screening_rows()
			self.assertEqual(len(adults), 15)
			self.assertTrue(any(a["full_name"] == "Esra Erboy" and a["pulse_bpm"] == 76 for a in adults))

	def test_04_word_image_vision_and_screening_panel_upload_tab(self) -> None:
		with isolated_real_db() as (_tmp_p, _db_p, _connect):
			persantil.save_app_settings({"gemini_enabled": "1", "gemini_api_key": "test-fake-key"})

			# 1. Okul Sağlık Taraması panelinde "📥 Şablondan / Dosyadan Yükle" sekmesi bulunmalı
			panel_out = revir_ui.run_action_headless(persantil.render_school_screening, {"action": "render"})
			self.assertIn("📥 Şablondan / Dosyadan Yükle", panel_out["main_html"])

			imp_out = revir_ui.run_action_headless(persantil.render_screening_import_section, {"action": "render"})
			self.assertIn("Öğrenci Boş Şablon", imp_out["main_html"])
			self.assertIn("Yetişkin Boş Şablon", imp_out["main_html"])

			# 2. Word (.docx) yetişkin tablosu okuma
			doc = docx.Document()
			tbl = doc.add_table(rows=2, cols=6)
			for idx, h in enumerate(["Ad Soyad", "Boy", "Kilo", "Tansiyon", "Şeker", "Nabız"]):
				tbl.cell(0, idx).text = h
			for idx, v in enumerate(["Selma Kaya", "164", "68,5", "125-80", "99", "78"]):
				tbl.cell(1, idx).text = v
			w_buf = io.BytesIO()
			doc.save(w_buf)
			w_recs, w_grp, _ = simport.parse_smart_screening_source(file_bytes=w_buf.getvalue(), file_name="yetiskin_liste.docx")
			self.assertEqual(w_grp, "adult")
			self.assertEqual(len(w_recs), 1)
			self.assertEqual((w_recs[0]["bp_sys"], w_recs[0]["bp_dia"]), (125, 80))

			# 3. Görsel (.jpg) -> Gemini Multimodal Vision AI ile okuma
			img = Image.new("RGB", (600, 300), (255, 255, 255))
			ImageDraw.Draw(img).text((20, 20), "Tarama Listesi", fill=(0, 0, 0))
			j_buf = io.BytesIO()
			img.save(j_buf, format="JPEG")

			fake_vision_json = '''{
				"mode": "rows",
				"target_group": "adult",
				"explanation": "Fotoğraftan 2 yetişkin tarama satırı okundu",
				"rows": [
					{"person_group": "adult", "full_name": "Murat Demir", "role": "Öğretmen", "height_cm": 178, "weight_kg": 84.5, "bp": "130/85", "glucose_mg_dl": 105, "pulse_bpm": 76},
					{"person_group": "adult", "full_name": "Berna Koç", "role": "Memur", "height_cm": 165, "weight_kg": 61.0, "bp": "110-70", "glucose_mg_dl": null, "pulse_bpm": 72}
				]
			}'''
			with mock.patch("gemini_assistant.call_gemini", return_value=(True, fake_vision_json, "gemini-3-flash-preview")):
				ok_v, v_rows, v_grp, _v_note = persantil._run_ai_smart_screening_parse(
					j_buf.getvalue(), "foto_tarama.jpg", "", "auto", "", "Öğretmen"
				)
			self.assertTrue(ok_v)
			self.assertEqual(v_grp, "adult")
			self.assertEqual(len(v_rows), 2)
			self.assertEqual((v_rows[0]["bp_sys"], v_rows[0]["bp_dia"]), (130, 85))
			self.assertIsNone(v_rows[1]["glucose_mg_dl"])


if __name__ == "__main__":
	unittest.main()
