"""
Genel Bakış Sağ Üst Günlük Takvim Yaprağı, Hover ✕ Kapatma ve Widget Çekmecesi (+ / Sürükle-Bırak) Testleri.
"""
from __future__ import annotations

import json
from pathlib import Path
import sys
import unittest
from unittest import mock

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


class TestOverviewDailyLeafAndWidgetTray(unittest.TestCase):
	def setUp(self) -> None:
		revir_ui.reset_session()
		self._ctx = isolated_real_db()
		self.tmp_p, self.db_p, self.connect = self._ctx.__enter__()

	def tearDown(self) -> None:
		self._ctx.__exit__(None, None, None)

	def test_overview_daily_leaf_top_right_placement_and_controls(self) -> None:
		with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
			"İmsak": "05:55", "Güneş": "07:25", "Öğle": "12:55", "İkindi": "15:25", "Akşam": "17:48", "Yatsı": "19:10",
		}):
			out = revir_ui.run_action_headless(persantil.render_overview, {"action": "render"})

		ov_html = str(out.get("main_html") or "")
		ov_scripts = str(out.get("scripts_html") or "")

		# 1. Günlük Takvim Yaprağı Genel Bakış içinde yer almalı
		self.assertIn('data-layout-widget-marker="overview_daily_leaf"', ov_html)
		self.assertIn("🗓️ GÜNLÜK TAKVİM YAPRAĞI", ov_html)
		self.assertIn("revir_ov_prayer_countdown", ov_html)
		self.assertIn("btn_ov_open_cal_event_popup", ov_html)

		# 2. Yeni kurulum varsayılan yerleşimi kullanıcının belirlediği 12 sütunlu düzende gelmeli
		cfg_line = ov_scripts.split("const cfg = ", 1)[1].split(";\n", 1)[0]
		parsed_cfg = json.loads(cfg_line)
		lay = parsed_cfg["layout"]

		self.assertEqual((lay["overview_quick_links"]["col"], lay["overview_quick_links"]["row"], lay["overview_quick_links"]["col_span"], lay["overview_quick_links"]["row_span"]), (0, 0, 7, 3))
		self.assertEqual((lay["overview_recent"]["col"], lay["overview_recent"]["row"], lay["overview_recent"]["col_span"], lay["overview_recent"]["row_span"]), (0, 3, 7, 9))
		self.assertEqual((lay["overview_pending"]["col"], lay["overview_pending"]["row"], lay["overview_pending"]["col_span"], lay["overview_pending"]["row_span"]), (0, 12, 7, 4))
		self.assertEqual((lay["overview_stats"]["col"], lay["overview_stats"]["row"], lay["overview_stats"]["col_span"], lay["overview_stats"]["row_span"]), (0, 16, 7, 4))
		self.assertEqual((lay["overview_outbreak"]["col"], lay["overview_outbreak"]["row"], lay["overview_outbreak"]["col_span"], lay["overview_outbreak"]["row_span"]), (0, 20, 7, 5))

		self.assertEqual((lay["overview_daily_leaf"]["col"], lay["overview_daily_leaf"]["row"], lay["overview_daily_leaf"]["col_span"], lay["overview_daily_leaf"]["row_span"]), (7, 0, 5, 12))
		self.assertEqual((lay["overview_meds"]["col"], lay["overview_meds"]["row"], lay["overview_meds"]["col_span"], lay["overview_meds"]["row_span"]), (7, 12, 5, 4))
		self.assertEqual((lay["overview_diabetes"]["col"], lay["overview_diabetes"]["row"], lay["overview_diabetes"]["col_span"], lay["overview_diabetes"]["row_span"]), (7, 16, 5, 4))
		self.assertEqual((lay["overview_calendar"]["col"], lay["overview_calendar"]["row"], lay["overview_calendar"]["col_span"], lay["overview_calendar"]["row_span"]), (7, 20, 5, 5))
		self.assertEqual((lay["overview_gemini"]["col"], lay["overview_gemini"]["row"], lay["overview_gemini"]["col_span"], lay["overview_gemini"]["row_span"]), (0, 25, 12, 5))
		self.assertTrue(lay["overview_daily_leaf"]["visible"])
		self.assertFalse(lay["overview_chart"]["visible"])

		# 3. Yaprak üzerindeki '+' butonuna tıklandığında etkinlik ekleme diyaloğu açılmalı
		with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
			"İmsak": "05:55", "Güneş": "07:25", "Öğle": "12:55", "İkindi": "15:25", "Akşam": "17:48", "Yatsı": "19:10",
		}):
			out_popup = revir_ui.run_action_headless(
				persantil.render_overview,
				{"action": "click", "key": "btn_ov_open_cal_event_popup"},
			)
		self.assertIsNotNone(out_popup.get("dialog_html"))

	def test_legacy_saved_overview_layout_migrates_for_top_right_leaf(self) -> None:
		# Daha önce overview_daily_leaf yokken kaydedilmiş tam genişlikli (col_span=12) profil sağ üstte yer açmalı
		old_saved = {
			"overview_stats": {
				"id": "overview_stats", "col": 0, "row": 0, "col_span": 12, "row_span": 4,
				"x_pct": 0, "y_px": 0, "w_pct": 100, "h_px": 260, "visible": True, "z": 1, "accent": "#0d9488",
			},
			"overview_quick_links": {
				"id": "overview_quick_links", "col": 0, "row": 4, "col_span": 12, "row_span": 2,
				"x_pct": 0, "y_px": 272, "w_pct": 100, "h_px": 124, "visible": True, "z": 1, "accent": "#0284c7",
			},
		}
		persantil.save_widget_layout("layout_overview", old_saved, "Varsayılan Düzen")

		defaults = {
			"overview_stats": {"id": "overview_stats", "col": 0, "row": 0, "col_span": 7, "row_span": 4, "x_pct": 0, "y_px": 0, "w_pct": 58.33, "h_px": 260, "visible": True},
			"overview_quick_links": {"id": "overview_quick_links", "col": 0, "row": 4, "col_span": 7, "row_span": 3, "x_pct": 0, "y_px": 272, "w_pct": 58.33, "h_px": 192, "visible": True},
			"overview_daily_leaf": {"id": "overview_daily_leaf", "col": 7, "row": 0, "col_span": 5, "row_span": 7, "x_pct": 58.33, "y_px": 0, "w_pct": 41.67, "h_px": 464, "visible": True},
		}
		merged = persantil.get_saved_widget_layout("layout_overview", defaults)
		self.assertEqual(merged["overview_stats"]["col_span"], 7)
		self.assertEqual(merged["overview_quick_links"]["col_span"], 7)
		self.assertEqual(merged["overview_daily_leaf"]["col"], 7)
		self.assertEqual(merged["overview_daily_leaf"]["row"], 0)

	def test_hover_close_x_and_widget_tray_drag_restore_architecture(self) -> None:
		# Kabuk HTML içinde Düzen çubuğu Widgetlar butonu, çekmece, + butonu ve sürükle-bırak motoru yer almalı
		shell = revir_ui._SHELL_HTML
		for marker in (
			'id="revir-layout-widgets-btn"',
			'id="revir-widget-tray"',
			'id="revir-widget-tray-list"',
			'id="revir-wtray-drag-chip"',
			"revir-box-close-x",
			"revir-wtray-add-btn",
			"window.__revirToggleWidgetTray",
			"window.__revirRefreshWidgetTray",
			"bindTrayItemGesture",
			"startFollowPlacementMode",
		):
			self.assertIn(marker, shell)

		# Genel Bakış Smart Grid üzerinde hover X kapatma ve dışarıdan sürükle-bırak API'si yer almalı
		with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
			"İmsak": "05:55", "Güneş": "07:25", "Öğle": "12:55", "İkindi": "15:25", "Akşam": "17:48", "Yatsı": "19:10",
		}):
			out = revir_ui.run_action_headless(persantil.render_overview, {"action": "render"})
		scripts = str(out.get("scripts_html") or "")
		for sm_marker in (
			"revir-card-close-x",
			"closeWidgetById",
			"window.__revirSmartCanvasApi",
			"previewExternalDrop",
			"restoreWidget",
		):
			self.assertIn(sm_marker, scripts)

		# Senkronizasyon üzerinden overview_daily_leaf kapatılıp ardından farklı bir konuma (örn. col=0, row=2) bırakıldığında kalıcı kaydedilmeli
		closed_sync = {
			"overview_daily_leaf": {
				"id": "overview_daily_leaf",
				"x_pct": 58.33, "y_px": 0, "w_pct": 41.67, "h_px": 464,
				"col": 7, "row": 0, "col_span": 5, "row_span": 7,
				"min_col_span": 4, "min_row_span": 4,
				"visible": False, "z": 1, "accent": "#0d9488",
			}
		}
		with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
			"İmsak": "05:55", "Güneş": "07:25", "Öğle": "12:55", "İkindi": "15:25", "Akşam": "17:48", "Yatsı": "19:10",
		}):
			revir_ui.run_action_headless(
				persantil.render_overview,
				{
					"action": "render",
					"pending_inputs": {
						"layout_overview_edit": "true",
						"_widget_layout_sync_layout_overview": json.dumps(closed_sync),
					},
				},
			)
		saved_after_close = persantil.get_saved_widget_layout("layout_overview", closed_sync)
		self.assertFalse(saved_after_close["overview_daily_leaf"]["visible"])

		restored_sync = {
			"overview_daily_leaf": {
				"id": "overview_daily_leaf",
				"x_pct": 0.0, "y_px": 136, "w_pct": 50.0, "h_px": 464,
				"col": 0, "row": 2, "col_span": 6, "row_span": 7,
				"min_col_span": 4, "min_row_span": 4,
				"visible": True, "z": 1, "accent": "#0d9488",
			}
		}
		with mock.patch("school_calendar.get_prayer_times_for_province", return_value={
			"İmsak": "05:55", "Güneş": "07:25", "Öğle": "12:55", "İkindi": "15:25", "Akşam": "17:48", "Yatsı": "19:10",
		}):
			revir_ui.run_action_headless(
				persantil.render_overview,
				{
					"action": "render",
					"pending_inputs": {
						"layout_overview_edit": "true",
						"_widget_layout_sync_layout_overview": json.dumps(restored_sync),
					},
				},
			)
		saved_after_restore = persantil.get_saved_widget_layout("layout_overview", restored_sync)
		self.assertTrue(saved_after_restore["overview_daily_leaf"]["visible"])
		self.assertEqual(saved_after_restore["overview_daily_leaf"]["col"], 0)
		self.assertEqual(saved_after_restore["overview_daily_leaf"]["row"], 2)


if __name__ == "__main__":
	unittest.main()
