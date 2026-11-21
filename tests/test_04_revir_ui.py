"""Unit tests for revir_ui.py (widgets, session, charts, caches, headless & HTTP server)."""
from __future__ import annotations

import json
import sys
import unittest
import urllib.request
from datetime import date, time
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import revir_ui


class TestRevirUI(unittest.TestCase):
    def setUp(self):
        revir_ui.reset_session()

    def test_session_state_uploaded_file_and_selection(self):
        ss = revir_ui.SessionState()
        ss.foo = 42
        self.assertEqual(ss.foo, 42)
        del ss.foo
        with self.assertRaises(AttributeError):
            _ = ss.foo

        uf = revir_ui.UploadedFile("deneme_ğüşiöç.txt", b"abc", "text/plain", "f1")
        self.assertEqual(uf.getvalue(), b"abc")
        self.assertEqual(uf.name, "deneme_ğüşiöç.txt")

        sel = revir_ui._SelectionState([1, 2])
        self.assertEqual(sel.get("rows"), [1, 2])
        self.assertEqual(sel["rows"], [1, 2])
        ev = revir_ui._DataframeSelectionEvent(sel)
        self.assertEqual(ev.get("selection")["rows"], [1, 2])
        self.assertEqual(ev["selection"]["rows"], [1, 2])

    def test_cache_decorators_and_clearing(self):
        calls = {"d": 0, "r": 0}

        @revir_ui.cache_data
        def _cached_d(x):
            calls["d"] += 1
            return x * 2

        @revir_ui.cache_resource
        def _cached_r(x):
            calls["r"] += 1
            return {"v": x}

        self.assertEqual(_cached_d(5), 10)
        self.assertEqual(_cached_d(5), 10)
        self.assertEqual(calls["d"], 1)
        _cached_d.clear()
        self.assertEqual(_cached_d(5), 10)
        self.assertEqual(calls["d"], 2)

        self.assertEqual(_cached_r("a")["v"], "a")
        self.assertEqual(_cached_r("a")["v"], "a")
        self.assertEqual(calls["r"], 1)
        _cached_r.clear()
        self.assertEqual(_cached_r("a")["v"], "a")
        self.assertEqual(calls["r"], 2)

        revir_ui.clear_ephemeral_widget_values()
        revir_ui.clear_dataframe_selection("df_k")

    def test_markdown_and_charts_and_df_formatting(self):
        md = revir_ui._render_markdown("# Başlık\n- madde 1\n- **kalın** `kod` [link](https://example.com)\n| A | B |\n|---|---|\n| 1 | 2 |")
        self.assertIn("<h1", md)
        self.assertIn("<table", md)

        svg = revir_ui._render_svg_bar_chart({"A": 10, "B": 20})
        self.assertIn("<svg", svg)
        svg_empty = revir_ui._render_svg_bar_chart({})
        self.assertIsInstance(svg_empty, str)

        self.assertEqual(revir_ui._format_df_cell(None, None), "")
        num_cfg = revir_ui.column_config.NumberColumn("X", format="%.2f")
        self.assertEqual(revir_ui._format_df_cell(3.14159, num_cfg), "3.14")
        fp = revir_ui._df_fingerprint(pd.DataFrame([{"a": 1, "b": "x"}]))
        self.assertIsInstance(fp, str)
        chk_cfg = revir_ui.column_config.CheckboxColumn("C")
        self.assertEqual(revir_ui._coerce_editor_value("true", chk_cfg, "bool"), True)
        self.assertEqual(revir_ui._coerce_editor_value("12.5", num_cfg, "float64"), 12.5)

    def test_all_widgets_and_dialog_and_server_headless(self):
        def _app():
            revir_ui.set_page_config(page_title="Test", layout="wide")
            revir_ui.markdown("## Merhaba ğüşiöç")
            revir_ui.caption("Alt yazı")
            revir_ui.write("Yazı", {"k": 1})
            revir_ui.text("Düz metin")
            revir_ui.code("print(1)", language="python")
            revir_ui.divider()
            revir_ui.success("Tamam")
            revir_ui.info("Bilgi")
            revir_ui.warning("Uyarı")
            revir_ui.error("Hata")
            revir_ui.metric("Öğrenci", 120, "+5")
            revir_ui.progress(0.5, text="%50")
            revir_ui.balloons()
            revir_ui.bar_chart({"Ocak": 3, "Şubat": 5})

            c1, c2 = revir_ui.columns(2)
            with c1:
                revir_ui.text_input("Ad", value="Ali", key="t_ad")
                revir_ui.text_area("Not", value="...", key="t_not")
                revir_ui.number_input("Boy", value=140.0, key="n_boy")
                revir_ui.checkbox("Onay", value=True, key="c_onay")
                revir_ui.toggle("Aktif", value=False, key="tg_aktif")
            with c2:
                revir_ui.selectbox("Sınıf", ["1-A", "2-A"], key="s_sinif")
                revir_ui.multiselect("Etiket", ["A", "B"], default=["A"], key="m_etiket")
                revir_ui.radio("Cinsiyet", ["E", "K"], key="r_cins")
                revir_ui.date_input("Tarih", value=date(2026, 11, 22), key="d_tar")
                revir_ui.time_input("Saat", value=time(9, 30), key="tm_saat")
                revir_ui.file_uploader("Dosya", key="f_up")

            with revir_ui.expander("Detay", expanded=True):
                revir_ui.dataframe(pd.DataFrame([{"Ad": "Ali", "Boy": 140}]), key="df_test")
                revir_ui.data_editor(pd.DataFrame([{"Ad": "Ali", "Sec": True}]), key="de_test")
                revir_ui.download_button("İndir", data=b"abc", file_name="a.txt", key="dl_btn")
                revir_ui.link_button("Site", url="https://example.com")
                revir_ui.button("Kaydet", key="btn_save")

            with revir_ui.form("f1"):
                revir_ui.form_submit_button("Gönder")

            t1, t2 = revir_ui.tabs(["Sekme 1", "Sekme 2"])
            with t1:
                revir_ui.text("S1")
            with t2:
                revir_ui.text("S2")

        out = revir_ui.run_action_headless(_app, {"action": "render"})
        self.assertIn("main_html", out)
        self.assertNotIn("Çalışma Zamanı Hatası", out["main_html"])
        self.assertTrue(out["balloons"])

        # Test LocalUIServer HTTP GET/POST & stop
        srv = revir_ui.LocalUIServer(_app, host="127.0.0.1", port=0)
        url = srv.start()
        try:
            with urllib.request.urlopen(url, timeout=5) as resp:
                self.assertEqual(resp.status, 200)
                body = resp.read().decode("utf-8")
                self.assertIn("<!DOCTYPE html>", body)
                # Gemini FAB button is image/SVG only (no visible text label),
                # voice mic / auto-TTS controls removed, and outside-click closes chat bubble
                fab_section = body.split('<button id="revir-fab-btn"')[1].split("</button>")[0]
                self.assertIn("<svg", fab_section)
                self.assertNotIn("Gemini Asistan", fab_section)
                self.assertNotIn("revir-chat-mic-btn", body)
                self.assertNotIn("revir-chat-tts-btn", body)
                self.assertNotIn('id="revir-voice-bar"', body)
                self.assertIn('ev.target.closest("#revir-chat-bubble, #revir-fab-btn")', body)
        finally:
            srv.stop()

    def test_native_audio_controller_status(self):
        ctrl = revir_ui._NativeAudioController()
        st = ctrl.get_voice_status()
        self.assertIsInstance(st, dict)
        self.assertFalse(ctrl.is_speaking())
        ctrl.cancel_recording()
        ctrl.stop_speaking()


if __name__ == "__main__":
    unittest.main(verbosity=2)
