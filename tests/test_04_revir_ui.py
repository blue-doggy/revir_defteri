"""Unit tests for revir_ui.py (widgets, session, charts, caches, headless & HTTP server)."""
from __future__ import annotations

import json
import re
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
                # Edge/corner table & modal resizing CSS/HTML/JS hooks
                self.assertIn("revir-table-outer", body)
                self.assertIn("revir-tbl-handle-r", body)
                self.assertIn("revir-tbl-handle-l", body)
                self.assertIn("revir-tbl-handle-b", body)
                self.assertIn("revir-tbl-handle-se", body)
                self.assertIn("revir-tbl-handle-sw", body)
                self.assertIn("revir-col-resizer", body)
                self.assertIn("revir-modal-resize-r", body)
                self.assertIn("revir-modal-resize-l", body)
                self.assertIn("revir-modal-resize-b", body)
                self.assertIn("revir-modal-resize-se", body)
                self.assertIn('id="revir-resize-badge"', body)
                self.assertIn("window.__revirInitResizableTables", body)
                self.assertIn("colResizeAt", body)
                self.assertIn("__revirLastResizeEndAt", body)
                # Collapsible UI polish & Page box layout (excluding calendar) hooks
                for marker in (
                    "revir-pop-close",
                    "revir-modal-dirty-pill",
                    "revir-chat-pin-btn",
                    "revir-chat-resize-nw",
                    "window.__revirToggleChatPin",
                    "window.__revirClosePopovers",
                    "isCalendarPageName",
                    "window.__revirToggleLayoutEdit",
                    "window.__revirInitBoxLayout",
                    "revir-box-res-se",
                    "revir_box_layout_v1",
                    "revir-layout-bar",
                    ".revir-grid-zone",
                    "revir-box-swap-pill",
                    "revir-box-span-tag",
                    "inferMinGridCols",
                    "container-type: inline-size",
                    ".revir-tab-pane.revir-grid-zone:not(.active)",
                    "ownTabsWrap",
                    "@media (max-width: 760px)",
                    "revir-settings-status-grid",
                    "window.__revirTabKey",
                ):
                    self.assertIn(marker, body)
                out_box_off = revir_ui.run_action_headless(lambda: (_app(), revir_ui.configure_box_layout(False)), {"action": "render"})
                self.assertFalse(out_box_off.get("box_layout_enabled"))
                out_box_on = revir_ui.run_action_headless(lambda: (_app(), revir_ui.configure_box_layout(True)), {"action": "render"})
                self.assertTrue(out_box_on.get("box_layout_enabled"))
        finally:
            srv.stop()

    def test_table_wrappers_and_modal_dialog_resize_markers(self):
        def _dlg_app():
            @revir_ui.dialog("Pencere İçi Tablo", width="large")
            def _open():
                revir_ui.dataframe(pd.DataFrame([{"Ad": "Ayşe", "Boy": 135}]), key="dlg_df_1")
                revir_ui.data_editor(pd.DataFrame([{"Sutun": "A", "Aktif": True}]), key="dlg_ed_1")
                revir_ui.markdown("| Sütun 1 | Sütun 2 |\n|---|---|\n| 10 | 20 |")

            if "open_test_dlg" not in revir_ui.session_state:
                revir_ui.session_state["open_test_dlg"] = True
                _open()

        out = revir_ui.run_action_headless(_dlg_app, {"action": "render"})
        dlg_html = out.get("dialog_html") or ""
        self.assertIn('class="revir-table-outer" data-tbl-key="dlg_df_1" data-tbl-explicit="1"', dlg_html)
        self.assertIn('class="revir-table-outer" data-tbl-key="dlg_ed_1" data-tbl-explicit="1"', dlg_html)
        self.assertIn('class="revir-md-table-wrap"', dlg_html)

    def test_lazy_tabs_render_only_the_active_panel(self):
        def _app():
            tabs = revir_ui.tabs(["Ağır Sekme", "İkinci Sekme"], lazy=True)
            tabs.render(0, lambda: revir_ui.text("İlk panel içeriği"))
            tabs.render(1, lambda: revir_ui.text("İkinci panel içeriği"))

        first = revir_ui.run_action_headless(_app, {"action": "render"})
        first_html = str(first.get("main_html") or "")
        self.assertIn('data-rerender-on-tab="true"', first_html)
        self.assertIn("İlk panel içeriği", first_html)
        self.assertNotIn("İkinci panel içeriği", first_html)

        tab_match = re.search(r'<div class="revir-tab-bar" id="([^"]+)"', first_html)
        if tab_match is None:
            self.fail("Sekme listesi HTML içinde bulunamadı.")
        revir_ui._ACTIVE_SESSION.active_tabs[tab_match.group(1)] = 1
        second = revir_ui.run_action_headless(_app, {"action": "render"})
        second_html = str(second.get("main_html") or "")
        self.assertIn("İkinci panel içeriği", second_html)
        self.assertNotIn("İlk panel içeriği", second_html)

        # Nested non-lazy tabs inside a lazy outer tab must keep distinct tab groups & inactive panes hidden
        def _nested_app():
            outer = revir_ui.tabs(["Ana 1", "Ana 2"], lazy=True)
            def _pane0():
                inner = revir_ui.tabs(["İç A", "İç B"], lazy=False)
                with inner[0]:
                    revir_ui.text("İçerik A")
                with inner[1]:
                    revir_ui.text("İçerik B")
            outer.render(0, _pane0)
            outer.render(1, lambda: revir_ui.text("Ana 2 içeriği"))

        nested_out = revir_ui.run_action_headless(_nested_app, {"action": "render"})
        nested_html = str(nested_out.get("main_html") or "")
        self.assertEqual(nested_html.count('data-rerender-on-tab="true"'), 1)
        self.assertIn("İçerik A", nested_html)
        self.assertIn("İçerik B", nested_html)
        self.assertIn(".revir-tab-pane.revir-grid-zone:not(.active)", revir_ui._SHELL_HTML)

    def test_native_audio_controller_status(self):
        ctrl = revir_ui._NativeAudioController()
        st = ctrl.get_voice_status()
        self.assertIsInstance(st, dict)
        self.assertFalse(ctrl.is_speaking())
        ctrl.cancel_recording()
        ctrl.stop_speaking()

    def test_dataframe_default_turkish_alphabetical_order_and_locked_headers(self):
        names = ["Zeynep", "İbrahim", "Işıl", "Çetin", "Cem", "Ömer", "Şule", "Ümit", "Ahmet"]
        sorted_names = sorted(names, key=revir_ui.tr_alpha_sort_key)
        self.assertEqual(
            sorted_names,
            ["Ahmet", "Cem", "Çetin", "Işıl", "İbrahim", "Ömer", "Şule", "Ümit", "Zeynep"],
        )

        df_in = pd.DataFrame([
            {"id": 10, "Ad Soyad": "Zeynep Kaya", "Sınıf": "4-A"},
            {"id": 20, "Ad Soyad": "Çetin Demir", "Sınıf": "1-B"},
            {"id": 30, "Ad Soyad": "Ahmet Yılmaz", "Sınıf": "3-C"},
        ])

        def _tbl_app():
            revir_ui.dataframe(
                df_in,
                hide_index=True,
                on_select="rerun",
                selection_mode=["single-row", "single-cell"],
                key="tbl_alpha_lock",
                column_config={"id": None},
            )

        out = revir_ui.run_action_headless(_tbl_app, {"action": "render"})
        html = str(out.get("main_html") or "")
        # Column header click sorting must be disabled (no onclick sort or sort arrows)
        self.assertNotIn('onclick="window.__revirSortTable', html)
        self.assertNotIn("revir-sort-ind", html)
        # Rows must render in default Turkish A-Z order by Ad Soyad
        self.assertLess(html.index("Ahmet Yılmaz"), html.index("Çetin Demir"))
        self.assertLess(html.index("Çetin Demir"), html.index("Zeynep Kaya"))
        # Original DataFrame row index (2 for Ahmet Yılmaz) preserved in cell click handler
        self.assertIn('__revirDfSelect("tbl_alpha_lock", 2,', html)


if __name__ == "__main__":
    unittest.main(verbosity=2)
