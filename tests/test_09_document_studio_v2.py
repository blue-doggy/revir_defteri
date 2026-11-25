# -*- coding: utf-8 -*-
"""A4 Canlı Belge Stüdyosu, PNG İmza, Okunaklı İlk Sayfa Kapakları ve Yan Yana Eylem Barı Testleri."""
from __future__ import annotations

import base64
import io
import json
import sys
import unittest
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import document_studio as dstudio
import persantil
import revir_ui
from tests.db_test_helper import isolated_real_db


def _make_sample_signature_png(with_white_bg: bool = True) -> bytes:
    from PIL import Image, ImageDraw

    bg = (255, 255, 255, 255) if with_white_bg else (255, 255, 255, 0)
    im = Image.new("RGBA", (360, 140), bg)
    draw = ImageDraw.Draw(im)
    draw.line([(40, 95), (120, 45), (190, 90), (280, 40)], fill=(15, 44, 89, 255), width=6)
    draw.line([(60, 105), (290, 105)], fill=(15, 44, 89, 255), width=3)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    im.close()
    return buf.getvalue()


def _make_sample_pdf(path: Path) -> None:
    pdf_bytes = dstudio.render_html_body_to_pdf_bytes(
        "<h1>T.C. OKUL REViR MÜDÜRLÜĞÜ</h1>"
        "<p>Öğrenci Sağlık ve Tarama İzlem Formu — Türkçe karakterler: ğüşiöçİĞÜŞÖÇ.</p>"
        "<table><thead><tr><th>Tarih</th><th>Ad Soyad</th><th>Not</th></tr></thead>"
        "<tbody><tr><td>25.11.2026</td><td>Ayşe Yılmaz</td><td>Ateş normal 36.6 °C</td></tr></tbody></table>",
        font_family="Arial",
        font_size_pt=11.0,
    )
    path.write_bytes(pdf_bytes)


def _make_sample_docx(path: Path) -> None:
    import docx

    d = docx.Document()
    d.add_heading("REViR VELİ İZİN VE BİLGİLENDİRME FORMU", level=1)
    d.add_paragraph("Sayın Veli, çocuğunuzun boy ve kilo persantil taraması planlanmıştır.")
    d.add_paragraph("Türkçe karakter testi: ğüşiöç ĞÜŞİÖÇ.")
    t = d.add_table(rows=3, cols=3)
    t.cell(0, 0).text = "Öğrenci No"
    t.cell(0, 1).text = "Ad Soyad"
    t.cell(0, 2).text = "Sınıf"
    t.cell(1, 0).text = "101"
    t.cell(1, 1).text = "Mehmet Demir"
    t.cell(1, 2).text = "4-A"
    t.cell(2, 0).text = "102"
    t.cell(2, 1).text = "Elif Kaya"
    t.cell(2, 2).text = "5-B"
    d.save(path)


def _make_sample_xlsx(path: Path) -> None:
    df = pd.DataFrame(
        [
            {"Sıra": i + 1, "Ad Soyad": f"Öğrenci {i + 1}", "Sınıf": "6-A", "Boy (cm)": 140 + i, "Kilo (kg)": 36 + i * 0.5, "Durum": "Normal"}
            for i in range(12)
        ]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="TaramaListesi")


def _make_sample_pptx(path: Path) -> None:
    import zipfile

    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        '</Types>'
    )
    slide1 = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:cSld><p:spTree><p:sp><p:txBody>'
        '<a:p><a:r><a:t>Okul Revir Hijyen Eğitimi Sunumu</a:t></a:r></a:p>'
        '<a:p><a:r><a:t>El yıkama ve sağlıklı beslenme ilkeleri (ğüşiöç)</a:t></a:r></a:p>'
        '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("ppt/slides/slide1.xml", slide1)


def _make_sample_png(path: Path) -> None:
    from PIL import Image, ImageDraw

    im = Image.new("RGB", (620, 860), "#ffffff")
    draw = ImageDraw.Draw(im)
    draw.rectangle([(24, 24), (596, 836)], outline="#0f172a", width=3)
    draw.text((48, 56), "TARAMA BELGESI", fill="#0f172a", font=dstudio.get_cover_pil_font(22, bold=True))
    im.save(path, format="PNG")
    im.close()


class TestDocumentStudioV2(unittest.TestCase):
    def setUp(self):
        self.ctx = isolated_real_db()
        self.tmp_p, self.db_p, self.connect = self.ctx.__enter__()
        revir_ui.reset_session()
        self.evrak_dir = self.tmp_p / "revir_evraklari"
        self.evrak_dir.mkdir(parents=True, exist_ok=True)
        self._orig_evrak = persantil.EVRAKLAR_DIR
        persantil.EVRAKLAR_DIR = self.evrak_dir

    def tearDown(self):
        persantil.EVRAKLAR_DIR = self._orig_evrak
        revir_ui.reset_session()
        self.ctx.__exit__(None, None, None)

    def test_readable_first_page_covers_all_formats(self):
        from PIL import Image

        pdf_p = self.evrak_dir / "ornek_form.pdf"
        docx_p = self.evrak_dir / "veli_izin.docx"
        xlsx_p = self.evrak_dir / "tarama_tablo.xlsx"
        pptx_p = self.evrak_dir / "egitim_sunum.pptx"
        txt_p = self.evrak_dir / "notlar.txt"
        png_p = self.evrak_dir / "taranan_belge.png"

        _make_sample_pdf(pdf_p)
        _make_sample_docx(docx_p)
        _make_sample_xlsx(xlsx_p)
        _make_sample_pptx(pptx_p)
        txt_p.write_text("Satır 1: Revir günlük kontrol listesi\nSatır 2: İlaç dolabı ısısı 21.4°C", encoding="utf-8")
        _make_sample_png(png_p)

        for p in (pdf_p, docx_p, xlsx_p, pptx_p, txt_p, png_p):
            cov = persantil.get_or_create_document_cover(p, force=True)
            self.assertIsNotNone(cov, f"Kapak üretilemedi: {p.name}")
            self.assertTrue(cov.exists())
            self.assertIn("cov_v2_", cov.name)
            self.assertGreater(cov.stat().st_size, 2500, f"Kapak boş/çok küçük: {p.name}")

            im = Image.open(cov)
            w, h = im.size
            # Excel >4 sütun yatay A4 (920x650), PPTX slayt (960x540), diğerleri yüksek çözünürlüklü sayfa
            if p == xlsx_p:
                self.assertEqual((w, h), (920, 650))
                # Eski sütun/satır girinti hatasını yakala: orta bölgede birden fazla satır mürekkebi olmalı
                mid_crop = im.crop((40, 140, 880, 420)).convert("L")
                dark_pixels = sum(1 for px in mid_crop.getdata() if px < 160)
                self.assertGreater(dark_pixels, 800)
            elif p == pptx_p:
                self.assertGreaterEqual(w, 800)
                self.assertGreater(w, h)
            else:
                self.assertGreaterEqual(w, 550)
                self.assertGreater(h, w)
            im.close()

    def test_png_signature_transparency_and_library(self):
        from PIL import Image

        raw_sig = _make_sample_signature_png(with_white_bg=True)
        clean_png = dstudio.prepare_signature_png_bytes(raw_sig, remove_white_bg=True)
        im = Image.open(io.BytesIO(clean_png)).convert("RGBA")
        # Köşe pikseli şeffaf olmalı (beyaz zemin temizlenmiş), ink pikselleri opak kalmalı
        corner_alpha = im.getpixel((0, 0))[3]
        self.assertEqual(corner_alpha, 0)
        opaque_count = sum(1 for _, _, _, a in im.getdata() if a >= 200)
        self.assertGreater(opaque_count, 300)
        self.assertLess(im.width, 360)  # Kenar beyazlıkları kırpılmış olmalı
        im.close()

        saved_path = dstudio.save_signature_to_library(self.evrak_dir, "Okul Hemşiresi İmza", raw_sig, remove_white_bg=True)
        self.assertTrue(saved_path.exists())
        sigs = dstudio.list_saved_signatures(self.evrak_dir)
        self.assertEqual(len(sigs), 1)
        self.assertIn("Okul Hemşiresi", sigs[0]["name"])

    def test_a4_pdf_docx_xlsx_image_editing_and_signature_stamping(self):
        import docx
        import openpyxl
        import pymupdf
        from PIL import Image

        sig_bytes = dstudio.prepare_signature_png_bytes(_make_sample_signature_png(with_white_bg=True))
        sig_b64 = base64.b64encode(sig_bytes).decode("ascii")

        # 1. PDF yazı tipi/boyut + imza damgalama
        pdf_src = self.evrak_dir / "form_src.pdf"
        pdf_dst = self.evrak_dir / "form_imzali.pdf"
        _make_sample_pdf(pdf_src)
        pdf_fields = {
            0: [
                {
                    "id": "t1",
                    "field_type": "text",
                    "value": "Onaylı Sağlık Taraması — Türkçe ğüşiöç",
                    "x0": 60.0,
                    "y0": 260.0,
                    "x1": 360.0,
                    "y1": 295.0,
                    "font_size": 13.0,
                    "font_family": "Times New Roman",
                    "bold": True,
                    "align": "center",
                    "color": "lacivert",
                },
                {
                    "id": "s1",
                    "field_type": "signature",
                    "value": "[İmza: Hemşire]",
                    "x0": 350.0,
                    "y0": 560.0,
                    "x1": 500.0,
                    "y1": 620.0,
                    "signature_b64": sig_b64,
                },
            ]
        }
        dstudio.save_filled_pdf(pdf_src, pdf_fields, pdf_dst)
        d_out = pymupdf.open(pdf_dst)
        txt_p0 = d_out[0].get_text("text").replace("\xa0", " ")
        imgs_p0 = d_out[0].get_images()
        d_out.close()
        self.assertIn("Onaylı Sağlık Taraması — Türkçe ğüşiöç", txt_p0)
        self.assertGreaterEqual(len(imgs_p0), 1)

        # 2. Word (.docx) yazı tipi, paragraf güncelleme ve PNG imza ekleme
        docx_src = self.evrak_dir / "izin_src.docx"
        docx_dst = self.evrak_dir / "izin_imzali.docx"
        _make_sample_docx(docx_src)
        paras, tables = dstudio._load_docx_structure(docx_src)
        first_idx = int(paras[0]["index"])
        dstudio._save_docx_structure(
            docx_src,
            docx_dst,
            edited_paras={first_idx: "GÜNCELLENMİŞ REViR ONAY FORMU"},
            appended_paras=["Eklenen revir açıklaması."],
            edited_tables={0: tables[0]},
            font_family="Times New Roman",
            font_size_pt=12.0,
            default_align="Ortala",
            signature_overlays=[{"x0": 340, "y0": 620, "x1": 490, "y1": 675, "signature_b64": sig_b64}],
            extra_text_overlays=[{"value": "Revir Onay Notu", "align": "center"}],
        )
        dx = docx.Document(docx_dst)
        full_dx_text = "\n".join(p.text for p in dx.paragraphs)
        self.assertIn("GÜNCELLENMİŞ REViR ONAY FORMU", full_dx_text)
        self.assertIn("Revir Onay Notu", full_dx_text)
        self.assertGreaterEqual(len(dx.inline_shapes), 1)

        # 3. Excel (.xlsx) A4 yazı tipi, sayfa düzeni ve PNG imza damgalama
        xl_src = self.evrak_dir / "tablo_src.xlsx"
        xl_dst = self.evrak_dir / "tablo_imzali.xlsx"
        _make_sample_xlsx(xl_src)
        df_xl = dstudio._load_xlsx_sheet_df(xl_src, "TaramaListesi")
        df_xl.loc[0, "Durum"] = "İmzalandı"
        dstudio._save_xlsx_sheet_df(
            xl_src,
            xl_dst,
            "TaramaListesi",
            df_xl,
            font_family="Arial",
            font_size_pt=11.0,
            landscape=True,
            signature_overlays=[{"x0": 520, "y0": 420, "x1": 680, "y1": 490, "signature_b64": sig_b64}],
        )
        wb_out = openpyxl.load_workbook(xl_dst)
        ws_out = wb_out["TaramaListesi"]
        self.assertEqual(str(ws_out.page_setup.paperSize), str(ws_out.PAPERSIZE_A4))
        self.assertEqual(ws_out.page_setup.orientation, ws_out.ORIENTATION_LANDSCAPE)
        self.assertGreaterEqual(len(getattr(ws_out, "_images", [])), 1)
        wb_out.close()

        # 4. Görsel (.png) üzerine A4 metin + şeffaf PNG imza işleme
        img_src = self.evrak_dir / "belge_img.png"
        img_dst = self.evrak_dir / "belge_img_imzali.png"
        _make_sample_png(img_src)
        dstudio.save_image_with_overlays(
            img_src,
            [
                {
                    "field_type": "text",
                    "value": "Revir Hemşiresi Onayı",
                    "x0": 60,
                    "y0": 420,
                    "x1": 340,
                    "y1": 460,
                    "_ref_w": 595.0,
                    "_ref_h": 825.0,
                    "font_family": "Arial",
                    "font_size": 14,
                    "bold": True,
                    "color": "lacivert",
                },
                {
                    "field_type": "signature",
                    "x0": 320,
                    "y0": 560,
                    "x1": 480,
                    "y1": 620,
                    "_ref_w": 595.0,
                    "_ref_h": 825.0,
                    "signature_b64": sig_b64,
                },
            ],
            img_dst,
        )
        self.assertTrue(img_dst.exists())
        im_out = Image.open(img_dst)
        self.assertEqual(im_out.size, (620, 860))
        im_out.close()

    def test_depot_side_by_side_actions_preview_and_a4_studios(self):
        pdf_p = self.evrak_dir / "a4_form.pdf"
        docx_p = self.evrak_dir / "a4_sozlesme.docx"
        xlsx_p = self.evrak_dir / "a4_tablo.xlsx"
        png_p = self.evrak_dir / "a4_resim.png"

        _make_sample_pdf(pdf_p)
        _make_sample_docx(docx_p)
        _make_sample_xlsx(xlsx_p)
        _make_sample_png(png_p)

        # 1. Evrak Deposu Kart Izgarası: tek satırda yan yana 5 eylem butonu & okunaklı kapaklar
        revir_ui.reset_session()
        out_depot = revir_ui.run_action_headless(
            lambda: persantil.render_documents_depot(),
            {"action": "render"},
        )
        html_depot = out_depot["main_html"]
        self.assertNotIn("Çalışma Zamanı Hatası", html_depot)
        self.assertIn("revir-doc-card-thumb", html_depot)
        self.assertIn("revir-doc-action-row", html_depot)
        self.assertIn("btn_studio_card_", html_depot)
        self.assertIn("btn_prev_", html_depot)
        self.assertIn("__revirDownload", html_depot)

        # 2. Detaylı Liste görünümü: aynı şekilde yan yana eylem çubuğu
        revir_ui.session_state["depot_layout_choice"] = "Detaylı Liste"
        out_list = revir_ui.run_action_headless(
            lambda: persantil.render_documents_depot(),
            {"action": "render"},
        )
        self.assertIn("btn_edit_list_", out_list["main_html"])
        self.assertIn("btn_prev_list_", out_list["main_html"])
        self.assertIn("revir-doc-action-row", out_list["main_html"])

        # 3. Belge Önizleme Penceresi (kart üzerindeki 👁️ butonuna tıklayarak): Düzenle, İndir, Ad Değiştir, Sil üstte tek yatay barda
        from hashlib import sha256 as _sha256

        pdf_h = _sha256(pdf_p.name.encode()).hexdigest()[:12]
        revir_ui.reset_session()
        out_prev = revir_ui.run_action_headless(
            lambda: persantil.render_documents_depot(),
            {"action": "click", "key": f"btn_prev_{pdf_h}"},
        )
        prev_html = out_prev.get("dialog_html") or ""
        self.assertIn("revir-doc-prev-actions", prev_html)
        self.assertIn("prev_open_studio_", prev_html)
        self.assertIn("__revirDownload", prev_html)
        self.assertIn("Ad Değiştir", prev_html)
        self.assertIn("Sil", prev_html)

        # 4. Canlı A4 Belgeyi Düzenle Stüdyosu tüm düzenlenebilir biçimler için hatasız açılmalı
        for target_doc, expected_token in (
            (pdf_p, "a4_stage_pdf_"),
            (docx_p, "a4_stage_docx_"),
            (xlsx_p, "TaramaListesi"),
            (png_p, "a4_stage_img_"),
        ):
            doc_h = _sha256(target_doc.name.encode()).hexdigest()[:12]
            revir_ui.reset_session()
            out_st = revir_ui.run_action_headless(
                lambda: persantil.render_documents_depot(),
                {"action": "click", "key": f"btn_studio_card_{doc_h}"},
            )
            dlg_html = out_st.get("dialog_html") or ""
            self.assertNotIn("Çalışma Zamanı Hatası", dlg_html, f"Stüdyo hata verdi: {target_doc.name}")
            self.assertIn(expected_token, dlg_html, f"Beklenen öğe yok ({target_doc.name}): {expected_token}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
