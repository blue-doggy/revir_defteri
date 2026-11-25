from __future__ import annotations

import copy
import html
import io
import json
import mimetypes
import re
import sqlite3
from datetime import date, datetime
from difflib import SequenceMatcher
from hashlib import sha256
from pathlib import Path
from typing import Any

import pandas as pd

import gemini_assistant as gai
import insulin_diabetes as idm
import revir_ui as ui

# ==============================================================================
# 1. YARDIMCI ARAÇLAR: TÜRKÇE KARAKTER FOLDING & UNICODE FONT BULUCU
# ==============================================================================

_TR_TRANS = str.maketrans("çğıöşü", "cgiosu")


def tr_fold(text: Any) -> str:
	s = re.sub(r"\s+", " ", str(text or "").replace("\xa0", " ")).strip()
	s = s.replace("İ", "i").replace("I", "ı").lower().replace("\u0307", "")
	return s.translate(_TR_TRANS)


def find_unicode_ttf_font(family: str = "arial", *, bold: bool = False, italic: bool = False) -> str | None:
	"""macOS, Windows ve Linux'ta Türkçe (ğşçıöüİ) destekli TTF/OTF font dosyasını bulur."""
	fam = tr_fold(family or "arial")
	if "times" in fam or "serif" in fam:
		if bold and italic:
			candidates = [
				"/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf",
				"C:/Windows/Fonts/timesbi.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf",
			]
		elif bold:
			candidates = [
				"/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
				"C:/Windows/Fonts/timesbd.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
			]
		elif italic:
			candidates = [
				"/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
				"C:/Windows/Fonts/timesi.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf",
			]
		else:
			candidates = [
				"/System/Library/Fonts/Supplemental/Times New Roman.ttf",
				"/System/Library/Fonts/Times.ttc",
				"C:/Windows/Fonts/times.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
				"/usr/share/fonts/truetype/freefont/FreeSerif.ttf",
			]
	elif "courier" in fam or "mono" in fam:
		if bold:
			candidates = [
				"/System/Library/Fonts/Supplemental/Courier New Bold.ttf",
				"C:/Windows/Fonts/courbd.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
			]
		else:
			candidates = [
				"/System/Library/Fonts/Supplemental/Courier New.ttf",
				"C:/Windows/Fonts/cour.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
			]
	elif "calibri" in fam or "tahoma" in fam or "verdana" in fam:
		if bold:
			candidates = [
				"/System/Library/Fonts/Supplemental/Verdana Bold.ttf",
				"/System/Library/Fonts/Supplemental/Tahoma Bold.ttf",
				"/System/Library/Fonts/Supplemental/Arial Bold.ttf",
				"C:/Windows/Fonts/calibrib.ttf",
				"C:/Windows/Fonts/tahomabd.ttf",
				"C:/Windows/Fonts/segoeuib.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
			]
		else:
			candidates = [
				"/System/Library/Fonts/Supplemental/Verdana.ttf",
				"/System/Library/Fonts/Supplemental/Tahoma.ttf",
				"C:/Windows/Fonts/calibri.ttf",
				"C:/Windows/Fonts/segoeui.ttf",
				"C:/Windows/Fonts/tahoma.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
			]
	else:
		if bold and italic:
			candidates = [
				"/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf",
				"C:/Windows/Fonts/arialbi.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
			]
		elif bold:
			candidates = [
				"/System/Library/Fonts/Supplemental/Arial Bold.ttf",
				"/System/Library/Fonts/Helvetica.ttc",
				"C:/Windows/Fonts/arialbd.ttf",
				"C:/Windows/Fonts/segoeuib.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
			]
		elif italic:
			candidates = [
				"/System/Library/Fonts/Supplemental/Arial Italic.ttf",
				"C:/Windows/Fonts/ariali.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
			]
		else:
			candidates = [
				"/System/Library/Fonts/Supplemental/Arial.ttf",
				"/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
				"/System/Library/Fonts/Supplemental/Tahoma.ttf",
				"/System/Library/Fonts/Supplemental/Verdana.ttf",
				"/Library/Fonts/Arial.ttf",
				"C:/Windows/Fonts/arial.ttf",
				"C:/Windows/Fonts/calibri.ttf",
				"C:/Windows/Fonts/segoeui.ttf",
				"C:/Windows/Fonts/tahoma.ttf",
				"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
				"/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
				"/usr/share/fonts/truetype/freefont/FreeSans.ttf",
			]

	# Genel güvenli yedek zincir (kalın/italik bulunamazsa düz Unicode TTF'ye düş)
	candidates.extend([
		"/System/Library/Fonts/Supplemental/Arial.ttf",
		"/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
		"C:/Windows/Fonts/arial.ttf",
		"C:/Windows/Fonts/calibri.ttf",
		"C:/Windows/Fonts/segoeui.ttf",
		"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
	])
	for p in candidates:
		try:
			if Path(p).exists():
				return p
		except Exception:
			continue
	return None


def get_cover_pil_font(size: int = 14, bold: bool = False, family: str = "arial", italic: bool = False):
	from PIL import ImageFont

	p = find_unicode_ttf_font(family=family, bold=bold, italic=italic)
	if p:
		try:
			return ImageFont.truetype(p, max(8, int(size)))
		except Exception:
			pass
	return ImageFont.load_default()


def _safe_stem(name: str, fallback: str = "belge") -> str:
	cleaned = re.sub(r'[\\/*?:"<>|]', "_", (name or "").strip())
	return cleaned or fallback


def _unique_target_path(folder: Path, stem: str, ext: str) -> Path:
	base = _safe_stem(stem)
	cand = folder / f"{base}{ext}"
	idx = 1
	while cand.exists():
		cand = folder / f"{base}_{idx}{ext}"
		idx += 1
	return cand


# ==============================================================================
# 2. AKILLI PDF YAZILABİLİR ALAN TARAYICI, ÖNİZLEYİCİ VE TÜRKÇE YAZICI (PyMuPDF)
# ==============================================================================

def _rects_overlap_strongly(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
	ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
	ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
	if ix1 <= ix0 or iy1 <= iy0:
		return False
	inter = (ix1 - ix0) * (iy1 - iy0)
	area_a = max(1.0, (a[2] - a[0]) * (a[3] - a[1]))
	area_b = max(1.0, (b[2] - b[0]) * (b[3] - b[1]))
	return (inter / min(area_a, area_b)) > 0.45


def scan_pdf_writable_fields(pdf_path: Path, page_idx: int = 0) -> list[dict[str, Any]]:
	"""
	PDF sayfasındaki yazı yazılabilecek alanları %100 çevrimdışı (yerel) tarar:
	1) Gerçek AcroForm form alanları (metin kutusu, onay kutusu, seçim listesi)
	2) Statik/düz PDF'lerde etiket sonrası boşluklar ('Adı Soyadı:', 'Tarih:', 'Okulu:' vb.)
	3) Alt çizgili / noktalı boşluklar ('_____', '.....') ve yatay vektör doldurma çizgileri.
	"""
	import pymupdf

	fields: list[dict[str, Any]] = []
	doc = pymupdf.open(pdf_path)
	try:
		if page_idx < 0 or page_idx >= len(doc):
			return []
		page = doc[page_idx]
		pw = float(page.rect.width or 595.0)
		ph = float(page.rect.height or 842.0)

		# 1. AcroForm interaktif alanları
		try:
			widgets = list(page.widgets() or [])
		except Exception:
			widgets = []
		for wi, w in enumerate(widgets):
			r = w.rect
			if r.width < 6 or r.height < 6:
				continue
			ft_num = int(getattr(w, "field_type", 7) or 7)
			ftype = "checkbox" if ft_num == 2 else "text"
			lbl = str(getattr(w, "field_label", "") or getattr(w, "field_name", "") or f"Form Alanı {wi + 1}").strip()
			val = str(getattr(w, "field_value", "") or "").strip()
			if val.lower() in ("off", "none"):
				val = ""
			fields.append({
				"id": f"af_{page_idx}_{wi}",
				"page": page_idx,
				"label": lbl[:60],
				"field_type": ftype,
				"source": "acroform",
				"widget_name": str(getattr(w, "field_name", "") or ""),
				"x0": round(float(r.x0), 1),
				"y0": round(float(r.y0), 1),
				"x1": round(float(r.x1), 1),
				"y1": round(float(r.y1), 1),
				"value": val,
				"font_size": 10,
				"color": "lacivert",
				"clear_bg": False,
			})

		# 2. Metin bloklarından etiket+boşluk ve altçizgi/nokta taraması
		txt_dict = page.get_text("dict") or {}
		lines_info: list[dict[str, Any]] = []
		for blk in txt_dict.get("blocks") or []:
			if blk.get("type") != 0:
				continue
			for ln in blk.get("lines") or []:
				spans = ln.get("spans") or []
				if not spans:
					continue
				full_txt = "".join(str(sp.get("text") or "") for sp in spans)
				clean_s = full_txt.strip()
				if not clean_s:
					continue
				lb = ln.get("bbox") or spans[0].get("bbox")
				lines_info.append({
					"text": clean_s,
					"raw": full_txt,
					"bbox": (float(lb[0]), float(lb[1]), float(lb[2]), float(lb[3])),
					"spans": spans,
				})

		def _has_text_in_rect(rx0: float, ry0: float, rx1: float, ry1: float) -> bool:
			for li in lines_info:
				bx0, by0, bx1, by1 = li["bbox"]
				if bx1 > rx0 + 3 and bx0 < rx1 - 3 and by1 > ry0 + 2 and by0 < ry1 - 2:
					# Sadece nokta/altçizgi ise boş kabul et
					stripped = re.sub(r"[._\-–—\s:]", "", li["text"])
					if stripped:
						return True
			return False

		def _add_candidate(lbl: str, rect: tuple[float, float, float, float], src: str = "detected", clear_bg: bool = False) -> None:
			x0, y0, x1, y1 = rect
			x0 = max(12.0, min(pw - 25.0, x0))
			x1 = max(x0 + 24.0, min(pw - 14.0, x1))
			y0 = max(10.0, min(ph - 18.0, y0))
			y1 = max(y0 + 14.0, min(ph - 10.0, y1))
			cand_r = (x0, y0, x1, y1)
			for ex in fields:
				if _rects_overlap_strongly(cand_r, (ex["x0"], ex["y0"], ex["x1"], ex["y1"])):
					return
			clean_lbl = re.sub(r"\s+", " ", str(lbl or "").replace("\xa0", " ")).strip(" :._-")
			fields.append({
				"id": f"det_{page_idx}_{len(fields) + 1}",
				"page": page_idx,
				"label": (clean_lbl or f"Yazı Alanı {len(fields) + 1}")[:55],
				"field_type": "text",
				"source": src,
				"widget_name": "",
				"x0": round(x0, 1),
				"y0": round(y0, 1),
				"x1": round(x1, 1),
				"y1": round(y1, 1),
				"value": "",
				"font_size": 10,
				"color": "lacivert",
				"clear_bg": clear_bg,
			})

		# 2a. Altçizgi / nokta dizisi (____ veya ....) veya ':' biten form etiketleri
		for li in lines_info:
			txt = li["text"]
			bx0, by0, bx1, by1 = li["bbox"]
			lh = max(14.0, by1 - by0 + 4.0)

			# Örn: "Adı Soyadı : ............" veya "Tarih: ________"
			m_dots = re.search(r"([._]{4,})", txt)
			if m_dots:
				prefix = txt[:m_dots.start()].strip(" :.-_")
				ratio_start = m_dots.start() / max(1, len(txt))
				ratio_end = m_dots.end() / max(1, len(txt))
				dx0 = bx0 + (bx1 - bx0) * ratio_start
				dx1 = max(dx0 + 60.0, bx0 + (bx1 - bx0) * ratio_end)
				_add_candidate(prefix or "Boş Çizgi Alanı", (dx0, by0 - 2.0, dx1, by0 + lh), src="underline", clear_bg=True)
				continue

			# Örn: "Adı Soyadı:" (sağında boşluk var)
			if txt.endswith(":") and len(txt) <= 44 and bx1 < pw * 0.80:
				# Aynı satırda sağdaki bir sonraki metnin başlangıcını bul
				next_x = pw - 28.0
				for other in lines_info:
					if other is li:
						continue
					ox0, oy0, ox1, oy1 = other["bbox"]
					if abs(oy0 - by0) <= 7.0 and ox0 > bx1 + 4.0:
						next_x = min(next_x, ox0 - 6.0)
				if next_x - (bx1 + 6.0) >= 48.0 and not _has_text_in_rect(bx1 + 5.0, by0 - 1.0, min(next_x, bx1 + 110.0), by1 + 1.0):
					_add_candidate(txt[:-1], (bx1 + 6.0, by0 - 2.0, next_x, by0 + lh), src="label_colon", clear_bg=False)

		# 2b. Vektörel yatay doldurma çizgileri ve boş form kutuları (page.get_drawings)
		try:
			drawings = page.get_drawings() or []
		except Exception:
			drawings = []
		for dr in drawings:
			dr_r = dr.get("rect")
			if not dr_r:
				continue
			dx0, dy0, dx1, dy1 = float(dr_r.x0), float(dr_r.y0), float(dr_r.x1), float(dr_r.y1)
			dw, dh = dx1 - dx0, dy1 - dy0
			# Yatay doldurma çizgisi (genişlik >= 55pt, yükseklik <= 3pt, sayfa genişliğinin %75'inden küçük ki başlık ayırıcısı olmasın)
			if 55.0 <= dw <= pw * 0.72 and dh <= 3.0 and 24.0 < dy0 < ph - 24.0:
				box_y0 = max(10.0, dy0 - 16.0)
				box_y1 = dy0 + 2.0
				if not _has_text_in_rect(dx0 + 2.0, box_y0 + 1.0, dx1 - 2.0, box_y1 - 1.0):
					# Solundaki veya hemen altındaki etiketi bul
					best_lbl = ""
					best_dist = 9999.0
					for li in lines_info:
						lx0, ly0, lx1, ly1 = li["bbox"]
						# Solunda aynı satırda
						if abs(ly1 - dy0) <= 10.0 and lx1 <= dx0 + 8.0 and (dx0 - lx1) < 90.0:
							d = dx0 - lx1
							if d < best_dist:
								best_dist = d
								best_lbl = li["text"]
						# Veya hemen altında ortalanmış (örn: İmza / Adı Soyadı)
						elif 0.0 <= (ly0 - dy0) <= 16.0 and lx0 >= dx0 - 25.0 and lx1 <= dx1 + 25.0:
							d = 120.0 + (ly0 - dy0)
							if d < best_dist:
								best_dist = d
								best_lbl = li["text"]
					_add_candidate(best_lbl or "Çizgi Üstü Alan", (dx0 + 2.0, box_y0, dx1 - 2.0, box_y1), src="line", clear_bg=False)
			# Küçük boş dikdörtgen form kutusu (ör. tablo boş hücresi / veri kutusu)
			elif 45.0 <= dw <= pw * 0.70 and 14.0 <= dh <= 34.0 and 24.0 < dy0 < ph - 24.0:
				if not _has_text_in_rect(dx0 + 3.0, dy0 + 2.0, dx1 - 3.0, dy1 - 2.0):
					best_lbl = ""
					for li in lines_info:
						lx0, ly0, lx1, ly1 = li["bbox"]
						if abs(((ly0 + ly1) / 2.0) - ((dy0 + dy1) / 2.0)) <= 12.0 and lx1 <= dx0 + 10.0 and (dx0 - lx1) < 120.0:
							best_lbl = li["text"]
							break
					if best_lbl:
						_add_candidate(best_lbl, (dx0 + 3.0, dy0 + 2.0, dx1 - 3.0, dy1 - 2.0), src="box", clear_bg=False)

		# Yukarıdan aşağıya, soldan sağa sırala (max 45 alan)
		fields.sort(key=lambda f: (round(f["y0"] / 8.0), f["x0"]))
		return fields[:45]
	finally:
		doc.close()


_PDF_COLOR_MAP = {
	"lacivert": (0.06, 0.19, 0.45),
	"siyah": (0.08, 0.08, 0.08),
	"kirmizi": (0.72, 0.11, 0.11),
	"yesil": (0.04, 0.44, 0.36),
}

_PIL_COLOR_MAP = {
	"lacivert": (15, 48, 115, 255),
	"siyah": (20, 20, 20, 255),
	"kirmizi": (184, 28, 28, 255),
	"yesil": (10, 112, 92, 255),
}

FONT_FAMILY_OPTIONS = ["Arial (Sans-Serif)", "Times New Roman (Serif)", "Calibri / Modern", "Courier New (Daktilo)"]
ALIGN_OPTIONS = ["Sola Hizalı", "Ortala", "Sağa Hizalı"]


def _font_family_code(family_lbl: str) -> str:
	fl = tr_fold(family_lbl)
	if "times" in fl or "serif" in fl:
		return "times"
	if "courier" in fl or "daktilo" in fl or "mono" in fl:
		return "courier"
	if "calibri" in fl or "modern" in fl or "verdana" in fl or "tahoma" in fl:
		return "calibri"
	return "arial"


def _css_font_family(family_lbl: str) -> str:
	code = _font_family_code(family_lbl)
	return {
		"times": '"Times New Roman", Times, Georgia, serif',
		"courier": '"Courier New", Courier, monospace',
		"calibri": 'Calibri, "Segoe UI", Tahoma, sans-serif',
		"arial": 'Arial, Helvetica, "Liberation Sans", sans-serif',
	}[code]


def _align_to_int(align_lbl: str) -> int:
	al = tr_fold(align_lbl)
	if "orta" in al or "center" in al:
		return 1
	if "sag" in al or "right" in al:
		return 2
	return 0


def _resolve_sig_bytes(f: dict[str, Any]) -> bytes | None:
	raw_b64 = str(f.get("signature_b64") or "").strip()
	if raw_b64:
		try:
			import base64 as _b64
			return _b64.b64decode(raw_b64)
		except Exception:
			pass
	sig_p_str = str(f.get("signature_path") or "").strip()
	if sig_p_str:
		try:
			sp = Path(sig_p_str)
			if sp.exists() and sp.is_file():
				return sp.read_bytes()
		except Exception:
			pass
	return None


def _apply_fields_to_pymupdf_page(page: Any, page_fields: list[dict[str, Any]], font_path: str | None, draw_guides: bool = False) -> None:
	import pymupdf

	font_cache: dict[str, str] = {}

	def _ensure_font_for(f_spec: dict[str, Any]) -> str:
		fam = _font_family_code(str(f_spec.get("font_family") or "arial"))
		is_b = bool(f_spec.get("bold"))
		is_i = bool(f_spec.get("italic"))
		cache_k = f"{fam}_{int(is_b)}_{int(is_i)}"
		if cache_k in font_cache:
			return font_cache[cache_k]
		fp = find_unicode_ttf_font(family=fam, bold=is_b, italic=is_i) or font_path
		fallback_map = {"times": "tiro", "courier": "cour", "calibri": "helv", "arial": "helv"}
		fn = fallback_map.get(fam, "helv")
		if fp:
			try:
				reg_name = f"revir_{cache_k}"
				page.insert_font(fontname=reg_name, fontfile=fp)
				fn = reg_name
			except Exception:
				pass
		font_cache[cache_k] = fn
		return fn

	acro_map: dict[str, Any] = {}
	try:
		for w in (page.widgets() or []):
			wname = str(getattr(w, "field_name", "") or "")
			if wname:
				acro_map[wname] = w
	except Exception:
		pass

	for idx, f in enumerate(page_fields):
		x0, y0, x1, y1 = float(f["x0"]), float(f["y0"]), float(f["x1"]), float(f["y1"])
		rect = pymupdf.Rect(x0, y0, x1, y1)
		ftype = str(f.get("field_type") or "text")
		val = str(f.get("value") or "")
		fsize = float(f.get("font_size") or 10.0)
		col_rgb = _PDF_COLOR_MAP.get(str(f.get("color") or "lacivert"), _PDF_COLOR_MAP["lacivert"])
		align_int = _align_to_int(str(f.get("align") or "left"))

		if ftype == "signature":
			sig_bytes = _resolve_sig_bytes(f)
			if draw_guides:
				shape = page.new_shape()
				shape.draw_rect(rect)
				shape.finish(color=(0.49, 0.23, 0.93), fill=(0.96, 0.95, 1.0), fill_opacity=0.28, width=1.1, dashes="[4 2]")
				badge_r = pymupdf.Rect(max(2.0, x0 - 1.0), max(2.0, y0 - 12.0), min(page.rect.width - 2.0, x0 + 48.0), max(12.0, y0 - 0.5))
				shape.draw_rect(badge_r)
				shape.finish(color=(0.49, 0.23, 0.93), fill=(0.49, 0.23, 0.93), width=0.5)
				shape.commit()
				page.insert_text((badge_r.x0 + 2.5, badge_r.y1 - 2.5), f"#{idx + 1} IMZA", fontsize=7.0, fontname="helv", color=(1, 1, 1))
			if sig_bytes:
				try:
					page.insert_image(rect, stream=sig_bytes, keep_proportion=True, overlay=True)
				except Exception:
					pass
			continue

		if draw_guides:
			stroke_c = (0.05, 0.58, 0.53) if val.strip() else (0.15, 0.39, 0.92)
			fill_c = (0.90, 0.99, 0.97) if val.strip() else (0.94, 0.97, 1.0)
			shape = page.new_shape()
			shape.draw_rect(rect)
			shape.finish(color=stroke_c, fill=fill_c, fill_opacity=0.42, width=1.0, dashes="[3 2]")
			badge_r = pymupdf.Rect(max(2.0, x0 - 1.0), max(2.0, y0 - 11.5), min(page.rect.width - 2.0, x0 + 21.0), max(12.0, y0 - 0.5))
			shape.draw_rect(badge_r)
			shape.finish(color=stroke_c, fill=stroke_c, width=0.5)
			shape.commit()
			page.insert_text((badge_r.x0 + 2.5, badge_r.y1 - 2.5), f"#{idx + 1}", fontsize=7.5, fontname="helv", color=(1, 1, 1))

		if not val.strip():
			continue

		# Gerçek AcroForm alanı ise (ve kayıt modundaysak) widget değerini güncelle
		if not draw_guides and f.get("source") == "acroform" and f.get("widget_name") in acro_map:
			w = acro_map[f["widget_name"]]
			try:
				if ftype == "checkbox":
					w.field_value = val.strip().lower() in ("1", "true", "evet", "x", "✓", "on", "yes")
				else:
					w.field_value = val
				w.update()
				continue
			except Exception:
				pass

		# Arka plan temizleme (noktalı/altçizgili alanları örtmek için)
		if f.get("clear_bg") and not draw_guides:
			sh_bg = page.new_shape()
			sh_bg.draw_rect(rect)
			sh_bg.finish(color=None, fill=(1, 1, 1))
			sh_bg.commit()

		fontname = _ensure_font_for(f)
		# Kutuya sığacak şekilde yazı boyutu ayarla ve Türkçe metni yerleştir
		cur_fs = fsize
		placed = False
		for _ in range(5):
			rc = page.insert_textbox(rect, val, fontsize=cur_fs, fontname=fontname, color=col_rgb, align=align_int)
			if rc >= 0:
				placed = True
				break
			cur_fs = max(6.5, cur_fs - 1.2)
		if not placed:
			page.insert_text((x0 + 2.0, min(y1 - 2.0, y0 + cur_fs + 1.0)), val[:140], fontsize=cur_fs, fontname=fontname, color=col_rgb)


def build_a4_html_shell(title: str, body_html: str, *, landscape: bool = False) -> str:
	"""Herhangi bir belge/tablo/metin içeriğini ekranda gerçek A4 kâğıdı (210×297 mm) ve yazdırmaya hazır HTML olarak sarar."""
	page_size = "A4 landscape" if landscape else "A4 portrait"
	sheet_w = "297mm" if landscape else "210mm"
	sheet_min_h = "210mm" if landscape else "297mm"
	pad = "12mm 14mm" if landscape else "15mm 18mm"
	return f"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
@page {{ size: {page_size}; margin: 12mm; }}
@media print {{
	html, body {{ background: #ffffff !important; margin: 0 !important; padding: 0 !important; }}
	.no-print {{ display: none !important; }}
	.a4-sheet {{ box-shadow: none !important; border: none !important; margin: 0 !important; width: 100% !important; min-height: auto !important; padding: 0 !important; }}
}}
* {{ box-sizing: border-box; }}
body {{
	margin: 0; padding: 18px 14px 28px 14px;
	background: #cbd5e1; color: #0f172a;
	font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
	font-size: 11pt; line-height: 1.5;
}}
.no-print {{
	max-width: {sheet_w}; margin: 0 auto 12px auto;
	display: flex; justify-content: space-between; align-items: center; gap: 8px;
}}
.a4-badge {{
	font-size: 12px; font-weight: 700; color: #334155;
	background: #ffffff; border: 1px solid #94a3b8; padding: 4px 12px; border-radius: 999px;
}}
.print-btn {{
	background: #0f766e; color: #ffffff; border: none; padding: 7px 15px;
	border-radius: 7px; font-weight: 700; font-size: 12.5px; cursor: pointer;
}}
.print-btn:hover {{ background: #115e59; }}
.a4-sheet {{
	width: min(100%, {sheet_w}); min-height: {sheet_min_h};
	margin: 0 auto 20px auto; padding: {pad};
	background: #ffffff; color: #0f172a;
	border: 1px solid #94a3b8; border-radius: 3px;
	box-shadow: 0 10px 30px rgba(15,23,42,0.22);
}}
h1 {{ font-size: 16pt; margin: 0 0 10px 0; color: #0f172a; border-bottom: 2px solid #0f766e; padding-bottom: 5px; }}
h2 {{ font-size: 13.5pt; margin: 14px 0 6px 0; color: #0f766e; }}
h3 {{ font-size: 11.5pt; margin: 10px 0 4px 0; color: #1e293b; }}
p {{ margin: 0 0 7px 0; white-space: pre-wrap; }}
table {{ width: 100%; border-collapse: collapse; margin: 10px 0 16px 0; font-size: 9.5pt; }}
th, td {{ border: 1px solid #cbd5e1; padding: 5px 7px; text-align: left; vertical-align: top; word-break: break-word; }}
th {{ background: #f1f5f9; font-weight: 700; color: #0f172a; }}
pre.a4-pre {{ white-space: pre-wrap; font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 9.5pt; margin: 0; }}
</style>
</head>
<body>
<div class="no-print">
	<span class="a4-badge">📄 A4 {"Yatay (297×210 mm)" if landscape else "Dikey (210×297 mm)"} Önizleme</span>
	<button class="print-btn" onclick="window.print()">🖨️ A4 Yazdır / PDF Kaydet</button>
</div>
<div class="a4-sheet">{body_html}</div>
</body>
</html>"""


def render_docx_a4_html(file_path: Path, paras_override: list[dict[str, Any]] | None = None, tables_override: list[pd.DataFrame] | None = None, appended_paras: list[str] | None = None) -> str:
	"""Word (.docx) belgesini gerçek A4 kâğıdı görünümünde HTML'e çevirir."""
	if paras_override is None or tables_override is None:
		paras, tables = _load_docx_structure(file_path)
	else:
		paras, tables = paras_override, tables_override
	parts: list[str] = []
	for p in paras:
		txt = str(p.get("text") or "").strip()
		if not txt:
			continue
		st = str(p.get("style") or "").lower()
		esc = html.escape(txt)
		if "heading 1" in st or "title" in st:
			parts.append(f"<h1>{esc}</h1>")
		elif "heading 2" in st:
			parts.append(f"<h2>{esc}</h2>")
		elif "heading 3" in st:
			parts.append(f"<h3>{esc}</h3>")
		else:
			parts.append(f"<p>{esc}</p>")
	for ap in (appended_paras or []):
		if str(ap).strip():
			parts.append(f"<p>{html.escape(str(ap).strip())}</p>")
	for df in tables:
		if df is None or df.empty:
			continue
		ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
		trs = []
		for _, row in df.head(120).iterrows():
			tds = "".join(f"<td>{html.escape('' if pd.isna(v) else str(v))}</td>" for v in row.tolist())
			trs.append(f"<tr>{tds}</tr>")
		parts.append(f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>")
	if not parts:
		parts.append("<p style='color:#64748b;'>Belge boş.</p>")
	return build_a4_html_shell(file_path.name, "\n".join(parts), landscape=False)


def render_xlsx_a4_html(file_path: Path, sheet_name: str, df_override: pd.DataFrame | None = None, landscape: bool = True) -> str:
	"""Excel (.xlsx) sayfasını A4 Yatay/Dikey kâğıt önizlemesine çevirir."""
	df = df_override if df_override is not None else _load_xlsx_sheet_df(file_path, sheet_name)
	hdr = f"<h1>{html.escape(file_path.stem)} — {html.escape(sheet_name)}</h1>"
	if df is None or df.empty:
		body = hdr + "<p style='color:#64748b;'>Sayfa boş.</p>"
	else:
		ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
		trs = []
		for _, row in df.head(150).iterrows():
			tds = "".join(f"<td>{html.escape('' if pd.isna(v) else str(v))}</td>" for v in row.tolist())
			trs.append(f"<tr>{tds}</tr>")
		body = hdr + f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>"
	return build_a4_html_shell(f"{file_path.name} ({sheet_name})", body, landscape=landscape)


def render_text_a4_html(file_path: Path, txt_content: str) -> str:
	body = f"<h1>{html.escape(file_path.name)}</h1><pre class='a4-pre'>{html.escape(txt_content)}</pre>"
	return build_a4_html_shell(file_path.name, body, landscape=False)


def render_pdf_page_preview_png(pdf_path: Path, page_idx: int, page_fields: list[dict[str, Any]], show_guides: bool = True, dpi: int = 135) -> bytes:
	import pymupdf

	font_path = find_unicode_ttf_font()
	doc = pymupdf.open(pdf_path)
	try:
		if page_idx < 0 or page_idx >= len(doc):
			return b""
		page = doc[page_idx]
		_apply_fields_to_pymupdf_page(page, page_fields, font_path, draw_guides=show_guides)
		pix = page.get_pixmap(dpi=dpi)
		return pix.tobytes("png")
	finally:
		doc.close()


def save_filled_pdf(src_pdf_path: Path, fields_by_page: dict[int, list[dict[str, Any]]], dest_pdf_path: Path) -> None:
	import pymupdf

	font_path = find_unicode_ttf_font()
	raw_bytes = src_pdf_path.read_bytes()
	doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
	try:
		for p_idx, p_fields in fields_by_page.items():
			if 0 <= int(p_idx) < len(doc) and p_fields:
				_apply_fields_to_pymupdf_page(doc[int(p_idx)], p_fields, font_path, draw_guides=False)
		dest_pdf_path.write_bytes(doc.tobytes(garbage=3, deflate=True))
	finally:
		doc.close()


def auto_match_pdf_fields_locally(
	fields: list[dict[str, Any]],
	app_cfg: dict[str, str],
	person: dict[str, Any] | None = None,
	hp: dict[str, Any] | None = None,
) -> int:
	"""
	%100 yerel ve çevrimdışı olarak PDF alan etiketlerini Okul Ayarları ve aktif kişinin
	gerçek verileriyle eşleştirip boş olan alanları doldurur. Doldurulan alan sayısını döndürür.
	"""
	p = person or {}
	h = hp or {}
	today_tr = date.today().strftime("%d.%m.%Y")
	dob_raw = str(p.get("birth_date") or "")
	dob_tr = dob_raw
	try:
		if len(dob_raw) >= 10:
			dob_tr = date.fromisoformat(dob_raw[:10]).strftime("%d.%m.%Y")
	except Exception:
		pass

	rules: list[tuple[tuple[str, ...], str]] = [
		(("okul adi", "kurum adi", "okulu", "kurumu"), str(app_cfg.get("school_name") or "")),
		(("hemsire", "revir sorumlusu", "gorevli"), str(app_cfg.get("nurse_name") or "")),
		(("mudur",), str(app_cfg.get("principal_name") or "")),
		(("ilce",), str(app_cfg.get("default_district") or "")),
		(("tarih", "date"), today_tr),
		(("adi soyadi", "ad soyad", "ogrenci adi", "hasta adi", "isim"), str(p.get("full_name") or "")),
		(("sinif", "sube"), str(p.get("class_name") or "")),
		(("okul no", "ogrenci no", "numara", "no"), str(p.get("student_no") or "")),
		(("dogum",), dob_tr),
		(("cinsiyet",), {"K": "Kız", "E": "Erkek"}.get(str(p.get("sex") or ""), "")),
		(("kan grubu",), str(h.get("blood_group") or "")),
		(("alerji",), str(h.get("allergies") or "")),
		(("kronik",), str(h.get("chronic_conditions") or "")),
		(("veli",), str(h.get("guardian_name") or p.get("guardian_name") or "")),
		(("telefon", "tel", "gsm", "iletisim"), str(h.get("guardian_phone") or p.get("guardian_phone") or "")),
		(("adres",), str(p.get("address") or "")),
	]

	filled = 0
	for f in fields:
		if str(f.get("value") or "").strip():
			continue
		lbl_f = tr_fold(f.get("label") or "")
		if not lbl_f:
			continue
		for keywords, val in rules:
			if not val.strip():
				continue
			if any(k in lbl_f for k in keywords):
				f["value"] = val.strip()
				filled += 1
				break
	return filled


# ==============================================================================
# 2B. EVRENSEL İLK SAYFA A4/EXCEL ÖNİZLEME, PNG İMZA VE CANLI A4 TUVAL MOTORU
# ==============================================================================

def _wrap_pil_lines(draw: Any, text: str, font: Any, max_width: int) -> list[str]:
	words = str(text or "").split()
	if not words:
		return [""]
	lines: list[str] = []
	cur = ""
	for w in words:
		cand = f"{cur} {w}".strip() if cur else w
		try:
			tw = draw.textlength(cand, font=font)
		except Exception:
			tw = len(cand) * 7.5
		if tw <= max_width or not cur:
			cur = cand
		else:
			lines.append(cur)
			cur = w
	if cur:
		lines.append(cur)
	return lines


def _excel_col_letter(idx: int) -> str:
	out = ""
	n = idx + 1
	while n > 0:
		n, rem = divmod(n - 1, 26)
		out = chr(65 + rem) + out
	return out


def _render_a4_table_sheet_image(
	*,
	title: str,
	headers: list[str],
	rows: list[list[str]],
	landscape: bool | None = None,
	accent_rgb: tuple[int, int, int] = (21, 128, 61),
	subtitle: str = "1. Sayfa Tablo Önizlemesi · A4",
) -> Any:
	from PIL import Image, ImageDraw

	ncols = max(1, min(8, len(headers) if headers else (max((len(r) for r in rows), default=4))))
	use_land = landscape if landscape is not None else (ncols >= 5)
	W, H = (920, 650) if use_land else (680, 962)

	img = Image.new("RGB", (W, H), "#ffffff")
	draw = ImageDraw.Draw(img)

	# Dış çerçeve & Excel/Tablo Üst Şerit
	draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
	draw.rectangle([(0, 0), (W, 44)], fill=accent_rgb)
	f_hdr = get_cover_pil_font(16, bold=True)
	f_sub = get_cover_pil_font(11, bold=False)
	f_title = get_cover_pil_font(15, bold=True)
	f_col_hdr = get_cover_pil_font(12, bold=True)
	f_cell = get_cover_pil_font(12, bold=False)
	f_cell_b = get_cover_pil_font(12, bold=True)

	draw.text((16, 12), f"📊 {subtitle}", fill="#ffffff", font=f_hdr)
	t_clean = str(title or "Tablo")
	if len(t_clean) > 58:
		t_clean = t_clean[:55] + "..."
	draw.text((18, 54), t_clean, fill="#0f172a", font=f_title)
	draw.line([(18, 80), (W - 18, 80)], fill="#cbd5e1", width=1)

	# Sütun genişliklerini içerik oranına göre hesapla
	margin_x = 18
	gutter_w = 34
	avail_w = W - (margin_x * 2) - gutter_w
	norm_hdrs = [(str(headers[i]) if i < len(headers) else _excel_col_letter(i)) for i in range(ncols)]
	norm_rows = [[(str(r[i]) if i < len(r) else "") for i in range(ncols)] for r in rows[:28]]

	weights = []
	for ci in range(ncols):
		max_len = max([len(norm_hdrs[ci])] + [len(r[ci]) for r in norm_rows[:18]] + [4])
		weights.append(max(6, min(28, max_len)))
	tot_w = max(1, sum(weights))
	col_widths = [max(52, int(avail_w * w / tot_w)) for w in weights]
	col_widths[-1] += avail_w - sum(col_widths)

	y_cur = 90
	# Harf başlık satırı (A, B, C...)
	draw.rectangle([(margin_x, y_cur), (margin_x + gutter_w, y_cur + 22)], fill="#e2e8f0", outline="#cbd5e1")
	x_c = margin_x + gutter_w
	for ci, cw in enumerate(col_widths):
		draw.rectangle([(x_c, y_cur), (x_c + cw, y_cur + 22)], fill="#f1f5f9", outline="#cbd5e1")
		lbl = _excel_col_letter(ci)
		tw = draw.textlength(lbl, font=f_col_hdr)
		draw.text((x_c + max(4, (cw - tw) / 2), y_cur + 4), lbl, fill="#475569", font=f_col_hdr)
		x_c += cw
	y_cur += 22

	def _fit_cell_text(txt: str, cw_px: int, font_obj: Any) -> str:
		s = str(txt or "").replace("\n", " ").strip()
		max_t = max(20, cw_px - 10)
		while s and draw.textlength(s, font=font_obj) > max_t:
			s = s[:-2] + "…"
			if len(s) <= 2:
				break
		return s

	# Tablo Başlık Satırı (Row 1)
	row_h = 26
	draw.rectangle([(margin_x, y_cur), (margin_x + gutter_w, y_cur + row_h)], fill="#e2e8f0", outline="#cbd5e1")
	draw.text((margin_x + 11, y_cur + 6), "1", fill="#475569", font=f_col_hdr)
	x_c = margin_x + gutter_w
	for ci, cw in enumerate(col_widths):
		draw.rectangle([(x_c, y_cur), (x_c + cw, y_cur + row_h)], fill="#dcfce7", outline="#86efac")
		ct = _fit_cell_text(norm_hdrs[ci], cw, f_cell_b)
		draw.text((x_c + 6, y_cur + 6), ct, fill="#14532d", font=f_cell_b)
		x_c += cw
	y_cur += row_h

	# Veri Satırları
	footer_y = H - 32
	for r_i, r_vals in enumerate(norm_rows):
		if y_cur + row_h > footer_y - 6:
			break
		draw.rectangle([(margin_x, y_cur), (margin_x + gutter_w, y_cur + row_h)], fill="#f1f5f9", outline="#cbd5e1")
		draw.text((margin_x + 8, y_cur + 6), str(r_i + 2), fill="#64748b", font=f_sub)
		row_bg = "#ffffff" if r_i % 2 == 0 else "#f8fafc"
		x_c = margin_x + gutter_w
		for ci, cw in enumerate(col_widths):
			draw.rectangle([(x_c, y_cur), (x_c + cw, y_cur + row_h)], fill=row_bg, outline="#e2e8f0")
			ct = _fit_cell_text(r_vals[ci], cw, f_cell)
			draw.text((x_c + 6, y_cur + 6), ct, fill="#1e293b", font=f_cell)
			x_c += cw
		y_cur += row_h

	draw.rectangle([(0, footer_y), (W, H)], fill="#f8fafc")
	draw.line([(0, footer_y), (W, footer_y)], fill="#e2e8f0", width=1)
	draw.text((16, footer_y + 8), f"📄 A4 Tablo Ekran Önizlemesi ({len(norm_rows)} satır · {ncols} sütun gösteriliyor)", fill="#475569", font=f_sub)
	return img


def render_html_body_to_pdf_bytes(
	body_html: str,
	*,
	font_family: str = "Arial",
	font_size_pt: float = 11.0,
	landscape: bool = False,
	margin_pt: float = 40.0,
	overlays_by_page: dict[int, list[dict[str, Any]]] | None = None,
) -> bytes:
	"""Verilen HTML gövdesini gerçek TTF Türkçe fontlarla vektör A4 PDF'e çevirir ve isteğe bağlı imza/metin damgaları basar."""
	import pymupdf

	fp = find_unicode_ttf_font(family=font_family)
	arch = pymupdf.Archive(str(Path(fp).parent)) if fp else None
	font_face = f'@font-face {{ font-family: "revirfont"; src: url("{Path(fp).name}"); }}\n' if fp else ""
	ff_css = '"revirfont", sans-serif' if fp else "sans-serif"
	css = (
		f"{font_face}"
		f"body {{ font-family: {ff_css}; font-size: {float(font_size_pt):.1f}pt; line-height: 1.45; color: #0f172a; }}\n"
		f"h1 {{ font-size: {min(24.0, float(font_size_pt) + 5.0):.1f}pt; color: #0f172a; border-bottom: 1.5px solid #0f766e; padding-bottom: 4px; margin: 0 0 8px 0; }}\n"
		f"h2 {{ font-size: {min(20.0, float(font_size_pt) + 2.5):.1f}pt; color: #0f766e; margin: 10px 0 5px 0; }}\n"
		f"h3 {{ font-size: {min(17.0, float(font_size_pt) + 1.0):.1f}pt; color: #1e293b; margin: 8px 0 4px 0; }}\n"
		"p { margin: 0 0 6px 0; }\n"
		"table { width: 100%; border-collapse: collapse; margin: 8px 0 12px 0; }\n"
		f"th, td {{ border: 0.75px solid #94a3b8; padding: 4px 6px; font-size: {max(7.5, float(font_size_pt) - 1.5):.1f}pt; vertical-align: top; }}\n"
		"th { background: #f1f5f9; font-weight: bold; }\n"
		"pre { white-space: pre-wrap; font-family: monospace; font-size: 9pt; }\n"
	)

	buf = io.BytesIO()
	writer = pymupdf.DocumentWriter(buf)
	story = pymupdf.Story(body_html or "<p></p>", user_css=css, archive=arch)
	mediabox = pymupdf.paper_rect("a4-l" if landscape else "a4")
	where = mediabox + (margin_pt, margin_pt, -margin_pt, -margin_pt)
	more = True
	guard = 0
	while more and guard < 40:
		dev = writer.begin_page(mediabox)
		more, _ = story.place(where)
		story.draw(dev, pymupdf.Identity)
		writer.end_page()
		guard += 1
	writer.close()

	pdf_bytes = buf.getvalue()
	if overlays_by_page:
		doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
		try:
			for p_idx, p_fields in overlays_by_page.items():
				if 0 <= int(p_idx) < len(doc) and p_fields:
					_apply_fields_to_pymupdf_page(doc[int(p_idx)], p_fields, fp, draw_guides=False)
			return doc.tobytes(garbage=3, deflate=True)
		finally:
			doc.close()
	return pdf_bytes


def _docx_to_inner_html(
	file_path: Path,
	paras_override: list[dict[str, Any]] | None = None,
	tables_override: list[pd.DataFrame] | None = None,
	appended_paras: list[str] | None = None,
	*,
	default_align: str = "",
) -> str:
	if paras_override is None or tables_override is None:
		paras, tables = _load_docx_structure(file_path)
	else:
		paras, tables = paras_override, tables_override

	align_css = {"Sola Hizalı": "left", "Ortala": "center", "Sağa Hizalı": "right"}.get(default_align, "")
	parts: list[str] = []
	for p in paras:
		txt = str(p.get("text") or "").strip()
		if not txt:
			continue
		st = str(p.get("style") or "").lower()
		p_align = str(p.get("align") or align_css or "")
		is_bold = bool(p.get("bold"))
		esc = html.escape(txt)
		st_attrs: list[str] = []
		if p_align:
			st_attrs.append(f"text-align:{p_align}")
		if is_bold:
			st_attrs.append("font-weight:bold")
		style_s = f' style="{";".join(st_attrs)}"' if st_attrs else ""
		if "heading 1" in st or "title" in st:
			parts.append(f"<h1{style_s}>{esc}</h1>")
		elif "heading 2" in st:
			parts.append(f"<h2{style_s}>{esc}</h2>")
		elif "heading 3" in st:
			parts.append(f"<h3{style_s}>{esc}</h3>")
		else:
			parts.append(f"<p{style_s}>{esc}</p>")
	for ap in (appended_paras or []):
		if str(ap).strip():
			parts.append(f"<p>{html.escape(str(ap).strip())}</p>")
	for df in tables:
		if df is None or df.empty:
			continue
		ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
		trs = []
		for _, row in df.head(120).iterrows():
			tds = "".join(f"<td>{html.escape('' if pd.isna(v) else str(v))}</td>" for v in row.tolist())
			trs.append(f"<tr>{tds}</tr>")
		parts.append(f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>")
	if not parts:
		parts.append("<p style='color:#64748b;'>Belge boş.</p>")
	return "\n".join(parts)


def _render_docx_first_page_snapshot(file_path: Path) -> Any:
	from PIL import Image, ImageDraw
	import zipfile

	# Öncelik: PyMuPDF Story ile gerçek vektör A4 sayfa 1 görseli üret (başlık + hizalama + tablolar birebir okunur)
	try:
		import pymupdf
		inner_html = _docx_to_inner_html(file_path)
		pdf_bytes = render_html_body_to_pdf_bytes(inner_html, font_family="Arial", font_size_pt=10.5, landscape=False, margin_pt=34.0)
		doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
		try:
			if len(doc) > 0:
				pix = doc[0].get_pixmap(dpi=115)
				im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
				# Köşeye şık A4 Word etiketi
				draw = ImageDraw.Draw(im)
				draw.rectangle([(0, 0), (im.width - 1, im.height - 1)], outline="#cbd5e1", width=2)
				return im
		finally:
			doc.close()
	except Exception:
		pass

	# Yedek: Doğrudan Pillow ile yüksek çözünürlüklü 680x962 A4 Sayfa 1 çizimi (küçük logo görseli metni asla ezmez)
	W, H = 680, 962
	img = Image.new("RGB", (W, H), "#ffffff")
	draw = ImageDraw.Draw(img)
	draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
	draw.rectangle([(0, 0), (W, 40)], fill="#0284c7")
	draw.text((16, 11), "📘 WORD BELGESİ · A4 SAYFA 1 ÖNİZLEMESİ", fill="#ffffff", font=get_cover_pil_font(14, bold=True))

	# Küçük logo varsa sadece üst sağ köşede zarif damga olarak göster
	try:
		with zipfile.ZipFile(file_path) as zf:
			for nm in zf.namelist():
				if nm.startswith("word/media/") and nm.lower().endswith((".png", ".jpg", ".jpeg")):
					logo_im = Image.open(io.BytesIO(zf.read(nm))).convert("RGB")
					logo_im.thumbnail((56, 56), Image.Resampling.LANCZOS)
					img.paste(logo_im, (W - 76, 50))
					break
	except Exception:
		pass

	f_h1 = get_cover_pil_font(17, bold=True)
	f_h2 = get_cover_pil_font(14, bold=True)
	f_body = get_cover_pil_font(13, bold=False)
	f_tbl = get_cover_pil_font(11, bold=False)
	f_foot = get_cover_pil_font(11, bold=False)

	y = 54
	try:
		paras, tables = _load_docx_structure(file_path)
	except Exception:
		paras, tables = ([{"style": "Heading 1", "text": file_path.stem}], [])

	for p in paras[:28]:
		if y > H - 90:
			break
		txt = str(p.get("text") or "").strip()
		if not txt:
			continue
		st = str(p.get("style") or "").lower()
		font_u = f_h1 if ("heading 1" in st or "title" in st) else (f_h2 if "heading 2" in st else f_body)
		col_u = "#0f766e" if "heading" in st else "#1e293b"
		step_y = 24 if font_u is f_h1 else (21 if font_u is f_h2 else 19)
		for ln in _wrap_pil_lines(draw, txt, font_u, W - 96)[:5]:
			if y > H - 65:
				break
			draw.text((28, y), ln, fill=col_u, font=font_u)
			y += step_y
		y += 4

	for df in tables[:2]:
		if df is None or df.empty or y > H - 110:
			break
		cols = list(df.columns)[:5]
		cw = max(80, (W - 56) // max(1, len(cols)))
		x_c = 28
		for ci, c_nm in enumerate(cols):
			draw.rectangle([(x_c, y), (x_c + cw, y + 24)], fill="#f1f5f9", outline="#94a3b8")
			draw.text((x_c + 5, y + 5), str(c_nm)[:16], fill="#0f172a", font=f_tbl)
			x_c += cw
		y += 24
		for _, row in df.head(12).iterrows():
			if y > H - 60:
				break
			x_c = 28
			for ci in range(len(cols)):
				draw.rectangle([(x_c, y), (x_c + cw, y + 22)], fill="#ffffff", outline="#cbd5e1")
				val = "" if pd.isna(row.iloc[ci]) else str(row.iloc[ci])
				draw.text((x_c + 5, y + 4), val[:16], fill="#334155", font=f_tbl)
				x_c += cw
			y += 22
		y += 10

	draw.rectangle([(0, H - 30), (W, H)], fill="#f8fafc")
	draw.line([(0, H - 30), (W, H - 30)], fill="#e2e8f0", width=1)
	draw.text((20, H - 22), "📄 Word (.docx) · A4 Sayfa 1 Önizlemesi", fill="#64748b", font=f_foot)
	return img


def _render_pptx_slide1_snapshot(file_path: Path) -> Any:
	from PIL import Image, ImageDraw
	import zipfile
	import xml.etree.ElementTree as ET

	texts: list[str] = []
	try:
		with zipfile.ZipFile(file_path) as zf:
			if "ppt/slides/slide1.xml" in zf.namelist():
				root = ET.fromstring(zf.read("ppt/slides/slide1.xml"))
				for elem in root.iter():
					if elem.tag.endswith("}t") and elem.text and elem.text.strip():
						texts.append(elem.text.strip())
	except Exception:
		pass

	W, H = 840, 520
	img = Image.new("RGB", (W, H), "#ffffff")
	draw = ImageDraw.Draw(img)
	draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
	draw.rectangle([(0, 0), (W, 46)], fill="#ea580c")
	draw.text((20, 13), "📙 POWERPOINT SUNUMU · SLAYT 1", fill="#ffffff", font=get_cover_pil_font(15, bold=True))

	title_s = texts[0] if texts else file_path.stem
	f_t = get_cover_pil_font(22, bold=True)
	f_b = get_cover_pil_font(15, bold=False)
	y = 74
	for ln in _wrap_pil_lines(draw, title_s, f_t, W - 64)[:3]:
		draw.text((32, y), ln, fill="#0f172a", font=f_t)
		y += 30
	draw.line([(32, y + 8), (W - 32, y + 8)], fill="#fed7aa", width=2)
	y += 24
	for bt in texts[1:14]:
		if y > H - 60:
			break
		for ln in _wrap_pil_lines(draw, f"• {bt}", f_b, W - 76)[:3]:
			if y > H - 56:
				break
			draw.text((40, y), ln, fill="#334155", font=f_b)
			y += 22
		y += 6

	draw.rectangle([(0, H - 30), (W, H)], fill="#fff7ed")
	draw.text((20, H - 22), f"📄 {file_path.name[:45]} · Slayt 1 Önizlemesi", fill="#9a3412", font=get_cover_pil_font(11))
	return img


def render_document_first_page_snapshot(file_path: Path, output_jpg_path: Path | None = None) -> bytes:
	"""Her evrak türü için okunabilir 1. sayfa (PDF/Word A4, Excel A4 tablo kesiti, Slayt 1, Metin, Görsel) kapağı üretir."""
	from PIL import Image, ImageDraw
	import pymupdf

	ext = file_path.suffix.lower()
	out_im: Any = None

	if ext == ".pdf":
		doc = pymupdf.open(file_path)
		try:
			if len(doc) > 0:
				page = doc[0]
				pix = page.get_pixmap(dpi=145)
				out_im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
				if max(out_im.width, out_im.height) > 960:
					out_im.thumbnail((960, 960), Image.Resampling.LANCZOS)
		finally:
			doc.close()

	elif ext in (".png", ".jpg", ".jpeg", ".webp"):
		src_im = Image.open(file_path).convert("RGB")
		use_land = src_im.width > src_im.height * 1.15
		W, H = (920, 650) if use_land else (680, 962)
		out_im = Image.new("RGB", (W, H), "#ffffff")
		draw = ImageDraw.Draw(out_im)
		draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
		src_im.thumbnail((W - 36, H - 56), Image.Resampling.LANCZOS)
		px = (W - src_im.width) // 2
		py = (H - 28 - src_im.height) // 2 + 4
		out_im.paste(src_im, (px, py))
		draw.rectangle([(0, H - 28), (W, H)], fill="#f8fafc")
		draw.text((16, H - 21), f"🖼️ A4 Görsel Önizleme · {file_path.name[:38]}", fill="#64748b", font=get_cover_pil_font(11))
		src_im.close()

	elif ext == ".docx":
		out_im = _render_docx_first_page_snapshot(file_path)

	elif ext in (".xlsx", ".xls", ".csv"):
		headers: list[str] = []
		rows: list[list[str]] = []
		sub = "Excel A4 Tablo (1. Sayfa)"
		if ext == ".csv":
			sub = "CSV Tablo Önizlemesi · A4"
			raw_txt = file_path.read_text(encoding="utf-8", errors="replace")
			lines = [ln for ln in raw_txt.splitlines() if ln.strip()][:32]
			sep = ";" if any(";" in ln for ln in lines[:5]) else ("," if any("," in ln for ln in lines[:5]) else "\t")
			parsed = [[c.strip() for c in ln.split(sep)] for ln in lines]
			if parsed:
				headers = parsed[0]
				rows = parsed[1:]
		else:
			import openpyxl
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			try:
				ws = wb.active
				sub = f"Excel · {ws.title[:22]} (A4 Tablo)"
				raw_rows = []
				for r in ws.iter_rows(values_only=True):
					vals = ["" if c is None else str(c).strip() for c in list(r)[:10]]
					while vals and vals[-1] == "":
						vals.pop()
					if any(vals):
						raw_rows.append(vals)
					if len(raw_rows) >= 30:
						break
				if raw_rows:
					max_c = max(len(x) for x in raw_rows)
					headers = raw_rows[0] + [""] * (max_c - len(raw_rows[0]))
					rows = [x + [""] * (max_c - len(x)) for x in raw_rows[1:]]
			finally:
				wb.close()
		out_im = _render_a4_table_sheet_image(
			title=file_path.stem,
			headers=headers or ["A", "B", "C", "D"],
			rows=rows,
			subtitle=sub,
		)

	elif ext == ".pptx":
		out_im = _render_pptx_slide1_snapshot(file_path)

	elif ext == ".txt":
		W, H = 680, 962
		out_im = Image.new("RGB", (W, H), "#ffffff")
		draw = ImageDraw.Draw(out_im)
		draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
		draw.rectangle([(0, 0), (W, 40)], fill="#334155")
		draw.text((16, 11), "📝 METİN BELGESİ · A4 SAYFA 1", fill="#ffffff", font=get_cover_pil_font(14, bold=True))
		txt = file_path.read_text(encoding="utf-8", errors="replace")[:5000]
		f_mono = get_cover_pil_font(13, bold=False, family="courier")
		y = 56
		for raw_ln in txt.splitlines():
			if y > H - 46:
				break
			wrapped = _wrap_pil_lines(draw, raw_ln, f_mono, W - 48) if raw_ln.strip() else [""]
			for wln in wrapped:
				if y > H - 46:
					break
				draw.text((24, y), wln, fill="#1e293b", font=f_mono)
				y += 19
		draw.rectangle([(0, H - 28), (W, H)], fill="#f8fafc")
		draw.text((16, H - 21), f"📄 {file_path.name[:42]} · 1. Sayfa", fill="#64748b", font=get_cover_pil_font(11))

	if out_im is None:
		W, H = 680, 962
		out_im = Image.new("RGB", (W, H), "#ffffff")
		draw = ImageDraw.Draw(out_im)
		draw.rectangle([(0, 0), (W - 1, H - 1)], outline="#cbd5e1", width=2)
		draw.rectangle([(0, 0), (W, 44)], fill="#475569")
		draw.text((16, 12), f"📄 {ext.upper().strip('.') or 'BELGE'} DOSYASI", fill="#ffffff", font=get_cover_pil_font(15, bold=True))
		f_big = get_cover_pil_font(18, bold=True)
		y = 90
		for ln in _wrap_pil_lines(draw, file_path.name, f_big, W - 64)[:5]:
			draw.text((32, y), ln, fill="#0f172a", font=f_big)
			y += 28

	out_buf = io.BytesIO()
	out_im.save(out_buf, format="JPEG", quality=88, optimize=True)
	jpg_bytes = out_buf.getvalue()
	out_im.close()
	if output_jpg_path is not None:
		output_jpg_path.parent.mkdir(parents=True, exist_ok=True)
		output_jpg_path.write_bytes(jpg_bytes)
	return jpg_bytes


def prepare_signature_png_bytes(
	raw_bytes: bytes,
	*,
	remove_white_bg: bool = True,
	max_width: int = 560,
) -> bytes:
	"""Yüklenen imza görüntüsünü şeffaf arka planlı ve kenar boşlukları kırpılmış PNG formatına çevirir."""
	from PIL import Image

	im = Image.open(io.BytesIO(raw_bytes)).convert("RGBA")
	if remove_white_bg:
		px = im.getdata()
		new_px = []
		for r, g, b, a in px:
			if a == 0:
				new_px.append((0, 0, 0, 0))
				continue
			bright = (r + g + b) / 3.0
			sat = max(r, g, b) - min(r, g, b)
			# Yakın beyaz kâğıt zeminini şeffaflaştır
			if bright > 230 and sat < 28:
				new_px.append((r, g, b, 0))
			elif bright > 205 and sat < 24:
				fade_a = int(a * max(0.0, (235.0 - bright) / 30.0))
				new_px.append((r, g, b, fade_a))
			else:
				new_px.append((r, g, b, a))
		im.putdata(new_px)

	# Şeffaf olmayan imza sınırlarına göre fazlalık boşluğu kırp
	alpha = im.getchannel("A")
	bbox = alpha.getbbox()
	if bbox:
		pad = 6
		crop_box = (
			max(0, bbox[0] - pad),
			max(0, bbox[1] - pad),
			min(im.width, bbox[2] + pad),
			min(im.height, bbox[3] + pad),
		)
		im = im.crop(crop_box)

	if im.width > max_width:
		ratio = max_width / float(im.width)
		im = im.resize((max_width, max(20, int(im.height * ratio))), Image.Resampling.LANCZOS)

	out = io.BytesIO()
	im.save(out, format="PNG")
	im.close()
	return out.getvalue()


def _signatures_dir(evrak_dir: Path) -> Path:
	sd = evrak_dir / ".signatures"
	sd.mkdir(parents=True, exist_ok=True)
	return sd


def list_saved_signatures(evrak_dir: Path) -> list[dict[str, Any]]:
	sd = _signatures_dir(evrak_dir)
	out: list[dict[str, Any]] = []
	for fp in sorted(sd.glob("*.png"), key=lambda x: x.stat().st_mtime, reverse=True):
		try:
			out.append({
				"name": fp.stem,
				"path": fp,
				"bytes": fp.read_bytes(),
			})
		except Exception:
			pass
	return out


def save_signature_to_library(evrak_dir: Path, label: str, raw_bytes: bytes, remove_white_bg: bool = True) -> Path:
	clean_png = prepare_signature_png_bytes(raw_bytes, remove_white_bg=remove_white_bg)
	stem = _safe_stem(label or f"imza_{datetime.now().strftime('%H%M%S')}", "imza")
	target = _signatures_dir(evrak_dir) / f"{stem}.png"
	target.write_bytes(clean_png)
	return target


def save_image_with_overlays(src_img_path: Path, overlays: list[dict[str, Any]], dest_path: Path) -> None:
	"""PNG/JPG/WEBP görseli veya A4 kâğıdı üzerine metin ve şeffaf PNG imza damgalarını kalıcı olarak işler."""
	from PIL import Image, ImageDraw

	base = Image.open(src_img_path).convert("RGBA")
	W, H = base.size
	draw = ImageDraw.Draw(base)

	for f in overlays:
		x0, y0, x1, y1 = float(f["x0"]), float(f["y0"]), float(f["x1"]), float(f["y1"])
		# Koordinatlar A4 referansı (pw/ph) veya doğrudan görsel pikseli üzerinden ölçeklenebilir
		ref_w = float(f.get("_ref_w") or W)
		ref_h = float(f.get("_ref_h") or H)
		px0 = int(round(x0 * W / max(1.0, ref_w)))
		py0 = int(round(y0 * H / max(1.0, ref_h)))
		px1 = int(round(x1 * W / max(1.0, ref_w)))
		py1 = int(round(y1 * H / max(1.0, ref_h)))
		box_w = max(16, px1 - px0)
		box_h = max(14, py1 - py0)

		if f.get("field_type") == "signature":
			sig_bytes = _resolve_sig_bytes(f)
			if sig_bytes:
				sig_im = Image.open(io.BytesIO(sig_bytes)).convert("RGBA")
				sig_im.thumbnail((box_w, box_h), Image.Resampling.LANCZOS)
				base.alpha_composite(sig_im, (max(0, px0), max(0, py0)))
				sig_im.close()
			continue

		val = str(f.get("value") or "")
		if not val.strip():
			continue
		if f.get("clear_bg"):
			draw.rectangle([(px0, py0), (px1, py1)], fill=(255, 255, 255, 255))
		fs_scaled = max(10, int(round(float(f.get("font_size") or 12) * (W / max(1.0, ref_w)))))
		pil_font = get_cover_pil_font(
			size=fs_scaled,
			bold=bool(f.get("bold")),
			family=_font_family_code(str(f.get("font_family") or "arial")),
			italic=bool(f.get("italic")),
		)
		col_rgba = _PIL_COLOR_MAP.get(str(f.get("color") or "lacivert"), _PIL_COLOR_MAP["lacivert"])
		y_t = py0 + 2
		for raw_ln in val.splitlines():
			for wln in _wrap_pil_lines(draw, raw_ln, pil_font, box_w):
				draw.text((px0 + 3, y_t), wln, fill=col_rgba, font=pil_font)
				y_t += int(fs_scaled * 1.25)

	ext = dest_path.suffix.lower()
	if ext in (".jpg", ".jpeg"):
		base.convert("RGB").save(dest_path, format="JPEG", quality=92)
	elif ext == ".webp":
		base.save(dest_path, format="WEBP", quality=92)
	else:
		base.save(dest_path, format="PNG")
	base.close()


def _sync_overlay_canvas_edits(fields: list[dict[str, Any]], raw_json: str) -> tuple[bool, int | None]:
	"""JS A4 sahnesinden gelen tıklama/sürükleme/boyutlandırma verilerini Python listesine uygular."""
	s = str(raw_json or "").strip()
	if not s:
		return False, None
	try:
		ev = json.loads(s)
	except Exception:
		return False, None
	if not isinstance(ev, dict):
		return False, None
	op = str(ev.get("op") or "")
	if op == "move":
		idx = ev.get("idx")
		if idx is not None and 0 <= int(idx) < len(fields):
			f = fields[int(idx)]
			f["x0"] = round(float(ev["x0"]), 1)
			f["y0"] = round(float(ev["y0"]), 1)
			f["x1"] = round(float(ev["x1"]), 1)
			f["y1"] = round(float(ev["y1"]), 1)
			return True, int(idx)
	elif op == "select":
		idx = ev.get("idx")
		if idx is not None and 0 <= int(idx) < len(fields):
			return True, int(idx)
	elif op == "cursor":
		return True, None
	return False, None


def _render_interactive_a4_overlay_stage(
	*,
	bg_img_b64: str,
	bg_mime: str = "image/png",
	page_w: float,
	page_h: float,
	fields: list[dict[str, Any]],
	stage_id: str,
	sync_key: str,
	cursor_xp: float,
	cursor_yp: float,
	selected_idx: int | None = None,
	show_guides: bool = True,
	badge_text: str = "📄 A4 Canlı Sayfa",
) -> None:
	"""Her tür A4 belge görüntüsü üzerinde tıkla-konumlandır, sürükle-taşı, köşeden boyutlandır ve canlı yazı/imza gösterimi sunar."""
	import base64 as _b64

	items_html: list[str] = []
	for idx, f in enumerate(fields):
		x0, y0, x1, y1 = float(f["x0"]), float(f["y0"]), float(f["x1"]), float(f["y1"])
		left_pct = max(0.0, min(98.0, (x0 / max(1.0, page_w)) * 100.0))
		top_pct = max(0.0, min(98.0, (y0 / max(1.0, page_h)) * 100.0))
		w_pct = max(2.0, min(99.0 - left_pct, ((x1 - x0) / max(1.0, page_w)) * 100.0))
		h_pct = max(1.8, min(99.0 - top_pct, ((y1 - y0) / max(1.0, page_h)) * 100.0))
		is_sel = (selected_idx == idx)
		sel_cls = " selected" if is_sel else ""
		ftype = str(f.get("field_type") or "text")

		if ftype == "signature":
			sig_b64 = str(f.get("signature_b64") or "")
			if not sig_b64:
				sb = _resolve_sig_bytes(f)
				if sb:
					sig_b64 = _b64.b64encode(sb).decode("ascii")
			sig_img_tag = f'<img src="data:image/png;base64,{sig_b64}" alt="PNG İmza" draggable="false" style="width:100%;height:100%;object-fit:contain;pointer-events:none;"/>' if sig_b64 else '<span style="font-size:10px;color:#6d28d9;">🖋️ PNG İmza</span>'
			guide_cls = " guided" if show_guides else ""
			items_html.append(
				f'<div class="revir-a4-ov-box is-sig{guide_cls}{sel_cls}" data-ov-idx="{idx}" '
				f'style="left:{left_pct:.2f}%;top:{top_pct:.2f}%;width:{w_pct:.2f}%;height:{h_pct:.2f}%;" '
				f'title="Sürükleyerek taşıyın, sağ alt köşeden boyutlandırın">'
				f'<span class="revir-a4-ov-badge sig">#{idx + 1} İmza</span>'
				f'{sig_img_tag}'
				f'<span class="revir-a4-ov-resize" data-ov-resize="{idx}"></span>'
				f'</div>'
			)
		else:
			val = str(f.get("value") or "")
			if not val.strip() and not show_guides:
				continue
			ff_css = _css_font_family(str(f.get("font_family") or "arial"))
			fs_pt = float(f.get("font_size") or 10.0)
			fw_css = "700" if f.get("bold") else "400"
			fi_css = "italic" if f.get("italic") else "normal"
			al_css = {0: "left", 1: "center", 2: "right"}[_align_to_int(str(f.get("align") or "left"))]
			col_hex = {"lacivert": "#0f3073", "siyah": "#141414", "kirmizi": "#b81c1c", "yesil": "#0a705c"}.get(str(f.get("color") or "lacivert"), "#0f3073")
			bg_css = "rgba(255,255,255,0.92)" if f.get("clear_bg") else ("rgba(240,253,250,0.48)" if (show_guides and val.strip()) else ("rgba(239,246,255,0.38)" if show_guides else "transparent"))
			guide_cls = " guided" if show_guides else ""
			disp_txt = html.escape(val) if val.strip() else f'<span style="opacity:0.42;font-size:9px;">{html.escape(str(f.get("label") or "")[:24])}</span>'
			items_html.append(
				f'<div class="revir-a4-ov-box{guide_cls}{sel_cls}" data-ov-idx="{idx}" data-base-pt="{fs_pt:.1f}" '
				f'style="left:{left_pct:.2f}%;top:{top_pct:.2f}%;width:{w_pct:.2f}%;min-height:{h_pct:.2f}%;'
				f'font-family:{ff_css};font-weight:{fw_css};font-style:{fi_css};text-align:{al_css};color:{col_hex};background:{bg_css};" '
				f'title="Tıklayıp seçin, sürükleyerek A4 üstünde istediğiniz noktaya taşıyın">'
				f'<span class="revir-a4-ov-badge">#{idx + 1}</span>'
				f'<span class="revir-a4-ov-txt">{disp_txt}</span>'
				f'<span class="revir-a4-ov-resize" data-ov-resize="{idx}"></span>'
				f'</div>'
			)

	cursor_html = (
		f'<div class="revir-a4-cursor-pin" style="left:{cursor_xp:.2f}%;top:{cursor_yp:.2f}%;" '
		f'title="Yazı veya imza için seçili konum (X:%{cursor_xp:.0f}, Y:%{cursor_yp:.0f})">'
		f'<span class="revir-a4-pin-dot"></span>'
		f'<span class="revir-a4-pin-lbl">📍 X:%{cursor_xp:.0f} Y:%{cursor_yp:.0f}</span>'
		f'</div>'
	)

	aspect_css = f"{page_w:.1f} / {page_h:.1f}"
	esc_stage = html.escape(stage_id, quote=True)
	esc_sync = html.escape(sync_key, quote=True)

	ui.markdown(
		f"""<style>
		.revir-field:has(input[data-key="{esc_sync}"]) {{ display: none !important; }}
		.revir-a4-live-wrap {{
			position: relative;
			width: 100%;
			max-width: 780px;
			aspect-ratio: {aspect_css};
			background: #ffffff;
			border: 1.5px solid #94a3b8;
			border-radius: 4px;
			box-shadow: 0 12px 32px rgba(15,23,42,0.18);
			overflow: hidden;
			cursor: crosshair;
			user-select: none;
		}}
		.revir-a4-live-bg {{
			width: 100%;
			height: 100%;
			display: block;
			object-fit: contain;
			pointer-events: none;
		}}
		.revir-a4-ov-box {{
			position: absolute;
			box-sizing: border-box;
			padding: 1px 3px;
			line-height: 1.22;
			white-space: pre-wrap;
			word-break: break-word;
			cursor: move;
			border-radius: 3px;
			border: 1px solid transparent;
			z-index: 4;
		}}
		.revir-a4-ov-box.guided {{
			border: 1px dashed rgba(13,148,136,0.72);
		}}
		.revir-a4-ov-box.is-sig.guided {{
			border: 1.5px dashed rgba(109,40,217,0.75);
		}}
		.revir-a4-ov-box.selected, .revir-a4-ov-box:hover {{
			border: 1.5px solid #0284c7 !important;
			box-shadow: 0 0 0 2px rgba(2,132,199,0.22);
			z-index: 8;
		}}
		.revir-a4-ov-badge {{
			display: none;
			position: absolute;
			top: -13px;
			left: -1px;
			background: #0d9488;
			color: #ffffff;
			font-size: 9px;
			font-weight: 800;
			font-family: Arial, sans-serif;
			padding: 0 4px;
			border-radius: 3px;
			line-height: 12px;
			pointer-events: none;
		}}
		.revir-a4-ov-badge.sig {{ background: #6d28d9; }}
		.revir-a4-ov-box.guided .revir-a4-ov-badge,
		.revir-a4-ov-box:hover .revir-a4-ov-badge,
		.revir-a4-ov-box.selected .revir-a4-ov-badge {{ display: inline-block; }}
		.revir-a4-ov-resize {{
			display: none;
			position: absolute;
			right: -4px;
			bottom: -4px;
			width: 10px;
			height: 10px;
			border-radius: 2px;
			background: #0284c7;
			border: 1px solid #ffffff;
			cursor: nwse-resize;
		}}
		.revir-a4-ov-box:hover .revir-a4-ov-resize,
		.revir-a4-ov-box.selected .revir-a4-ov-resize,
		.revir-a4-ov-box.is-sig .revir-a4-ov-resize {{ display: block; }}
		.revir-a4-cursor-pin {{
			position: absolute;
			transform: translate(-5px, -5px);
			pointer-events: none;
			z-index: 10;
			display: flex;
			align-items: center;
			gap: 4px;
		}}
		.revir-a4-pin-dot {{
			width: 10px;
			height: 10px;
			border-radius: 50%;
			background: #e11d48;
			border: 2px solid #ffffff;
			box-shadow: 0 1px 4px rgba(0,0,0,0.4);
		}}
		.revir-a4-pin-lbl {{
			background: rgba(15,23,42,0.86);
			color: #ffffff;
			font-size: 9.5px;
			font-weight: 700;
			padding: 1px 6px;
			border-radius: 999px;
			white-space: nowrap;
		}}
		</style>""",
		unsafe_allow_html=True,
	)
	ui.markdown(
		f"""<div class="revir-a4-stage" style="padding:12px;">
			<div style="display:flex;justify-content:space-between;align-items:center;width:100%;max-width:780px;gap:8px;flex-wrap:wrap;">
				<span class="revir-a4-page-badge">{html.escape(badge_text)}</span>
				<span style="font-size:11px;font-weight:700;color:#475569;">🖱️ Konum seçmek için kâğıda tıklayın · Kutuları/imzayı sürükleyip boyutlandırın</span>
			</div>
			<div id="{esc_stage}" class="revir-a4-live-wrap" data-pw="{page_w:.2f}" data-ph="{page_h:.2f}" data-sync-key="{esc_sync}">
				<img class="revir-a4-live-bg" src="data:{bg_mime};base64,{bg_img_b64}" alt="A4 Önizleme" draggable="false"/>
				{''.join(items_html)}
				{cursor_html}
			</div>
		</div>""",
		unsafe_allow_html=True,
	)

	stage_js = r"""
<script>
(function(){
	var stage = document.getElementById(__STAGE_ID__);
	if (!stage || stage.dataset.a4Bound === "1") return;
	stage.dataset.a4Bound = "1";
	var pw = parseFloat(stage.dataset.pw) || 595.0;
	var ph = parseFloat(stage.dataset.ph) || 842.0;
	var syncKey = stage.dataset.syncKey || "";

	function scaleFonts() {
		var r = stage.getBoundingClientRect();
		if (!r.width) return;
		var factor = r.width / pw;
		stage.querySelectorAll(".revir-a4-ov-box[data-base-pt]").forEach(function(box){
			var pt = parseFloat(box.dataset.basePt) || 10.0;
			box.style.fontSize = Math.max(7.0, pt * factor).toFixed(2) + "px";
		});
	}
	scaleFonts();
	if (typeof ResizeObserver !== "undefined") {
		try { new ResizeObserver(scaleFonts).observe(stage); } catch (e) {}
	}

	function emitSync(obj) {
		if (window.__revirFieldChange && syncKey) {
			window.__revirFieldChange(syncKey, JSON.stringify(obj));
		}
	}

	var dragState = null;

	stage.addEventListener("mousedown", function(ev){
		var resHandle = ev.target.closest ? ev.target.closest("[data-ov-resize]") : null;
		var box = ev.target.closest ? ev.target.closest("[data-ov-idx]") : null;
		var sr = stage.getBoundingClientRect();
		if (!sr.width || !sr.height) return;

		if (resHandle && box) {
			ev.preventDefault();
			ev.stopPropagation();
			dragState = {
				mode: "resize",
				idx: parseInt(box.dataset.ovIdx, 10),
				el: box,
				startX: ev.clientX,
				startY: ev.clientY,
				origLeftPct: parseFloat(box.style.left) || 0,
				origTopPct: parseFloat(box.style.top) || 0,
				origWpx: box.getBoundingClientRect().width,
				origHpx: box.getBoundingClientRect().height,
				moved: false
			};
			return;
		}
		if (box) {
			ev.preventDefault();
			ev.stopPropagation();
			dragState = {
				mode: "move",
				idx: parseInt(box.dataset.ovIdx, 10),
				el: box,
				startX: ev.clientX,
				startY: ev.clientY,
				origLeftPct: parseFloat(box.style.left) || 0,
				origTopPct: parseFloat(box.style.top) || 0,
				moved: false
			};
			return;
		}
		// Boş A4 kâğıda tıklandı -> imleç hedefi yerleştir
		var xp = Math.max(1.0, Math.min(96.0, ((ev.clientX - sr.left) / sr.width) * 100.0));
		var yp = Math.max(1.0, Math.min(96.0, ((ev.clientY - sr.top) / sr.height) * 100.0));
		var pin = stage.querySelector(".revir-a4-cursor-pin");
		if (pin) {
			pin.style.left = xp.toFixed(2) + "%";
			pin.style.top = yp.toFixed(2) + "%";
			var lbl = pin.querySelector(".revir-a4-pin-lbl");
			if (lbl) lbl.textContent = "📍 X:%" + Math.round(xp) + " Y:%" + Math.round(yp);
		}
		emitSync({ op: "cursor", xp: Number(xp.toFixed(1)), yp: Number(yp.toFixed(1)) });
	});

	window.addEventListener("mousemove", function(ev){
		if (!dragState) return;
		var sr = stage.getBoundingClientRect();
		if (!sr.width || !sr.height) return;
		var dx = ev.clientX - dragState.startX;
		var dy = ev.clientY - dragState.startY;
		if (Math.abs(dx) > 2 || Math.abs(dy) > 2) dragState.moved = true;
		if (dragState.mode === "move") {
			var nl = Math.max(0.0, Math.min(95.0, dragState.origLeftPct + (dx / sr.width) * 100.0));
			var nt = Math.max(0.0, Math.min(96.0, dragState.origTopPct + (dy / sr.height) * 100.0));
			dragState.el.style.left = nl.toFixed(2) + "%";
			dragState.el.style.top = nt.toFixed(2) + "%";
		} else if (dragState.mode === "resize") {
			var nwPct = Math.max(3.0, Math.min(98.0 - dragState.origLeftPct, ((dragState.origWpx + dx) / sr.width) * 100.0));
			var nhPct = Math.max(2.0, Math.min(98.0 - dragState.origTopPct, ((dragState.origHpx + dy) / sr.height) * 100.0));
			dragState.el.style.width = nwPct.toFixed(2) + "%";
			if (dragState.el.classList.contains("is-sig")) {
				dragState.el.style.height = nhPct.toFixed(2) + "%";
			} else {
				dragState.el.style.minHeight = nhPct.toFixed(2) + "%";
			}
		}
	});

	window.addEventListener("mouseup", function(){
		if (!dragState) return;
		var ds = dragState;
		dragState = null;
		if (!ds.moved) {
			emitSync({ op: "select", idx: ds.idx });
			return;
		}
		var sr = stage.getBoundingClientRect();
		var br = ds.el.getBoundingClientRect();
		if (!sr.width || !sr.height) return;
		var x0 = Math.max(2.0, ((br.left - sr.left) / sr.width) * pw);
		var y0 = Math.max(2.0, ((br.top - sr.top) / sr.height) * ph);
		var x1 = Math.min(pw - 2.0, ((br.right - sr.left) / sr.width) * pw);
		var y1 = Math.min(ph - 2.0, ((br.bottom - sr.top) / sr.height) * ph);
		emitSync({ op: "move", idx: ds.idx, x0: x0, y0: y0, x1: x1, y1: y1 });
	});
})();
</script>
""".replace("__STAGE_ID__", json.dumps(stage_id))
	ui.components.v1.html(stage_js, height=0)


# ==============================================================================
# 3. WORD (.DOCX) & EXCEL (.XLSX) OKUMA / DÜZENLEME / KAYDETME MOTORLARI
# ==============================================================================

def _load_docx_structure(file_path: Path) -> tuple[list[dict[str, Any]], list[pd.DataFrame]]:
	import docx

	d = docx.Document(file_path)
	paras: list[dict[str, Any]] = []
	for i, p in enumerate(d.paragraphs):
		style_nm = str(p.style.name if p.style else "Normal")
		paras.append({
			"index": i,
			"style": style_nm,
			"text": p.text or "",
		})
	tables_dfs: list[pd.DataFrame] = []
	for tbl in d.tables:
		raw_rows = [[c.text.strip() for c in row.cells] for row in tbl.rows]
		if not raw_rows:
			tables_dfs.append(pd.DataFrame())
			continue
		max_c = max(len(r) for r in raw_rows)
		norm_rows = [r + [""] * (max_c - len(r)) for r in raw_rows]
		if len(norm_rows) > 1:
			raw_hdrs = norm_rows[0]
			seen: dict[str, int] = {}
			hdrs = []
			for ci, h in enumerate(raw_hdrs):
				base = h.strip() or f"Sütun {ci + 1}"
				if base in seen:
					seen[base] += 1
					hdrs.append(f"{base} ({seen[base]})")
				else:
					seen[base] = 1
					hdrs.append(base)
			tables_dfs.append(pd.DataFrame(norm_rows[1:], columns=hdrs))
		else:
			hdrs = [f"Sütun {ci + 1}" for ci in range(max_c)]
			tables_dfs.append(pd.DataFrame(norm_rows, columns=hdrs))
	return paras, tables_dfs


def _save_docx_structure(
	src_path: Path,
	dest_path: Path,
	edited_paras: dict[int, str],
	appended_paras: list[str],
	edited_tables: dict[int, pd.DataFrame],
	*,
	font_family: str = "",
	font_size_pt: float | None = None,
	default_align: str = "",
	signature_overlays: list[dict[str, Any]] | None = None,
	extra_text_overlays: list[dict[str, Any]] | None = None,
) -> None:
	import docx
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.shared import Cm, Pt

	d = docx.Document(src_path)
	align_enum = {
		"Sola Hizalı": WD_ALIGN_PARAGRAPH.LEFT,
		"Ortala": WD_ALIGN_PARAGRAPH.CENTER,
		"Sağa Hizalı": WD_ALIGN_PARAGRAPH.RIGHT,
	}.get(default_align)
	docx_font_name = {
		"times": "Times New Roman",
		"courier": "Courier New",
		"calibri": "Calibri",
		"arial": "Arial",
	}.get(_font_family_code(font_family)) if font_family else None

	for p_idx, new_txt in edited_paras.items():
		if 0 <= p_idx < len(d.paragraphs):
			p = d.paragraphs[p_idx]
			if p.text != new_txt:
				if p.runs:
					p.runs[0].text = new_txt
					for r in p.runs[1:]:
						r.text = ""
				else:
					p.text = new_txt

	for add_t in appended_paras:
		if str(add_t).strip():
			d.add_paragraph(str(add_t).strip())

	# İsteğe bağlı serbest metin notlarını ekle
	for ov in (extra_text_overlays or []):
		val = str(ov.get("value") or "").strip()
		if val:
			p_ov = d.add_paragraph(val)
			ov_al = {
				1: WD_ALIGN_PARAGRAPH.CENTER,
				2: WD_ALIGN_PARAGRAPH.RIGHT,
			}.get(_align_to_int(str(ov.get("align") or "left")), WD_ALIGN_PARAGRAPH.LEFT)
			p_ov.alignment = ov_al
			ov_fn = {
				"times": "Times New Roman",
				"courier": "Courier New",
				"calibri": "Calibri",
				"arial": "Arial",
			}.get(_font_family_code(str(ov.get("font_family") or "arial")), "Arial")
			for r in p_ov.runs:
				r.font.name = ov_fn
				if ov.get("font_size"):
					r.font.size = Pt(float(ov["font_size"]))
				r.bold = bool(ov.get("bold"))
				r.italic = bool(ov.get("italic"))

	# Toplu tipografi ayarı uygulanacaksa normal paragrafları güncelle
	if docx_font_name or font_size_pt or align_enum is not None:
		for p in d.paragraphs:
			st_nm = str(p.style.name if p.style else "").lower()
			is_heading = ("heading" in st_nm or "title" in st_nm)
			if align_enum is not None and not is_heading:
				p.alignment = align_enum
			for r in p.runs:
				if docx_font_name:
					r.font.name = docx_font_name
				if font_size_pt and not is_heading:
					r.font.size = Pt(float(font_size_pt))

	# PNG imzaları belgeye ekle
	for sig in (signature_overlays or []):
		sb = _resolve_sig_bytes(sig)
		if not sb:
			continue
		p_sig = d.add_paragraph()
		# Yatay konum yüzde > 55 ise sağa hizalı, 35..55 orta, altı sol
		x0_val = float(sig.get("x0") or 360.0)
		if x0_val >= 320.0:
			p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
		elif x0_val >= 190.0:
			p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
		else:
			p_sig.alignment = WD_ALIGN_PARAGRAPH.LEFT
		w_pt = max(60.0, float(sig.get("x1") or (x0_val + 130.0)) - x0_val)
		w_cm = max(1.8, min(9.0, w_pt / 28.346))
		run_s = p_sig.add_run()
		run_s.add_picture(io.BytesIO(sb), width=Cm(round(w_cm, 2)))

	for t_idx, df in edited_tables.items():
		if not (0 <= t_idx < len(d.tables)) or df is None:
			continue
		tbl = d.tables[t_idx]
		has_header_row = len(tbl.rows) > 1
		cols = [str(c) for c in df.columns]
		new_rows_vals: list[list[str]] = []
		if has_header_row:
			new_rows_vals.append([re.sub(r"\s+\(\d+\)$", "", c) for c in cols])
		for _, row in df.iterrows():
			new_rows_vals.append(["" if pd.isna(v) else str(v) for v in row.tolist()])

		while len(tbl.rows) < len(new_rows_vals):
			tbl.add_row()
		while len(tbl.rows) > len(new_rows_vals) and len(tbl.rows) > 1:
			tbl._tbl.remove(tbl.rows[-1]._tr)

		for ri, r_vals in enumerate(new_rows_vals):
			if ri >= len(tbl.rows):
				break
			cells = tbl.rows[ri].cells
			for ci, val in enumerate(r_vals):
				if ci < len(cells):
					cells[ci].text = val
					if docx_font_name or font_size_pt:
						for cp in cells[ci].paragraphs:
							for cr in cp.runs:
								if docx_font_name:
									cr.font.name = docx_font_name
								if font_size_pt:
									cr.font.size = Pt(max(8.0, float(font_size_pt) - 1.0))

	d.save(dest_path)


def _load_xlsx_sheet_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
	import openpyxl

	wb = openpyxl.load_workbook(file_path, data_only=False)
	try:
		if sheet_name not in wb.sheetnames:
			return pd.DataFrame()
		ws = wb[sheet_name]
		rows_raw = [list(r) for r in ws.iter_rows(values_only=True)]
		while rows_raw and all(v is None or str(v).strip() == "" for v in rows_raw[-1]):
			rows_raw.pop()
		if not rows_raw:
			return pd.DataFrame(columns=["Sütun 1", "Sütun 2", "Sütun 3"])
		max_c = max(len(r) for r in rows_raw)
		norm = [["" if v is None else str(v) for v in (list(r) + [None] * (max_c - len(r)))] for r in rows_raw]
		raw_hdr = norm[0]
		seen: dict[str, int] = {}
		hdrs = []
		for ci, h in enumerate(raw_hdr):
			base = h.strip() or f"Sütun {ci + 1}"
			if base in seen:
				seen[base] += 1
				hdrs.append(f"{base} ({seen[base]})")
			else:
				seen[base] = 1
				hdrs.append(base)
		body = norm[1:] if len(norm) > 1 else []
		return pd.DataFrame(body, columns=hdrs)
	finally:
		wb.close()


def _save_xlsx_sheet_df(
	src_path: Path,
	dest_path: Path,
	sheet_name: str,
	df: pd.DataFrame,
	new_sheet_title: str = "",
	*,
	font_family: str = "",
	font_size_pt: float | None = None,
	landscape: bool | None = None,
	signature_overlays: list[dict[str, Any]] | None = None,
) -> None:
	import openpyxl
	from openpyxl.styles import Font

	wb = openpyxl.load_workbook(src_path)
	try:
		if new_sheet_title.strip():
			ws = wb.create_sheet(title=new_sheet_title.strip()[:31])
		elif sheet_name in wb.sheetnames:
			ws = wb[sheet_name]
		else:
			ws = wb.active

		xl_font_name = {
			"times": "Times New Roman",
			"courier": "Courier New",
			"calibri": "Calibri",
			"arial": "Arial",
		}.get(_font_family_code(font_family)) if font_family else None

		if landscape is not None:
			ws.page_setup.paperSize = ws.PAPERSIZE_A4
			ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if landscape else ws.ORIENTATION_PORTRAIT

		# Mevcut veri satırlarını temizle (üst başlık stilini koruyarak)
		old_max_r = ws.max_row
		old_max_c = ws.max_column
		cols = [re.sub(r"\s+\(\d+\)$", "", str(c)) for c in df.columns]

		for ci, col_nm in enumerate(cols, start=1):
			cell = ws.cell(row=1, column=ci, value=col_nm)
			if xl_font_name or font_size_pt:
				cell.font = Font(
					name=xl_font_name or (cell.font.name if cell.font else "Calibri"),
					size=float(font_size_pt) if font_size_pt else (cell.font.size if cell.font else 11),
					bold=True,
				)
			elif cell.font is None or not cell.font.bold:
				cell.font = Font(bold=True)

		for ri, (_, row) in enumerate(df.iterrows(), start=2):
			for ci, val in enumerate(row.tolist(), start=1):
				v_str = "" if pd.isna(val) else str(val).strip()
				out_v: Any = v_str
				if v_str and not v_str.startswith("0") and len(v_str) <= 9:
					try:
						if re.match(r"^-?\d+$", v_str):
							out_v = int(v_str)
						elif re.match(r"^-?\d+[.,]\d+$", v_str):
							out_v = float(v_str.replace(",", "."))
					except Exception:
						out_v = v_str
				dc = ws.cell(row=ri, column=ci, value=out_v)
				if xl_font_name or font_size_pt:
					dc.font = Font(
						name=xl_font_name or (dc.font.name if dc.font else "Calibri"),
						size=float(font_size_pt) if font_size_pt else (dc.font.size if dc.font else 10),
					)

		new_max_r = len(df) + 1
		for r_clear in range(new_max_r + 1, old_max_r + 1):
			for c_clear in range(1, max(old_max_c, len(cols)) + 1):
				ws.cell(row=r_clear, column=c_clear, value=None)

		# PNG imzaları Excel çalışma sayfasına resim olarak ekle
		for s_idx, sig in enumerate(signature_overlays or []):
			sb = _resolve_sig_bytes(sig)
			if not sb:
				continue
			try:
				from openpyxl.drawing.image import Image as XLImage
				xl_img = XLImage(io.BytesIO(sb))
				xl_img.width = 140
				xl_img.height = 60
				col_idx = max(1, min(max(1, len(cols)), int(round((float(sig.get("x0") or 240.0) / 842.0) * max(1, len(cols)))) + 1))
				row_anchor = new_max_r + 2 + (s_idx * 4)
				anchor_cell = f"{_excel_col_letter(col_idx - 1)}{row_anchor}"
				ws.add_image(xl_img, anchor_cell)
			except Exception:
				pass

		wb.save(dest_path)
	finally:
		wb.close()


# ==============================================================================
# 4. EVRAK STÜDYOSU DİYALOĞU (WORD / EXCEL / PDF FORM TARAMA & DOLDURMA)
# ==============================================================================

def _render_overlay_and_signature_controls(
	*,
	fields: list[dict[str, Any]],
	page_idx: int,
	pw: float,
	ph: float,
	evrak_dir: Path,
	ns: str,
	flash_k: str,
	cursor_xp_k: str,
	cursor_yp_k: str,
	sel_idx_k: str,
	allow_position_controls: bool = True,
) -> None:
	"""Tüm stüdyolar (PDF, Word, Excel, Görsel) için ortak: seçili konuma yazı/PNG imza ekleme ve canlı tipografi düzenleyicisi."""
	import base64 as _b64

	cur_xp = float(ui.session_state.get(cursor_xp_k, 18.0))
	cur_yp = float(ui.session_state.get(cursor_yp_k, 78.0))
	sel_idx = ui.session_state.get(sel_idx_k)
	if sel_idx is not None and not (0 <= int(sel_idx) < len(fields)):
		sel_idx = None
		ui.session_state[sel_idx_k] = None

	# 1. Aktif seçili katman veya tüm katmanlar için canlı tipografi / içerik düzenleyici
	if sel_idx is not None and fields:
		sf = fields[int(sel_idx)]
		sf_id = sf["id"]
		is_sig = str(sf.get("field_type") or "") == "signature"
		with ui.container(border=True):
			hdr_c1, hdr_c2 = ui.columns([3.2, 0.8])
			with hdr_c1:
				ui.markdown(f"**🎯 Seçili Katman (#{int(sel_idx) + 1} · {html.escape(str(sf.get('label') or ('PNG İmza' if is_sig else 'Metin')))})**")
			with hdr_c2:
				if ui.button("🗑️ Kaldır", key=f"{ns}_sel_del_{sf_id}", use_container_width=True):
					fields.pop(int(sel_idx))
					ui.session_state[sel_idx_k] = None
					ui.rerun(scope="fragment")

			if not is_sig:
				sf["value"] = ui.text_input(
					"İçerik",
					value=str(sf.get("value") or ""),
					key=f"{ns}_sel_val_{sf_id}",
				)
				st1, st2, st3 = ui.columns([1.8, 0.9, 1.1])
				ff_cur = str(sf.get("font_family") or FONT_FAMILY_OPTIONS[0])
				if ff_cur not in FONT_FAMILY_OPTIONS:
					ff_cur = FONT_FAMILY_OPTIONS[0]
				sf["font_family"] = st1.selectbox(
					"Yazı Tipi",
					FONT_FAMILY_OPTIONS,
					index=FONT_FAMILY_OPTIONS.index(ff_cur),
					key=f"{ns}_sel_ff_{sf_id}",
				)
				sf["font_size"] = int(
					st2.number_input("Boyut (pt)", min_value=6, max_value=36, value=int(sf.get("font_size") or 10), step=1, key=f"{ns}_sel_fs_{sf_id}")
				)
				al_cur = str(sf.get("align") or ALIGN_OPTIONS[0])
				if al_cur not in ALIGN_OPTIONS:
					al_cur = ALIGN_OPTIONS[0]
				sf["align"] = st3.selectbox("Hizalama", ALIGN_OPTIONS, index=ALIGN_OPTIONS.index(al_cur), key=f"{ns}_sel_al_{sf_id}")

				sb1, sb2, sb3, sb4 = ui.columns([0.8, 0.8, 1.2, 1.4])
				sf["bold"] = sb1.checkbox("Kalın (B)", value=bool(sf.get("bold")), key=f"{ns}_sel_b_{sf_id}")
				sf["italic"] = sb2.checkbox("İtalik (I)", value=bool(sf.get("italic")), key=f"{ns}_sel_i_{sf_id}")
				col_opts = ["lacivert", "siyah", "kirmizi", "yesil"]
				sf["color"] = sb3.selectbox("Renk", col_opts, index=col_opts.index(sf.get("color", "lacivert")) if sf.get("color") in col_opts else 0, key=f"{ns}_sel_col_{sf_id}")
				sf["clear_bg"] = sb4.checkbox("Nokta/çizgi ört", value=bool(sf.get("clear_bg")), key=f"{ns}_sel_bg_{sf_id}")

			if allow_position_controls:
				sp1, sp2, sp3, sp4 = ui.columns(4)
				sf["x0"] = float(sp1.number_input("Sol X (pt)", min_value=0.0, max_value=pw, value=round(float(sf["x0"]), 1), step=4.0, key=f"{ns}_sel_x0_{sf_id}"))
				sf["y0"] = float(sp2.number_input("Üst Y (pt)", min_value=0.0, max_value=ph, value=round(float(sf["y0"]), 1), step=4.0, key=f"{ns}_sel_y0_{sf_id}"))
				cur_w = max(20.0, float(sf["x1"]) - float(sf["x0"]))
				new_w = float(sp3.number_input("Genişlik (pt)", min_value=20.0, max_value=pw, value=round(cur_w, 1), step=8.0, key=f"{ns}_sel_w_{sf_id}"))
				sf["x1"] = min(pw - 2.0, sf["x0"] + new_w)
				cur_h = max(14.0, float(sf["y1"]) - float(sf["y0"]))
				new_h = float(sp4.number_input("Yükseklik (pt)", min_value=14.0, max_value=ph, value=round(cur_h, 1), step=6.0, key=f"{ns}_sel_h_{sf_id}"))
				sf["y1"] = min(ph - 2.0, sf["y0"] + new_h)

	# 2. Seçili A4 Konumuna Hızlı Metin Ekleme
	with ui.container(border=True):
		ui.markdown(f"##### ➕ A4 Konuma Yazı Ekle &nbsp;<small style='color:#0f766e;'>(📍 X:%{cur_xp:.0f} · Y:%{cur_yp:.0f})</small>", unsafe_allow_html=True)
		m_lbl = ui.text_input("Etiket", value="Serbest Metin", key=f"{ns}_man_lbl")
		m_val = ui.text_input("Yazılacak Metin", placeholder="Sol A4 kâğıda tıklayıp buraya metin yazın (Örn: ✓, Ad Soyad, Açıklama...)", key=f"{ns}_man_val")
		mc1, mc2, mc3 = ui.columns(3)
		m_ff = mc1.selectbox("Yazı Tipi", FONT_FAMILY_OPTIONS, index=0, key=f"{ns}_man_ff")
		m_fs = int(mc2.number_input("Punto (pt)", min_value=6, max_value=36, value=11, step=1, key=f"{ns}_man_fs"))
		m_col = mc3.selectbox("Renk", ["lacivert", "siyah", "kirmizi", "yesil"], index=0, key=f"{ns}_man_col")

		md1, md2, md3, md4 = ui.columns([1.0, 1.0, 1.2, 1.0])
		m_xp = float(md1.number_input("Yatay (%)", min_value=1.0, max_value=95.0, value=round(cur_xp, 1), step=2.0, key=f"{ns}_man_xp"))
		m_yp = float(md2.number_input("Dikey (%)", min_value=1.0, max_value=96.0, value=round(cur_yp, 1), step=2.0, key=f"{ns}_man_yp"))
		m_wp = float(md3.number_input("Genişlik (%)", min_value=6.0, max_value=92.0, value=36.0, step=2.0, key=f"{ns}_man_wp"))
		m_bold = md4.checkbox("Kalın (B)", value=False, key=f"{ns}_man_bold")

		if ui.button("➕ Bu Konuma Metin Ekle", type="primary", use_container_width=True, key=f"{ns}_man_add"):
			nx0 = round(pw * m_xp / 100.0, 1)
			ny0 = round(ph * m_yp / 100.0, 1)
			nx1 = round(min(pw - 8.0, nx0 + pw * m_wp / 100.0), 1)
			ny1 = round(min(ph - 8.0, ny0 + max(18.0, m_fs * 1.9)), 1)
			fields.append({
				"id": f"man_{page_idx}_{len(fields) + 1}_{int(datetime.now().timestamp()) % 10000}",
				"page": page_idx,
				"label": m_lbl.strip() or "Serbest Metin",
				"field_type": "text",
				"source": "manual",
				"widget_name": "",
				"x0": nx0, "y0": ny0, "x1": nx1, "y1": ny1,
				"value": m_val,
				"font_family": m_ff,
				"font_size": m_fs,
				"bold": m_bold,
				"italic": False,
				"align": ALIGN_OPTIONS[0],
				"color": m_col,
				"clear_bg": False,
			})
			ui.session_state[sel_idx_k] = len(fields) - 1
			ui.session_state[flash_k] = f"➕ '{m_lbl}' A4 sayfa üzerinde (X:%{m_xp:.0f}, Y:%{m_yp:.0f}) konumuna eklendi."
			ui.rerun(scope="fragment")

	# 3. PNG İmza Yükleme, Kurum İmza Kütüphanesi ve A4 Sayfaya Damgalama
	with ui.container(border=True):
		ui.markdown(f"##### 🖋️ PNG İmza / Kaşe At &nbsp;<small style='color:#6d28d9;'>(📍 Hedef: X:%{cur_xp:.0f} · Y:%{cur_yp:.0f})</small>", unsafe_allow_html=True)
		saved_sigs = list_saved_signatures(evrak_dir)
		sig_up = ui.file_uploader(
			"PNG veya Fotoğraf İmza Yükle (.png / .jpg)",
			type=["png", "jpg", "jpeg", "webp"],
			accept_multiple_files=False,
			key=f"{ns}_sig_upload",
		)
		sc_opt1, sc_opt2 = ui.columns([1.4, 1.1])
		rm_white = sc_opt1.checkbox("🪄 Beyaz arka planı otomatik şeffaf PNG yap", value=True, key=f"{ns}_sig_rm_bg")
		sig_wp = float(sc_opt2.number_input("İmza Genişliği (%)", min_value=8.0, max_value=55.0, value=24.0, step=2.0, key=f"{ns}_sig_wp"))

		chosen_sig_bytes: bytes | None = None
		chosen_sig_label = "İmza"
		if sig_up is not None:
			raw_s_bytes = sig_up.read()
			if raw_s_bytes:
				chosen_sig_bytes = prepare_signature_png_bytes(raw_s_bytes, remove_white_bg=rm_white)
				chosen_sig_label = Path(sig_up.name).stem
				if ui.button("💾 Bu İmzayı Kurum İmza Kütüphanesine Kaydet (Tekrar Kullan)", type="secondary", use_container_width=True, key=f"{ns}_sig_save_lib"):
					save_signature_to_library(evrak_dir, chosen_sig_label, raw_s_bytes, remove_white_bg=rm_white)
					ui.session_state[flash_k] = f"💾 '{chosen_sig_label}' PNG imza olarak kurum kütüphanesine kaydedildi."
					ui.rerun(scope="fragment")
		elif saved_sigs:
			sig_names = [s["name"] for s in saved_sigs]
			sel_sn = ui.selectbox("📂 Kayıtlı Kurum PNG İmzaları", sig_names, index=0, key=f"{ns}_sig_lib_pick")
			matched_s = next((s for s in saved_sigs if s["name"] == sel_sn), saved_sigs[0])
			chosen_sig_bytes = matched_s["bytes"]
			chosen_sig_label = matched_s["name"]

		if chosen_sig_bytes:
			b64_prev = _b64.b64encode(chosen_sig_bytes).decode("ascii")
			ui.markdown(
				f"<div style='background:#f8fafc;border:1px dashed #cbd5e1;border-radius:8px;padding:6px;text-align:center;margin:4px 0;'>"
				f"<img src='data:image/png;base64,{b64_prev}' style='max-height:56px;max-width:100%;object-fit:contain;' alt='İmza Önizleme'/>"
				f"</div>",
				unsafe_allow_html=True,
			)
			if ui.button("🖋️ Seçili A4 Noktasına PNG İmza At", type="primary", use_container_width=True, key=f"{ns}_sig_stamp_btn"):
				from PIL import Image as _PImg
				im_s = _PImg.open(io.BytesIO(chosen_sig_bytes))
				aspect_s = im_s.height / max(1, im_s.width)
				im_s.close()
				sx0 = round(pw * m_xp / 100.0, 1)
				sy0 = round(ph * m_yp / 100.0, 1)
				sw_pt = round(pw * sig_wp / 100.0, 1)
				sh_pt = round(max(24.0, min(ph * 0.35, sw_pt * aspect_s)), 1)
				sx1 = round(min(pw - 6.0, sx0 + sw_pt), 1)
				sy1 = round(min(ph - 6.0, sy0 + sh_pt), 1)
				fields.append({
					"id": f"sig_{page_idx}_{len(fields) + 1}_{int(datetime.now().timestamp()) % 10000}",
					"page": page_idx,
					"label": f"PNG İmza ({chosen_sig_label})",
					"field_type": "signature",
					"source": "signature",
					"widget_name": "",
					"x0": sx0, "y0": sy0, "x1": sx1, "y1": sy1,
					"value": chosen_sig_label,
					"signature_b64": _b64.b64encode(chosen_sig_bytes).decode("ascii"),
				})
				ui.session_state[sel_idx_k] = len(fields) - 1
				ui.session_state[flash_k] = f"🖋️ PNG imza A4 sayfasında (X:%{m_xp:.0f}, Y:%{m_yp:.0f}) konumuna yerleştirildi!"
				ui.rerun(scope="fragment")
		else:
			ui.caption("💡 Yukarıdan bir PNG/JPG imza yükleyin; beyaz arka planı otomatik şeffaf PNG'ye çevrilip istediğiniz noktaya basılır.")


def render_document_studio_dialog(file_path: Path, ctx_api: dict[str, Any]) -> None:
	"""
	Evrak Deposundaki Word (.docx), Excel (.xlsx), PDF (.pdf), Görsel (.png/.jpg/.webp) veya Metin belgesini
	canlı A4 görsel tuval, yazı tipi/boyutu ayarları ve şeffaf PNG imza desteğiyle doğrudan düzenler.
	"""
	if not file_path.exists():
		ui.error("Dosya bulunamadı.")
		return

	ext = file_path.suffix.lower()
	fhash = sha256(str(file_path.name).encode("utf-8")).hexdigest()[:10]
	evrak_dir: Path = ctx_api["EVRAKLAR_DIR"]
	is_gai_on, gai_key, gai_model = ctx_api["_get_active_gemini_config"]()
	app_cfg = ctx_api["get_app_settings"]()
	active_pid = ui.session_state.get("active_person_id")
	active_person = dict(ctx_api["get_person_by_id"](int(active_pid))) if active_pid else None

	h1, h2 = ui.columns([7.2, 2.8])
	with h1:
		ui.markdown(f"#### 🛠️ Belge Düzenleme & A4 İmza Stüdyosu — `{file_path.name}`")
		act_badge = f" · 📌 Aktif Kişi: **{ctx_api['mask_kvkk_value'](active_person['full_name'], 'name')}** ({active_person.get('class_name', '—')})" if active_person else ""
		ui.caption(f"🏷️ Tür: **{ext.upper().strip('.')}** · 🏫 Kurum: **{app_cfg.get('school_name', 'Okul Reviri')}**{act_badge}")
	with h2:
		mime_t = mimetypes.guess_type(file_path)[0] or "application/octet-stream"
		ui.download_button(
			"📥 Güncel Dosyayı İndir",
			data=lambda _p=file_path: _p.read_bytes(),
			file_name=file_path.name,
			mime=mime_t,
			type="secondary",
			use_container_width=True,
			key=f"studio_dl_{fhash}",
		)

	flash_k = f"studio_flash_{fhash}"
	f_msg = ui.session_state.pop(flash_k, None)
	if f_msg:
		ui.success(f_msg)

	if ext == ".pdf":
		_render_pdf_studio(file_path, fhash, evrak_dir, ctx_api, app_cfg, active_person, is_gai_on, gai_key, gai_model, flash_k)
	elif ext == ".docx":
		_render_docx_studio(file_path, fhash, evrak_dir, ctx_api, app_cfg, active_person, is_gai_on, gai_key, gai_model, flash_k)
	elif ext == ".xlsx":
		_render_xlsx_studio(file_path, fhash, evrak_dir, ctx_api, app_cfg, active_person, is_gai_on, gai_key, gai_model, flash_k)
	elif ext in (".png", ".jpg", ".jpeg", ".webp"):
		_render_image_studio(file_path, fhash, evrak_dir, ctx_api, flash_k)
	elif ext in (".txt", ".csv"):
		_render_text_studio(file_path, fhash, evrak_dir, ctx_api, flash_k)
	else:
		ui.info(f"`{ext}` biçimi için doğrudan düzenleme şu an desteklenmiyor. Word (`.docx`), Excel (`.xlsx`), PDF (`.pdf`), Resim (`.png`/`.jpg`) veya Metin (`.txt`/`.csv`) dosyalarını bu stüdyoda düzenleyebilirsiniz.")


def _render_pdf_studio(
	file_path: Path,
	fhash: str,
	evrak_dir: Path,
	ctx_api: dict[str, Any],
	app_cfg: dict[str, str],
	active_person: dict[str, Any] | None,
	is_gai_on: bool,
	gai_key: str,
	gai_model: str,
	flash_k: str,
) -> None:
	import base64 as _b64
	import pymupdf

	doc_meta = pymupdf.open(file_path)
	num_pages = len(doc_meta)
	page_dims = [(float(doc_meta[i].rect.width), float(doc_meta[i].rect.height)) for i in range(num_pages)]
	doc_meta.close()

	if num_pages == 0:
		ui.warning("PDF dosyasında sayfa bulunamadı.")
		return

	state_key = f"pdf_studio_fields_{fhash}"
	if state_key not in ui.session_state:
		ui.session_state[state_key] = {}

	tb1, tb2, tb3, tb4 = ui.columns([1.3, 1.5, 1.8, 1.4])
	with tb1:
		page_opts = [f"Sayfa {i + 1} / {num_pages}" for i in range(num_pages)]
		sel_pg_lbl = ui.selectbox("📄 Çalışılan Sayfa", page_opts, index=0, key=f"pdf_pg_sel_{fhash}")
		cur_pg = page_opts.index(sel_pg_lbl)
	pw, ph = page_dims[cur_pg]

	if cur_pg not in ui.session_state[state_key]:
		ui.session_state[state_key][cur_pg] = scan_pdf_writable_fields(file_path, cur_pg)

	cur_fields: list[dict[str, Any]] = ui.session_state[state_key][cur_pg]
	ns = f"pdf_st_{fhash}_{cur_pg}"
	cursor_xp_k = f"{ns}_cxp"
	cursor_yp_k = f"{ns}_cyp"
	sel_idx_k = f"{ns}_sel_idx"
	sync_k = f"{ns}_a4_sync"

	# A4 Canlı Tuval JS->Python senkronizasyonu
	raw_sync = ui.text_input("", value="", key=sync_k, label_visibility="collapsed")
	if raw_sync.strip():
		changed, picked_idx = _sync_overlay_canvas_edits(cur_fields, raw_sync)
		if changed:
			try:
				ev_d = json.loads(raw_sync)
				if ev_d.get("op") == "cursor":
					ui.session_state[cursor_xp_k] = float(ev_d.get("xp", 18.0))
					ui.session_state[cursor_yp_k] = float(ev_d.get("yp", 78.0))
					ui.session_state[f"{ns}_man_xp"] = round(float(ev_d.get("xp", 18.0)), 1)
					ui.session_state[f"{ns}_man_yp"] = round(float(ev_d.get("yp", 78.0)), 1)
				elif ev_d.get("op") == "move" and picked_idx is not None:
					sf = cur_fields[picked_idx]
					sf_id = sf["id"]
					ui.session_state[f"{ns}_sel_x0_{sf_id}"] = round(float(sf["x0"]), 1)
					ui.session_state[f"{ns}_sel_y0_{sf_id}"] = round(float(sf["y0"]), 1)
					ui.session_state[f"{ns}_sel_w_{sf_id}"] = round(max(20.0, float(sf["x1"]) - float(sf["x0"])), 1)
					ui.session_state[f"{ns}_sel_h_{sf_id}"] = round(max(14.0, float(sf["y1"]) - float(sf["y0"])), 1)
					ui.session_state[f"pdf_fx0_{fhash}_{cur_pg}_{sf_id}"] = round(float(sf["x0"]), 1)
					ui.session_state[f"pdf_fy0_{fhash}_{cur_pg}_{sf_id}"] = round(float(sf["y0"]), 1)
					ui.session_state[f"pdf_fw_{fhash}_{cur_pg}_{sf_id}"] = round(max(20.0, float(sf["x1"]) - float(sf["x0"])), 1)
			except Exception:
				pass
			if picked_idx is not None:
				ui.session_state[sel_idx_k] = picked_idx
		ui.session_state[sync_k] = ""

	with tb2:
		ui.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
		if ui.button("🔍 Yazılabilir Alanları Tara", type="secondary", use_container_width=True, key=f"pdf_rescan_{fhash}_{cur_pg}"):
			ui.session_state[state_key][cur_pg] = scan_pdf_writable_fields(file_path, cur_pg)
			ui.session_state[flash_k] = f"✅ Sayfa {cur_pg + 1} yerel olarak tarandı: {len(ui.session_state[state_key][cur_pg])} yazılabilir alan bulundu."
			ui.rerun(scope="fragment")

	with tb3:
		ui.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
		if ui.button("⚡ Okul & Kişi Bilgilerini Otomatik Eşleştir", type="secondary", use_container_width=True, key=f"pdf_automatch_{fhash}_{cur_pg}"):
			hp_dict = None
			if active_person:
				try:
					with ctx_api["connect_db"]() as _c:
						_hr = _c.execute("SELECT * FROM health_profiles WHERE student_id=?", (int(active_person["id"]),)).fetchone()
						if _hr:
							hp_dict = dict(_hr)
				except Exception:
					pass
			cnt = auto_match_pdf_fields_locally(cur_fields, app_cfg, active_person, hp_dict)
			ui.session_state[flash_k] = f"⚡ {cnt} alan okul ayarları ve seçili kişi bilgileriyle otomatik dolduruldu!" if cnt else "Eşleşen boş alan bulunamadı; manuel girebilir veya yeni alan ekleyebilirsiniz."
			ui.rerun(scope="fragment")

	with tb4:
		ui.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
		show_guides = ui.checkbox("🟩 Kutu Numaralarını Göster", value=True, key=f"pdf_guides_{fhash}")

	if is_gai_on:
		with ui.expander("✨ Gemini AI PDF Form Tarayıcı & Akıllı Doldurma Yardımcısı (İsteğe Bağlı)", expanded=False):
			ui.caption("Taranmış resim PDF'lerinde bile yazı yazılabilecek kutuları görsel olarak bulabilir veya talimatınıza göre alanları doldurabilir.")
			ai_pdf_inst = ui.text_input(
				"💬 Gemini'ye Talimat (Opsiyonel)",
				placeholder="Örn: Boş alanları bul ve bugünkü tarih, okul adı ve istirahat/sevk açıklamasıyla doldur",
				key=f"pdf_ai_inst_{fhash}_{cur_pg}",
			)
			ga1, ga2 = ui.columns(2)
			with ga1:
				if ui.button("👁️ Gemini Vision ile Sayfadaki Boş Alanları Tara", type="primary", use_container_width=True, key=f"pdf_ai_scan_{fhash}_{cur_pg}"):
					raw_png = render_pdf_page_preview_png(file_path, cur_pg, [], show_guides=False, dpi=130)
					anon_ctx = ""
					if active_person:
						try:
							anon_ctx, _, _ = ctx_api["_build_person_anon_bundle"](int(active_person["id"]))
						except Exception:
							anon_ctx = ""
					school_ctx = f"Okul: {app_cfg.get('school_name', '')} | Hemşire: {app_cfg.get('nurse_name', '')} | Tarih: {date.today().strftime('%d.%m.%Y')}"
					with ui.spinner("Gemini sayfa görüntüsünü tarayıp yazılabilir kutuları tespit ediyor..."):
						ok_sc, sc_res, used_m = gai.scan_pdf_form_fields_with_gemini(
							api_key=gai_key,
							model=gai_model,
							page_png_bytes=raw_png,
							page_number=cur_pg + 1,
							school_context=school_ctx,
							anon_student_context=anon_ctx,
							user_instruction=ai_pdf_inst,
						)
					if ok_sc and sc_res.get("fields"):
						new_ai_fields = []
						for idx_a, af in enumerate(sc_res["fields"]):
							x0 = round(pw * af["x_pct"] / 100.0, 1)
							y0 = round(ph * af["y_pct"] / 100.0, 1)
							x1 = round(min(pw - 10.0, x0 + pw * af["w_pct"] / 100.0), 1)
							y1 = round(min(ph - 10.0, y0 + ph * af["h_pct"] / 100.0), 1)
							sval = str(af.get("suggested_value") or "")
							if active_person and "[ÖĞRENCİ_ADI]" in sval:
								sval = sval.replace("[ÖĞRENCİ_ADI]", str(active_person.get("full_name") or ""))
							new_ai_fields.append({
								"id": f"gai_{cur_pg}_{idx_a + 1}",
								"page": cur_pg,
								"label": af["label"],
								"field_type": af.get("field_type", "text"),
								"source": "gemini_vision",
								"widget_name": "",
								"x0": x0, "y0": y0, "x1": x1, "y1": y1,
								"value": sval,
								"font_family": FONT_FAMILY_OPTIONS[0],
								"font_size": int(af.get("font_size") or 10),
								"color": "lacivert",
								"clear_bg": False,
							})
						ui.session_state[state_key][cur_pg] = new_ai_fields
						ui.session_state[flash_k] = f"✨ Gemini Vision ({used_m}) Sayfa {cur_pg + 1} üzerinde {len(new_ai_fields)} yazılabilir alan buldu!"
						ui.rerun(scope="fragment")
					else:
						ui.error(f"Gemini Vision tarama hatası: {sc_res.get('error', 'Alan bulunamadı')}")

			with ga2:
				if ui.button("🪄 Mevcut Alanları Gemini ile Akıllı Doldur", type="secondary", disabled=not cur_fields, use_container_width=True, key=f"pdf_ai_fill_{fhash}_{cur_pg}"):
					if not ai_pdf_inst.strip():
						ui.warning("Lütfen üstteki kutuya alanların nasıl doldurulacağına dair kısa bir talimat yazın.")
					else:
						anon_ctx = ""
						if active_person:
							try:
								anon_ctx, _, _ = ctx_api["_build_person_anon_bundle"](int(active_person["id"]))
							except Exception:
								anon_ctx = ""
						school_ctx = f"Okul: {app_cfg.get('school_name', '')} | Hemşire: {app_cfg.get('nurse_name', '')} | Tarih: {date.today().strftime('%d.%m.%Y')}"
						f_payload = {"fields": [{"id": f["id"], "label": f["label"], "current_value": f["value"]} for f in cur_fields]}
						with ui.spinner("Gemini form alanları için Türkçe içerik hazırlıyor..."):
							ok_af, af_res, used_m = gai.assist_office_document_with_gemini(
								api_key=gai_key,
								model=gai_model,
								doc_type="pdf_fields",
								doc_payload=f_payload,
								instruction=ai_pdf_inst,
								school_context=school_ctx,
								anon_student_context=anon_ctx,
							)
						if ok_af and isinstance(af_res.get("field_values"), dict):
							fvals = af_res["field_values"]
							upd = 0
							for f in cur_fields:
								if f["id"] in fvals and str(fvals[f["id"]]).strip():
									nv = str(fvals[f["id"]])
									if active_person:
										nv = nv.replace("[ÖĞRENCİ_ADI]", str(active_person.get("full_name") or ""))
									f["value"] = nv
									upd += 1
							ui.session_state[flash_k] = f"✨ Gemini ({used_m}) {upd} form alanını doldurdu! {af_res.get('explanation', '')}"
							ui.rerun(scope="fragment")
						else:
							ui.error(f"Gemini doldurma hatası: {af_res.get('error', 'Ayrıştırılamadı')}")

	# 2 Sütunlu Çalışma Alanı: Sol Tıklanabilir & Sürüklenebilir A4 Kâğıt + Sağ Form / Tipografi / PNG İmza Paneli
	col_prev, col_edit = ui.columns([1.28, 1.02], gap="large", sticky_first_col=True)

	with col_edit:
		st_tab_fields, st_tab_add, st_tab_save = ui.tabs([
			f"✍️ Alanlar ({len(cur_fields)})",
			"🖋️ Yazı Tipi, Konum & PNG İmza",
			"💾 Kaydet & Dışa Aktar",
		])

		with st_tab_fields:
			if not cur_fields:
				ui.info("Bu sayfada otomatik form alanı tespit edilmedi. Soldaki A4 kâğıda tıklayıp **'🖋️ Yazı Tipi, Konum & PNG İmza'** sekmesinden yazı veya PNG imza ekleyebilirsiniz.")
			else:
				ui.caption("Soldaki A4 kâğıttaki kutulara tıklayıp sürükleyebilir veya buradan metin ve yazı tipi/punto ayarlarını düzenleyebilirsiniz.")
				for f_idx, f in enumerate(cur_fields):
					fid = f["id"]
					is_sig_f = (f.get("field_type") == "signature")
					src_tag = {"acroform": "Form", "gemini_vision": "AI", "underline": "Çizgi", "label_colon": "Etiket", "line": "Çizgi", "box": "Kutu", "manual": "Serbest", "signature": "PNG İmza"}.get(str(f.get("source")), "Alan")
					with ui.container(border=True):
						fc1, fc2 = ui.columns([3.4, 0.9])
						with fc1:
							if is_sig_f:
								ui.markdown(f"🖋️ **#{f_idx + 1} · {html.escape(str(f.get('label') or 'PNG İmza'))}** `[{src_tag}]`")
								ui.caption(f"Konum: X={f['x0']:.0f}pt, Y={f['y0']:.0f}pt · Genişlik={f['x1']-f['x0']:.0f}pt")
							else:
								new_v = ui.text_input(
									f"#{f_idx + 1} · {f['label']} [{src_tag}]",
									value=str(f.get("value") or ""),
									placeholder="Yazılacak metin (boş bırakılırsa yazılmaz)...",
									key=f"pdf_fv_{fhash}_{cur_pg}_{fid}",
								)
								f["value"] = new_v
						with fc2:
							ui.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
							if ui.button("🗑️", key=f"pdf_fdel_{fhash}_{cur_pg}_{fid}", help="Bu alanı kaldır", use_container_width=True):
								cur_fields.pop(f_idx)
								if ui.session_state.get(sel_idx_k) == f_idx:
									ui.session_state[sel_idx_k] = None
								ui.rerun(scope="fragment")
						if not is_sig_f:
							with ui.expander(f"⚙️ Yazı Tipi, Boyut & Konum (#{f_idx + 1})", expanded=(ui.session_state.get(sel_idx_k) == f_idx)):
								ft1, ft2, ft3 = ui.columns([1.8, 0.9, 1.1])
								ff_v = str(f.get("font_family") or FONT_FAMILY_OPTIONS[0])
								if ff_v not in FONT_FAMILY_OPTIONS:
									ff_v = FONT_FAMILY_OPTIONS[0]
								f["font_family"] = ft1.selectbox("Yazı Tipi", FONT_FAMILY_OPTIONS, index=FONT_FAMILY_OPTIONS.index(ff_v), key=f"pdf_fff_{fhash}_{cur_pg}_{fid}")
								f["font_size"] = int(ft2.number_input("Punto", min_value=6, max_value=36, value=int(f.get("font_size") or 10), step=1, key=f"pdf_fs_{fhash}_{cur_pg}_{fid}"))
								al_v = str(f.get("align") or ALIGN_OPTIONS[0])
								if al_v not in ALIGN_OPTIONS:
									al_v = ALIGN_OPTIONS[0]
								f["align"] = ft3.selectbox("Hiza", ALIGN_OPTIONS, index=ALIGN_OPTIONS.index(al_v), key=f"pdf_fal_{fhash}_{cur_pg}_{fid}")
								t1, t2, t3, t4 = ui.columns(4)
								f["x0"] = float(t1.number_input("Sol (X)", min_value=0.0, max_value=pw, value=float(f["x0"]), step=4.0, key=f"pdf_fx0_{fhash}_{cur_pg}_{fid}"))
								f["y0"] = float(t2.number_input("Üst (Y)", min_value=0.0, max_value=ph, value=float(f["y0"]), step=4.0, key=f"pdf_fy0_{fhash}_{cur_pg}_{fid}"))
								w_cur = max(20.0, float(f["x1"]) - float(f["x0"]))
								new_w = float(t3.number_input("Genişlik", min_value=20.0, max_value=pw, value=round(w_cur, 1), step=10.0, key=f"pdf_fw_{fhash}_{cur_pg}_{fid}"))
								f["x1"] = min(pw - 4.0, f["x0"] + new_w)
								f["bold"] = t4.checkbox("Kalın", value=bool(f.get("bold")), key=f"pdf_fb_{fhash}_{cur_pg}_{fid}")
								c1, c2 = ui.columns(2)
								col_opts = ["lacivert", "siyah", "kirmizi", "yesil"]
								cur_ci = col_opts.index(f.get("color", "lacivert")) if f.get("color") in col_opts else 0
								f["color"] = c1.selectbox("Renk", col_opts, index=cur_ci, key=f"pdf_fc_{fhash}_{cur_pg}_{fid}")
								f["clear_bg"] = c2.checkbox("Arkadaki noktaları/çizgiyi ört", value=bool(f.get("clear_bg")), key=f"pdf_fbg_{fhash}_{cur_pg}_{fid}")

		with st_tab_add:
			_render_overlay_and_signature_controls(
				fields=cur_fields,
				page_idx=cur_pg,
				pw=pw,
				ph=ph,
				evrak_dir=evrak_dir,
				ns=ns,
				flash_k=flash_k,
				cursor_xp_k=cursor_xp_k,
				cursor_yp_k=cursor_yp_k,
				sel_idx_k=sel_idx_k,
			)

		with st_tab_save:
			filled_total = sum(
				1 for flist in ui.session_state[state_key].values()
				for f in flist
				if (f.get("field_type") == "signature" or str(f.get("value") or "").strip())
			)
			ui.info(f"📝 Tüm sayfalarda doldurulmuş metin veya PNG imza içeren toplam **{filled_total}** alan kayda hazır.")
			new_copy_stem = ui.text_input(
				"Farklı Kaydet Dosya Adı (.pdf otomatik eklenir)",
				value=f"{file_path.stem}_imzali",
				key=f"pdf_save_stem_{fhash}",
			)
			sv1, sv2 = ui.columns(2)
			with sv1:
				if ui.button("📄 Yeni PDF Olarak Depoya Kaydet", type="primary", use_container_width=True, key=f"pdf_save_copy_{fhash}"):
					target_p = _unique_target_path(evrak_dir, new_copy_stem, ".pdf")
					save_filled_pdf(file_path, ui.session_state[state_key], target_p)
					ctx_api["get_or_create_document_cover"](target_p, force=True)
					ui.session_state[flash_k] = f"✅ Doldurulmuş & imzalı PDF `{target_p.name}` olarak Evrak Deposuna kaydedildi!"
					ui.rerun(scope="fragment")
			with sv2:
				if ui.button("💾 Mevcut PDF Üzerine Yaz", type="secondary", use_container_width=True, key=f"pdf_save_over_{fhash}"):
					save_filled_pdf(file_path, ui.session_state[state_key], file_path)
					ctx_api["get_or_create_document_cover"](file_path, force=True)
					ui.session_state[flash_k] = f"✅ `{file_path.name}` dosyası güncellendi!"
					ui.rerun(scope="fragment")

	with col_prev:
		try:
			bg_png = render_pdf_page_preview_png(file_path, cur_pg, [], show_guides=False, dpi=140)
			if bg_png:
				b64_bg = _b64.b64encode(bg_png).decode("ascii")
				_render_interactive_a4_overlay_stage(
					bg_img_b64=b64_bg,
					bg_mime="image/png",
					page_w=pw,
					page_h=ph,
					fields=cur_fields,
					stage_id=f"a4_stage_pdf_{fhash}_{cur_pg}",
					sync_key=sync_k,
					cursor_xp=float(ui.session_state.get(cursor_xp_k, 18.0)),
					cursor_yp=float(ui.session_state.get(cursor_yp_k, 78.0)),
					selected_idx=ui.session_state.get(sel_idx_k),
					show_guides=show_guides,
					badge_text=f"📄 A4 · Sayfa {cur_pg + 1} / {num_pages}",
				)
		except Exception as exc:
			ui.error(f"PDF önizleme oluşturulamadı: {exc}")


def _render_docx_studio(
	file_path: Path,
	fhash: str,
	evrak_dir: Path,
	ctx_api: dict[str, Any],
	app_cfg: dict[str, str],
	active_person: dict[str, Any] | None,
	is_gai_on: bool,
	gai_key: str,
	gai_model: str,
	flash_k: str,
) -> None:
	import base64 as _b64
	import pymupdf

	st_key = f"docx_studio_data_{fhash}"
	if st_key not in ui.session_state:
		paras, tables = _load_docx_structure(file_path)
		ui.session_state[st_key] = {
			"paras": paras,
			"tables": tables,
			"appended": [],
			"overlays": [],
		}
	st = ui.session_state[st_key]
	paras: list[dict[str, Any]] = st["paras"]
	tables: list[pd.DataFrame] = st["tables"]
	overlays: list[dict[str, Any]] = st.setdefault("overlays", [])

	# Üst Tipografi ve Canlı A4 Çalışma Kipi Şeridi
	tp1, tp2, tp3, tp4 = ui.columns([1.4, 0.9, 1.1, 1.6])
	with tp1:
		doc_font_family = ui.selectbox("🔤 Belge Yazı Tipi", FONT_FAMILY_OPTIONS, index=0, key=f"docx_ff_{fhash}")
	with tp2:
		doc_font_size = int(ui.number_input("📏 Punto (pt)", min_value=8, max_value=22, value=11, step=1, key=f"docx_fs_{fhash}"))
	with tp3:
		doc_align = ui.selectbox("📐 Hizalama", ALIGN_OPTIONS, index=0, key=f"docx_al_{fhash}")
	with tp4:
		live_mode = ui.radio(
			"Canlı A4 Modu",
			["🖋️ A4 Konum & PNG İmza", "🖱️ A4 Üstünde Metin Düzenle"],
			horizontal=True,
			key=f"docx_live_mode_{fhash}",
		)

	# Hızlı Bul & Değiştir + Okul/Kişi Yer Tutucu Şeridi (%100 Çevrimdışı)
	with ui.expander("🔎 Hızlı Bul & Değiştir + Okul / Öğrenci Bilgilerini Otomatik Yerleştir", expanded=False):
		fr1, fr2, fr3 = ui.columns([1.8, 1.8, 1.2])
		find_txt = fr1.text_input("Aranacak Metin", placeholder="Örn: [TARİH] veya 2025-2026", key=f"docx_find_{fhash}")
		rep_txt = fr2.text_input("Yeni Metin", placeholder="Örn: 21.11.2026", key=f"docx_rep_{fhash}")
		with fr3:
			ui.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
			if ui.button("🔄 Tümünde Değiştir", type="primary", use_container_width=True, key=f"docx_do_rep_{fhash}"):
				if find_txt:
					hit = 0
					for p in paras:
						if find_txt in p["text"]:
							p["text"] = p["text"].replace(find_txt, rep_txt)
							ui.session_state[f"docx_p_{fhash}_{p['index']}"] = p["text"]
							hit += 1
					for df in tables:
						if not df.empty:
							for col in df.columns:
								mask = df[col].astype(str).str.contains(re.escape(find_txt), na=False)
								if mask.any():
									df[col] = df[col].astype(str).str.replace(find_txt, rep_txt, regex=False)
									hit += int(mask.sum())
					ui.session_state[flash_k] = f"🔄 {hit} yerde değiştirme uygulandı."
					ui.rerun(scope="fragment")

		if ui.button("⚡ Yer Tutucuları ([OKUL_ADI], [TARIH], [HEMSIRE_ADI], [OGRENCI_ADI], [SINIF]) Otomatik Doldur", use_container_width=True, key=f"docx_fill_ph_{fhash}"):
			ph_map = {
				"[OKUL_ADI]": str(app_cfg.get("school_name") or ""),
				"[HEMSIRE_ADI]": str(app_cfg.get("nurse_name") or ""),
				"[MUDUR_ADI]": str(app_cfg.get("principal_name") or ""),
				"[TARIH]": date.today().strftime("%d.%m.%Y"),
				"[TARİH]": date.today().strftime("%d.%m.%Y"),
				"[OGRENCI_ADI]": str(active_person.get("full_name") or "") if active_person else "",
				"[ÖĞRENCİ_ADI]": str(active_person.get("full_name") or "") if active_person else "",
				"[SINIF]": str(active_person.get("class_name") or "") if active_person else "",
			}
			hit = 0
			for p in paras:
				for k_ph, v_ph in ph_map.items():
					if v_ph and k_ph in p["text"]:
						p["text"] = p["text"].replace(k_ph, v_ph)
						ui.session_state[f"docx_p_{fhash}_{p['index']}"] = p["text"]
						hit += 1
			for df in tables:
				for col in df.columns:
					for k_ph, v_ph in ph_map.items():
						if v_ph:
							s_col = df[col].astype(str)
							if s_col.str.contains(re.escape(k_ph), na=False).any():
								df[col] = s_col.str.replace(k_ph, v_ph, regex=False)
								hit += 1
			ui.session_state[flash_k] = f"⚡ {hit} yer tutucu otomatik dolduruldu!" if hit else "Belgede eşleşen köşeli parantezli yer tutucu bulunamadı."
			ui.rerun(scope="fragment")

	if is_gai_on:
		with ui.expander("✨ Gemini AI Word Belge Düzenleme Yardımcısı (İsteğe Bağlı)", expanded=False):
			w_inst = ui.text_input(
				"💬 Belgede ne düzenlensin / eklensin?",
				placeholder="Örn: Dili resmi MEB yazışma üslubuna çevir | Eksik imza ve tarih satırı ekle | Tabloyu seçili öğrenci bilgisine göre doldur",
				key=f"docx_ai_inst_{fhash}",
			)
			if ui.button("✨ Gemini ile Düzenleme Önerisi Hazırla", type="primary", use_container_width=True, key=f"docx_ai_run_{fhash}"):
				if not w_inst.strip():
					ui.warning("Lütfen bir talimat yazın.")
				else:
					anon_ctx = ""
					if active_person:
						try:
							anon_ctx, _, _ = ctx_api["_build_person_anon_bundle"](int(active_person["id"]))
						except Exception:
							anon_ctx = ""
					payload = {
						"paragraphs": [{"index": p["index"], "text": p["text"]} for p in paras if p["text"].strip()][:60],
						"tables": [
							{"table_index": ti, "headers": list(df.columns), "rows": df.head(25).values.tolist()}
							for ti, df in enumerate(tables[:4])
						],
					}
					school_ctx = f"Okul: {app_cfg.get('school_name', '')} | Hemşire: {app_cfg.get('nurse_name', '')} | Tarih: {date.today().strftime('%d.%m.%Y')}"
					with ui.spinner("Gemini Word belgesini inceliyor..."):
						ok_dw, dw_res, used_m = gai.assist_office_document_with_gemini(
							api_key=gai_key,
							model=gai_model,
							doc_type="docx",
							doc_payload=payload,
							instruction=w_inst,
							school_context=school_ctx,
							anon_student_context=anon_ctx,
						)
					if ok_dw:
						ui.session_state[f"docx_ai_prop_{fhash}"] = {"res": dw_res, "model": used_m}
						ui.rerun(scope="fragment")
					else:
						ui.error(f"Gemini hatası: {dw_res.get('error')}")

			prop = ui.session_state.get(f"docx_ai_prop_{fhash}")
			if isinstance(prop, dict):
				pr = prop["res"]
				ui.info(f"🤖 **Gemini Önerisi ({prop['model']}):** {pr.get('explanation', '')}")
				p_edits = pr.get("paragraph_edits") or []
				p_adds = pr.get("append_paragraphs") or []
				t_edits = pr.get("table_edits") or []
				ui.caption(f"Önerilen değişiklikler: **{len(p_edits)}** paragraf düzenlemesi · **{len(p_adds)}** yeni paragraf · **{len(t_edits)}** tablo hücresi")
				ap_c1, ap_c2 = ui.columns(2)
				with ap_c1:
					if ui.button("✅ Gemini Önerilerini Belgeye Uygula", type="primary", use_container_width=True, key=f"docx_ai_apply_{fhash}"):
						p_by_idx = {p["index"]: p for p in paras}
						for pe in p_edits:
							try:
								pi = int(pe["index"])
								nt = str(pe.get("new_text") or "")
								if active_person:
									nt = nt.replace("[ÖĞRENCİ_ADI]", str(active_person.get("full_name") or ""))
								if pi in p_by_idx:
									p_by_idx[pi]["text"] = nt
									ui.session_state[f"docx_p_{fhash}_{pi}"] = nt
							except Exception:
								pass
						for pa in p_adds:
							if str(pa).strip():
								nt = str(pa).strip()
								if active_person:
									nt = nt.replace("[ÖĞRENCİ_ADI]", str(active_person.get("full_name") or ""))
								st["appended"].append(nt)
						for te in t_edits:
							try:
								ti, ri, ci = int(te["table_index"]), int(te["row_index"]), int(te["col_index"])
								if 0 <= ti < len(tables) and 0 <= ri < len(tables[ti]) and 0 <= ci < len(tables[ti].columns):
									tables[ti].iat[ri, ci] = str(te.get("new_text") or "")
							except Exception:
								pass
						ui.session_state.pop(f"docx_ai_prop_{fhash}", None)
						ui.session_state[flash_k] = "✅ Gemini önerileri düzenleyiciye aktarıldı!"
						ui.rerun(scope="fragment")
				with ap_c2:
					if ui.button("✖ Öneriyi İptal Et", use_container_width=True, key=f"docx_ai_dismiss_{fhash}"):
						ui.session_state.pop(f"docx_ai_prop_{fhash}", None)
						ui.rerun(scope="fragment")

	pw, ph = 595.0, 842.0
	ns = f"docx_st_{fhash}"
	cursor_xp_k = f"{ns}_cxp"
	cursor_yp_k = f"{ns}_cyp"
	sel_idx_k = f"{ns}_sel_idx"
	sync_k = f"{ns}_a4_sync"
	inline_sync_k = f"{ns}_inline_sync"

	raw_inline = ui.text_input("", value="", key=inline_sync_k, label_visibility="collapsed")
	if raw_inline.strip():
		try:
			iev = json.loads(raw_inline)
			if isinstance(iev, dict) and iev.get("op") == "p_edit":
				pi_ed = int(iev["idx"])
				new_t = str(iev.get("text") or "")
				for p in paras:
					if int(p["index"]) == pi_ed:
						p["text"] = new_t
						ui.session_state[f"docx_p_{fhash}_{pi_ed}"] = new_t
						break
		except Exception:
			pass
		ui.session_state[inline_sync_k] = ""

	raw_sync = ui.text_input("", value="", key=sync_k, label_visibility="collapsed")
	if raw_sync.strip():
		changed, picked_idx = _sync_overlay_canvas_edits(overlays, raw_sync)
		if changed:
			try:
				ev_d = json.loads(raw_sync)
				if ev_d.get("op") == "cursor":
					ui.session_state[cursor_xp_k] = float(ev_d.get("xp", 18.0))
					ui.session_state[cursor_yp_k] = float(ev_d.get("yp", 78.0))
					ui.session_state[f"{ns}_man_xp"] = round(float(ev_d.get("xp", 18.0)), 1)
					ui.session_state[f"{ns}_man_yp"] = round(float(ev_d.get("yp", 78.0)), 1)
				elif ev_d.get("op") == "move" and picked_idx is not None:
					sf = overlays[picked_idx]
					sf_id = sf["id"]
					ui.session_state[f"{ns}_sel_x0_{sf_id}"] = round(float(sf["x0"]), 1)
					ui.session_state[f"{ns}_sel_y0_{sf_id}"] = round(float(sf["y0"]), 1)
					ui.session_state[f"{ns}_sel_w_{sf_id}"] = round(max(20.0, float(sf["x1"]) - float(sf["x0"])), 1)
					ui.session_state[f"{ns}_sel_h_{sf_id}"] = round(max(14.0, float(sf["y1"]) - float(sf["y0"])), 1)
			except Exception:
				pass
			if picked_idx is not None:
				ui.session_state[sel_idx_k] = picked_idx
		ui.session_state[sync_k] = ""

	ui.markdown(
		f"<style>.revir-field:has(input[data-key=\"{html.escape(inline_sync_k, quote=True)}\"]){{display:none!important}}</style>",
		unsafe_allow_html=True,
	)

	# 2 Sütunlu Canlı A4 Stüdyo
	col_prev, col_edit = ui.columns([1.28, 1.02], gap="large", sticky_first_col=True)

	with col_edit:
		t_p, t_t, t_ov, t_s = ui.tabs([
			f"📝 Metin ({sum(1 for p in paras if p['text'].strip())})",
			f"📊 Tablolar ({len(tables)})",
			f"🖋️ Konum & PNG İmza ({len(overlays)})",
			"💾 Kaydet & PDF",
		])

		with t_p:
			non_empty = [p for p in paras if p["text"].strip()]
			if not non_empty:
				ui.info("Bu Word belgesinde metin paragrafı bulunmuyor; aşağıdan yeni paragraf ekleyebilir veya Tablolar sekmesini düzenleyebilirsiniz.")
			for p in non_empty[:80]:
				pi = p["index"]
				h_val = 90 if len(p["text"]) > 130 else 62
				p["text"] = ui.text_area(
					f"Paragraf #{pi + 1} ({p['style']})",
					value=p["text"],
					height=h_val,
					key=f"docx_p_{fhash}_{pi}",
				)
			ui.divider()
			ui.markdown("##### ➕ Belge Sonuna Yeni Paragraf Ekle")
			new_p_txt = ui.text_area("Eklenecek Metin", placeholder="Belgenin sonuna eklemek istediğiniz paragrafı yazın...", height=76, key=f"docx_newp_{fhash}")
			if ui.button("➕ Paragrafı Belgeye Ekle", key=f"docx_addp_{fhash}"):
				if new_p_txt.strip():
					st["appended"].append(new_p_txt.strip())
					ui.session_state[f"docx_newp_{fhash}"] = ""
					ui.session_state[flash_k] = "➕ Yeni paragraf belge sonuna eklendi."
					ui.rerun(scope="fragment")
			if st["appended"]:
				ui.caption(f"Belge sonuna eklenecek **{len(st['appended'])}** yeni paragraf:")
				for ai, at in enumerate(st["appended"]):
					ac1, ac2 = ui.columns([4.5, 0.7])
					ac1.markdown(f"• {html.escape(at)}")
					if ac2.button("✖", key=f"docx_rm_app_{fhash}_{ai}"):
						st["appended"].pop(ai)
						ui.rerun(scope="fragment")

		with t_t:
			if not tables:
				ui.info("Bu Word belgesinde tablo bulunmuyor.")
			else:
				for ti, df in enumerate(tables):
					ui.markdown(f"##### 📊 Tablo #{ti + 1}")
					tables[ti] = ui.data_editor(
						df,
						num_rows="dynamic",
						use_container_width=True,
						key=f"docx_tbl_{fhash}_{ti}",
					)

		with t_ov:
			_render_overlay_and_signature_controls(
				fields=overlays,
				page_idx=0,
				pw=pw,
				ph=ph,
				evrak_dir=evrak_dir,
				ns=ns,
				flash_k=flash_k,
				cursor_xp_k=cursor_xp_k,
				cursor_yp_k=cursor_yp_k,
				sel_idx_k=sel_idx_k,
			)

		with t_s:
			new_stem = ui.text_input("Farklı Kaydet Dosya Adı", value=f"{file_path.stem}_duzenlenmis", key=f"docx_sv_stem_{fhash}")
			sig_ovs = [ov for ov in overlays if ov.get("field_type") == "signature"]
			txt_ovs = [ov for ov in overlays if ov.get("field_type") != "signature" and str(ov.get("value") or "").strip()]

			s1, s2 = ui.columns(2)
			with s1:
				if ui.button("💾 Word (.docx) Üzerine Kaydet", type="primary", use_container_width=True, key=f"docx_sv_over_{fhash}"):
					edited_map = {p["index"]: p["text"] for p in paras}
					tbl_map = {ti: df for ti, df in enumerate(tables)}
					_save_docx_structure(
						file_path, file_path, edited_map, st["appended"], tbl_map,
						font_family=doc_font_family, font_size_pt=float(doc_font_size), default_align=doc_align,
						signature_overlays=sig_ovs, extra_text_overlays=txt_ovs,
					)
					ctx_api["get_or_create_document_cover"](file_path, force=True)
					ui.session_state.pop(st_key, None)
					ui.session_state[flash_k] = f"✅ `{file_path.name}` (yazı tipi + imzalarla) başarıyla güncellendi!"
					ui.rerun(scope="fragment")
			with s2:
				if ui.button("📄 Yeni Word Kopyası (.docx)", type="secondary", use_container_width=True, key=f"docx_sv_copy_{fhash}"):
					target_p = _unique_target_path(evrak_dir, new_stem, ".docx")
					edited_map = {p["index"]: p["text"] for p in paras}
					tbl_map = {ti: df for ti, df in enumerate(tables)}
					_save_docx_structure(
						file_path, target_p, edited_map, st["appended"], tbl_map,
						font_family=doc_font_family, font_size_pt=float(doc_font_size), default_align=doc_align,
						signature_overlays=sig_ovs, extra_text_overlays=txt_ovs,
					)
					ctx_api["get_or_create_document_cover"](target_p, force=True)
					ui.session_state[flash_k] = f"✅ `{target_p.name}` Evrak Deposuna yeni kopya olarak kaydedildi!"
					ui.rerun(scope="fragment")

			if ui.button("📑 Tam Konumlu İmzalı & Şekilli A4 PDF Olarak Kaydet (.pdf)", type="primary", use_container_width=True, key=f"docx_sv_pdf_{fhash}"):
				inner_h = _docx_to_inner_html(file_path, paras_override=paras, tables_override=tables, appended_paras=st["appended"], default_align=doc_align)
				ovs_by_pg: dict[int, list[dict[str, Any]]] = {}
				for ov in overlays:
					pg_i = int(ov.get("page") or 0)
					ovs_by_pg.setdefault(pg_i, []).append(ov)
				pdf_bytes = render_html_body_to_pdf_bytes(
					inner_h,
					font_family=doc_font_family,
					font_size_pt=float(doc_font_size),
					landscape=False,
					overlays_by_page=ovs_by_pg,
				)
				pdf_target = _unique_target_path(evrak_dir, new_stem, ".pdf")
				pdf_target.write_bytes(pdf_bytes)
				ctx_api["get_or_create_document_cover"](pdf_target, force=True)
				ui.session_state[flash_k] = f"📑 Tam konumlu PNG imzalı A4 PDF `{pdf_target.name}` olarak Evrak Deposuna kaydedildi!"
				ui.rerun(scope="fragment")

	with col_prev:
		if "Doğrudan" in live_mode or "Metin Düzenle" in live_mode:
			ff_css = _css_font_family(doc_font_family)
			al_css = {"Sola Hizalı": "left", "Ortala": "center", "Sağa Hizalı": "right"}.get(doc_align, "left")
			p_nodes: list[str] = []
			for p in paras:
				txt = str(p.get("text") or "").strip()
				if not txt:
					continue
				pi = int(p["index"])
				st_l = str(p.get("style") or "").lower()
				tag = "h1" if ("heading 1" in st_l or "title" in st_l) else ("h2" if "heading 2" in st_l else ("h3" if "heading 3" in st_l else "p"))
				p_nodes.append(
					f'<{tag} contenteditable="true" data-live-pidx="{pi}" '
					f'style="outline:none;border-radius:4px;padding:2px 4px;margin:0 0 6px 0;transition:background 0.15s;" '
					f'onfocus="this.style.background=\'rgba(13,148,136,0.08)\';this.style.boxShadow=\'inset 0 0 0 1px #0d9488\';" '
					f'onblur="this.style.background=\'transparent\';this.style.boxShadow=\'none\';if(window.__revirFieldChange){{window.__revirFieldChange(\'{html.escape(inline_sync_k, quote=True)}\', JSON.stringify({{op:\'p_edit\',idx:{pi},text:this.innerText}}));}}">'
					f'{html.escape(txt)}</{tag}>'
				)
			for ap in st["appended"]:
				if str(ap).strip():
					p_nodes.append(f"<p>{html.escape(str(ap).strip())}</p>")
			for df in tables:
				if df is None or df.empty:
					continue
				ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
				trs = []
				for _, r in df.head(40).iterrows():
					cells_html = "".join(
						"<td>{}</td>".format(html.escape(str(v) if not pd.isna(v) else ""))
						for v in r.tolist()
					)
					trs.append(f"<tr>{cells_html}</tr>")
				p_nodes.append(f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>")

			ui.markdown(
				f"<div class='revir-a4-stage' style='padding:12px;'>"
				f"<div style='display:flex;justify-content:space-between;width:100%;max-width:780px;align-items:center;'>"
				f"<span class='revir-a4-page-badge'>🖱️ A4 Kâğıt Üstünde Canlı Düzenleme</span>"
				f"<span style='font-size:11px;font-weight:700;color:#475569;'>Paragraflara tıklayıp doğrudan yazabilirsiniz</span>"
				f"</div>"
				f"<div class='revir-a4-sheet' style='font-family:{ff_css};font-size:{doc_font_size}pt;text-align:{al_css};'>{''.join(p_nodes)}</div>"
				f"</div>",
				unsafe_allow_html=True,
			)
		else:
			inner_h = _docx_to_inner_html(file_path, paras_override=paras, tables_override=tables, appended_paras=st["appended"], default_align=doc_align)
			raw_pdf_bytes = render_html_body_to_pdf_bytes(inner_h, font_family=doc_font_family, font_size_pt=float(doc_font_size), landscape=False)
			doc_v = pymupdf.open(stream=raw_pdf_bytes, filetype="pdf")
			try:
				n_pgs = max(1, len(doc_v))
				cur_pg_idx = 0
				if n_pgs > 1:
					sel_dp = ui.selectbox("📄 A4 Sayfası", [f"Sayfa {i + 1} / {n_pgs}" for i in range(n_pgs)], index=0, key=f"docx_pg_{fhash}")
					cur_pg_idx = min(n_pgs - 1, int(sel_dp.split()[1]) - 1)
				pg_obj = doc_v[cur_pg_idx]
				pw, ph = float(pg_obj.rect.width), float(pg_obj.rect.height)
				bg_png = pg_obj.get_pixmap(dpi=135).tobytes("png")
			finally:
				doc_v.close()
			b64_bg = _b64.b64encode(bg_png).decode("ascii")
			pg_overlays = [ov for ov in overlays if int(ov.get("page") or 0) == cur_pg_idx]
			_render_interactive_a4_overlay_stage(
				bg_img_b64=b64_bg,
				bg_mime="image/png",
				page_w=pw,
				page_h=ph,
				fields=overlays,
				stage_id=f"a4_stage_docx_{fhash}_{cur_pg_idx}",
				sync_key=sync_k,
				cursor_xp=float(ui.session_state.get(cursor_xp_k, 62.0)),
				cursor_yp=float(ui.session_state.get(cursor_yp_k, 80.0)),
				selected_idx=ui.session_state.get(sel_idx_k),
				show_guides=True,
				badge_text=f"📄 Word A4 Canlı Sahne · Sayfa {cur_pg_idx + 1} / {n_pgs}",
			)


def _render_xlsx_studio(
	file_path: Path,
	fhash: str,
	evrak_dir: Path,
	ctx_api: dict[str, Any],
	app_cfg: dict[str, str],
	active_person: dict[str, Any] | None,
	is_gai_on: bool,
	gai_key: str,
	gai_model: str,
	flash_k: str,
) -> None:
	import base64 as _b64
	import openpyxl
	import pymupdf

	wb_m = openpyxl.load_workbook(file_path, read_only=True)
	sheet_names = list(wb_m.sheetnames)
	wb_m.close()

	sh1, sh2, sh3, sh4 = ui.columns([1.8, 1.5, 1.1, 1.2])
	with sh1:
		sel_sheet = ui.selectbox("📑 Excel Sayfası (Sheet)", sheet_names, index=0, key=f"xlsx_sh_{fhash}")
	with sh2:
		xl_ff = ui.selectbox("🔤 Yazı Tipi", FONT_FAMILY_OPTIONS, index=0, key=f"xlsx_ff_{fhash}")
	with sh3:
		xl_fs = int(ui.number_input("📏 Punto", min_value=7, max_value=18, value=10, step=1, key=f"xlsx_fs_{fhash}"))
	with sh4:
		xl_orient = ui.selectbox("📐 A4 Yönü", ["A4 Yatay", "A4 Dikey"], index=0, key=f"xlsx_orient_{fhash}")
	is_land = ("Yatay" in xl_orient)

	new_sh_nm = ui.text_input("➕ Yeni Sayfa Adı (Opsiyonel)", placeholder="Boş bırakılırsa seçili sayfa güncellenir", key=f"xlsx_newsh_{fhash}")

	df_key = f"xlsx_studio_df_{fhash}_{sel_sheet}"
	if df_key not in ui.session_state:
		ui.session_state[df_key] = _load_xlsx_sheet_df(file_path, sel_sheet)
	ov_key = f"xlsx_studio_ov_{fhash}_{sel_sheet}"
	if ov_key not in ui.session_state:
		ui.session_state[ov_key] = []

	cur_df: pd.DataFrame = ui.session_state[df_key]
	overlays: list[dict[str, Any]] = ui.session_state[ov_key]

	pw, ph = (842.0, 595.0) if is_land else (595.0, 842.0)
	ns = f"xlsx_st_{fhash}_{sel_sheet}"
	cursor_xp_k = f"{ns}_cxp"
	cursor_yp_k = f"{ns}_cyp"
	sel_idx_k = f"{ns}_sel_idx"
	sync_k = f"{ns}_a4_sync"

	raw_sync = ui.text_input("", value="", key=sync_k, label_visibility="collapsed")
	if raw_sync.strip():
		changed, picked_idx = _sync_overlay_canvas_edits(overlays, raw_sync)
		if changed:
			try:
				ev_d = json.loads(raw_sync)
				if ev_d.get("op") == "cursor":
					ui.session_state[cursor_xp_k] = float(ev_d.get("xp", 70.0))
					ui.session_state[cursor_yp_k] = float(ev_d.get("yp", 82.0))
					ui.session_state[f"{ns}_man_xp"] = round(float(ev_d.get("xp", 70.0)), 1)
					ui.session_state[f"{ns}_man_yp"] = round(float(ev_d.get("yp", 82.0)), 1)
				elif ev_d.get("op") == "move" and picked_idx is not None:
					sf = overlays[picked_idx]
					sf_id = sf["id"]
					ui.session_state[f"{ns}_sel_x0_{sf_id}"] = round(float(sf["x0"]), 1)
					ui.session_state[f"{ns}_sel_y0_{sf_id}"] = round(float(sf["y0"]), 1)
					ui.session_state[f"{ns}_sel_w_{sf_id}"] = round(max(20.0, float(sf["x1"]) - float(sf["x0"])), 1)
					ui.session_state[f"{ns}_sel_h_{sf_id}"] = round(max(14.0, float(sf["y1"]) - float(sf["y0"])), 1)
			except Exception:
				pass
			if picked_idx is not None:
				ui.session_state[sel_idx_k] = picked_idx
		ui.session_state[sync_k] = ""

	if is_gai_on:
		with ui.expander("✨ Gemini AI Excel Tablo Yardımcısı (Veri Girme & Düzenleme)", expanded=False):
			x_inst = ui.text_input(
				"💬 Tabloda ne yapılsın?",
				placeholder="Örn: Eksik durum hücrelerini sınıflandır | Yeni özet satırı ekle | Seçili öğrenciye göre 1 yeni kayıt satırı ekle",
				key=f"xlsx_ai_inst_{fhash}_{sel_sheet}",
			)
			if ui.button("✨ Gemini ile Tablo Önerisi Hazırla", type="primary", use_container_width=True, key=f"xlsx_ai_run_{fhash}_{sel_sheet}"):
				if not x_inst.strip():
					ui.warning("Lütfen bir talimat yazın.")
				else:
					anon_ctx = ""
					if active_person:
						try:
							anon_ctx, _, _ = ctx_api["_build_person_anon_bundle"](int(active_person["id"]))
						except Exception:
							anon_ctx = ""
					payload = {
						"sheet": sel_sheet,
						"columns": list(cur_df.columns),
						"rows": cur_df.head(45).to_dict(orient="records"),
					}
					school_ctx = f"Okul: {app_cfg.get('school_name', '')} | Hemşire: {app_cfg.get('nurse_name', '')} | Tarih: {date.today().strftime('%d.%m.%Y')}"
					with ui.spinner("Gemini Excel sayfasını inceliyor..."):
						ok_x, x_res, used_m = gai.assist_office_document_with_gemini(
							api_key=gai_key,
							model=gai_model,
							doc_type="xlsx",
							doc_payload=payload,
							instruction=x_inst,
							school_context=school_ctx,
							anon_student_context=anon_ctx,
						)
					if ok_x:
						ui.session_state[f"xlsx_ai_prop_{fhash}_{sel_sheet}"] = {"res": x_res, "model": used_m}
						ui.rerun(scope="fragment")
					else:
						ui.error(f"Gemini hatası: {x_res.get('error')}")

			xprop = ui.session_state.get(f"xlsx_ai_prop_{fhash}_{sel_sheet}")
			if isinstance(xprop, dict):
				xr = xprop["res"]
				ui.info(f"🤖 **Gemini Önerisi ({xprop['model']}):** {xr.get('explanation', '')}")
				c_edits = xr.get("cell_edits") or []
				a_rows = xr.get("append_rows") or []
				ui.caption(f"Önerilen: **{len(c_edits)}** hücre güncellemesi · **{len(a_rows)}** yeni satır")
				xa1, xa2 = ui.columns(2)
				with xa1:
					if ui.button("✅ Tabloya Uygula", type="primary", use_container_width=True, key=f"xlsx_ai_apply_{fhash}_{sel_sheet}"):
						df_mod = cur_df.copy()
						for ce in c_edits:
							try:
								ri = int(ce["row_index"])
								cn = str(ce["col_name"])
								if cn in df_mod.columns and 0 <= ri < len(df_mod):
									df_mod.at[df_mod.index[ri], cn] = str(ce.get("value") or "")
							except Exception:
								pass
						if a_rows:
							extra_rows = []
							for ar in a_rows:
								if isinstance(ar, dict):
									extra_rows.append({c: str(ar.get(c, "")) for c in df_mod.columns})
							if extra_rows:
								df_mod = pd.concat([df_mod, pd.DataFrame(extra_rows)], ignore_index=True)
						ui.session_state[df_key] = df_mod
						ui.session_state.pop(f"xlsx_ai_prop_{fhash}_{sel_sheet}", None)
						ui.session_state[flash_k] = "✅ Gemini önerileri Excel tablosuna uygulandı!"
						ui.rerun(scope="fragment")
				with xa2:
					if ui.button("✖ İptal", use_container_width=True, key=f"xlsx_ai_dismiss_{fhash}_{sel_sheet}"):
						ui.session_state.pop(f"xlsx_ai_prop_{fhash}_{sel_sheet}", None)
						ui.rerun(scope="fragment")

	xt1, xt2, xt3 = ui.tabs([
		"📊 Tablo Hücrelerini Düzenle",
		f"📄 Canlı A4 & PNG İmza Stüdyosu ({len(overlays)})",
		"💾 Kaydet & PDF İndir",
	])

	with xt1:
		ui.caption("💡 Hücreye tıklayıp doğrudan yazabilirsiniz; `Enter` tuşu bir alt satıra geçer, alttan yeni satır ekleyebilirsiniz.")
		edited_df = ui.data_editor(
			cur_df,
			num_rows="dynamic",
			use_container_width=True,
			key=f"xlsx_ed_{fhash}_{sel_sheet}",
		)
		ui.session_state[df_key] = edited_df

	with xt2:
		df_live = ui.session_state[df_key]
		xl_inner = f"<h1>{html.escape(file_path.stem)} — {html.escape(sel_sheet)}</h1>"
		if df_live is not None and not df_live.empty:
			ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in df_live.columns)
			trs = []
			for _, r in df_live.head(60).iterrows():
				cells_html = "".join(
					"<td>{}</td>".format(html.escape(str(v) if not pd.isna(v) else ""))
					for v in r.tolist()
				)
				trs.append(f"<tr>{cells_html}</tr>")
			xl_inner += f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>"
		else:
			xl_inner += "<p>Sayfa boş.</p>"

		c_xl_prev, c_xl_ctrl = ui.columns([1.3, 1.0], gap="large")
		with c_xl_ctrl:
			_render_overlay_and_signature_controls(
				fields=overlays,
				page_idx=0,
				pw=pw,
				ph=ph,
				evrak_dir=evrak_dir,
				ns=ns,
				flash_k=flash_k,
				cursor_xp_k=cursor_xp_k,
				cursor_yp_k=cursor_yp_k,
				sel_idx_k=sel_idx_k,
			)
		with c_xl_prev:
			pdf_xl = render_html_body_to_pdf_bytes(xl_inner, font_family=xl_ff, font_size_pt=float(xl_fs), landscape=is_land, margin_pt=28.0)
			doc_xl = pymupdf.open(stream=pdf_xl, filetype="pdf")
			try:
				bg_png = doc_xl[0].get_pixmap(dpi=130).tobytes("png")
			finally:
				doc_xl.close()
			_render_interactive_a4_overlay_stage(
				bg_img_b64=_b64.b64encode(bg_png).decode("ascii"),
				bg_mime="image/png",
				page_w=pw,
				page_h=ph,
				fields=overlays,
				stage_id=f"a4_stage_xlsx_{fhash}_{sel_sheet}",
				sync_key=sync_k,
				cursor_xp=float(ui.session_state.get(cursor_xp_k, 70.0)),
				cursor_yp=float(ui.session_state.get(cursor_yp_k, 82.0)),
				selected_idx=ui.session_state.get(sel_idx_k),
				show_guides=True,
				badge_text=f"📊 Excel {xl_orient} · Sayfa 1",
			)

	with xt3:
		edited_df = ui.session_state[df_key]
		sig_ovs = [ov for ov in overlays if ov.get("field_type") == "signature"]
		new_stem = ui.text_input("Farklı Kaydet Dosya Adı", value=f"{file_path.stem}_duzenlenmis", key=f"xlsx_sv_stem_{fhash}")
		sv1, sv2 = ui.columns(2)
		with sv1:
			if ui.button("💾 Mevcut Excel Dosyasının Üzerine Kaydet", type="primary", use_container_width=True, key=f"xlsx_sv_over_{fhash}"):
				_save_xlsx_sheet_df(
					file_path, file_path, sel_sheet, edited_df, new_sheet_title=new_sh_nm,
					font_family=xl_ff, font_size_pt=float(xl_fs), landscape=is_land, signature_overlays=sig_ovs,
				)
				ctx_api["get_or_create_document_cover"](file_path, force=True)
				ui.session_state[flash_k] = f"✅ `{file_path.name}` ({new_sh_nm.strip() or sel_sheet}) kaydedildi!"
				ui.rerun(scope="fragment")
		with sv2:
			if ui.button("📊 Yeni Kopya Olarak Depoya Kaydet (.xlsx)", type="secondary", use_container_width=True, key=f"xlsx_sv_copy_{fhash}"):
				target_p = _unique_target_path(evrak_dir, new_stem, ".xlsx")
				_save_xlsx_sheet_df(
					file_path, target_p, sel_sheet, edited_df, new_sheet_title=new_sh_nm,
					font_family=xl_ff, font_size_pt=float(xl_fs), landscape=is_land, signature_overlays=sig_ovs,
				)
				ctx_api["get_or_create_document_cover"](target_p, force=True)
				ui.session_state[flash_k] = f"✅ `{target_p.name}` Evrak Deposuna kaydedildi!"
				ui.rerun(scope="fragment")

		if ui.button("📑 Tablo & İmzaları A4 PDF Olarak Kaydet (.pdf)", type="primary", use_container_width=True, key=f"xlsx_sv_pdf_{fhash}"):
			xl_inner = f"<h1>{html.escape(file_path.stem)} — {html.escape(sel_sheet)}</h1>"
			if edited_df is not None and not edited_df.empty:
				ths = "".join(f"<th>{html.escape(str(c))}</th>" for c in edited_df.columns)
				trs = []
				for _, r in edited_df.head(150).iterrows():
					cells_html = "".join(
						"<td>{}</td>".format(html.escape(str(v) if not pd.isna(v) else ""))
						for v in r.tolist()
					)
					trs.append(f"<tr>{cells_html}</tr>")
				xl_inner += f"<table><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table>"
			pdf_bytes = render_html_body_to_pdf_bytes(
				xl_inner, font_family=xl_ff, font_size_pt=float(xl_fs), landscape=is_land, margin_pt=28.0,
				overlays_by_page={0: overlays} if overlays else None,
			)
			pdf_target = _unique_target_path(evrak_dir, new_stem, ".pdf")
			pdf_target.write_bytes(pdf_bytes)
			ctx_api["get_or_create_document_cover"](pdf_target, force=True)
			ui.session_state[flash_k] = f"📑 İmzalı Excel tablosu `{pdf_target.name}` olarak A4 PDF kaydedildi!"
			ui.rerun(scope="fragment")


def _render_image_studio(
	file_path: Path,
	fhash: str,
	evrak_dir: Path,
	ctx_api: dict[str, Any],
	flash_k: str,
) -> None:
	import base64 as _b64
	from PIL import Image
	import pymupdf

	im = Image.open(file_path)
	iw, ih = float(im.width), float(im.height)
	im.close()

	st_key = f"img_studio_ov_{fhash}"
	if st_key not in ui.session_state:
		ui.session_state[st_key] = []
	overlays: list[dict[str, Any]] = ui.session_state[st_key]

	ns = f"img_st_{fhash}"
	cursor_xp_k = f"{ns}_cxp"
	cursor_yp_k = f"{ns}_cyp"
	sel_idx_k = f"{ns}_sel_idx"
	sync_k = f"{ns}_a4_sync"

	raw_sync = ui.text_input("", value="", key=sync_k, label_visibility="collapsed")
	if raw_sync.strip():
		changed, picked_idx = _sync_overlay_canvas_edits(overlays, raw_sync)
		if changed:
			try:
				ev_d = json.loads(raw_sync)
				if ev_d.get("op") == "cursor":
					ui.session_state[cursor_xp_k] = float(ev_d.get("xp", 25.0))
					ui.session_state[cursor_yp_k] = float(ev_d.get("yp", 80.0))
					ui.session_state[f"{ns}_man_xp"] = round(float(ev_d.get("xp", 25.0)), 1)
					ui.session_state[f"{ns}_man_yp"] = round(float(ev_d.get("yp", 80.0)), 1)
			except Exception:
				pass
			if picked_idx is not None:
				ui.session_state[sel_idx_k] = picked_idx
		ui.session_state[sync_k] = ""

	# A4 orantılı mantıksal sayfa ölçüsü
	pw, ph = 595.0, max(240.0, round(595.0 * ih / max(1.0, iw), 1))
	for ov in overlays:
		ov["_ref_w"] = pw
		ov["_ref_h"] = ph

	col_prev, col_edit = ui.columns([1.28, 1.02], gap="large", sticky_first_col=True)
	with col_edit:
		t1, t2 = ui.tabs([f"🖋️ Metin & PNG İmza Ekle ({len(overlays)})", "💾 Kaydet & A4 PDF"])
		with t1:
			_render_overlay_and_signature_controls(
				fields=overlays,
				page_idx=0,
				pw=pw,
				ph=ph,
				evrak_dir=evrak_dir,
				ns=ns,
				flash_k=flash_k,
				cursor_xp_k=cursor_xp_k,
				cursor_yp_k=cursor_yp_k,
				sel_idx_k=sel_idx_k,
			)
		with t2:
			new_stem = ui.text_input("Farklı Kaydet Dosya Adı", value=f"{file_path.stem}_imzali", key=f"img_sv_stem_{fhash}")
			s1, s2 = ui.columns(2)
			with s1:
				if ui.button("💾 Mevcut Görsel Üzerine Kaydet", type="primary", use_container_width=True, key=f"img_sv_over_{fhash}"):
					save_image_with_overlays(file_path, overlays, file_path)
					ctx_api["get_or_create_document_cover"](file_path, force=True)
					ui.session_state[flash_k] = f"✅ `{file_path.name}` güncellendi!"
					ui.rerun(scope="fragment")
			with s2:
				if ui.button("🖼️ Yeni Görsel Kopyası Kaydet", type="secondary", use_container_width=True, key=f"img_sv_copy_{fhash}"):
					tp = _unique_target_path(evrak_dir, new_stem, file_path.suffix.lower())
					save_image_with_overlays(file_path, overlays, tp)
					ctx_api["get_or_create_document_cover"](tp, force=True)
					ui.session_state[flash_k] = f"✅ `{tp.name}` kaydedildi!"
					ui.rerun(scope="fragment")
			if ui.button("📑 A4 PDF Belgesi Olarak Kaydet (.pdf)", type="primary", use_container_width=True, key=f"img_sv_pdf_{fhash}"):
				buf_img = io.BytesIO()
				tmp_out = evrak_dir / f".tmp_img_{fhash}.png"
				save_image_with_overlays(file_path, overlays, tmp_out)
				doc_i = pymupdf.open()
				pg_i = doc_i.new_page(width=595, height=842)
				pg_i.insert_image(pymupdf.Rect(24, 24, 571, 818), filename=str(tmp_out), keep_proportion=True)
				pdf_tp = _unique_target_path(evrak_dir, new_stem, ".pdf")
				pdf_tp.write_bytes(doc_i.tobytes(garbage=3, deflate=True))
				doc_i.close()
				tmp_out.unlink(missing_ok=True)
				ctx_api["get_or_create_document_cover"](pdf_tp, force=True)
				ui.session_state[flash_k] = f"📑 `{pdf_tp.name}` A4 PDF olarak Evrak Deposuna kaydedildi!"
				ui.rerun(scope="fragment")

	with col_prev:
		b64_img = _b64.b64encode(file_path.read_bytes()).decode("ascii")
		mime_i = mimetypes.guess_type(file_path)[0] or "image/png"
		_render_interactive_a4_overlay_stage(
			bg_img_b64=b64_img,
			bg_mime=mime_i,
			page_w=pw,
			page_h=ph,
			fields=overlays,
			stage_id=f"a4_stage_img_{fhash}",
			sync_key=sync_k,
			cursor_xp=float(ui.session_state.get(cursor_xp_k, 25.0)),
			cursor_yp=float(ui.session_state.get(cursor_yp_k, 80.0)),
			selected_idx=ui.session_state.get(sel_idx_k),
			show_guides=True,
			badge_text="🖼️ Görsel / A4 Üstü Düzenleme & PNG İmza",
		)


def _render_text_studio(file_path: Path, fhash: str, evrak_dir: Path, ctx_api: dict[str, Any], flash_k: str) -> None:
	raw_txt = file_path.read_text(encoding="utf-8", errors="replace")
	edited = ui.text_area("📝 Dosya İçeriği", value=raw_txt, height=380, key=f"txt_ed_{fhash}")
	if ui.button("💾 Değişiklikleri Kaydet", type="primary", use_container_width=True, key=f"txt_sv_{fhash}"):
		file_path.write_text(edited or "", encoding="utf-8")
		ctx_api["get_or_create_document_cover"](file_path, force=True)
		ui.session_state[flash_k] = f"✅ `{file_path.name}` kaydedildi!"
		ui.rerun(scope="fragment")


def create_new_office_document_in_depot(
	evrak_dir: Path,
	doc_kind: str,
	title: str,
	app_cfg: dict[str, str],
) -> Path:
	"""Evrak deposunda sıfırdan temiz bir Word (.docx) veya Excel (.xlsx) belgesi oluşturur."""
	evrak_dir.mkdir(parents=True, exist_ok=True)
	school_nm = str(app_cfg.get("school_name") or "Okul Reviri")
	nurse_nm = str(app_cfg.get("nurse_name") or "Okul Revir Sorumlusu")
	today_tr = date.today().strftime("%d.%m.%Y")

	if doc_kind == "xlsx":
		import openpyxl
		from openpyxl.styles import Font

		target = _unique_target_path(evrak_dir, title or f"Revir_Tablo_{date.today().isoformat()}", ".xlsx")
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = "Liste"
		headers = ["Sıra", "Tarih", "Adı Soyadı", "Sınıfı", "Açıklama / İşlem", "Notlar"]
		for ci, h in enumerate(headers, start=1):
			c = ws.cell(row=1, column=ci, value=h)
			c.font = Font(bold=True)
		ws.cell(row=2, column=1, value=1)
		ws.cell(row=2, column=2, value=today_tr)
		wb.save(target)
		return target

	import docx

	target = _unique_target_path(evrak_dir, title or f"Revir_Belge_{date.today().isoformat()}", ".docx")
	d = docx.Document()
	d.add_heading(f"{school_nm} — {title or 'Resmi Revir Belgesi'}", level=1)
	d.add_paragraph(f"Tarih: {today_tr}")
	d.add_paragraph("Konu: ")
	d.add_paragraph("Açıklama: ")
	d.add_paragraph(f"Düzenleyen: {nurse_nm}")
	d.save(target)
	return target


# ==============================================================================
# 5. GEMİNİ AKILLI EVRAK OKUYUCU & ÖĞRENCİ EŞLEŞTİRİCİ + VERİTABANINA İŞLEME
# ==============================================================================

def rank_matching_people(connect_db_fn, hints: dict[str, Any], limit: int = 12) -> list[dict[str, Any]]:
	"""Belgeden çıkan isim, okul no ve sınıf ipuçlarına göre yerel SQLite kişilerini puanlayarak sıralar."""
	name_q = tr_fold(hints.get("name_hint") or "")
	no_q = re.sub(r"\D+", "", str(hints.get("student_no_hint") or ""))
	cls_q = tr_fold(hints.get("class_hint") or "")

	with connect_db_fn() as conn:
		rows = conn.execute(
			"""SELECT s.id, s.full_name, s.student_no, s.birth_date, s.sex, s.person_type,
					  COALESCE(c.name, 'Sınıfsız') AS class_name
			   FROM students s
			   LEFT JOIN classes c ON c.id = s.class_id
			   ORDER BY s.full_name ASC"""
		).fetchall()

	scored: list[tuple[float, dict[str, Any]]] = []
	for r in rows:
		rd = dict(r)
		score = 0.0
		r_no = re.sub(r"\D+", "", str(rd.get("student_no") or ""))
		r_nm = tr_fold(rd.get("full_name") or "")
		r_cls = tr_fold(rd.get("class_name") or "")
		if no_q and r_no and no_q == r_no:
			score += 1.2
		if name_q and r_nm:
			if name_q == r_nm:
				score += 1.5
			elif name_q in r_nm or r_nm in name_q:
				score += 0.95
			else:
				ratio = SequenceMatcher(None, name_q, r_nm).ratio()
				# Kelime kesişimi bonusu
				w_q = set(name_q.split())
				w_r = set(r_nm.split())
				overlap = len(w_q & w_r) / max(1, len(w_q))
				score += max(ratio * 0.9, overlap * 0.95)
		if cls_q and r_cls and (cls_q == r_cls or cls_q in r_cls):
			score += 0.25
		rd["_match_score"] = round(score, 3)
		scored.append((score, rd))

	scored.sort(key=lambda x: (-x[0], x[1]["full_name"]))
	return [x[1] for x in scored[:limit]] if (name_q or no_q or cls_q) else [x[1] for x in scored]


def _append_unique_csv(existing: str | None, addition: str | None) -> str:
	cur = [x.strip() for x in str(existing or "").split(",") if x.strip()]
	cur_f = {tr_fold(x) for x in cur}
	for item in [x.strip() for x in str(addition or "").split(",") if x.strip()]:
		if tr_fold(item) not in cur_f:
			cur.append(item)
			cur_f.add(tr_fold(item))
	return ", ".join(cur)


def apply_extracted_clinical_record(
	connect_db_fn,
	student_id: int,
	extracted: dict[str, Any],
	*,
	apply_profile: bool = True,
	apply_labs: bool = True,
	apply_vitals: bool = True,
	apply_measurements: bool = True,
	apply_meds: bool = True,
	apply_disability: bool = False,
	apply_diabetes: bool = False,
	apply_visit_note: bool = False,
	apply_vaccinations: bool = False,
	apply_screenings: bool = False,
	source_filename: str = "",
	source_bytes: bytes | None = None,
	age_in_months_fn=None,
	moh_schedule: list[dict[str, Any]] | None = None,
	save_screening_fn=None,
) -> list[str]:
	"""Onaylanmış Gemini evrak çıkarımını gerçek SQLite tablolarına (kesin şema uyumuyla) işler."""
	applied_log: list[str] = []
	doc_dt = str(extracted.get("document_date") or "").strip()
	if not re.match(r"^\d{4}-\d{2}-\d{2}$", doc_dt):
		doc_dt = date.today().isoformat()
	now_iso = datetime.now().isoformat(timespec="minutes")

	def _parse_ref_bounds(ref_s: str) -> tuple[float | None, float | None]:
		nums = re.findall(r"-?\d+(?:[.,]\d+)?", str(ref_s or ""))
		if len(nums) >= 2:
			try:
				return float(nums[0].replace(",", ".")), float(nums[1].replace(",", "."))
			except Exception:
				return None, None
		return None, None

	with connect_db_fn() as conn:
		# 1. Sağlık Profili (health_profiles)
		if apply_profile:
			hp_in = extracted.get("health_profile") or {}
			if any(str(hp_in.get(k) or "").strip() for k in ("allergies", "chronic_conditions", "blood_group", "regular_medications", "emergency_notes")):
				cur_hp = conn.execute("SELECT * FROM health_profiles WHERE student_id=?", (student_id,)).fetchone()
				cur_d = dict(cur_hp) if cur_hp else {}
				new_al = _append_unique_csv(cur_d.get("allergies"), hp_in.get("allergies"))
				new_ch = _append_unique_csv(cur_d.get("chronic_conditions"), hp_in.get("chronic_conditions"))
				new_bg = str(hp_in.get("blood_group") or "").strip() or str(cur_d.get("blood_group") or "")
				new_rm = _append_unique_csv(cur_d.get("medications"), hp_in.get("regular_medications"))
				em_old = str(cur_d.get("emergency_notes") or "").strip()
				em_add = str(hp_in.get("emergency_notes") or "").strip()
				new_em = (em_old + ("\n" if em_old and em_add else "") + em_add) if (em_add and em_add not in em_old) else em_old
				if cur_hp:
					conn.execute(
						"""UPDATE health_profiles
						   SET allergies=?, chronic_conditions=?, blood_group=?, medications=?, emergency_notes=?, updated_at=?
						   WHERE student_id=?""",
						(new_al or None, new_ch or None, new_bg or None, new_rm or None, new_em or None, now_iso, student_id),
					)
				else:
					conn.execute(
						"""INSERT INTO health_profiles(student_id, allergies, chronic_conditions, blood_group, medications, emergency_notes, updated_at)
						   VALUES (?, ?, ?, ?, ?, ?, ?)""",
						(student_id, new_al or None, new_ch or None, new_bg or None, new_rm or None, new_em or None, now_iso),
					)
				applied_log.append("📋 Sağlık profili (alerji / kronik / kan grubu / ilaç / acil not) güncellendi")

		# 2. Laboratuvar Tahlilleri (lab_reports + lab_results)
		if apply_labs:
			labs = extracted.get("lab_results") or []
			valid_labs = [lr for lr in labs if isinstance(lr, dict) and str(lr.get("test_name") or "").strip()]
			if valid_labs:
				first_dt = str(valid_labs[0].get("measured_on") or doc_dt).strip()
				if not re.match(r"^\d{4}-\d{2}-\d{2}$", first_dt):
					first_dt = doc_dt
				mime_t = mimetypes.guess_type(source_filename or "rapor.pdf")[0] or "application/pdf"
				cur_rep = conn.execute(
					"""INSERT INTO lab_reports(student_id, report_date, file_name, content_type, content, extracted_text)
					   VALUES (?, ?, ?, ?, ?, ?)""",
					(
						student_id,
						first_dt,
						source_filename or "Gemini_Evrak_Okuyucu",
						mime_t,
						source_bytes if source_bytes else b"",
						str(extracted.get("clinical_summary") or ""),
					),
				)
				rep_id = int(cur_rep.lastrowid)
				lab_ins = 0
				for lr in valid_labs:
					v_num = lr.get("value_num")
					val_txt_raw = lr.get("value_text") if lr.get("value_text") not in (None, "") else lr.get("result_value")
					ref_raw = lr.get("ref_range") if lr.get("ref_range") not in (None, "") else lr.get("reference_range")
					flag_raw = lr.get("flag") if lr.get("flag") not in (None, "") else lr.get("status")
					try:
						v_num = float(str(v_num).replace(",", ".")) if v_num is not None and str(v_num).strip() != "" else None
					except Exception:
						v_num = None
					if v_num is None:
						try:
							m_n = re.search(r"-?\d+(?:[.,]\d+)?", str(val_txt_raw or ""))
							v_num = float(m_n.group(0).replace(",", ".")) if m_n else None
						except Exception:
							v_num = None
					r_low, r_high = _parse_ref_bounds(str(ref_raw or ""))
					val_store: Any = v_num if v_num is not None else (str(val_txt_raw or "").strip() or "0")
					conn.execute(
						"""INSERT INTO lab_results(report_id, test_name, value, unit, ref_low, ref_high, lab_flag)
						   VALUES (?, ?, ?, ?, ?, ?, ?)""",
						(
							rep_id,
							str(lr.get("test_name") or "").strip(),
							val_store,
							str(lr.get("unit") or "").strip() or None,
							r_low,
							r_high,
							str(flag_raw or "Normal").strip() or "Normal",
						),
					)
					lab_ins += 1
				if lab_ins:
					applied_log.append(f"🧪 {lab_ins} laboratuvar tahlil parametresi işlendi")

		# 3. Vital Bulgular (vital_readings)
		if apply_vitals:
			vt = extracted.get("vitals") or {}
			has_vt = any(vt.get(k) is not None and str(vt.get(k)).strip() != "" for k in ("temperature_c", "pulse_bpm", "spo2_pct", "systolic_bp", "diastolic_bp", "glucose_mg_dl"))
			if has_vt:
				def _fnum(v: Any) -> float | None:
					try:
						return float(str(v).replace(",", ".")) if v is not None and str(v).strip() != "" else None
					except Exception:
						return None
				m_at = str(vt.get("measured_at") or "").strip() or f"{doc_dt} 09:00"
				conn.execute(
					"""INSERT INTO vital_readings(student_id, measured_at, systolic_mmhg, diastolic_mmhg, pulse_bpm, temperature_c, spo2_pct, glucose_mg_dl, notes)
					   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						student_id, m_at,
						_fnum(vt.get("systolic_bp")),
						_fnum(vt.get("diastolic_bp")),
						_fnum(vt.get("pulse_bpm")),
						_fnum(vt.get("temperature_c")),
						_fnum(vt.get("spo2_pct")),
						_fnum(vt.get("glucose_mg_dl")),
						f"Evraktan aktarıldı ({source_filename})",
					),
				)
				applied_log.append("💓 Vital bulgu kaydı eklendi")

		# 4. Antropometrik Ölçümler (measurements)
		if apply_measurements:
			ms = extracted.get("measurements") or {}
			try:
				h_cm = float(str(ms.get("height_cm")).replace(",", ".")) if ms.get("height_cm") not in (None, "") else None
			except Exception:
				h_cm = None
			try:
				w_kg = float(str(ms.get("weight_kg")).replace(",", ".")) if ms.get("weight_kg") not in (None, "") else None
			except Exception:
				w_kg = None
			try:
				hc_cm = float(str(ms.get("head_circ_cm")).replace(",", ".")) if ms.get("head_circ_cm") not in (None, "") else None
			except Exception:
				hc_cm = None
			if h_cm or w_kg or hc_cm:
				bmi_v = round(w_kg / ((h_cm / 100.0) ** 2), 2) if (h_cm and w_kg and h_cm > 30) else None
				m_on = str(ms.get("measured_on") or doc_dt).strip()
				if not re.match(r"^\d{4}-\d{2}-\d{2}$", m_on):
					m_on = doc_dt
				age_m = None
				try:
					b_row = conn.execute("SELECT birth_date FROM students WHERE id=?", (student_id,)).fetchone()
					if b_row and b_row["birth_date"] and age_in_months_fn:
						age_m = age_in_months_fn(str(b_row["birth_date"]), date.fromisoformat(m_on))
				except Exception:
					age_m = None
				conn.execute(
					"""INSERT INTO measurements(student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
					   VALUES (?, ?, ?, ?, ?, ?, ?)""",
					(student_id, m_on, h_cm, w_kg, hc_cm, age_m, bmi_v),
				)
				applied_log.append("📏 Boy / Kilo / Baş çevresi ölçümü kaydedildi")

		# 5. İlaç Takip Sistemi (student_medications)
		if apply_meds:
			meds = extracted.get("medications") or []
			med_ins = 0
			for md in meds:
				if not isinstance(md, dict) or not str(md.get("medication_name") or "").strip():
					continue
				conn.execute(
					"""INSERT INTO student_medications(student_id, medication_name, med_type, dosage_pattern, scheduled_times, start_date, special_notes, status, created_at)
					   VALUES (?, ?, ?, ?, ?, ?, ?, 'Aktif', ?)""",
					(
						student_id,
						str(md.get("medication_name") or "").strip(),
						str(md.get("med_type") or "Düzenli İlaç").strip(),
						str(md.get("dosage") or "").strip() or None,
						str(md.get("schedule_times") or "").strip() or None,
						doc_dt,
						f"{str(md.get('notes') or '').strip()} [Evrak: {source_filename}]".strip(),
						now_iso,
					),
				)
				med_ins += 1
			if med_ins:
				applied_log.append(f"💊 {med_ins} ilaç kaydı İlaç Takip Sistemine eklendi")

		# 6. Engellilik / Özel Gereksinim Profili (students.has_disability + health_profiles)
		if apply_disability:
			dis = extracted.get("disability") or {}
			if dis.get("has_disability"):
				cur_hp = conn.execute("SELECT disability_profile_json FROM health_profiles WHERE student_id=?", (student_id,)).fetchone()
				try:
					cur_dp = json.loads(cur_hp["disability_profile_json"] or "{}") if (cur_hp and cur_hp["disability_profile_json"]) else {}
				except Exception:
					cur_dp = {}
				cur_dp["has_disability"] = True
				if dis.get("disability_rate") not in (None, ""):
					try:
						cur_dp["disability_rate"] = int(float(dis["disability_rate"]))
					except Exception:
						pass
				if dis.get("categories"):
					cats = list(dict.fromkeys((cur_dp.get("categories") or []) + [str(c) for c in dis["categories"] if str(c).strip()]))
					cur_dp["categories"] = cats
				if str(dis.get("mobility_status") or "").strip():
					cur_dp["mobility_status"] = str(dis["mobility_status"]).strip()
				if str(dis.get("school_adaptations") or "").strip():
					cur_dp["school_adaptations"] = str(dis["school_adaptations"]).strip()
				if str(dis.get("report_summary") or "").strip():
					cur_dp["notes"] = str(dis["report_summary"]).strip()
				rate_val = int(cur_dp.get("disability_rate") or 0)
				dp_json = json.dumps(cur_dp, ensure_ascii=False)
				conn.execute("UPDATE students SET has_disability=1 WHERE id=?", (student_id,))
				if cur_hp:
					conn.execute(
						"UPDATE health_profiles SET disability_rate=?, disability_profile_json=?, updated_at=? WHERE student_id=?",
						(rate_val, dp_json, now_iso, student_id),
					)
				else:
					conn.execute(
						"INSERT INTO health_profiles(student_id, disability_rate, disability_profile_json, updated_at) VALUES (?, ?, ?, ?)",
						(student_id, rate_val, dp_json, now_iso),
					)
				applied_log.append("♿ Özel gereksinim / engellilik profili güncellendi")

		# 8. Revir İzlem Notu (visits)
		if apply_visit_note and str(extracted.get("visit_note_suggestion") or extracted.get("clinical_summary") or "").strip():
			v_note = str(extracted.get("visit_note_suggestion") or extracted.get("clinical_summary") or "").strip()
			conn.execute(
				"""INSERT INTO visits(student_id, visit_date, visit_time, complaint, observation, action, follow_up_status)
				   VALUES (?, ?, ?, ?, ?, ?, 'Kapalı')""",
				(
					student_id,
					doc_dt,
					datetime.now().strftime("%H:%M"),
					f"Belge / Rapor Kaydı ({extracted.get('document_type', 'Sağlık Belgesi')})",
					f"{v_note} [Dosya: {source_filename}]",
					"Dış rapor / evrak okunup dosyaya işlendi",
				),
			)
			applied_log.append("🩺 Revir başvuru / izlem notu eklendi")

	# 7. Diyabet / İnsülin Profili (Dışarıda kendi kısa işleminde yaz — SQLite kilit çakışmasını önler!)
	if apply_diabetes:
		diab = extracted.get("diabetes") or {}
		if diab.get("is_diabetic"):
			with connect_db_fn() as _dc:
				idm.ensure_diabetes_tables(_dc)
			cur_p = idm.get_diabetes_profile(connect_db_fn, student_id)
			upd_p = dict(cur_p)
			upd_p["is_diabetic"] = 1
			if str(diab.get("diabetes_type") or "").strip():
				upd_p["diabetes_type"] = str(diab["diabetes_type"]).strip()
			for num_k in ("icr_breakfast", "icr_lunch", "icr_snack", "isf", "target_bg"):
				if diab.get(num_k) not in (None, ""):
					try:
						upd_p[num_k] = float(str(diab[num_k]).replace(",", "."))
					except Exception:
						pass
			if str(diab.get("insulin_brand") or "").strip():
				upd_p["insulin_brand"] = str(diab["insulin_brand"]).strip()
			idm.save_diabetes_profile(connect_db_fn, student_id, upd_p)
			applied_log.append("💉 Diyabet / İnsülin profili (İKO / İDF) güncellendi")

	# 9. Aşı Takibi (student_vaccinations)
	if apply_vaccinations:
		vacs = extracted.get("vaccinations") or []
		if isinstance(vacs, dict):
			vacs = [vacs]
		sched_list = [v for v in (moh_schedule or []) if isinstance(v, dict) and v.get("code")]
		sched_by_code = {str(v["code"]): v for v in sched_list}

		def _resolve_vaccine_meta(rec: dict[str, Any]) -> tuple[str, dict[str, Any] | None]:
			raw_code = str(rec.get("vaccine_code") or "").strip()
			if raw_code and raw_code in sched_by_code:
				return raw_code, sched_by_code[raw_code]
			v_fold = tr_fold(rec.get("vaccine_name") or "")
			d_fold = tr_fold(rec.get("dose_name") or "")
			if v_fold and sched_list:
				for meta in sched_list:
					mv = tr_fold(meta.get("vaccine") or "")
					md = tr_fold(meta.get("dose") or "")
					if (v_fold in mv or mv in v_fold) and (not d_fold or d_fold in md or md in d_fold):
						return str(meta["code"]), meta
			if raw_code:
				return raw_code, None
			return "", None

		vac_ins = 0
		with connect_db_fn() as _vc:
			for v_rec in vacs:
				if not isinstance(v_rec, dict):
					continue
				v_code, v_meta = _resolve_vaccine_meta(v_rec)
				if not v_code:
					continue
				v_name = str(v_meta["vaccine"] if v_meta else (v_rec.get("vaccine_name") or v_code)).strip()
				d_name = str(v_meta["dose"] if v_meta else (v_rec.get("dose_name") or "1. Doz")).strip()
				p_lbl = str(v_meta["period_label"] if v_meta else "Diğer").strip()
				p_mos = int(v_meta.get("period_months") or 0) if v_meta else 0
				adm_d = str(v_rec.get("administered_date") or "").strip()[:10] or None
				fac_v = str(v_rec.get("facility") or "").strip() or "Aile Sağlığı Merkezi (ASM)"
				lot_v = str(v_rec.get("lot_number") or "").strip() or None
				not_v = str(v_rec.get("notes") or "").strip() or None
				_vc.execute(
					"""INSERT INTO student_vaccinations(
						student_id, vaccine_code, vaccine_name, dose_name, period_label,
						period_months, is_done, administered_date, facility, lot_number, notes, updated_at
					) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
					ON CONFLICT(student_id, vaccine_code) DO UPDATE SET
						is_done=excluded.is_done,
						administered_date=COALESCE(excluded.administered_date, student_vaccinations.administered_date),
						facility=COALESCE(excluded.facility, student_vaccinations.facility),
						lot_number=COALESCE(excluded.lot_number, student_vaccinations.lot_number),
						notes=COALESCE(excluded.notes, student_vaccinations.notes),
						updated_at=excluded.updated_at""",
					(
						student_id, v_code, v_name, d_name, p_lbl, p_mos,
						1 if v_rec.get("is_done", True) else 0,
						adm_d, fac_v, lot_v, not_v, now_iso,
					),
				)
				vac_ins += 1
		if vac_ins:
			applied_log.append(f"💉 {vac_ins} aşı dozu takvim kaydına işlendi")

	# 10. Okul Taraması (school_screenings)
	if apply_screenings:
		scr = extracted.get("school_screenings") or {}
		if isinstance(scr, list):
			scr = scr[0] if scr and isinstance(scr[0], dict) else {}
		if isinstance(scr, dict) and any(
			scr.get(k) not in (None, "")
			for k in ("height_cm", "weight_kg", "head_circ_cm", "temperature_c", "pulse_bpm", "spo2_pct", "wears_glasses", "vision_right", "vision_left", "hair_screening", "hearing_right", "hearing_left", "health_behavior", "psychosocial_status", "notes")
		):
			s_dt = str(scr.get("screening_date") or doc_dt).strip()[:10]
			if not re.match(r"^\d{4}-\d{2}-\d{2}$", s_dt):
				s_dt = doc_dt

			def _sflt(v: Any) -> float | None:
				try:
					return float(str(v).replace(",", ".")) if v not in (None, "") else None
				except Exception:
					return None

			if callable(save_screening_fn):
				save_screening_fn(
					student_id=student_id,
					screening_date=s_dt,
					height_cm=_sflt(scr.get("height_cm")),
					weight_kg=_sflt(scr.get("weight_kg")),
					head_circ_cm=_sflt(scr.get("head_circ_cm")),
					temperature_c=_sflt(scr.get("temperature_c")),
					pulse_bpm=_sflt(scr.get("pulse_bpm")),
					spo2_pct=_sflt(scr.get("spo2_pct")),
					wears_glasses=str(scr.get("wears_glasses") or "").strip() or None,
					vision_right=str(scr.get("vision_right") or "").strip() or None,
					vision_left=str(scr.get("vision_left") or "").strip() or None,
					hair_screening=str(scr.get("hair_screening") or "").strip() or None,
					hearing_right=str(scr.get("hearing_right") or "").strip() or None,
					hearing_left=str(scr.get("hearing_left") or "").strip() or None,
					health_behavior=str(scr.get("health_behavior") or "").strip() or None,
					psychosocial_status=str(scr.get("psychosocial_status") or "").strip() or None,
					notes=str(scr.get("notes") or f"Sohbet / Evraktan aktarıldı ({source_filename})").strip(),
				)
			else:
				with connect_db_fn() as _sc:
					_sc.execute(
						"""INSERT INTO school_screenings(
							student_id, screening_date, height_cm, weight_kg, head_circ_cm,
							temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right,
							vision_left, hair_screening, hearing_right, hearing_left,
							health_behavior, harmful_habits, psychosocial_status, notes, screened_by, created_at
						) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(
							student_id, s_dt, _sflt(scr.get("height_cm")), _sflt(scr.get("weight_kg")), _sflt(scr.get("head_circ_cm")),
							_sflt(scr.get("temperature_c")), _sflt(scr.get("pulse_bpm")), _sflt(scr.get("spo2_pct")),
							str(scr.get("wears_glasses") or "").strip() or None,
							str(scr.get("vision_right") or "").strip() or None,
							str(scr.get("vision_left") or "").strip() or None,
							str(scr.get("hair_screening") or "").strip() or None,
							str(scr.get("hearing_right") or "").strip() or None,
							str(scr.get("hearing_left") or "").strip() or None,
							str(scr.get("health_behavior") or "").strip() or None,
							str(scr.get("harmful_habits") or "").strip() or None,
							str(scr.get("psychosocial_status") or "").strip() or None,
							str(scr.get("notes") or f"Sohbet / Evraktan aktarıldı ({source_filename})").strip(),
							None, now_iso,
						),
					)
					_sc.execute(
						"UPDATE health_profiles SET last_screened_at=COALESCE(?, last_screened_at), updated_at=? WHERE student_id=?",
						(s_dt, now_iso, student_id),
					)
			applied_log.append("🏫 Okul sağlık taraması kaydı işlendi")

	return applied_log


def render_smart_document_importer_dialog(ctx_api: dict[str, Any], initial_file_path: Path | None = None) -> None:
	"""
	Herhangi bir sağlık raporunu, tahlili, reçeteyi, ÇÖZGER/diyabet raporunu veya sınıf/öğrenci listesini
	Gemini ile okuyup ait olduğu öğrenciyi bulan ve hemşire onayıyla doğrudan veritabanına işleyen akıllı pencere.
	"""
	is_gai_on, gai_key, gai_model = ctx_api["_get_active_gemini_config"]()
	if not is_gai_on:
		ui.warning("🤖 Bu özellik için **Ayarlar → Yapay Zekâ & KVKK** bölümünden yapay zekâ asistanının etkinleştirilmesi gerekir. Yapay zekâ kapalıyken tüm manuel form ve Excel/PDF yükleme ekranlarını normal şekilde kullanabilirsiniz.")
		return

	ui.caption("Yüklediğiniz veya Evrak Deposundan seçtiğiniz belgeyi Gemini okur, hangi öğrenciye ait olduğunu yerel veritabanından eşleştirir ve çıkardığı tüm klinik bilgileri onayınızla doğrudan dosyasına işler.")

	src_bytes: bytes | None = None
	src_name: str = ""
	if initial_file_path and initial_file_path.exists():
		src_bytes = initial_file_path.read_bytes()
		src_name = initial_file_path.name
		ui.info(f"📄 **Seçili Depo Belgesi:** `{src_name}`")
	else:
		evrak_dir: Path = ctx_api["EVRAKLAR_DIR"]
		depot_files = sorted([f for f in evrak_dir.iterdir() if f.is_file() and not f.name.startswith(".")], key=lambda x: x.stat().st_mtime, reverse=True) if evrak_dir.exists() else []
		s_c1, s_c2 = ui.columns([1.4, 1.2])
		with s_c1:
			up_f = ui.file_uploader(
				"📤 Yeni Belge Yükle (PDF, Fotoğraf, Word, Excel, TXT)",
				type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "csv", "txt"],
				key="smart_imp_uploader",
			)
		with s_c2:
			dep_opts = ["— Seçiniz —"] + [f.name for f in depot_files[:40]]
			picked_dep = ui.selectbox("📁 Veya Evrak Deposundan Seç", dep_opts, index=0, key="smart_imp_depot_sel")
		if up_f is not None:
			src_bytes = up_f.read()
			src_name = up_f.name
		elif picked_dep != "— Seçiniz —":
			p_cand = evrak_dir / picked_dep
			if p_cand.exists():
				src_bytes = p_cand.read_bytes()
				src_name = p_cand.name

	extra_note = ui.text_input(
		"💬 Ek İpucu / Talimat (Opsiyonel)",
		placeholder="Örn: Bu bir tahlil ve ÇÖZGER raporu, alerji ve ilaçları da çıkar",
		key="smart_imp_extra_note",
	)

	res_key = f"smart_imp_result_{sha256(src_name.encode('utf-8')).hexdigest()[:10]}" if src_name else "smart_imp_result_none"
	if ui.button("✨ Belgeyi Gemini ile Oku & Öğrenciyi Eşleştir", type="primary", disabled=(not src_bytes), use_container_width=True, key="btn_smart_imp_analyze"):
		ext_txt = ""
		if src_name.lower().endswith(".pdf") and src_bytes:
			try:
				ext_txt = ctx_api["extract_pdf_text"](src_bytes)
			except Exception:
				ext_txt = ""
		with ui.spinner("Gemini belgeyi okuyor ve klinik bilgileri ayrıştırıyor..."):
			strict_pdf_mode = False
			if callable(ctx_api.get("get_app_settings")):
				try:
					strict_pdf_mode = str(ctx_api["get_app_settings"]().get("gemini_strict_kvkk_pdf", "0")).strip() == "1"
				except Exception:
					strict_pdf_mode = False
			ok_ex, ex_data, used_m = gai.extract_clinical_record_from_document_with_gemini(
				api_key=gai_key,
				model=gai_model,
				file_bytes=src_bytes or b"",
				file_name=src_name,
				extracted_text=ext_txt,
				extra_instruction=extra_note,
				strict_kvkk_pdf=strict_pdf_mode,
			)
		ui.session_state[res_key] = {"ok": ok_ex, "data": ex_data, "model": used_m, "name": src_name}
		ui.rerun(scope="fragment")

	saved_res = ui.session_state.get(res_key)
	if not isinstance(saved_res, dict):
		return
	if not saved_res.get("ok"):
		ui.error(f"Belge analiz hatası: {saved_res.get('data', {}).get('error', 'Bilinmeyen hata')}")
		return

	data = saved_res["data"]
	hints = data.get("detected_person_hints") or {}

	with ui.container(border=True):
		ui.markdown(f"##### 🤖 Analiz Tamamlandı ({saved_res.get('model')}) · Tür: **`{data.get('document_type', 'Sağlık Belgesi')}`**")
		if data.get("clinical_summary"):
			ui.markdown(f"📝 **Klinik Özet:** {data['clinical_summary']}")
		if hints.get("name_hint") or hints.get("student_no_hint") or hints.get("class_hint"):
			ui.caption(
				f"🔎 Belgeden okunan kişi ipucu: **Ad:** `{hints.get('name_hint') or '—'}` · "
				f"**No:** `{hints.get('student_no_hint') or '—'}` · **Sınıf:** `{hints.get('class_hint') or '—'}`"
			)

		# Öğrenci Eşleştirme Seçicisi
		ranked = rank_matching_people(ctx_api["connect_db"], hints, limit=25)
		match_opts = []
		lbl_to_pid: dict[str, int] = {}
		for rp in ranked:
			pct = int(min(99, rp.get("_match_score", 0) * 65))
			tag = f"🎯 %{pct} eşleşme · " if rp.get("_match_score", 0) >= 0.45 else ""
			sno = f" (No: {rp['student_no']})" if rp.get("student_no") else ""
			lbl = f"{tag}{rp['full_name']} · {rp['class_name']}{sno}"
			match_opts.append(lbl)
			lbl_to_pid[lbl] = int(rp["id"])

		if not match_opts:
			ui.warning("Veritabanında henüz kayıtlı kişi yok. Önce öğrenci oluşturun.")
			return

		picked_p_lbl = ui.selectbox("👤 Bilgilerin İşleneceği Hedef Öğrenci / Kişi (En yüksek eşleşme en üsttedir)", match_opts, index=0, key=f"smart_imp_target_{res_key}")
		target_pid = lbl_to_pid[picked_p_lbl]

		# Çıkarılan Alanların Onay Kutuları ve Önizlemeleri
		hp_d = data.get("health_profile") or {}
		labs_d = data.get("lab_results") or []
		vit_d = data.get("vitals") or {}
		meas_d = data.get("measurements") or {}
		meds_d = data.get("medications") or []
		vacs_d = data.get("vaccinations") or []
		dis_d = data.get("disability") or {}
		diab_d = data.get("diabetes") or {}
		v_note_d = str(data.get("visit_note_suggestion") or "").strip()

		has_hp = any(str(hp_d.get(k) or "").strip() for k in ("allergies", "chronic_conditions", "blood_group", "regular_medications", "emergency_notes"))
		has_vt = any(vit_d.get(k) not in (None, "") for k in ("temperature_c", "pulse_bpm", "spo2_pct", "systolic_bp", "diastolic_bp", "glucose_mg_dl"))
		has_ms = any(meas_d.get(k) not in (None, "") for k in ("height_cm", "weight_kg", "head_circ_cm"))

		ui.markdown("###### ✅ İşlenecek Bölümleri Seçin ve Onaylayın:")
		chk1, chk2 = ui.columns(2)
		with chk1:
			c_hp = ui.checkbox(f"📋 Sağlık Profilini Güncelle {'(Veri bulundu)' if has_hp else '(Veri yok)'}", value=has_hp, disabled=not has_hp, key=f"si_chk_hp_{res_key}")
			if has_hp:
				ui.caption(f"Alerji: `{hp_d.get('allergies') or '—'}` · Kronik: `{hp_d.get('chronic_conditions') or '—'}` · Kan: `{hp_d.get('blood_group') or '—'}`")
			c_labs = ui.checkbox(f"🧪 Laboratuvar Sonuçlarını İşle ({len(labs_d)} parametre)", value=bool(labs_d), disabled=not labs_d, key=f"si_chk_lab_{res_key}")
			c_vt = ui.checkbox(f"💓 Vital Bulguları Kaydet {'(Veri bulundu)' if has_vt else '(Veri yok)'}", value=has_vt, disabled=not has_vt, key=f"si_chk_vt_{res_key}")
			c_ms = ui.checkbox(f"📏 Boy / Kilo Ölçümü Kaydet {'(Veri bulundu)' if has_ms else '(Veri yok)'}", value=has_ms, disabled=not has_ms, key=f"si_chk_ms_{res_key}")
			c_vac = ui.checkbox(f"💉 Aşı Takvimine İşle ({len(vacs_d)} doz)", value=bool(vacs_d), disabled=not vacs_d, key=f"si_chk_vac_{res_key}")
		with chk2:
			c_meds = ui.checkbox(f"💊 İlaç Takip Listesine Ekle ({len(meds_d)} ilaç)", value=bool(meds_d), disabled=not meds_d, key=f"si_chk_med_{res_key}")
			c_dis = ui.checkbox(f"♿ Özel Gereksinim / Engel Profilini Güncelle {'(Tespit edildi)' if dis_d.get('has_disability') else '(Yok)'}", value=bool(dis_d.get("has_disability")), disabled=not dis_d.get("has_disability"), key=f"si_chk_dis_{res_key}")
			c_diab = ui.checkbox(f"💉 Diyabet / İnsülin Profilini Güncelle {'(Tespit edildi)' if diab_d.get('is_diabetic') else '(Yok)'}", value=bool(diab_d.get("is_diabetic")), disabled=not diab_d.get("is_diabetic"), key=f"si_chk_diab_{res_key}")
			c_vis = ui.checkbox("🩺 Revir İzlem / Belge Notu Olarak Kaydet", value=bool(v_note_d), key=f"si_chk_vis_{res_key}")

		if labs_d:
			ui.dataframe(pd.DataFrame(labs_d)[["test_name", "value_text", "unit", "ref_range", "flag"]].rename(columns={"test_name": "Test", "value_text": "Sonuç", "unit": "Birim", "ref_range": "Referans", "flag": "Durum"}), use_container_width=True, height=180)
		if meds_d:
			ui.dataframe(pd.DataFrame(meds_d), use_container_width=True, height=140)
		if vacs_d:
			ui.dataframe(pd.DataFrame(vacs_d), use_container_width=True, height=140)
		if data.get("inferred_risks_and_actions"):
			ui.info("🧠 **Belgeden Çıkarılan Dikkat Noktaları:**\n" + "\n".join(f"- {x}" for x in data["inferred_risks_and_actions"]))

		save_to_depot = False
		if initial_file_path is None and src_bytes:
			save_to_depot = ui.checkbox("📁 Bu belgeyi ayrıca Revir Evrak Deposuna da arşivle", value=True, key=f"si_chk_dep_{res_key}")

		if ui.button("✅ Onayla ve Seçili Verileri Kişi Dosyasına İşle", type="primary", use_container_width=True, key=f"btn_si_commit_{res_key}"):
			logs = apply_extracted_clinical_record(
				ctx_api["connect_db"],
				target_pid,
				data,
				apply_profile=c_hp,
				apply_labs=c_labs,
				apply_vitals=c_vt,
				apply_measurements=c_ms,
				apply_meds=c_meds,
				apply_disability=c_dis,
				apply_diabetes=c_diab,
				apply_visit_note=c_vis,
				apply_vaccinations=c_vac,
				source_filename=src_name,
				source_bytes=src_bytes,
				age_in_months_fn=ctx_api.get("age_in_months"),
				moh_schedule=ctx_api.get("MOH_VACCINE_SCHEDULE"),
				save_screening_fn=ctx_api.get("save_school_screening"),
			)
			if save_to_depot and src_bytes and src_name:
				evrak_dir: Path = ctx_api["EVRAKLAR_DIR"]
				evrak_dir.mkdir(parents=True, exist_ok=True)
				dst = _unique_target_path(evrak_dir, Path(src_name).stem, Path(src_name).suffix.lower() or ".pdf")
				dst.write_bytes(src_bytes)
				ctx_api["get_or_create_document_cover"](dst, force=True)
				logs.append(f"📁 Belge `{dst.name}` olarak Evrak Deposuna arşivlendi")
			ctx_api["sync_after_person_mutation"](target_pid)
			ctx_api["set_active_person"](target_pid)
			ui.session_state.pop(res_key, None)
			if logs:
				ui.success("✅ İşlem Tamamlandı:\n\n" + "\n".join(f"- {x}" for x in logs))
			else:
				ui.info("İşlenecek aktif bir bölüm seçilmedi.")
			ui.rerun(scope="fragment")
