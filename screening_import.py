"""
Okul Sağlık Taraması — Akıllı Excel Şablon Üreticisi & 3 Katmanlı Hibrit İçe Aktarım Motoru (screening_import.py)

Özellikler:
1. Öğrenci ve Yetişkin/Personel için resmi, veri doğrulamalı (.xlsx) tarama şablonları
   (hem boş örnekli şablon hem mevcut okul/personel listesi önceden doldurulmuş hazır şablon).
2. Yüksek başarı oranlı çoklu kaynak okuyucu (.xlsx, .xls, .csv, .tsv, .docx, .pdf, .jpg, .png, .webp + Kopyala-Yapıştır metin):
   - Üst başlık satırlarını atlayıp gerçek tablo başlığını otomatik bulma,
   - Kopyala-yapıştır metinlerde alt satıra kaymış (wrapped) son sütun değerlerini otomatik birleştirme,
   - '-' / 'yok' / 'boş' yer tutucularını temizleme, Türkçe virgüllü ondalık ('77,6') ve tansiyon ('115-60', '120/80') çözümleme,
   - Yetişkin / Öğrenci tablosunu sütun yapısından otomatik algılama,
   - Görsel ön-işleme (EXIF düzeltme, kontrast iyileştirme, ölçekleme) + Yerel OCR + Gemini Multimodal Vision AI desteği.
3. Türkçe karakter duyarsız + bulanık (fuzzy) isim eşleştirme ve tek-transaction hızlı SQLite toplu kayıt motoru.
"""
from __future__ import annotations

from datetime import date, datetime
from difflib import SequenceMatcher
import html
import io
import math
import re
from typing import Any, Callable

import pandas as pd

import revir_ui as ui

_PLACEHOLDER_TOKENS: frozenset[str] = frozenset({
	"", "-", "—", "–", "--", "---", ".", "..", "?",
	"yok", "boş", "bos", "ölçülmedi", "olculmedi",
	"yapılmadı", "yapilmadi", "n/a", "na", "null", "none", "nan", "<na>",
})

_TR_FOLD_MAP = str.maketrans({
	"İ": "i", "I": "i", "ı": "i",
	"Ğ": "g", "ğ": "g",
	"Ü": "u", "ü": "u",
	"Ş": "s", "ş": "s",
	"Ö": "o", "ö": "o",
	"Ç": "c", "ç": "c",
})

STUDENT_TEMPLATE_COLUMNS: list[str] = [
	"Sınıf",
	"Okul No",
	"Ad Soyad",
	"Cinsiyet",
	"Doğum Tarihi",
	"Boy (cm)",
	"Kilo (kg)",
	"Baş Çevresi (cm)",
	"Ateş (°C)",
	"Nabız (/dk)",
	"SpO2 (%)",
	"Gözlük",
	"Sağ Göz",
	"Sol Göz",
	"Saç / Hijyen",
	"Sağ Kulak",
	"Sol Kulak",
	"Tarama Notu",
]

ADULT_TEMPLATE_COLUMNS: list[str] = [
	"Ad Soyad",
	"Görev / Branş",
	"Cinsiyet",
	"Doğum Tarihi",
	"Boy (cm)",
	"Kilo (kg)",
	"Tansiyon (mmHg)",
	"Kan Şekeri (mg/dL)",
	"KŞ Zamanı",
	"Nabız (/dk)",
	"Ateş (°C)",
	"SpO2 (%)",
	"Görme Durumu",
	"Kronik / Tarama Notu",
]


def tr_fold(val: object) -> str:
	s = str(val or "").strip().translate(_TR_FOLD_MAP).lower()
	return re.sub(r"\s+", " ", s).strip()


def tr_col_key(val: object) -> str:
	return re.sub(r"[^a-z0-9]", "", tr_fold(val))


def is_placeholder_or_blank(val: object) -> bool:
	if val is None:
		return True
	if isinstance(val, float) and math.isnan(val):
		return True
	try:
		if pd.isna(val):
			return True
	except (TypeError, ValueError):
		pass
	s = str(val).strip()
	if not s:
		return True
	return tr_fold(s) in _PLACEHOLDER_TOKENS


def clean_cell_str(val: object) -> str | None:
	if is_placeholder_or_blank(val):
		return None
	s = re.sub(r"\s+", " ", str(val)).strip()
	return s or None


def parse_flex_float(val: object) -> float | None:
	if is_placeholder_or_blank(val):
		return None
	if isinstance(val, (int, float)) and not isinstance(val, bool):
		return None if math.isnan(float(val)) else float(val)
	s = str(val).strip()
	# Yalnızca tek sayı içeren değerleri yakala (115-60 tansiyon gibi çift sayıları tek sayı sanma)
	if re.search(r"\d+\s*[/\-]\s*\d+", s):
		return None
	s_clean = re.sub(r"[^\d,.\-]", "", s).replace(",", ".")
	if not s_clean or s_clean in {".", "-", "-."}:
		return None
	try:
		v = float(s_clean)
		return None if math.isnan(v) else v
	except ValueError:
		return None


def normalize_height_smart(val: object) -> float | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	if 0.45 <= num <= 2.35:
		num = num * 100.0
	elif 450.0 <= num <= 2350.0:
		num = num / 10.0
	num = round(num, 1)
	return num if 40.0 <= num <= 235.0 else None


def normalize_weight_smart(val: object, *, is_adult: bool = False) -> float | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	if 2000.0 <= num <= 200000.0:
		num = num / 1000.0
	elif (not is_adult) and (130.0 < num <= 999.0):
		num = num / 10.0
	elif is_adult and (230.0 < num <= 1999.0):
		num = num / 10.0
	num = round(num, 2)
	return num if 1.0 <= num <= 300.0 else None


def normalize_head_circ_smart(val: object) -> float | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	if 300.0 <= num <= 650.0:
		num = num / 10.0
	num = round(num, 1)
	return num if 30.0 <= num <= 65.0 else None


def normalize_bp_smart(bp_or_sys: object, dia_val: object = None) -> tuple[int | None, int | None]:
	if is_placeholder_or_blank(bp_or_sys) and is_placeholder_or_blank(dia_val):
		return (None, None)

	def _scale_comp(n: float | None, is_sys: bool) -> int | None:
		if n is None or n <= 0:
			return None
		if is_sys:
			if 6.0 <= n <= 24.0:
				n *= 10.0
			iv = int(round(n))
			return iv if 50 <= iv <= 260 else None
		if 3.0 <= n <= 15.0:
			n *= 10.0
		iv = int(round(n))
		return iv if 30 <= iv <= 160 else None

	if not is_placeholder_or_blank(bp_or_sys) and is_placeholder_or_blank(dia_val):
		raw = re.sub(r"\bmmhg\b", "", str(bp_or_sys), flags=re.IGNORECASE).strip()
		m_pair = re.search(r"(\d{1,3}(?:[.,]\d+)?)\s*[/\-\\:; ]\s*(\d{1,3}(?:[.,]\d+)?)", raw)
		if m_pair:
			s1 = float(m_pair.group(1).replace(",", "."))
			d1 = float(m_pair.group(2).replace(",", "."))
			so, do = _scale_comp(s1, True), _scale_comp(d1, False)
			if so and do and so < do:
				so2, do2 = _scale_comp(d1, True), _scale_comp(s1, False)
				if so2 and do2 and so2 > do2:
					return (so2, do2)
			return (so, do)
		digits = re.sub(r"\s+", "", raw)
		if digits.endswith(".0"):
			digits = digits[:-2]
		if digits.isdigit() and len(digits) in (4, 5, 6):
			cs, cd = (int(digits[:2]), int(digits[2:])) if len(digits) == 4 else (int(digits[:3]), int(digits[3:]))
			if 50 <= cs <= 260 and 30 <= cd <= 160 and cs > cd:
				return (cs, cd)

	sn = parse_flex_float(bp_or_sys)
	dn = parse_flex_float(dia_val)
	return (_scale_comp(sn, True), _scale_comp(dn, False))


def normalize_pulse_smart(val: object) -> int | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	iv = int(round(num))
	return iv if 30 <= iv <= 240 else None


def normalize_glucose_smart(val: object) -> float | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	num = round(num, 1)
	return num if 20.0 <= num <= 800.0 else None


def normalize_temp_smart(val: object) -> float | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	if 330.0 <= num <= 430.0:
		num = num / 10.0
	num = round(num, 1)
	return num if 33.0 <= num <= 43.0 else None


def normalize_spo2_smart(val: object) -> int | None:
	num = parse_flex_float(val)
	if num is None or num <= 0:
		return None
	if 500.0 <= num <= 999.0:
		num = int(num // 10)
	iv = int(round(num))
	return iv if 50 <= iv <= 100 else None


def normalize_sex_smart(val: object) -> str | None:
	if is_placeholder_or_blank(val):
		return None
	f = tr_fold(val)
	if f in {"e", "erkek", "m", "male", "bay", "oglan", "boy"}:
		return "E"
	if f in {"k", "kiz", "kadin", "f", "female", "bayan", "girl"}:
		return "K"
	return None


def normalize_date_smart(val: object) -> str | None:
	if is_placeholder_or_blank(val):
		return None
	if isinstance(val, (datetime, date, pd.Timestamp)):
		return val.strftime("%Y-%m-%d")
	s = str(val).strip().replace("/", "-").replace(".", "-")
	m = re.search(r"\b(\d{4}-\d{1,2}-\d{1,2}|\d{1,2}-\d{1,2}-\d{4})\b", s)
	if not m:
		# Yalnızca doğum yılı girilmişse (ör. 1984) -> 1984-01-01
		if re.fullmatch(r"(19[4-9]\d|20[0-2]\d)", s):
			return f"{s}-01-01"
		return None
	parts = m.group(1).split("-")
	try:
		if len(parts[0]) == 4:
			y, mo, dy = int(parts[0]), int(parts[1]), int(parts[2])
		else:
			dy, mo, y = int(parts[0]), int(parts[1]), int(parts[2])
		return date(y, mo, dy).isoformat()
	except ValueError:
		return None


# ==============================================================================
# 1. RESMİ ÖĞRENCİ VE YETİŞKİN EXCEL TARAMA ŞABLONLARI (.xlsx)
# ==============================================================================

def _style_screening_workbook(wb, ws, headers: list[str], col_widths: list[int], accent_hex: str, rules_lines: list[str], validations: list[tuple[int, str]]) -> bytes:
	from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
	from openpyxl.utils import get_column_letter
	from openpyxl.worksheet.datavalidation import DataValidation

	hdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
	hdr_fill = PatternFill(start_color=accent_hex, end_color=accent_hex, fill_type="solid")
	row_font = Font(name="Segoe UI", size=10)
	zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
	thin_border = Border(
		left=Side(style="thin", color="CBD5E1"),
		right=Side(style="thin", color="CBD5E1"),
		top=Side(style="thin", color="CBD5E1"),
		bottom=Side(style="thin", color="CBD5E1"),
	)
	center_al = Alignment(horizontal="center", vertical="center")
	left_al = Alignment(horizontal="left", vertical="center")

	for c_idx in range(1, len(headers) + 1):
		cell = ws.cell(row=1, column=c_idx)
		cell.font = hdr_font
		cell.fill = hdr_fill
		cell.alignment = center_al
		cell.border = thin_border
	ws.row_dimensions[1].height = 28
	ws.freeze_panes = "A2"

	max_r = max(ws.max_row, 2)
	for r_i in range(2, max_r + 1):
		ws.row_dimensions[r_i].height = 21
		for c_i in range(1, len(headers) + 1):
			c = ws.cell(row=r_i, column=c_i)
			c.font = row_font
			c.border = thin_border
			if r_i % 2 == 1 and c.value is not None:
				c.fill = zebra_fill
			h_lbl = headers[c_i - 1]
			c.alignment = left_al if ("Ad Soyad" in h_lbl or "Not" in h_lbl or "Görev" in h_lbl) else center_al

	for i, w in enumerate(col_widths, start=1):
		ws.column_dimensions[get_column_letter(i)].width = w

	for col_idx_1, formula_list in validations:
		col_ltr = get_column_letter(col_idx_1)
		dv = DataValidation(type="list", formula1=f'"{formula_list}"', allow_blank=True)
		ws.add_data_validation(dv)
		dv.add(f"{col_ltr}2:{col_ltr}500")

	ws_info = wb.create_sheet("Bilgi_ve_Kurallar")
	ws_info.column_dimensions["A"].width = 95
	title_cell = ws_info.cell(row=1, column=1, value="📌 Okul Sağlık Taraması Şablon Kullanım Kılavuzu")
	title_cell.font = Font(name="Segoe UI", size=12, bold=True, color=accent_hex)
	for idx_l, ln in enumerate(rules_lines, start=3):
		c_inf = ws_info.cell(row=idx_l, column=1, value=ln)
		c_inf.font = Font(name="Segoe UI", size=10)

	buf = io.BytesIO()
	wb.save(buf)
	return buf.getvalue()


def build_student_screening_excel_template(
	connect_db_fn: Callable | None = None,
	*,
	include_roster: bool = False,
	target_class: str | None = None,
) -> bytes:
	from openpyxl import Workbook

	wb = Workbook()
	ws = wb.active
	ws.title = "Ogrenci_Taramasi"
	ws.append(STUDENT_TEMPLATE_COLUMNS)

	roster_written = 0
	if include_roster and callable(connect_db_fn):
		with connect_db_fn() as conn:
			if target_class and target_class != "Sivil / Diğer":
				rows = conn.execute(
					"""SELECT c.name AS cls_name, s.student_no, s.full_name, s.sex, s.birth_date
					   FROM students s
					   JOIN classes c ON c.id = s.class_id
					   WHERE c.name = ? AND COALESCE(s.person_type, 'Öğrenci') = 'Öğrenci'
					   ORDER BY CAST(s.student_no AS INTEGER), s.student_no, s.full_name""",
					(target_class,),
				).fetchall()
			else:
				rows = conn.execute(
					"""SELECT c.name AS cls_name, s.student_no, s.full_name, s.sex, s.birth_date
					   FROM students s
					   JOIN classes c ON c.id = s.class_id
					   WHERE c.name != 'Sivil / Diğer' AND COALESCE(s.person_type, 'Öğrenci') = 'Öğrenci'
					   ORDER BY c.name, CAST(s.student_no AS INTEGER), s.student_no, s.full_name"""
				).fetchall()
		for r in rows:
			ws.append([
				r["cls_name"] or "",
				r["student_no"] or "",
				r["full_name"] or "",
				r["sex"] or "",
				r["birth_date"] or "",
				None, None, None, None, None, None,
				None, None, None, None, None, None, None,
			])
			roster_written += 1

	if roster_written == 0:
		cls_ex = target_class or "4-A"
		samples = [
			[cls_ex, "101", "Ali Asaf Doğan", "E", "2016-09-10", 138.0, 32.5, None, 36.6, 84, 98, "Hayır (Kullanmıyor)", "Tam / Normal (10/10)", "Tam / Normal (10/10)", "Temiz / Parazit Yok", "Normal / İyi Duyuyor", "Normal / İyi Duyuyor", ""],
			[cls_ex, "102", "Zeynep Bal", "K", "2016-05-31", 135.5, 30.2, None, 36.5, 82, 99, "Evet (Sürekli)", "Gözlüklü Tam", "Gözlüklü Tam", "Temiz / Parazit Yok", "Normal / İyi Duyuyor", "Normal / İyi Duyuyor", ""],
		]
		for s in samples:
			ws.append(s)

	widths = [14, 11, 26, 10, 14, 12, 12, 16, 11, 12, 11, 20, 18, 18, 20, 18, 18, 28]
	rules = [
		"1. 'Ad Soyad' dışındaki tüm sütunlar isteğe bağlıdır; ölçemediğiniz alanları boş bırakabilir veya '-' yazabilirsiniz.",
		"2. Boy (cm): 138, 138.5 veya 1,38 biçiminde yazılabilir (metre girilirse otomatik cm'ye çevrilir).",
		"3. Kilo (kg): 32.5 veya 32,5 biçiminde virgülle/noktayla yazılabilir.",
		"4. Kayıtlı olmayan öğrenciler otomatik oluşturulur; kayıtlı öğrencilerin ise tarama ve persantil verileri güncellenir.",
		"5. Yalnızca belirli sütunlardan oluşan sade bir tablo (ör. sadece Sınıf, Ad Soyad, Boy, Kilo) yükleseniz de sistem otomatik tanır.",
	]
	validations = [
		(4, "E,K"),
		(12, "Hayır (Kullanmıyor),Evet (Sürekli),Bazen / Sadece Okurken,Gözlük Önerildi / Muayene Bekleniyor"),
		(15, "Temiz / Parazit Yok,Kepeklenme / Kuruluk,Sirke Tespit Edildi,Canlı Bit Tespit Edildi"),
	]
	return _style_screening_workbook(wb, ws, STUDENT_TEMPLATE_COLUMNS, widths, "0D9488", rules, validations)


def build_adult_screening_excel_template(
	connect_db_fn: Callable | None = None,
	*,
	include_roster: bool = False,
) -> bytes:
	from openpyxl import Workbook

	wb = Workbook()
	ws = wb.active
	ws.title = "Yetiskin_Taramasi"
	ws.append(ADULT_TEMPLATE_COLUMNS)

	roster_written = 0
	if include_roster and callable(connect_db_fn):
		with connect_db_fn() as conn:
			rows = conn.execute(
				"""SELECT s.full_name, s.notes AS role_title, s.sex, s.birth_date
				   FROM students s
				   LEFT JOIN classes c ON c.id = s.class_id
				   WHERE c.name = 'Sivil / Diğer'
				      OR (s.person_type IS NOT NULL AND trim(s.person_type) != '' AND trim(s.person_type) != 'Öğrenci')
				   ORDER BY s.full_name COLLATE NOCASE"""
			).fetchall()
		for r in rows:
			ws.append([
				r["full_name"] or "",
				str(r["role_title"] or "").strip() or "Öğretmen",
				r["sex"] or "",
				r["birth_date"] or "",
				None, None, None, None, "Açlık", None, None, None, "Normal", "",
			])
			roster_written += 1

	if roster_written == 0:
		samples = [
			["Müberra Görgü", "Öğretmen", "K", "1985-04-12", 168, 77.6, "115-60", 96, "Açlık", 89, 36.6, 98, "Normal", ""],
			["Şebnem Ülker", "Personel", "K", "1988-09-20", 159, 63.9, "100-60", 103, "Açlık", 68, None, None, "Gözlük / Lens Kullanıyor", ""],
		]
		for s in samples:
			ws.append(s)

	widths = [26, 20, 10, 14, 12, 12, 16, 17, 13, 12, 11, 11, 22, 30]
	rules = [
		"1. 'Ad Soyad' zorunludur; sistemde kaydı olmayan yetişkin/personel için otomatik yetişkin profili ('Sivil / Diğer') oluşturulur.",
		"2. Tansiyon (mmHg): '120/80', '115-60', '12/8' veya '12080' biçimlerinin tümü otomatik kabul edilir.",
		"3. Kan Şekeri (mg/dL), Boy (cm), Kilo (kg) ve Nabız (/dk) alanlarında boş veya '-' olan hücreler güvenle atlanır.",
		"4. Yalnızca 'Ad Soyad, Boy, Kilo, Tansiyon, Şeker, Nabız' sütunlarından oluşan sade listeler de doğrudan yüklenebilir veya yapıştırılabilir.",
	]
	validations = [
		(3, "K,E"),
		(9, "Açlık,Tokluk,Rastgele"),
		(13, "Normal,Gözlük / Lens Kullanıyor,Kontrol Önerildi"),
	]
	return _style_screening_workbook(wb, ws, ADULT_TEMPLATE_COLUMNS, widths, "1E40AF", rules, validations)


# ==============================================================================
# 2. AKILLI TABLO, SARILMIŞ SATIR (WRAPPED-LINE), OCR & AI OKUMA MOTORU
# ==============================================================================

_HEADER_SYNONYMS: dict[str, tuple[str, ...]] = {
	"class_name": ("sinif", "sube", "sinifsube", "class", "grade"),
	"student_no": ("okulno", "ogrencino", "numara", "studentno"),
	"full_name": ("adsoyad", "adisoyadi", "isim", "isimsoyisim", "ogrenciadi", "personeladi", "kisi", "fullname", "name"),
	"first_name": ("ad", "adi", "firstname"),
	"last_name": ("soyad", "soyadi", "lastname"),
	"role": ("gorev", "brans", "unvan", "statu", "gorevbrans", "meslek", "birim", "departman", "role"),
	"sex": ("cinsiyet", "sex", "gender"),
	"birth_date": ("dogumtarihi", "dogum", "dogumyili", "dogumgunu", "bdate", "birthdate", "dob"),
	"height_cm": ("boy", "boycm", "boyu", "height", "heightcm"),
	"weight_kg": ("kilo", "kilokg", "agirlik", "kilosu", "weight", "weightkg"),
	"head_circ_cm": ("bascevresi", "bascevresicm", "headcirc"),
	"bp": ("tansiyon", "tansiyonmmhg", "kanbasinci", "ta", "kb", "bp", "bloodpressure", "sistolikdiyastolik"),
	"bp_sys": ("sistolik", "buyuktansiyon", "sys", "systolic"),
	"bp_dia": ("diyastolik", "kucuktansiyon", "dia", "diastolic"),
	"glucose_mg_dl": ("seker", "kansekeri", "kansekerimgdl", "ks", "aks", "tks", "glukoz", "glikoz", "glucose", "bg"),
	"glucose_context": ("kszaman", "kszamani", "sekerdurumu", "acliktokluk", "olcumzamani"),
	"pulse_bpm": ("nabiz", "nabizdk", "kalphizi", "kalpatisi", "pulse", "pulsebpm", "hr"),
	"temperature_c": ("ates", "atesc", "vucutisisi", "sicaklik", "temp", "temperature"),
	"spo2_pct": ("spo2", "oksijen", "saturasyon", "oksijensaturasyonu"),
	"wears_glasses": ("gozluk", "gozlukkullanimi", "lens"),
	"vision_right": ("saggoz", "gorme", "gormedurumu", "goz", "vision"),
	"vision_left": ("solgoz", "visionleft"),
	"hair_screening": ("sac", "sactaramasi", "sachijyen", "hijyen", "pedikuloz", "bit"),
	"hearing_right": ("sagkulak", "isitme", "kulak", "hearing"),
	"hearing_left": ("solkulak", "hearingleft"),
	"health_behavior": ("saglikdavranisi", "beslenme"),
	"psychosocial_status": ("psikososyal", "psikososyaldurum", "uyum"),
	"notes": ("not", "notlar", "taramanotu", "kronik", "kroniktaramanotu", "aciklama", "sevk"),
}


def map_column_header(raw_col: object) -> str | None:
	ck = tr_col_key(raw_col)
	if not ck:
		return None
	# Tam eşleşme öncelikli
	for canonical, syns in _HEADER_SYNONYMS.items():
		if ck in syns:
			return canonical
	# Kapsama eşleşmesi (uzun anahtarlar önce)
	for canonical, syns in _HEADER_SYNONYMS.items():
		if canonical in ("first_name", "last_name"):
			continue
		for syn in syns:
			if len(syn) >= 3 and syn in ck:
				return canonical
	return None


def _looks_like_orphan_continuation(tokens: list[str]) -> bool:
	"""Alt satıra kaymış tekil sayısal/tansiyon değeri mi (ör. Esra Erboy satırından taşan '76') kontrol eder."""
	non_empty = [t.strip() for t in tokens if t.strip()]
	if not non_empty or len(non_empty) > 2:
		return False
	for tok in non_empty:
		if is_placeholder_or_blank(tok):
			continue
		# Harf içeriyorsa kişi adıdır, devam satırı değildir (yalnızca mmHg/mg/cm/kg birimleri hariç)
		alpha_clean = re.sub(r"\b(cm|kg|gr|mmhg|mg|dl|dk|bpm)\b", "", tr_fold(tok))
		if re.search(r"[a-z]", alpha_clean):
			return False
	return True


def stitch_and_parse_delimited_text(raw_text: str) -> pd.DataFrame | None:
	"""
	Excel / Word / WhatsApp / Sohbetten yapıştırılan sekme, noktalı virgül, boru (|) veya çoklu boşluk
	ayrımlı metinleri okur; alt satıra kaymış (wrapped) hücreleri bir önceki satırla otomatik birleştirir.
	"""
	if not raw_text or not str(raw_text).strip():
		return None
	lines = [ln.rstrip("\r") for ln in str(raw_text).splitlines() if ln.strip()]
	if len(lines) < 2:
		return None

	# Ayırıcı tespiti
	sample_top = "\n".join(lines[:6])
	if "\t" in sample_top:
		splitter = lambda s: [c.strip() for c in s.split("\t")]
	elif "|" in sample_top and sample_top.count("|") >= 2:
		splitter = lambda s: [c.strip() for c in s.strip("| ").split("|")]
	elif ";" in sample_top and sample_top.count(";") >= 2:
		splitter = lambda s: [c.strip() for c in s.split(";")]
	elif "," in sample_top and lines[0].count(",") >= 3:
		# Ondalık virgülle karışmaması için yalnızca başlıkta >=3 virgül varsa ve boşluklu tablo değilse
		splitter = lambda s: [c.strip() for c in s.split(",")]
	else:
		splitter = lambda s: [c.strip() for c in re.split(r"\s{2,}", s.strip())]

	tokenized_rows: list[list[str]] = [splitter(ln) for ln in lines]

	# Başlık satırını ilk 8 satır içinde bul
	hdr_idx = -1
	expected_cols = 0
	for i, r_toks in enumerate(tokenized_rows[:8]):
		hits = sum(1 for cell in r_toks if map_column_header(cell) is not None)
		if hits >= 2:
			hdr_idx = i
			# Sondaki tamamen boş başlıkları kırp
			while r_toks and not r_toks[-1].strip():
				r_toks.pop()
			expected_cols = len(r_toks)
			break

	if hdr_idx == -1:
		# Başlıksız yapıştırma: ör. "Müberra Görgü   168   77,6   115-60   96   89"
		stitched_no_hdr: list[list[str]] = []
		for r_toks in tokenized_rows:
			clean_r = [t for t in r_toks if t != ""]
			if not clean_r:
				continue
			# Tek satır boşlukla ayrılmışsa ("Ali Kaya 168 77,6 115-60 96 89") akıllı böl
			if len(clean_r) == 1:
				parts_ws = r_toks[0].split()
				name_w: list[str] = []
				rest_w: list[str] = []
				for pw in parts_ws:
					if not rest_w and re.search(r"[A-Za-zÇĞİÖŞÜçğıöşü]", pw) and not re.search(r"\d", pw):
						name_w.append(pw)
					else:
						rest_w.append(pw)
				if name_w and rest_w:
					clean_r = [" ".join(name_w), *rest_w]
			if stitched_no_hdr and _looks_like_orphan_continuation(clean_r):
				stitched_no_hdr[-1].extend(clean_r)
			else:
				stitched_no_hdr.append(clean_r)
		if not stitched_no_hdr:
			return None
		max_w = max(len(r) for r in stitched_no_hdr)
		if max_w < 2:
			return None
		padded = [r + [""] * (max_w - len(r)) for r in stitched_no_hdr]
		return pd.DataFrame(padded, columns=[f"Sutun_{i+1}" for i in range(max_w)])

	headers = [c if c else f"Sutun_{j+1}" for j, c in enumerate(tokenized_rows[hdr_idx][:expected_cols])]
	stitched: list[list[str]] = []
	for r_toks in tokenized_rows[hdr_idx + 1 :]:
		# Sondaki boş sekme hücrelerini kırp ki eksik sütun sayısı doğru görünsün
		trimmed = list(r_toks)
		while trimmed and trimmed[-1] == "":
			trimmed.pop()
		if not trimmed:
			continue
		if stitched and len(stitched[-1]) < expected_cols and _looks_like_orphan_continuation(trimmed):
			stitched[-1].extend(trimmed)
			continue
		# Sohbet mesajlarında tablonun altına yazılan serbest cümleleri ("yetişkinlere profil oluşturup...") atla
		if expected_cols >= 3 and len(trimmed) == 1 and not _looks_like_orphan_continuation(trimmed):
			continue
		stitched.append(trimmed)

	if not stitched:
		return None
	norm_rows = [
		(r[:expected_cols] + [""] * max(0, expected_cols - len(r)))
		for r in stitched
	]
	return pd.DataFrame(norm_rows, columns=headers)


def find_and_promote_header_row(df_raw: pd.DataFrame) -> pd.DataFrame:
	"""Üstte kurum başlığı / boş satırlar bulunan Excel veya Word tablolarında gerçek başlık satırını bulup sütun yapar."""
	if df_raw is None or df_raw.empty:
		return df_raw
	cur_hits = sum(1 for c in df_raw.columns if map_column_header(c) is not None)
	if cur_hits >= 2:
		return df_raw

	best_idx = -1
	best_hits = cur_hits
	scan_n = min(15, len(df_raw))
	for i in range(scan_n):
		vals = [str(v).strip() if not is_placeholder_or_blank(v) else "" for v in df_raw.iloc[i].tolist()]
		hits = sum(1 for v in vals if v and map_column_header(v) is not None)
		if hits > best_hits and hits >= 2:
			best_hits = hits
			best_idx = i

	if best_idx >= 0:
		new_cols: list[str] = []
		seen_c: dict[str, int] = {}
		for j, v in enumerate(df_raw.iloc[best_idx].tolist()):
			base_c = str(v).strip() if not is_placeholder_or_blank(v) else f"Sutun_{j+1}"
			cnt = seen_c.get(base_c, 0)
			seen_c[base_c] = cnt + 1
			new_cols.append(base_c if cnt == 0 else f"{base_c}_{cnt+1}")
		sub_df = df_raw.iloc[best_idx + 1 :].copy().reset_index(drop=True)
		sub_df.columns = new_cols
		return sub_df
	return df_raw


def extract_records_from_dataframe(
	df_in: pd.DataFrame,
	*,
	target_group: str = "auto",
	fallback_class: str = "",
	default_role: str = "Öğretmen / Personel",
) -> tuple[list[dict[str, Any]], str]:
	"""DataFrame içinden öğrenci veya yetişkin tarama kayıtlarını çıkarır ve algılanan grubu ('student'|'adult') döndürür."""
	if df_in is None or df_in.empty:
		return [], (target_group if target_group in ("student", "adult") else "student")

	df = find_and_promote_header_row(df_in)
	col_map: dict[str, str] = {}
	for col in df.columns:
		canon = map_column_header(col)
		if canon and canon not in col_map:
			col_map[canon] = col

	# Başlıksız tablo sezgisel sütun eşlemesi
	if "full_name" not in col_map and "first_name" not in col_map and len(df.columns) >= 2:
		cols = list(df.columns)
		sample_r = df.iloc[0].tolist()
		has_bp_like = any(
			(not is_placeholder_or_blank(v)) and bool(re.search(r"^\d{2,3}\s*[/\-]\s*\d{2,3}$", str(v).strip()))
			for v in sample_r
		)
		if has_bp_like or target_group == "adult":
			col_map["full_name"] = cols[0]
			if len(cols) >= 2:
				col_map["height_cm"] = cols[1]
			if len(cols) >= 3:
				col_map["weight_kg"] = cols[2]
			if len(cols) >= 4:
				col_map["bp"] = cols[3]
			if len(cols) >= 5:
				col_map["glucose_mg_dl"] = cols[4]
			if len(cols) >= 6:
				col_map["pulse_bpm"] = cols[5]
		elif len(cols) >= 6 and not fallback_class:
			col_map["class_name"] = cols[0]
			col_map["full_name"] = cols[1]
			col_map["sex"] = cols[2]
			col_map["birth_date"] = cols[3]
			col_map["height_cm"] = cols[4]
			col_map["weight_kg"] = cols[5]
		elif len(cols) >= 5 and fallback_class:
			col_map["full_name"] = cols[0]
			col_map["sex"] = cols[1]
			col_map["birth_date"] = cols[2]
			col_map["height_cm"] = cols[3]
			col_map["weight_kg"] = cols[4]
		else:
			# 2-4 sütunlu yalın liste: Ad Soyad, Boy, Kilo (veya Sınıf, Ad Soyad, Boy, Kilo)
			first_txt = str(sample_r[0] or "").strip()
			if re.match(r"^\d{1,2}\s*[-/]\s*[A-Za-zÇĞİÖŞÜ]$", first_txt) and len(cols) >= 3:
				col_map["class_name"] = cols[0]
				col_map["full_name"] = cols[1]
				if len(cols) >= 3:
					col_map["height_cm"] = cols[2]
				if len(cols) >= 4:
					col_map["weight_kg"] = cols[3]
			else:
				col_map["full_name"] = cols[0]
				if len(cols) >= 2:
					col_map["height_cm"] = cols[1]
				if len(cols) >= 3:
					col_map["weight_kg"] = cols[2]

	# Otomatik grup algılama
	if target_group in ("adult", "student"):
		resolved_group = target_group
	else:
		has_adult_cols = any(k in col_map for k in ("role", "bp", "bp_sys", "bp_dia", "glucose_mg_dl"))
		has_stu_cols = any(k in col_map for k in ("class_name", "student_no", "head_circ_cm", "hair_screening")) or bool(fallback_class)
		resolved_group = "adult" if (has_adult_cols and not has_stu_cols) else "student"

	is_adult = (resolved_group == "adult")
	records: list[dict[str, Any]] = []

	for _, row in df.iterrows():
		fn = clean_cell_str(row.get(col_map["full_name"])) if "full_name" in col_map else None
		if not fn and ("first_name" in col_map or "last_name" in col_map):
			p1 = clean_cell_str(row.get(col_map.get("first_name", ""))) or ""
			p2 = clean_cell_str(row.get(col_map.get("last_name", ""))) or ""
			fn = f"{p1} {p2}".strip() or None
		if not fn:
			continue
		fn_fold = tr_fold(fn)
		if fn_fold in {"ad soyad", "adi soyadi", "isim", "ogrenci adi", "personel adi", "toplam", "ortalama"}:
			continue
		# Sadece rakamdan oluşan sıra numarası satırlarını atla
		if fn.replace(".", "").isdigit():
			continue

		cls_val = (clean_cell_str(row.get(col_map["class_name"])) if "class_name" in col_map else None) or fallback_class or ""
		sno_val = clean_cell_str(row.get(col_map["student_no"])) if "student_no" in col_map else None
		if sno_val and sno_val.endswith(".0") and sno_val[:-2].isdigit():
			sno_val = sno_val[:-2]
		role_val = (clean_cell_str(row.get(col_map["role"])) if "role" in col_map else None) or (default_role if is_adult else "")
		sex_val = normalize_sex_smart(row.get(col_map["sex"])) if "sex" in col_map else None
		dob_val = normalize_date_smart(row.get(col_map["birth_date"])) if "birth_date" in col_map else None

		h_val = normalize_height_smart(row.get(col_map["height_cm"])) if "height_cm" in col_map else None
		w_val = normalize_weight_smart(row.get(col_map["weight_kg"]), is_adult=is_adult) if "weight_kg" in col_map else None
		hc_val = normalize_head_circ_smart(row.get(col_map["head_circ_cm"])) if "head_circ_cm" in col_map else None

		sys_v, dia_v = (None, None)
		if "bp" in col_map:
			sys_v, dia_v = normalize_bp_smart(row.get(col_map["bp"]))
		elif "bp_sys" in col_map or "bp_dia" in col_map:
			sys_v, dia_v = normalize_bp_smart(
				row.get(col_map.get("bp_sys", "")),
				row.get(col_map.get("bp_dia", "")),
			)

		glu_v = normalize_glucose_smart(row.get(col_map["glucose_mg_dl"])) if "glucose_mg_dl" in col_map else None
		gctx_v = (clean_cell_str(row.get(col_map["glucose_context"])) if "glucose_context" in col_map else None) or ("Açlık" if glu_v is not None else None)
		pls_v = normalize_pulse_smart(row.get(col_map["pulse_bpm"])) if "pulse_bpm" in col_map else None
		tmp_v = normalize_temp_smart(row.get(col_map["temperature_c"])) if "temperature_c" in col_map else None
		spo_v = normalize_spo2_smart(row.get(col_map["spo2_pct"])) if "spo2_pct" in col_map else None

		wears_v = clean_cell_str(row.get(col_map["wears_glasses"])) if "wears_glasses" in col_map else None
		vr_v = clean_cell_str(row.get(col_map["vision_right"])) if "vision_right" in col_map else None
		vl_v = clean_cell_str(row.get(col_map["vision_left"])) if "vision_left" in col_map else None
		hair_v = clean_cell_str(row.get(col_map["hair_screening"])) if "hair_screening" in col_map else None
		hr_v = clean_cell_str(row.get(col_map["hearing_right"])) if "hearing_right" in col_map else None
		hl_v = clean_cell_str(row.get(col_map["hearing_left"])) if "hearing_left" in col_map else None
		beh_v = clean_cell_str(row.get(col_map["health_behavior"])) if "health_behavior" in col_map else None
		psy_v = clean_cell_str(row.get(col_map["psychosocial_status"])) if "psychosocial_status" in col_map else None
		notes_v = clean_cell_str(row.get(col_map["notes"])) if "notes" in col_map else None

		records.append({
			"person_group": "adult" if is_adult else "student",
			"class_name": "Sivil / Diğer" if is_adult else cls_val,
			"student_no": None if is_adult else sno_val,
			"role": role_val if is_adult else "",
			"full_name": fn,
			"sex": sex_val,
			"birth_date": dob_val,
			"height_cm": h_val,
			"weight_kg": w_val,
			"head_circ_cm": None if is_adult else hc_val,
			"bp_sys": sys_v,
			"bp_dia": dia_v,
			"glucose_mg_dl": glu_v,
			"glucose_context": gctx_v,
			"pulse_bpm": pls_v,
			"temperature_c": tmp_v,
			"spo2_pct": spo_v,
			"wears_glasses": wears_v,
			"vision_right": vr_v,
			"vision_left": vl_v or (vr_v if is_adult and vr_v else None),
			"hair_screening": hair_v,
			"hearing_right": hr_v,
			"hearing_left": hl_v,
			"health_behavior": beh_v,
			"psychosocial_status": psy_v,
			"notes": notes_v,
		})

	return records, resolved_group


def preprocess_image_for_ocr_and_vision(image_bytes: bytes) -> tuple[bytes, Any]:
	"""Görselin EXIF yönünü düzeltir, kontrastı iyileştirir ve OCR/Vision için ideal çözünürlüğe getirir."""
	from PIL import Image, ImageOps

	img = Image.open(io.BytesIO(image_bytes))
	try:
		img = ImageOps.exif_transpose(img)
	except Exception:
		pass
	img = img.convert("RGB")
	w, h = img.size
	if max(w, h) < 1100:
		scale = 1400.0 / float(max(w, h, 1))
		img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
	elif max(w, h) > 2400:
		scale = 2200.0 / float(max(w, h))
		img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
	try:
		img = ImageOps.autocontrast(img, cutoff=1)
	except Exception:
		pass
	out_b = io.BytesIO()
	img.save(out_b, format="PNG")
	return out_b.getvalue(), img


def parse_image_screening_ocr_v2(
	image_bytes: bytes,
	ocr_reader_fn: Callable[[], Any],
	*,
	target_group: str = "auto",
	fallback_class: str = "",
) -> tuple[list[dict[str, Any]], str]:
	"""Geliştirilmiş yerel OCR tablo satır/sütun okuyucu (hem öğrenci hem yetişkin tansiyon/şeker/nabız destekli)."""
	import numpy as np

	_clean_png, pil_img = preprocess_image_for_ocr_and_vision(image_bytes)
	reader = ocr_reader_fn()
	results = reader.readtext(np.asarray(pil_img), detail=1)
	items: list[dict[str, Any]] = []
	for bbox, text, prob in results:
		t = str(text or "").strip()
		if not t:
			continue
		ys = [p[1] for p in bbox]
		xs = [p[0] for p in bbox]
		items.append({
			"text": t,
			"y": sum(ys) / len(ys),
			"h": max(ys) - min(ys),
			"min_x": min(xs),
			"prob": prob,
		})
	if not items:
		return [], (target_group if target_group in ("student", "adult") else "student")

	items.sort(key=lambda it: it["y"])
	med_h = sorted(it["h"] for it in items)[len(items) // 2]
	row_thresh = max(12.0, med_h * 0.62)

	rows: list[list[dict[str, Any]]] = []
	cur_r = [items[0]]
	for it in items[1:]:
		if abs(it["y"] - (sum(x["y"] for x in cur_r) / len(cur_r))) <= row_thresh:
			cur_r.append(it)
		else:
			cur_r.sort(key=lambda x: x["min_x"])
			rows.append(cur_r)
			cur_r = [it]
	if cur_r:
		cur_r.sort(key=lambda x: x["min_x"])
		rows.append(cur_r)

	tsv_lines = ["\t".join(it["text"] for it in r) for r in rows]
	df_try = stitch_and_parse_delimited_text("\n".join(tsv_lines))
	if df_try is not None and not df_try.empty:
		recs, grp = extract_records_from_dataframe(df_try, target_group=target_group, fallback_class=fallback_class)
		if recs:
			return recs, grp

	# Başlık yakalanamadıysa satır içi akıllı ayrıştırma
	is_adult_hint = (target_group == "adult") or any(
		re.search(r"\b\d{2,3}\s*[/\-]\s*\d{2,3}\b", " ".join(it["text"] for it in r)) for r in rows
	)
	out_recs: list[dict[str, Any]] = []
	for r in rows:
		tokens = [it["text"].strip() for it in r if it["text"].strip()]
		joined_fold = tr_fold(" ".join(tokens))
		if any(h in joined_fold for h in ("ad soyad", "boy", "tansiyon", "kan sekeri")):
			continue
		name_parts: list[str] = []
		nums: list[float] = []
		sys_v, dia_v = (None, None)
		sex_v = None
		dob_v = None
		for tok in tokens:
			if is_placeholder_or_blank(tok):
				continue
			if sys_v is None and re.match(r"^\d{2,3}\s*[/\-]\s*\d{2,3}$", tok):
				sys_v, dia_v = normalize_bp_smart(tok)
				continue
			dt_c = normalize_date_smart(tok)
			if dt_c and not dob_v and len(tok) >= 8:
				dob_v = dt_c
				continue
			sx_c = normalize_sex_smart(tok)
			if sx_c and not sex_v and len(tok) <= 5:
				sex_v = sx_c
				continue
			nf = parse_flex_float(tok)
			if nf is not None:
				nums.append(nf)
				continue
			name_parts.append(tok)
		fn = " ".join(name_parts).strip()
		if not fn or len(fn) < 3:
			continue
		h_v = normalize_height_smart(nums[0]) if len(nums) >= 1 else None
		w_v = normalize_weight_smart(nums[1], is_adult=is_adult_hint) if len(nums) >= 2 else None
		if h_v is None and w_v is None and len(nums) >= 2:
			h_v = normalize_height_smart(max(nums[:2]))
			w_v = normalize_weight_smart(min(nums[:2]), is_adult=is_adult_hint)
		glu_v = normalize_glucose_smart(nums[2]) if (is_adult_hint and len(nums) >= 3) else None
		pls_v = normalize_pulse_smart(nums[3] if (is_adult_hint and len(nums) >= 4) else (nums[2] if (not is_adult_hint and len(nums) >= 3) else None))
		out_recs.append({
			"person_group": "adult" if is_adult_hint else "student",
			"class_name": "Sivil / Diğer" if is_adult_hint else fallback_class,
			"student_no": None,
			"role": "Personel / Öğretmen" if is_adult_hint else "",
			"full_name": fn,
			"sex": sex_v,
			"birth_date": dob_v,
			"height_cm": h_v,
			"weight_kg": w_v,
			"head_circ_cm": None,
			"bp_sys": sys_v,
			"bp_dia": dia_v,
			"glucose_mg_dl": glu_v,
			"glucose_context": "Açlık" if glu_v is not None else None,
			"pulse_bpm": pls_v,
			"temperature_c": None,
			"spo2_pct": None,
			"wears_glasses": None,
			"vision_right": None,
			"vision_left": None,
			"hair_screening": None,
			"hearing_right": None,
			"hearing_left": None,
			"health_behavior": None,
			"psychosocial_status": None,
			"notes": None,
		})
	return out_recs, ("adult" if is_adult_hint else "student")


def parse_smart_screening_source(
	*,
	file_bytes: bytes | None = None,
	file_name: str = "",
	raw_text: str = "",
	target_group: str = "auto",
	fallback_class: str = "",
	default_role: str = "Öğretmen / Personel",
	ocr_reader_fn: Callable[[], Any] | None = None,
	extract_docx_fn: Callable[[bytes], tuple[list[tuple[str, pd.DataFrame]], list[str]]] | None = None,
) -> tuple[list[dict[str, Any]], str, str]:
	"""
	Birleşik deterministik + yerel OCR tarama okuyucu.
	Dönüş: (kayıt_listesi, algılanan_grup ['student'|'adult'], kaynak_açıklaması)
	"""
	if raw_text and str(raw_text).strip() and not file_bytes:
		df_txt = stitch_and_parse_delimited_text(raw_text)
		if df_txt is not None and not df_txt.empty:
			recs, grp = extract_records_from_dataframe(
				df_txt, target_group=target_group, fallback_class=fallback_class, default_role=default_role
			)
			return recs, grp, f"Yapıştırılan metinden {len(recs)} satır okundu."
		return [], (target_group if target_group in ("student", "adult") else "student"), "Metinde tablo satırı bulunamadı."

	if not file_bytes:
		return [], "student", "Dosya verisi boş."

	low_fn = (file_name or "").lower().strip()

	if low_fn.endswith((".png", ".jpg", ".jpeg", ".bmp", ".webp")):
		if callable(ocr_reader_fn):
			recs, grp = parse_image_screening_ocr_v2(
				file_bytes, ocr_reader_fn, target_group=target_group, fallback_class=fallback_class
			)
			return recs, grp, f"Görsel OCR ile {len(recs)} satır okundu."
		return [], (target_group if target_group in ("student", "adult") else "student"), "Görsel okumak için Gemini AI veya OCR gerekli."

	if low_fn.endswith(".pdf"):
		try:
			import pymupdf
			all_recs_pdf: list[dict[str, Any]] = []
			det_grp = target_group if target_group in ("student", "adult") else "student"
			with pymupdf.open(stream=file_bytes, filetype="pdf") as doc:
				full_txt_pages: list[str] = []
				for page in doc:
					try:
						tabs = page.find_tables()
						for tb in (tabs.tables if tabs else []):
							df_tb = pd.DataFrame(tb.extract())
							r_p, g_p = extract_records_from_dataframe(
								df_tb, target_group=target_group, fallback_class=fallback_class, default_role=default_role
							)
							if r_p:
								all_recs_pdf.extend(r_p)
								det_grp = g_p
					except Exception:
						pass
					ptxt = (page.get_text("text") or "").strip()
					if ptxt:
						full_txt_pages.append(ptxt)
				if not all_recs_pdf and full_txt_pages:
					df_pt = stitch_and_parse_delimited_text("\n".join(full_txt_pages))
					if df_pt is not None:
						all_recs_pdf, det_grp = extract_records_from_dataframe(
							df_pt, target_group=target_group, fallback_class=fallback_class, default_role=default_role
						)
				if not all_recs_pdf and callable(ocr_reader_fn) and len(doc) > 0:
					pix = doc[0].get_pixmap(dpi=200)
					all_recs_pdf, det_grp = parse_image_screening_ocr_v2(
						pix.tobytes("png"), ocr_reader_fn, target_group=target_group, fallback_class=fallback_class
					)
			return all_recs_pdf, det_grp, f"PDF belgesinden {len(all_recs_pdf)} satır okundu."
		except Exception as exc:
			return [], "student", f"PDF okuma hatası: {exc}"

	if low_fn.endswith(".docx"):
		all_recs_w: list[dict[str, Any]] = []
		det_grp = target_group if target_group in ("student", "adult") else "student"
		if callable(extract_docx_fn):
			tables_found, para_lines = extract_docx_fn(file_bytes)
		else:
			import docx as _docx
			doc_o = _docx.Document(io.BytesIO(file_bytes))
			tables_found = []
			for tbl in doc_o.tables:
				r_vals = []
				for row in tbl.rows:
					seen_tc = set()
					cells = []
					for c in row.cells:
						tc_id = id(getattr(c, "_tc", c))
						if tc_id in seen_tc:
							continue
						seen_tc.add(tc_id)
						cells.append(re.sub(r"\s+", " ", (c.text or "").strip()))
					if any(cells):
						r_vals.append(cells)
				if r_vals:
					max_c = max(len(x) for x in r_vals)
					padded = [x + [""] * (max_c - len(x)) for x in r_vals]
					tables_found.append(("", pd.DataFrame(padded)))
			para_lines = [(p.text or "").strip() for p in doc_o.paragraphs if (p.text or "").strip()]

		for cls_h, df_t in tables_found:
			r_t, g_t = extract_records_from_dataframe(
				df_t, target_group=target_group, fallback_class=cls_h or fallback_class, default_role=default_role
			)
			if r_t:
				all_recs_w.extend(r_t)
				det_grp = g_t
		if not all_recs_w and para_lines:
			df_p = stitch_and_parse_delimited_text("\n".join(para_lines))
			if df_p is not None:
				all_recs_w, det_grp = extract_records_from_dataframe(
					df_p, target_group=target_group, fallback_class=fallback_class, default_role=default_role
				)
		return all_recs_w, det_grp, f"Word (.docx) belgesinden {len(all_recs_w)} satır okundu."

	if low_fn.endswith((".xlsx", ".xls")):
		sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=object)
		all_recs_x: list[dict[str, Any]] = []
		det_grp = target_group if target_group in ("student", "adult") else "student"
		for s_name, s_df in (sheets.items() if isinstance(sheets, dict) else [("Sayfa1", sheets)]):
			if s_df is None or s_df.empty:
				continue
			if "bilgi" in tr_fold(s_name) or "kural" in tr_fold(s_name) or "kilavuz" in tr_fold(s_name):
				continue
			cls_from_sheet = ""
			if re.match(r"^(\d{1,2}\s*[-/ ]?\s*[A-Za-zÇĞİÖŞÜ]|anasinifi)", tr_fold(s_name)):
				cls_from_sheet = str(s_name).strip()
			r_s, g_s = extract_records_from_dataframe(
				s_df,
				target_group=target_group,
				fallback_class=fallback_class or cls_from_sheet,
				default_role=default_role,
			)
			if r_s:
				all_recs_x.extend(r_s)
				det_grp = g_s
		return all_recs_x, det_grp, f"Excel dosyasından {len(all_recs_x)} satır okundu."

	if low_fn.endswith((".csv", ".tsv", ".txt")):
		decoded = ""
		for enc in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
			try:
				decoded = file_bytes.decode(enc)
				break
			except Exception:
				continue
		df_c = stitch_and_parse_delimited_text(decoded)
		if df_c is None:
			try:
				df_c = pd.read_csv(io.StringIO(decoded), sep=None, engine="python", dtype=object)
			except Exception:
				df_c = pd.DataFrame()
		recs_c, grp_c = extract_records_from_dataframe(
			df_c, target_group=target_group, fallback_class=fallback_class, default_role=default_role
		)
		return recs_c, grp_c, f"Metin/CSV dosyasından {len(recs_c)} satır okundu."

	raise ValueError(f"Desteklenmeyen dosya türü: {file_name}")


def normalize_ai_extracted_records(
	raw_rows: list[dict[str, Any]],
	*,
	target_group: str = "auto",
	fallback_class: str = "",
	default_role: str = "Öğretmen / Personel",
) -> list[dict[str, Any]]:
	"""Gemini AI'dan (sohbet veya tablo/görsel okuyucu) gelen sözlük listesini standart tarama kaydı şemasına dönüştürür."""
	out: list[dict[str, Any]] = []
	for rr in (raw_rows or []):
		if not isinstance(rr, dict):
			continue
		fn = clean_cell_str(rr.get("full_name") or rr.get("name") or rr.get("ad_soyad"))
		if not fn:
			continue
		p_grp_raw = tr_fold(rr.get("person_group") or rr.get("person_type") or "")
		if target_group in ("adult", "student"):
			is_adult = (target_group == "adult")
		elif p_grp_raw in {"adult", "yetiskin", "sivil", "personel", "ogretmen", "staff"} or "sivil" in p_grp_raw:
			is_adult = True
		elif p_grp_raw in {"student", "ogrenci"}:
			is_adult = False
		else:
			has_adult_f = any(not is_placeholder_or_blank(rr.get(k)) for k in ("bp", "bp_sys", "systolic_bp", "systolic_mmhg", "glucose_mg_dl", "seker", "role"))
			has_stu_f = any(not is_placeholder_or_blank(rr.get(k)) for k in ("student_no", "head_circ_cm", "hair_screening")) or (
				not is_placeholder_or_blank(rr.get("class_name")) and str(rr.get("class_name")).strip() != "Sivil / Diğer"
			)
			is_adult = bool(has_adult_f and not has_stu_f)

		sys_v, dia_v = normalize_bp_smart(
			rr.get("bp") or rr.get("tansiyon") or rr.get("bp_sys") or rr.get("systolic_bp") or rr.get("systolic_mmhg"),
			rr.get("bp_dia") or rr.get("diastolic_bp") or rr.get("diastolic_mmhg"),
		)
		glu_v = normalize_glucose_smart(rr.get("glucose_mg_dl") or rr.get("glucose") or rr.get("seker"))
		vr = clean_cell_str(rr.get("vision_right") or rr.get("vision") or rr.get("gorme"))
		vl = clean_cell_str(rr.get("vision_left")) or (vr if is_adult else None)
		hr = clean_cell_str(rr.get("hearing_right") or rr.get("hearing"))
		hl = clean_cell_str(rr.get("hearing_left")) or hr

		out.append({
			"person_group": "adult" if is_adult else "student",
			"class_name": "Sivil / Diğer" if is_adult else (clean_cell_str(rr.get("class_name")) or fallback_class or ""),
			"student_no": None if is_adult else clean_cell_str(rr.get("student_no")),
			"role": (clean_cell_str(rr.get("role")) or default_role) if is_adult else "",
			"full_name": fn,
			"sex": normalize_sex_smart(rr.get("sex")),
			"birth_date": normalize_date_smart(rr.get("birth_date")),
			"height_cm": normalize_height_smart(rr.get("height_cm") or rr.get("boy")),
			"weight_kg": normalize_weight_smart(rr.get("weight_kg") or rr.get("kilo"), is_adult=is_adult),
			"head_circ_cm": None if is_adult else normalize_head_circ_smart(rr.get("head_circ_cm")),
			"bp_sys": sys_v,
			"bp_dia": dia_v,
			"glucose_mg_dl": glu_v,
			"glucose_context": clean_cell_str(rr.get("glucose_context")) or ("Açlık" if glu_v is not None else None),
			"pulse_bpm": normalize_pulse_smart(rr.get("pulse_bpm") or rr.get("pulse") or rr.get("nabiz")),
			"temperature_c": normalize_temp_smart(rr.get("temperature_c") or rr.get("temperature")),
			"spo2_pct": normalize_spo2_smart(rr.get("spo2_pct") or rr.get("spo2")),
			"wears_glasses": clean_cell_str(rr.get("wears_glasses")),
			"vision_right": vr,
			"vision_left": vl,
			"hair_screening": clean_cell_str(rr.get("hair_screening")),
			"hearing_right": hr,
			"hearing_left": hl,
			"health_behavior": clean_cell_str(rr.get("health_behavior")),
			"psychosocial_status": clean_cell_str(rr.get("psychosocial_status")),
			"notes": clean_cell_str(rr.get("notes")),
		})
	return out


# ==============================================================================
# 3. BULANIK TÜRKÇE İSİM EŞLEŞTİRME & TEK-TRANSACTION HIZLI TOPLU KAYIT
# ==============================================================================

def _find_matching_person(
	rec: dict[str, Any],
	db_people: list[dict[str, Any]],
	by_fold_adult: dict[str, dict[str, Any]],
	by_fold_student: dict[str, dict[str, Any]],
	by_sno: dict[tuple[str, str], dict[str, Any]],
) -> tuple[dict[str, Any] | None, str]:
	is_adult = (rec.get("person_group") == "adult")
	fn = str(rec.get("full_name") or "").strip()
	fn_f = tr_fold(fn)
	if not fn_f:
		return None, "new"

	if not is_adult:
		sno = str(rec.get("student_no") or "").strip()
		cls_f = tr_fold(rec.get("class_name") or "")
		if sno and cls_f and (cls_f, sno) in by_sno:
			return by_sno[(cls_f, sno)], "exact_no"

	primary_map = by_fold_adult if is_adult else by_fold_student
	if fn_f in primary_map:
		return primary_map[fn_f], "exact_name"

	# Yakın OCR / yazım hatası eşleştirmesi (>=0.91 benzerlik ve aynı ilk ad kökü)
	toks = fn_f.split()
	best_p = None
	best_ratio = 0.0
	for p in db_people:
		p_is_adult = bool(p["is_adult"])
		if p_is_adult != is_adult:
			continue
		pf = p["fold_name"]
		if not pf:
			continue
		ptoks = pf.split()
		if toks and ptoks and len(toks) == len(ptoks) and toks[0][:2] == ptoks[0][:2]:
			ratio = SequenceMatcher(None, fn_f, pf).ratio()
			if ratio >= 0.91 and ratio > best_ratio:
				best_ratio = ratio
				best_p = p
	if best_p is not None:
		return best_p, "fuzzy_name"
	return None, "new"


def preview_match_and_alerts(connect_db_fn: Callable, records: list[dict[str, Any]]) -> list[dict[str, Any]]:
	"""Kayıtları veritabanıyla ön-eşleştirir, VKİ ve klinik uyarı rozetlerini üretir."""
	with connect_db_fn() as conn:
		rows = conn.execute(
			"""SELECT s.id, s.full_name, s.student_no, s.sex, s.birth_date, s.person_type, s.notes,
			          c.name AS class_name
			   FROM students s
			   LEFT JOIN classes c ON c.id = s.class_id"""
		).fetchall()

	db_people: list[dict[str, Any]] = []
	by_fold_adult: dict[str, dict[str, Any]] = {}
	by_fold_student: dict[str, dict[str, Any]] = {}
	by_sno: dict[tuple[str, str], dict[str, Any]] = {}
	for r in rows:
		d = dict(r)
		c_nm = str(d.get("class_name") or "").strip()
		p_tp = str(d.get("person_type") or "").strip()
		is_ad = (c_nm == "Sivil / Diğer") or (p_tp != "" and p_tp != "Öğrenci")
		d["is_adult"] = is_ad
		d["fold_name"] = tr_fold(d.get("full_name") or "")
		db_people.append(d)
		if d["fold_name"]:
			if is_ad:
				by_fold_adult[d["fold_name"]] = d
			else:
				by_fold_student[d["fold_name"]] = d
		if (not is_ad) and d.get("student_no") and c_nm:
			by_sno[(tr_fold(c_nm), str(d["student_no"]).strip())] = d

	enriched: list[dict[str, Any]] = []
	for rec in records:
		matched, m_kind = _find_matching_person(rec, db_people, by_fold_adult, by_fold_student, by_sno)
		item = dict(rec)
		item["matched_id"] = int(matched["id"]) if matched else None
		item["match_kind"] = m_kind
		h = item.get("height_cm")
		w = item.get("weight_kg")
		bmi = round(float(w) / ((float(h) / 100.0) ** 2), 1) if (h and w and float(h) > 0) else None
		item["bmi"] = bmi

		alerts: list[str] = []
		sys_v, dia_v = item.get("bp_sys"), item.get("bp_dia")
		if sys_v and dia_v:
			if sys_v >= 180 or dia_v >= 120:
				alerts.append(f"🚨 Kriz TA ({sys_v}/{dia_v})")
			elif sys_v >= 140 or dia_v >= 90:
				alerts.append(f"⚠️ Yüksek TA ({sys_v}/{dia_v})")
			elif sys_v < 90 or dia_v < 60:
				alerts.append(f"⚠️ Düşük TA ({sys_v}/{dia_v})")
		glu = item.get("glucose_mg_dl")
		if glu:
			if glu < 70:
				alerts.append(f"🚨 Düşük KŞ ({glu:g})")
			elif glu >= 140:
				alerts.append(f"⚠️ Yüksek KŞ ({glu:g})")
		pls = item.get("pulse_bpm")
		if pls and (pls >= 120 or pls < 50):
			alerts.append(f"⚠️ Nabız ({pls}/dk)")
		if bmi and item.get("person_group") == "adult" and bmi >= 30.0:
			alerts.append(f"VKİ {bmi:.1f} (Obez)")
		item["clinical_alert"] = " · ".join(alerts)
		enriched.append(item)
	return enriched


def commit_smart_screening_import(
	connect_db_fn: Callable,
	upsert_measurement_conn_fn: Callable,
	upsert_vital_conn_fn: Callable,
	records: list[dict[str, Any]],
	*,
	screening_date: str | None = None,
	screened_by: str | None = None,
	default_class: str | None = None,
	default_role: str = "Öğretmen / Personel",
	default_country: str = "Türkiye",
	default_province: str | None = None,
	default_district: str | None = None,
	on_after_commit_fn: Callable[[], None] | None = None,
) -> dict[str, int]:
	"""
	Öğrenci ve/veya yetişkin tarama listesini TEK BİR SQLite transaction içinde
	profillere, school_screenings, measurements, vital_readings ve health_profiles tablolarına yazar.
	"""
	scr_date = normalize_date_smart(screening_date) or date.today().isoformat()
	scr_date_obj = date.fromisoformat(scr_date)
	now_iso = datetime.now().isoformat(timespec="seconds")
	scr_by = (str(screened_by or "").strip()) or "Okul Revir Sorumlusu"

	added_students = 0
	added_adults = 0
	updated_profiles = 0
	screened_count = 0
	vitals_count = 0

	with connect_db_fn() as conn:
		conn.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sivil / Diğer')")
		conn.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sınıf Belirtilmedi')")
		cls_rows = conn.execute("SELECT id, name FROM classes").fetchall()
		class_map = {str(r["name"]): int(r["id"]) for r in cls_rows}
		class_fold_map = {tr_fold(r["name"]): int(r["id"]) for r in cls_rows}
		civ_cid = class_map["Sivil / Diğer"]
		unassigned_cid = class_map["Sınıf Belirtilmedi"]

		stu_rows = conn.execute(
			"""SELECT s.id, s.full_name, s.student_no, s.class_id, s.sex, s.birth_date, s.person_type, s.notes,
			          c.name AS class_name
			   FROM students s
			   LEFT JOIN classes c ON c.id = s.class_id"""
		).fetchall()
		db_people: list[dict[str, Any]] = []
		by_fold_adult: dict[str, dict[str, Any]] = {}
		by_fold_student: dict[str, dict[str, Any]] = {}
		by_sno: dict[tuple[str, str], dict[str, Any]] = {}
		for sr in stu_rows:
			d = dict(sr)
			c_nm = str(d.get("class_name") or "").strip()
			p_tp = str(d.get("person_type") or "").strip()
			is_ad = (c_nm == "Sivil / Diğer") or (p_tp != "" and p_tp != "Öğrenci")
			d["is_adult"] = is_ad
			d["fold_name"] = tr_fold(d.get("full_name") or "")
			db_people.append(d)
			if d["fold_name"]:
				if is_ad:
					by_fold_adult[d["fold_name"]] = d
				else:
					by_fold_student[d["fold_name"]] = d
			if (not is_ad) and d.get("student_no") and c_nm:
				by_sno[(tr_fold(c_nm), str(d["student_no"]).strip())] = d

		def _resolve_student_cid(raw_cls: str) -> int:
			c_s = str(raw_cls or default_class or "").strip()
			if not c_s or c_s == "Sivil / Diğer":
				return unassigned_cid
			if c_s in class_map:
				return class_map[c_s]
			cf = tr_fold(c_s)
			if cf in class_fold_map:
				return class_fold_map[cf]
			for existing_nm, existing_id in class_map.items():
				if existing_nm != "Sivil / Diğer" and (existing_nm.endswith(f"/ {c_s}") or cf in tr_fold(existing_nm)):
					return existing_id
			conn.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (c_s,))
			nid = int(conn.execute("SELECT id FROM classes WHERE name=?", (c_s,)).fetchone()["id"])
			class_map[c_s] = nid
			class_fold_map[cf] = nid
			return nid

		for raw_r in records:
			fn = clean_cell_str(raw_r.get("full_name"))
			if not fn:
				continue
			is_adult = (raw_r.get("person_group") == "adult")
			matched, _mk = _find_matching_person(raw_r, db_people, by_fold_adult, by_fold_student, by_sno)

			sex_in = normalize_sex_smart(raw_r.get("sex"))
			dob_in = normalize_date_smart(raw_r.get("birth_date"))
			h_val = normalize_height_smart(raw_r.get("height_cm"))
			w_val = normalize_weight_smart(raw_r.get("weight_kg"), is_adult=is_adult)
			hc_val = None if is_adult else normalize_head_circ_smart(raw_r.get("head_circ_cm"))
			sys_v, dia_v = normalize_bp_smart(raw_r.get("bp_sys") if raw_r.get("bp_sys") is not None else raw_r.get("bp"), raw_r.get("bp_dia"))
			glu_v = normalize_glucose_smart(raw_r.get("glucose_mg_dl"))
			gctx_v = clean_cell_str(raw_r.get("glucose_context")) or ("Açlık" if glu_v is not None else None)
			pls_i = normalize_pulse_smart(raw_r.get("pulse_bpm"))
			pls_v = float(pls_i) if pls_i is not None else None
			tmp_v = normalize_temp_smart(raw_r.get("temperature_c"))
			spo_i = normalize_spo2_smart(raw_r.get("spo2_pct"))
			spo_v = float(spo_i) if spo_i is not None else None

			wears_v = clean_cell_str(raw_r.get("wears_glasses"))
			vr_v = clean_cell_str(raw_r.get("vision_right"))
			vl_v = clean_cell_str(raw_r.get("vision_left")) or (vr_v if is_adult else None)
			if is_adult and vr_v and not wears_v:
				if "Gözlük" in vr_v or "Lens" in vr_v:
					wears_v = "Evet (Sürekli)"
				elif "Normal" in vr_v:
					wears_v = "Hayır (Kullanmıyor)"
			hair_v = clean_cell_str(raw_r.get("hair_screening"))
			hr_v = clean_cell_str(raw_r.get("hearing_right"))
			hl_v = clean_cell_str(raw_r.get("hearing_left"))
			beh_v = clean_cell_str(raw_r.get("health_behavior"))
			psy_v = clean_cell_str(raw_r.get("psychosocial_status"))
			notes_v = clean_cell_str(raw_r.get("notes"))

			if matched:
				pid = int(matched["id"])
				sex_final = sex_in or matched.get("sex")
				dob_final = dob_in or matched.get("birth_date")
				if is_adult:
					role_final = clean_cell_str(raw_r.get("role")) or matched.get("notes") or default_role
					conn.execute(
						"UPDATE students SET sex=COALESCE(?, sex), birth_date=COALESCE(?, birth_date), notes=COALESCE(?, notes) WHERE id=?",
						(sex_final, dob_final, role_final, pid),
					)
				else:
					cid_final = _resolve_student_cid(str(raw_r.get("class_name") or "")) if raw_r.get("class_name") else int(matched["class_id"])
					sno_final = clean_cell_str(raw_r.get("student_no")) or matched.get("student_no")
					conn.execute(
						"UPDATE students SET class_id=?, student_no=COALESCE(?, student_no), sex=COALESCE(?, sex), birth_date=COALESCE(?, birth_date) WHERE id=?",
						(cid_final, sno_final, sex_final, dob_final, pid),
					)
				updated_profiles += 1
			else:
				sex_final = sex_in or (None if is_adult else "E")
				dob_final = dob_in
				if is_adult:
					role_final = clean_cell_str(raw_r.get("role")) or default_role
					cur_ins = conn.execute(
						"""INSERT INTO students(class_id, person_type, full_name, sex, birth_date, notes, country, province, district)
						   VALUES (?, 'Sivil / diğer kişi', ?, ?, ?, ?, ?, ?, ?)""",
						(civ_cid, fn, sex_final, dob_final, role_final, default_country, default_province, default_district),
					)
					pid = int(cur_ins.lastrowid)
					added_adults += 1
					new_entry = {"id": pid, "full_name": fn, "fold_name": tr_fold(fn), "is_adult": True, "sex": sex_final, "birth_date": dob_final, "notes": role_final}
					db_people.append(new_entry)
					by_fold_adult[new_entry["fold_name"]] = new_entry
				else:
					cid_final = _resolve_student_cid(str(raw_r.get("class_name") or ""))
					sno_final = clean_cell_str(raw_r.get("student_no"))
					cur_ins = conn.execute(
						"""INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date, country, province, district)
						   VALUES (?, 'Öğrenci', ?, ?, ?, ?, ?, ?, ?)""",
						(cid_final, sno_final, fn, sex_final, dob_final, default_country, default_province, default_district),
					)
					pid = int(cur_ins.lastrowid)
					added_students += 1
					new_entry = {"id": pid, "full_name": fn, "fold_name": tr_fold(fn), "is_adult": False, "class_id": cid_final, "student_no": sno_final, "sex": sex_final, "birth_date": dob_final}
					db_people.append(new_entry)
					by_fold_student[new_entry["fold_name"]] = new_entry

			has_any_screening = any(
				v is not None
				for v in (h_val, w_val, hc_val, sys_v, dia_v, glu_v, pls_v, tmp_v, spo_v, wears_v, vr_v, vl_v, hair_v, hr_v, hl_v, beh_v, psy_v, notes_v)
			)
			if has_any_screening:
				conn.execute(
					"""INSERT INTO school_screenings(
						student_id, screening_date, height_cm, weight_kg, head_circ_cm,
						temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right,
						vision_left, hair_screening, hearing_right, hearing_left,
						health_behavior, harmful_habits, psychosocial_status, notes, screened_by, created_at
					) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, ?, ?)""",
					(
						pid, scr_date, h_val, w_val, hc_val,
						tmp_v, pls_v, spo_v, wears_v, vr_v,
						vl_v, hair_v, hr_v, hl_v,
						beh_v, psy_v, notes_v or ("Toplu Tarama Aktarımı (Yetişkin)" if is_adult else "Toplu Tarama Aktarımı"),
						scr_by, now_iso,
					),
				)
				conn.execute(
					"""INSERT INTO health_profiles(
						student_id, updated_at, wears_glasses, vision_right, vision_left,
						hair_screening, hearing_right, hearing_left, health_behavior,
						psychosocial_status, last_screened_at, reviewed_by
					) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
					ON CONFLICT(student_id) DO UPDATE SET
						updated_at = excluded.updated_at,
						wears_glasses = COALESCE(excluded.wears_glasses, health_profiles.wears_glasses),
						vision_right = COALESCE(excluded.vision_right, health_profiles.vision_right),
						vision_left = COALESCE(excluded.vision_left, health_profiles.vision_left),
						hair_screening = COALESCE(excluded.hair_screening, health_profiles.hair_screening),
						hearing_right = COALESCE(excluded.hearing_right, health_profiles.hearing_right),
						hearing_left = COALESCE(excluded.hearing_left, health_profiles.hearing_left),
						health_behavior = COALESCE(excluded.health_behavior, health_profiles.health_behavior),
						psychosocial_status = COALESCE(excluded.psychosocial_status, health_profiles.psychosocial_status),
						last_screened_at = excluded.last_screened_at,
						reviewed_by = COALESCE(excluded.reviewed_by, health_profiles.reviewed_by)""",
					(pid, now_iso, wears_v, vr_v, vl_v, hair_v, hr_v, hl_v, beh_v, psy_v, scr_date, scr_by),
				)
				if h_val is not None or w_val is not None or hc_val is not None:
					upsert_measurement_conn_fn(
						conn,
						pid,
						scr_date_obj,
						h_val,
						w_val,
						birth_date=dob_final,
						sex=sex_final,
						head_circ_cm=hc_val,
					)
				if any(v is not None for v in (sys_v, dia_v, pls_v, tmp_v, spo_v, glu_v)):
					upsert_vital_conn_fn(
						conn,
						pid,
						f"{scr_date}T09:30",
						sys_v,
						dia_v,
						pls_v,
						tmp_v,
						glu_v,
						glucose_context=gctx_v if glu_v is not None else None,
						notes="Okul Taraması (Yetişkin)" if is_adult else "Okul Taraması",
						spo2=spo_v,
						upsert_same_slot=True,
						upsert_screening_day=True,
					)
					vitals_count += 1
				screened_count += 1

	if callable(on_after_commit_fn):
		on_after_commit_fn()

	return {
		"added_students": added_students,
		"added_adults": added_adults,
		"updated_profiles": updated_profiles,
		"screened_count": screened_count,
		"vitals_count": vitals_count,
		"total_processed": added_students + added_adults + updated_profiles,
	}


# ==============================================================================
# 4. ŞABLONDAN & DOSYADAN / GÖRSELDEN / METİNDEN YÜKLE SEKME ARAYÜZÜ
# ==============================================================================

def render_smart_template_and_ai_import_tab(
	*,
	connect_db_fn: Callable,
	get_app_settings_fn: Callable[[], dict[str, Any]],
	get_gemini_cfg_fn: Callable[[], tuple[bool, str, str]],
	ocr_reader_fn: Callable[[], Any],
	extract_docx_fn: Callable,
	ai_parse_fn: Callable,
	commit_fn: Callable[[list[dict[str, Any]], str, str, str | None, str], dict[str, int]],
	merge_editor_fn: Callable[[pd.DataFrame, object], pd.DataFrame],
	target_class: str | None = None,
	key_prefix: str = "smart_scr_imp",
) -> None:
	kp = f"{key_prefix}_{target_class or 'all'}"
	app_cfg = get_app_settings_fn()
	is_ai_on, ai_key, ai_model = get_gemini_cfg_fn()

	flash_msg = ui.session_state.pop(f"_flash_{kp}", None)
	if flash_msg:
		ui.success(flash_msg)

	# 1. Şablon İndirme Merkezi (Öğrenci & Yetişkin)
	with ui.container(border=True):
		ui.markdown("#### 📥 Resmi Excel Tarama Şablonları (Öğrenci & Yetişkin)")
		ui.caption("Taramadan önce boş şablonu veya sistemde kayıtlı öğrenci/personel isimleri hazır doldurulmuş şablonu indirip doldurarak sıfır hatayla geri yükleyebilirsiniz.")
		tc1, tc2, tc3, tc4 = ui.columns(4)
		with tc1:
			ui.download_button(
				"🎒 Öğrenci Boş Şablon (.xlsx)",
				data=lambda _tc=target_class: build_student_screening_excel_template(connect_db_fn, include_roster=False, target_class=_tc),
				file_name=f"ogrenci_tarama_sablonu_{target_class or 'genel'}.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key=f"dl_stu_blank_{kp}",
				use_container_width=True,
			)
		with tc2:
			ui.download_button(
				"📋 Kayıtlı Öğrencili Şablon (.xlsx)",
				data=lambda _tc=target_class: build_student_screening_excel_template(connect_db_fn, include_roster=True, target_class=_tc),
				file_name=f"ogrenci_hazir_liste_tarama_{target_class or 'tum_okul'}.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key=f"dl_stu_roster_{kp}",
				use_container_width=True,
			)
		with tc3:
			ui.download_button(
				"🧑 Yetişkin Boş Şablon (.xlsx)",
				data=lambda: build_adult_screening_excel_template(connect_db_fn, include_roster=False),
				file_name="yetiskin_personel_tarama_sablonu.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key=f"dl_adult_blank_{kp}",
				use_container_width=True,
			)
		with tc4:
			ui.download_button(
				"👨‍🏫 Kayıtlı Personelli Şablon (.xlsx)",
				data=lambda: build_adult_screening_excel_template(connect_db_fn, include_roster=True),
				file_name="yetiskin_hazir_liste_tarama.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key=f"dl_adult_roster_{kp}",
				use_container_width=True,
			)

	# 2. Yükleme, Kopyala-Yapıştır & AI Ayarları
	with ui.container(border=True):
		cfg1, cfg2, cfg3, cfg4 = ui.columns([1.35, 1.0, 1.15, 1.2])
		with cfg1:
			grp_labels = [
				"🤖 Otomatik Algıla (Öğrenci / Yetişkin)",
				"🎒 Öğrenci Sağlık Taraması",
				"🧑 Yetişkin / Personel Taraması",
			]
			grp_choice = ui.selectbox("Hedef Tarama Grubu", grp_labels, index=0 if not target_class else 1, key=f"grp_{kp}")
			grp_code = "student" if "Öğrenci Sağlık" in grp_choice else ("adult" if "Yetişkin" in grp_choice else "auto")
		with cfg2:
			scr_dt = ui.date_input("🗓️ Tarama Tarihi", value=date.today(), key=f"dt_{kp}")
		with cfg3:
			scr_nurse = ui.text_input("👩‍⚕️ Taramayı Yapan", value=app_cfg.get("nurse_name", "Okul Revir Sorumlusu"), key=f"nrs_{kp}")
		with cfg4:
			def_role_or_cls = ui.text_input(
				"Varsayılan Görev / Sınıf",
				value=target_class or ("Öğretmen / Personel" if grp_code == "adult" else ""),
				placeholder="Örn: Öğretmen veya 4-A",
				key=f"def_rc_{kp}",
			)

		in_l, in_r = ui.columns([1.35, 1.35])
		with in_l:
			up_file = ui.file_uploader(
				"📄 Dosya veya Fotoğraf Yükle (.xlsx, .xls, .docx, .csv, .pdf, .jpg, .png, .webp)",
				type=["xlsx", "xls", "docx", "csv", "tsv", "txt", "pdf", "png", "jpg", "jpeg", "webp", "bmp"],
				key=f"uploader_{kp}",
			)
		with in_r:
			pasted_txt = ui.text_area(
				"📋 Veya Tablo / Metin Yapıştır (Excel, Word, WhatsApp)",
				placeholder="Ad Soyad\tBoy\tKilo\tTansiyon\tŞeker\tNabız\nMüberra Görgü\t168\t77,6\t115-60\t96\t89",
				height=92,
				key=f"paste_{kp}",
			)

	has_file = up_file is not None
	has_paste = bool(str(pasted_txt or "").strip())
	if not has_file and not has_paste:
		return

	f_bytes = up_file.getvalue() if has_file else None
	f_name = up_file.name if has_file else "yapistirilan_tablo.tsv"
	src_sig = f"{grp_code}:{f_name}:{len(f_bytes) if f_bytes else 0}:{hash(pasted_txt or '')}"
	ai_cache_key = f"ai_parsed_{kp}"
	is_image_or_pdf = bool(has_file and f_name.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".pdf")))

	parsed_rows: list[dict[str, Any]] = []
	detected_grp = grp_code if grp_code != "auto" else "student"
	source_note = ""

	cached_ai = ui.session_state.get(ai_cache_key)
	use_cached_ai = isinstance(cached_ai, dict) and cached_ai.get("sig") == src_sig and bool(cached_ai.get("rows"))

	if use_cached_ai:
		parsed_rows = cached_ai["rows"]
		detected_grp = cached_ai.get("group") or detected_grp
		source_note = f"🤖 Yapay Zekâ Çözümlemesi: {cached_ai.get('note', '')}"
	else:
		# Görsellerde Gemini açıksa doğrudan yüksek doğruluklu Multimodal Vision kullan, kapalıysa yerel OCR çalıştır
		if is_image_or_pdf and is_ai_on and f_bytes and f_name.lower().endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp")):
			with ui.spinner("🤖 Gemini Vision AI görseldeki tarama tablosunu okuyor..."):
				ok_v, ai_rows_v, ai_grp_v, ai_note_v = ai_parse_fn(
					f_bytes, f_name, "", grp_code, target_class or "", def_role_or_cls or "Öğretmen / Personel"
				)
			if ok_v and ai_rows_v:
				ui.session_state[ai_cache_key] = {"sig": src_sig, "rows": ai_rows_v, "group": ai_grp_v, "note": ai_note_v}
				parsed_rows = ai_rows_v
				detected_grp = ai_grp_v
				source_note = f"🤖 Gemini Vision AI: {ai_note_v}"
		if not parsed_rows:
			fb_cls = target_class or (def_role_or_cls if grp_code == "student" else "")
			fb_role = def_role_or_cls if grp_code != "student" and def_role_or_cls else "Öğretmen / Personel"
			with ui.spinner("Veriler ayrıştırılıyor ve sütunlar normalize ediliyor..."):
				parsed_rows, detected_grp, source_note = parse_smart_screening_source(
					file_bytes=f_bytes,
					file_name=f_name,
					raw_text=pasted_txt if not has_file else "",
					target_group=grp_code,
					fallback_class=fb_cls,
					default_role=fb_role,
					ocr_reader_fn=ocr_reader_fn,
					extract_docx_fn=extract_docx_fn,
				)

	if is_ai_on:
		ai_c1, ai_c2 = ui.columns([2.3, 1.0])
		with ai_c1:
			btn_ai_lbl = (
				"🤖 Gemini AI Vision / Akıllı Tablo Çözümleyici ile Yeniden Oku & Doğrula"
				if parsed_rows
				else "🤖 Karmaşık Görseli / Belgeyi Gemini AI ile Çözümle"
			)
			if ui.button(btn_ai_lbl, type="secondary" if parsed_rows else "primary", use_container_width=True, key=f"btn_ai_{kp}"):
				with ui.spinner("Gemini AI belgeyi / görseli inceliyor ve tüm tarama alanlarını çıkarıyor..."):
					ok_a, rows_a, grp_a, note_a = ai_parse_fn(
						f_bytes,
						f_name,
						pasted_txt if not has_file else "",
						grp_code,
						target_class or "",
						def_role_or_cls or "Öğretmen / Personel",
					)
				if ok_a and rows_a:
					ui.session_state[ai_cache_key] = {"sig": src_sig, "rows": rows_a, "group": grp_a, "note": note_a}
					ui.rerun()
				else:
					ui.error(f"AI ayrıştırma hatası: {note_a}")
		with ai_c2:
			if use_cached_ai and ui.button("↩️ Standart Okumaya Dön", use_container_width=True, key=f"btn_ai_rst_{kp}"):
				ui.session_state.pop(ai_cache_key, None)
				ui.rerun()

	if not parsed_rows:
		ui.warning("⚠️ Kaynaktan geçerli bir tarama satırı okunamadı. Sütun başlıklarını kontrol edebilir veya yukarıdaki Gemini AI butonunu kullanabilirsiniz.")
		return

	enriched = preview_match_and_alerts(connect_db_fn, parsed_rows)
	matched_n = sum(1 for r in enriched if r["matched_id"] is not None)
	new_n = len(enriched) - matched_n
	alert_n = sum(1 for r in enriched if r["clinical_alert"])
	is_adult_mode = sum(1 for r in enriched if r.get("person_group") == "adult") >= (len(enriched) / 2.0)

	ui.caption(
		f"✓ **{html.escape(source_note)}** &nbsp;·&nbsp; Algılanan Mod: **{'🧑 Yetişkin / Personel' if is_adult_mode else '🎒 Öğrenci'}** "
		f"&nbsp;·&nbsp; Toplam: **{len(enriched)}** &nbsp;·&nbsp; Eşleşen: **{matched_n}** &nbsp;·&nbsp; Yeni Profil: **{new_n}**"
		+ (f" &nbsp;·&nbsp; ⚠️ Klinik Dikkat: **{alert_n}**" if alert_n else ""),
		unsafe_allow_html=True,
	)

	# 3. Düzenlenebilir Önizleme Tablosu (Kayıt Öncesi Hücre Düzeltme İmkanı)
	editor_rows: list[dict[str, Any]] = []
	for r in enriched:
		st_lbl = "🟢 Eşleşti" if r["matched_id"] else "➕ Yeni Profil"
		if is_adult_mode:
			bp_s = f"{r['bp_sys']}/{r['bp_dia']}" if (r.get("bp_sys") and r.get("bp_dia")) else ""
			editor_rows.append({
				"Durum": st_lbl,
				"Ad Soyad": r.get("full_name") or "",
				"Görev / Branş": r.get("role") or (def_role_or_cls or "Öğretmen / Personel"),
				"Boy (cm)": r.get("height_cm"),
				"Kilo (kg)": r.get("weight_kg"),
				"VKİ": r.get("bmi"),
				"Tansiyon": bp_s,
				"Şeker (mg/dL)": r.get("glucose_mg_dl"),
				"Nabız (/dk)": r.get("pulse_bpm"),
				"Görme": r.get("vision_right") or "",
				"Not": r.get("notes") or "",
				"Klinik Uyarı": r.get("clinical_alert") or "—",
			})
		else:
			editor_rows.append({
				"Durum": st_lbl,
				"Sınıf": r.get("class_name") or (target_class or ""),
				"Okul No": r.get("student_no") or "",
				"Ad Soyad": r.get("full_name") or "",
				"Cinsiyet": r.get("sex") or "",
				"Doğum Tarihi": r.get("birth_date") or "",
				"Boy (cm)": r.get("height_cm"),
				"Kilo (kg)": r.get("weight_kg"),
				"VKİ": r.get("bmi"),
				"Nabız (/dk)": r.get("pulse_bpm"),
				"Ateş (°C)": r.get("temperature_c"),
				"Gözlük": r.get("wears_glasses") or "",
				"Sağ Göz": r.get("vision_right") or "",
				"Sol Göz": r.get("vision_left") or "",
				"Saç / Hijyen": r.get("hair_screening") or "",
				"Not": r.get("notes") or "",
			})

	preview_df = pd.DataFrame(editor_rows)
	ed_key = f"ed_prev_{kp}_{abs(hash(src_sig)) % 100000}"
	edited_df = ui.data_editor(
		preview_df,
		hide_index=True,
		use_container_width=True,
		key=ed_key,
		column_config={
			"Durum": ui.column_config.TextColumn("Durum", disabled=True),
			"VKİ": ui.column_config.NumberColumn("VKİ", disabled=True),
			"Klinik Uyarı": ui.column_config.TextColumn("Klinik Uyarı", disabled=True),
		},
	)

	btn_commit_lbl = (
		f"🚀 {len(preview_df)} Yetişkinin Profil ve Tarama Verilerini Kaydet"
		if is_adult_mode
		else f"🚀 {len(preview_df)} Öğrencinin Profil ve Tarama Verilerini Kaydet"
	)
	if ui.button(btn_commit_lbl, type="primary", use_container_width=True, key=f"btn_commit_{kp}"):
		merged_df = merge_editor_fn(preview_df, ui.session_state.get(ed_key))
		if edited_df is not None and not edited_df.empty and merged_df.equals(preview_df):
			merged_df = edited_df
		final_recs, _ = extract_records_from_dataframe(
			merged_df.drop(columns=[c for c in ("Durum", "VKİ", "Klinik Uyarı") if c in merged_df.columns]),
			target_group="adult" if is_adult_mode else "student",
			fallback_class=target_class or "",
			default_role=def_role_or_cls or "Öğretmen / Personel",
		)
		res = commit_fn(
			final_recs,
			scr_dt.isoformat(),
			scr_nurse,
			target_class,
			def_role_or_cls or "Öğretmen / Personel",
		)
		ui.session_state.pop(f"paste_{kp}", None)
		ui.session_state.pop(ai_cache_key, None)
		ui.session_state[f"_flash_{kp}"] = (
			f"✅ Aktarım tamamlandı: {res['screened_count']} kişinin tarama/ölçüm verisi kaydedildi "
			f"(Yeni yetişkin: {res['added_adults']}, Yeni öğrenci: {res['added_students']}, Güncellenen profil: {res['updated_profiles']})."
		)
		ui.rerun()
