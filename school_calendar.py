# -*- coding: utf-8 -*-
"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Akıllı Okul & Revir Takvimi,
İnteraktif Takvim Yaprağı, 81 İl Ezan Vakitleri Canlı Sayacı, MEB Belirli Gün ve Haftalar
ve Gemini AI Planlama Modülü.
"""
from __future__ import annotations

import calendar
import html
import json
import math
import re
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from typing import Any, Callable

import revir_ui as ui
from gemini_assistant import (
	generate_calendar_events_with_ai,
	generate_daily_leaf_content_with_ai,
)

# ---------------------------------------------------------------------------
# SABİTLER, RENK PLANI VE SVG SEMBOLLER (TÜRK BAYRAĞI & HİLAL)
# ---------------------------------------------------------------------------

TR_MONTH_NAMES = {
	1: "Ocak",
	2: "Şubat",
	3: "Mart",
	4: "Nisan",
	5: "Mayıs",
	6: "Haziran",
	7: "Temmuz",
	8: "Ağustos",
	9: "Eylül",
	10: "Ekim",
	11: "Kasım",
	12: "Aralık",
}

TR_WEEKDAY_NAMES = {
	0: "Pazartesi",
	1: "Salı",
	2: "Çarşamba",
	3: "Perşembe",
	4: "Cuma",
	5: "Cumartesi",
	6: "Pazar",
}

TR_WEEKDAY_SHORT = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]

HIJRI_MONTH_NAMES = {
	1: "Muharrem",
	2: "Safer",
	3: "Rebiülevvel",
	4: "Rebiülahir",
	5: "Cemaziyelevvel",
	6: "Cemaziyelahir",
	7: "Recep",
	8: "Şaban",
	9: "Ramazan",
	10: "Şevval",
	11: "Zilkade",
	12: "Zilhicce",
}

QUOTE_MODE_OPTIONS = [
	"Karışık (Hadis-i Şerif + Özlü Söz / Günün Yazısı)",
	"Sadece Hadis-i Şerif ve Hikmetli Sözler",
	"Sadece Özlü Söz, Atasözü ve Eğitim/Sağlık Vecizeleri",
	"Kapalı (Gösterme)",
]

CATEGORY_LABELS: dict[str, tuple[str, str, str]] = {
	"milli_bayram": ("Milli Bayram & Ulusal Gün", "#dc2626", "tr_flag"),
	"dini_bayram": ("Dini Bayram & Mübarek Gün", "#d97706", "crescent"),
	"resmi_tatil": ("Resmi Tatil (Memur İzinli)", "#dc2626", "tr_flag"),
	"yarim_gun": ("Yarım Gün Mesai (Arife / 28 Ekim)", "#ea580c", "half_day"),
	"secim_idari_tatil": ("Seçim Tatili / İdari İzin / Ara Tatil", "#7c3aed", "ballot"),
	"osym_sinav": ("ÖSYM / MEB Sınavı", "#2563eb", "exam"),
	"meslek_gunu": ("Meslek Grubu Günü", "#0d9488", "profession"),
	"belirli_gun": ("MEB Belirli Gün ve Hafta", "#0284c7", "event"),
	"revir_etkinlik": ("Revir Taraması / Aşı / Etkinlik", "#059669", "health"),
}

SYMBOL_OPTIONS: dict[str, str] = {
	"tr_flag": "🇹🇷 Türk Bayrağı (Milli Bayram / Ulusal Gün)",
	"crescent": "🌙 Hilal Sembolü (Dini Bayram / Kandil)",
	"half_day": "½ Yarım Gün Mesai Rozeti",
	"exam": "📝 ÖSYM / MEB Sınavı",
	"profession": "👩‍⚕️ Meslek Grubu Günü",
	"ballot": "🗳️ Seçim Tatili / İdari İzin",
	"health": "🩺 Revir Taraması / Aşı / Sağlık",
	"event": "📌 MEB Belirli Gün ve Hafta / Etkinlik",
}

# Özel Çizim SVG Türk Bayrağı ve Altın Hilal Rozetleri
SVG_TR_FLAG_MINI = (
	'<svg width="18" height="13" viewBox="0 0 30 20" style="vertical-align:-2px;border-radius:2.5px;'
	'box-shadow:0 1px 2px rgba(0,0,0,0.22);display:inline-block;flex-shrink:0;" title="Milli Bayram / Ulusal Gün">'
	'<rect width="30" height="20" fill="#E30A17"/>'
	'<circle cx="11.5" cy="10" r="5" fill="#ffffff"/>'
	'<circle cx="12.8" cy="10" r="4" fill="#E30A17"/>'
	'<polygon points="16.2,10 19.6,11.1 17.5,8.2 17.5,11.8 19.6,8.9" fill="#ffffff"/>'
	'</svg>'
)

SVG_CRESCENT_MINI = (
	'<svg width="15" height="15" viewBox="0 0 24 24" style="vertical-align:-2px;display:inline-block;flex-shrink:0;" '
	'title="Dini Bayram / Kandil / Mübarek Gün">'
	'<circle cx="12" cy="12" r="11" fill="#fef3c7" stroke="#f59e0b" stroke-width="1.5"/>'
	'<path d="M14.8 5.2A7 7 0 1 0 18.8 16.5A5.6 5.6 0 1 1 14.8 5.2Z" fill="#d97706"/>'
	'<circle cx="16.8" cy="8.5" r="1.3" fill="#f59e0b"/>'
	'</svg>'
)


def render_symbol_badge(symbol_type: str, compact: bool = False) -> str:
	"""Sembol türüne göre takvim hücresinde veya yaprağında gösterilecek HTML ikonunu döndürür."""
	st = (symbol_type or "event").strip()
	if st == "tr_flag":
		return SVG_TR_FLAG_MINI
	if st == "crescent":
		return SVG_CRESCENT_MINI
	if st == "half_day":
		return (
			'<span style="background:#ffedd5;color:#c2410c;border:1px solid #fdba74;'
			'border-radius:4px;padding:0 4px;font-size:10px;font-weight:900;line-height:1.3;">½</span>'
		)
	if st == "exam":
		return '<span style="font-size:12.5px;" title="ÖSYM / MEB Sınavı">📝</span>'
	if st == "profession":
		return '<span style="font-size:12.5px;" title="Meslek Grubu Günü">👩‍⚕️</span>'
	if st == "ballot":
		return '<span style="font-size:12.5px;" title="Seçim Tatili / İdari İzin">🗳️</span>'
	if st == "health":
		return '<span style="font-size:12.5px;" title="Revir / Sağlık Etkinliği">🩺</span>'
	return '<span style="font-size:12px;" title="MEB Belirli Gün ve Hafta / Etkinlik">📌</span>'


# ---------------------------------------------------------------------------
# TÜRKİYE 81 İL KOORDİNATLARI (ÇEVRİMDIŞI & ÇEVRİMİÇİ EZAN VAKTİ MOTORU İÇİN)
# ---------------------------------------------------------------------------

TR_PROVINCE_COORDS: dict[str, tuple[float, float]] = {
	"Adana": (37.0000, 35.3213), "Adıyaman": (37.7648, 38.2786), "Afyonkarahisar": (38.7507, 30.5567),
	"Ağrı": (39.7191, 43.0503), "Amasya": (40.6499, 35.8353), "Ankara": (39.9334, 32.8597),
	"Antalya": (36.8969, 30.7133), "Artvin": (41.1828, 41.8183), "Aydın": (37.8560, 27.8416),
	"Balıkesir": (39.6484, 27.8826), "Bilecik": (40.1451, 29.9799), "Bingöl": (38.8855, 40.4966),
	"Bitlis": (38.4006, 42.1095), "Bolu": (40.7392, 31.6089), "Burdur": (37.7203, 30.2908),
	"Bursa": (40.1885, 29.0610), "Çanakkale": (40.1553, 26.4142), "Çankırı": (40.6013, 33.6134),
	"Çorum": (40.5506, 34.9556), "Denizli": (37.7765, 29.0864), "Diyarbakır": (37.9144, 40.2306),
	"Edirne": (41.6771, 26.5557), "Elazığ": (38.6810, 39.2264), "Erzincan": (39.7500, 39.5000),
	"Erzurum": (39.9043, 41.2679), "Eskişehir": (39.7767, 30.5206), "Gaziantep": (37.0662, 37.3833),
	"Giresun": (40.9128, 38.3895), "Gümüşhane": (40.4603, 39.4814), "Hakkari": (37.5833, 43.7333),
	"Hatay": (36.2021, 36.1606), "Isparta": (37.7648, 30.5566), "Mersin": (36.8121, 34.6415),
	"İstanbul": (41.0082, 28.9784), "İzmir": (38.4237, 27.1428), "Kars": (40.6013, 43.0975),
	"Kastamonu": (41.3887, 33.7827), "Kayseri": (38.7312, 35.4787), "Kırklareli": (41.7333, 27.2167),
	"Kırşehir": (39.1425, 34.1709), "Kocaeli": (40.8533, 29.8815), "Konya": (37.8746, 32.4932),
	"Kütahya": (39.4167, 29.9833), "Malatya": (38.3552, 38.3095), "Manisa": (38.6191, 27.4289),
	"Kahramanmaraş": (37.5858, 36.9371), "Mardin": (37.3212, 40.7245), "Muğla": (37.2153, 28.3636),
	"Muş": (38.9462, 41.7539), "Nevşehir": (38.6244, 34.7239), "Niğde": (37.9667, 34.6833),
	"Ordu": (40.9839, 37.8764), "Rize": (41.0201, 40.5234), "Sakarya": (40.7569, 30.3783),
	"Samsun": (41.2867, 36.3300), "Siirt": (37.9333, 41.9500), "Sinop": (42.0231, 35.1531),
	"Sivas": (39.7477, 37.0179), "Tekirdağ": (40.9833, 27.5167), "Tokat": (40.3167, 36.5500),
	"Trabzon": (41.0027, 39.7168), "Tunceli": (39.1079, 39.5401), "Şanlıurfa": (37.1591, 38.7969),
	"Uşak": (38.6823, 29.4082), "Van": (38.4891, 43.4089), "Yozgat": (39.8181, 34.8147),
	"Zonguldak": (41.4564, 31.7987), "Aksaray": (38.3687, 34.0370), "Bayburt": (40.2552, 40.2249),
	"Karaman": (37.1759, 33.2287), "Kırıkkale": (39.8468, 33.5153), "Batman": (37.8812, 41.1351),
	"Şırnak": (37.5164, 42.4611), "Bartın": (41.6344, 32.3375), "Ardahan": (41.1105, 42.7022),
	"Iğdır": (39.9167, 44.0333), "Yalova": (40.6500, 29.2667), "Karabük": (41.2061, 32.6204),
	"Kilis": (36.7184, 37.1212), "Osmaniye": (37.0742, 36.2478), "Düzce": (40.8438, 31.1565),
}


# ---------------------------------------------------------------------------
# DİYANET TAKVİMİ İLE KALİBRE EDİLMİŞ HİCRİ TAKVİM HESAPLAYICI (HATA #8 DÜZELTMESİ)
# ---------------------------------------------------------------------------

# Diyanet İşleri Başkanlığı Resmi Hicri Ay Başlangıç Çıpaları (Miladi Tarih -> (Hicri Yıl, Hicri Ay))
# Bu çıpalar sayesinde 1 Ramazan, 1 Şevval (Ramazan Bayramı), 10 Zilhicce (Kurban Bayramı) ve
# 1 Muharrem (Hicri Yılbaşı) Diyanet Takvimi ile %100 birebir örtüşür.
_DIYANET_HIJRI_MONTH_ANCHORS: list[tuple[date, int, int]] = [
	(date(2024, 7, 7), 1446, 1),    # 1 Muharrem 1446
	(date(2025, 1, 1), 1446, 7),    # 1 Recep 1446
	(date(2025, 1, 31), 1446, 8),   # 1 Şaban 1446
	(date(2025, 3, 1), 1446, 9),    # 1 Ramazan 1446
	(date(2025, 3, 30), 1446, 10),  # 1 Şevval 1446 (Ramazan Bayramı 1. Gün)
	(date(2025, 4, 29), 1446, 11),  # 1 Zilkade 1446
	(date(2025, 5, 28), 1446, 12),  # 1 Zilhicce 1446 (6 Haziran = 10 Zilhicce Kurban Bayramı)
	(date(2025, 6, 26), 1447, 1),   # 1 Muharrem 1447 (Hicri Yılbaşı)
	(date(2025, 7, 26), 1447, 2),   # 1 Safer 1447
	(date(2025, 8, 24), 1447, 3),   # 1 Rebiülevvel 1447 (3 Eylül = 11/12 Rebiülevvel Mevlid)
	(date(2025, 9, 23), 1447, 4),   # 1 Rebiülahir 1447
	(date(2025, 10, 23), 1447, 5),  # 1 Cemaziyelevvel 1447
	(date(2025, 11, 21), 1447, 6),  # 1 Cemaziyelahir 1447
	(date(2025, 12, 21), 1447, 7),  # 1 Recep 1447
	(date(2026, 1, 20), 1447, 8),   # 1 Şaban 1447
	(date(2026, 2, 19), 1447, 9),   # 1 Ramazan 1447
	(date(2026, 3, 20), 1447, 10),  # 1 Şevval 1447 (Ramazan Bayramı 1. Gün)
	(date(2026, 4, 18), 1447, 11),  # 1 Zilkade 1447
	(date(2026, 5, 18), 1447, 12),  # 1 Zilhicce 1447 (27 Mayıs = 10 Zilhicce Kurban Bayramı)
	(date(2026, 6, 16), 1448, 1),   # 1 Muharrem 1448 (Hicri Yılbaşı)
	(date(2026, 7, 16), 1448, 2),   # 1 Safer 1448
	(date(2026, 8, 14), 1448, 3),   # 1 Rebiülevvel 1448
	(date(2026, 9, 13), 1448, 4),   # 1 Rebiülahir 1448
	(date(2026, 10, 12), 1448, 5),  # 1 Cemaziyelevvel 1448
	(date(2026, 11, 11), 1448, 6),  # 1 Cemaziyelahir 1448
	(date(2026, 12, 10), 1448, 7),  # 1 Recep 1448
	(date(2027, 1, 9), 1448, 8),    # 1 Şaban 1448
	(date(2027, 2, 8), 1448, 9),    # 1 Ramazan 1448
	(date(2027, 3, 9), 1448, 10),   # 1 Şevval 1448 (Ramazan Bayramı 1. Gün)
	(date(2027, 4, 8), 1448, 11),   # 1 Zilkade 1448
	(date(2027, 5, 7), 1448, 12),   # 1 Zilhicce 1448 (16 Mayıs = 10 Zilhicce Kurban Bayramı)
	(date(2027, 6, 6), 1449, 1),    # 1 Muharrem 1449 (Hicri Yılbaşı)
]


def gregorian_to_hijri_str(d: date) -> str:
	"""
	Miladi tarihi Diyanet İşleri Başkanlığı takvimiyle kalibre edilmiş Türkçe Hicri takvim
	metnine (Örn: '1 Ramazan 1447', '1 Muharrem 1448') dönüştürür.
	"""
	# Önce Diyanet çıpa tablosunda aralık eşleşmesi kontrol et
	if _DIYANET_HIJRI_MONTH_ANCHORS[0][0] <= d <= (_DIYANET_HIJRI_MONTH_ANCHORS[-1][0] + timedelta(days=29)):
		chosen_anchor = _DIYANET_HIJRI_MONTH_ANCHORS[0]
		for anc in _DIYANET_HIJRI_MONTH_ANCHORS:
			if anc[0] <= d:
				chosen_anchor = anc
			else:
				break
		anc_date, h_year, h_month = chosen_anchor
		h_day = (d - anc_date).days + 1
		if 1 <= h_day <= 30:
			return f"{h_day} {HIJRI_MONTH_NAMES.get(h_month, 'Muharrem')} {h_year}"

	# Çıpa aralığı dışındaki yıllar için Diyanet ofsetli (-1 gün) astronomik Jülyen dönüşümü
	jd = int((1461 * (d.year + 4800 + int((d.month - 14) / 12))) / 4)
	jd += int((367 * (d.month - 2 - 12 * int((d.month - 14) / 12))) / 12)
	jd -= int((3 * int((d.year + 4900 + int((d.month - 14) / 12)) / 100)) / 4)
	jd += d.day - 32076  # Diyanet rü'yet-i hilal kalibrasyonu

	l_val = jd - 1948440 + 10632
	n_val = int((l_val - 1) / 10631)
	l_val = l_val - 10631 * n_val + 354
	j_val = (int((10985 - l_val) / 5316)) * (int((50 * l_val) / 17719)) + (int(l_val / 5670)) * (int((43 * l_val) / 15238))
	l_val = l_val - (int((30 - j_val) / 15)) * (int((17719 * j_val) / 50)) - (int(j_val / 16)) * (int((15238 * j_val) / 43)) + 29
	h_month = int((24 * l_val) / 709)
	h_day = max(1, min(30, l_val - int((709 * h_month) / 24)))
	h_year = 30 * n_val + j_val - 30
	m_name = HIJRI_MONTH_NAMES.get(h_month, "Muharrem")
	return f"{h_day} {m_name} {h_year}"


# ---------------------------------------------------------------------------
# DİYANET UYUMLU 81 İL EZAN VAKİTLERİ MOTORU (ONLINE + ASTRONOMİK OFFLINE)
# ---------------------------------------------------------------------------

def _compute_offline_prayer_times(province: str, d: date) -> dict[str, str]:
	"""
	İnternet bağlantısı olmadığında Türkiye'nin 81 ili için Diyanet açılarına
	(İmsak -18°, Yatsı -17°, temkin payları) göre astronomik güneş konumundan ezan vakitlerini hesaplar.
	"""
	lat, lon = TR_PROVINCE_COORDS.get(province, TR_PROVINCE_COORDS["Ankara"])
	tz = 3.0  # Türkiye sabit GMT+3
	doy = d.timetuple().tm_yday

	gamma = 2.0 * math.pi / 365.0 * (doy - 1 + 0.5)
	eq_time = 229.18 * (
		0.000075
		+ 0.001868 * math.cos(gamma)
		- 0.032077 * math.sin(gamma)
		- 0.014615 * math.cos(2 * gamma)
		- 0.040849 * math.sin(2 * gamma)
	)
	decl = (
		0.006918
		- 0.399912 * math.cos(gamma)
		+ 0.070257 * math.sin(gamma)
		- 0.006758 * math.cos(2 * gamma)
		+ 0.000907 * math.sin(2 * gamma)
		- 0.002697 * math.cos(3 * gamma)
		+ 0.00148 * math.sin(3 * gamma)
	)
	lat_rad = math.radians(lat)

	def _hour_angle(alt_deg: float) -> float:
		alt_rad = math.radians(alt_deg)
		cos_ha = (math.sin(alt_rad) - math.sin(lat_rad) * math.sin(decl)) / (math.cos(lat_rad) * math.cos(decl))
		cos_ha = max(-0.999, min(0.999, cos_ha))
		return math.degrees(math.acos(cos_ha)) / 15.0

	noon = 12.0 + tz - (lon / 15.0) - (eq_time / 60.0)
	ha_sun = _hour_angle(-0.833)
	ha_fajr = _hour_angle(-18.0)
	ha_isha = _hour_angle(-17.0)

	# Asr-ı Evvel (Diyanet standart gölge = 1 + tan(|lat - decl|))
	asr_alt = math.degrees(math.atan(1.0 / (1.0 + math.tan(abs(lat_rad - decl)))))
	ha_asr = _hour_angle(asr_alt)

	# Diyanet temkin süreleri (dakika)
	imsak_h = noon - ha_fajr - (2.0 / 60.0)
	gunes_h = noon - ha_sun - (6.0 / 60.0)
	ogle_h = noon + (5.0 / 60.0)
	ikindi_h = noon + ha_asr + (4.0 / 60.0)
	aksam_h = noon + ha_sun + (7.0 / 60.0)
	yatsi_h = noon + ha_isha + (2.0 / 60.0)

	def _fmt(h_val: float) -> str:
		total_m = int(round((h_val % 24.0) * 60.0))
		hh = (total_m // 60) % 24
		mm = total_m % 60
		return f"{hh:02d}:{mm:02d}"

	return {
		"İmsak": _fmt(imsak_h),
		"Güneş": _fmt(gunes_h),
		"Öğle": _fmt(ogle_h),
		"İkindi": _fmt(ikindi_h),
		"Akşam": _fmt(aksam_h),
		"Yatsı": _fmt(yatsi_h),
		"source": "Yerleşik Diyanet Astronomik Hesaplama",
	}


_PRAYER_MEMORY_CACHE: dict[str, dict[str, str]] = {}


def get_prayer_times_for_province(province: str, d: date, connect_db_fn: Callable | None = None) -> dict[str, str]:
	"""
	Seçili il ve tarih için Diyanet İşleri Başkanlığı (Method 13) ezan vakitlerini döndürür.
	Önce bellek ve SQLite önbelleğine bakar, yoksa çevrimiçi Diyanet/Aladhan servisinden çeker;
	çevrimdışıysa yerleşik 81 il astronomik motoruyla anında hesaplar.
	"""
	prov_clean = (province or "Ankara").strip() or "Ankara"
	if prov_clean not in TR_PROVINCE_COORDS:
		for k in TR_PROVINCE_COORDS:
			if k.lower() == prov_clean.lower():
				prov_clean = k
				break
		else:
			prov_clean = "Ankara"

	date_iso = d.isoformat()
	cache_key = f"{prov_clean}|{date_iso}"
	if cache_key in _PRAYER_MEMORY_CACHE:
		return _PRAYER_MEMORY_CACHE[cache_key]

	if connect_db_fn is not None:
		try:
			with connect_db_fn() as conn:
				row = conn.execute(
					"SELECT prayer_cache_json FROM calendar_daily_notes WHERE note_date = ?",
					(date_iso,),
				).fetchone()
				if row and row["prayer_cache_json"]:
					cached_map = json.loads(row["prayer_cache_json"])
					if isinstance(cached_map, dict) and prov_clean in cached_map:
						_PRAYER_MEMORY_CACHE[cache_key] = cached_map[prov_clean]
						return cached_map[prov_clean]
		except Exception:
			pass

	lat, lon = TR_PROVINCE_COORDS.get(prov_clean, TR_PROVINCE_COORDS["Ankara"])
	result: dict[str, str] | None = None
	try:
		d_str = d.strftime("%d-%m-%Y")
		url = f"https://api.aladhan.com/v1/timings/{d_str}?latitude={lat}&longitude={lon}&method=13"
		req = urllib.request.Request(url, headers={"User-Agent": "RevirDefteri/1.15"})
		with urllib.request.urlopen(req, timeout=2.2) as resp:
			payload = json.loads(resp.read().decode("utf-8"))
			timings = (payload.get("data") or {}).get("timings") or {}
			if timings.get("Fajr") and timings.get("Maghrib"):
				result = {
					"İmsak": str(timings.get("Imsak") or timings.get("Fajr"))[:5],
					"Güneş": str(timings.get("Sunrise"))[:5],
					"Öğle": str(timings.get("Dhuhr"))[:5],
					"İkindi": str(timings.get("Asr"))[:5],
					"Akşam": str(timings.get("Maghrib"))[:5],
					"Yatsı": str(timings.get("Isha"))[:5],
					"source": "Diyanet Takvimi (Çevrimiçi)",
				}
	except Exception:
		result = None

	if not result:
		result = _compute_offline_prayer_times(prov_clean, d)

	_PRAYER_MEMORY_CACHE[cache_key] = result

	if connect_db_fn is not None:
		try:
			with connect_db_fn() as conn:
				row = conn.execute(
					"SELECT prayer_cache_json FROM calendar_daily_notes WHERE note_date = ?",
					(date_iso,),
				).fetchone()
				existing_cache = {}
				if row and row["prayer_cache_json"]:
					try:
						existing_cache = json.loads(row["prayer_cache_json"]) or {}
					except Exception:
						existing_cache = {}
				existing_cache[prov_clean] = result
				cache_str = json.dumps(existing_cache, ensure_ascii=False)
				if row:
					conn.execute(
						"UPDATE calendar_daily_notes SET prayer_cache_json = ?, updated_at = ? WHERE note_date = ?",
						(cache_str, datetime.now().isoformat(timespec="seconds"), date_iso),
					)
				else:
					conn.execute(
						"""INSERT INTO calendar_daily_notes
						   (note_date, prayer_cache_json, updated_at)
						   VALUES (?, ?, ?)""",
						(date_iso, cache_str, datetime.now().isoformat(timespec="seconds")),
					)
		except Exception:
			pass

	return result


# ---------------------------------------------------------------------------
# YERLEŞİK TÜRKİYE RESMİ & DİNİ TATİLLER, YARIM GÜNLER, MEB BELİRLİ GÜN VE HAFTALAR
# ---------------------------------------------------------------------------

FIXED_NATIONAL_DAYS: list[dict[str, Any]] = [
	{
		"month": 1, "day": 1,
		"title": "Yılbaşı Tatili",
		"category": "resmi_tatil", "symbol_type": "event",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "Yeni takvim yılının ilk günü — Devlet memurları ve okullar için tam gün resmi tatil.",
	},
	{
		"month": 3, "day": 18,
		"title": "18 Mart Şehitler Günü ve Çanakkale Deniz Zaferi",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": False, "is_half_day": False,
		"profession_group": "Türk Silahlı Kuvvetleri, Şehit Aileleri ve Gaziler",
		"description": "18 Mart Çanakkale Deniz Zaferi'nin yıl dönümü ve Aziz Şehitlerimizi Anma Günü.",
	},
	{
		"month": 4, "day": 23,
		"title": "23 Nisan Ulusal Egemenlik ve Çocuk Bayramı",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "TBMM'nin açılışı ve Ulusal Egemenlik ve Çocuk Bayramı — Devlet memurları ve okullar için resmi tatil.",
	},
	{
		"month": 5, "day": 1,
		"title": "1 Mayıs Emek ve Dayanışma Günü",
		"category": "resmi_tatil", "symbol_type": "profession",
		"is_civil_holiday": True, "is_half_day": False,
		"profession_group": "Tüm Emekçiler, İşçiler ve Kamu Çalışanları",
		"description": "Emek ve Dayanışma Günü — Devlet memurları ve okullar için tam gün resmi tatil.",
	},
	{
		"month": 5, "day": 19,
		"title": "19 Mayıs Atatürk'ü Anma, Gençlik ve Spor Bayramı",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "Kurtuluş Savaşı'nın başlangıcı — Devlet memurları ve okullar için tam gün resmi tatil.",
	},
	{
		"month": 7, "day": 15,
		"title": "15 Temmuz Demokrasi ve Millî Birlik Günü",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "Demokrasi ve Millî Birlik Günü — Devlet memurları için tam gün resmi tatil.",
	},
	{
		"month": 8, "day": 30,
		"title": "30 Ağustos Zafer Bayramı",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "Büyük Taarruz ve Başkomutanlık Meydan Muharebesi Zaferi — Tam gün resmi tatil.",
	},
	{
		"month": 10, "day": 28,
		"title": "29 Ekim Cumhuriyet Bayramı Arifesi (Saat 13:00'ten Sonra Yarım Gün)",
		"category": "yarim_gun", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": True,
		"start_time": "13:00", "end_time": "17:30",
		"description": "28 Ekim günü saat 13:00'ten itibaren devlet memurları ve okullar için yarım gün resmi tatildir.",
	},
	{
		"month": 10, "day": 29,
		"title": "29 Ekim Cumhuriyet Bayramı",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": True, "is_half_day": False,
		"description": "Cumhuriyetin İlanı — Devlet memurları ve okullar için tam gün resmi milli bayram tatili.",
	},
	{
		"month": 11, "day": 10,
		"title": "10 Kasım Atatürk'ü Anma Günü ve Atatürk Haftası (10-16 Kasım)",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": False, "is_half_day": False,
		"start_time": "09:05",
		"description": "Cumhuriyetimizin kurucusu Gazi Mustafa Kemal Atatürk'ü Anma Günü ve MEB Atatürk Haftası (10-16 Kasım).",
	},
]

# Diyanet Takvimine Göre Yıllara Özel Dini Bayramlar, Arifeler ve Kandiller
RELIGIOUS_DAYS_BY_YEAR: dict[int, list[dict[str, Any]]] = {
	2025: [
		{"date": "2025-01-02", "title": "Regaip Kandili", "holiday": False, "half": False, "desc": "Üç ayların müjdecisi mübarek Regaip Kandili."},
		{"date": "2025-01-26", "title": "Miraç Kandili", "holiday": False, "half": False, "desc": "Mübarek Miraç Kandili gecesi."},
		{"date": "2025-02-13", "title": "Berat Kandili", "holiday": False, "half": False, "desc": "Ramazan ayının müjdecisi Berat Kandili."},
		{"date": "2025-03-01", "title": "Ramazan Ayı Başlangıcı", "holiday": False, "half": False, "desc": "On bir ayın sultanı Ramazan-ı Şerif'in ilk günü."},
		{"date": "2025-03-26", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Bin aydan daha hayırlı olan mübarek Kadir Gecesi."},
		{"date": "2025-03-29", "title": "Ramazan Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
		{"date": "2025-03-30", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Resmi tatil."},
		{"date": "2025-03-31", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2025-04-01", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2025-06-05", "title": "Kurban Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
		{"date": "2025-06-06", "title": "Kurban Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 1. Günü — Resmi tatil."},
		{"date": "2025-06-07", "title": "Kurban Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 2. Günü — Resmi tatil."},
		{"date": "2025-06-08", "title": "Kurban Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 3. Günü — Resmi tatil."},
		{"date": "2025-06-09", "title": "Kurban Bayramı (4. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 4. Günü — Resmi tatil."},
		{"date": "2025-06-26", "title": "Hicri Yılbaşı (1 Muharrem 1447)", "holiday": False, "half": False, "desc": "Hicri 1447 yılının ilk günü."},
		{"date": "2025-07-05", "title": "Aşure Günü (10 Muharrem)", "holiday": False, "half": False, "desc": "Paylaşma ve bereket günü olan 10 Muharrem Aşure Günü."},
		{"date": "2025-09-03", "title": "Mevlid Kandili", "holiday": False, "half": False, "desc": "Peygamber Efendimiz Hz. Muhammed'in (s.a.v.) dünyaya teşrifi — Mevlid Kandili."},
		{"date": "2025-12-25", "title": "Regaip Kandili (1447)", "holiday": False, "half": False, "desc": "Üç ayların ilk kandili olan mübarek Regaip Kandili."},
	],
	2026: [
		{"date": "2026-01-15", "title": "Miraç Kandili", "holiday": False, "half": False, "desc": "Mübarek Miraç Kandili gecesi."},
		{"date": "2026-02-02", "title": "Berat Kandili", "holiday": False, "half": False, "desc": "Ramazan ayının müjdecisi mübarek Berat Kandili."},
		{"date": "2026-02-19", "title": "Ramazan Ayı Başlangıcı (1 Ramazan 1447)", "holiday": False, "half": False, "desc": "On bir ayın sultanı Ramazan-ı Şerif'in ilk günü."},
		{"date": "2026-03-16", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Bin aydan daha hayırlı olan mübarek Kadir Gecesi."},
		{"date": "2026-03-19", "title": "Ramazan Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları ve okullar için yarım gün resmi tatil."},
		{"date": "2026-03-20", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-03-21", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2026-03-22", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2026-05-26", "title": "Kurban Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları ve okullar için yarım gün resmi tatil."},
		{"date": "2026-05-27", "title": "Kurban Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 1. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-28", "title": "Kurban Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 2. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-29", "title": "Kurban Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 3. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-30", "title": "Kurban Bayramı (4. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 4. Günü — Resmi tatil."},
		{"date": "2026-06-16", "title": "Hicri Yılbaşı (1 Muharrem 1448)", "holiday": False, "half": False, "desc": "Hicri 1448 takvim yılının ilk günü."},
		{"date": "2026-06-25", "title": "Aşure Günü (10 Muharrem 1448)", "holiday": False, "half": False, "desc": "10 Muharrem Aşure Günü."},
		{"date": "2026-08-24", "title": "Mevlid Kandili", "holiday": False, "half": False, "desc": "Peygamber Efendimiz Hz. Muhammed'in (s.a.v.) dünyaya teşrifi — Mevlid Kandili."},
		{"date": "2026-12-10", "title": "Regaip Kandili (1448)", "holiday": False, "half": False, "desc": "Üç ayların müjdecisi mübarek Regaip Kandili."},
	],
	2027: [
		{"date": "2027-01-04", "title": "Miraç Kandili", "holiday": False, "half": False, "desc": "Mübarek Miraç Kandili gecesi."},
		{"date": "2027-01-22", "title": "Berat Kandili", "holiday": False, "half": False, "desc": "Mübarek Berat Kandili gecesi."},
		{"date": "2027-02-08", "title": "Ramazan Ayı Başlangıcı (1 Ramazan 1448)", "holiday": False, "half": False, "desc": "Ramazan-ı Şerif'in ilk günü."},
		{"date": "2027-03-05", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Mübarek Kadir Gecesi."},
		{"date": "2027-03-08", "title": "Ramazan Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
		{"date": "2027-03-09", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Resmi tatil."},
		{"date": "2027-03-10", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2027-03-11", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2027-05-15", "title": "Kurban Bayramı Arifesi (Yarım Gün Resmi Tatil)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Yarım gün resmi tatil."},
		{"date": "2027-05-16", "title": "Kurban Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 1. Günü — Resmi tatil."},
		{"date": "2027-05-17", "title": "Kurban Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 2. Günü — Resmi tatil."},
		{"date": "2027-05-18", "title": "Kurban Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 3. Günü — Resmi tatil."},
		{"date": "2027-05-19", "title": "Kurban Bayramı (4. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 4. Günü — Resmi tatil."},
		{"date": "2027-06-06", "title": "Hicri Yılbaşı (1 Muharrem 1449)", "holiday": False, "half": False, "desc": "Hicri 1449 yılının ilk günü."},
		{"date": "2027-06-15", "title": "Aşure Günü", "holiday": False, "half": False, "desc": "10 Muharrem Aşure Günü."},
		{"date": "2027-08-14", "title": "Mevlid Kandili", "holiday": False, "half": False, "desc": "Mübarek Mevlid Kandili."},
	],
}

# MEB Belirli Gün ve Haftalar Çizelgesi + Meslek Grubu Günleri ((Ay, Gün) -> Bilgi)
# Not: FIXED_NATIONAL_DAYS içindeki (3,18), (5,1), (10,29), (11,10) ile mükerrer başlık oluşturmaz (Hata #9 düzeltmesi)
PROFESSION_AND_SPECIAL_DAYS: dict[tuple[int, int], dict[str, str]] = {
	(1, 5): {"title": "MEB - Verem Eğitimi ve Propaganda Haftası (Ocak İlk Haftası)", "prof": "Göğüs Hastalıkları ve Halk Sağlığı Hemşireleri", "cat": "belirli_gun"},
	(1, 7): {"title": "MEB - Beyaz Baston Görme Engelliler Haftası (7-14 Ocak)", "prof": "Özel Eğitim ve Rehabilitasyon Uzmanları", "cat": "belirli_gun"},
	(1, 10): {"title": "10 Ocak Çalışan Gazeteciler Günü & İdareciler Günü", "prof": "Gazeteciler, Basın Emekçileri ve Mülki İdareciler", "cat": "meslek_gunu"},
	(1, 12): {"title": "MEB - Enerji Tasarrufu Haftası (Ocak 2. Haftası)", "prof": "Enerji ve Çevre Mühendisleri", "cat": "belirli_gun"},
	(1, 25): {"title": "MEB - Cüzzam Haftası (25-31 Ocak)", "prof": "Dermatoloji ve Enfeksiyon Kontrol Hemşireleri", "cat": "belirli_gun"},
	(1, 26): {"title": "Dünya Gümrük Günü", "prof": "Gümrük Muhafaza ve Ticaret Personeli", "cat": "meslek_gunu"},
	(2, 4): {"title": "4 Şubat Dünya Kanser Günü", "prof": "Onkoloji Hekimleri ve Koruyucu Sağlık Çalışanları", "cat": "belirli_gun"},
	(2, 9): {"title": "9 Şubat Dünya Sigarayı Bırakma Günü", "prof": "Göğüs Hastalıkları ve Yeşilay Danışmanları", "cat": "belirli_gun"},
	(2, 28): {"title": "MEB - 28 Şubat Sivil Savunma Günü", "prof": "AFAD, Arama Kurtarma ve Sivil Savunma Uzmanları", "cat": "meslek_gunu"},
	(3, 1): {"title": "MEB - Yeşilay Haftası (1-7 Mart) & Muhasebeciler Günü", "prof": "Yeşilay Danışmanları, Rehber Öğretmenler ve Mali Müşavirler", "cat": "belirli_gun"},
	(3, 2): {"title": "MEB - Girişimcilik Haftası (Mart İlk Haftası)", "prof": "Girişimciler ve Mesleki Eğitim Öğretmenleri", "cat": "belirli_gun"},
	(3, 8): {"title": "MEB - 8 Mart Dünya Kadınlar Günü & Bilim ve Teknoloji Haftası (8-14 Mart)", "prof": "Bilim İnsanları, Araştırmacılar ve Tüm Emekçi Kadınlar", "cat": "belirli_gun"},
	(3, 12): {"title": "MEB - 12 Mart İstiklâl Marşı'nın Kabulü ve Mehmet Akif Ersoy'u Anma Günü", "prof": "Edebiyatçılar, Şairler ve Eğitimciler", "cat": "milli_bayram"},
	(3, 14): {"title": "MEB - 14 Mart Tıp Bayramı", "prof": "Hekimler, Hemşireler ve Tüm Sağlık Çalışanları", "cat": "meslek_gunu"},
	(3, 15): {"title": "MEB - Tüketiciyi Koruma Haftası (15-21 Mart)", "prof": "Tüketici Hakları Uzmanları ve Denetçiler", "cat": "belirli_gun"},
	(3, 18): {"title": "MEB - Yaşlılara Saygı Haftası (18-24 Mart)", "prof": "Geriatri Uzmanları ve Sosyal Hizmet Çalışanları", "cat": "belirli_gun"},
	(3, 21): {"title": "MEB - Orman Haftası (21-26 Mart), Türk Dünyası Günü & Down Sendromu Farkındalık Günü", "prof": "Orman Mühendisleri ve Özel Eğitim Öğretmenleri", "cat": "belirli_gun"},
	(3, 22): {"title": "MEB - 22 Mart Dünya Su Günü", "prof": "Çevre ve Su Mühendisleri", "cat": "belirli_gun"},
	(3, 24): {"title": "24 Mart Dünya Tüberküloz (Verem) Günü", "prof": "Halk Sağlığı ve Göğüs Hastalıkları Hemşireleri", "cat": "belirli_gun"},
	(3, 27): {"title": "MEB - 27 Mart Dünya Tiyatrolar Günü", "prof": "Tiyatro Sanatçıları ve Drama Öğretmenleri", "cat": "meslek_gunu"},
	(3, 30): {"title": "MEB - Kütüphaneler Haftası (Mart Son Haftası)", "prof": "Kütüphaneciler ve Arşiv Uzmanları", "cat": "meslek_gunu"},
	(4, 1): {"title": "MEB - Kanser Haftası (1-7 Nisan)", "prof": "Onkoloji Uzmanları ve Halk Sağlığı Tarama Ekipleri", "cat": "belirli_gun"},
	(4, 2): {"title": "MEB - 2 Nisan Dünya Otizm Farkındalık Günü", "prof": "Çocuk Psikiyatristleri ve Özel Eğitimciler", "cat": "belirli_gun"},
	(4, 5): {"title": "5 Nisan Avukatlar Günü", "prof": "Avukatlar ve Hukukçular", "cat": "meslek_gunu"},
	(4, 7): {"title": "MEB - 7 Nisan Dünya Sağlık Günü & Sağlık ve Sosyal Güvenlik Haftası (7-13 Nisan) & KVKK Günü", "prof": "Tüm Sağlık Profesyonelleri, Okul Hemşireleri ve Veri Güvenliği Uzmanları", "cat": "meslek_gunu"},
	(4, 10): {"title": "10 Nisan Polis Haftası ve Türk Polis Teşkilatı Kuruluş Günü", "prof": "Emniyet Mensupları ve Polisler", "cat": "meslek_gunu"},
	(4, 14): {"title": "MEB - Kutlu Doğum / Peygamberimizi Anma ve Kardeşlik Haftası", "prof": "Din Kültürü ve Ahlak Bilgisi Öğretmenleri", "cat": "belirli_gun"},
	(4, 15): {"title": "MEB - Turizm Haftası (15-22 Nisan)", "prof": "Turizm Çalışanları ve Rehberler", "cat": "meslek_gunu"},
	(4, 21): {"title": "MEB - 21-28 Nisan Ebeler Haftası", "prof": "Ebeler ve Ana-Çocuk Sağlığı Çalışanları", "cat": "meslek_gunu"},
	(4, 26): {"title": "MEB - 26 Nisan Dünya Fikri Mülkiyet Günü", "prof": "Patent Vekilleri ve Buluşçular", "cat": "belirli_gun"},
	(4, 28): {"title": "MEB - Kardeşlik Haftası (28 Nisan - 4 Mayıs) & Dünya İş Sağlığı ve Güvenliği Günü", "prof": "İşyeri Hekimleri, İSG Uzmanları ve Kurum Hemşireleri", "cat": "meslek_gunu"},
	(4, 29): {"title": "MEB - 29 Nisan Kût'ül-Amâre Zaferi", "prof": "Tarih Öğretmenleri ve Askerî Tarihçiler", "cat": "milli_bayram"},
	(5, 3): {"title": "MEB - Trafik ve İlkyardım Haftası & Bilişim Haftası (Mayıs İlk Haftası)", "prof": "Paramedikler, ATT'ler, İlkyardım Eğitmenleri ve Bilişim Öğretmenleri", "cat": "meslek_gunu"},
	(5, 4): {"title": "MEB - İş Sağlığı ve Güvenliği Haftası (4-10 Mayıs)", "prof": "İSG Uzmanları ve Okul Güvenlik Ekipleri", "cat": "belirli_gun"},
	(5, 6): {"title": "6 Mayıs Hıdırellez Kültür ve Bahar Bayramı", "prof": "Kültür ve Halkbilim Çalışanları", "cat": "belirli_gun"},
	(5, 10): {"title": "MEB - Engelliler Haftası (10-16 Mayıs), Vakıflar Haftası & Psikologlar Günü", "prof": "Psikologlar, PDR Uzmanları ve Özel Eğitimciler", "cat": "meslek_gunu"},
	(5, 12): {"title": "MEB - 12 Mayıs Dünya Hemşireler Günü & Hemşirelik Haftası (12-18 Mayıs)", "prof": "Okul Revir Hemşireleri ve Tüm Hemşireler", "cat": "meslek_gunu"},
	(5, 14): {"title": "14 Mayıs Eczacılık Günü & Dünya Çiftçiler Günü", "prof": "Eczacılar ve Çiftçiler / Ziraat Mühendisleri", "cat": "meslek_gunu"},
	(5, 15): {"title": "MEB - Aile Haftası (15-21 Mayıs)", "prof": "Aile Danışmanları ve Sosyal Hizmet Uzmanları", "cat": "belirli_gun"},
	(5, 18): {"title": "MEB - Müzeler Haftası (18-24 Mayıs)", "prof": "Müzeciler, Arkeologlar ve Sanat Tarihçileri", "cat": "meslek_gunu"},
	(5, 21): {"title": "MEB - 21 Mayıs Dünya Süt Günü", "prof": "Gıda Mühendisleri ve Beslenme Uzmanları", "cat": "belirli_gun"},
	(5, 25): {"title": "MEB - 25 Mayıs Etik Günü", "prof": "Kamu Etik Kurulu ve Eğitimciler", "cat": "belirli_gun"},
	(5, 29): {"title": "MEB - 29 Mayıs İstanbul'un Fethi", "prof": "Tarihçiler ve Eğitimciler", "cat": "milli_bayram"},
	(6, 5): {"title": "MEB - 5 Haziran Dünya Çevre Günü & Çevre Koruma Haftası (Haziran 2. Haftası)", "prof": "Çevre Mühendisleri ve Doğa Koruma Görevlileri", "cat": "meslek_gunu"},
	(6, 6): {"title": "6 Haziran Diyetisyenler Günü", "prof": "Diyetisyenler ve Beslenme Uzmanları", "cat": "meslek_gunu"},
	(6, 14): {"title": "14 Haziran Dünya Gönüllü Kan Bağışçıları Günü", "prof": "Kızılay ve Kan Merkezi Sağlık Çalışanları", "cat": "belirli_gun"},
	(7, 1): {"title": "1 Temmuz Denizcilik ve Kabotaj Bayramı", "prof": "Denizciler, Kaptanlar ve Gemi Adamları", "cat": "meslek_gunu"},
	(7, 24): {"title": "24 Temmuz Gazeteciler ve Basın Bayramı", "prof": "Basın ve Medya Çalışanları", "cat": "meslek_gunu"},
	(8, 26): {"title": "26 Ağustos Malazgirt Zaferi", "prof": "Tarih Öğretmenleri ve Türk Silahlı Kuvvetleri", "cat": "milli_bayram"},
	(9, 1): {"title": "1 Eylül Dünya Barış Günü & Adli Yıl Açılışı", "prof": "Hakimler, Savcılar ve Hukuk Çalışanları", "cat": "meslek_gunu"},
	(9, 8): {"title": "8 Eylül Dünya Fizyoterapistler Günü & Dünya Okuma Yazma Günü", "prof": "Fizyoterapistler ve Sınıf Öğretmenleri", "cat": "meslek_gunu"},
	(9, 15): {"title": "MEB - İlköğretim Haftası (Eylül 3. Haftası)", "prof": "Öğretmenler, Okul İdarecileri ve Okul Hemşireleri", "cat": "belirli_gun"},
	(9, 19): {"title": "MEB - 19 Eylül Gaziler Günü", "prof": "Kahraman Gazilerimiz ve Şehit Yakınları", "cat": "milli_bayram"},
	(9, 22): {"title": "MEB - Ahilik Kültürü Haftası", "prof": "Esnaf, Sanatkârlar ve Meslek Lisesi Öğretmenleri", "cat": "belirli_gun"},
	(9, 25): {"title": "MEB - İtfaiyecilik Haftası (25 Eylül - 1 Ekim) & Dünya Okul Sütü Günü", "prof": "İtfaiyeciler ve Okul Beslenme Görevlileri", "cat": "meslek_gunu"},
	(9, 26): {"title": "MEB - 26 Eylül Türk Dil Bayramı", "prof": "Türkçe ve Edebiyat Öğretmenleri, Dilbilimciler", "cat": "belirli_gun"},
	(9, 29): {"title": "29 Eylül Dünya Kalp Günü", "prof": "Kardiyologlar ve Koruyucu Sağlık Ekibi", "cat": "belirli_gun"},
	(10, 1): {"title": "1 Ekim Dünya Yaşlılar Günü & Camiler ve Din Görevlileri Haftası (1-7 Ekim)", "prof": "Din Görevlileri ve Geriatri Uzmanları", "cat": "meslek_gunu"},
	(10, 4): {"title": "MEB - 4 Ekim Hayvanları Koruma Günü", "prof": "Veteriner Hekimler", "cat": "meslek_gunu"},
	(10, 10): {"title": "10 Ekim Dünya Ruh Sağlığı Günü", "prof": "Psikiyatristler, Psikologlar ve Okul Rehber Öğretmenleri", "cat": "meslek_gunu"},
	(10, 13): {"title": "13 Ekim Ankara'nın Başkent Oluşu & Dünya Afet Risklerinin Azaltılması Günü", "prof": "AFAD ve Kamu Yönetimi Çalışanları", "cat": "milli_bayram"},
	(10, 15): {"title": "15 Ekim Dünya El Yıkama Günü", "prof": "Okul Sağlığı Hemşireleri ve Enfeksiyon Kontrol Ekibi", "cat": "belirli_gun"},
	(10, 19): {"title": "19 Ekim Muhtarlar Günü", "prof": "Mahalle ve Köy Muhtarları", "cat": "meslek_gunu"},
	(10, 24): {"title": "MEB - 24 Ekim Birleşmiş Milletler Günü", "prof": "Diplomatlar ve Uluslararası İlişkiler Uzmanları", "cat": "belirli_gun"},
	(10, 30): {"title": "MEB - Kızılay Haftası (29 Ekim - 4 Kasım)", "prof": "Kızılay Çalışanları ve İnsani Yardım Gönüllüleri", "cat": "belirli_gun"},
	(11, 2): {"title": "MEB - Lösemili Çocuklar Haftası (2-8 Kasım)", "prof": "Pediatrik Hematoloji Hekimleri ve Hemşireleri", "cat": "belirli_gun"},
	(11, 3): {"title": "MEB - Organ Bağışı ve Nakli Haftası (3-9 Kasım)", "prof": "Organ Nakli Koordinatörleri ve Cerrahlar", "cat": "belirli_gun"},
	(11, 8): {"title": "8 Kasım Dünya Radyoloji Günü", "prof": "Radyoloji Uzmanları ve Teknikerleri", "cat": "meslek_gunu"},
	(11, 12): {"title": "MEB - 12 Kasım Afet Eğitimi Hazırlık Günü", "prof": "AFAD, Okul İş Güvenliği ve İlk Yardım Ekipleri", "cat": "belirli_gun"},
	(11, 14): {"title": "MEB - 14 Kasım Dünya Diyabet Günü", "prof": "Endokrinologlar, Diyabet Eğitim Hemşireleri ve Okul Hemşireleri", "cat": "belirli_gun"},
	(11, 20): {"title": "MEB - 20 Kasım Dünya Çocuk Hakları Günü & Dünya Felsefe Günü", "prof": "Çocuk Sağlığı Uzmanları, Felsefe Öğretmenleri ve Eğitimciler", "cat": "belirli_gun"},
	(11, 22): {"title": "MEB - 22 Kasım Diş Hekimleri Günü & Ağız ve Diş Sağlığı Haftası (21-27 Kasım)", "prof": "Diş Hekimleri ve Ağız-Diş Sağlığı Çalışanları", "cat": "meslek_gunu"},
	(11, 24): {"title": "MEB - 24 Kasım Öğretmenler Günü", "prof": "Öğretmenler ve Eğitim Camiası", "cat": "meslek_gunu"},
	(12, 3): {"title": "MEB - 3 Aralık Dünya Engelliler Günü & Vakıf Haftası", "prof": "Özel Eğitim, Fizyoterapi ve Sosyal Hizmet Uzmanları", "cat": "belirli_gun"},
	(12, 4): {"title": "4 Aralık Dünya Madenciler Günü", "prof": "Maden İşçileri ve Maden Mühendisleri", "cat": "meslek_gunu"},
	(12, 5): {"title": "MEB - 5 Aralık Dünya Mühendisler Günü & Türk Kadınına Seçme ve Seçilme Hakkı Günü", "prof": "Mühendisler ve Hukukçular", "cat": "meslek_gunu"},
	(12, 10): {"title": "MEB - 10 Aralık Dünya İnsan Hakları Günü & İnsan Hakları ve Demokrasi Haftası (10-16 Aralık)", "prof": "Hukukçular, Öğretmenler ve İnsan Hakları Savunucuları", "cat": "belirli_gun"},
	(12, 12): {"title": "MEB - Tutum, Yatırım ve Türk Malları (Yerli Malı) Haftası (12-18 Aralık)", "prof": "Yerli Üreticiler, Esnaf, Öğretmenler ve Veliler", "cat": "belirli_gun"},
	(12, 22): {"title": "22 Aralık Dünya Paramedikler Günü", "prof": "Paramedikler ve 112 Acil Sağlık Ekibi", "cat": "meslek_gunu"},
	(12, 27): {"title": "27 Aralık Atatürk'ün Ankara'ya Gelişi", "prof": "Tarih Öğretmenleri ve Kamu Görevlileri", "cat": "milli_bayram"},
}

# Hafta Boyu Süren Resmi MEB Belirli Gün ve Haftalar (Takvim Yaprağında o haftanın tüm günlerinde görünür!)
MEB_WEEK_RANGES: list[dict[str, Any]] = [
	{"m1": 1, "d1": 1, "m2": 1, "d2": 7, "title": "Verem Eğitimi ve Propaganda Haftası", "focus": "Solunum hijyeni, sınıf havalandırması ve bulaşıcı hastalık farkındalığı"},
	{"m1": 1, "d1": 7, "m2": 1, "d2": 14, "title": "Beyaz Baston Görme Engelliler Haftası", "focus": "Okulda görme taramaları ve erişilebilirlik farkındalığı"},
	{"m1": 1, "d1": 11, "m2": 1, "d2": 17, "title": "Enerji Tasarrufu Haftası", "focus": "Kaynakların verimli kullanımı"},
	{"m1": 3, "d1": 1, "m2": 3, "d2": 7, "title": "Yeşilay Haftası (1–7 Mart)", "focus": "Zararlı alışkanlıklardan ve teknoloji/tütün bağımlılığından korunma"},
	{"m1": 3, "d1": 8, "m2": 3, "d2": 14, "title": "Bilim ve Teknoloji Haftası (8–14 Mart)", "focus": "Bilimsel düşünce ve sağlık teknolojileri"},
	{"m1": 3, "d1": 15, "m2": 3, "d2": 21, "title": "Tüketiciyi Koruma Haftası (15–21 Mart)", "focus": "Güvenli gıda, etiket okuma ve bilinçli tüketim"},
	{"m1": 3, "d1": 18, "m2": 3, "d2": 24, "title": "Yaşlılara Saygı Haftası (18–24 Mart)", "focus": "Kuşaklar arası dayanışma ve büyüklere saygı"},
	{"m1": 3, "d1": 21, "m2": 3, "d2": 26, "title": "Orman Haftası (21–26 Mart)", "focus": "Doğa sevgisi, temiz hava ve çevre sağlığı"},
	{"m1": 3, "d1": 28, "m2": 4, "d2": 3, "title": "Kütüphaneler Haftası", "focus": "Okuma alışkanlığı ve bilgi okuryazarlığı"},
	{"m1": 4, "d1": 1, "m2": 4, "d2": 7, "title": "Kanser Haftası (1–7 Nisan)", "focus": "Erken tanı, sağlıklı beslenme ve hareketli yaşam"},
	{"m1": 4, "d1": 7, "m2": 4, "d2": 13, "title": "Dünya Sağlık Haftası & Sağlık ve Sosyal Güvenlik Haftası (7–13 Nisan)", "focus": "Okul sağlığı taramaları, kişisel hijyen ve koruyucu sağlık"},
	{"m1": 4, "d1": 15, "m2": 4, "d2": 22, "title": "Turizm Haftası (15–22 Nisan)", "focus": "Kültürel miras ve doğa bilinci"},
	{"m1": 4, "d1": 21, "m2": 4, "d2": 28, "title": "Ebeler Haftası (21–28 Nisan)", "focus": "Anne ve çocuk sağlığı"},
	{"m1": 4, "d1": 28, "m2": 5, "d2": 4, "title": "Kardeşlik Haftası (28 Nisan – 4 Mayıs)", "focus": "Akran zorbalığını önleme, empati ve okul içi uyum"},
	{"m1": 5, "d1": 1, "m2": 5, "d2": 7, "title": "Trafik ve İlkyardım Haftası & Bilişim Haftası (Mayıs 1. Haftası)", "focus": "Revir ilk yardım eğitimi, okul servis güvenliği ve güvenli internet"},
	{"m1": 5, "d1": 4, "m2": 5, "d2": 10, "title": "İş Sağlığı ve Güvenliği Haftası (4–10 Mayıs)", "focus": "Okul ve laboratuvar güvenliği, kaza önleme"},
	{"m1": 5, "d1": 10, "m2": 5, "d2": 16, "title": "Engelliler Haftası (10–16 Mayıs) & Vakıflar Haftası", "focus": "Özel gereksinimli öğrencilerle kapsayıcı eğitim ve yardımlaşma"},
	{"m1": 5, "d1": 12, "m2": 5, "d2": 18, "title": "Hemşirelik Haftası (12–18 Mayıs)", "focus": "Okul sağlığı hemşireliği ve şefkatli bakım"},
	{"m1": 5, "d1": 18, "m2": 5, "d2": 24, "title": "Müzeler Haftası (18–24 Mayıs)", "focus": "Tarih ve kültür bilinci"},
	{"m1": 6, "d1": 5, "m2": 6, "d2": 11, "title": "Çevre Koruma Haftası (Haziran 2. Haftası)", "focus": "Temiz okul, geri dönüşüm ve su tasarrufu"},
	{"m1": 9, "d1": 14, "m2": 9, "d2": 20, "title": "İlköğretim Haftası (Eylül 3. Haftası)", "focus": "Okula uyum, yeni eğitim yılı sağlık künyesi ve hijyen kontrolü"},
	{"m1": 9, "d1": 25, "m2": 10, "d2": 1, "title": "İtfaiyecilik ve Yangınından Korunma Haftası (25 Eylül – 1 Ekim)", "focus": "Acil durum tahliye ve yangın güvenliği"},
	{"m1": 10, "d1": 29, "m2": 11, "d2": 4, "title": "Kızılay Haftası (29 Ekim – 4 Kasım)", "focus": "İlk yardım, kan bağışı ve afet dayanışması"},
	{"m1": 11, "d1": 2, "m2": 11, "d2": 8, "title": "Lösemili Çocuklar Haftası (2–8 Kasım)", "focus": "Lösemi farkındalığı ve bağışıklık dostu hijyen"},
	{"m1": 11, "d1": 3, "m2": 11, "d2": 9, "title": "Organ Bağışı ve Nakli Haftası (3–9 Kasım)", "focus": "Yaşam kurtaran organ bağışı bilinci"},
	{"m1": 11, "d1": 10, "m2": 11, "d2": 16, "title": "Atatürk Haftası (10–16 Kasım)", "focus": "Atatürk'ün bilim, eğitim ve sağlığa verdiği önem"},
	{"m1": 11, "d1": 21, "m2": 11, "d2": 27, "title": "Ağız ve Diş Sağlığı Haftası (21–27 Kasım)", "focus": "Doğru diş fırçalama ve okul diş sağlığı taraması"},
	{"m1": 12, "d1": 10, "m2": 12, "d2": 16, "title": "İnsan Hakları ve Demokrasi Haftası (10–16 Aralık)", "focus": "Çocuk hakları, eşitlik ve saygı"},
	{"m1": 12, "d1": 12, "m2": 12, "d2": 18, "title": "Tutum, Yatırım ve Türk Malları (Yerli Malı) Haftası (12–18 Aralık)", "focus": "Sağlıklı mevsim meyve-sebzeleri, alerjen kontrolü ve dengeli beslenme"},
]


def get_active_meb_weeks_for_date(d: date) -> list[dict[str, str]]:
	"""Seçili tarihin içine düştüğü MEB Belirli Gün ve Haftalarını döndürür."""
	matches: list[dict[str, str]] = []
	for w in MEB_WEEK_RANGES:
		try:
			start_d = date(d.year, w["m1"], w["d1"])
			end_d = date(d.year, w["m2"], w["d2"])
			if start_d <= d <= end_d:
				matches.append({"title": w["title"], "focus": w["focus"]})
		except Exception:
			pass
	return matches


# Haftanın gününe göre varsayılan meslek/topluluk takdiri
WEEKDAY_DEFAULT_PROFESSION_THEME: dict[int, tuple[str, str]] = {
	0: ("Okul Sağlığı Hemşireleri & Koruyucu Sağlık Ekibi", "Haftanın ilk iş gününde öğrencilerin sağlığı için görev yapan okul revir sorumluları."),
	1: ("Öğretmenler & Okul Rehberlik (PDR) Servisi", "Öğrencilerin akademik ve psikososyal gelişimine rehberlik eden eğitimciler."),
	2: ("Hekimler, Aile Sağlığı & Toplum Sağlığı Çalışanları", "Birinci basamak sağlık hizmetleri ve koruyucu hekimlik emekçileri."),
	3: ("Kamu Görevlileri, İdareciler & Bürokratlar", "Kamu hizmetlerinin aksamadan yürütülmesini sağlayan devlet memurları."),
	4: ("Destek Personeli, Beslenme & Hijyen Görevlileri", "Okul hijyeni, güvenli beslenme ve sağlıklı eğitim ortamını koruyan emekçiler."),
	5: ("112 Acil Sağlık, Nöbetçi Eczane & Güvenlik Görevlileri", "Hafta sonu nöbet başında toplum sağlığı ve güvenliğini koruyanlar."),
	6: ("Aileler, Veliler & Gönüllü Yardım Kuruluşları", "Çocukların sağlıklı büyümesinde okul ile iş birliği yapan aileler."),
}


def _nth_weekday_of_month(year: int, month: int, weekday: int, nth: int) -> date:
	"""
	Belirtilen yıl ve ayda `nth` (1..5) sıradaki `weekday` (5=Cumartesi, 6=Pazar) tarihini döndürür.
	ÖSYM ve MEB sınavlarının her yılda gerçek Cumartesi/Pazar gününe düşmesini sağlar (Hata #6 düzeltmesi).
	"""
	first_day = date(year, month, 1)
	diff = (weekday - first_day.weekday()) % 7
	day_num = 1 + diff + (nth - 1) * 7
	_, max_d = calendar.monthrange(year, month)
	while day_num > max_d:
		day_num -= 7
	return date(year, month, day_num)


# Yerleşik ÖSYM & MEB Merkezi Sınav Takvimi (Her yıla hafta sonu hizalı + 2026 kesin tarihli)
BUILTIN_OSYM_MEB_EXAMS: list[dict[str, Any]] = [
	{"code": "MSÜ", "month": 3, "day_2026": 1, "weekday": 6, "nth": 1, "time": "10:15", "title": "ÖSYM - MSÜ (Millî Savunma Üniversitesi Askerî Öğrenci Aday Belirleme Sınavı)", "desc": "ÖSYM tarafından Pazar günü uygulanan MSÜ sınavı (Saat: 10:15)."},
	{"code": "YDS / YÖKDİL", "month": 3, "day_2026": 8, "weekday": 6, "nth": 2, "time": "10:15", "title": "ÖSYM - YÖKDİL / 1 (Yükseköğretim Kurumları Yabancı Dil Sınavı)", "desc": "ÖSYM YÖKDİL İlkbahar dönemi Pazar oturumu."},
	{"code": "YDS / YÖKDİL", "month": 4, "day_2026": 5, "weekday": 6, "nth": 1, "time": "10:15", "title": "ÖSYM - YDS / 1 (Yabancı Dil Bilgisi Seviye Tespit Sınavı)", "desc": "ÖSYM YDS İlkbahar dönemi Pazar oturumu."},
	{"code": "İOKBS Bursluluk", "month": 4, "day_2026": 26, "weekday": 6, "nth": 4, "time": "10:00", "title": "MEB - İOKBS (İlköğretim ve Ortaöğretim Kurumları Bursluluk Sınavı)", "desc": "MEB merkezi bursluluk Pazar sınavı."},
	{"code": "ALES", "month": 5, "day_2026": 10, "weekday": 6, "nth": 2, "time": "10:15", "title": "ÖSYM - ALES / 1 (Akademik Personel ve Lisansüstü Eğitimi Giriş Sınavı)", "desc": "ÖSYM ALES-1 Pazar oturumu."},
	{"code": "LGS", "month": 6, "day_2026": 14, "weekday": 6, "nth": 2, "time": "09:30", "title": "MEB - LGS (Liselere Geçiş Sistemi Merkezi Sınavı)", "desc": "8. sınıf öğrencileri için MEB LGS Sözel (09:30) ve Sayısal (11:30) oturumları."},
	{"code": "YKS (TYT-AYT-YDT)", "month": 6, "day_2026": 20, "weekday": 5, "nth": 3, "time": "10:15", "title": "ÖSYM - YKS 1. Oturum: TYT (Temel Yeterlilik Testi)", "desc": "Yükseköğretim Kurumları Sınavı Cumartesi TYT oturumu (Saat 10:15)."},
	{"code": "YKS (TYT-AYT-YDT)", "month": 6, "day_2026": 21, "weekday": 6, "nth": 3, "time": "10:15", "title": "ÖSYM - YKS 2. Oturum: AYT & 3. Oturum: YDT", "desc": "YKS Pazar Alan Yeterlilik Testleri (10:15) ve Yabancı Dil Testi (15:45)."},
	{"code": "DGS", "month": 7, "day_2026": 19, "weekday": 6, "nth": 3, "time": "10:15", "title": "ÖSYM - DGS (Dikey Geçiş Sınavı)", "desc": "Meslek yüksekokulu mezunları için Dikey Geçiş Sınavı."},
	{"code": "KPSS", "month": 9, "day_2026": 6, "weekday": 6, "nth": 1, "time": "10:15", "title": "ÖSYM - KPSS Lisans (Genel Yetenek - Genel Kültür & Eğitim Bilimleri)", "desc": "Kamu Personel Seçme Sınavı Lisans Pazar oturumu."},
	{"code": "KPSS", "month": 10, "day_2026": 4, "weekday": 6, "nth": 1, "time": "10:15", "title": "ÖSYM - KPSS Ön Lisans Oturumu", "desc": "Kamu Personel Seçme Sınavı Ön Lisans Pazar oturumu."},
	{"code": "KPSS", "month": 10, "day_2026": 25, "weekday": 6, "nth": 4, "time": "10:15", "title": "ÖSYM - KPSS Ortaöğretim (Lise) Oturumu", "desc": "Kamu Personel Seçme Sınavı Ortaöğretim Pazar oturumu."},
	{"code": "ALES", "month": 11, "day_2026": 22, "weekday": 6, "nth": 4, "time": "10:15", "title": "ÖSYM - ALES / 3 (Sonbahar Dönemi)", "desc": "ÖSYM ALES-3 Pazar oturumu."},
]


def _resolve_exam_date_for_year(year: int, ex: dict[str, Any]) -> str:
	"""2026 için kesin günü, diğer yıllar için ilgili haftanın Cumartesi/Pazar gününü döndürür."""
	if year == 2026 and ex.get("day_2026"):
		return f"{year:04d}-{ex['month']:02d}-{int(ex['day_2026']):02d}"
	d_obj = _nth_weekday_of_month(year, int(ex["month"]), int(ex.get("weekday", 6)), int(ex.get("nth", 2)))
	return d_obj.isoformat()


# ---------------------------------------------------------------------------
# YERLEŞİK HADİS-İ ŞERİF VE ÖZLÜ SÖZ / ATASÖZÜ KÜTÜPHANESİ (ÇEVRİMDIŞI DESTEK)
# ---------------------------------------------------------------------------

BUILTIN_HADITHS: list[tuple[str, str]] = [
	("İki nimet vardır ki insanların çoğu onların kıymetini bilmekte aldanmıştır: Sağlık ve boş vakit.", "Buhârî, Rikâk, 1"),
	("İnsanların en hayırlısı, insanlara en çok faydalı olandır.", "Kudâî, Müsnedü'ş-Şihâb, II, 223"),
	("Kolaylaştırınız, zorlaştırmayınız; müjdeleyiniz, nefret ettirmeyiniz.", "Buhârî, İlim, 11"),
	("Temizlik imanın yarısıdır.", "Müslim, Tahâret, 1"),
	("Merhamet etmeyene merhamet olunmaz.", "Buhârî, Edeb, 18"),
	("Hiçbir baba, çocuğuna güzel terbiyeden daha kıymetli bir miras bırakmamıştır.", "Tirmizî, Birr, 33"),
	("Sizin en hayırlınız, ahlakı en güzel olanınızdır.", "Buhârî, Edeb, 38"),
	("Allah, bir kulun yaptığı işi sağlam ve güzel yapmasından hoşnut olur.", "Taberânî, el-Mu'cemü'l-Evsat"),
	("Hastayı ziyaret edin, aç olanı doyurun ve sıkıntıda olana el uzatın.", "Buhârî, Merdâ, 4"),
	("Küçüklerimize merhamet etmeyen, büyüklerimize saygı göstermeyen bizden değildir.", "Tirmizî, Birr, 15"),
	("Güzel söz sadakadır; kardeşinin yüzüne tebessüm etmen de sadakadır.", "Buhârî, Edeb, 34; Tirmizî, Birr, 36"),
	("Beş şey gelmeden önce beş şeyin kıymetini bil: İhtiyarlığından önce gençliğinin, hastalığından önce sağlığının...", "Hâkim, el-Müstedrek, IV, 341"),
	("İlim öğrenmek için yola çıkan kimseye Allah cennet yolunu kolaylaştırır.", "Müslim, Zikir, 39"),
	("Mümin, elinden ve dilinden insanların güvende olduğu kimsedir.", "Tirmizî, Îmân, 12"),
	("Tedavi olunuz; çünkü Allah yarattığı her hastalığın şifasını da yaratmıştır.", "Ebû Dâvûd, Tıb, 1"),
]

BUILTIN_QUOTES: list[tuple[str, str]] = [
	("Halk içinde muteber bir nesne yok devlet gibi; olmaya devlet cihanda bir nefes sıhhat gibi.", "Kanuni Sultan Süleyman (Muhibbî)"),
	("Beni Türk hekimlerine emanet ediniz.", "Gazi Mustafa Kemal Atatürk"),
	("Eğitimdir ki bir milleti ya özgür, bağımsız, şanlı, yüksek bir topluluk halinde yaşatır; ya da esaret ve sefalete terk eder.", "Gazi Mustafa Kemal Atatürk"),
	("Güneş girmeyen eve doktor girer.", "Türk Atasözü"),
	("Sağlam kafa sağlam vücutta bulunur.", "Gazi Mustafa Kemal Atatürk"),
	("Bir okul açan, bir hapishane kapatır.", "Victor Hugo"),
	("Hastalık yoktur, hasta vardır; her çocuğun büyüme ve gelişimi kendine özgü bir yolculuktur.", "Klinik Tıp İlkesi"),
	("Şefkat ve merhamet, en güçlü ilaçtan daha hızlı iyileştirir.", "Florence Nightingale"),
	("Çocuk spor ve sağlıkla büyüdüğünde geleceğin en sağlam temeli atılmış olur.", "Okul Sağlığı Özdeyişi"),
	("İlim ilim bilmektir, ilim kendin bilmektir; sen kendini bilmezsen ya nice okumaktır.", "Yunus Emre"),
	("Sabır, acı bir kök; fakat meyvesi çok tatlıdır.", "Mevlânâ Celâleddîn-i Rûmî"),
	("Damlaya damlaya göl olur; küçük koruyucu sağlık alışkanlıkları ömür boyu sıhhat kazandırır.", "Türk Atasözü & Sağlık Notu"),
	("Bir mum diğer bir mumu tutuşturmakla ışığından hiçbir şey kaybetmez.", "Mevlânâ Celâleddîn-i Rûmî"),
	("Temizlik sağlığın, düzenli takip ise güvenli okulun anahtarıdır.", "Okul Revir İlkeleri"),
	("En büyük zenginlik sağlık, en büyük mutluluk ise huzurdur.", "Atasözü"),
]


def get_builtin_daily_quote_and_hadith(d: date) -> dict[str, str]:
	"""Seçili tarihe göre deterministik olarak yerleşik Hadis-i Şerif ve Özlü Söz döndürür."""
	idx_h = (d.toordinal() * 7 + d.day) % len(BUILTIN_HADITHS)
	idx_q = (d.toordinal() * 11 + d.month) % len(BUILTIN_QUOTES)
	h_txt, h_src = BUILTIN_HADITHS[idx_h]
	q_txt, q_aut = BUILTIN_QUOTES[idx_q]
	return {
		"hadith_text": h_txt,
		"hadith_source": h_src,
		"quote_text": q_txt,
		"quote_author": q_aut,
	}


# ---------------------------------------------------------------------------
# VERİTABANI YARDIMCILARI VE OTOMATİK RESMİ TATİL / MEB BELİRLİ GÜN YÜKLEYİCİ
# ---------------------------------------------------------------------------

def ensure_year_builtin_holidays_seeded(year: int, connect_db_fn: Callable, include_special_days: bool = False) -> int:
	"""
	Belirtilen yıl için resmi tatiller, yarım gün arife/28 Ekim çalışmaları, dini bayramlar
	ve isteğe bağlı olarak MEB Belirli Gün ve Haftaları + Meslek Günlerini veritabanına işler.
	"""
	added = 0
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db_fn() as conn:
		existing_rows = conn.execute(
			"SELECT event_date, title FROM calendar_events WHERE substr(event_date, 1, 4) = ?",
			(str(year),),
		).fetchall()
		existing_keys = {(r["event_date"], (r["title"] or "").strip().lower()) for r in existing_rows}

		# 1. Sabit Miladi Resmi Tatiller ve Ulusal Günler
		for item in FIXED_NATIONAL_DAYS:
			ev_date = f"{year:04d}-{item['month']:02d}-{item['day']:02d}"
			t_key = (ev_date, item["title"].strip().lower())
			if t_key in existing_keys:
				continue
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'builtin', 0, ?)""",
				(
					ev_date,
					item["title"],
					item["category"],
					item["symbol_type"],
					1 if item.get("is_civil_holiday") else 0,
					1 if item.get("is_half_day") else 0,
					item.get("profession_group", ""),
					item.get("start_time", ""),
					item.get("end_time", ""),
					item.get("description", ""),
					now_iso,
				),
			)
			existing_keys.add(t_key)
			added += 1

		# 2. Yıla Özel Dini Bayramlar, Arifeler ve Kandiller
		for r_item in RELIGIOUS_DAYS_BY_YEAR.get(year, []):
			ev_date = r_item["date"]
			t_key = (ev_date, r_item["title"].strip().lower())
			if t_key in existing_keys:
				continue
			cat = "yarim_gun" if r_item.get("half") else "dini_bayram"
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, ?, 'crescent', ?, ?, '', ?, '', ?, 'builtin', 0, ?)""",
				(
					ev_date,
					r_item["title"],
					cat,
					1 if r_item.get("holiday") else 0,
					1 if r_item.get("half") else 0,
					"13:00" if r_item.get("half") else "",
					r_item.get("desc", ""),
					now_iso,
				),
			)
			existing_keys.add(t_key)
			added += 1

		# 3. MEB Belirli Gün ve Haftaları + Meslek Grubu Günleri
		if include_special_days:
			for (m_no, d_no), sp_info in PROFESSION_AND_SPECIAL_DAYS.items():
				ev_date = f"{year:04d}-{m_no:02d}-{d_no:02d}"
				t_key = (ev_date, sp_info["title"].strip().lower())
				if t_key in existing_keys:
					continue
				cat = sp_info.get("cat", "belirli_gun")
				sym = "tr_flag" if cat == "milli_bayram" else ("profession" if cat == "meslek_gunu" else "event")
				conn.execute(
					"""INSERT INTO calendar_events
					   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
					    profession_group, start_time, end_time, description, source, is_completed, created_at)
					   VALUES (?, ?, ?, ?, 0, 0, ?, '', '', ?, 'builtin', 0, ?)""",
					(
						ev_date,
						sp_info["title"],
						cat,
						sym,
						sp_info.get("prof", ""),
						f"MEB Belirli Gün ve Haftalar / Meslek Grubu: {sp_info.get('prof', '')}",
						now_iso,
					),
				)
				existing_keys.add(t_key)
				added += 1

	return added


def seed_selected_exams_to_db(year: int, selected_exams: list[str], connect_db_fn: Callable) -> int:
	"""Seçilen ÖSYM ve MEB sınavlarını ilgili yılın gerçek hafta sonu günlerine hizalayarak işler."""
	added = 0
	now_iso = datetime.now().isoformat(timespec="seconds")
	sel_set = set(selected_exams)
	with connect_db_fn() as conn:
		existing_rows = conn.execute(
			"SELECT event_date, title FROM calendar_events WHERE substr(event_date, 1, 4) = ?",
			(str(year),),
		).fetchall()
		existing_keys = {(r["event_date"], (r["title"] or "").strip().lower()) for r in existing_rows}

		for ex in BUILTIN_OSYM_MEB_EXAMS:
			if sel_set and ex["code"] not in sel_set:
				continue
			ev_date = _resolve_exam_date_for_year(year, ex)
			t_key = (ev_date, ex["title"].strip().lower())
			if t_key in existing_keys:
				continue
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, 'osym_sinav', 'exam', 0, 0, 'Öğrenciler, Öğretmenler ve Sınav Görevlileri', ?, '', ?, 'builtin_exam', 0, ?)""",
				(ev_date, ex["title"], ex["time"], ex["desc"], now_iso),
			)
			existing_keys.add(t_key)
			added += 1
	return added


def save_ai_generated_events_to_db(events: list[dict[str, Any]], connect_db_fn: Callable, upsert: bool = False) -> int:
	"""Gemini AI tarafından üretilen takvim kayıtlarını tekilleştirerek SQLite veritabanına kaydeder."""
	added = 0
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db_fn() as conn:
		for ev in events:
			if not isinstance(ev, dict):
				continue
			raw_date = str(ev.get("event_date") or ev.get("date") or ev.get("tarih") or "").strip()[:10]
			if re.match(r"^\d{2}[./-]\d{2}[./-]\d{4}$", raw_date):
				d_p, m_p, y_p = re.split(r"[./-]", raw_date)
				ev_date = f"{y_p}-{m_p}-{d_p}"
			else:
				ev_date = raw_date
			title = str(ev.get("title") or ev.get("name") or ev.get("baslik") or "").strip()
			if len(ev_date) != 10 or not title:
				continue
			try:
				date.fromisoformat(ev_date)
			except Exception:
				continue
			cat = str(ev.get("category") or "belirli_gun").strip()
			if cat not in CATEGORY_LABELS:
				cat = "belirli_gun"
			sym = str(ev.get("symbol_type") or CATEGORY_LABELS[cat][2]).strip()
			if sym not in SYMBOL_OPTIONS:
				sym = CATEGORY_LABELS[cat][2]
			is_half = 1 if (bool(ev.get("is_half_day")) or cat == "yarim_gun") else 0
			is_hol = 1 if (bool(ev.get("is_civil_holiday")) or is_half or cat in ("resmi_tatil", "secim_idari_tatil")) else 0
			prof = str(ev.get("profession_group") or "").strip()
			st_time = str(ev.get("start_time") or "").strip()[:5]
			en_time = str(ev.get("end_time") or "").strip()[:5]
			desc = str(ev.get("description") or "").strip()
			dup = conn.execute(
				"SELECT id FROM calendar_events WHERE event_date = ? AND lower(trim(title)) = lower(trim(?))",
				(ev_date, title),
			).fetchone()
			if dup:
				if upsert:
					conn.execute(
						"""UPDATE calendar_events
						   SET category=?, symbol_type=?, is_civil_holiday=?, is_half_day=?,
						       profession_group=?, start_time=?, end_time=?, description=?
						   WHERE id=?""",
						(cat, sym, is_hol, is_half, prof, st_time, en_time, desc, int(dup["id"])),
					)
					added += 1
				continue
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'gemini_ai', 0, ?)""",
				(ev_date, title, cat, sym, is_hol, is_half, prof, st_time, en_time, desc, now_iso),
			)
			added += 1
	return added


# Resmi MEB Çalışma Takvimi Ön Tanımlı Eğitim Yılı, Ara Tatil ve Yarıyıl (Sömestr) Tarihleri
MEB_ACADEMIC_CALENDAR_PRESETS: dict[str, dict[str, str]] = {
	"2024-2025": {
		"start_date": "2024-09-09",
		"break1_start": "2024-11-11",
		"break1_end": "2024-11-15",
		"semester_start": "2025-01-20",
		"semester_end": "2025-01-31",
		"break2_start": "2025-03-31",
		"break2_end": "2025-04-04",
		"end_date": "2025-06-20",
	},
	"2025-2026": {
		"start_date": "2025-09-08",
		"break1_start": "2025-11-10",
		"break1_end": "2025-11-14",
		"semester_start": "2026-01-19",
		"semester_end": "2026-01-30",
		"break2_start": "2026-03-16",
		"break2_end": "2026-03-20",
		"end_date": "2026-06-26",
	},
	"2026-2027": {
		"start_date": "2026-09-07",
		"break1_start": "2026-11-09",
		"break1_end": "2026-11-13",
		"semester_start": "2027-01-18",
		"semester_end": "2027-01-29",
		"break2_start": "2027-04-05",
		"break2_end": "2027-04-09",
		"end_date": "2027-06-18",
	},
	"2027-2028": {
		"start_date": "2027-09-06",
		"break1_start": "2027-11-08",
		"break1_end": "2027-11-12",
		"semester_start": "2028-01-24",
		"semester_end": "2028-02-04",
		"break2_start": "2028-04-10",
		"break2_end": "2028-04-14",
		"end_date": "2028-06-16",
	},
	"2028-2029": {
		"start_date": "2028-09-11",
		"break1_start": "2028-11-13",
		"break1_end": "2028-11-17",
		"semester_start": "2029-01-22",
		"semester_end": "2029-02-02",
		"break2_start": "2029-04-09",
		"break2_end": "2029-04-13",
		"end_date": "2029-06-22",
	},
}


def get_default_meb_academic_break_dates(
	academic_year_name: str,
	start_date_iso: str = "",
	end_date_iso: str = "",
) -> dict[str, str]:
	"""
	Eğitim öğretim yılı adına (örn. '2026-2027') göre MEB standart başlangıç, bitiş,
	1. dönem ara tatili, yarıyıl (sömestr) tatili ve 2. dönem ara tatili tarihlerini döndürür.
	"""
	ay_norm = re.sub(r"\s+", "", str(academic_year_name or "")).replace("/", "-").replace("_", "-")
	m_yr = re.search(r"(20\d{2})-(20\d{2})", ay_norm)
	if m_yr:
		key_yr = f"{m_yr.group(1)}-{m_yr.group(2)}"
		if key_yr in MEB_ACADEMIC_CALENDAR_PRESETS:
			return dict(MEB_ACADEMIC_CALENDAR_PRESETS[key_yr])
		y1, y2 = int(m_yr.group(1)), int(m_yr.group(2))
	else:
		try:
			y1 = date.fromisoformat(str(start_date_iso)[:10]).year
		except Exception:
			y1 = date.today().year
		y2 = y1 + 1
		key_yr = f"{y1}-{y2}"
		if key_yr in MEB_ACADEMIC_CALENDAR_PRESETS:
			return dict(MEB_ACADEMIC_CALENDAR_PRESETS[key_yr])

	sd = _nth_weekday_of_month(y1, 9, 0, 2)  # Eylül 2. Pazartesi
	b1_s = _nth_weekday_of_month(y1, 11, 0, 2)  # Kasım 2. Pazartesi
	b1_e = b1_s + timedelta(days=4)
	sem_s = _nth_weekday_of_month(y2, 1, 0, 3)  # Ocak 3. Pazartesi
	sem_e = sem_s + timedelta(days=11)  # 2. Cuma
	b2_s = _nth_weekday_of_month(y2, 4, 0, 1)  # Nisan 1. Pazartesi
	b2_e = b2_s + timedelta(days=4)
	ed = _nth_weekday_of_month(y2, 6, 4, 3)  # Haziran 3. Cuma
	return {
		"start_date": sd.isoformat(),
		"break1_start": b1_s.isoformat(),
		"break1_end": b1_e.isoformat(),
		"semester_start": sem_s.isoformat(),
		"semester_end": sem_e.isoformat(),
		"break2_start": b2_s.isoformat(),
		"break2_end": b2_e.isoformat(),
		"end_date": ed.isoformat(),
	}


def sync_academic_year_breaks_to_calendar(
	connect_db_fn: Callable,
	*,
	academic_year_name: str,
	start_date: str,
	end_date: str,
	break1_start: str,
	break1_end: str,
	semester_start: str,
	semester_end: str,
	break2_start: str,
	break2_end: str,
) -> dict[str, int]:
	"""
	Eğitim öğretim yılı açılış/kapanış günlerini, 1. ve 2. dönem ara tatillerini ve
	yarıyıl (sömestr / 15 tatil) günlerini calendar_events tablosuna işler.
	"""
	ay_lbl = (academic_year_name or "").strip() or "Eğitim Yılı"
	now_iso = datetime.now().isoformat(timespec="seconds")

	def _parse_d(val: str) -> date | None:
		try:
			return date.fromisoformat(str(val or "").strip()[:10])
		except Exception:
			return None

	def _iter_days(d_start: date | None, d_end: date | None, max_span: int = 35) -> list[date]:
		if not d_start or not d_end or d_end < d_start:
			return []
		span = min((d_end - d_start).days, max_span)
		return [d_start + timedelta(days=i) for i in range(span + 1)]

	d_start = _parse_d(start_date)
	d_end = _parse_d(end_date)
	d_b1_s, d_b1_e = _parse_d(break1_start), _parse_d(break1_end)
	d_sem_s, d_sem_e = _parse_d(semester_start), _parse_d(semester_end)
	d_b2_s, d_b2_e = _parse_d(break2_start), _parse_d(break2_end)

	b1_days = _iter_days(d_b1_s, d_b1_e)
	sem_days = _iter_days(d_sem_s, d_sem_e)
	b2_days = _iter_days(d_b2_s, d_b2_e)

	records_to_insert: list[tuple[str, str, str, str, int, int, str, str]] = []

	if d_start:
		records_to_insert.append((
			d_start.isoformat(),
			f"🏫 {ay_lbl} Eğitim Öğretim Yılı Başlangıcı (1. Dönem Açılış)",
			"belirli_gun",
			"event",
			0,
			0,
			"Öğrenciler, Öğretmenler ve Okul Sağlığı Ekibi",
			f"{ay_lbl} eğitim öğretim yılı 1. dönem ders başlangıcı.",
		))

	for idx, cur_d in enumerate(b1_days, start=1):
		records_to_insert.append((
			cur_d.isoformat(),
			f"🍂 1. Dönem Ara Tatili ({idx}. Gün)",
			"secim_idari_tatil",
			"ballot",
			1,
			0,
			"Öğrenciler ve Öğretmenler",
			f"{ay_lbl} MEB 1. Dönem Kasım Ara Tatili ({d_b1_s.strftime('%d.%m.%Y')} - {d_b1_e.strftime('%d.%m.%Y')}).",
		))

	if d_sem_s:
		first_term_end = d_sem_s - timedelta(days=3 if d_sem_s.weekday() == 0 else 1)
		records_to_insert.append((
			first_term_end.isoformat(),
			f"🎓 1. Dönem Sonu (Karne Günü)",
			"belirli_gun",
			"event",
			0,
			0,
			"Öğrenciler ve Öğretmenler",
			f"{ay_lbl} eğitim öğretim yılı 1. dönem karne günü.",
		))

	for idx, cur_d in enumerate(sem_days, start=1):
		records_to_insert.append((
			cur_d.isoformat(),
			f"⛄ Yarıyıl Tatili / Sömestr ({idx}. Gün)",
			"secim_idari_tatil",
			"ballot",
			1,
			0,
			"Öğrenciler ve Öğretmenler",
			f"{ay_lbl} MEB Yarıyıl / Sömestr (15 Tatil) Dönemi ({d_sem_s.strftime('%d.%m.%Y')} - {d_sem_e.strftime('%d.%m.%Y')}).",
		))

	if d_sem_e:
		second_term_start = d_sem_e + timedelta(days=3 if d_sem_e.weekday() == 4 else 1)
		records_to_insert.append((
			second_term_start.isoformat(),
			f"🏫 2. Dönem Başlangıcı",
			"belirli_gun",
			"event",
			0,
			0,
			"Öğrenciler ve Öğretmenler",
			f"{ay_lbl} eğitim öğretim yılı 2. dönem ders başlangıcı.",
		))

	for idx, cur_d in enumerate(b2_days, start=1):
		records_to_insert.append((
			cur_d.isoformat(),
			f"🌸 2. Dönem Ara Tatili ({idx}. Gün)",
			"secim_idari_tatil",
			"ballot",
			1,
			0,
			"Öğrenciler ve Öğretmenler",
			f"{ay_lbl} MEB 2. Dönem Ara Tatili ({d_b2_s.strftime('%d.%m.%Y')} - {d_b2_e.strftime('%d.%m.%Y')}).",
		))

	if d_end:
		records_to_insert.append((
			d_end.isoformat(),
			f"🏁 {ay_lbl} Eğitim Öğretim Yılı Sonu (Karne & Kapanış Günü)",
			"belirli_gun",
			"event",
			0,
			0,
			"Öğrenciler, Öğretmenler ve Okul Sağlığı Ekibi",
			f"{ay_lbl} eğitim öğretim yılı 2. dönem sonu ve yaz tatili başlangıcı.",
		))

	with connect_db_fn() as conn:
		conn.execute("DELETE FROM calendar_events WHERE source = 'meb_akademik_takvim'")
		for ev_date, title, cat, sym, is_hol, is_half, prof, desc in records_to_insert:
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, ?, ?, ?, ?, ?, '', '', ?, 'meb_akademik_takvim', 0, ?)""",
				(ev_date, title, cat, sym, is_hol, is_half, prof, desc, now_iso),
			)
		conn.commit()

	return {
		"total_events": len(records_to_insert),
		"break1_days": len(b1_days),
		"semester_days": len(sem_days),
		"break2_days": len(b2_days),
	}



def _sync_selected_date_when_month_changes(new_year: int, new_month: int, today: date) -> None:
	"""
	Ay veya yıl değiştirildiğinde sağdaki Takvim Yaprağının eski ayda takılı kalmasını önler (Hata #4/#7 düzeltmesi).
	"""
	cur_iso = str(ui.session_state.get("cal_selected_date") or today.isoformat())[:10]
	try:
		cur_d = date.fromisoformat(cur_iso)
	except Exception:
		cur_d = today
	if cur_d.year == new_year and cur_d.month == new_month:
		return
	if today.year == new_year and today.month == new_month:
		ui.session_state["cal_selected_date"] = today.isoformat()
	else:
		_, max_d = calendar.monthrange(new_year, new_month)
		target_day = min(cur_d.day, max_d)
		ui.session_state["cal_selected_date"] = date(new_year, new_month, target_day).isoformat()


def _handle_calendar_event_inline_click(connect_db_fn: Callable, *, set_flash: bool = True) -> bool:
	"""Takvim kutucukları içindeki '✓' (yapıldı) ve '✕' (kaldır) hızlı işlem butonlarını işler."""
	try:
		ctx = ui._get_ctx()
		clicked_wid = getattr(ctx, "clicked_button", None)
	except Exception:
		return False
	if not isinstance(clicked_wid, str):
		return False

	if clicked_wid.startswith("btn_cal_ev_done_"):
		try:
			ev_id = int(clicked_wid[len("btn_cal_ev_done_"):])
		except ValueError:
			return False
		with connect_db_fn() as conn:
			row = conn.execute(
				"SELECT title, is_completed FROM calendar_events WHERE id = ?",
				(ev_id,),
			).fetchone()
			if row is not None:
				new_done = 0 if bool(row["is_completed"]) else 1
				conn.execute(
					"UPDATE calendar_events SET is_completed = ? WHERE id = ?",
					(new_done, ev_id),
				)
				conn.commit()
				if set_flash:
					if new_done == 1:
						ui.session_state["_cal_expand_completed"] = True
						ui.session_state["_cal_flash_msg"] = f"✅ '{row['title']}' tamamlandı olarak işaretlendi ve Tamamlanan Etkinlikler bölümüne taşındı."
					else:
						ui.session_state["_cal_flash_msg"] = f"↩️ '{row['title']}' tekrar Günlük Takvim Yaprağı'na geri alındı."
		ctx.clicked_button = None
		return True

	if clicked_wid.startswith("btn_cal_ev_del_"):
		try:
			ev_id = int(clicked_wid[len("btn_cal_ev_del_"):])
		except ValueError:
			return False
		with connect_db_fn() as conn:
			row = conn.execute(
				"SELECT title FROM calendar_events WHERE id = ?",
				(ev_id,),
			).fetchone()
			if row is not None:
				conn.execute("DELETE FROM calendar_events WHERE id = ?", (ev_id,))
				conn.commit()
				if set_flash:
					ui.session_state["_cal_flash_msg"] = f"✕ '{row['title']}' takvimden silindi."
		ctx.clicked_button = None
		return True

	return False


# ---------------------------------------------------------------------------
# ANA TAKVİM VE TAKVİM YAPRAĞI EKRANI RENDER FONKSİYONU
# ---------------------------------------------------------------------------

def render_calendar_and_planning_page(
	*,
	connect_db_fn: Callable,
	get_app_settings_fn: Callable,
	save_app_settings_fn: Callable,
	get_active_gemini_config_fn: Callable,
	get_tr_provinces_fn: Callable,
	is_section_visible_fn: Callable | None = None,
) -> None:
	"""
	GELİŞİM & PLAN -> 📅 Takvim & Planlama ana sayfasını oluşturur.
	- Sol Panel: Her gün için TEK ve doğrudan tıklanabilir kutucuklu Aylık Takvim Izgarası
	- Sağ Panel: Ezan vakitleri, canlı sayaç ve takvim yaprağı arkası (Hadis/Özlü Söz)
	  doğrudan içine gömülü Birleşik 'Günlük Takvim Yaprağı' + Etkinlik Ekleme/Yönetim
	"""
	today = date.today()
	cfg = get_app_settings_fn()
	ai_enabled, ai_key, ai_model = get_active_gemini_config_fn()

	# Doğrudan gün kutucuğuna tıklamayı yakala (altında ikinci bir buton olmadan tek kutucuk tıklaması!)
	try:
		clicked_wid = getattr(ui._get_ctx(), "clicked_button", None)
		if isinstance(clicked_wid, str) and clicked_wid.startswith("btn_cal_day_"):
			picked_iso = clicked_wid[len("btn_cal_day_"):][:10]
			date.fromisoformat(picked_iso)
			ui.session_state["cal_selected_date"] = picked_iso
	except Exception:
		pass

	_handle_calendar_event_inline_click(connect_db_fn, set_flash=True)

	if "cal_selected_date" not in ui.session_state:
		ui.session_state["cal_selected_date"] = today.isoformat()
	try:
		sel_date_obj = date.fromisoformat(str(ui.session_state["cal_selected_date"])[:10])
	except Exception:
		sel_date_obj = today
		ui.session_state["cal_selected_date"] = today.isoformat()

	if "cal_view_year" not in ui.session_state:
		ui.session_state["cal_view_year"] = sel_date_obj.year
	if "cal_view_month" not in ui.session_state:
		ui.session_state["cal_view_month"] = sel_date_obj.month

	view_year = int(ui.session_state.get("cal_view_year") or today.year)
	view_month = int(ui.session_state.get("cal_view_month") or today.month)

	# Eğer seçili tarih görüntülenen ayın dışındaysa senkronize et
	if sel_date_obj.year != view_year or sel_date_obj.month != view_month:
		_sync_selected_date_when_month_changes(view_year, view_month, today)
		sel_date_obj = date.fromisoformat(str(ui.session_state["cal_selected_date"])[:10])

	# Yılın temel resmi ve dini tatillerini ilk açılışta otomatik hazırla
	# (Silinen tatillerin sonraki oturumlarda geri gelmemesi için yıl kalıcı olarak işaretlenir)
	seeded_years = {y.strip() for y in str(cfg.get("_cal_auto_seeded_years") or "").split(",") if y.strip()}
	seed_flag_key = f"_cal_seeded_{view_year}"
	if (
		not ui.session_state.get(seed_flag_key)
		and str(view_year) not in seeded_years
		and str(cfg.get("calendar_auto_load_holidays", "1")) != "0"
	):
		ensure_year_builtin_holidays_seeded(view_year, connect_db_fn, include_special_days=False)
		ui.session_state[seed_flag_key] = True
		seeded_years.add(str(view_year))
		save_app_settings_fn({"_cal_auto_seeded_years": ",".join(sorted(seeded_years))})

	top_h1, top_h2 = ui.columns([2.7, 1.15])
	with top_h1:
		ui.markdown("## 📅 Okul & Revir Takvimi ve Günlük Takvim Yaprağı")
		ui.caption(
			"Gün kutucuklarına doğrudan tıklayarak sağdaki **Takvim Yaprağını** açın. "
			"Ezan vakti ili, görünüm ve yapay zeka ile toplu tatil/sınav/belirli gün ekleme işlemleri **Ayarlar → Takvim & Yemekhane → Takvim Ayarları** bölümündedir."
		)
	with top_h2:
		ui.write("")
		if ui.button("⚙️ Takvim Ayarları & AI Sihirbazı", key="btn_goto_cal_settings_top", type="secondary", use_container_width=True):
			ui.session_state["settings_vertical_nav"] = "📅 Takvim & Yemekhane"
			ui.session_state["_force_tab_label"] = "🗓️ Takvim Ayarları"
			ui.session_state["target_page"] = "⚙️ Ayarlar"
			ui.session_state["main_menu_page"] = "⚙️ Ayarlar"
			ui.rerun()

	flash_msg = ui.session_state.pop("_cal_flash_msg", None)
	if flash_msg:
		ui.success(flash_msg)

	# Ay Navigasyon Çubuğu
	nav_c1, nav_c2, nav_c3, nav_c4, nav_c5 = ui.columns([1.0, 1.35, 1.05, 1.0, 1.0])
	with nav_c1:
		if ui.button("◀ Önceki Ay", key="btn_cal_prev_month", use_container_width=True):
			if view_month == 1:
				view_month = 12
				view_year -= 1
			else:
				view_month -= 1
			ui.session_state["cal_view_year"] = view_year
			ui.session_state["cal_view_month"] = view_month
			ui.session_state["cal_sel_month_box"] = TR_MONTH_NAMES[view_month]
			ui.session_state["cal_sel_year_box"] = view_year
			_sync_selected_date_when_month_changes(view_year, view_month, today)
			ui.rerun()
	with nav_c2:
		month_labels = list(TR_MONTH_NAMES.values())
		cur_m_label = TR_MONTH_NAMES.get(view_month, "Ocak")
		picked_m_label = ui.selectbox(
			"Ay",
			month_labels,
			index=month_labels.index(cur_m_label),
			key="cal_sel_month_box",
			label_visibility="collapsed",
		)
		new_m = month_labels.index(picked_m_label) + 1
		if new_m != view_month:
			view_month = new_m
			ui.session_state["cal_view_month"] = view_month
			_sync_selected_date_when_month_changes(view_year, view_month, today)
			ui.rerun()
	with nav_c3:
		year_opts = list(range(2024, 2031))
		if view_year not in year_opts:
			year_opts.append(view_year)
			year_opts.sort()
		picked_y = int(
			ui.selectbox(
				"Yıl",
				year_opts,
				index=year_opts.index(view_year),
				key="cal_sel_year_box",
				label_visibility="collapsed",
			)
		)
		if picked_y != view_year:
			view_year = picked_y
			ui.session_state["cal_view_year"] = view_year
			_sync_selected_date_when_month_changes(view_year, view_month, today)
			ui.rerun()
	with nav_c4:
		if ui.button("📍 Bugüne Git", key="btn_cal_goto_today", type="secondary", use_container_width=True):
			ui.session_state["cal_view_year"] = today.year
			ui.session_state["cal_view_month"] = today.month
			ui.session_state["cal_selected_date"] = today.isoformat()
			ui.session_state["cal_sel_month_box"] = TR_MONTH_NAMES[today.month]
			ui.session_state["cal_sel_year_box"] = today.year
			ui.rerun()
	with nav_c5:
		if ui.button("Sonraki Ay ▶", key="btn_cal_next_month", use_container_width=True):
			if view_month == 12:
				view_month = 1
				view_year += 1
			else:
				view_month += 1
			ui.session_state["cal_view_year"] = view_year
			ui.session_state["cal_view_month"] = view_month
			ui.session_state["cal_sel_month_box"] = TR_MONTH_NAMES[view_month]
			ui.session_state["cal_sel_year_box"] = view_year
			_sync_selected_date_when_month_changes(view_year, view_month, today)
			ui.rerun()

	# -----------------------------------------------------------------------
	# GÖRÜNTÜLENEN AYIN VERİLERİNİ VERİTABANINDAN ÇEKME VE MESAİ ANALİZİ
	# -----------------------------------------------------------------------
	month_prefix = f"{view_year:04d}-{view_month:02d}-"
	with connect_db_fn() as conn:
		month_rows = conn.execute(
			"""SELECT id, event_date, title, category, symbol_type, is_civil_holiday,
			          is_half_day, profession_group, start_time, end_time, description,
			          source, is_completed
			   FROM calendar_events
			   WHERE event_date LIKE ?
			   ORDER BY is_civil_holiday DESC, is_half_day DESC, start_time ASC, id ASC""",
			(month_prefix + "%",),
		).fetchall()

		visit_counts_by_day: dict[str, int] = {}
		try:
			for vr in conn.execute(
				"SELECT visit_date, COUNT(*) AS cnt FROM visits WHERE visit_date LIKE ? GROUP BY visit_date",
				(month_prefix + "%",),
			).fetchall():
				visit_counts_by_day[str(vr["visit_date"])[:10]] = int(vr["cnt"] or 0)
		except Exception:
			pass

	events_by_date: dict[str, list[dict[str, Any]]] = {}
	for r in month_rows:
		d_key = str(r["event_date"])[:10]
		events_by_date.setdefault(d_key, []).append(dict(r))

	_, days_in_month = calendar.monthrange(view_year, view_month)
	workdays_count = 0.0
	weekend_days_count = 0
	civil_holiday_count = 0
	half_day_count = 0
	admin_leave_count = 0
	total_events_count = len(month_rows)

	for d_num in range(1, days_in_month + 1):
		d_obj = date(view_year, view_month, d_num)
		d_iso = d_obj.isoformat()
		day_evs = events_by_date.get(d_iso, [])
		is_wknd = d_obj.weekday() >= 5
		has_full_hol = any(bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")) for e in day_evs)
		has_half_hol = any(bool(e.get("is_half_day")) for e in day_evs)
		has_admin_leave = any(e.get("category") == "secim_idari_tatil" or e.get("symbol_type") == "ballot" for e in day_evs)

		if has_full_hol:
			civil_holiday_count += 1
		elif has_half_hol:
			half_day_count += 1
		elif has_admin_leave:
			admin_leave_count += 1

		if is_wknd:
			weekend_days_count += 1
		else:
			if has_full_hol or has_admin_leave:
				pass
			elif has_half_hol:
				workdays_count += 0.5
			else:
				workdays_count += 1.0

	workdays_str = f"{int(workdays_count)}" if workdays_count.is_integer() else f"{workdays_count:.1f}".replace(".", ",")
	admin_badge_str = f" · {admin_leave_count} idari izin" if admin_leave_count else ""

	# Ay Özeti & Renk / Sembol Rehberi Şeridi
	ui.markdown(
		f"""<div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;
		padding:10px 14px;border-radius:12px;background:var(--bg-card);border:1px solid var(--line);margin:6px 0 12px 0;">
			<div style="display:flex;flex-wrap:wrap;align-items:center;gap:12px;">
				<span style="font-size:15px;font-weight:800;color:var(--text-heading);">🗓️ {TR_MONTH_NAMES[view_month]} {view_year}</span>
				<span style="font-size:12.5px;background:#ecfdf5;color:#047857;border:1px solid #a7f3d0;padding:3px 9px;border-radius:999px;font-weight:700;">
					🏢 Net İş Günü: {workdays_str} gün
				</span>
				<span style="font-size:12.5px;background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;padding:3px 9px;border-radius:999px;font-weight:700;">
					🏖️ Hafta Sonu: {weekend_days_count} gün
				</span>
				<span style="font-size:12.5px;background:#fef2f2;color:#dc2626;border:1px solid #fecaca;padding:3px 9px;border-radius:999px;font-weight:800;">
					🇹🇷 Resmi Tatil: {civil_holiday_count} tam{f' · {half_day_count} yarım gün' if half_day_count else ''}{admin_badge_str}
				</span>
				<span style="font-size:12.5px;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;padding:3px 9px;border-radius:999px;font-weight:700;">
					📌 Kayıtlı Gün/Etkinlik: {total_events_count}
				</span>
			</div>
			<div style="display:flex;flex-wrap:wrap;align-items:center;gap:10px;font-size:11.5px;color:var(--text-muted);font-weight:600;">
				<span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:#ffffff;border:1px solid #94a3b8;vertical-align:-1px;"></span> İş Günü</span>
				<span><span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:#e2e8f0;border:1px solid #94a3b8;vertical-align:-1px;"></span> Hafta Sonu</span>
				<span style="color:#dc2626;font-weight:800;"><span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:#fee2e2;border:1px solid #dc2626;vertical-align:-1px;"></span> Memur Tatili (Kırmızı)</span>
				<span>{SVG_TR_FLAG_MINI} Milli Bayram</span>
				<span>{SVG_CRESCENT_MINI} Dini Bayram/Kandil</span>
				<span><b>½</b> Yarım Gün</span>
				<span>📝 Sınav</span>
			</div>
		</div>""",
		unsafe_allow_html=True,
	)

	show_grid_sec = True
	show_leaf_sec = True
	if is_section_visible_fn is not None:
		try:
			show_grid_sec = bool(is_section_visible_fn("calendar_monthly_grid"))
			show_leaf_sec = bool(is_section_visible_fn("calendar_daily_leaf"))
		except Exception:
			pass

	if not show_grid_sec or not show_leaf_sec:
		hidden_names = []
		if not show_grid_sec:
			hidden_names.append("Aylık Takvim Izgarası")
		if not show_leaf_sec:
			hidden_names.append("Günlük Takvim Yaprağı")
		hc1, hc2 = ui.columns([3.0, 1.2])
		with hc1:
			ui.info(f"👁️ Gizlenmiş takvim bölümü var: **{', '.join(hidden_names)}**.")
		with hc2:
			if ui.button("👁️ Gizli Takvim Bölümlerini Geri Aç", key="btn_restore_hidden_cal_sections", type="primary", use_container_width=True):
				try:
					vis_raw = cfg.get("page_sections_visibility", "{}")
					vis_map = json.loads(vis_raw) if vis_raw else {}
				except Exception:
					vis_map = {}
				vis_map["calendar_monthly_grid"] = True
				vis_map["calendar_daily_leaf"] = True
				save_app_settings_fn({"page_sections_visibility": json.dumps(vis_map, ensure_ascii=False)})
				ui.rerun()

	ui.markdown(
		"""<style>
		.revir-cal-split-anchor { display: none; }
		.revir-cal-split-anchor + .revir-cols {
			grid-template-columns: minmax(520px, 1.48fr) minmax(350px, 1fr) !important;
			align-items: start !important;
			gap: 18px !important;
		}
		.revir-cal-split-anchor + .revir-cols > .revir-col:nth-child(2) {
			position: sticky;
			top: 8px;
		}
		.revir-cal-grid-wrap {
			width: 100%;
			overflow-x: auto;
			padding-bottom: 2px;
		}
		.revir-cal-grid-wrap .revir-cols-7 {
			gap: 6px !important;
			margin-bottom: 0 !important;
		}
		@media (max-width: 1100px) {
			.revir-cal-split-anchor + .revir-cols {
				grid-template-columns: minmax(0, 1fr) !important;
			}
			.revir-cal-split-anchor + .revir-cols > .revir-col:nth-child(2) {
				position: static;
			}
		}
		</style>""",
		unsafe_allow_html=True,
	)

	if show_grid_sec and show_leaf_sec:
		ui.markdown('<div class="revir-cal-split-anchor"></div>', unsafe_allow_html=True)
		cal_col, leaf_col = ui.columns([1.48, 1.0], gap="medium")
	elif show_grid_sec:
		cal_col = ui.container()
		leaf_col = None
	elif show_leaf_sec:
		cal_col = None
		leaf_col = ui.container()
	else:
		return

	# -----------------------------------------------------------------------
	# SOL PANEL: 7 SÜTUNLU AYLIK İNTERAKTİF TAKVİM IZGARASI (TEK TIKLANABİLİR KUTUCUK)
	# -----------------------------------------------------------------------
	if cal_col is not None:
		with cal_col:
			ui.markdown('<div class="revir-cal-grid-wrap">', unsafe_allow_html=True)
			hdr_cols = ui.columns(7, gap="small")
			for w_idx, w_short in enumerate(TR_WEEKDAY_SHORT):
				is_wk = w_idx >= 5
				hdr_bg = "#fee2e2" if w_idx == 6 else ("#e2e8f0" if is_wk else "var(--bg-card)")
				hdr_fg = "#dc2626" if w_idx == 6 else ("#334155" if is_wk else "var(--text-heading)")
				with hdr_cols[w_idx]:
					ui.markdown(
						f"""<div style="text-align:center;padding:5px 3px;border-radius:8px;background:{hdr_bg};
						border:1px solid var(--line);font-size:11.5px;font-weight:800;color:{hdr_fg};margin-bottom:4px;">
							{w_short}
						</div>""",
						unsafe_allow_html=True,
					)

			cal_matrix = calendar.Calendar(firstweekday=0).monthdayscalendar(view_year, view_month)
			for week_idx, week_days in enumerate(cal_matrix):
				w_cols = ui.columns(7, gap="small")
				for wd_idx, day_num in enumerate(week_days):
					with w_cols[wd_idx]:
						if day_num == 0:
							ui.markdown(
								'<div style="min-height:82px;border-radius:9px;background:rgba(148,163,184,0.06);'
								'border:1px dashed rgba(148,163,184,0.22);margin-bottom:5px;"></div>',
								unsafe_allow_html=True,
							)
							continue

						cell_date = date(view_year, view_month, day_num)
						cell_iso = cell_date.isoformat()
						is_today = (cell_date == today)
						is_selected = (cell_date == sel_date_obj)
						is_weekend = (wd_idx >= 5)
						is_sunday = (wd_idx == 6)

						cell_events = events_by_date.get(cell_iso, [])
						sp_builtin = PROFESSION_AND_SPECIAL_DAYS.get((view_month, day_num))
						revir_v_cnt = visit_counts_by_day.get(cell_iso, 0)

						has_full_holiday = any(bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")) for e in cell_events)
						has_half_day = any(bool(e.get("is_half_day")) for e in cell_events)
						has_tr_flag = any(e.get("symbol_type") == "tr_flag" for e in cell_events) or (
							sp_builtin and sp_builtin.get("cat") == "milli_bayram"
						)
						has_crescent = any(e.get("symbol_type") == "crescent" for e in cell_events)
						has_exam = any(e.get("symbol_type") == "exam" or e.get("category") == "osym_sinav" for e in cell_events)
						has_ballot = any(e.get("symbol_type") == "ballot" or e.get("category") == "secim_idari_tatil" for e in cell_events)
						has_prof = any(e.get("symbol_type") == "profession" or e.get("category") in ("meslek_gunu", "belirli_gun") for e in cell_events) or bool(sp_builtin)

						if has_full_holiday:
							cell_bg = "linear-gradient(180deg, #fef2f2 0%, #fee2e2 100%)"
							cell_border = "2.5px solid #dc2626" if is_selected else "1.5px solid #fca5a5"
							num_color = "#dc2626"
							status_tag = '<span style="font-size:9px;font-weight:800;color:#ffffff;background:#dc2626;padding:1px 4px;border-radius:4px;">TATİL</span>'
						elif has_half_day and is_weekend:
							cell_bg = "linear-gradient(180deg, #fef2f2 0%, #fee2e2 100%)"
							cell_border = "2.5px solid #dc2626" if is_selected else "1.5px solid #fca5a5"
							num_color = "#dc2626"
							status_tag = '<span style="font-size:8.5px;font-weight:800;color:#b91c1c;background:#fecaca;padding:1px 4px;border-radius:4px;">ARİFE · TATİL</span>'
						elif has_half_day:
							cell_bg = "linear-gradient(180deg, #fff7ed 0%, #ffedd5 100%)"
							cell_border = "2.5px solid #ea580c" if is_selected else "1.5px solid #fdba74"
							num_color = "#dc2626"
							status_tag = '<span style="font-size:9px;font-weight:800;color:#9a3412;background:#fed7aa;padding:1px 4px;border-radius:4px;">½ YARIM GÜN</span>'
						elif has_ballot:
							cell_bg = "linear-gradient(180deg, #f5f3ff 0%, #ede9fe 100%)"
							cell_border = "2.5px solid #7c3aed" if is_selected else "1.5px solid #c4b5fd"
							num_color = "#6d28d9"
							status_tag = '<span style="font-size:9px;font-weight:800;color:#5b21b6;background:#ddd6fe;padding:1px 4px;border-radius:4px;">İDARİ İZİN</span>'
						elif is_weekend:
							cell_bg = "linear-gradient(180deg, #f1f5f9 0%, #e2e8f0 100%)"
							cell_border = "2.5px solid var(--teal)" if is_selected else "1px solid #cbd5e1"
							num_color = "#dc2626" if is_sunday else "#334155"
							status_tag = '<span style="font-size:8.8px;font-weight:700;color:#64748b;">H.Sonu</span>'
						else:
							cell_bg = "var(--bg-card)"
							cell_border = "2.5px solid var(--teal)" if is_selected else "1px solid var(--line)"
							num_color = "var(--text-heading)"
							status_tag = '<span style="font-size:8.8px;font-weight:600;color:#059669;">İş Günü</span>'

						if is_today and not is_selected:
							cell_border = "2px solid #0284c7"

						top_symbols: list[str] = []
						if has_tr_flag:
							top_symbols.append(SVG_TR_FLAG_MINI)
						if has_crescent:
							top_symbols.append(SVG_CRESCENT_MINI)
						if has_exam:
							top_symbols.append('<span title="ÖSYM / MEB Sınavı" style="font-size:11px;">📝</span>')
						if has_prof and not has_tr_flag:
							top_symbols.append('<span title="MEB Belirli Gün / Meslek Günü" style="font-size:11px;">👩‍⚕️</span>')
						symbols_html = "".join(top_symbols)

						shown_events = cell_events[:2]
						done_in_cell = sum(1 for e in cell_events if bool(e.get("is_completed")))
						done_hidden = sum(1 for e in cell_events[2:] if bool(e.get("is_completed")))

						pills_html_list: list[str] = []
						for ev in shown_events:
							ev_done = bool(ev.get("is_completed"))
							ev_title = html.escape(str(ev.get("title") or ""))
							ev_hol = bool(ev.get("is_civil_holiday"))
							ev_cat = str(ev.get("category") or "belirli_gun")
							if ev_done:
								p_bg, p_fg, p_bd = "#dcfce7", "#047857", "#34d399"
								title_done_css = "text-decoration:line-through;opacity:0.86;"
								pill_txt = f"✓ {ev_title}"
								ev_tooltip = f"{ev_title} (Tamamlandı)"
							else:
								title_done_css = ""
								pill_txt = ev_title
								ev_tooltip = ev_title
								if ev_hol:
									p_bg, p_fg, p_bd = "#fee2e2", "#b91c1c", "#fca5a5"
								elif ev_cat == "dini_bayram":
									p_bg, p_fg, p_bd = "#fef3c7", "#b45309", "#fde68a"
								elif ev_cat == "osym_sinav":
									p_bg, p_fg, p_bd = "#dbeafe", "#1d4ed8", "#93c5fd"
								elif ev_cat == "secim_idari_tatil":
									p_bg, p_fg, p_bd = "#ede9fe", "#6d28d9", "#c4b5fd"
								elif ev_cat == "revir_etkinlik":
									p_bg, p_fg, p_bd = "#d1fae5", "#047857", "#6ee7b7"
								else:
									p_bg, p_fg, p_bd = "#f1f5f9", "#334155", "#cbd5e1"
							pills_html_list.append(
								f'<div style="font-size:9.3px;font-weight:700;color:{p_fg};background:{p_bg};'
								f'border:1px solid {p_bd};border-radius:4px;padding:1px 4px;margin-top:2px;'
								f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;{title_done_css}" title="{ev_tooltip}">'
								f'{pill_txt}</div>'
							)

						if sp_builtin and len(pills_html_list) < 2:
							existing_titles_low = {str(e.get("title") or "").lower() for e in cell_events}
							if sp_builtin["title"].lower() not in existing_titles_low:
								sp_t = html.escape(sp_builtin["title"])
								pills_html_list.append(
									f'<div style="font-size:9px;font-weight:600;color:#0f766e;background:#f0fdfa;'
									f'border:1px solid #99f6e4;border-radius:4px;padding:1px 4px;margin-top:2px;'
									f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{sp_t}">'
									f'{sp_t}</div>'
								)

						if len(cell_events) > 2:
							extra_cnt = len(cell_events) - 2
							done_extra_note = f" · ✅ {done_hidden} bitti" if done_hidden > 0 else ""
							pills_html_list.append(
								f'<div style="font-size:9px;font-weight:800;color:#475569;margin-top:1px;">+{extra_cnt} kayıt{done_extra_note}</div>'
							)
						elif revir_v_cnt > 0:
							pills_html_list.append(
								f'<div style="font-size:9px;font-weight:700;color:#0369a1;margin-top:1px;">🩺 {revir_v_cnt} revir</div>'
							)
						elif done_in_cell > 0:
							pills_html_list.append(
								f'<div style="font-size:8.8px;font-weight:800;color:#047857;margin-top:1px;">✅ {done_in_cell} tamamlandı</div>'
							)

						pills_joined = "".join(pills_html_list)
						today_dot = (
							'<span style="background:#0284c7;color:#fff;font-size:8px;font-weight:800;padding:1px 4px;border-radius:4px;margin-left:2px;">BUGÜN</span>'
							if is_today else ""
						)
						sel_badge = (
							'<span style="background:var(--teal);color:#fff;font-size:8px;font-weight:800;padding:1px 4px;border-radius:4px;margin-left:2px;">SEÇİLİ</span>'
							if is_selected and not is_today else ""
						)
						sel_shadow = "0 0 0 2.5px rgba(13,148,136,0.3)" if is_selected else "0 1px 3px rgba(0,0,0,0.05)"

						# SADECE TEK KUTUCUK: Doğrudan bu kutucuğa tıklandığında o gün seçilir (altında ikinci buton yok!)
						ui.markdown(
							f"""<div role="button" tabindex="0"
							onclick="window.__revirClick('btn_cal_day_{cell_iso}')"
							onkeydown="if(event.key==='Enter'||event.key===' '){{event.preventDefault();window.__revirClick('btn_cal_day_{cell_iso}');}}"
							title="{day_num} {TR_MONTH_NAMES[view_month]} {view_year} takvim yaprağını açmak için tıklayın"
							style="min-height:82px;padding:5px 6px;border-radius:9px;background:{cell_bg};
							border:{cell_border};box-shadow:{sel_shadow};cursor:pointer;user-select:none;
							transition:transform 0.1s ease,box-shadow 0.1s ease;margin-bottom:5px;
							display:flex;flex-direction:column;justify-content:space-between;">
								<div>
									<div style="display:flex;align-items:center;justify-content:space-between;gap:2px;">
										<span style="font-size:14.5px;font-weight:900;color:{num_color};line-height:1.1;">{day_num}{today_dot}{sel_badge}</span>
										<span style="display:flex;align-items:center;gap:2px;">{symbols_html}</span>
									</div>
									<div style="margin-top:2px;">{status_tag}</div>
									<div style="margin-top:3px;">{pills_joined}</div>
								</div>
							</div>""",
							unsafe_allow_html=True,
						)
			ui.markdown('</div>', unsafe_allow_html=True)

			# Sol Panel Altı: Bu Ayın Etkinlik Özeti & Hızlı Kategori Filtresi
			with connect_db_fn() as _mconn:
				_ym_prefix = f"{view_year:04d}-{view_month:02d}-"
				_month_ev_rows = [
					dict(r)
					for r in _mconn.execute(
						"""SELECT id, event_date, title, category, is_civil_holiday, is_half_day, start_time, is_completed
						   FROM calendar_events WHERE event_date LIKE ?
						   ORDER BY event_date ASC, start_time ASC, id ASC""",
						(f"{_ym_prefix}%",),
					).fetchall()
				]
			with ui.expander(f"🔎 {TR_MONTH_NAMES[view_month]} {view_year} Etkinlik Özeti & Kategori Filtresi ({len(_month_ev_rows)})", expanded=False):
				_cat_f = ui.radio(
					"Kategori Filtresi",
					["🌐 Tümü", "🇹🇷 Resmi Tatil", "📝 Sınav", "🩺 Revir / Tarama", "🎗️ Belirli Gün"],
					horizontal=True,
					key=f"cal_month_cat_filter_{view_year}_{view_month}",
					label_visibility="collapsed",
				)
				_f_evs = []
				for _mev in _month_ev_rows:
					_mc = str(_mev.get("category") or "")
					if _cat_f == "🇹🇷 Resmi Tatil" and not (_mev.get("is_civil_holiday") or _mc in ("milli_bayram", "resmi_tatil", "yarim_gun", "secim_idari_tatil")):
						continue
					if _cat_f == "📝 Sınav" and "sinav" not in _mc:
						continue
					if _cat_f == "🩺 Revir / Tarama" and "revir" not in _mc:
						continue
					if _cat_f == "🎗️ Belirli Gün" and _mc not in ("meb_belirli_gun", "meslek_gunu"):
						continue
					_f_evs.append(_mev)
				if not _f_evs:
					ui.caption("Seçili kategoride bu ay için kayıtlı etkinlik bulunmuyor.")
				else:
					for _mev in _f_evs[:25]:
						_ed = str(_mev["event_date"])[:10]
						_is_mev_done = bool(_mev.get("is_completed"))
						_st_ico = "✅" if _is_mev_done else ("🔴" if _mev.get("is_civil_holiday") else "📌")
						_done_lbl_s = " · `(Tamamlandı)`" if _is_mev_done else ""
						_tm_s = f" · {_mev['start_time']}" if _mev.get("start_time") else ""
						_mc1, _mc2 = ui.columns([4.2, 1.0], vertical_alignment="center")
						with _mc1:
							ui.caption(f"{_st_ico} **{_ed}**{_tm_s} — {_mev['title']}{_done_lbl_s}")
						with _mc2:
							if ui.button("Güne Git", key=f"cal_jump_ev_{_mev['id']}", use_container_width=True):
								ui.session_state["cal_selected_date"] = _ed
								ui.rerun()

	# -----------------------------------------------------------------------
	# SAĞ PANEL: SEÇİLİ GÜNE AİT BİRLEŞİK "TAKVİM YAPRAĞI"
	# -----------------------------------------------------------------------
	if leaf_col is not None:
		with leaf_col:
			_render_interactive_calendar_leaf(
				sel_date=sel_date_obj,
				connect_db_fn=connect_db_fn,
				cfg=cfg,
				save_app_settings_fn=save_app_settings_fn,
				ai_enabled=ai_enabled,
				ai_key=ai_key,
				ai_model=ai_model,
				get_tr_provinces_fn=get_tr_provinces_fn,
			)


# ---------------------------------------------------------------------------
# SAĞ PANEL: BİRLEŞİK TAKVİM YAPRAĞI (EZAN VAKTİ + YAPRAK ARKASI YAZISI İÇERide)
# ---------------------------------------------------------------------------

def _render_interactive_calendar_leaf(
	*,
	sel_date: date,
	connect_db_fn: Callable,
	cfg: dict[str, str],
	save_app_settings_fn: Callable,
	ai_enabled: bool,
	ai_key: str,
	ai_model: str,
	get_tr_provinces_fn: Callable,
) -> None:
	today = date.today()
	sel_iso = sel_date.isoformat()
	weekday_idx = sel_date.weekday()
	weekday_name = TR_WEEKDAY_NAMES[weekday_idx]
	month_name = TR_MONTH_NAMES[sel_date.month]
	hijri_str = gregorian_to_hijri_str(sel_date)
	day_of_year = sel_date.timetuple().tm_yday
	days_in_year = 366 if calendar.isleap(sel_date.year) else 365
	remaining_days = days_in_year - day_of_year

	with connect_db_fn() as conn:
		day_events = [
			dict(r)
			for r in conn.execute(
				"""SELECT id, event_date, title, category, symbol_type, is_civil_holiday,
				          is_half_day, profession_group, start_time, end_time, description,
				          source, is_completed
				   FROM calendar_events
				   WHERE event_date = ?
				   ORDER BY is_civil_holiday DESC, start_time ASC, id ASC""",
				(sel_iso,),
			).fetchall()
		]
		leaf_note_row = conn.execute(
			"""SELECT quote_text, quote_source, hadith_text, hadith_source,
			          day_significance, profession_override, custom_note
			   FROM calendar_daily_notes
			   WHERE note_date = ?""",
			(sel_iso,),
		).fetchone()
		completed_all_events = [
			dict(r)
			for r in conn.execute(
				"""SELECT id, event_date, title, category, symbol_type, is_civil_holiday,
				          is_half_day, start_time, end_time, description, is_completed
				   FROM calendar_events
				   WHERE is_completed = 1
				   ORDER BY CASE WHEN event_date = ? THEN 0 ELSE 1 END,
				            event_date DESC, start_time ASC, id DESC
				   LIMIT 100""",
				(sel_iso,),
			).fetchall()
		]

	sp_builtin = PROFESSION_AND_SPECIAL_DAYS.get((sel_date.month, sel_date.day))
	active_meb_weeks = get_active_meb_weeks_for_date(sel_date)
	has_full_hol = any(bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")) for e in day_events)
	has_half_hol = any(bool(e.get("is_half_day")) for e in day_events)
	has_admin_leave = any(e.get("category") == "secim_idari_tatil" or e.get("symbol_type") == "ballot" for e in day_events)
	is_weekend = weekday_idx >= 5
	is_sunday = weekday_idx == 6

	if has_full_hol:
		header_accent = "#dc2626"
		num_color = "#dc2626"
		work_status_badge = "🇹🇷 RESMİ TATİL (DEVLET MEMURLARI İZİNLİ)"
		work_status_bg = "#fee2e2"
		work_status_fg = "#b91c1c"
	elif has_half_hol and is_weekend:
		header_accent = "#dc2626"
		num_color = "#dc2626"
		work_status_badge = "🏖️ HAFTA SONU & ARİFE GÜNÜ (MEMURLAR TAM GÜN İZİNLİ)"
		work_status_bg = "#fee2e2"
		work_status_fg = "#b91c1c"
	elif has_half_hol:
		header_accent = "#ea580c"
		num_color = "#dc2626"
		work_status_badge = "½ YARIM GÜN MESAİ (13:00'TEN SONRA RESMİ TATİL)"
		work_status_bg = "#ffedd5"
		work_status_fg = "#c2410c"
	elif has_admin_leave:
		header_accent = "#7c3aed"
		num_color = "#6d28d9"
		work_status_badge = "🗳️ SEÇİM TATİLİ / İDARİ İZİN (OKUL & KURUM TATİLİ)"
		work_status_bg = "#ede9fe"
		work_status_fg = "#5b21b6"
	elif is_weekend:
		header_accent = "#475569"
		num_color = "#dc2626" if is_sunday else "#334155"
		work_status_badge = "🏖️ HAFTA SONU TATİLİ"
		work_status_bg = "#e2e8f0"
		work_status_fg = "#334155"
	else:
		header_accent = "#0f766e"
		num_color = "#0f172a"
		work_status_badge = "🏢 TAM İŞ GÜNÜ (MEMUR & OKUL MESAİSİ)"
		work_status_bg = "#d1fae5"
		work_status_fg = "#065f46"

	# 1. Meslek Grubu Bölümü (Ayarlardan açık/kapalı kontrolü)
	show_profession = str(cfg.get("calendar_show_profession", "1")).strip() != "0"
	prof_section_html = ""
	if show_profession:
		prof_groups: list[str] = []
		for e in day_events:
			pg = str(e.get("profession_group") or "").strip()
			if pg and pg not in prof_groups:
				prof_groups.append(pg)
		if sp_builtin and sp_builtin.get("prof") and sp_builtin["prof"] not in prof_groups:
			prof_groups.append(sp_builtin["prof"])
		ai_prof_override = (leaf_note_row["profession_override"] if leaf_note_row and leaf_note_row["profession_override"] else "").strip()
		if ai_prof_override and ai_prof_override not in prof_groups:
			prof_groups.append(ai_prof_override)

		default_wd_prof, default_wd_desc = WEEKDAY_DEFAULT_PROFESSION_THEME[weekday_idx]
		ai_significance = (leaf_note_row["day_significance"] if leaf_note_row and leaf_note_row["day_significance"] else "").strip()

		if prof_groups:
			active_prof_title = " · ".join(prof_groups)
			active_prof_sub = (
				ai_significance
				or (sp_builtin["title"] if sp_builtin else "")
				or next((e["title"] for e in day_events if e.get("profession_group")), "Günün Özel Meslek Grubu")
			)
		else:
			active_prof_title = default_wd_prof
			active_prof_sub = ai_significance or default_wd_desc

		prof_section_html = f"""<div style="margin-top:10px;padding:8px 11px;border-radius:9px;background:#f0fdfa;border:1px solid #99f6e4;">
			<div style="font-size:10.5px;font-weight:800;color:#0f766e;text-transform:uppercase;letter-spacing:0.5px;">
				👩‍⚕️ GÜNÜN MESLEK GRUBU & ANLAMI
			</div>
			<div style="font-size:13px;font-weight:800;color:#134e4a;margin-top:2px;">
				{html.escape(active_prof_title)}
			</div>
			<div style="font-size:11.3px;color:#0f766e;margin-top:1px;">
				{html.escape(active_prof_sub)}
			</div>
		</div>"""

	# 2. MEB Belirli Gün ve Haftalar Şeridi (Ayarlardan açık/kapalı kontrolü)
	show_meb_weeks = str(cfg.get("calendar_show_meb_weeks", "1")).strip() != "0"
	meb_weeks_section_html = ""
	if show_meb_weeks and active_meb_weeks:
		meb_weeks_html_list: list[str] = []
		for mw in active_meb_weeks:
			meb_weeks_html_list.append(
				f"""<div style="padding:6px 10px;border-radius:8px;background:#eff6ff;border:1px solid #bfdbfe;margin-bottom:5px;">
					<div style="font-size:12px;font-weight:800;color:#1e40af;">🏫 {html.escape(mw['title'])}</div>
					<div style="font-size:11px;color:#1e3a8a;margin-top:1px;">Revir & Okul Odağı: {html.escape(mw['focus'])}</div>
				</div>"""
			)
		meb_weeks_section_html = (
			f'<div style="margin-top:9px;"><div style="font-size:10.5px;font-weight:800;color:#1e40af;'
			f'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">🏫 BU HAFTANIN MEB BELİRLİ GÜN VE HAFTALARI</div>'
			f'{"".join(meb_weeks_html_list)}</div>'
		)

	# 3. Günün Tatil / Sınav / Planlanan Etkinlik Özeti (Tamamlananlar yapraktan çıkar, alttaki 'Tamamlanan Etkinlikler'de gösterilir)
	active_day_events = [e for e in day_events if not bool(e.get("is_completed"))]
	completed_day_events = [e for e in day_events if bool(e.get("is_completed"))]
	existing_titles_low: set[str] = {str(e.get("title") or "").strip().lower() for e in day_events if str(e.get("title") or "").strip()}
	special_badges_html: list[str] = []
	for e in active_day_events:
		ev_id = int(e["id"])
		sym_html = render_symbol_badge(e.get("symbol_type") or "event")
		raw_ttl = str(e.get("title") or "").strip()
		e_ttl = html.escape(raw_ttl)
		e_desc = html.escape(str(e.get("description") or ""))
		t_rng = " - ".join(filter(None, [e.get("start_time"), e.get("end_time")]))
		t_span = f'<span style="background:rgba(15,23,42,0.08);padding:1px 6px;border-radius:4px;font-size:10.5px;font-weight:800;flex-shrink:0;">⏰ {html.escape(t_rng)}</span>' if t_rng else ""
		is_h = bool(e.get("is_civil_holiday"))
		b_bg = "#fef2f2" if is_h else "#f8fafc"
		b_bd = "#fca5a5" if is_h else "#cbd5e1"
		b_fg = "#b91c1c" if is_h else "#1e293b"
		special_badges_html.append(
			f"""<div class="revir-leaf-ev-card" style="background:{b_bg};border:1px solid {b_bd};color:{b_fg};">
				<button type="button" class="revir-leaf-ev-del" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_del_{ev_id}');" onkeydown="event.stopPropagation();" title="Bu etkinliği takvimden sil / kaldır">✕</button>
				<button type="button" role="checkbox" aria-checked="false" aria-label="✓ Tamamlandı" class="revir-leaf-ev-done" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_done_{ev_id}');" onkeydown="event.stopPropagation();" title="✓ Tamamlandı — Etkinliği tamamlandı olarak işaretle">✓</button>
				<div style="font-size:12.5px;font-weight:800;color:{b_fg};display:flex;align-items:center;justify-content:space-between;gap:6px;">
					<span style="display:flex;align-items:center;gap:6px;min-width:0;flex:1;">{sym_html} <span>{e_ttl}</span></span>
					{t_span}
				</div>
				{f'<div style="font-size:11.2px;color:#475569;margin-top:3px;">{e_desc}</div>' if e_desc else ''}
			</div>"""
		)
	if sp_builtin and sp_builtin["title"].lower() not in existing_titles_low:
		sym_html = render_symbol_badge("tr_flag" if sp_builtin.get("cat") == "milli_bayram" else "profession")
		special_badges_html.append(
			f"""<div style="padding:7px 10px;border-radius:8px;background:#f0fdfa;border:1px solid #99f6e4;margin-bottom:5px;">
				<div style="font-size:12.5px;font-weight:800;color:#0f766e;display:flex;align-items:center;gap:6px;">
					{sym_html} <span>{html.escape(sp_builtin['title'])}</span>
				</div>
			</div>"""
		)

	if completed_day_events and not active_day_events:
		done_note = '<div style="font-size:12px;color:#059669;font-style:italic;margin-top:4px;">✅ Bu güne ait tüm etkinlikler tamamlandı (aşağıdaki <b>Tamamlanan Etkinlikler</b> bölümünden inceleyebilirsiniz).</div>'
		special_block = "".join(special_badges_html) + done_note
	elif special_badges_html:
		special_block = "".join(special_badges_html)
	else:
		special_block = '<div style="font-size:12px;color:#64748b;font-style:italic;">Bu güne kayıtlı resmi tatil, sınav veya etkinlik bulunmuyor.</div>'

	# 4. TAKVİM YAPRAĞI İÇİNE GÖMÜLÜ EZAN VAKİTLERİ VE CANLI GERİ SAYIM SAYACI
	show_prayer = str(cfg.get("calendar_show_prayer_times", "1")).strip() != "0"
	default_prov = (cfg.get("default_province") or "Ankara").strip() or "Ankara"
	default_dist = (cfg.get("default_district") or "").strip()
	prayer_leaf_html = ""
	prayer_script_html = ""

	if show_prayer:
		loc_disp = f"{default_prov} / {default_dist}" if default_dist else default_prov
		sel_ptimes = get_prayer_times_for_province(default_prov, sel_date, connect_db_fn)
		today_ptimes = sel_ptimes if sel_date == today else get_prayer_times_for_province(default_prov, today, connect_db_fn)
		tmr_ptimes = get_prayer_times_for_province(default_prov, today + timedelta(days=1), connect_db_fn)

		vakit_keys = ["İmsak", "Güneş", "Öğle", "İkindi", "Akşam", "Yatsı"]
		now_dt = datetime.now()
		now_hm = now_dt.strftime("%H:%M")

		next_vakit_name = "İmsak"
		next_vakit_hm = today_ptimes.get("İmsak", "05:30")
		current_vakit_name = "Yatsı"
		is_target_tomorrow = False
		for idx_v, vk in enumerate(vakit_keys):
			vhm = today_ptimes.get(vk, "00:00")
			if now_hm < vhm:
				next_vakit_name = vk
				next_vakit_hm = vhm
				current_vakit_name = vakit_keys[idx_v - 1] if idx_v > 0 else "Yatsı"
				break
		else:
			is_target_tomorrow = True
			next_vakit_name = "İmsak (Yarın)"
			next_vakit_hm = tmr_ptimes.get("İmsak", today_ptimes.get("İmsak", "05:30"))
			current_vakit_name = "Yatsı"

		today_schedule_json = json.dumps(
			[{"name": vk, "time": today_ptimes.get(vk, "00:00")} for vk in vakit_keys],
			ensure_ascii=False,
		)
		tmr_imsak_json = json.dumps(tmr_ptimes.get("İmsak", today_ptimes.get("İmsak", "05:30")))

		v_cards_html: list[str] = []
		for vk in vakit_keys:
			vhm = sel_ptimes.get(vk, "--:--")
			is_cur_v = (vk == current_vakit_name and sel_date == today)
			vc_bg = "#0f766e" if is_cur_v else "#f8fafc"
			vc_fg = "#ffffff" if is_cur_v else "#0f172a"
			vc_sub = "#ccfbf1" if is_cur_v else "#64748b"
			vc_bd = "2px solid #0d9488" if is_cur_v else "1px solid #cbd5e1"
			v_cards_html.append(
				f"""<div data-vakit-card="{vk}" style="text-align:center;padding:5px 3px;border-radius:7px;background:{vc_bg};border:{vc_bd};">
					<div style="font-size:10px;font-weight:700;color:{vc_sub};">{vk}</div>
					<div style="font-size:13px;font-weight:900;color:{vc_fg};margin-top:1px;">{vhm}</div>
				</div>"""
			)

		try:
			nh, nm = [int(x) for x in next_vakit_hm.split(":")[:2]]
			target_dt = now_dt.replace(hour=nh, minute=nm, second=0, microsecond=0)
			if is_target_tomorrow or target_dt <= now_dt:
				target_dt += timedelta(days=1)
			diff_sec = max(0, int((target_dt - now_dt).total_seconds()))
			init_h = diff_sec // 3600
			init_m = (diff_sec % 3600) // 60
			init_s = diff_sec % 60
			init_countdown_txt = f"{init_h:02d} saat {init_m:02d} dk {init_s:02d} sn"
		except Exception:
			init_countdown_txt = "-- saat -- dk -- sn"

		diff_day_note = (
			f"Vakitler: {sel_date.strftime('%d.%m.%Y')} · Sayaç: Bugün ({today.strftime('%d.%m.%Y')})"
			if sel_date != today
			else f"{sel_ptimes.get('source', 'Diyanet Takvimi')}"
		)

		prayer_leaf_html = f"""<div style="margin-top:11px;padding-top:10px;border-top:1.5px dashed #cbd5e1;">
			<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
				<span style="font-size:10.8px;font-weight:800;color:#0f766e;text-transform:uppercase;letter-spacing:0.5px;">
					🕌 EZAN VAKİTLERİ ({html.escape(loc_disp)})
				</span>
				<span style="font-size:10px;color:#64748b;font-weight:600;">{html.escape(diff_day_note)}</span>
			</div>
			<div style="display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:4px;margin-bottom:7px;">
				{''.join(v_cards_html)}
			</div>
			<div style="padding:7px 11px;border-radius:8px;background:linear-gradient(90deg,#0f766e 0%,#0284c7 100%);
			color:#ffffff;display:flex;align-items:center;justify-content:space-between;font-size:12px;font-weight:800;">
				<span>⏳ <span id="revir_next_prayer_label">{html.escape(next_vakit_name)} ({html.escape(next_vakit_hm)})</span> vaktine kalan:</span>
				<span id="revir_prayer_countdown" style="background:rgba(255,255,255,0.2);padding:2px 8px;border-radius:5px;font-family:monospace;font-size:13px;letter-spacing:0.3px;">
					{init_countdown_txt}
				</span>
			</div>
		</div>"""

		prayer_script_html = f"""<script>
		(function() {{
			var schedule = {today_schedule_json};
			var tomorrowImsak = {tmr_imsak_json};
			function tickPrayerCountdown() {{
				var el = document.getElementById('revir_prayer_countdown');
				var lbl = document.getElementById('revir_next_prayer_label');
				if (!el || !schedule || !schedule.length) return;
				var now = new Date();
				var target = null;
				var targetName = '';
				var targetHM = '';
				for (var i = 0; i < schedule.length; i++) {{
					var parts = schedule[i].time.split(':');
					if (parts.length < 2) continue;
					var dt = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(parts[0], 10), parseInt(parts[1], 10), 0);
					if (dt > now) {{
						target = dt;
						targetName = schedule[i].name;
						targetHM = schedule[i].time;
						break;
					}}
				}}
				if (!target) {{
					var p0 = (tomorrowImsak || schedule[0].time).split(':');
					target = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, parseInt(p0[0], 10), parseInt(p0[1], 10), 0);
					targetName = 'İmsak (Yarın)';
					targetHM = tomorrowImsak || schedule[0].time;
				}}
				var diff = Math.max(0, Math.floor((target.getTime() - now.getTime()) / 1000));
				var hh = String(Math.floor(diff / 3600)).padStart(2, '0');
				var mm = String(Math.floor((diff % 3600) / 60)).padStart(2, '0');
				var ss = String(diff % 60).padStart(2, '0');
				el.textContent = hh + ' saat ' + mm + ' dk ' + ss + ' sn';
				if (lbl) lbl.textContent = targetName + ' (' + targetHM + ')';
			}}
			if (window._revirPrayerTimer) clearInterval(window._revirPrayerTimer);
			tickPrayerCountdown();
			window._revirPrayerTimer = setInterval(tickPrayerCountdown, 1000);
		}})();
		</script>"""

	# 5. TAKVİM YAPRAĞI İÇİNE GÖMÜLÜ "TAKVİM YAPRAĞI ARKASI" (HADİS-İ ŞERİF / ÖZLÜ SÖZ / GÜNÜN YAZISI)
	quote_mode = cfg.get("calendar_quote_mode", QUOTE_MODE_OPTIONS[0])
	if quote_mode not in QUOTE_MODE_OPTIONS:
		quote_mode = QUOTE_MODE_OPTIONS[0]

	quote_leaf_html = ""
	if quote_mode != "Kapalı (Gösterme)":
		builtin_qh = get_builtin_daily_quote_and_hadith(sel_date)
		hadith_txt = (leaf_note_row["hadith_text"] if leaf_note_row and leaf_note_row["hadith_text"] else "") or builtin_qh["hadith_text"]
		hadith_src = (leaf_note_row["hadith_source"] if leaf_note_row and leaf_note_row["hadith_source"] else "") or builtin_qh["hadith_source"]
		quote_txt = (leaf_note_row["quote_text"] if leaf_note_row and leaf_note_row["quote_text"] else "") or builtin_qh["quote_text"]
		quote_src = (leaf_note_row["quote_source"] if leaf_note_row and leaf_note_row["quote_source"] else "") or builtin_qh["quote_author"]
		leaf_back_txt = (leaf_note_row["custom_note"] if leaf_note_row and leaf_note_row["custom_note"] else "")

		show_hadith = ("Hadis" in quote_mode) or ("Karışık" in quote_mode)
		show_quote = ("Özlü Söz" in quote_mode) or ("Karışık" in quote_mode)

		back_parts: list[str] = []
		if show_hadith:
			back_parts.append(
				f"""<div style="padding:8px 11px;border-radius:8px;background:#fffbeb;border-left:3.5px solid #d97706;margin-bottom:6px;color:#78350f;">
					<div style="font-size:10px;font-weight:800;color:#b45309;text-transform:uppercase;">🌙 GÜNÜN HADİS-İ ŞERİFİ</div>
					<div style="font-size:12.2px;font-style:italic;font-weight:600;margin-top:2px;">“{html.escape(hadith_txt)}”</div>
					<div style="font-size:10.5px;font-weight:800;color:#92400e;margin-top:2px;text-align:right;">— {html.escape(hadith_src)}</div>
				</div>"""
			)
		if show_quote:
			back_parts.append(
				f"""<div style="padding:8px 11px;border-radius:8px;background:#f0fdf4;border-left:3.5px solid #059669;margin-bottom:6px;color:#064e3b;">
					<div style="font-size:10px;font-weight:800;color:#047857;text-transform:uppercase;">💡 GÜNÜN ÖZLÜ SÖZÜ / VECİZESİ</div>
					<div style="font-size:12.2px;font-style:italic;font-weight:600;margin-top:2px;">“{html.escape(quote_txt)}”</div>
					<div style="font-size:10.5px;font-weight:800;color:#065f46;margin-top:2px;text-align:right;">— {html.escape(quote_src)}</div>
				</div>"""
			)
		if leaf_back_txt:
			back_parts.append(
				f"""<div style="padding:7px 10px;border-radius:7px;background:#f8fafc;border:1px solid #cbd5e1;font-size:11.5px;color:#334155;">
					<b>📖 Takvim Yaprağı Arkası Yazısı:</b> {html.escape(leaf_back_txt)}
				</div>"""
			)
		if back_parts:
			quote_leaf_html = f"""<div style="margin-top:11px;padding-top:10px;border-top:1.5px dashed #cbd5e1;">
				<div style="font-size:10.5px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">
					📜 TAKVİM YAPRAĞI ARKASI · GÜNÜN SÖZÜ & YAZISI
				</div>
				{''.join(back_parts)}
			</div>"""

	ui.markdown(
		"""<style>
		.revir-leaf-ev-card {
			position: relative;
			min-height: 58px;
			padding: 8px 40px 8px 10px;
			border-radius: 9px;
			margin-bottom: 6px;
			transition: box-shadow 0.16s ease, border-color 0.16s ease;
		}
		.revir-leaf-ev-card:hover {
			box-shadow: 0 4px 12px rgba(15, 23, 42, 0.09);
		}
		.revir-leaf-ev-del {
			position: absolute;
			top: 6px;
			right: 7px;
			width: 22px;
			height: 22px;
			padding: 0;
			border-radius: 6px;
			display: inline-flex;
			align-items: center;
			justify-content: center;
			font-size: 12px;
			font-weight: 900;
			line-height: 1;
			cursor: pointer;
			background: rgba(255, 255, 255, 0.95);
			color: #dc2626;
			border: 1px solid #fca5a5;
			opacity: 0;
			pointer-events: none;
			transform: scale(0.9);
			transition: opacity 0.15s ease, transform 0.15s ease, background 0.15s ease, color 0.15s ease;
		}
		.revir-leaf-ev-del:hover {
			background: #dc2626;
			color: #ffffff;
			border-color: #b91c1c;
		}
		.revir-leaf-ev-done {
			position: absolute;
			top: 32px;
			right: 7px;
			width: 22px;
			height: 22px;
			padding: 0;
			border-radius: 6px;
			display: inline-flex;
			align-items: center;
			justify-content: center;
			font-size: 13px;
			font-weight: 900;
			line-height: 1;
			cursor: pointer;
			background: #ffffff;
			color: transparent;
			border: 2px solid #10b981;
			opacity: 0;
			pointer-events: none;
			transform: scale(0.9);
			transition: opacity 0.15s ease, transform 0.15s ease, background 0.15s ease, color 0.15s ease, border-color 0.15s ease;
		}
		.revir-leaf-ev-done:hover,
		.revir-leaf-ev-done:focus-visible {
			background: #10b981;
			color: #ffffff;
			border-color: #059669;
		}
		.revir-leaf-ev-card:hover .revir-leaf-ev-del,
		.revir-leaf-ev-card:focus-within .revir-leaf-ev-del {
			opacity: 1;
			pointer-events: auto;
			transform: scale(1);
		}
		.revir-leaf-ev-card:hover .revir-leaf-ev-done,
		.revir-leaf-ev-card:focus-within .revir-leaf-ev-done {
			opacity: 1;
			pointer-events: auto;
			transform: scale(1);
		}
		@media (hover: none), (max-width: 760px) {
			.revir-leaf-ev-card .revir-leaf-ev-del,
			.revir-leaf-ev-card .revir-leaf-ev-done {
				opacity: 1;
				pointer-events: auto;
				transform: none;
			}
		}
		.revir-done-ev-card {
			position: relative;
			padding: 8px 34px 28px 10px;
			border-radius: 9px;
			background: #ecfdf5;
			border: 1px solid #6ee7b7;
			margin-bottom: 6px;
			color: #064e3b;
		}
		.revir-done-ev-undo {
			position: absolute;
			bottom: 5px;
			right: 6px;
			height: 22px;
			padding: 0 8px;
			border-radius: 6px;
			display: inline-flex;
			align-items: center;
			justify-content: center;
			gap: 3px;
			font-size: 10.5px;
			font-weight: 800;
			line-height: 1;
			cursor: pointer;
			background: #ffffff;
			color: #047857;
			border: 1px solid #6ee7b7;
			opacity: 0.88;
			transition: opacity 0.15s ease, background 0.15s ease, color 0.15s ease;
		}
		.revir-done-ev-undo:hover {
			opacity: 1;
			background: #059669;
			color: #ffffff;
			border-color: #047857;
		}
		.revir-done-ev-card .revir-leaf-ev-del {
			opacity: 0.78;
			pointer-events: auto;
			transform: scale(1);
		}
		.revir-done-ev-card:hover .revir-leaf-ev-del {
			opacity: 1;
		}
		</style>""",
		unsafe_allow_html=True,
	)

	# BİRLEŞİK TEK PARÇA TAKVİM YAPRAĞI KARTI (Tarih + Meslek + MEB Haftası + Tatil + Ezan Vakti + Yaprak Arkası Yazısı)
	ui.markdown(
		f"""<div style="border-radius:13px;overflow:hidden;border:2px solid {header_accent};background:#fffdf9;
		box-shadow:0 8px 22px rgba(15,23,42,0.08);margin-bottom:8px;color:#1e293b;">
			<div style="background:{header_accent};color:#ffffff;padding:7px 13px;display:flex;align-items:center;justify-content:space-between;">
				<span style="font-size:11.5px;font-weight:800;letter-spacing:0.6px;">🗓️ GÜNLÜK TAKVİM YAPRAĞI</span>
				<span style="font-size:11px;font-weight:700;opacity:0.95;">🌙 {html.escape(hijri_str)}</span>
			</div>
			<div style="padding:10px 13px;">
				<div style="display:flex;align-items:center;justify-content:space-between;border-bottom:1.5px dashed #cbd5e1;padding-bottom:8px;">
					<div>
						<div style="font-size:12px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.8px;">
							{html.escape(month_name)} {sel_date.year}
						</div>
						<div style="display:flex;align-items:center;gap:9px;margin:2px 0;">
							<span style="font-size:40px;font-weight:900;color:{num_color};line-height:1.0;">
								{sel_date.day:02d}
							</span>
							<button type="button"
							onclick="event.stopPropagation();window.__revirClick('btn_open_cal_event_popup');"
							title="{sel_date.strftime('%d.%m.%Y')} — Etkinlik & Plan Yönetimi (Ekle / Düzenle)"
							style="width:30px;height:30px;border-radius:50%;border:1.5px solid {header_accent};
							background:{header_accent};color:#ffffff;font-size:19px;font-weight:900;
							line-height:1;display:inline-flex;align-items:center;justify-content:center;
							cursor:pointer;box-shadow:0 3px 10px rgba(15,23,42,0.16);">+</button>
						</div>
						<div style="font-size:14.5px;font-weight:800;color:{num_color};">
							{html.escape(weekday_name)}
						</div>
					</div>
					<div style="text-align:right;">
						<div style="display:inline-block;padding:3px 9px;border-radius:999px;background:{work_status_bg};color:{work_status_fg};font-size:10.2px;font-weight:900;margin-bottom:5px;">
							{work_status_badge}
						</div>
						<div style="font-size:11px;color:#475569;font-weight:700;">
							Yılın <b>{day_of_year}.</b> Günü · Kalan: <b>{remaining_days}</b>
						</div>
						<div style="font-size:10.5px;color:#64748b;margin-top:2px;">
							Miladi: {sel_date.strftime('%d.%m.%Y')}
						</div>
					</div>
				</div>

				{prof_section_html}
				{meb_weeks_section_html}

				<div style="margin-top:10px;">
					<div style="font-size:10.5px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;">
						📌 GÜNÜN TATİL, SINAV VE ETKİNLİKLERİ
					</div>
					{special_block}
				</div>

				{prayer_leaf_html}
				{quote_leaf_html}
			</div>
		</div>
		{prayer_script_html}""",
		unsafe_allow_html=True,
	)

	# Yaprak Altı Hızlı Aksiyonlar (Gemini ile Yaprak Yazısı Yenileme & Takvim Ayarları Kısayolu)
	leaf_act1, leaf_act2 = ui.columns(2)
	with leaf_act1:
		if ai_enabled and ai_key:
			if ui.button("✨ Gemini ile Yaprak Yazısını Yenile", key=f"btn_ai_leaf_{sel_iso}", type="secondary", use_container_width=True):
				meb_ctx = ", ".join([w["title"] for w in active_meb_weeks])
				day_ctx = ", ".join(filter(None, [
					", ".join([e["title"] for e in day_events]),
					sp_builtin["title"] if sp_builtin else "",
					meb_ctx,
				]))
				ok, res, used_m = generate_daily_leaf_content_with_ai(
					api_key=ai_key,
					model=ai_model,
					target_date=sel_iso,
					quote_mode=quote_mode,
					day_context=day_ctx,
					province=default_prov,
				)
				if ok:
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO calendar_daily_notes
							   (note_date, quote_text, quote_source, hadith_text, hadith_source,
							    day_significance, profession_override, custom_note, updated_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
							   ON CONFLICT(note_date) DO UPDATE SET
							     quote_text = excluded.quote_text,
							     quote_source = excluded.quote_source,
							     hadith_text = excluded.hadith_text,
							     hadith_source = excluded.hadith_source,
							     day_significance = excluded.day_significance,
							     profession_override = excluded.profession_override,
							     custom_note = excluded.custom_note,
							     updated_at = excluded.updated_at""",
							(
								sel_iso,
								str(res.get("quote_text") or "").strip(),
								str(res.get("quote_author") or "").strip(),
								str(res.get("hadith_text") or "").strip(),
								str(res.get("hadith_source") or "").strip(),
								str(res.get("day_significance") or "").strip(),
								str(res.get("profession_group") or "").strip(),
								str(res.get("leaf_back_note") or "").strip(),
								datetime.now().isoformat(timespec="seconds"),
							),
						)
					ui.session_state["_cal_flash_msg"] = f"✨ {sel_date.strftime('%d.%m.%Y')} takvim yaprağı yazısı Gemini AI ({used_m}) ile güncellendi."
					ui.rerun()
				else:
					ui.error(f"Gemini hatası: {res.get('error')}")
	with leaf_act2:
		if ui.button("⚙️ Takvim Ayarları (İl / Ezan / AI)", key="btn_goto_cal_settings_leaf", type="secondary", use_container_width=True):
			ui.session_state["settings_vertical_nav"] = "📅 Takvim & Yemekhane"
			ui.session_state["_force_tab_label"] = "🗓️ Takvim Ayarları"
			ui.session_state["target_page"] = "⚙️ Ayarlar"
			ui.session_state["main_menu_page"] = "⚙️ Ayarlar"
			ui.rerun()

	# Tamamlanan Etkinlikler Bölümü (Yapraktan kaldırılan ama aylık takvimde görünmeye devam eden etkinlikler)
	_expand_done_box = bool(ui.session_state.pop("_cal_expand_completed", False))
	with ui.expander(f"✅ Tamamlanan Etkinlikler ({len(completed_all_events)})", expanded=_expand_done_box):
		if not completed_all_events:
			ui.caption("Henüz tamamlanmış etkinlik yok. Yapraktaki etkinlik kutucuğuna tıklayarak tamamladığınız kayıtlar burada listelenir.")
		else:
			ui.caption("Aylık takvimde görünmeye devam eder; **↩️ Geri Al** ile yaprağa döndürebilir veya **✕** ile silebilirsiniz.")
			for ce in completed_all_events:
				ce_id = int(ce["id"])
				ce_d_str = str(ce.get("event_date") or "")[:10]
				try:
					ce_d_obj = date.fromisoformat(ce_d_str)
					ce_m_name = TR_MONTH_NAMES.get(ce_d_obj.month, str(ce_d_obj.month))
					ce_date_lbl = f"{ce_d_obj.day:02d} {ce_m_name} {ce_d_obj.year}"
				except Exception:
					ce_date_lbl = ce_d_str
				ce_sel_badge = (
					"<span style='background:#0d9488;color:#ffffff;padding:1px 5px;border-radius:4px;font-size:9.5px;font-weight:800;'>SEÇİLİ GÜN</span>"
					if ce_d_str == sel_iso
					else ""
				)
				ce_sym_html = render_symbol_badge(ce.get("symbol_type") or "event")
				ce_ttl = html.escape(str(ce.get("title") or "").strip())
				ce_desc = html.escape(str(ce.get("description") or "").strip())
				ce_t_rng = " - ".join(filter(None, [ce.get("start_time"), ce.get("end_time")]))
				ce_t_span = (
					f'<span style="background:rgba(6,78,59,0.1);padding:1px 6px;border-radius:4px;font-size:10.5px;font-weight:800;">⏰ {html.escape(ce_t_rng)}</span>'
					if ce_t_rng
					else ""
				)
				ui.markdown(
					f"""<div class="revir-done-ev-card">
						<button type="button" class="revir-leaf-ev-del" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_del_{ce_id}');" onkeydown="event.stopPropagation();" title="Bu etkinliği takvimden kalıcı olarak sil">✕</button>
						<div style="display:flex;align-items:center;justify-content:space-between;gap:6px;flex-wrap:wrap;">
							<span style="display:inline-flex;align-items:center;gap:6px;min-width:0;flex:1;flex-wrap:wrap;">
								<span style="background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;padding:1px 6px;border-radius:5px;font-size:10.5px;font-weight:800;">📅 {html.escape(ce_date_lbl)}</span>
								{ce_sel_badge}
								{ce_sym_html}
								<span style="font-size:12.5px;font-weight:800;text-decoration:line-through;opacity:0.88;">{ce_ttl}</span>
							</span>
							{ce_t_span}
						</div>
						{f'<div style="font-size:11px;color:#047857;margin-top:3px;">{ce_desc}</div>' if ce_desc else ''}
						<button type="button" class="revir-done-ev-undo" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_done_{ce_id}');" onkeydown="event.stopPropagation();" title="Tamamlandı işaretini geri al ve tekrar Takvim Yaprağı'na koy">↩️ Geri Al</button>
					</div>""",
					unsafe_allow_html=True,
				)

	# -----------------------------------------------------------------------
	# ETKİNLİK & PLAN YÖNETİMİ POP-UP DİYALOG TETİKLEYİCİSİ (+ BUTONU)
	# -----------------------------------------------------------------------
	try:
		clicked_wid = getattr(ui._get_ctx(), "clicked_button", None)
		if clicked_wid == "btn_open_cal_event_popup":
			open_calendar_event_manager_dialog(
				sel_iso=sel_iso,
				connect_db_fn=connect_db_fn,
				get_active_gemini_config_fn=lambda: (ai_enabled, ai_key, ai_model),
			)
	except Exception:
		pass


@ui.dialog("📌 Etkinlik & Plan Yönetimi", width="large")
def open_calendar_event_manager_dialog(
	*,
	sel_iso: str,
	connect_db_fn: Callable,
	get_active_gemini_config_fn: Callable,
) -> None:
	"""Günlük Takvim Yaprağı üzerindeki '+' butonuna basıldığında açılan Etkinlik & Plan Yönetimi pop-up penceresi."""
	try:
		sel_date = date.fromisoformat(str(sel_iso)[:10])
	except Exception:
		sel_date = date.today()
		sel_iso = sel_date.isoformat()

	month_name = TR_MONTH_NAMES[sel_date.month]
	weekday_name = TR_WEEKDAY_NAMES[sel_date.weekday()]
	ai_enabled, ai_key, ai_model = get_active_gemini_config_fn()

	dlg_flash = ui.session_state.pop("_cal_dlg_flash", None)
	if dlg_flash:
		ui.success(dlg_flash)

	ui.markdown(f"### 📅 {sel_date.strftime('%d')} {month_name} {sel_date.year} ({weekday_name}) — Etkinlik & Plan Yönetimi")

	with connect_db_fn() as conn:
		day_events = [
			dict(r)
			for r in conn.execute(
				"""SELECT id, event_date, title, category, symbol_type, is_civil_holiday,
				          is_half_day, profession_group, start_time, end_time, description,
				          source, is_completed
				   FROM calendar_events
				   WHERE event_date = ?
				   ORDER BY is_civil_holiday DESC, start_time ASC, id ASC""",
				(sel_iso,),
			).fetchall()
		]

	if day_events:
		with ui.container(border=True):
			done_cnt_dlg = sum(1 for e in day_events if bool(e.get("is_completed")))
			active_cnt_dlg = len(day_events) - done_cnt_dlg
			ui.markdown("#### 📋 Bu Güne Kayıtlı Etkinlikler / Tatiller")
			if done_cnt_dlg:
				ui.caption(f"⏳ Aktif: **{active_cnt_dlg}** · ✅ Tamamlandı: **{done_cnt_dlg}**")
			for ev in day_events:
				ev_id = int(ev["id"])
				ev_done = bool(ev.get("is_completed"))
				t_range = " - ".join(filter(None, [ev.get("start_time"), ev.get("end_time")]))
				t_badge = f"<span style='background:rgba(15,23,42,0.07);padding:1px 6px;border-radius:4px;font-size:11px;font-weight:800;'>⏰ {html.escape(t_range)}</span>" if t_range else ""
				done_pill = (
					"<span style='display:inline-block;background:#dcfce7;color:#166534;border:1px solid #86efac;"
					"padding:1px 8px;border-radius:999px;font-size:11px;font-weight:800;'>✅ Tamamlandı</span>"
					if ev_done
					else "<span style='display:inline-block;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;"
					"padding:1px 7px;border-radius:999px;font-size:10.5px;font-weight:700;'>⏳ Planlandı</span>"
				)
				ttl_html = html.escape(str(ev.get("title") or ""))
				ttl_decor = "text-decoration:line-through;opacity:0.78;color:#065f46;" if ev_done else "color:var(--text-heading);"
				card_style = "background:#f0fdf4;border:1px solid #86efac;" if ev_done else "background:var(--bg-card);border:1px solid var(--line);"
				desc_html = f"<div style='font-size:12px;color:#475569;margin-top:3px;'>{html.escape(str(ev.get('description') or ''))}</div>" if ev.get("description") else ""
				row_c1, row_c2, row_c3 = ui.columns([2.45, 0.95, 0.45], vertical_alignment="center")
				with row_c1:
					ui.markdown(
						f"""<div style="padding:7px 10px;border-radius:8px;{card_style}">
							<div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px;">
								{t_badge}
								<span style="font-size:13.5px;font-weight:800;{ttl_decor}">{ttl_html}</span>
								{done_pill}
							</div>
							{desc_html}
						</div>""",
						unsafe_allow_html=True,
					)
				with row_c2:
					done_lbl = "↩️ Geri Al" if ev_done else "✓ Tamamlandı"
					if ui.button(done_lbl, key=f"dlg_btn_cal_done_{ev_id}", help="Tamamlandı durumunu değiştir", use_container_width=True):
						with connect_db_fn() as conn:
							conn.execute("UPDATE calendar_events SET is_completed = ? WHERE id = ?", (0 if ev_done else 1, ev_id))
						ui.rerun()
				with row_c3:
					if ui.button("✕", key=f"dlg_btn_cal_del_{ev_id}", help="Bu etkinliği takvimden kaldır", use_container_width=True):
						with connect_db_fn() as conn:
							conn.execute("DELETE FROM calendar_events WHERE id = ?", (ev_id,))
						ui.session_state["_cal_dlg_flash"] = f"✕ '{ev['title']}' takvimden kaldırıldı."
						ui.rerun()
	else:
		ui.info("Bu güne henüz kayıtlı bir etkinlik, tatil veya sınav bulunmuyor. Aşağıdan hemen ekleyebilirsiniz.")

	with ui.container(border=True):
		ui.markdown(f"#### ➕ {sel_date.strftime('%d.%m.%Y')} Gününe Yeni Etkinlik / Tatil / Sınav Ekle")
		if ai_enabled and ai_key:
			q_ai_txt = ui.text_input(
				"✍️ Yapay Zekâya (Gemini AI) Yazarak Hızlı Etkinlik / Gün / Saat / Durum Ekle",
				placeholder=f"Örn: '{sel_date.strftime('%d.%m.%Y')} saat 10:30 4-A sınıfı göz taraması' veya 'Cuma günü idari izin'...",
				key=f"dlg_cal_quick_ai_txt_{sel_iso}",
			)
			if ui.button("✨ Yazılanı Gemini AI ile Takvime İşle", key=f"dlg_btn_quick_ai_add_{sel_iso}", type="secondary", use_container_width=True):
				if not q_ai_txt.strip():
					ui.warning("Lütfen eklenecek etkinlik, gün, saat veya durumu yazın.")
				else:
					ok, res, used_m = generate_calendar_events_with_ai(
						api_key=ai_key,
						model=ai_model,
						year=sel_date.year,
						month=sel_date.month,
						preset_type="election_or_custom",
						custom_prompt=f"(Seçili gün bağlamı: {sel_iso}) {q_ai_txt.strip()}",
					)
					if ok:
						ev_items = res.get("events") or []
						added = save_ai_generated_events_to_db(ev_items, connect_db_fn, upsert=True)
						if ev_items and isinstance(ev_items[0], dict) and ev_items[0].get("event_date"):
							try:
								d0 = date.fromisoformat(str(ev_items[0]["event_date"])[:10])
								ui.session_state["cal_selected_date"] = d0.isoformat()
								ui.session_state["cal_view_year"] = d0.year
								ui.session_state["cal_view_month"] = d0.month
							except Exception:
								pass
						ui.session_state["_cal_dlg_flash"] = f"✨ Gemini AI ({used_m}): {res.get('summary') or ''} ({added} kayıt takvime eklendi)."
						ui.rerun()
					else:
						ui.error(f"Gemini hatası: {res.get('error')}")
			ui.divider()

		cat_keys = list(CATEGORY_LABELS.keys())
		cat_disp = [CATEGORY_LABELS[k][0] for k in cat_keys]
		sym_keys = list(SYMBOL_OPTIONS.keys())
		sym_disp = [SYMBOL_OPTIONS[k] for k in sym_keys]

		picked_cat_disp = ui.selectbox(
			"Kategori Seçin",
			cat_disp,
			index=cat_keys.index("revir_etkinlik"),
			key=f"dlg_cal_add_cat_live_{sel_iso}",
		)
		picked_cat = cat_keys[cat_disp.index(picked_cat_disp)]
		default_sym = CATEGORY_LABELS[picked_cat][2]
		default_hol = picked_cat in ("milli_bayram", "resmi_tatil", "yarim_gun", "secim_idari_tatil")
		default_half = picked_cat == "yarim_gun"

		with ui.form(f"dlg_form_add_cal_event_{sel_iso}_{picked_cat}"):
			new_title = ui.text_input("Etkinlik / Tatil / Sınav Başlığı", placeholder="Örn: 3-A Sınıfı Boy-Kilo Taraması veya Seçim İdari İzni")
			picked_sym_disp = ui.selectbox("Takvim Sembolü", sym_disp, index=sym_keys.index(default_sym))
			picked_sym = sym_keys[sym_disp.index(picked_sym_disp)]

			tc1, tc2 = ui.columns(2)
			with tc1:
				new_start = ui.text_input("Başlangıç Saati", placeholder="09:00")
			with tc2:
				new_end = ui.text_input("Bitiş Saati", placeholder="12:30")

			hc1, hc2 = ui.columns(2)
			with hc1:
				new_is_hol = ui.checkbox("Devlet Memuru Resmi Tatili (Kırmızı Yaz)", value=default_hol)
			with hc2:
				new_is_half = ui.checkbox("Yarım Gün Mesai (½ Rozeti)", value=default_half)

			new_prof = ui.text_input("İlgili Meslek Grubu (Opsiyonel)", placeholder="Örn: Okul Hemşireleri ve Öğretmenler")
			new_desc = ui.text_area("Açıklama / Not", placeholder="Günün anlamı veya planlanan etkinlik detayı...", height=60)

			if ui.form_submit_button("💾 Takvime Kaydet", type="primary", use_container_width=True):
				if not new_title.strip():
					ui.warning("Lütfen bir başlık girin.")
				else:
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO calendar_events
							   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
							    profession_group, start_time, end_time, description, source, is_completed, created_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'manual', 0, ?)""",
							(
								sel_iso,
								new_title.strip(),
								picked_cat,
								picked_sym,
								1 if (new_is_hol or new_is_half) else 0,
								1 if new_is_half else 0,
								new_prof.strip(),
								new_start.strip()[:5],
								new_end.strip()[:5],
								new_desc.strip(),
								datetime.now().isoformat(timespec="seconds"),
							),
						)
					ui.session_state["_cal_flash_msg"] = f"✅ '{new_title.strip()}' ({sel_date.strftime('%d.%m.%Y')}) takvime eklendi."
					ui.close_dialog()
					ui.rerun()


# ---------------------------------------------------------------------------
# AYARLAR -> KURUM, YIL & ŞUBE YAPISI -> 🗓️ TAKVİM AYARLARI SEKMESİ
# ---------------------------------------------------------------------------

def render_calendar_settings_tab(
	*,
	connect_db_fn: Callable,
	current_cfg: dict[str, str],
	save_app_settings_fn: Callable,
	get_active_gemini_config_fn: Callable,
	get_tr_provinces_fn: Callable,
) -> None:
	"""
	Ayarlar -> 🏫 Kurum, Yıl & Şube Yapısı -> 🗓️ Takvim Ayarları sekmesini oluşturur:
	1. Takvim Yaprağı görünüm bölümleri, Ezan Vakitleri görünürlüğü, İl değiştirme ve Hadis/Özlü Söz tercihi
	2. Yapay Zeka (Gemini AI) ve Hazır Paketlerle Resmi/Dini Tatil, Yarım Gün, MEB Belirli Gün ve Haftalar, ÖSYM/MEB Sınavı ekleme
	3. Ekstra eklemek istediğimiz etkinlikleri, tatilleri veya özel durumları (günü, saati ile) yazarak ekleyebileceğimiz kutucuk
	"""
	today = date.today()
	ai_enabled, ai_key, ai_model = get_active_gemini_config_fn()
	ai_provider_lbl = "Gemini AI"
	all_provs = get_tr_provinces_fn() or sorted(TR_PROVINCE_COORDS.keys())
	cur_prov = (current_cfg.get("default_province") or "Ankara").strip() or "Ankara"
	if cur_prov not in all_provs:
		cur_prov = "Ankara"

	ui.markdown("### 🗓️ Takvim Ayarları")
	ui.caption(
		"**Takvim & Planlama** yaprağı bölümlerini, ezan vakti ilini, günün sözünü yönetin veya yapay zekâ / hazır paketlerle tatil, sınav ve etkinlik ekleyin."
	)

	cal_cfg_flash = ui.session_state.pop("_cal_settings_flash", None)
	if cal_cfg_flash:
		ui.success(cal_cfg_flash)

	# -----------------------------------------------------------------------
	# 1. BÖLÜM: TAKVİM YAPRAĞI GÖRÜNÜM, EZAN VAKTİ VE İL AYARLARI
	# -----------------------------------------------------------------------
	with ui.container(border=True):
		ui.markdown("#### 🕌 Yaprak İçeriği, Ezan Vakti İli ve Sözler")
		s_c1, s_c2 = ui.columns(2)
		with s_c1:
			sel_prov = ui.selectbox(
				"🏙️ Ezan Vakitleri İli (Varsayılan İl ile Senkronize)",
				all_provs,
				index=all_provs.index(cur_prov),
				key="cfg_cal_tab_province",
			)
		cur_q_mode = current_cfg.get("calendar_quote_mode", QUOTE_MODE_OPTIONS[0])
		if cur_q_mode not in QUOTE_MODE_OPTIONS:
			cur_q_mode = QUOTE_MODE_OPTIONS[0]
		with s_c2:
			sel_q_mode = ui.selectbox(
				"📜 Takvim Yaprağı Arkası Yazı Tercihi (Hadis / Özlü Söz)",
				QUOTE_MODE_OPTIONS,
				index=QUOTE_MODE_OPTIONS.index(cur_q_mode),
				key="cfg_cal_tab_quote_mode",
			)

		tg_c1, tg_c2, tg_c3, tg_c4 = ui.columns(4)
		with tg_c1:
			sel_show_prayer = ui.toggle(
				"🕌 Ezan Vakitleri ve Canlı Sayacı Göster",
				value=(str(current_cfg.get("calendar_show_prayer_times", "1")) != "0"),
				key="cfg_cal_tab_show_prayer",
			)
		with tg_c2:
			sel_show_prof = ui.toggle(
				"👩‍⚕️ Günün Meslek Grubu Bölümünü Göster",
				value=(str(current_cfg.get("calendar_show_profession", "1")) != "0"),
				key="cfg_cal_tab_show_prof",
			)
		with tg_c3:
			sel_show_meb = ui.toggle(
				"🏫 MEB Belirli Gün & Haftalar Şeridini Göster",
				value=(str(current_cfg.get("calendar_show_meb_weeks", "1")) != "0"),
				key="cfg_cal_tab_show_meb",
			)
		with tg_c4:
			sel_auto_hol = ui.toggle(
				"🇹🇷 Resmi & Dini Tatilleri Otomatik İşaretle",
				value=(str(current_cfg.get("calendar_auto_load_holidays", "1")) != "0"),
				key="cfg_cal_tab_auto_hol",
			)

		if ui.button("💾 Takvim Yaprağı, İl ve Görünüm Ayarlarını Kaydet", type="primary", key="btn_save_cal_tab_settings"):
			save_app_settings_fn({
				"default_province": sel_prov,
				"calendar_quote_mode": sel_q_mode,
				"calendar_show_prayer_times": "1" if sel_show_prayer else "0",
				"calendar_show_profession": "1" if sel_show_prof else "0",
				"calendar_show_meb_weeks": "1" if sel_show_meb else "0",
				"calendar_auto_load_holidays": "1" if sel_auto_hol else "0",
			})
			ui.session_state["_cal_settings_flash"] = f"✅ Takvim yaprağı ayarları kaydedildi (Ezan vakti ili: {sel_prov}, Söz tercihi: {sel_q_mode})."
			ui.rerun()

	# -----------------------------------------------------------------------
	# 2. BÖLÜM: YAPAY ZEKA & HAZIR PAKETLERLE RESMİ TATİL, MEB BELİRLİ GÜNLER VE ÖSYM SINAVLARI EKLEME
	# -----------------------------------------------------------------------
	with ui.container(border=True):
		ui.markdown("#### ✨ Otomatik Tatil, MEB Günleri & Sınav Ekleme")
		ui.caption("Hedef yıl/ayı seçerek resmi tatilleri, yarım gün arifeleri, MEB Belirli Günlerini ve ÖSYM/MEB sınavlarını takvime işleyin.")

		ym_c1, ym_c2 = ui.columns(2)
		year_opts = list(range(2024, 2031))
		def_yr = int(ui.session_state.get("cal_view_year") or today.year)
		if def_yr not in year_opts:
			def_yr = today.year
		with ym_c1:
			target_year = int(ui.selectbox("📅 Hedef Yıl", year_opts, index=year_opts.index(def_yr), key="cfg_cal_target_year"))
		month_labels = list(TR_MONTH_NAMES.values())
		def_mn = int(ui.session_state.get("cal_view_month") or today.month)
		with ym_c2:
			picked_m_lbl = ui.selectbox("🗓️ Hedef Ay (Ay Bazlı İşlemler İçin)", month_labels, index=def_mn - 1, key="cfg_cal_target_month")
			target_month = month_labels.index(picked_m_lbl) + 1

		ui.markdown("##### 🇹🇷 A) Resmi Tatiller, Arifeler ve MEB Belirli Günler")
		p_c1, p_c2, p_c3 = ui.columns(3)
		with p_c1:
			if ui.button(f"📥 {target_year} Resmi & Dini Tatiller + Yarım Günleri İşle", key="cfg_btn_seed_holidays", type="primary", use_container_width=True):
				cnt = ensure_year_builtin_holidays_seeded(target_year, connect_db_fn, include_special_days=False)
				ui.session_state["_cal_settings_flash"] = f"✅ {target_year} yılı resmi tatilleri, yarım gün arifeler ve dini bayramlar takvime işlendi ({cnt} yeni kayıt)."
				ui.rerun()
		with p_c2:
			if ui.button(f"🏫 {target_year} MEB Belirli Gün ve Haftalar + Meslek Günlerini İşle", key="cfg_btn_seed_meb", type="primary", use_container_width=True):
				cnt = ensure_year_builtin_holidays_seeded(target_year, connect_db_fn, include_special_days=True)
				ui.session_state["_cal_settings_flash"] = f"✅ {target_year} yılı MEB Belirli Gün ve Haftalar çizelgesi ile meslek günleri takvime işlendi ({cnt} yeni kayıt)."
				ui.rerun()
		with p_c3:
			if ai_enabled and ai_key:
				if ui.button(f"✨ {ai_provider_lbl} ile {TR_MONTH_NAMES[target_month]} {target_year} MEB Günlerini Ekle", key="cfg_btn_ai_meb_month", type="secondary", use_container_width=True):
					ok, res, used_m = generate_calendar_events_with_ai(
						api_key=ai_key,
						model=ai_model,
						year=target_year,
						month=target_month,
						preset_type="meb_belirli_gunler",
					)
					if ok:
						added = save_ai_generated_events_to_db(res.get("events") or [], connect_db_fn)
						ui.session_state["_cal_settings_flash"] = f"✨ {ai_provider_lbl} ({used_m}) ile {TR_MONTH_NAMES[target_month]} {target_year} için {added} MEB Belirli Gün ve Haftası eklendi!"
						ui.rerun()
					else:
						ui.error(f"AI hatası: {res.get('error')}")
			else:
				ui.caption("💡 Yapay zekâ açık olduğunda ay/yıl bazlı ek MEB belirli gün taraması yapabilirsiniz.")

		ui.divider()
		ui.markdown("##### 📝 B) ÖSYM ve MEB Merkezi Sınavları")
		exam_choices = ["YKS (TYT-AYT-YDT)", "LGS", "KPSS", "MSÜ", "ALES", "DGS", "YDS / YÖKDİL", "İOKBS Bursluluk"]
		picked_exams = ui.multiselect(
			"Takvimde İşaretlenecek ÖSYM / MEB Sınavlarını Seçin",
			exam_choices,
			default=["YKS (TYT-AYT-YDT)", "LGS", "KPSS", "MSÜ"],
			key="cfg_cal_multiselect_exams",
		)
		ex_b1, ex_b2 = ui.columns(2)
		with ex_b1:
			if ui.button(f"📝 Seçili Sınavları {target_year} Takvimine İşle (Hızlı)", key="cfg_btn_seed_exams", type="primary", use_container_width=True):
				cnt = seed_selected_exams_to_db(target_year, picked_exams, connect_db_fn)
				ui.session_state["_cal_settings_flash"] = f"✅ Seçilen sınavlardan {cnt} oturum {target_year} takvimine 📝 sembolüyle işlendi."
				ui.rerun()
		with ex_b2:
			if ai_enabled and ai_key:
				if ui.button(f"✨ {ai_provider_lbl} ile {target_year} ÖSYM Sınav Takvimini İşle", key="cfg_btn_ai_exams", type="secondary", use_container_width=True):
					ok, res, used_m = generate_calendar_events_with_ai(
						api_key=ai_key,
						model=ai_model,
						year=target_year,
						month=None,
						preset_type="osym_exams",
						selected_exams=picked_exams,
					)
					if ok:
						added = save_ai_generated_events_to_db(res.get("events") or [], connect_db_fn)
						ui.session_state["_cal_settings_flash"] = f"✨ {ai_provider_lbl} ({used_m}) ile {added} ÖSYM/MEB sınav tarihi takvime eklendi."
						ui.rerun()
					else:
						ui.error(f"AI hatası: {res.get('error')}")

	# -----------------------------------------------------------------------
	# 3. BÖLÜM: EKSTRA EKLEMEK İSTEDİĞİMİZ ŞEYLERİ VEYA DURUMLARI YAZARAK EKLEME KUTUCUĞU
	# -----------------------------------------------------------------------
	with ui.container(border=True):
		ui.markdown("#### ✍️ Serbest Yazarak veya Manuel Etkinlik Ekleme")
		ui.caption(
			"Özel tatilleri (seçim/kar tatili, ara tatil, idari izin), taramaları veya toplantıları **günü ve saatiyle** yazarak ya da manuel form ile takvime ekleyin."
		)

		free_txt_prompt = ui.text_area(
			f"🤖 {ai_provider_lbl} ile Yazarak Etkinlik / Tatil Ekle",
			placeholder=(
				f"Örn: '25 Kasım {target_year} Çarşamba saat 14:00-15:30 arası Veli Toplantısı ekle', "
				f"'18 Aralık {target_year} Cuma gününü kar tatili (idari izin) olarak işaretle', "
				f"'12-16 Kasım {target_year} tarihlerini MEB 1. Dönem Ara Tatili olarak takvime işle'..."
			),
			height=85,
			key="cfg_cal_free_text_box",
		)
		ft_c1, ft_c2 = ui.columns([1.6, 1.4])
		with ft_c1:
			if ui.button(f"✨ {ai_provider_lbl} ile Analiz Et ve Takvime Ekle", key="cfg_btn_ai_free_add", type="primary", use_container_width=True):
				if not free_txt_prompt.strip():
					ui.warning("Lütfen eklenecek etkinlik, gün, saat veya tatil durumunu yukarıdaki kutucuğa yazın.")
				elif not (ai_enabled and ai_key):
					ui.error("Yapay Zekâ şu an kapalı veya API anahtarı girilmemiş. Lütfen Ayarlar → Yapay Zekâ & KVKK bölümünden aktif edin veya aşağıdaki manuel ekleme kutusunu kullanın.")
				else:
					ok, res, used_m = generate_calendar_events_with_ai(
						api_key=ai_key,
						model=ai_model,
						year=target_year,
						month=target_month,
						preset_type="election_or_custom",
						custom_prompt=free_txt_prompt.strip(),
					)
					if ok:
						ev_items = res.get("events") or []
						added = save_ai_generated_events_to_db(ev_items, connect_db_fn, upsert=True)
						if ev_items and isinstance(ev_items[0], dict) and ev_items[0].get("event_date"):
							try:
								d0 = date.fromisoformat(str(ev_items[0]["event_date"])[:10])
								ui.session_state["cal_selected_date"] = d0.isoformat()
								ui.session_state["cal_view_year"] = d0.year
								ui.session_state["cal_view_month"] = d0.month
							except Exception:
								pass
						summary_txt = res.get("summary") or f"{added} kayıt takvime işlendi."
						ui.session_state["_cal_settings_flash"] = f"✨ {ai_provider_lbl} ({used_m}): {summary_txt} ({added} kayıt takvime eklendi)."
						ui.rerun()
					else:
						ui.error(f"AI hatası: {res.get('error')}")
		with ft_c2:
			if ui.button("📅 Takvim & Planlama Sayfasına Git", key="cfg_btn_goto_calendar_page", type="secondary", use_container_width=True):
				ui.session_state["target_page"] = "📅 Takvim & Planlama"
				ui.session_state["main_menu_page"] = "📅 Takvim & Planlama"
				ui.rerun()

		with ui.expander("📌 Manuel Tarih, Saat ve Durum Seçerek Doğrudan Takvime Ekle (AI Kapalıyken de Çalışır)", expanded=False):
			cat_keys = list(CATEGORY_LABELS.keys())
			cat_disp = [CATEGORY_LABELS[k][0] for k in cat_keys]
			sym_keys = list(SYMBOL_OPTIONS.keys())
			sym_disp = [SYMBOL_OPTIONS[k] for k in sym_keys]

			m_cat_disp = ui.selectbox("Kategori", cat_disp, index=cat_keys.index("revir_etkinlik"), key="cfg_cal_manual_cat")
			m_cat = cat_keys[cat_disp.index(m_cat_disp)]
			def_sym = CATEGORY_LABELS[m_cat][2]
			def_hol = m_cat in ("milli_bayram", "resmi_tatil", "yarim_gun", "secim_idari_tatil")
			def_half = m_cat == "yarim_gun"

			with ui.form(f"cfg_form_manual_cal_add_{m_cat}"):
				mc1, mc2, mc3 = ui.columns([1.3, 1.0, 1.0])
				with mc1:
					m_date_str = ui.text_input("Tarih (YYYY-AA-GG)", value=str(ui.session_state.get("cal_selected_date") or today.isoformat())[:10])
				with mc2:
					m_start = ui.text_input("Başlangıç Saati", placeholder="09:00")
				with mc3:
					m_end = ui.text_input("Bitiş Saati", placeholder="12:00")

				m_title = ui.text_input("Etkinlik / Tatil / Durum Başlığı", placeholder="Örn: Seçim Sonrası İdari İzin veya 4-B Göz Taraması")
				m_sym_disp = ui.selectbox("Takvim Sembolü", sym_disp, index=sym_keys.index(def_sym))
				m_sym = sym_keys[sym_disp.index(m_sym_disp)]

				mh1, mh2 = ui.columns(2)
				with mh1:
					m_is_hol = ui.checkbox("Devlet Memuru Resmi Tatili (Kırmızı Yaz)", value=def_hol)
				with mh2:
					m_is_half = ui.checkbox("Yarım Gün Mesai (½ Rozeti)", value=def_half)

				m_prof = ui.text_input("İlgili Meslek Grubu (Opsiyonel)", placeholder="Örn: Okul Hemşireleri ve Öğretmenler")
				m_desc = ui.text_area("Açıklama / Not", placeholder="Detaylı bilgi veya açıklama...", height=55)

				if ui.form_submit_button("💾 Manuel Kaydı Takvime Ekle", type="primary", use_container_width=True):
					clean_d = m_date_str.strip()[:10]
					try:
						d_parsed = date.fromisoformat(clean_d)
					except Exception:
						d_parsed = None
					if not d_parsed:
						ui.error("Lütfen tarihi YYYY-AA-GG biçiminde girin (Örn: 2026-11-25).")
					elif not m_title.strip():
						ui.warning("Lütfen bir başlık girin.")
					else:
						with connect_db_fn() as conn:
							conn.execute(
								"""INSERT INTO calendar_events
								   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
								    profession_group, start_time, end_time, description, source, is_completed, created_at)
								   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'manual', 0, ?)""",
								(
									clean_d,
									m_title.strip(),
									m_cat,
									m_sym,
									1 if (m_is_hol or m_is_half) else 0,
									1 if m_is_half else 0,
									m_prof.strip(),
									m_start.strip()[:5],
									m_end.strip()[:5],
									m_desc.strip(),
									datetime.now().isoformat(timespec="seconds"),
								),
							)
						ui.session_state["cal_selected_date"] = clean_d
						ui.session_state["cal_view_year"] = d_parsed.year
						ui.session_state["cal_view_month"] = d_parsed.month
						ui.session_state["_cal_settings_flash"] = f"✅ '{m_title.strip()}' ({d_parsed.strftime('%d.%m.%Y')}) takvime eklendi."
						ui.rerun()


# ---------------------------------------------------------------------------
# GENEL BAKIŞ (OVERVIEW) TAKVİM & YAKLAŞAN ETKİNLİKLER BİLEŞENİ
# ---------------------------------------------------------------------------

def render_overview_calendar_widget(
	*,
	connect_db_fn: Callable,
	get_active_gemini_config_fn: Callable,
) -> None:
	"""Genel Bakış (Bento Sağ Panel) üzerinde bugünkü ve yaklaşan 14 günlük takvim etkinliklerini gösterir."""
	today_d = date.today()
	today_iso = today_d.isoformat()
	end_iso = (today_d + timedelta(days=14)).isoformat()

	# Yılın temel tatilleri henüz eklenmemişse otomatik hazırla
	ov_seed_key = f"_cal_seeded_{today_d.year}"
	if not ui.session_state.get(ov_seed_key):
		try:
			ensure_year_builtin_holidays_seeded(today_d.year, connect_db_fn, include_special_days=False)
			ui.session_state[ov_seed_key] = True
		except Exception:
			pass

	_handle_calendar_event_inline_click(connect_db_fn, set_flash=False)

	with connect_db_fn() as conn:
		rows = conn.execute(
			"""SELECT id, event_date, title, category, symbol_type, is_civil_holiday,
			          is_half_day, start_time, end_time, is_completed
			   FROM calendar_events
			   WHERE event_date >= ? AND event_date <= ?
			   ORDER BY event_date ASC, is_civil_holiday DESC, start_time ASC, id ASC
			   LIMIT 12""",
			(today_iso, end_iso),
		).fetchall()

	with ui.container(border=True):
		ov_h1, ov_h2, ov_h3 = ui.columns([2.0, 0.5, 1.0])
		with ov_h1:
			ui.markdown("##### 📅 Takvim & Yaklaşan Etkinlikler")
		with ov_h2:
			if ui.button("➕", key="ov_bento_cal_add", help="Bugüne etkinlik / plan ekle", use_container_width=True):
				open_calendar_event_manager_dialog(
					sel_iso=today_iso,
					connect_db_fn=connect_db_fn,
					get_active_gemini_config_fn=get_active_gemini_config_fn,
				)
		with ov_h3:
			if ui.button("Takvime Git ↗", key="ov_bento_cal_go", use_container_width=True):
				ui.session_state["cal_selected_date"] = today_iso
				ui.session_state["cal_view_year"] = today_d.year
				ui.session_state["cal_view_month"] = today_d.month
				ui.session_state["target_page"] = "📅 Takvim & Planlama"
				ui.session_state["main_menu_page"] = "📅 Takvim & Planlama"
				ui.rerun()

		active_weeks = get_active_meb_weeks_for_date(today_d)
		hijri_today = gregorian_to_hijri_str(today_d)
		wd_today = TR_WEEKDAY_NAMES[today_d.weekday()]
		today_hol = any(str(r["event_date"])[:10] == today_iso and bool(r["is_civil_holiday"]) for r in rows)
		if today_hol:
			st_badge = "<span style='background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:800;'>🇹🇷 RESMİ TATİL</span>"
		elif today_d.weekday() >= 5:
			st_badge = "<span style='background:#e2e8f0;color:#475569;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:800;'>🏖️ HAFTA SONU</span>"
		else:
			st_badge = "<span style='background:#d1fae5;color:#047857;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:800;'>🏢 İŞ GÜNÜ</span>"

		ui.markdown(
			f"""<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;
			padding:7px 10px;border-radius:8px;background:var(--bg-card);border:1px solid var(--line);margin-bottom:6px;">
				<span style="font-size:12px;font-weight:800;color:var(--text-heading);">
					🗓️ {today_d.day:02d} {TR_MONTH_NAMES[today_d.month]} {today_d.year} {wd_today} · <span style="color:var(--text-muted);font-weight:700;">🌙 {html.escape(hijri_today)}</span>
				</span>
				{st_badge}
			</div>""",
			unsafe_allow_html=True,
		)
		if active_weeks:
			wk_titles = " · ".join(w["title"] for w in active_weeks[:2])
			ui.caption(f"🏫 **Bu Hafta:** {wk_titles}")

		if not rows:
			ui.caption("Önümüzdeki 14 gün içinde kayıtlı takvim etkinliği veya resmi tatil bulunmuyor. Sağ üstteki **➕** butonuyla hemen ekleyebilirsiniz.")
			return

		for r in rows[:6]:
			ev_id = int(r["id"])
			ev_d_str = str(r["event_date"])[:10]
			try:
				ev_d = date.fromisoformat(ev_d_str)
				if ev_d == today_d:
					d_lbl = "BUGÜN"
				elif ev_d == today_d + timedelta(days=1):
					d_lbl = "YARIN"
				else:
					d_lbl = f"{ev_d.day:02d} {TR_MONTH_NAMES[ev_d.month][:3]}"
			except Exception:
				d_lbl = ev_d_str

			sym_html = render_symbol_badge(r["symbol_type"] or "event")
			t_rng = " - ".join(filter(None, [r["start_time"], r["end_time"]]))
			t_html = f"<span style='color:var(--teal);font-weight:700;font-size:11.5px;'>⏰ {html.escape(t_rng)}</span>" if t_rng else ""
			is_hol = bool(r["is_civil_holiday"])
			is_done = bool(r["is_completed"])
			d_bg = "#d1fae5" if is_done else ("#fee2e2" if is_hol else ("#dbeafe" if ev_d_str == today_iso else "#f1f5f9"))
			d_fg = "#047857" if is_done else ("#dc2626" if is_hol else ("#1d4ed8" if ev_d_str == today_iso else "#334155"))
			ttl_style = "text-decoration:line-through;opacity:0.65;" if is_done else ""
			done_badge = "<span style='background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;padding:0 5px;border-radius:4px;font-size:10px;font-weight:800;flex-shrink:0;'>✓ Tamamlandı</span>" if is_done else ""
			done_btn_css = "background:#10b981;color:#ffffff;border:1px solid #059669;" if is_done else "background:#ffffff;color:#059669;border:1px solid #6ee7b7;"
			done_title = "Tamamlandı işaretini geri al" if is_done else "Tamamlandı olarak işaretle"
			row_tint_css = "background:#ecfdf5;border-color:#6ee7b7;" if is_done else ""

			ui.markdown(
				f"""<div class="revir-op-row" style="padding:6px 9px;display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:4px;{row_tint_css}">
					<div style="display:flex;align-items:center;gap:7px;min-width:0;flex:1;">
						<span style="background:{d_bg};color:{d_fg};padding:2px 7px;border-radius:6px;font-size:11px;font-weight:800;white-space:nowrap;">{html.escape(d_lbl)}</span>
						{sym_html}
						<span style="font-weight:700;font-size:13px;color:var(--text-heading);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;{ttl_style}">{html.escape(str(r['title'] or ''))}</span>
						{done_badge}
					</div>
					<div style="display:inline-flex;align-items:center;gap:5px;white-space:nowrap;flex-shrink:0;">
						{t_html}
						<button type="button" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_done_{ev_id}');" onkeydown="event.stopPropagation();" title="{done_title}" style="width:24px;height:24px;padding:0;border-radius:6px;display:inline-flex;align-items:center;justify-content:center;font-size:12.5px;font-weight:900;line-height:1;cursor:pointer;{done_btn_css}">✓</button>
						<button type="button" onclick="event.stopPropagation();window.__revirClick('btn_cal_ev_del_{ev_id}');" onkeydown="event.stopPropagation();" title="Bu etkinliği takvimden kaldır" style="width:24px;height:24px;padding:0;border-radius:6px;display:inline-flex;align-items:center;justify-content:center;font-size:12.5px;font-weight:900;line-height:1;cursor:pointer;background:#ffffff;color:#dc2626;border:1px solid #fca5a5;">✕</button>
					</div>
				</div>""",
				unsafe_allow_html=True,
			)

		if len(rows) > 6:
			ui.caption(f"… ve önümüzdeki 14 gün içinde {len(rows) - 6} takvim kaydı daha.")


