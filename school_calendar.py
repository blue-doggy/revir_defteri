# -*- coding: utf-8 -*-
"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Akıllı Okul & Revir Takvimi,
İnteraktif Takvim Yaprağı, 81 İl Ezan Vakitleri Canlı Sayacı ve Gemini AI Planlama Modülü.
"""
from __future__ import annotations

import calendar
import html
import json
import math
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
	"belirli_gun": ("Belirli Gün ve Hafta", "#0284c7", "event"),
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
	"event": "📌 Standart Etkinlik / Belirli Gün",
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
	return '<span style="font-size:12px;" title="Etkinlik / Belirli Gün">📌</span>'


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
# HİCRİ TAKVİM HESAPLAYICI (KUWAITI / UMM AL-QURA YAKLAŞIMI)
# ---------------------------------------------------------------------------

def gregorian_to_hijri_str(d: date) -> str:
	"""Miladi tarihi Türkçe Hicri takvim metnine (Örn: '20 Cemaziyelevvel 1448') dönüştürür."""
	jd = int((1461 * (d.year + 4800 + int((d.month - 14) / 12))) / 4)
	jd += int((367 * (d.month - 2 - 12 * int((d.month - 14) / 12))) / 12)
	jd -= int((3 * int((d.year + 4900 + int((d.month - 14) / 12)) / 100)) / 4)
	jd += d.day - 32075

	l_val = jd - 1948440 + 10632
	n_val = int((l_val - 1) / 10631)
	l_val = l_val - 10631 * n_val + 354
	j_val = (int((10985 - l_val) / 5316)) * (int((50 * l_val) / 17719)) + (int(l_val / 5670)) * (int((43 * l_val) / 15238))
	l_val = l_val - (int((30 - j_val) / 15)) * (int((17719 * j_val) / 50)) - (int(j_val / 16)) * (int((15238 * j_val) / 43)) + 29
	h_month = int((24 * l_val) / 709)
	h_day = l_val - int((709 * h_month) / 24)
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

	# Asr-ı Evvel (Şafii/Hanefi Diyanet standart gölge = 1 + tan(|lat - decl|))
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
		# Büyük/küçük harf eşleştirmesi
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

	# SQLite önbellek kontrolü
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

	# Çevrimiçi Diyanet (Method 13) sorgusu (kısa timeout ile arayüzü bloklamaz)
	lat, lon = TR_PROVINCE_COORDS.get(prov_clean, TR_PROVINCE_COORDS["Ankara"])
	result: dict[str, str] | None = None
	try:
		d_str = d.strftime("%d-%m-%Y")
		url = f"https://api.aladhan.com/v1/timings/{d_str}?latitude={lat}&longitude={lon}&method=13"
		req = urllib.request.Request(url, headers={"User-Agent": "RevirDefteri/1.14"})
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

	# SQLite önbelleğine kaydet
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
# YERLEŞİK TÜRKİYE RESMİ & DİNİ TATİLLER, YARIM GÜNLER, MESLEK GÜNLERİ VE SINAVLAR
# ---------------------------------------------------------------------------

# Sabit Miladi Resmi Tatiller ve Ulusal Günler (Ay, Gün)
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
		"title": "18 Mart Çanakkale Zaferi ve Şehitleri Anma Günü",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": False, "is_half_day": False,
		"description": "Çanakkale Deniz Zaferi'nin yıl dönümü ve Aziz Şehitlerimizi Anma Günü.",
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
		"title": "10 Kasım Atatürk'ü Anma Günü",
		"category": "milli_bayram", "symbol_type": "tr_flag",
		"is_civil_holiday": False, "is_half_day": False,
		"start_time": "09:05",
		"description": "Cumhuriyetimizin kurucusu Gazi Mustafa Kemal Atatürk'ü Anma Günü ve Atatürk Haftası başlangıcı.",
	},
]

# Diyanet Takvimine Göre Yıllara Özel Dini Bayramlar, Arifeler ve Kandiller (2025, 2026, 2027, 2028)
RELIGIOUS_DAYS_BY_YEAR: dict[int, list[dict[str, Any]]] = {
	2025: [
		{"date": "2025-01-02", "title": "Regaip Kandili", "holiday": False, "half": False, "desc": "Üç ayların müjdecisi mübarek Regaip Kandili."},
		{"date": "2025-01-26", "title": "Miraç Kandili", "holiday": False, "half": False, "desc": "Mübarek Miraç Kandili gecesi."},
		{"date": "2025-02-13", "title": "Berat Kandili", "holiday": False, "half": False, "desc": "Ramazan ayının müjdecisi Berat Kandili."},
		{"date": "2025-03-01", "title": "Ramazan Ayı Başlangıcı", "holiday": False, "half": False, "desc": "On bir ayın sultanı Ramazan-ı Şerif'in ilk günü."},
		{"date": "2025-03-26", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Bin aydan daha hayırlı olan mübarek Kadir Gecesi."},
		{"date": "2025-03-29", "title": "Ramazan Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
		{"date": "2025-03-30", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Resmi tatil."},
		{"date": "2025-03-31", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2025-04-01", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2025-06-05", "title": "Kurban Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
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
		{"date": "2026-02-19", "title": "Ramazan Ayı Başlangıcı", "holiday": False, "half": False, "desc": "On bir ayın sultanı Ramazan-ı Şerif'in ilk günü."},
		{"date": "2026-03-16", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Bin aydan daha hayırlı olan mübarek Kadir Gecesi."},
		{"date": "2026-03-19", "title": "Ramazan Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları ve okullar için yarım gün resmi tatil."},
		{"date": "2026-03-20", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-03-21", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2026-03-22", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2026-05-26", "title": "Kurban Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları ve okullar için yarım gün resmi tatil."},
		{"date": "2026-05-27", "title": "Kurban Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 1. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-28", "title": "Kurban Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 2. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-29", "title": "Kurban Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 3. Günü — Devlet memurları ve okullar için resmi tatil."},
		{"date": "2026-05-30", "title": "Kurban Bayramı (4. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 4. Günü — Resmi tatil."},
		{"date": "2026-06-16", "title": "Hicri Yılbaşı (1 Muharrem 1448)", "holiday": False, "half": False, "desc": "Hicri 1448 takvim yılının ilk günü."},
		{"date": "2026-06-25", "title": "Aşure Günü (10 Muharrem)", "holiday": False, "half": False, "desc": "10 Muharrem Aşure Günü."},
		{"date": "2026-08-24", "title": "Mevlid Kandili", "holiday": False, "half": False, "desc": "Peygamber Efendimiz Hz. Muhammed'in (s.a.v.) dünyaya teşrifi — Mevlid Kandili."},
		{"date": "2026-12-10", "title": "Regaip Kandili (1448)", "holiday": False, "half": False, "desc": "Üç ayların müjdecisi mübarek Regaip Kandili."},
	],
	2027: [
		{"date": "2027-01-04", "title": "Miraç Kandili", "holiday": False, "half": False, "desc": "Mübarek Miraç Kandili gecesi."},
		{"date": "2027-01-22", "title": "Berat Kandili", "holiday": False, "half": False, "desc": "Mübarek Berat Kandili gecesi."},
		{"date": "2027-02-08", "title": "Ramazan Ayı Başlangıcı", "holiday": False, "half": False, "desc": "Ramazan-ı Şerif'in ilk günü."},
		{"date": "2027-03-05", "title": "Kadir Gecesi", "holiday": False, "half": False, "desc": "Mübarek Kadir Gecesi."},
		{"date": "2027-03-08", "title": "Ramazan Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Ramazan Bayramı Arifesi — Saat 13:00'ten itibaren devlet memurları için yarım gün resmi tatil."},
		{"date": "2027-03-09", "title": "Ramazan Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 1. Günü — Resmi tatil."},
		{"date": "2027-03-10", "title": "Ramazan Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 2. Günü — Resmi tatil."},
		{"date": "2027-03-11", "title": "Ramazan Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Ramazan Bayramı 3. Günü — Resmi tatil."},
		{"date": "2027-05-15", "title": "Kurban Bayramı Arifesi (Yarım Gün Mesai)", "holiday": True, "half": True, "desc": "Kurban Bayramı Arifesi — Yarım gün resmi tatil."},
		{"date": "2027-05-16", "title": "Kurban Bayramı (1. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 1. Günü — Resmi tatil."},
		{"date": "2027-05-17", "title": "Kurban Bayramı (2. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 2. Günü — Resmi tatil."},
		{"date": "2027-05-18", "title": "Kurban Bayramı (3. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 3. Günü — Resmi tatil."},
		{"date": "2027-05-19", "title": "Kurban Bayramı (4. Gün)", "holiday": True, "half": False, "desc": "Kurban Bayramı 4. Günü — Resmi tatil."},
		{"date": "2027-06-06", "title": "Hicri Yılbaşı (1 Muharrem 1449)", "holiday": False, "half": False, "desc": "Hicri 1449 yılının ilk günü."},
		{"date": "2027-06-15", "title": "Aşure Günü", "holiday": False, "half": False, "desc": "10 Muharrem Aşure Günü."},
		{"date": "2027-08-14", "title": "Mevlid Kandili", "holiday": False, "half": False, "desc": "Mübarek Mevlid Kandili."},
	],
}

# Meslek Grubu Günleri ve Belirli Gün / Haftalar Sözlüğü ((Ay, Gün) -> Bilgi)
PROFESSION_AND_SPECIAL_DAYS: dict[tuple[int, int], dict[str, str]] = {
	(1, 7): {"title": "Beyaz Baston Görme Engelliler Haftası (7-14 Ocak)", "prof": "Özel Eğitim ve Rehabilitasyon Uzmanları", "cat": "belirli_gun"},
	(1, 10): {"title": "10 Ocak Çalışan Gazeteciler Günü & İdareciler Günü", "prof": "Gazeteciler, Basın Emekçileri ve Mülki İdareciler", "cat": "meslek_gunu"},
	(1, 26): {"title": "Dünya Gümrük Günü", "prof": "Gümrük Muhafaza ve Ticaret Personeli", "cat": "meslek_gunu"},
	(2, 4): {"title": "4 Şubat Dünya Kanser Günü", "prof": "Onkoloji Hekimleri ve Koruyucu Sağlık Çalışanları", "cat": "belirli_gun"},
	(2, 9): {"title": "9 Şubat Dünya Sigarayı Bırakma Günü", "prof": "Göğüs Hastalıkları ve Yeşilay Danışmanları", "cat": "belirli_gun"},
	(2, 28): {"title": "28 Şubat Sivil Savunma Günü", "prof": "AFAD, Arama Kurtarma ve Sivil Savunma Uzmanları", "cat": "meslek_gunu"},
	(3, 1): {"title": "Yeşilay Haftası Başlangıcı (1-7 Mart) & Muhasebeciler Günü", "prof": "Mali Müşavirler, Muhasebeciler ve Bağımlılık Danışmanları", "cat": "meslek_gunu"},
	(3, 8): {"title": "8 Mart Dünya Kadınlar Günü", "prof": "Tüm Emekçi Kadınlar", "cat": "belirli_gun"},
	(3, 12): {"title": "12 Mart İstiklâl Marşı'nın Kabulü ve Mehmet Akif Ersoy'u Anma Günü", "prof": "Edebiyatçılar, Şairler ve Eğitimciler", "cat": "milli_bayram"},
	(3, 14): {"title": "14 Mart Tıp Bayramı", "prof": "Hekimler, Hemşireler ve Tüm Sağlık Çalışanları", "cat": "meslek_gunu"},
	(3, 18): {"title": "18 Mart Şehitleri Anma Günü ve Çanakkale Deniz Zaferi & Yaşlılara Saygı Haftası", "prof": "Geriatri Uzmanları ve Sosyal Hizmet Çalışanları", "cat": "milli_bayram"},
	(3, 21): {"title": "21 Mart Nevruz Günü & Dünya Down Sendromu Farkındalık Günü", "prof": "Özel Eğitim Öğretmenleri ve Çocuk Gelişimciler", "cat": "belirli_gun"},
	(3, 22): {"title": "22 Mart Dünya Su Günü", "prof": "Çevre ve Su Mühendisleri", "cat": "belirli_gun"},
	(3, 24): {"title": "24 Mart Dünya Tüberküloz (Verem) Günü", "prof": "Halk Sağlığı ve Göğüs Hastalıkları Hemşireleri", "cat": "belirli_gun"},
	(3, 30): {"title": "Kütüphaneler Haftası", "prof": "Kütüphaneciler ve Arşiv Uzmanları", "cat": "meslek_gunu"},
	(4, 2): {"title": "2 Nisan Dünya Otizm Farkındalık Günü", "prof": "Çocuk Psikiyatristleri ve Özel Eğitimciler", "cat": "belirli_gun"},
	(4, 5): {"title": "5 Nisan Avukatlar Günü", "prof": "Avukatlar ve Hukukçular", "cat": "meslek_gunu"},
	(4, 7): {"title": "7 Nisan Dünya Sağlık Günü & Sağlık Haftası (7-13 Nisan)", "prof": "Tüm Sağlık Profesyonelleri ve Okul Hemşireleri", "cat": "meslek_gunu"},
	(4, 10): {"title": "10 Nisan Polis Haftası ve Türk Polis Teşkilatı Kuruluş Günü", "prof": "Emniyet Mensupları ve Polisler", "cat": "meslek_gunu"},
	(4, 15): {"title": "Turizm Haftası (15-22 Nisan)", "prof": "Turizm Çalışanları ve Rehberler", "cat": "meslek_gunu"},
	(4, 21): {"title": "21-28 Nisan Ebeler Haftası", "prof": "Ebeler ve Ana-Çocuk Sağlığı Çalışanları", "cat": "meslek_gunu"},
	(4, 28): {"title": "28 Nisan Dünya İş Sağlığı ve Güvenliği Günü", "prof": "İşyeri Hekimleri, İSG Uzmanları ve Kurum Hemşireleri", "cat": "meslek_gunu"},
	(5, 1): {"title": "1 Mayıs Emek ve Dayanışma Günü", "prof": "Tüm İşçiler, Emekçiler ve Kamu Görevlileri", "cat": "meslek_gunu"},
	(5, 3): {"title": "Trafik ve İlkyardım Haftası (Mayıs İlk Haftası)", "prof": "Paramedikler, ATT'ler ve İlkyardım Eğitmenleri", "cat": "meslek_gunu"},
	(5, 6): {"title": "6 Mayıs Hıdırellez Kültür ve Bahar Bayramı", "prof": "Kültür ve Halkbilim Çalışanları", "cat": "belirli_gun"},
	(5, 10): {"title": "10-16 Mayıs Engelliler Haftası & Psikologlar Günü", "prof": "Psikologlar, PDR Uzmanları ve Özel Eğitimciler", "cat": "meslek_gunu"},
	(5, 12): {"title": "12 Mayıs Dünya Hemşireler Günü & Hemşirelik Haftası (12-18 Mayıs)", "prof": "Okul Revir Hemşireleri ve Tüm Hemşireler", "cat": "meslek_gunu"},
	(5, 14): {"title": "14 Mayıs Eczacılık Günü & Dünya Çiftçiler Günü", "prof": "Eczacılar ve Çiftçiler / Ziraat Mühendisleri", "cat": "meslek_gunu"},
	(5, 18): {"title": "18 Mayıs Müzeler Günü", "prof": "Müzeciler, Arkeologlar ve Sanat Tarihçileri", "cat": "meslek_gunu"},
	(5, 21): {"title": "21 Mayıs Dünya Süt Günü", "prof": "Gıda Mühendisleri ve Beslenme Uzmanları", "cat": "belirli_gun"},
	(6, 5): {"title": "5 Haziran Dünya Çevre Günü", "prof": "Çevre Mühendisleri ve Doğa Koruma Görevlileri", "cat": "meslek_gunu"},
	(6, 6): {"title": "6 Haziran Diyetisyenler Günü", "prof": "Diyetisyenler ve Beslenme Uzmanları", "cat": "meslek_gunu"},
	(6, 14): {"title": "14 Haziran Dünya Gönüllü Kan Bağışçıları Günü", "prof": "Kızılay ve Kan Merkezi Sağlık Çalışanları", "cat": "belirli_gun"},
	(7, 1): {"title": "1 Temmuz Denizcilik ve Kabotaj Bayramı", "prof": "Denizciler, Kaptanlar ve Gemi Adamları", "cat": "meslek_gunu"},
	(7, 24): {"title": "24 Temmuz Gazeteciler ve Basın Bayramı", "prof": "Basın ve Medya Çalışanları", "cat": "meslek_gunu"},
	(9, 1): {"title": "1 Eylül Dünya Barış Günü & Adli Yıl Açılışı", "prof": "Hakimler, Savcılar ve Hukuk Çalışanları", "cat": "meslek_gunu"},
	(9, 8): {"title": "8 Eylül Dünya Fizyoterapistler Günü", "prof": "Fizyoterapistler", "cat": "meslek_gunu"},
	(9, 15): {"title": "İlköğretim Haftası (Eylül 3. Haftası)", "prof": "Öğretmenler, Okul İdarecileri ve Okul Hemşireleri", "cat": "belirli_gun"},
	(9, 25): {"title": "25 Eylül - 1 Ekim İtfaiyecilik Haftası", "prof": "İtfaiyeciler ve Yangın Güvenliği Personeli", "cat": "meslek_gunu"},
	(9, 29): {"title": "29 Eylül Dünya Kalp Günü", "prof": "Kardiyologlar ve Koruyucu Sağlık Ekibi", "cat": "belirli_gun"},
	(10, 4): {"title": "4 Ekim Hayvanları Koruma Günü", "prof": "Veteriner Hekimler", "cat": "meslek_gunu"},
	(10, 10): {"title": "10 Ekim Dünya Ruh Sağlığı Günü", "prof": "Psikiyatristler, Psikologlar ve Okul Rehber Öğretmenleri", "cat": "meslek_gunu"},
	(10, 15): {"title": "15 Ekim Dünya El Yıkama Günü", "prof": "Okul Sağlığı Hemşireleri ve Enfeksiyon Kontrol Ekibi", "cat": "belirli_gun"},
	(10, 19): {"title": "19 Ekim Muhtarlar Günü", "prof": "Mahalle ve Köy Muhtarları", "cat": "meslek_gunu"},
	(10, 29): {"title": "29 Ekim - 4 Kasım Kızılay Haftası", "prof": "Kızılay Çalışanları ve İnsani Yardım Gönüllüleri", "cat": "belirli_gun"},
	(11, 2): {"title": "2-8 Kasım Lösemili Çocuklar Haftası", "prof": "Pediatrik Hematoloji Hekimleri ve Hemşireleri", "cat": "belirli_gun"},
	(11, 3): {"title": "3-9 Kasım Organ Bağışı ve Nakli Haftası", "prof": "Organ Nakli Koordinatörleri ve Cerrahlar", "cat": "belirli_gun"},
	(11, 8): {"title": "8 Kasım Dünya Radyoloji Günü", "prof": "Radyoloji Uzmanları ve Teknikerleri", "cat": "meslek_gunu"},
	(11, 14): {"title": "14 Kasım Dünya Diyabet Günü", "prof": "Endokrinologlar, Diyabet Eğitim Hemşireleri ve Okul Hemşireleri", "cat": "belirli_gun"},
	(11, 20): {"title": "20 Kasım Dünya Çocuk Hakları Günü", "prof": "Çocuk Sağlığı Uzmanları ve Eğitimciler", "cat": "belirli_gun"},
	(11, 22): {"title": "22 Kasım Diş Hekimleri Günü & Ağız ve Diş Sağlığı Haftası", "prof": "Diş Hekimleri ve Ağız-Diş Sağlığı Çalışanları", "cat": "meslek_gunu"},
	(11, 24): {"title": "24 Kasım Öğretmenler Günü", "prof": "Öğretmenler ve Eğitim Camiası", "cat": "meslek_gunu"},
	(12, 3): {"title": "3 Aralık Dünya Engelliler Günü", "prof": "Özel Eğitim, Fizyoterapi ve Sosyal Hizmet Uzmanları", "cat": "belirli_gun"},
	(12, 4): {"title": "4 Aralık Dünya Madenciler Günü", "prof": "Maden İşçileri ve Maden Mühendisleri", "cat": "meslek_gunu"},
	(12, 5): {"title": "5 Aralık Dünya Mühendisler Günü & Türk Kadınına Seçme ve Seçilme Hakkı Günü", "prof": "Mühendisler", "cat": "meslek_gunu"},
	(12, 10): {"title": "10 Aralık Dünya İnsan Hakları Günü", "prof": "Hukukçular ve İnsan Hakları Savunucuları", "cat": "belirli_gun"},
	(12, 12): {"title": "12-18 Aralık Tutum, Yatırım ve Türk Malları (Yerli Malı) Haftası", "prof": "Yerli Üreticiler, Esnaf ve Sanayiciler", "cat": "belirli_gun"},
	(12, 22): {"title": "22 Aralık Dünya Paramedikler Günü", "prof": "Paramedikler ve 112 Acil Sağlık Ekibi", "cat": "meslek_gunu"},
}

# Haftanın gününe göre varsayılan meslek/topluluk takdiri (Özel bir meslek günü olmayan günler için)
WEEKDAY_DEFAULT_PROFESSION_THEME: dict[int, tuple[str, str]] = {
	0: ("Okul Sağlığı Hemşireleri & Koruyucu Sağlık Ekibi", "Haftanın ilk iş gününde öğrencilerin sağlığı için görev yapan okul revir sorumluları."),
	1: ("Öğretmenler & Okul Rehberlik (PDR) Servisi", "Öğrencilerin akademik ve psikososyal gelişimine rehberlik eden eğitimciler."),
	2: ("Hekimler, Aile Sağlığı & Toplum Sağlığı Çalışanları", "Birinci basamak sağlık hizmetleri ve koruyucu hekimlik emekçileri."),
	3: ("Kamu Görevlileri, İdareciler & Bürokratlar", "Kamu hizmetlerinin aksamadan yürütülmesini sağlayan devlet memurları."),
	4: ("Destek Personeli, Beslenme & Hijyen Görevlileri", "Okul hijyeni, güvenli beslenme ve sağlıklı eğitim ortamını koruyan emekçiler."),
	5: ("112 Acil Sağlık, Nöbetçi Eczane & Güvenlik Görevlileri", "Hafta sonu nöbet başında toplum sağlığı ve güvenliğini koruyanlar."),
	6: ("Aileler, Veliler & Gönüllü Yardım Kuruluşları", "Çocukların sağlıklı büyümesinde okul ile iş birliği yapan aileler."),
}

# Yerleşik ÖSYM & MEB Merkezi Sınav Takvimi (Çevrimdışı Hızlı Yükleme İçin)
BUILTIN_OSYM_MEB_EXAMS: list[dict[str, Any]] = [
	{"code": "MSÜ", "month": 3, "day": 1, "time": "10:15", "title": "ÖSYM - MSÜ (Millî Savunma Üniversitesi Askerî Öğrenci Aday Belirleme Sınavı)", "desc": "ÖSYM tarafından uygulanan MSÜ sınavı (Saat: 10:15)."},
	{"code": "YDS / YÖKDİL", "month": 3, "day": 8, "time": "10:15", "title": "ÖSYM - YÖKDİL / 1 (Yükseköğretim Kurumları Yabancı Dil Sınavı)", "desc": "ÖSYM YÖKDİL İlkbahar dönemi sınavı."},
	{"code": "YDS / YÖKDİL", "month": 4, "day": 5, "time": "10:15", "title": "ÖSYM - YDS / 1 (Yabancı Dil Bilgisi Seviye Tespit Sınavı)", "desc": "ÖSYM YDS İlkbahar dönemi sınavı."},
	{"code": "İOKBS Bursluluk", "month": 4, "day": 26, "time": "10:00", "title": "MEB - İOKBS (İlköğretim ve Ortaöğretim Kurumları Bursluluk Sınavı)", "desc": "MEB merkezi bursluluk sınavı."},
	{"code": "ALES", "month": 5, "day": 10, "time": "10:15", "title": "ÖSYM - ALES / 1 (Akademik Personel ve Lisansüstü Eğitimi Giriş Sınavı)", "desc": "ÖSYM ALES-1 oturumu."},
	{"code": "LGS", "month": 6, "day": 14, "time": "09:30", "title": "MEB - LGS (Liselere Geçiş Sistemi Merkezi Sınavı)", "desc": "8. sınıf öğrencileri için MEB LGS Sözel (09:30) ve Sayısal (11:30) oturumları."},
	{"code": "YKS (TYT-AYT-YDT)", "month": 6, "day": 20, "time": "10:15", "title": "ÖSYM - YKS 1. Oturum: TYT (Temel Yeterlilik Testi)", "desc": "Yükseköğretim Kurumları Sınavı TYT oturumu (Saat 10:15)."},
	{"code": "YKS (TYT-AYT-YDT)", "month": 6, "day": 21, "time": "10:15", "title": "ÖSYM - YKS 2. Oturum: AYT & 3. Oturum: YDT", "desc": "YKS Alan Yeterlilik Testleri (10:15) ve Yabancı Dil Testi (15:45)."},
	{"code": "DGS", "month": 7, "day": 19, "time": "10:15", "title": "ÖSYM - DGS (Dikey Geçiş Sınavı)", "desc": "Meslek yüksekokulu mezunları için Dikey Geçiş Sınavı."},
	{"code": "KPSS", "month": 9, "day": 6, "time": "10:15", "title": "ÖSYM - KPSS Lisans (Genel Yetenek - Genel Kültür & Eğitim Bilimleri)", "desc": "Kamu Personel Seçme Sınavı Lisans oturumu."},
	{"code": "KPSS", "month": 10, "day": 4, "time": "10:15", "title": "ÖSYM - KPSS Ön Lisans Oturumu", "desc": "Kamu Personel Seçme Sınavı Ön Lisans oturumu."},
	{"code": "KPSS", "month": 10, "day": 25, "time": "10:15", "title": "ÖSYM - KPSS Ortaöğretim (Lise) Oturumu", "desc": "Kamu Personel Seçme Sınavı Ortaöğretim oturumu."},
	{"code": "ALES", "month": 11, "day": 22, "time": "10:15", "title": "ÖSYM - ALES / 3 (Sonbahar Dönemi)", "desc": "ÖSYM ALES-3 oturumu."},
]


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
# VERİTABANI YARDIMCILARI VE OTOMATİK RESMİ TATİL YÜKLEYİCİ
# ---------------------------------------------------------------------------

def ensure_year_builtin_holidays_seeded(year: int, connect_db_fn: Callable, include_special_days: bool = False) -> int:
	"""
	Belirtilen yıl için resmi tatiller, yarım gün arife/28 Ekim çalışmaları ve dini bayramlar
	henüz yüklenmemişse veritabanına ekler.
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

		# 3. İsteğe bağlı olarak Meslek Günleri ve Belirli Günler
		if include_special_days:
			for (m_no, d_no), sp_info in PROFESSION_AND_SPECIAL_DAYS.items():
				ev_date = f"{year:04d}-{m_no:02d}-{d_no:02d}"
				t_key = (ev_date, sp_info["title"].strip().lower())
				if t_key in existing_keys:
					continue
				cat = sp_info.get("cat", "meslek_gunu")
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
						f"Günün Meslek Grubu: {sp_info.get('prof', '')}",
						now_iso,
					),
				)
				existing_keys.add(t_key)
				added += 1

	return added


def seed_selected_exams_to_db(year: int, selected_exams: list[str], connect_db_fn: Callable) -> int:
	"""Seçilen ÖSYM ve MEB sınavlarını ilgili yılın takvimine işler."""
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
			ev_date = f"{year:04d}-{ex['month']:02d}-{ex['day']:02d}"
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


def save_ai_generated_events_to_db(events: list[dict[str, Any]], connect_db_fn: Callable) -> int:
	"""Gemini AI tarafından üretilen takvim kayıtlarını tekilleştirerek SQLite veritabanına kaydeder."""
	added = 0
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db_fn() as conn:
		for ev in events:
			if not isinstance(ev, dict):
				continue
			ev_date = str(ev.get("event_date") or "").strip()[:10]
			title = str(ev.get("title") or "").strip()
			if len(ev_date) != 10 or not title:
				continue
			dup = conn.execute(
				"SELECT id FROM calendar_events WHERE event_date = ? AND lower(trim(title)) = lower(trim(?))",
				(ev_date, title),
			).fetchone()
			if dup:
				continue
			cat = str(ev.get("category") or "belirli_gun").strip()
			if cat not in CATEGORY_LABELS:
				cat = "belirli_gun"
			sym = str(ev.get("symbol_type") or CATEGORY_LABELS[cat][2]).strip()
			if sym not in SYMBOL_OPTIONS:
				sym = CATEGORY_LABELS[cat][2]
			is_hol = 1 if bool(ev.get("is_civil_holiday")) else 0
			is_half = 1 if bool(ev.get("is_half_day")) else 0
			prof = str(ev.get("profession_group") or "").strip()
			st_time = str(ev.get("start_time") or "").strip()[:5]
			en_time = str(ev.get("end_time") or "").strip()[:5]
			desc = str(ev.get("description") or "").strip()
			conn.execute(
				"""INSERT INTO calendar_events
				   (event_date, title, category, symbol_type, is_civil_holiday, is_half_day,
				    profession_group, start_time, end_time, description, source, is_completed, created_at)
				   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'gemini_ai', 0, ?)""",
				(ev_date, title, cat, sym, is_hol, is_half, prof, st_time, en_time, desc, now_iso),
			)
			added += 1
	return added


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
) -> None:
	"""
	GELİŞİM & PLAN -> 📅 Takvim & Planlama ana sayfasını oluşturur:
	- Üst Bar: Ay/Yıl seçici, Önceki/Sonraki ay butonları, Ayın İş Günü / Hafta Sonu / Resmi Tatil özeti
	- ✨ Gemini AI Takvim Sihirbazı & Hızlı Tatil / ÖSYM Sınavı / Meslek Günü İşaretleyici
	- Sol Panel (%62): 7 sütunlu renk kodlu ve bayrak/hilal sembollü Aylık Takvim Izgarası
	- Sağ Panel (%38): Seçili Güne Ait Klasik & Modern 'Takvim Yaprağı' (Meslek Grubu, Tatil,
	  İl Bazlı Canlı Ezan Vakti Geri Sayım Sayacı, Ayarlara Göre Hadis/Özlü Söz ve Etkinlik Yönetimi)
	"""
	today = date.today()
	cfg = get_app_settings_fn()
	ai_enabled, ai_key, ai_model = get_active_gemini_config_fn()

	# Seçili tarih, ay ve yıl state yönetimi
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

	# Yılın temel resmi ve dini tatillerini ilk açılışta otomatik hazırla
	seed_flag_key = f"_cal_seeded_{view_year}"
	if not ui.session_state.get(seed_flag_key) and str(cfg.get("calendar_auto_load_holidays", "1")) != "0":
		ensure_year_builtin_holidays_seeded(view_year, connect_db_fn, include_special_days=False)
		ui.session_state[seed_flag_key] = True

	# Başlık ve Üst Navigasyon
	ui.markdown("## 📅 Okul & Revir Takvimi, Resmi Tatiller ve Günlük Takvim Yaprağı")
	ui.caption(
		"İş günleri, hafta sonları, devlet memurları için **kırmızı** resmi tatiller ve **½ yarım gün** arife mesaileri, "
		"🇹🇷 milli bayramlar, 🌙 dini bayramlar/kandiller, 📝 ÖSYM/MEB sınavları, 👩‍⚕️ meslek günleri ve 🕌 il bazlı canlı ezan vakti sayacı."
	)

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
			ui.rerun()

	# -----------------------------------------------------------------------
	# GEMINI AI TAKVİM SİHİRBAZI & HIZLI TATİL / SINAV / MESLEK GÜNÜ PANELİ
	# -----------------------------------------------------------------------
	with ui.expander("✨ Gemini AI Takvim Sihirbazı · Resmi/Dini Tatil, Yarım Gün, ÖSYM Sınavları ve Özel Gün İşaretleyici", expanded=False):
		wiz_t1, wiz_t2, wiz_t3 = ui.tabs([
			"🇹🇷 Resmi & Dini Tatiller / Meslek Günleri",
			"📝 ÖSYM & MEB Sınav Takvimi",
			"🤖 Gemini AI Serbest Takvim Komutu (Seçim Tatili, Ara Tatil vb.)",
		])

		with wiz_t1:
			ui.caption(
				f"**{view_year}** yılı için Türkiye Cumhuriyeti resmi tatillerini, devlet memurları yarım gün arife/28 Ekim mesailerini, "
				"milli bayramları (🇹🇷), dini bayram ve kandilleri (🌙) veya 365 günlük meslek grubu & belirli günleri tek tıkla takvime işleyin."
			)
			wc1, wc2, wc3 = ui.columns(3)
			with wc1:
				if ui.button(f"📥 {view_year} Resmi & Dini Tatilleri + Yarım Günleri İşle", key="btn_seed_official_holidays", type="primary", use_container_width=True):
					cnt = ensure_year_builtin_holidays_seeded(view_year, connect_db_fn, include_special_days=False)
					ui.session_state["_cal_flash_msg"] = f"✅ {view_year} yılı resmi tatilleri, yarım gün arife mesaileri ve dini bayramlar kontrol edildi ({cnt} yeni kayıt eklendi)."
					ui.rerun()
			with wc2:
				if ui.button(f"👩‍⚕️ {view_year} Meslek Günleri & Belirli Gün/Haftaları İşle", key="btn_seed_profession_days", type="secondary", use_container_width=True):
					cnt = ensure_year_builtin_holidays_seeded(view_year, connect_db_fn, include_special_days=True)
					ui.session_state["_cal_flash_msg"] = f"✅ {view_year} yılı meslek günleri (Tıp Bayramı, Hemşireler Günü, Öğretmenler Günü vb.) ve belirli gün/haftalar takvime işlendi ({cnt} yeni kayıt)."
					ui.rerun()
			with wc3:
				if ai_enabled and ai_key:
					if ui.button(f"✨ Gemini AI ile {TR_MONTH_NAMES[view_month]} {view_year} Özel Günlerini Zenginleştir", key="btn_ai_enrich_month", type="secondary", use_container_width=True):
						ok, res, used_m = generate_calendar_events_with_ai(
							api_key=ai_key,
							model=ai_model,
							year=view_year,
							month=view_month,
							preset_type="special_days",
						)
						if ok:
							added = save_ai_generated_events_to_db(res.get("events") or [], connect_db_fn)
							ui.session_state["_cal_flash_msg"] = f"✨ Gemini AI ({used_m}) ile {TR_MONTH_NAMES[view_month]} {view_year} için {added} yeni özel gün ve meslek günü takvime işlendi!"
							ui.rerun()
						else:
							ui.error(f"Gemini hatası: {res.get('error')}")
				else:
					ui.caption("💡 Gemini AI açık olduğunda aya özel ek ulusal/uluslararası sağlık ve meslek günlerini yapay zekayla da tarayabilirsiniz.")

		with wiz_t2:
			ui.caption("Takvimde işaretlenmesini istediğiniz **ÖSYM** ve **MEB** merkezi sınavlarını seçin:")
			exam_choices = ["YKS (TYT-AYT-YDT)", "LGS", "KPSS", "MSÜ", "ALES", "DGS", "YDS / YÖKDİL", "İOKBS Bursluluk"]
			picked_exams = ui.multiselect(
				"İşaretlenecek Sınavlar",
				exam_choices,
				default=["YKS (TYT-AYT-YDT)", "LGS", "KPSS", "MSÜ"],
				key="cal_multiselect_exams",
			)
			ex_c1, ex_c2 = ui.columns(2)
			with ex_c1:
				if ui.button(f"📝 Seçili Sınavları {view_year} Takvimine İşle (Hızlı)", key="btn_seed_builtin_exams", type="primary", use_container_width=True):
					cnt = seed_selected_exams_to_db(view_year, picked_exams, connect_db_fn)
					ui.session_state["_cal_flash_msg"] = f"✅ Seçilen sınavlardan {cnt} oturum {view_year} takvimine 📝 sembolüyle eklendi."
					ui.rerun()
			with ex_c2:
				if ai_enabled and ai_key:
					if ui.button(f"✨ Gemini AI ile {view_year} Güncel ÖSYM Sınavlarını Sorgula & İşle", key="btn_ai_osym_exams", type="secondary", use_container_width=True):
						ok, res, used_m = generate_calendar_events_with_ai(
							api_key=ai_key,
							model=ai_model,
							year=view_year,
							month=None,
							preset_type="osym_exams",
							selected_exams=picked_exams,
						)
						if ok:
							added = save_ai_generated_events_to_db(res.get("events") or [], connect_db_fn)
							ui.session_state["_cal_flash_msg"] = f"✨ Gemini AI ({used_m}) ile {added} ÖSYM/MEB sınav tarihi takvime eklendi."
							ui.rerun()
						else:
							ui.error(f"Gemini hatası: {res.get('error')}")

		with wiz_t3:
			if not (ai_enabled and ai_key):
				ui.info("🤖 Serbest doğal dil ile takvim işaretlemek için **Ayarlar → Yapay Zekâ (Gemini)** bölümünden Gemini AI'yı aktif hale getirebilirsiniz. Manuel etkinlik ve tatil ekleme formu sağdaki **Takvim Yaprağı** altında her zaman aktiftir.")
			else:
				ai_cal_prompt = ui.text_area(
					"Gemini AI'ya Takvim Talimatı Verin",
					placeholder=(
						f"Örn: {TR_MONTH_NAMES[view_month]} {view_year} içindeki MEB ara tatilini, seçim ertesi idari izin gününü, "
						"okul genel göz ve boy-kilo tarama haftasını veya istediğiniz özel tatil/sınav günlerini yazın..."
					),
					height=78,
					key="cal_ai_free_prompt",
				)
				if ui.button("✨ Gemini AI ile Analiz Et ve Takvime İşle", key="btn_run_ai_free_cal", type="primary"):
					if not ai_cal_prompt.strip():
						ui.warning("Lütfen eklenecek tatil, sınav veya etkinlik talimatını yazın.")
					else:
						ok, res, used_m = generate_calendar_events_with_ai(
							api_key=ai_key,
							model=ai_model,
							year=view_year,
							month=view_month,
							preset_type="election_or_custom",
							custom_prompt=ai_cal_prompt.strip(),
						)
						if ok:
							added = save_ai_generated_events_to_db(res.get("events") or [], connect_db_fn)
							summary_txt = res.get("summary") or f"{added} kayıt oluşturuldu."
							ui.session_state["_cal_flash_msg"] = f"✨ Gemini AI ({used_m}): {summary_txt} ({added} yeni takvim kaydı eklendi)."
							ui.rerun()
						else:
							ui.error(f"Gemini hatası: {res.get('error')}")

	# -----------------------------------------------------------------------
	# GÖRÜNTÜLENEN AYIN VERİLERİNİ VERİTABANINDAN ÇEKME VE ANALİZ
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

		# Aynı aydaki klinik izlem / aşı / tarama sayılarını da bilgi amaçlı çekelim
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

	# Ayın günlerini ve mesai istatistiklerini hesapla
	_, days_in_month = calendar.monthrange(view_year, view_month)
	workdays_count = 0.0
	weekend_days_count = 0
	civil_holiday_count = 0
	half_day_count = 0
	total_events_count = len(month_rows)

	for d_num in range(1, days_in_month + 1):
		d_obj = date(view_year, view_month, d_num)
		d_iso = d_obj.isoformat()
		day_evs = events_by_date.get(d_iso, [])
		is_wknd = d_obj.weekday() >= 5
		has_full_hol = any(bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")) for e in day_evs)
		has_half_hol = any(bool(e.get("is_half_day")) for e in day_evs)

		if has_full_hol:
			civil_holiday_count += 1
		elif has_half_hol:
			half_day_count += 1

		if is_wknd:
			weekend_days_count += 1
		else:
			if has_full_hol:
				pass
			elif has_half_hol:
				workdays_count += 0.5
			else:
				workdays_count += 1.0

	workdays_str = f"{int(workdays_count)}" if workdays_count.is_integer() else f"{workdays_count:.1f}".replace(".", ",")

	# Ay Özeti & Renk / Sembol Rehberi Şeridi
	ui.markdown(
		f"""<div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;
		padding:10px 14px;border-radius:12px;background:var(--bg-card);border:1px solid var(--line);margin:6px 0 12px 0;">
			<div style="display:flex;flex-wrap:wrap;align-items:center;gap:14px;">
				<span style="font-size:15px;font-weight:800;color:var(--text-heading);">🗓️ {TR_MONTH_NAMES[view_month]} {view_year}</span>
				<span style="font-size:12.5px;background:#ecfdf5;color:#047857;border:1px solid #a7f3d0;padding:3px 9px;border-radius:999px;font-weight:700;">
					🏢 Net İş Günü: {workdays_str} gün
				</span>
				<span style="font-size:12.5px;background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;padding:3px 9px;border-radius:999px;font-weight:700;">
					🏖️ Hafta Sonu: {weekend_days_count} gün
				</span>
				<span style="font-size:12.5px;background:#fef2f2;color:#dc2626;border:1px solid #fecaca;padding:3px 9px;border-radius:999px;font-weight:800;">
					🇹🇷 Resmi Tatil: {civil_holiday_count} tam{f' · {half_day_count} yarım gün' if half_day_count else ''}
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

	# -----------------------------------------------------------------------
	# İKİ SÜTUNLU ANA YERLEŞİM: SOL = AYLIK TAKVİM, SAĞ = TAKVİM YAPRAĞI
	# -----------------------------------------------------------------------
	cal_col, leaf_col = ui.columns([1.68, 1.05], gap="medium")

	with cal_col:
		# Haftanın Günleri Başlığı (Pzt - Paz)
		hdr_cols = ui.columns(7)
		for w_idx, w_short in enumerate(TR_WEEKDAY_SHORT):
			is_wk = w_idx >= 5
			hdr_bg = "#fee2e2" if w_idx == 6 else ("#e2e8f0" if is_wk else "var(--bg-card)")
			hdr_fg = "#dc2626" if w_idx == 6 else ("#334155" if is_wk else "var(--text-heading)")
			with hdr_cols[w_idx]:
				ui.markdown(
					f"""<div style="text-align:center;padding:6px 4px;border-radius:8px;background:{hdr_bg};
					border:1px solid var(--line);font-size:12.5px;font-weight:800;color:{hdr_fg};margin-bottom:4px;">
						{w_short}
					</div>""",
					unsafe_allow_html=True,
				)

		cal_matrix = calendar.Calendar(firstweekday=0).monthdayscalendar(view_year, view_month)
		for week_idx, week_days in enumerate(cal_matrix):
			w_cols = ui.columns(7)
			for wd_idx, day_num in enumerate(week_days):
				with w_cols[wd_idx]:
					if day_num == 0:
						ui.markdown(
							'<div style="min-height:112px;border-radius:10px;background:rgba(148,163,184,0.06);'
							'border:1px dashed rgba(148,163,184,0.22);margin-bottom:4px;"></div>',
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
					has_prof = any(e.get("symbol_type") == "profession" or e.get("category") == "meslek_gunu" for e in cell_events) or bool(sp_builtin)

					# Renk Planı:
					# 1. Devlet Memuru Resmi Tatili -> Kırmızı Vurgulu (#fef2f2 zemin, #dc2626 kırmızı yazı)
					# 2. Yarım Gün Mesai -> Turuncu-Kırmızı (#fff7ed zemin, #ea580c / #dc2626)
					# 3. Hafta Sonu -> Belirgin Arduvaz/Mavi-Gri (#f1f5f9 zemin)
					# 4. İş Günü -> Beyaz / Temiz Kart Zemini
					if has_full_holiday:
						cell_bg = "linear-gradient(180deg, #fef2f2 0%, #fee2e2 100%)"
						cell_border = "2px solid #dc2626" if is_selected else "1.5px solid #fca5a5"
						num_color = "#dc2626"
						status_tag = '<span style="font-size:9.5px;font-weight:800;color:#ffffff;background:#dc2626;padding:1px 5px;border-radius:4px;">TATİL</span>'
					elif has_half_day:
						cell_bg = "linear-gradient(180deg, #fff7ed 0%, #ffedd5 100%)"
						cell_border = "2px solid #ea580c" if is_selected else "1.5px solid #fdba74"
						num_color = "#dc2626"
						status_tag = '<span style="font-size:9.5px;font-weight:800;color:#9a3412;background:#fed7aa;padding:1px 5px;border-radius:4px;">½ YARIM GÜN</span>'
					elif has_ballot:
						cell_bg = "linear-gradient(180deg, #f5f3ff 0%, #ede9fe 100%)"
						cell_border = "2px solid #7c3aed" if is_selected else "1.5px solid #c4b5fd"
						num_color = "#6d28d9"
						status_tag = '<span style="font-size:9.5px;font-weight:800;color:#5b21b6;background:#ddd6fe;padding:1px 5px;border-radius:4px;">İDARİ İZİN</span>'
					elif is_weekend:
						cell_bg = "linear-gradient(180deg, #f1f5f9 0%, #e2e8f0 100%)"
						cell_border = "2px solid var(--teal)" if is_selected else "1px solid #cbd5e1"
						num_color = "#dc2626" if is_sunday else "#334155"
						status_tag = '<span style="font-size:9px;font-weight:700;color:#64748b;">H.Sonu</span>'
					else:
						cell_bg = "var(--bg-card)"
						cell_border = "2px solid var(--teal)" if is_selected else "1px solid var(--line)"
						num_color = "var(--text-heading)"
						status_tag = '<span style="font-size:9px;font-weight:600;color:#059669;">İş Günü</span>'

					if is_today and not is_selected:
						cell_border = "2px solid #0284c7"

					# Üst sağ sembol dizisi (Bayrak, Hilal, Sınav, Meslek)
					top_symbols: list[str] = []
					if has_tr_flag:
						top_symbols.append(SVG_TR_FLAG_MINI)
					if has_crescent:
						top_symbols.append(SVG_CRESCENT_MINI)
					if has_exam:
						top_symbols.append('<span title="ÖSYM / MEB Sınavı" style="font-size:12px;">📝</span>')
					if has_prof and not has_tr_flag:
						top_symbols.append('<span title="Meslek / Belirli Gün" style="font-size:11.5px;">👩‍⚕️</span>')
					symbols_html = "".join(top_symbols)

					# Hücre içi etkinlik/tatil başlık etiketleri (en fazla 2 satır + taşma sayısı)
					pills_html_list: list[str] = []
					for ev in cell_events[:2]:
						ev_title = html.escape(str(ev.get("title") or ""))
						ev_hol = bool(ev.get("is_civil_holiday"))
						ev_cat = str(ev.get("category") or "belirli_gun")
						if ev_hol:
							p_bg, p_fg, p_bd = "#fee2e2", "#b91c1c", "#fca5a5"
						elif ev_cat == "dini_bayram":
							p_bg, p_fg, p_bd = "#fef3c7", "#b45309", "#fde68a"
						elif ev_cat == "osym_sinav":
							p_bg, p_fg, p_bd = "#dbeafe", "#1d4ed8", "#93c5fd"
						elif ev_cat == "revir_etkinlik":
							p_bg, p_fg, p_bd = "#d1fae5", "#047857", "#6ee7b7"
						else:
							p_bg, p_fg, p_bd = "#f1f5f9", "#334155", "#cbd5e1"
						pills_html_list.append(
							f'<div style="font-size:10px;font-weight:700;color:{p_fg};background:{p_bg};'
							f'border:1px solid {p_bd};border-radius:4px;padding:1.5px 4px;margin-top:2px;'
							f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{ev_title}">'
							f'{ev_title}</div>'
						)

					# Eğer veritabanında kayıt yoksa ama yerleşik meslek/belirli gün varsa onu hafif etiketle göster
					if not cell_events and sp_builtin:
						sp_t = html.escape(sp_builtin["title"])
						pills_html_list.append(
							f'<div style="font-size:9.5px;font-weight:600;color:#0f766e;background:#f0fdfa;'
							f'border:1px solid #99f6e4;border-radius:4px;padding:1px 4px;margin-top:2px;'
							f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="{sp_t}">'
							f'{sp_t}</div>'
						)

					if len(cell_events) > 2:
						pills_html_list.append(
							f'<div style="font-size:9.5px;font-weight:800;color:#475569;margin-top:1px;">+{len(cell_events) - 2} etkinlik daha</div>'
						)
					elif revir_v_cnt > 0:
						pills_html_list.append(
							f'<div style="font-size:9.5px;font-weight:700;color:#0369a1;margin-top:1px;">🩺 {revir_v_cnt} revir kaydı</div>'
						)

					pills_joined = "".join(pills_html_list)
					today_dot = (
						'<span style="background:#0284c7;color:#fff;font-size:8.5px;font-weight:800;padding:1px 4px;border-radius:4px;margin-left:3px;">BUGÜN</span>'
						if is_today else ""
					)
					sel_shadow = "0 0 0 2px rgba(13,148,136,0.22)" if is_selected else "0 1px 2px rgba(0,0,0,0.04)"

					ui.markdown(
						f"""<div style="min-height:86px;padding:6px 7px;border-radius:10px 10px 4px 4px;background:{cell_bg};
						border:{cell_border};box-shadow:{sel_shadow};display:flex;flex-direction:column;justify-content:space-between;">
							<div>
								<div style="display:flex;align-items:center;justify-content:space-between;gap:2px;">
									<span style="font-size:15.5px;font-weight:900;color:{num_color};line-height:1.1;">{day_num}{today_dot}</span>
									<span style="display:flex;align-items:center;gap:3px;">{symbols_html}</span>
								</div>
								<div style="margin-top:2px;">{status_tag}</div>
								<div style="margin-top:3px;">{pills_joined}</div>
							</div>
						</div>""",
						unsafe_allow_html=True,
					)
					btn_lbl = f"✓ {day_num} {TR_MONTH_NAMES[view_month][:3]}" if is_selected else f"{day_num} {TR_MONTH_NAMES[view_month][:3]}"
					if ui.button(
						btn_lbl,
						key=f"btn_cal_day_{cell_iso}",
						type=("primary" if is_selected else "secondary"),
						use_container_width=True,
					):
						ui.session_state["cal_selected_date"] = cell_iso
						ui.rerun()

	# -----------------------------------------------------------------------
	# SAĞ PANEL: SEÇİLİ GÜNE AİT İNTERAKTİF "TAKVİM YAPRAĞI"
	# -----------------------------------------------------------------------
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
# SAĞ PANEL: TAKVİM YAPRAĞI DETAY BİLEŞENİ
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
	sel_iso = sel_date.isoformat()
	weekday_idx = sel_date.weekday()
	weekday_name = TR_WEEKDAY_NAMES[weekday_idx]
	month_name = TR_MONTH_NAMES[sel_date.month]
	hijri_str = gregorian_to_hijri_str(sel_date)
	day_of_year = sel_date.timetuple().tm_yday
	days_in_year = 366 if calendar.isleap(sel_date.year) else 365
	remaining_days = days_in_year - day_of_year

	# Seçili günün veritabanı kayıtları ve günlük yaprak notu
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
			          day_significance, custom_note
			   FROM calendar_daily_notes
			   WHERE note_date = ?""",
			(sel_iso,),
		).fetchone()

	sp_builtin = PROFESSION_AND_SPECIAL_DAYS.get((sel_date.month, sel_date.day))
	has_full_hol = any(bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")) for e in day_events)
	has_half_hol = any(bool(e.get("is_half_day")) for e in day_events)
	is_weekend = weekday_idx >= 5
	is_sunday = weekday_idx == 6

	if has_full_hol:
		header_accent = "#dc2626"
		num_color = "#dc2626"
		work_status_badge = "🇹🇷 RESMİ TATİL (DEVLET MEMURLARI İZİNLİ)"
		work_status_bg = "#fee2e2"
		work_status_fg = "#b91c1c"
	elif has_half_hol:
		header_accent = "#ea580c"
		num_color = "#dc2626"
		work_status_badge = "½ YARIM GÜN MESAİ (13:00'TEN SONRA RESMİ TATİL)"
		work_status_bg = "#ffedd5"
		work_status_fg = "#c2410c"
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

	# Günün Meslek Grubu Tespiti
	prof_groups: list[str] = []
	for e in day_events:
		pg = str(e.get("profession_group") or "").strip()
		if pg and pg not in prof_groups:
			prof_groups.append(pg)
	if sp_builtin and sp_builtin.get("prof") and sp_builtin["prof"] not in prof_groups:
		prof_groups.append(sp_builtin["prof"])

	default_wd_prof, default_wd_desc = WEEKDAY_DEFAULT_PROFESSION_THEME[weekday_idx]
	if prof_groups:
		active_prof_title = " · ".join(prof_groups)
		active_prof_sub = (
			sp_builtin["title"]
			if sp_builtin
			else next((e["title"] for e in day_events if e.get("profession_group")), "Günün Özel Meslek Grubu")
		)
	else:
		active_prof_title = default_wd_prof
		active_prof_sub = default_wd_desc

	# Günün Tatil / Belirli Gün Özeti
	special_badges_html: list[str] = []
	for e in day_events:
		sym_html = render_symbol_badge(e.get("symbol_type") or "event")
		e_ttl = html.escape(str(e.get("title") or ""))
		e_desc = html.escape(str(e.get("description") or ""))
		is_h = bool(e.get("is_civil_holiday"))
		b_bg = "#fef2f2" if is_h else "#f8fafc"
		b_bd = "#fca5a5" if is_h else "#cbd5e1"
		b_fg = "#b91c1c" if is_h else "#1e293b"
		special_badges_html.append(
			f"""<div style="padding:7px 10px;border-radius:8px;background:{b_bg};border:1px solid {b_bd};margin-bottom:6px;">
				<div style="font-size:12.5px;font-weight:800;color:{b_fg};display:flex;align-items:center;gap:6px;">
					{sym_html} <span>{e_ttl}</span>
				</div>
				{f'<div style="font-size:11.2px;color:#475569;margin-top:2px;">{e_desc}</div>' if e_desc else ''}
			</div>"""
		)
	if not day_events and sp_builtin:
		sym_html = render_symbol_badge("tr_flag" if sp_builtin.get("cat") == "milli_bayram" else "profession")
		special_badges_html.append(
			f"""<div style="padding:7px 10px;border-radius:8px;background:#f0fdfa;border:1px solid #99f6e4;margin-bottom:6px;">
				<div style="font-size:12.5px;font-weight:800;color:#0f766e;display:flex;align-items:center;gap:6px;">
					{sym_html} <span>{html.escape(sp_builtin['title'])}</span>
				</div>
			</div>"""
		)

	special_block = (
		"".join(special_badges_html)
		if special_badges_html
		else '<div style="font-size:12px;color:#64748b;font-style:italic;">Bu güne kayıtlı resmi tatil veya sınav bulunmuyor. Aşağıdan yeni etkinlik ekleyebilirsiniz.</div>'
	)

	# 1. KLASİK TAKVİM YAPRAĞI KARTI
	ui.markdown(
		f"""<div style="border-radius:14px;overflow:hidden;border:2px solid {header_accent};background:#fffdf9;
		box-shadow:0 8px 22px rgba(15,23,42,0.08);margin-bottom:10px;color:#1e293b;">
			<div style="background:{header_accent};color:#ffffff;padding:8px 14px;display:flex;align-items:center;justify-content:space-between;">
				<span style="font-size:12px;font-weight:800;letter-spacing:0.6px;">🗓️ GÜNLÜK TAKVİM YAPRAĞI</span>
				<span style="font-size:11.5px;font-weight:700;opacity:0.95;">🌙 {html.escape(hijri_str)}</span>
			</div>
			<div style="padding:12px 16px;">
				<div style="display:flex;align-items:center;justify-content:space-between;border-bottom:1.5px dashed #cbd5e1;padding-bottom:10px;">
					<div>
						<div style="font-size:13px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.8px;">
							{html.escape(month_name)} {sel_date.year}
						</div>
						<div style="font-size:46px;font-weight:900;color:{num_color};line-height:1.0;margin:2px 0;">
							{sel_date.day:02d}
						</div>
						<div style="font-size:16px;font-weight:800;color:{num_color};">
							{html.escape(weekday_name)}
						</div>
					</div>
					<div style="text-align:right;">
						<div style="display:inline-block;padding:4px 10px;border-radius:999px;background:{work_status_bg};color:{work_status_fg};font-size:11px;font-weight:900;margin-bottom:6px;">
							{work_status_badge}
						</div>
						<div style="font-size:11.5px;color:#475569;font-weight:700;">
							Yılın <b>{day_of_year}.</b> Günü · Kalan: <b>{remaining_days}</b>
						</div>
						<div style="font-size:11px;color:#64748b;margin-top:2px;">
							Miladi: {sel_date.strftime('%d.%m.%Y')}
						</div>
					</div>
				</div>

				<div style="margin-top:10px;padding:8px 11px;border-radius:9px;background:#f0fdfa;border:1px solid #99f6e4;">
					<div style="font-size:10.5px;font-weight:800;color:#0f766e;text-transform:uppercase;letter-spacing:0.5px;">
						👩‍⚕️ GÜNÜN MESLEK GRUBU & TOPLULUK TAKDİRİ
					</div>
					<div style="font-size:13px;font-weight:800;color:#134e4a;margin-top:2px;">
						{html.escape(active_prof_title)}
					</div>
					<div style="font-size:11.3px;color:#0f766e;margin-top:1px;">
						{html.escape(active_prof_sub)}
					</div>
				</div>

				<div style="margin-top:10px;">
					<div style="font-size:10.5px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;">
						📌 GÜNÜN TATİL, SINAV VE BELİRLİ GÜNLERİ
					</div>
					{special_block}
				</div>
			</div>
		</div>""",
		unsafe_allow_html=True,
	)

	# -----------------------------------------------------------------------
	# 2. SEÇİLİ İLE GÖRE EZAN VAKİTLERİ VE CANLI GERİ SAYIM SAYACI
	# -----------------------------------------------------------------------
	show_prayer = str(cfg.get("calendar_show_prayer_times", "1")).strip() != "0"
	if show_prayer:
		default_prov = (cfg.get("default_province") or "Ankara").strip() or "Ankara"
		default_dist = (cfg.get("default_district") or "").strip()
		all_provs = get_tr_provinces_fn() or sorted(TR_PROVINCE_COORDS.keys())
		if default_prov not in all_provs:
			default_prov = "Ankara"

		with ui.container(border=True):
			pr_hdr1, pr_hdr2 = ui.columns([1.8, 1.2])
			with pr_hdr1:
				loc_disp = f"{default_prov} / {default_dist}" if default_dist else default_prov
				ui.markdown(f"#### 🕌 Ezan Vakitleri ({html.escape(loc_disp)})")
			with pr_hdr2:
				picked_prov = ui.selectbox(
					"İl Seçimi",
					all_provs,
					index=all_provs.index(default_prov),
					key="cal_leaf_province_picker",
					label_visibility="collapsed",
				)
				if picked_prov != default_prov:
					save_app_settings_fn({"default_province": picked_prov})
					ui.rerun()

			ptimes = get_prayer_times_for_province(default_prov, sel_date, connect_db_fn)
			vakit_keys = ["İmsak", "Güneş", "Öğle", "İkindi", "Akşam", "Yatsı"]
			now_dt = datetime.now()
			now_hm = now_dt.strftime("%H:%M")

			# Aktif ve sıradaki vakti bul
			next_vakit_name = "İmsak"
			next_vakit_hm = ptimes.get("İmsak", "05:30")
			current_vakit_name = "Yatsı"
			for idx_v, vk in enumerate(vakit_keys):
				vhm = ptimes.get(vk, "00:00")
				if now_hm < vhm:
					next_vakit_name = vk
					next_vakit_hm = vhm
					current_vakit_name = vakit_keys[idx_v - 1] if idx_v > 0 else "Yatsı"
					break
			else:
				# Yatsı geçtiyse yarının İmsak vakti
				tmr_pt = get_prayer_times_for_province(default_prov, date.today() + timedelta(days=1), connect_db_fn)
				next_vakit_name = "İmsak (Yarın)"
				next_vakit_hm = tmr_pt.get("İmsak", ptimes.get("İmsak", "05:30"))
				current_vakit_name = "Yatsı"

			# Canlı JavaScript Geri Sayım Sayacı için tüm vakitlerin listesi
			vakit_json_list = json.dumps(
				[{"name": vk, "time": ptimes.get(vk, "00:00")} for vk in vakit_keys],
				ensure_ascii=False,
			)

			v_cards_html: list[str] = []
			for vk in vakit_keys:
				vhm = ptimes.get(vk, "--:--")
				is_cur_v = (vk == current_vakit_name and sel_date == date.today())
				vc_bg = "#0f766e" if is_cur_v else "var(--bg-card)"
				vc_fg = "#ffffff" if is_cur_v else "var(--text-heading)"
				vc_sub = "#ccfbf1" if is_cur_v else "var(--text-muted)"
				vc_bd = "2px solid #0d9488" if is_cur_v else "1px solid var(--line)"
				v_cards_html.append(
					f"""<div style="text-align:center;padding:6px 4px;border-radius:8px;background:{vc_bg};border:{vc_bd};">
						<div style="font-size:10.5px;font-weight:700;color:{vc_sub};">{vk}</div>
						<div style="font-size:13.5px;font-weight:900;color:{vc_fg};margin-top:1px;">{vhm}</div>
					</div>"""
				)

			# Kalan süreyi Python tarafında da başlangıç değeri olarak hesapla (ilk render anında anında görünsün)
			try:
				nh, nm = [int(x) for x in next_vakit_hm.split(":")[:2]]
				target_dt = now_dt.replace(hour=nh, minute=nm, second=0, microsecond=0)
				if target_dt <= now_dt:
					target_dt += timedelta(days=1)
				diff_sec = max(0, int((target_dt - now_dt).total_seconds()))
				init_h = diff_sec // 3600
				init_m = (diff_sec % 3600) // 60
				init_s = diff_sec % 60
				init_countdown_txt = f"{init_h:02d} saat {init_m:02d} dk {init_s:02d} sn"
			except Exception:
				init_countdown_txt = "-- saat -- dk -- sn"

			ui.markdown(
				f"""<div style="display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:5px;margin:4px 0 8px 0;">
					{''.join(v_cards_html)}
				</div>
				<div style="padding:8px 12px;border-radius:9px;background:linear-gradient(90deg,#0f766e 0%,#0284c7 100%);
				color:#ffffff;display:flex;align-items:center;justify-content:space-between;font-size:12.5px;font-weight:800;">
					<span>⏳ <span id="revir_next_prayer_label">{html.escape(next_vakit_name)} ({html.escape(next_vakit_hm)})</span> vaktine kalan:</span>
					<span id="revir_prayer_countdown" style="background:rgba(255,255,255,0.18);padding:2px 9px;border-radius:6px;font-family:monospace;font-size:13.5px;letter-spacing:0.4px;">
						{init_countdown_txt}
					</span>
				</div>
				<div style="font-size:10.5px;color:var(--text-muted);margin-top:4px;text-align:right;">
					Kaynak: {html.escape(ptimes.get('source', 'Diyanet Takvimi'))} · Ayarlar → Varsayılan Konum ile senkronize
				</div>
				<script>
				(function() {{
					var schedule = {vakit_json_list};
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
							var p0 = schedule[0].time.split(':');
							target = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, parseInt(p0[0], 10), parseInt(p0[1], 10), 0);
							targetName = schedule[0].name + ' (Yarın)';
							targetHM = schedule[0].time;
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
				</script>""",
				unsafe_allow_html=True,
			)

	# -----------------------------------------------------------------------
	# 3. GÜNÜN HADİS-İ ŞERİFİ / ÖZLÜ SÖZÜ (AYARLARDAN YÖNETİLEN & GEMINI DESTEKLİ)
	# -----------------------------------------------------------------------
	quote_mode = cfg.get("calendar_quote_mode", QUOTE_MODE_OPTIONS[0])
	if quote_mode != "Kapalı (Gösterme)":
		builtin_qh = get_builtin_daily_quote_and_hadith(sel_date)
		hadith_txt = (leaf_note_row["hadith_text"] if leaf_note_row and leaf_note_row["hadith_text"] else "") or builtin_qh["hadith_text"]
		hadith_src = (leaf_note_row["hadith_source"] if leaf_note_row and leaf_note_row["hadith_source"] else "") or builtin_qh["hadith_source"]
		quote_txt = (leaf_note_row["quote_text"] if leaf_note_row and leaf_note_row["quote_text"] else "") or builtin_qh["quote_text"]
		quote_src = (leaf_note_row["quote_source"] if leaf_note_row and leaf_note_row["quote_source"] else "") or builtin_qh["quote_author"]
		leaf_back_txt = (leaf_note_row["custom_note"] if leaf_note_row and leaf_note_row["custom_note"] else "")

		with ui.container(border=True):
			ui.markdown("#### 📜 Takvim Yaprağı Arkası: Günün Sözü & Hadis-i Şerif")
			show_hadith = ("Hadis" in quote_mode) or ("Karışık" in quote_mode)
			show_quote = ("Özlü Söz" in quote_mode) or ("Karışık" in quote_mode)

			if show_hadith:
				ui.markdown(
					f"""<div style="padding:9px 12px;border-radius:9px;background:#fffbeb;border-left:4px solid #d97706;margin-bottom:7px;color:#78350f;">
						<div style="font-size:10.5px;font-weight:800;color:#b45309;text-transform:uppercase;">🌙 GÜNÜN HADİS-İ ŞERİFİ</div>
						<div style="font-size:12.8px;font-style:italic;font-weight:600;margin-top:3px;">“{html.escape(hadith_txt)}”</div>
						<div style="font-size:11px;font-weight:800;color:#92400e;margin-top:3px;text-align:right;">— {html.escape(hadith_src)}</div>
					</div>""",
					unsafe_allow_html=True,
				)
			if show_quote:
				ui.markdown(
					f"""<div style="padding:9px 12px;border-radius:9px;background:#f0fdf4;border-left:4px solid #059669;margin-bottom:7px;color:#064e3b;">
						<div style="font-size:10.5px;font-weight:800;color:#047857;text-transform:uppercase;">💡 GÜNÜN ÖZLÜ SÖZÜ / VECİZESİ</div>
						<div style="font-size:12.8px;font-style:italic;font-weight:600;margin-top:3px;">“{html.escape(quote_txt)}”</div>
						<div style="font-size:11px;font-weight:800;color:#065f46;margin-top:3px;text-align:right;">— {html.escape(quote_src)}</div>
					</div>""",
					unsafe_allow_html=True,
				)
			if leaf_back_txt:
				ui.markdown(
					f"""<div style="padding:8px 11px;border-radius:8px;background:var(--bg-card);border:1px solid var(--line);font-size:12px;color:var(--text-body);">
						<b>📖 Takvim Yaprağı Notu:</b> {html.escape(leaf_back_txt)}
					</div>""",
					unsafe_allow_html=True,
				)

			qm_c1, qm_c2 = ui.columns([1.45, 1.35])
			with qm_c1:
				new_q_mode = ui.selectbox(
					"Gösterim Tercihi (Ayarlar)",
					QUOTE_MODE_OPTIONS,
					index=(QUOTE_MODE_OPTIONS.index(quote_mode) if quote_mode in QUOTE_MODE_OPTIONS else 0),
					key="cal_leaf_quote_mode_sel",
					label_visibility="collapsed",
				)
				if new_q_mode != quote_mode:
					save_app_settings_fn({"calendar_quote_mode": new_q_mode})
					ui.rerun()
			with qm_c2:
				if ai_enabled and ai_key:
					if ui.button("✨ Gemini ile Bu Güne Özel Üret", key=f"btn_ai_leaf_{sel_iso}", type="secondary", use_container_width=True):
						day_ctx = ", ".join([e["title"] for e in day_events]) or (sp_builtin["title"] if sp_builtin else "")
						ok, res, used_m = generate_daily_leaf_content_with_ai(
							api_key=ai_key,
							model=ai_model,
							target_date=sel_iso,
							quote_mode=quote_mode,
							day_context=day_ctx,
							province=cfg.get("default_province", "Ankara"),
						)
						if ok:
							with connect_db_fn() as conn:
								conn.execute(
									"""INSERT INTO calendar_daily_notes
									   (note_date, quote_text, quote_source, hadith_text, hadith_source,
									    day_significance, custom_note, updated_at)
									   VALUES (?, ?, ?, ?, ?, ?, ?, ?)
									   ON CONFLICT(note_date) DO UPDATE SET
									     quote_text = excluded.quote_text,
									     quote_source = excluded.quote_source,
									     hadith_text = excluded.hadith_text,
									     hadith_source = excluded.hadith_source,
									     day_significance = excluded.day_significance,
									     custom_note = excluded.custom_note,
									     updated_at = excluded.updated_at""",
									(
										sel_iso,
										str(res.get("quote_text") or "").strip(),
										str(res.get("quote_author") or "").strip(),
										str(res.get("hadith_text") or "").strip(),
										str(res.get("hadith_source") or "").strip(),
										str(res.get("day_significance") or "").strip(),
										str(res.get("leaf_back_note") or "").strip(),
										datetime.now().isoformat(timespec="seconds"),
									),
								)
							ui.session_state["_cal_flash_msg"] = f"✨ {sel_date.strftime('%d.%m.%Y')} takvim yaprağı yazısı Gemini AI ({used_m}) ile güncellendi."
							ui.rerun()
						else:
							ui.error(f"Gemini hatası: {res.get('error')}")

	# -----------------------------------------------------------------------
	# 4. SEÇİLİ GÜNÜN ETKİNLİKLERİ & YENİ ETKİNLİK / TATİL OLUŞTURMA FORMU
	# -----------------------------------------------------------------------
	with ui.container(border=True):
		ui.markdown(f"#### 📌 {sel_date.strftime('%d')} {month_name} {sel_date.year} — Etkinlik & Plan Yönetimi")
		if day_events:
			for ev in day_events:
				ev_id = int(ev["id"])
				ev_done = bool(ev.get("is_completed"))
				t_range = " - ".join(filter(None, [ev.get("start_time"), ev.get("end_time")]))
				t_badge = f"`{t_range}` · " if t_range else ""
				strike_open = "~~" if ev_done else ""
				strike_close = "~~" if ev_done else ""
				row_c1, row_c2, row_c3 = ui.columns([2.6, 0.55, 0.55])
				with row_c1:
					ui.markdown(f"{t_badge}{strike_open}**{ev['title']}**{strike_close}")
					if ev.get("description"):
						ui.caption(ev["description"])
				with row_c2:
					done_lbl = "↩️" if ev_done else "✓"
					if ui.button(done_lbl, key=f"btn_cal_done_{ev_id}", help="Tamamlandı durumunu değiştir", use_container_width=True):
						with connect_db_fn() as conn:
							conn.execute("UPDATE calendar_events SET is_completed = ? WHERE id = ?", (0 if ev_done else 1, ev_id))
						ui.rerun()
				with row_c3:
					if ui.button("🗑️", key=f"btn_cal_del_{ev_id}", help="Bu kaydı sil", use_container_width=True):
						with connect_db_fn() as conn:
							conn.execute("DELETE FROM calendar_events WHERE id = ?", (ev_id,))
						ui.session_state["_cal_flash_msg"] = f"🗑️ '{ev['title']}' takvimden silindi."
						ui.rerun()

		with ui.expander(f"➕ {sel_date.strftime('%d.%m.%Y')} Gününe Yeni Etkinlik / Tatil / Sınav Ekle", expanded=(len(day_events) == 0)):
			with ui.form(f"form_add_cal_event_{sel_iso}"):
				new_title = ui.text_input("Etkinlik / Tatil / Sınav Başlığı", placeholder="Örn: 3-A Sınıfı Boy-Kilo Taraması veya Seçim İdari İzni")
				fc1, fc2 = ui.columns(2)
				cat_keys = list(CATEGORY_LABELS.keys())
				cat_disp = [CATEGORY_LABELS[k][0] for k in cat_keys]
				with fc1:
					picked_cat_disp = ui.selectbox("Kategori", cat_disp, index=cat_keys.index("revir_etkinlik"))
					picked_cat = cat_keys[cat_disp.index(picked_cat_disp)]
				sym_keys = list(SYMBOL_OPTIONS.keys())
				sym_disp = [SYMBOL_OPTIONS[k] for k in sym_keys]
				default_sym = CATEGORY_LABELS[picked_cat][2]
				with fc2:
					picked_sym_disp = ui.selectbox("Takvim Sembolü", sym_disp, index=sym_keys.index(default_sym))
					picked_sym = sym_keys[sym_disp.index(picked_sym_disp)]

				tc1, tc2 = ui.columns(2)
				with tc1:
					new_start = ui.text_input("Başlangıç Saati", placeholder="09:00")
				with tc2:
					new_end = ui.text_input("Bitiş Saati", placeholder="12:30")

				hc1, hc2 = ui.columns(2)
				with hc1:
					new_is_hol = ui.checkbox("Devlet Memuru Resmi Tatili (Kırmızı Yaz)", value=(picked_cat in ("milli_bayram", "resmi_tatil", "yarim_gun")))
				with hc2:
					new_is_half = ui.checkbox("Yarım Gün Mesai (½ Rozeti)", value=(picked_cat == "yarim_gun"))

				new_prof = ui.text_input("İlgili Meslek Grubu (Opsiyonel)", placeholder="Örn: Okul Hemşireleri ve Öğretmenler")
				new_desc = ui.text_area("Açıklama / Not", placeholder="Günün anlamı veya planlanan etkinlik detayı...", height=65)

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
						ui.rerun()
