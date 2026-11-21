"""
Okul Revir Defteri — Pediatrik Diyabet & İnsülin Doz Yönetim Merkezi (insulin_diabetes.py)

Özellikler:
- Hekim reçeteli İKO (Sabah/Öğle/Ara Öğün), İDF, Hedef KŞ, Düzeltme Eşiği, Kalem Hassasiyeti ve Maksimum Bolus Sınırı
- Kesin deterministik pediatrik bolus matematiği + Aktif İnsülin (IOB / Stacking Koruması)
- Aylık Okul Yemek Programı yönetimi + Gemini AI (PDF / Fotoğraf / Excel / Word / Metin) otomatik karbonhidrat çıkarıcı
- Tabağın tüketim yüzdesi (%0–%150) + Türk Okul/Kantin hızlı ekstra yiyecek kataloğu + Gemini tabak fotoğrafı/metin karb tahminleyici
- Fiziksel aktivite / beden eğitimi indirim motoru + egzersiz öncesi KŞ/Keton güvenlik kalkanı
- 15-15 Hipoglisemi & Hiperglisemi/Keton acil müdahale protokolü
- Enjeksiyon bölgesi rotasyon takibi (lipohipertrofi önleme & engel/uzuv profili duyarlılığı)
- Öğrencinin yaşı, boyu, kilosu, WHO persantil/Z-skoru ve TDD/kg verileriyle entegre Gemini AI klinik yorum + Veli SMS taslağı
- Glisemik trend (TIR %70–180) analizi + Resmi MEB antetli Word (.docx) & Excel (.xlsx) aylık plan ve izlem çıktıları
"""
from __future__ import annotations

import html
import io
import json
import math
import re
import sqlite3
from datetime import date, datetime, time, timedelta
from typing import Any

import pandas as pd

import gemini_assistant as gai
import revir_ui as ui


# ==============================================================================
# 1. SABİTLER & KLİNİK REFERANS KATALOGLARI
# ==============================================================================

DIABETES_TYPE_OPTIONS: list[str] = [
	"Tip 1 Diyabet (T1DM)",
	"Tip 2 Diyabet (Pediatrik/Erişkin)",
	"MODY / Monogenik Diyabet",
	"Steroid / İkincil Diyabet",
]

INSULIN_BRAND_OPTIONS: list[str] = [
	"NovoRapid (İnsülin Aspart)",
	"Fiasp (Ultra Hızlı Aspart)",
	"Humalog (İnsülin Lispro)",
	"Apidra (İnsülin Glulisin)",
	"Lyumjev (Ultra Hızlı Lispro)",
	"Actrapid / Humulin R (Kısa Etkili)",
	"Diğer Hızlı Etkili İnsülin",
]

DELIVERY_METHOD_OPTIONS: list[str] = [
	"İnsülin Kalemi (0.5 Ü Yarım Üniteli — Pediatrik)",
	"İnsülin Kalemi (1.0 Ü Tam Üniteli)",
	"İnsülin Pompası (0.1 Ü Hassasiyet)",
	"İnsülin Pompası (0.05 Ü Hassasiyet)",
]

PEN_INCREMENT_MAP: dict[str, float] = {
	"İnsülin Kalemi (0.5 Ü Yarım Üniteli — Pediatrik)": 0.5,
	"İnsülin Kalemi (1.0 Ü Tam Üniteli)": 1.0,
	"İnsülin Pompası (0.1 Ü Hassasiyet)": 0.1,
	"İnsülin Pompası (0.05 Ü Hassasiyet)": 0.05,
}

MEAL_TYPE_OPTIONS: list[str] = [
	"Öğle Yemeği",
	"Kahvaltı",
	"İkindi / Ara Öğün",
	"Sadece Düzeltme (Öğünsüz)",
]

ACTIVITY_PRESETS: dict[str, dict[str, Any]] = {
	"Yok / Normal Sınıf Dersi": {"pct": 0.0, "desc": "Ek azaltma yok (%0)"},
	"Hafif Aktivite (Bahçe Oyunu / Yürüyüş, 20-30 dk)": {"pct": 15.0, "desc": "Öğün bolusundan -%15 azaltma"},
	"Orta Şiddet (Beden Eğitimi / Basketbol, 30-60 dk)": {"pct": 25.0, "desc": "Öğün bolusundan -%25 azaltma"},
	"Yoğun Aktivite (Koşu / Yüzme / Turnuva, >45 dk)": {"pct": 40.0, "desc": "Öğün bolusundan -%40 azaltma"},
	"Özel Yüzde Belirle": {"pct": 20.0, "desc": "Hekim planına göre özel azaltma %"},
}

INJECTION_SITE_OPTIONS: list[str] = [
	"Karın — Sağ",
	"Karın — Sol",
	"Üst Kol — Sağ",
	"Üst Kol — Sol",
	"Uyluk (Bacak) — Sağ",
	"Uyluk (Bacak) — Sol",
	"Kalça — Sağ",
	"Kalça — Sol",
	"İnsülin Pompası (İnfüzyon Seti)",
]

KETONE_OPTIONS: list[str] = [
	"Ölçülmedi",
	"Negatif (0 / <0.6 mmol/L)",
	"Eser / Sınırda (0.6–1.5 mmol/L)",
	"Pozitif (+) / Yüksek (>1.5 mmol/L — Acil!)",
]

# Türk Okul Kantini & Beslenme Çantası Hızlı Ekstra Yiyecek Kataloğu
QUICK_EXTRA_FOODS: list[dict[str, Any]] = [
	{"label": "🍞 1 Dilim Ekmek (15g)", "name": "1 Dilim Ekmek (25g)", "carbs_g": 15.0, "gi": "Yüksek"},
	{"label": "🥨 Yarım Simit (28g)", "name": "Yarım Susamlı Simit", "carbs_g": 28.0, "gi": "Yüksek"},
	{"label": "🥪 Yarım Kaşarlı Tost (30g)", "name": "Yarım Kaşarlı Tost", "carbs_g": 30.0, "gi": "Orta"},
	{"label": "🍎 1 Orta Elma (15g)", "name": "1 Orta Boy Elma", "carbs_g": 15.0, "gi": "Düşük"},
	{"label": "🍌 1 Orta Muz (23g)", "name": "1 Orta Boy Muz", "carbs_g": 23.0, "gi": "Orta"},
	{"label": "🥛 1 Kutu Süt 200ml (10g)", "name": "1 Kutu Sade Süt (200 ml)", "carbs_g": 10.0, "gi": "Düşük"},
	{"label": "🍫 Kakaolu Süt 200ml (22g)", "name": "1 Kutu Kakaolu Süt (200 ml)", "carbs_g": 22.0, "gi": "Orta"},
	{"label": "🥛 1 Kutu Ayran (4g)", "name": "1 Küçük Kutu Ayran (170 ml)", "carbs_g": 4.0, "gi": "Düşük"},
	{"label": "🧃 Meyve Suyu 200ml (22g)", "name": "1 Kutu Meyve Suyu (200 ml)", "carbs_g": 22.0, "gi": "Yüksek"},
	{"label": "🍪 4 Adet Bisküvi (18g)", "name": "4 Adet Sade/Tam Buğday Bisküvi", "carbs_g": 18.0, "gi": "Orta"},
	{"label": "🥨 1 Paket Çubuk Kraker (28g)", "name": "1 Küçük Paket Çubuk Kraker (40g)", "carbs_g": 28.0, "gi": "Yüksek"},
	{"label": "🧁 1 Dilim Kek / Poğaça (32g)", "name": "1 Adet Poğaça / Dilim Kek", "carbs_g": 32.0, "gi": "Yüksek"},
]

TR_MONTH_NAMES: dict[int, str] = {
	1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
	7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık",
}


# ==============================================================================
# 2. VERİTABANI ŞEMASI & CRUD YARDIMCILARI
# ==============================================================================

def ensure_diabetes_tables(connection: sqlite3.Connection) -> None:
	"""Diyabet profili, aylık yemek menüsü ve insülin bolus defteri tablolarını oluşturur."""
	connection.executescript(
		"""
		CREATE TABLE IF NOT EXISTS diabetes_profiles (
			student_id INTEGER PRIMARY KEY REFERENCES students(id) ON DELETE CASCADE,
			is_diabetic INTEGER NOT NULL DEFAULT 1,
			diabetes_type TEXT NOT NULL DEFAULT 'Tip 1 Diyabet (T1DM)',
			insulin_brand TEXT DEFAULT 'NovoRapid (İnsülin Aspart)',
			basal_insulin TEXT DEFAULT '',
			delivery_method TEXT DEFAULT 'İnsülin Kalemi (0.5 Ü Yarım Üniteli — Pediatrik)',
			pen_increment REAL NOT NULL DEFAULT 0.5,
			icr_breakfast REAL NOT NULL DEFAULT 15.0,
			icr_lunch REAL NOT NULL DEFAULT 15.0,
			icr_snack REAL NOT NULL DEFAULT 18.0,
			isf REAL NOT NULL DEFAULT 50.0,
			target_bg REAL NOT NULL DEFAULT 120.0,
			correction_threshold REAL NOT NULL DEFAULT 150.0,
			hypo_threshold REAL NOT NULL DEFAULT 70.0,
			hyper_ketone_threshold REAL NOT NULL DEFAULT 250.0,
			insulin_duration_hours REAL NOT NULL DEFAULT 3.5,
			tdd_units REAL DEFAULT 0.0,
			max_single_bolus REAL NOT NULL DEFAULT 12.0,
			cgm_device TEXT DEFAULT '',
			glucagon_location TEXT DEFAULT '',
			physician_name TEXT DEFAULT '',
			emergency_protocol TEXT DEFAULT '',
			updated_at TEXT NOT NULL
		);
		CREATE TABLE IF NOT EXISTS school_meal_menus (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			year_month TEXT NOT NULL,
			menu_date TEXT NOT NULL,
			meal_type TEXT NOT NULL DEFAULT 'Öğle Yemeği',
			items_json TEXT NOT NULL,
			total_carbs_g REAL NOT NULL DEFAULT 0.0,
			source_note TEXT DEFAULT '',
			updated_at TEXT NOT NULL,
			UNIQUE(menu_date, meal_type)
		);
		CREATE TABLE IF NOT EXISTS insulin_bolus_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
			log_date TEXT NOT NULL,
			log_time TEXT NOT NULL,
			meal_type TEXT NOT NULL DEFAULT 'Öğle Yemeği',
			bg_mg_dl REAL,
			ketone_status TEXT DEFAULT '',
			menu_carbs_g REAL NOT NULL DEFAULT 0.0,
			menu_consumed_pct REAL NOT NULL DEFAULT 100.0,
			extra_foods_json TEXT DEFAULT '[]',
			extra_carbs_g REAL NOT NULL DEFAULT 0.0,
			total_carbs_g REAL NOT NULL DEFAULT 0.0,
			activity_level TEXT DEFAULT 'Yok / Normal Sınıf Dersi',
			activity_duration_min INTEGER DEFAULT 0,
			activity_reduction_pct REAL NOT NULL DEFAULT 0.0,
			icr_used REAL NOT NULL DEFAULT 15.0,
			isf_used REAL NOT NULL DEFAULT 50.0,
			iob_units REAL NOT NULL DEFAULT 0.0,
			carb_dose_units REAL NOT NULL DEFAULT 0.0,
			correction_dose_units REAL NOT NULL DEFAULT 0.0,
			calculated_dose_units REAL NOT NULL DEFAULT 0.0,
			applied_dose_units REAL NOT NULL DEFAULT 0.0,
			injection_site TEXT DEFAULT '',
			meal_summary TEXT DEFAULT '',
			ai_commentary TEXT DEFAULT '',
			administered_by TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);
		CREATE INDEX IF NOT EXISTS idx_menus_ym_date ON school_meal_menus(year_month, menu_date);
		CREATE INDEX IF NOT EXISTS idx_bolus_stu_date ON insulin_bolus_logs(student_id, log_date DESC, log_time DESC);
		"""
	)


def default_diabetes_profile(student_id: int) -> dict[str, Any]:
	return {
		"student_id": int(student_id),
		"is_diabetic": 0,
		"diabetes_type": "Tip 1 Diyabet (T1DM)",
		"insulin_brand": "NovoRapid (İnsülin Aspart)",
		"basal_insulin": "",
		"delivery_method": "İnsülin Kalemi (0.5 Ü Yarım Üniteli — Pediatrik)",
		"pen_increment": 0.5,
		"icr_breakfast": 15.0,
		"icr_lunch": 15.0,
		"icr_snack": 18.0,
		"isf": 50.0,
		"target_bg": 120.0,
		"correction_threshold": 150.0,
		"hypo_threshold": 70.0,
		"hyper_ketone_threshold": 250.0,
		"insulin_duration_hours": 3.5,
		"tdd_units": 0.0,
		"max_single_bolus": 12.0,
		"cgm_device": "",
		"glucagon_location": "",
		"physician_name": "",
		"emergency_protocol": "",
		"updated_at": "",
		"_exists": False,
	}


def get_diabetes_profile(connect_db_fn, student_id: int) -> dict[str, Any]:
	base = default_diabetes_profile(student_id)
	try:
		with connect_db_fn() as conn:
			ensure_diabetes_tables(conn)
			row = conn.execute("SELECT * FROM diabetes_profiles WHERE student_id=?", (int(student_id),)).fetchone()
		if row:
			d = dict(row)
			base.update(d)
			base["_exists"] = True
	except Exception:
		pass
	return base


def save_diabetes_profile(connect_db_fn, student_id: int, data: dict[str, Any], sync_chronic: bool = False) -> None:
	pid = int(student_id)
	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		conn.execute(
			"""INSERT INTO diabetes_profiles(
				student_id, is_diabetic, diabetes_type, insulin_brand, basal_insulin,
				delivery_method, pen_increment, icr_breakfast, icr_lunch, icr_snack,
				isf, target_bg, correction_threshold, hypo_threshold, hyper_ketone_threshold,
				insulin_duration_hours, tdd_units, max_single_bolus, cgm_device,
				glucagon_location, physician_name, emergency_protocol, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(student_id) DO UPDATE SET
				is_diabetic=excluded.is_diabetic,
				diabetes_type=excluded.diabetes_type,
				insulin_brand=excluded.insulin_brand,
				basal_insulin=excluded.basal_insulin,
				delivery_method=excluded.delivery_method,
				pen_increment=excluded.pen_increment,
				icr_breakfast=excluded.icr_breakfast,
				icr_lunch=excluded.icr_lunch,
				icr_snack=excluded.icr_snack,
				isf=excluded.isf,
				target_bg=excluded.target_bg,
				correction_threshold=excluded.correction_threshold,
				hypo_threshold=excluded.hypo_threshold,
				hyper_ketone_threshold=excluded.hyper_ketone_threshold,
				insulin_duration_hours=excluded.insulin_duration_hours,
				tdd_units=excluded.tdd_units,
				max_single_bolus=excluded.max_single_bolus,
				cgm_device=excluded.cgm_device,
				glucagon_location=excluded.glucagon_location,
				physician_name=excluded.physician_name,
				emergency_protocol=excluded.emergency_protocol,
				updated_at=excluded.updated_at""",
			(
				pid,
				1 if data.get("is_diabetic", True) else 0,
				str(data.get("diabetes_type") or "Tip 1 Diyabet (T1DM)"),
				str(data.get("insulin_brand") or "NovoRapid (İnsülin Aspart)"),
				str(data.get("basal_insulin") or "").strip(),
				str(data.get("delivery_method") or "İnsülin Kalemi (0.5 Ü Yarım Üniteli — Pediatrik)"),
				float(data.get("pen_increment") or 0.5),
				max(1.0, float(data.get("icr_breakfast") or 15.0)),
				max(1.0, float(data.get("icr_lunch") or 15.0)),
				max(1.0, float(data.get("icr_snack") or 18.0)),
				max(5.0, float(data.get("isf") or 50.0)),
				max(70.0, float(data.get("target_bg") or 120.0)),
				max(100.0, float(data.get("correction_threshold") or 150.0)),
				max(50.0, float(data.get("hypo_threshold") or 70.0)),
				max(180.0, float(data.get("hyper_ketone_threshold") or 250.0)),
				max(2.0, min(6.0, float(data.get("insulin_duration_hours") or 3.5))),
				max(0.0, float(data.get("tdd_units") or 0.0)),
				max(1.0, float(data.get("max_single_bolus") or 12.0)),
				str(data.get("cgm_device") or "").strip(),
				str(data.get("glucagon_location") or "").strip(),
				str(data.get("physician_name") or "").strip(),
				str(data.get("emergency_protocol") or "").strip(),
				now_iso,
			),
		)
		if sync_chronic and data.get("is_diabetic", True):
			hp_row = conn.execute("SELECT chronic_conditions FROM health_profiles WHERE student_id=?", (pid,)).fetchone()
			cur_ch = str(hp_row["chronic_conditions"] or "").strip() if hp_row else ""
			if "diyabet" not in cur_ch.lower():
				new_ch = f"{cur_ch}, {data.get('diabetes_type', 'Tip 1 Diyabet')}".strip(", ")
				conn.execute(
					"""INSERT INTO health_profiles(student_id, chronic_conditions, updated_at)
					   VALUES (?, ?, ?)
					   ON CONFLICT(student_id) DO UPDATE SET chronic_conditions=excluded.chronic_conditions, updated_at=excluded.updated_at""",
					(pid, new_ch, now_iso),
				)


def upsert_school_meal_menu(
	connect_db_fn,
	menu_date: str,
	meal_type: str,
	items: list[dict[str, Any]],
	source_note: str = "",
) -> None:
	mdate = str(menu_date).strip()[:10]
	ym = mdate[:7]
	mtype = str(meal_type or "Öğle Yemeği").strip()
	tot_c = round(sum(max(0.0, float(it.get("carbs_g") or 0.0)) for it in items), 1)
	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		conn.execute(
			"""INSERT INTO school_meal_menus(year_month, menu_date, meal_type, items_json, total_carbs_g, source_note, updated_at)
			   VALUES (?, ?, ?, ?, ?, ?, ?)
			   ON CONFLICT(menu_date, meal_type) DO UPDATE SET
				 year_month=excluded.year_month,
				 items_json=excluded.items_json,
				 total_carbs_g=excluded.total_carbs_g,
				 source_note=excluded.source_note,
				 updated_at=excluded.updated_at""",
			(ym, mdate, mtype, json.dumps(items, ensure_ascii=False), tot_c, source_note, now_iso),
		)


def get_school_menus_for_month(connect_db_fn, year_month: str) -> list[dict[str, Any]]:
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		rows = conn.execute(
			"SELECT * FROM school_meal_menus WHERE year_month=? ORDER BY menu_date ASC, meal_type ASC",
			(str(year_month)[:7],),
		).fetchall()
	out: list[dict[str, Any]] = []
	for r in rows:
		d = dict(r)
		try:
			d["items"] = json.loads(d.get("items_json") or "[]")
		except Exception:
			d["items"] = []
		out.append(d)
	return out


def get_school_menu_for_date(connect_db_fn, menu_date: str, meal_type: str = "Öğle Yemeği") -> dict[str, Any] | None:
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		row = conn.execute(
			"SELECT * FROM school_meal_menus WHERE menu_date=? AND meal_type=?",
			(str(menu_date)[:10], str(meal_type)),
		).fetchone()
		if not row:
			row = conn.execute(
				"SELECT * FROM school_meal_menus WHERE menu_date=? ORDER BY id ASC LIMIT 1",
				(str(menu_date)[:10],),
			).fetchone()
	if not row:
		return None
	d = dict(row)
	try:
		d["items"] = json.loads(d.get("items_json") or "[]")
	except Exception:
		d["items"] = []
	return d


def delete_school_menu_by_id(connect_db_fn, menu_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		conn.execute("DELETE FROM school_meal_menus WHERE id=?", (int(menu_id),))


# ==============================================================================
# 3. KESİN PEDİATRİK BOLUS MATEMATİĞİ, AKTİF İNSÜLİN (IOB) & GÜVENLİK MOTORU
# ==============================================================================

def calculate_active_iob(
	connect_db_fn,
	student_id: int,
	ref_dt: datetime,
	duration_hours: float = 3.5,
) -> tuple[float, list[dict[str, Any]]]:
	"""Son `duration_hours` saat içindeki uygulanan boluslardan kalan aktif insülini (IOB) hesaplar."""
	dur = max(2.0, min(6.0, float(duration_hours or 3.5)))
	cutoff = ref_dt - timedelta(hours=dur)
	d_min = cutoff.date().isoformat()
	d_max = ref_dt.date().isoformat()
	recent_active: list[dict[str, Any]] = []
	total_iob = 0.0
	try:
		with connect_db_fn() as conn:
			ensure_diabetes_tables(conn)
			rows = conn.execute(
				"""SELECT id, log_date, log_time, applied_dose_units, meal_type, injection_site
				   FROM insulin_bolus_logs
				   WHERE student_id=? AND log_date >= ? AND log_date <= ? AND applied_dose_units > 0
				   ORDER BY log_date DESC, log_time DESC""",
				(int(student_id), d_min, d_max),
			).fetchall()
		for r in rows:
			try:
				dt_str = f"{str(r['log_date'])[:10]} {str(r['log_time'])[:5]}"
				bolus_dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
			except Exception:
				continue
			diff_hr = (ref_dt - bolus_dt).total_seconds() / 3600.0
			if 0.0 <= diff_hr < dur:
				dose = float(r["applied_dose_units"] or 0.0)
				rem_ratio = max(0.0, 1.0 - (diff_hr / dur))
				rem_iob = round(dose * rem_ratio, 2)
				total_iob += rem_iob
				recent_active.append({
					"id": int(r["id"]),
					"time": str(r["log_time"])[:5],
					"date": str(r["log_date"])[:10],
					"dose": dose,
					"elapsed_min": int(round(diff_hr * 60)),
					"remaining_iob": rem_iob,
					"site": str(r["injection_site"] or ""),
				})
	except Exception:
		pass
	return round(total_iob, 2), recent_active


def get_last_injection_site(connect_db_fn, student_id: int) -> tuple[str, str]:
	try:
		with connect_db_fn() as conn:
			ensure_diabetes_tables(conn)
			row = conn.execute(
				"""SELECT injection_site, log_date, log_time FROM insulin_bolus_logs
				   WHERE student_id=? AND injection_site IS NOT NULL AND injection_site != ''
				   ORDER BY log_date DESC, log_time DESC, id DESC LIMIT 1""",
				(int(student_id),),
			).fetchone()
		if row:
			return str(row["injection_site"]), f"{row['log_date']} {row['log_time']}"
	except Exception:
		pass
	return "", ""


def _round_to_increment(val: float, inc: float) -> float:
	if val <= 0:
		return 0.0
	step = float(inc or 0.5)
	if step <= 0:
		step = 0.5
	return round(round(val / step) * step, 2)


def calculate_bolus_recommendation(
	*,
	profile: dict[str, Any],
	meal_type: str,
	bg_mg_dl: float | None,
	ketone_status: str = "Ölçülmedi",
	menu_carbs_raw: float = 0.0,
	menu_consumed_pct: float = 100.0,
	extra_carbs_g: float = 0.0,
	activity_level: str = "Yok / Normal Sınıf Dersi",
	activity_duration_min: int = 0,
	custom_activity_pct: float | None = None,
	iob_units: float = 0.0,
	student_weight_kg: float | None = None,
	student_age_years: float | None = None,
) -> dict[str, Any]:
	"""
	ISPAD / Çocuk Endokrinolojisi standartlarına göre deterministik insülin bolus hesabı ve güvenlik değerlendirmesi.
	"""
	if meal_type == "Kahvaltı":
		icr = float(profile.get("icr_breakfast") or 15.0)
	elif meal_type == "İkindi / Ara Öğün":
		icr = float(profile.get("icr_snack") or 18.0)
	else:
		icr = float(profile.get("icr_lunch") or 15.0)
	icr = max(1.0, icr)

	isf = max(5.0, float(profile.get("isf") or 50.0))
	target_bg = float(profile.get("target_bg") or 120.0)
	corr_thresh = float(profile.get("correction_threshold") or 150.0)
	hypo_thresh = float(profile.get("hypo_threshold") or 70.0)
	hyper_thresh = float(profile.get("hyper_ketone_threshold") or 250.0)
	pen_inc = float(profile.get("pen_increment") or 0.5)
	max_bolus = float(profile.get("max_single_bolus") or 12.0)
	tdd = float(profile.get("tdd_units") or 0.0)

	if meal_type == "Sadece Düzeltme (Öğünsüz)":
		eff_menu_carbs = 0.0
		eff_extra_carbs = 0.0
	else:
		eff_menu_carbs = round(max(0.0, float(menu_carbs_raw or 0.0)) * (max(0.0, float(menu_consumed_pct)) / 100.0), 1)
		eff_extra_carbs = round(max(0.0, float(extra_carbs_g or 0.0)), 1)
	total_carbs = round(eff_menu_carbs + eff_extra_carbs, 1)

	if activity_level == "Özel Yüzde Belirle" and custom_activity_pct is not None:
		act_pct = max(0.0, min(70.0, float(custom_activity_pct)))
	else:
		act_pct = float(ACTIVITY_PRESETS.get(activity_level, {}).get("pct", 0.0))

	raw_carb_dose = (total_carbs / icr) if total_carbs > 0 else 0.0
	adj_carb_dose = raw_carb_dose * (1.0 - act_pct / 100.0)

	is_hypo = (bg_mg_dl is not None and bg_mg_dl < hypo_thresh)
	is_severe_hypo = (bg_mg_dl is not None and bg_mg_dl < 54.0)
	is_hyper_ketone = (bg_mg_dl is not None and bg_mg_dl >= hyper_thresh)
	ketone_positive = ("Pozitif" in str(ketone_status))
	ketone_trace = ("Eser" in str(ketone_status))

	gross_corr = 0.0
	net_corr = 0.0
	if (not is_hypo) and (bg_mg_dl is not None) and (bg_mg_dl > corr_thresh):
		gross_corr = (bg_mg_dl - target_bg) / isf
		net_corr = max(0.0, gross_corr - max(0.0, float(iob_units or 0.0)))

	unrounded_total = 0.0 if is_hypo else max(0.0, adj_carb_dose + net_corr)
	recommended_dose = 0.0 if is_hypo else _round_to_increment(unrounded_total, pen_inc)

	exceeds_max = recommended_dose > max_bolus
	tdd_per_kg = round(tdd / float(student_weight_kg), 2) if (student_weight_kg and student_weight_kg > 0 and tdd > 0) else None

	# Güvenlik uyarıları ve pediatrik kurallar
	alerts: list[dict[str, str]] = []
	if is_severe_hypo:
		alerts.append({
			"level": "critical",
			"title": "🚨 AĞIR HİPOGLİSEMİ (<54 mg/dL) — İNSÜLİN KESİNLİKLE YAPILMAZ!",
			"body": "Bilinç açıksa hemen 15-20 g hızlı karbonhidrat (150 ml meyve suyu veya 4 kesme şeker) verin, 15 dk sonra KŞ tekrar ölçün. Bilinç bulanıksa veya yutamıyorsa HİÇBİR ŞEY İÇİRMEYİN, GLUKAGON uygulayıp 112 ve veliyi arayın!",
		})
	elif is_hypo:
		alerts.append({
			"level": "critical",
			"title": f"🚨 HİPOGLİSEMİ (<{hypo_thresh:.0f} mg/dL) — 15-15 ACİL PROTOKOLÜ!",
			"body": "Önce 15 g hızlı etkili karbonhidrat verip 15 dk dinlendirin ve KŞ'yi tekrar ölçün. Kan şekeri en az 80–100 mg/dL üzerine çıkmadan insülin enjeksiyonu yapılmaz!",
		})

	if ketone_positive:
		alerts.append({
			"level": "critical",
			"title": "🚨 KETON POZİTİF (+) — EGZERSİZ YASAK & ACİL HEKİM/VELİ BİLDİRİMİ!",
			"body": "Pozitif keton varlığında beden eğitimi / fiziksel aktivite kesinlikle yapılmaz (DKA riski). Bol su içirilmeli, veli ve çocuk endokrinoloji hekimi derhal bilgilendirilmelidir.",
		})
	elif is_hyper_ketone:
		alerts.append({
			"level": "warning",
			"title": f"⚠️ YÜKSEK KAN ŞEKERİ (≥{hyper_thresh:.0f} mg/dL) — KETON KONTROLÜ GEREKLİ!",
			"body": "İdrar veya kan ketonu kontrol edilmeli, öğrenciye bol su içirilmelidir. Keton pozitif çıkarsa egzersiz yapılmamalıdır.",
		})
	elif ketone_trace:
		alerts.append({
			"level": "warning",
			"title": "⚠️ Eser / Sınırda Keton — Sıvı Alımı & 1-2 Saat Sonra Tekrar Ölçüm",
			"body": "Bol su içirin, egzersiz şiddetini sınırlandırın ve 1-2 saat içinde kan şekeri + keton tekrarı yapın.",
		})

	if act_pct > 0 and bg_mg_dl is not None:
		if 70.0 <= bg_mg_dl < 100.0 and not is_hypo:
			alerts.append({
				"level": "warning",
				"title": "🏃 Egzersiz Öncesi Düşük-Normal KŞ (70–99 mg/dL)",
				"body": "Fiziksel aktivite öncesi hipoglisemiyi önlemek için ek 15 g karbonhidrat (1 küçük meyve/kraker) verilmesi ve 15 dk sonra tekrar KŞ bakılması önerilir.",
			})
		elif bg_mg_dl >= hyper_thresh:
			alerts.append({
				"level": "warning",
				"title": "🏃 Yüksek Şekerle Egzersiz Uyarısı",
				"body": "KŞ ≥ 250 mg/dL iken keton negatif olduğu teyit edilmeden fiziksel aktiviteye başlanmamalıdır.",
			})

	if iob_units > 0.1:
		alerts.append({
			"level": "info",
			"title": f"⏱️ Aktif İnsülin (IOB) Koruması Devrede: {iob_units:.2f} Ü",
			"body": f"Son {profile.get('insulin_duration_hours', 3.5)} saat içinde uygulanan dozdan kalan {iob_units:.2f} Ü aktif insülin çifte doz (stacking) riskini önlemek için düzeltme dozundan düşüldü.",
		})

	if exceeds_max:
		alerts.append({
			"level": "critical",
			"title": f"🛑 Maksimum Tek Bolus Sınırı Aşıldı ({recommended_dose:.2f} Ü > {max_bolus:.1f} Ü)!",
			"body": "Hesaplanan doz profildeki tek seferlik maksimum güvenlik sınırını aşıyor. Karbonhidrat girişini, İKO/İDF değerlerini doğrulayın ve hekim/veli onayı olmadan uygulamayın!",
		})

	# TDD/kg pediatrik referans kontrolü
	tdd_kg_note = ""
	if tdd_per_kg is not None:
		if tdd_per_kg < 0.45:
			tdd_kg_note = f"{tdd_per_kg:.2f} Ü/kg/gün (Balayı / Düşük İhtiyaç Dönemi)"
		elif tdd_per_kg <= 1.0:
			tdd_kg_note = f"{tdd_per_kg:.2f} Ü/kg/gün (Okul Çağı Tipik Pediatrik Aralık: 0.5–1.0 Ü/kg)"
		elif tdd_per_kg <= 1.5:
			tdd_kg_note = f"{tdd_per_kg:.2f} Ü/kg/gün (Ergenlik / Büyüme Atılımı Aralığı: 1.0–1.5 Ü/kg)"
		else:
			tdd_kg_note = f"{tdd_per_kg:.2f} Ü/kg/gün (Yüksek Doz İhtiyacı — Endokrinoloji Değerlendirmesi)"

	# Adım adım şeffaf formül dökümü
	formula_steps: list[str] = []
	if meal_type != "Sadece Düzeltme (Öğünsüz)":
		formula_steps.append(
			f"1. **Toplam Karbonhidrat:** ({menu_carbs_raw:.1f}g menü × %{menu_consumed_pct:.0f} = {eff_menu_carbs:.1f}g) + {eff_extra_carbs:.1f}g ekstra = **{total_carbs:.1f} g**"
		)
		formula_steps.append(
			f"2. **Ham Öğün Bolusu:** {total_carbs:.1f} g ÷ {icr:.1f} ({meal_type} İKO) = **{raw_carb_dose:.2f} Ü**"
		)
		if act_pct > 0:
			formula_steps.append(
				f"3. **Fiziksel Aktivite İndirimi (-%{act_pct:.0f}):** {raw_carb_dose:.2f} Ü × {(1.0 - act_pct/100.0):.2f} = **{adj_carb_dose:.2f} Ü**"
			)
	else:
		formula_steps.append("1. **Öğün Bolusu:** Öğünsüz düzeltme modu (**0.00 Ü**)")

	if is_hypo:
		formula_steps.append(f"4. **Düzeltme Bolusu:** KŞ ({bg_mg_dl:.0f} mg/dL) hipoglisemi sınırının altında → **DOZ KİLİTLENDİ (0.0 Ü)**")
	elif bg_mg_dl is not None and bg_mg_dl > corr_thresh:
		if iob_units > 0:
			formula_steps.append(
				f"4. **Düzeltme Bolusu (IOB Korumalı):** (({bg_mg_dl:.0f} − {target_bg:.0f}) ÷ {isf:.0f}) − {iob_units:.2f} Ü IOB = **{net_corr:.2f} Ü** *(Ham: {gross_corr:.2f} Ü)*"
			)
		else:
			formula_steps.append(
				f"4. **Düzeltme Bolusu:** ({bg_mg_dl:.0f} − {target_bg:.0f}) ÷ {isf:.0f} (İDF) = **{net_corr:.2f} Ü**"
			)
	else:
		bg_txt = f"{bg_mg_dl:.0f} mg/dL ≤ {corr_thresh:.0f}" if bg_mg_dl is not None else "KŞ girilmedi"
		formula_steps.append(f"4. **Düzeltme Bolusu:** {bg_txt} → **0.00 Ü**")

	formula_steps.append(
		f"5. **Nihai Önerilen Doz ({pen_inc} Ü hassasiyet):** {unrounded_total:.2f} Ü → **`{recommended_dose:.2f} Ü`**"
	)

	return {
		"icr_used": icr,
		"isf_used": isf,
		"target_bg": target_bg,
		"effective_menu_carbs_g": eff_menu_carbs,
		"extra_carbs_g": eff_extra_carbs,
		"total_carbs_g": total_carbs,
		"raw_carb_dose": round(raw_carb_dose, 2),
		"activity_reduction_pct": act_pct,
		"adjusted_carb_dose": round(adj_carb_dose, 2),
		"gross_correction_dose": round(gross_corr, 2),
		"iob_units": round(iob_units, 2),
		"net_correction_dose": round(net_corr, 2),
		"unrounded_total_units": round(unrounded_total, 2),
		"recommended_dose_units": recommended_dose,
		"pen_increment": pen_inc,
		"is_hypo": is_hypo,
		"is_severe_hypo": is_severe_hypo,
		"is_hyper_ketone": is_hyper_ketone,
		"ketone_positive": ketone_positive,
		"exceeds_max_bolus": exceeds_max,
		"tdd_per_kg": tdd_per_kg,
		"tdd_kg_note": tdd_kg_note,
		"alerts": alerts,
		"formula_steps": formula_steps,
	}


# ==============================================================================
# 4. RAPORLAMA: AYLIK YEMEK-İNSÜLİN PLANI & İNSÜLİN DEFTERİ (WORD / EXCEL)
# ==============================================================================

def build_monthly_insulin_plan_df(connect_db_fn, student_id: int, year_month: str) -> tuple[pd.DataFrame, dict[str, Any]]:
	"""Seçili öğrencinin profiline ve aylık okul yemek listesine göre tüm ayın tahmini öğün insülin tablosunu oluşturur."""
	prof = get_diabetes_profile(connect_db_fn, student_id)
	menus = get_school_menus_for_month(connect_db_fn, year_month)
	pen_inc = float(prof.get("pen_increment") or 0.5)
	rows: list[dict[str, Any]] = []
	for m in menus:
		mtype = str(m.get("meal_type") or "Öğle Yemeği")
		icr = float(prof.get("icr_breakfast") if mtype == "Kahvaltı" else (prof.get("icr_snack") if mtype == "İkindi / Ara Öğün" else prof.get("icr_lunch")) or 15.0)
		tot_c = float(m.get("total_carbs_g") or 0.0)
		items = m.get("items") or []
		dish_str = ", ".join(f"{it.get('name')} ({it.get('carbs_g', 0):.0f}g)" for it in items)
		gi_notes = [f"{it.get('name')}: {it.get('fat_protein_note')}" for it in items if str(it.get("fat_protein_note") or "").strip()]
		full_dose = _round_to_increment(tot_c / icr, pen_inc) if tot_c > 0 else 0.0
		half_dose = _round_to_increment((tot_c * 0.5) / icr, pen_inc) if tot_c > 0 else 0.0
		act_dose = _round_to_increment((tot_c / icr) * 0.75, pen_inc) if tot_c > 0 else 0.0
		try:
			dt_lbl = date.fromisoformat(str(m["menu_date"])[:10]).strftime("%d.%m.%Y")
		except Exception:
			dt_lbl = str(m["menu_date"])
		rows.append({
			"Tarih": dt_lbl,
			"Öğün": mtype,
			"Menü İçeriği (Tahmini Karb)": dish_str or "—",
			"Toplam Karb (g)": round(tot_c, 1),
			"İKO (g/Ü)": round(icr, 1),
			"Tam Porsiyon Doz (Ü)": full_dose,
			"Yarım Porsiyon (%50) Doz (Ü)": half_dose,
			"Beden Eğitimi (-%25) Doz (Ü)": act_dose,
			"Glisemik / Yağ-Protein Notu": " | ".join(gi_notes) if gi_notes else "Normal",
		})
	return pd.DataFrame(rows), prof


def get_student_bolus_logs_df(connect_db_fn, student_id: int, days_back: int = 30) -> pd.DataFrame:
	cutoff = (date.today() - timedelta(days=int(days_back))).isoformat()
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		df = pd.read_sql_query(
			"""SELECT id, log_date AS Tarih, log_time AS Saat, meal_type AS Öğün,
					  bg_mg_dl AS Kan_Sekeri, ketone_status AS Keton,
					  menu_carbs_g AS Menu_Karb, menu_consumed_pct AS Tuketim_Yuzde,
					  extra_carbs_g AS Ekstra_Karb, total_carbs_g AS Toplam_Karb,
					  activity_level AS Aktivite, activity_reduction_pct AS Aktivite_Indirim,
					  iob_units AS Aktif_IOB, calculated_dose_units AS Hesaplanan_Doz,
					  applied_dose_units AS Uygulanan_Doz, injection_site AS Enjeksiyon_Bolgesi,
					  meal_summary AS Ogun_Ozeti, administered_by AS Uygulayan, notes AS Notlar
			   FROM insulin_bolus_logs
			   WHERE student_id=? AND log_date >= ?
			   ORDER BY log_date DESC, log_time DESC, id DESC""",
			conn,
			params=(int(student_id), cutoff),
		)
	return df


def compute_glycemic_tir_stats(logs_df: pd.DataFrame, hypo_thresh: float = 70.0, hyper_high: float = 180.0) -> dict[str, Any]:
	if logs_df is None or logs_df.empty:
		return {
			"total_logs": 0, "bg_count": 0, "avg_bg": None,
			"tir_pct": None, "hypo_cnt": 0, "hyper_cnt": 0,
			"avg_dose": None, "total_dose": 0.0, "site_counts": {},
		}
	bg_vals = pd.to_numeric(logs_df["Kan_Sekeri"], errors="coerce").dropna()
	doses = pd.to_numeric(logs_df["Uygulanan_Doz"], errors="coerce").dropna()
	bg_n = int(len(bg_vals))
	hypo_c = int((bg_vals < hypo_thresh).sum()) if bg_n else 0
	hyper_c = int((bg_vals > hyper_high).sum()) if bg_n else 0
	in_range = int(((bg_vals >= hypo_thresh) & (bg_vals <= hyper_high)).sum()) if bg_n else 0
	tir_pct = round(in_range / bg_n * 100.0, 1) if bg_n else None

	sites = logs_df["Enjeksiyon_Bolgesi"].dropna().astype(str).str.strip()
	site_counts = sites[sites != ""].value_counts().to_dict()
	return {
		"total_logs": int(len(logs_df)),
		"bg_count": bg_n,
		"avg_bg": round(float(bg_vals.mean()), 1) if bg_n else None,
		"tir_pct": tir_pct,
		"hypo_cnt": hypo_c,
		"hyper_cnt": hyper_c,
		"avg_dose": round(float(doses.mean()), 2) if len(doses) else None,
		"total_dose": round(float(doses.sum()), 2) if len(doses) else 0.0,
		"site_counts": site_counts,
	}


# ==============================================================================
# 5. ANA ARAYÜZ: 💉 İNSÜLİN & DİYABET YÖNETİM MERKEZİ (4 SEKME)
# ==============================================================================

def _build_student_growth_bundle(ctx_api: dict[str, Any], person: sqlite3.Row) -> dict[str, Any]:
	pid = int(person["id"])
	birth_s = str(person["birth_date"] or "")
	sex_s = str(person["sex"] or "")
	age_m = ctx_api["age_in_months"](birth_s, date.today()) if birth_s else None
	age_y = round(age_m / 12.0, 1) if age_m is not None else None
	meas = ctx_api["get_person_latest_measurements"](pid)
	h_cm = meas.get("height_cm") if meas.get("has_data") else None
	w_kg = meas.get("weight_kg") if meas.get("has_data") else None
	bmi_v = meas.get("bmi") if meas.get("has_data") else None

	h_ev = ctx_api["evaluate_measurement"](h_cm, "height", sex_s, age_m) if (h_cm and sex_s and age_m is not None and age_m <= 228) else None
	w_ev = ctx_api["evaluate_measurement"](w_kg, "weight", sex_s, age_m) if (w_kg and sex_s and age_m is not None and age_m <= 228) else None
	b_ev = ctx_api["evaluate_measurement"](bmi_v, "bmi", sex_s, age_m) if (bmi_v and sex_s and age_m is not None and age_m <= 228) else None

	return {
		"yas_ay": age_m,
		"yas_yil": age_y,
		"yas_etiketi": f"{age_m // 12} yaş {age_m % 12} ay" if age_m is not None else "Girilmedi",
		"cinsiyet": {"K": "Kız", "E": "Erkek"}.get(sex_s, "Belirtilmedi"),
		"boy_cm": round(float(h_cm), 1) if h_cm else None,
		"kilo_kg": round(float(w_kg), 1) if w_kg else None,
		"vki_bmi": round(float(bmi_v), 2) if bmi_v else None,
		"boy_persantil": f"%{h_ev['percentile']:.1f}p (Z: {h_ev['z']:+.2f})" if h_ev else "—",
		"kilo_persantil": f"%{w_ev['percentile']:.1f}p (Z: {w_ev['z']:+.2f})" if w_ev else "—",
		"vki_persantil": f"%{b_ev['percentile']:.1f}p ({b_ev['status']})" if b_ev else "—",
		"son_olcum_tarihi": meas.get("latest_measured_on") if meas.get("has_data") else None,
	}


def render_insulin_diabetes_page(ctx_api: dict[str, Any]) -> None:
	"""
	`persantil.py` tarafından çağrılan ana '💉 İnsülin & Diyabet Yönetimi' sayfası.
	ctx_api: connect_db, get_person_by_id, set_active_person, get_app_settings,
	         _get_active_gemini_config, _build_person_anon_bundle, get_person_latest_measurements,
	         evaluate_measurement, age_in_months, build_ai_custom_report_docx,
	         build_ai_custom_report_excel, mask_kvkk_value, get_person_disability_profile,
	         format_disability_badge_text, sync_after_person_mutation, clear_all_data_caches
	"""
	connect_db_fn = ctx_api["connect_db"]
	with connect_db_fn() as conn:
		ensure_diabetes_tables(conn)
		diab_rows = conn.execute(
			"""SELECT dp.student_id, dp.diabetes_type, dp.icr_lunch, dp.isf, s.full_name, c.name AS class_name
			   FROM diabetes_profiles dp
			   JOIN students s ON s.id = dp.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   WHERE dp.is_diabetic = 1
			   ORDER BY s.full_name"""
		).fetchall()
		all_stu_rows = conn.execute(
			"""SELECT s.id, s.full_name, s.student_no, COALESCE(c.name, 'Sivil / Diğer') AS class_name,
					  COALESCE(dp.is_diabetic, 0) AS is_diab
			   FROM students s
			   LEFT JOIN classes c ON c.id = s.class_id
			   LEFT JOIN diabetes_profiles dp ON dp.student_id = s.id
			   ORDER BY COALESCE(dp.is_diabetic, 0) DESC, c.name ASC, s.full_name ASC"""
		).fetchall()

	ui.markdown("## 💉 Okulda Diyabet & İnsülin Doz Yönetim Merkezi")
	ui.caption(
		"Öğrenci profili (yaş, boy, kilo, WHO persantil, İKO/İDF), aylık okul yemek programı, ekstra yiyecekler, fiziksel aktivite ve aktif insülin (IOB) korumasıyla kesin pediatrik bolus hesabı + Gemini AI klinik değerlendirme."
	)
	ui.markdown(
		"<div style='background:rgba(2,132,199,0.08);border:1px solid rgba(2,132,199,0.28);border-left:4px solid #0284c7;"
		"padding:8px 12px;border-radius:8px;margin-bottom:10px;font-size:12.5px;'>"
		"⚕️ <strong>Klinik Güvenlik Notu:</strong> Bu modül, çocuğun çocuk endokrinoloji hekimi tarafından belirlenen İKO / İDF / Hedef KŞ parametrelerine göre okul hemşiresine karar desteği sağlar. "
		"Her uygulama öncesi kan şekeri ölçümü doğrulanmalı, hekim reçetesi esas alınmalıdır."
		"</div>",
		unsafe_allow_html=True,
	)

	if not all_stu_rows:
		ui.info("Sistemde henüz kayıtlı öğrenci bulunmuyor. Önce '👥 Öğrenciler' sayfasından öğrenci ekleyin.")
		return

	# 1-Tık Diyabet Takipli Öğrenci Hızlı Geçiş Hapları
	is_sec_vis = ctx_api.get("is_section_visible") or (lambda _k: True)
	active_pid = ui.session_state.get("active_person_id")
	if diab_rows and is_sec_vis("insulin_quick_chips"):
		ui.caption("⚡ **Diyabet Takipli Öğrenciler (Tek Tıkla Geçiş):**")
		chip_n = min(5, len(diab_rows))
		chip_cols = ui.columns(chip_n)
		for d_idx, dr in enumerate(diab_rows):
			d_sid = int(dr["student_id"])
			is_cur_d = (active_pid is not None and int(active_pid) == d_sid)
			d_nm = ctx_api["mask_kvkk_value"](dr["full_name"], "name")
			btn_lbl = f"{'✅ ' if is_cur_d else '💙 '}{d_nm} ({dr['class_name'] or '—'})"
			with chip_cols[d_idx % chip_n]:
				if ui.button(btn_lbl, key=f"ins_quick_diab_chip_{d_sid}", type=("primary" if is_cur_d else "secondary"), use_container_width=True):
					ctx_api["set_active_person"](d_sid)
					ui.session_state.pop("ins_main_person_picker", None)
					ui.rerun()

	# Hızlı Öğrenci Seçici (Diyabetli öğrenciler 💙 etiketiyle en üstte!)
	opts_labels: list[str] = []
	lbl_to_id: dict[str, int] = {}
	sel_idx = 0
	for idx, sr in enumerate(all_stu_rows):
		tag = "💙 [Diyabet Takipli] " if int(sr["is_diab"] or 0) == 1 else "👤 "
		sno = f" (No: {sr['student_no']})" if sr["student_no"] else ""
		lbl = f"{tag}{sr['full_name']} · {sr['class_name']}{sno}"
		opts_labels.append(lbl)
		lbl_to_id[lbl] = int(sr["id"])
		if active_pid and int(sr["id"]) == int(active_pid):
			sel_idx = idx

	# Aktif kişi başka yerden (chip / Cmd+K) değiştiyse picker'ı senkronize et
	if ui.session_state.get("ins_picker_synced_pid") != active_pid:
		ui.session_state["ins_main_person_picker"] = opts_labels[sel_idx]
		ui.session_state["ins_picker_synced_pid"] = active_pid

	top_s1, top_s2 = ui.columns([2.6, 1.4])
	with top_s1:
		picked_lbl = ui.selectbox(
			"🎓 Diyabet / İnsülin Takibi Yapılacak Öğrenci (Diyabetli öğrenciler en üsttedir)",
			opts_labels,
			index=sel_idx,
			key="ins_main_person_picker",
		)
	sel_pid = lbl_to_id.get(picked_lbl, int(all_stu_rows[0]["id"]))
	if sel_pid != active_pid:
		ctx_api["set_active_person"](sel_pid)
		ui.session_state["ins_picker_synced_pid"] = sel_pid

	person = ctx_api["get_person_by_id"](sel_pid)
	if not person:
		ui.warning("Seçili kişi bulunamadı.")
		return

	prof = get_diabetes_profile(connect_db_fn, sel_pid)
	growth = _build_student_growth_bundle(ctx_api, person)
	dp_dis = ctx_api["get_person_disability_profile"](sel_pid, person=person)
	now_dt = datetime.now()
	iob_now, active_boluses = calculate_active_iob(connect_db_fn, sel_pid, now_dt, float(prof.get("insulin_duration_hours") or 3.5))
	last_site, last_site_dt = get_last_injection_site(connect_db_fn, sel_pid)

	with top_s2:
		ui.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
		ui.metric("💙 Diyabet Takipli Öğrenci", f"{len(diab_rows)} öğrenci")

	# Öğrenci Klinik Özet Bandı (Yaş, Boy, Kilo, VKİ/Persantil, İKO/İDF, IOB, Engel)
	disp_name = ctx_api["mask_kvkk_value"](person["full_name"], "name")
	bw_str = (
		f"📏 Boy: <b>{growth['boy_cm']} cm</b> ({growth['boy_persantil']}) · ⚖️ Kilo: <b>{growth['kilo_kg']} kg</b> ({growth['kilo_persantil']}) · VKİ: <b>{growth['vki_bmi']}</b>"
		if (growth["boy_cm"] and growth["kilo_kg"])
		else "📏 Boy/Kilo ölçümü girilmemiş *(AI daha doğru yorum için boy-kilo ölçümünü kullanır)*"
	)
	prof_status_html = (
		f"<span style='background:#0284c7;color:#fff;padding:2px 9px;border-radius:999px;font-size:11.5px;font-weight:700;'>💙 {html.escape(str(prof['diabetes_type']))}</span> "
		f"<b>İKO (S/Ö/A):</b> 1/{prof['icr_breakfast']:.0f}g · 1/{prof['icr_lunch']:.0f}g · 1/{prof['icr_snack']:.0f}g &nbsp;|&nbsp; "
		f"<b>İDF:</b> {prof['isf']:.0f} &nbsp;|&nbsp; <b>Hedef:</b> {prof['target_bg']:.0f} mg/dL &nbsp;|&nbsp; "
		f"<b>Aktif IOB:</b> <code>{iob_now:.2f} Ü</code>"
		if prof.get("_exists") and prof.get("is_diabetic")
		else "<span style='background:#f59e0b;color:#fff;padding:2px 9px;border-radius:999px;font-size:11.5px;font-weight:700;'>⚠️ Diyabet Profili Henüz Tanımlanmadı</span> — '⚙️ Diyabet Profili' sekmesinden hekim İKO/İDF değerlerini kaydedin (varsayılan 1/15g ve İDF 50 ile önizleme yapılıyor)."
	)
	dis_line = f"<div style='margin-top:4px;color:#6d28d9;font-size:12px;'>♿ <b>Özel Gereksinim/Engel:</b> {html.escape(ctx_api['format_disability_badge_text'](dp_dis))}</div>" if dp_dis.get("has_disability") else ""
	site_line = f" · 💉 <b>Son Enjeksiyon:</b> {html.escape(last_site)} ({html.escape(last_site_dt)})" if last_site else ""
	gluc_line = f" · 🆘 <b>Glukagon:</b> {html.escape(str(prof.get('glucagon_location') or ''))}" if prof.get("glucagon_location") else ""

	with ui.container(border=True):
		ui.markdown(
			f"<div style='display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px;'>"
			f"<div><strong style='font-size:1.05rem;'>👤 {html.escape(disp_name)}</strong> · {html.escape(str(person['class_name']))} · {growth['cinsiyet']} · 🎂 {growth['yas_etiketi']}</div>"
			f"<div style='font-size:12.5px;'>{prof_status_html}</div>"
			f"</div>"
			f"<div style='margin-top:5px;font-size:12.5px;color:var(--text-muted);'>{bw_str}{site_line}{gluc_line}</div>"
			f"{dis_line}",
			unsafe_allow_html=True,
		)

	tab_calc, tab_menu, tab_prof, tab_logs = ui.tabs([
		"🧮 İnsülin Doz Hesaplayıcı & AI Yorum",
		"📅 Aylık Yemek Programı & İnsülin Planı",
		"⚙️ Öğrenci Diyabet Profili (İKO / İDF)",
		"📈 Uygulama Defteri, TIR Trendi & Raporlar",
	])

	with tab_calc:
		_render_tab_bolus_calculator(ctx_api, person, sel_pid, prof, growth, dp_dis, iob_now, active_boluses, last_site)
	with tab_menu:
		_render_tab_monthly_menu(ctx_api, person, sel_pid, prof)
	with tab_prof:
		_render_tab_diabetes_profile(ctx_api, person, sel_pid, prof, growth)
	with tab_logs:
		_render_tab_logs_and_trends(ctx_api, person, sel_pid, prof, growth)


# ==============================================================================
# SEKME 1: AKILLI BOLUS HESAPLAYICI + MENÜ + EKSTRA YİYECEK + AKTİVİTE + GEMİNİ
# ==============================================================================

def _render_tab_bolus_calculator(
	ctx_api: dict[str, Any],
	person: sqlite3.Row,
	student_id: int,
	prof: dict[str, Any],
	growth: dict[str, Any],
	dp_dis: dict[str, Any],
	iob_now: float,
	active_boluses: list[dict[str, Any]],
	last_site: str,
) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_gai_on, gai_key, gai_model = ctx_api["_get_active_gemini_config"]()
	disp_name = ctx_api["mask_kvkk_value"](person["full_name"], "name")

	# Ekstra yiyecek oturum listesi
	extras_key = f"ins_extras_list_{student_id}"
	if extras_key not in ui.session_state or not isinstance(ui.session_state[extras_key], list):
		ui.session_state[extras_key] = []

	flash_k = f"ins_calc_flash_{student_id}"
	f_msg = ui.session_state.pop(flash_k, None)
	if f_msg:
		ui.success(f_msg)

	col_in, col_out = ui.columns([1.22, 0.98], gap="large", sticky_last_col=True)

	with col_in:
		with ui.container(border=True):
			hdr_m1, hdr_m2, hdr_m3 = ui.columns([2.3, 0.85, 0.85])
			with hdr_m1:
				ui.markdown("##### 1️⃣ Öğün, Kan Şekeri ve Keton Durumu")
			with hdr_m2:
				if ui.button("📅 Bugün", key=f"ins_dt_today_{student_id}", type="secondary", use_container_width=True):
					ui.session_state[f"ins_calc_date_{student_id}"] = date.today()
					ui.rerun()
			with hdr_m3:
				if ui.button("➡️ Yarın", key=f"ins_dt_tom_{student_id}", type="secondary", use_container_width=True):
					ui.session_state[f"ins_calc_date_{student_id}"] = date.today() + timedelta(days=1)
					ui.rerun()
			d1, d2, d3 = ui.columns(3)
			sel_date = d1.date_input("Tarih", value=date.today(), key=f"ins_calc_date_{student_id}")
			sel_time = d2.time_input("Saat", value=datetime.now().replace(second=0, microsecond=0).time(), key=f"ins_calc_time_{student_id}")
			meal_type = d3.selectbox("Öğün Türü", MEAL_TYPE_OPTIONS, index=0, key=f"ins_calc_meal_{student_id}")

			b1, b2 = ui.columns([1.15, 1.35])
			bg_raw = b1.text_input(
				"🩸 Güncel Kan Şekeri (mg/dL)",
				value="",
				placeholder="Örn: 165 (boş bırakılabilir)",
				key=f"ins_calc_bg_{student_id}",
			)
			bg_val: float | None = None
			if bg_raw and str(bg_raw).strip():
				try:
					parsed_bg = float(str(bg_raw).replace(",", ".").strip())
					if 20.0 <= parsed_bg <= 900.0:
						bg_val = parsed_bg
				except ValueError:
					ui.warning("Kan şekeri sayısal olmalıdır (örn: 145).")
			ketone_st = b2.selectbox("🧪 Keton Durumu (KŞ ≥ 250 ise kontrol edin)", KETONE_OPTIONS, index=0, key=f"ins_calc_ket_{student_id}")

		# 2. Günün Aylık Yemek Menüsü
		menu_carbs_raw = 0.0
		menu_consumed_pct = 100.0
		selected_menu_dish_names: list[str] = []

		if meal_type != "Sadece Düzeltme (Öğünsüz)":
			with ui.container(border=True):
				ui.markdown(f"##### 2️⃣ Günün Okul Yemek Menüsü ({sel_date.strftime('%d.%m.%Y')} · {meal_type})")
				day_menu = get_school_menu_for_date(connect_db_fn, sel_date.isoformat(), meal_type)
				if day_menu and day_menu.get("items"):
					m_items = day_menu["items"]
					item_labels = [
						f"{it['name']} — {it.get('portion', '1 porsiyon')} ({float(it.get('carbs_g') or 0):.0f} g karb)"
						for it in m_items
					]
					picked_dishes = ui.multiselect(
						"Öğrencinin Tabak / Menüden Aldığı Yemekler (Yemediğini çıkarabilirsiniz)",
						item_labels,
						default=item_labels,
						key=f"ins_menu_pick_{student_id}_{sel_date.isoformat()}_{meal_type}",
					)
					for lbl_i, it in zip(item_labels, m_items):
						if lbl_i in picked_dishes:
							menu_carbs_raw += float(it.get("carbs_g") or 0.0)
							selected_menu_dish_names.append(f"{it['name']} ({float(it.get('carbs_g') or 0):.0f}g)")
							if str(it.get("fat_protein_note") or "").strip():
								ui.caption(f"⚠️ **{it['name']}:** {it['fat_protein_note']}")

					cons_c1, cons_c2 = ui.columns([1.4, 1.0])
					with cons_c1:
						cons_opts = {
							"🍽️ Tamamını Yedi (%100)": 100.0,
							"🍽️ Çoğunu Yedi (%75)": 75.0,
							"🍽️ Yarısını Yedi (%50)": 50.0,
							"🍽️ Çok Az Yedi (%25)": 25.0,
							"🍽️ Ekstra / Büyük Porsiyon (%125)": 125.0,
							"🍽️ Hiç Yemedi (%0)": 0.0,
						}
						cons_lbl = ui.selectbox(
							"Tüketilen Porsiyon Oranı",
							list(cons_opts.keys()),
							index=0,
							key=f"ins_menu_cons_{student_id}_{sel_date.isoformat()}",
						)
						menu_consumed_pct = cons_opts[cons_lbl]
					with cons_c2:
						manual_menu_c = ui.number_input(
							"Menü Karb Düzeltme (g)",
							min_value=0.0,
							max_value=300.0,
							step=1.0,
							value=round(menu_carbs_raw, 1),
							key=f"ins_menu_ovr_{student_id}_{sel_date.isoformat()}_{len(picked_dishes)}",
						)
						menu_carbs_raw = float(manual_menu_c)
				else:
					ui.info(
						f"📅 `{sel_date.strftime('%d.%m.%Y')}` tarihi için kayıtlı yemek menüsü bulunamadı. "
						"**'📅 Aylık Yemek Programı'** sekmesinden okulun aylık menüsünü yükleyebilir veya aşağıya doğrudan öğün karbonhidratını girebilirsiniz."
					)
					menu_carbs_raw = float(
						ui.number_input(
							"Manuel Öğün Karbonhidratı (g)",
							min_value=0.0,
							max_value=300.0,
							step=1.0,
							value=0.0,
							key=f"ins_manual_meal_carbs_{student_id}",
						)
					)
					menu_consumed_pct = 100.0

			# 3. Ekstra Yiyecekler / Kantin / Beslenme Çantası + Gemini AI Karb Tahminleyici
			with ui.container(border=True):
				ui.markdown("##### 3️⃣ Ekstra Yedikleri / Yiyecekleri *(Kantin, Beslenme Çantası, Ara Öğün)*")
				ui.caption("Tek tıkla sık tüketilen okul yiyeceklerinden ekleyin veya Gemini AI'ya yazılı/fotoğraflı karb hesaplattırın:")
				q_cols = ui.columns(4)
				for qi, qf in enumerate(QUICK_EXTRA_FOODS[:8]):
					with q_cols[qi % 4]:
						if ui.button(qf["label"], key=f"ins_qf_{student_id}_{qi}", type="secondary", use_container_width=True):
							ui.session_state[extras_key].append({"name": qf["name"], "carbs_g": qf["carbs_g"], "gi": qf["gi"]})
							ui.rerun()

				ex_c1, ex_c2, ex_c3 = ui.columns([2.0, 1.0, 0.9])
				ex_name = ex_c1.text_input("Ekstra Yiyecek Adı", placeholder="Örn: 1 kase yoğurt / ev yapımı kek", key=f"ins_ex_nm_{student_id}")
				ex_cg = ex_c2.number_input("Karb (g)", min_value=0.0, max_value=200.0, step=1.0, value=15.0, key=f"ins_ex_cg_{student_id}")
				with ex_c3:
					ui.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
					if ui.button("➕ Ekle", key=f"ins_ex_add_{student_id}", use_container_width=True):
						if ex_name.strip():
							ui.session_state[extras_key].append({"name": ex_name.strip(), "carbs_g": float(ex_cg), "gi": "Orta"})
							ui.rerun()

				if is_gai_on:
					with ui.expander("✨ Gemini AI ile Yiyecek Metni veya Tabak Fotoğrafından Karbonhidrat Hesapla", expanded=False):
						ai_food_txt = ui.text_input(
							"Ne yedi / yiyecek? (Serbest Türkçe yazın)",
							placeholder="Örn: 1 küçük tabak spagetti, 1 dilim karpuz ve yarım bardak ayran",
							key=f"ins_ai_food_txt_{student_id}",
						)
						ai_food_img = ui.file_uploader(
							"📸 Veya Yemek / Atıştırmalık Fotoğrafı Yükle (Opsiyonel)",
							type=["png", "jpg", "jpeg", "webp"],
							key=f"ins_ai_food_img_{student_id}",
						)
						if ui.button("✨ Gemini ile Karbonhidrat Hesapla ve Listeye Ekle", type="primary", key=f"ins_btn_ai_food_{student_id}", use_container_width=True):
							f_bytes = ai_food_img.getvalue() if ai_food_img is not None else None
							f_name = ai_food_img.name if ai_food_img is not None else ""
							with ui.spinner("Gemini porsiyon karbonhidratlarını hesaplıyor..."):
								ok_ef, ef_res, used_ef = gai.estimate_extra_foods_carbs_with_gemini(
									api_key=gai_key,
									model=gai_model,
									food_description=ai_food_txt,
									file_bytes=f_bytes,
									file_name=f_name,
									student_age_years=growth.get("yas_yil"),
								)
							if ok_ef and ef_res.get("items"):
								for it in ef_res["items"]:
									ui.session_state[extras_key].append({
										"name": f"{it['name']} ({it.get('portion', '')})".strip(),
										"carbs_g": float(it["carbs_g"]),
										"gi": it.get("gi", "Orta"),
									})
								note_g = ef_res.get("glycemic_note") or ""
								ui.session_state[flash_k] = f"✅ Gemini ({used_ef}) {len(ef_res['items'])} kalem (~{ef_res['total_carbs_g']:.1f}g karb) ekledi! {note_g}"
								ui.rerun()
							else:
								ui.error(f"Karbonhidrat tahmin hatası: {ef_res.get('error', 'Ayrıştırılamadı')}")

				# Ekli Ekstra Yiyecekler Listesi
				cur_extras: list[dict[str, Any]] = ui.session_state[extras_key]
				if cur_extras:
					ex_tot_lbl = sum(float(x.get("carbs_g") or 0.0) for x in cur_extras)
					ui.markdown(f"**Eklenen Ekstra Yiyecekler (Toplam: `{ex_tot_lbl:.1f} g`):**")
					for idx_e, ex_it in enumerate(cur_extras):
						ec_l, ec_r = ui.columns([3.8, 0.7])
						ec_l.markdown(f"• {html.escape(str(ex_it['name']))} — **`{float(ex_it['carbs_g']):.1f} g`** *(Gİ: {ex_it.get('gi', 'Orta')})*")
						if ec_r.button("✖", key=f"ins_ex_rm_{student_id}_{idx_e}", use_container_width=True):
							ui.session_state[extras_key].pop(idx_e)
							ui.rerun()
					if ui.button("🗑️ Tüm Ekstraları Temizle", key=f"ins_ex_clr_{student_id}", type="secondary"):
						ui.session_state[extras_key] = []
						ui.rerun()

		# 4. Fiziksel Aktivite / Beden Eğitimi
		with ui.container(border=True):
			ui.markdown("##### 4️⃣ Planlanan / Yapılan Fiziksel Aktivite & Beden Eğitimi")
			act1, act2 = ui.columns([2.2, 1.0])
			act_level = act1.selectbox("🏃 Aktivite Şiddeti", list(ACTIVITY_PRESETS.keys()), index=0, key=f"ins_act_lvl_{student_id}")
			act_dur = act2.number_input("Süre (dk)", min_value=0, max_value=240, step=5, value=0 if act_level == "Yok / Normal Sınıf Dersi" else 40, key=f"ins_act_dur_{student_id}")
			custom_act_pct = None
			if act_level == "Özel Yüzde Belirle":
				custom_act_pct = float(ui.number_input("Özel İnsülin Azaltma Yüzdesi (%)", min_value=0.0, max_value=70.0, step=5.0, value=20.0, key=f"ins_act_cust_{student_id}"))
			act_note = ui.text_input("Aktivite Detayı (Opsiyonel)", placeholder="Örn: Öğleden sonra 2 saat beden eğitimi / futbol turnuvası", key=f"ins_act_note_{student_id}")

	# Kesin deterministik bolus hesabını çalıştır
	cur_extras_final: list[dict[str, Any]] = ui.session_state.get(extras_key) or []
	extra_carbs_sum = round(sum(float(x.get("carbs_g") or 0.0) for x in cur_extras_final), 1)
	calc_ref_dt = datetime.combine(sel_date, sel_time)
	iob_for_calc, _ = calculate_active_iob(connect_db_fn, student_id, calc_ref_dt, float(prof.get("insulin_duration_hours") or 3.5))

	calc = calculate_bolus_recommendation(
		profile=prof,
		meal_type=meal_type,
		bg_mg_dl=bg_val,
		ketone_status=ketone_st,
		menu_carbs_raw=menu_carbs_raw,
		menu_consumed_pct=menu_consumed_pct,
		extra_carbs_g=extra_carbs_sum,
		activity_level=act_level,
		activity_duration_min=int(act_dur or 0),
		custom_activity_pct=custom_act_pct,
		iob_units=iob_for_calc,
		student_weight_kg=growth.get("kilo_kg"),
		student_age_years=growth.get("yas_yil"),
	)

	with col_out:
		# Güvenlik Uyarıları
		for al in calc["alerts"]:
			if al["level"] == "critical":
				ui.error(f"**{al['title']}**\n\n{al['body']}")
			elif al["level"] == "warning":
				ui.warning(f"**{al['title']}**\n\n{al['body']}")
			else:
				ui.info(f"**{al['title']}**\n\n{al['body']}")

		# Hipoglisemi (<70) ise 15-15 Kuralı Canlı Geri Sayım Sayacı
		if calc["is_hypo"]:
			clk_id = f"hypo_clk_{student_id}"
			ui.markdown(
				f"<div class='revir-hypo-timer'>"
				f"<div><div style='font-size:12px;font-weight:800;color:#dc2626;'>⏱️ 15-15 KURALI CANLI GERİ SAYIM</div>"
				f"<div style='font-size:11.5px;color:var(--text-muted);'>15g hızlı karbonhidrat verin ve 15 dk sonra KŞ tekrar ölçün</div></div>"
				f"<div style='display:flex;align-items:center;gap:8px;'>"
				f"<span id='{clk_id}' class='revir-hypo-clock'>15:00</span>"
				f"<button type='button' class='revir-btn revir-btn-primary revir-btn-sm' onclick=\"window.__revirStartHypoTimer('{clk_id}', 900)\">▶ Başlat</button>"
				f"<button type='button' class='revir-btn revir-btn-sm' onclick=\"window.__revirResetHypoTimer('{clk_id}')\">↺</button>"
				f"</div></div>",
				unsafe_allow_html=True,
			)

		# Büyük Önerilen Doz Kartı (Karanlık Mod Uyumlu)
		is_danger_card = bool(calc["is_hypo"] or calc["exceeds_max_bolus"] or calc["ketone_positive"])
		hero_cls = "revir-dose-hero danger" if is_danger_card else "revir-dose-hero"
		dose_big = "0.0 Ü (DURDURULDU)" if calc["is_hypo"] else f"{calc['recommended_dose_units']:.2f} Ünite"

		with ui.container(border=True):
			ui.markdown(
				f"<div class='{hero_cls}'>"
				f"<div style='font-size:12px;font-weight:700;color:var(--text-muted);letter-spacing:0.4px;'>HESAPLANAN ÖNERİLEN BOLUS DOZU ({html.escape(str(prof.get('insulin_brand', 'Hızlı İnsülin')))})</div>"
				f"<div class='revir-dose-hero-val'>💉 {dose_big}</div>"
				f"<div style='font-size:12px;color:var(--text-muted);'>Toplam Karb: <b>{calc['total_carbs_g']:.1f} g</b> · Öğün Bolusu: <b>{calc['adjusted_carb_dose']:.2f} Ü</b> · Düzeltme: <b>{calc['net_correction_dose']:.2f} Ü</b> · IOB: <b>{calc['iob_units']:.2f} Ü</b></div>"
				f"</div>",
				unsafe_allow_html=True,
			)

			st_c1, st_c2, st_c3 = ui.columns(3)
			st_c1.metric("Toplam Karb", f"{calc['total_carbs_g']:.1f} g")
			st_c2.metric("Öğün / Düzeltme", f"{calc['adjusted_carb_dose']:.2f} / {calc['net_correction_dose']:.2f} Ü")
			st_c3.metric("Aktif IOB", f"{calc['iob_units']:.2f} Ü")

			if calc.get("tdd_kg_note"):
				ui.caption(f"⚖️ **Günlük İnsülin / Kilo (TDD/kg):** {calc['tdd_kg_note']}")

			with ui.expander("📐 Adım Adım Hesaplama Formül Dökümü (Denetim İçin)", expanded=True):
				for step in calc["formula_steps"]:
					ui.markdown(f"- {step}")

		# Gemini AI Klinik Değerlendirme Kartı
		is_sec_vis_calc = ctx_api.get("is_section_visible") or (lambda _k: True)
		ai_res_key = f"ins_ai_review_{student_id}"
		ai_comment_for_log = ""
		if is_sec_vis_calc("insulin_ai_review"):
			with ui.container(border=True):
				ui.markdown("##### 🤖 Gemini AI Pediatrik Diyabet Klinik Yorumu")
				ui.caption("Öğrencinin yaşı, boyu, kilosu, WHO persantili, TDD/kg oranı, menüsü, ekstra yiyecekleri, aktivitesi ve varsa engel/alerji profiliyle birlikte yorumlar.")
				ai_q_extra = ui.text_input(
					"💬 Gemini'ye Özel Soru / Ek Not (Opsiyonel)",
					placeholder="Örn: Öğrenci biraz halsiz görünüyor, 1 saat sonra futbol oynayacak, nelere dikkat edelim?",
					key=f"ins_ai_q_{student_id}",
				)
				if ui.button(
					"✨ Gemini AI ile Dozu, Büyüme Profilini & Menüyü Yorumla",
					type="primary",
					disabled=not is_gai_on,
					use_container_width=True,
					key=f"ins_btn_ai_analyze_{student_id}",
				):
					anon_ctx, _, full_nm = ctx_api["_build_person_anon_bundle"](student_id)
					app_cfg = ctx_api["get_app_settings"]()
					logs_recent_df = get_student_bolus_logs_df(connect_db_fn, student_id, days_back=14)
					recent_list = logs_recent_df[["Tarih", "Saat", "Öğün", "Kan_Sekeri", "Toplam_Karb", "Uygulanan_Doz", "Aktivite"]].head(10).to_dict(orient="records") if not logs_recent_df.empty else []
					meal_payload = {
						"ogun_turu": meal_type,
						"menu_yemekleri": selected_menu_dish_names,
						"menu_tuketim_yuzdesi": menu_consumed_pct,
						"ekstra_yiyecekler": cur_extras_final,
						"fiziksel_aktivite": act_level,
						"aktivite_suresi_dk": int(act_dur or 0),
						"aktivite_notu": act_note,
						"son_enjeksiyon_bolgesi": last_site,
					}
					with ui.spinner("Gemini yaş, boy, kilo, WHO persantil, yemek menüsü ve aktivite verilerini analiz ediyor..."):
						ok_ia, ia_data, used_ia = gai.analyze_insulin_dose_with_gemini(
							api_key=gai_key,
							model=gai_model,
							anon_student_context=anon_ctx,
							growth_info=growth,
							diabetes_profile={k: v for k, v in prof.items() if k not in ("student_id", "_exists", "physician_name")},
							calc_result=calc,
							meal_details=meal_payload,
							recent_bg_trend=recent_list,
							school_name=app_cfg.get("school_name", "Okul Reviri"),
							nurse_name=app_cfg.get("nurse_name", "Okul Revir Sorumlusu"),
							extra_instruction=ai_q_extra,
						)
					ui.session_state[ai_res_key] = {"ok": ok_ia, "data": ia_data, "model": used_ia}
					ui.rerun()

				ia_saved = ui.session_state.get(ai_res_key)
				if isinstance(ia_saved, dict):
					if not ia_saved.get("ok"):
						ui.error(f"Gemini AI hatası: {ia_saved.get('data', {}).get('error', 'Bilinmeyen hata')}")
					else:
						d_ia = ia_saved["data"]
						badge = d_ia.get("safety_badge", "✅ Değerlendirildi")
						ui.markdown(f"**{badge}** <small>({ia_saved.get('model')})</small>", unsafe_allow_html=True)
						ui.markdown(f"📏 **Yaş / Boy / Kilo & TDD/kg:** {d_ia.get('anthropometry_tdd_comment', '—')}")
						ui.markdown(f"🍽️ **Menü, Ekstralar & Glisemik Analiz:** {d_ia.get('meal_glycemic_comment', '—')}")
						ui.markdown(f"🏃 **Aktivite & Hipoglisemi Kalkanı:** {d_ia.get('activity_hypo_comment', '—')}")
						ui.info(f"👩‍⚕️ **Hemşire İzlem Planı:** {d_ia.get('nurse_action_plan', '—')}")
						sms_raw = str(d_ia.get("guardian_sms_draft") or "")
						if sms_raw:
							sms_final = sms_raw.replace("[ÖĞRENCİ_ADI]", disp_name).replace("[Öğrenci Adı]", disp_name)
							ui.text_area("📱 Veli Bilgilendirme (WhatsApp / SMS) Taslağı", value=sms_final, height=90, key=f"ins_sms_draft_{student_id}")
						ai_comment_for_log = (
							f"[{badge}] {d_ia.get('anthropometry_tdd_comment', '')} | "
							f"{d_ia.get('meal_glycemic_comment', '')} | {d_ia.get('nurse_action_plan', '')}"
						).strip(" |")

		# Uygulama & Rotasyon Kayıt Formu
		with ui.container(border=True):
			ui.markdown("##### 💉 Uygulamayı Onayla & İnsülin Defterine Kaydet")
			# Enjeksiyon bölgesi rotasyon önerisi + uzuv engeli uyarısı
			def_site_idx = 0
			if last_site in INJECTION_SITE_OPTIONS:
				def_site_idx = (INJECTION_SITE_OPTIONS.index(last_site) + 1) % len(INJECTION_SITE_OPTIONS)
				ui.caption(f"🔄 **Rotasyon Önerisi:** Son enjeksiyon `{last_site}` bölgesine yapılmıştı; lipohipertrofiyi önlemek için sıradaki farklı bölge önerildi.")
			if dp_dis.get("has_disability"):
				lmap = dp_dis.get("limb_status") or {}
				limb_warns = []
				if str(lmap.get("right_arm") or "").startswith(("Doğuştan", "Kullanamıyor")):
					limb_warns.append("Sağ Kol")
				if str(lmap.get("left_arm") or "").startswith(("Doğuştan", "Kullanamıyor")):
					limb_warns.append("Sol Kol")
				if limb_warns:
					ui.warning(f"♿ **Uzuv Profili Hatırlatması:** Öğrencinin profilinde **{', '.join(limb_warns)}** kısıtlılığı/yokluğu kayıtlı — uygun farklı enjeksiyon bölgesi seçin.")

			ap1, ap2 = ui.columns(2)
			applied_dose = ap1.number_input(
				"Uygulanan Gerçek Doz (Ü)",
				min_value=0.0,
				max_value=60.0,
				step=float(prof.get("pen_increment") or 0.5),
				value=float(calc["recommended_dose_units"]),
				key=f"ins_applied_dose_{student_id}_{calc['recommended_dose_units']}",
			)
			inj_site = ap2.selectbox("Enjeksiyon Bölgesi", INJECTION_SITE_OPTIONS, index=def_site_idx, key=f"ins_inj_site_{student_id}")
			log_note = ui.text_input("Hemşire Uygulama Notu", placeholder="Örn: Yemekten 10 dk önce yapıldı, 2 saat sonra KŞ kontrolü planlandı", key=f"ins_log_note_{student_id}")
			sync_vital_chk = ui.checkbox("💓 Ölçülen kan şekerini 'Vital Takip' kayıtlarına da işle", value=True, key=f"ins_sync_vital_{student_id}")

			if ui.button("💾 İnsülin Dozunu & Kan Şekerini Kaydet", type="primary", use_container_width=True, key=f"ins_btn_save_log_{student_id}"):
				app_cfg = ctx_api["get_app_settings"]()
				nurse_nm = app_cfg.get("nurse_name") or "Okul Revir Sorumlusu"
				meal_sum_parts = []
				if selected_menu_dish_names:
					meal_sum_parts.append(f"Menü (%{menu_consumed_pct:.0f}): " + ", ".join(selected_menu_dish_names))
				elif menu_carbs_raw > 0:
					meal_sum_parts.append(f"Manuel öğün: {menu_carbs_raw:.1f}g")
				if cur_extras_final:
					meal_sum_parts.append("Ekstra: " + ", ".join(f"{x['name']} ({x['carbs_g']:.0f}g)" for x in cur_extras_final))
				if act_note:
					meal_sum_parts.append(f"Aktivite: {act_note}")

				with connect_db_fn() as conn:
					ensure_diabetes_tables(conn)
					conn.execute(
						"""INSERT INTO insulin_bolus_logs(
							student_id, log_date, log_time, meal_type, bg_mg_dl, ketone_status,
							menu_carbs_g, menu_consumed_pct, extra_foods_json, extra_carbs_g, total_carbs_g,
							activity_level, activity_duration_min, activity_reduction_pct,
							icr_used, isf_used, iob_units, carb_dose_units, correction_dose_units,
							calculated_dose_units, applied_dose_units, injection_site,
							meal_summary, ai_commentary, administered_by, notes, created_at
						) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(
							student_id,
							sel_date.isoformat(),
							sel_time.strftime("%H:%M"),
							meal_type,
							bg_val,
							ketone_st,
							calc["effective_menu_carbs_g"],
							float(menu_consumed_pct),
							json.dumps(cur_extras_final, ensure_ascii=False),
							calc["extra_carbs_g"],
							calc["total_carbs_g"],
							act_level,
							int(act_dur or 0),
							calc["activity_reduction_pct"],
							calc["icr_used"],
							calc["isf_used"],
							calc["iob_units"],
							calc["adjusted_carb_dose"],
							calc["net_correction_dose"],
							calc["recommended_dose_units"],
							float(applied_dose),
							inj_site,
							" | ".join(meal_sum_parts),
							ai_comment_for_log,
							nurse_nm,
							log_note.strip(),
							datetime.now().isoformat(timespec="minutes"),
						),
					)
					if sync_vital_chk and bg_val is not None:
						g_ctx = "Açlık / Öğün Öncesi" if meal_type in ("Kahvaltı", "Öğle Yemeği") else "Rastgele / Ara Öğün"
						conn.execute(
							"""INSERT INTO vital_readings(student_id, measured_at, glucose_mg_dl, glucose_context, notes)
							   VALUES (?, ?, ?, ?, ?)""",
							(
								student_id,
								f"{sel_date.isoformat()} {sel_time.strftime('%H:%M')}",
								bg_val,
								g_ctx,
								f"İnsülin modülü ({meal_type}) · Uygulanan: {applied_dose:.2f} Ü ({inj_site})",
							),
						)
				ui.session_state[extras_key] = []
				ctx_api["sync_after_person_mutation"](student_id)
				ui.session_state[flash_k] = f"✅ {applied_dose:.2f} Ü insülin uygulaması ({inj_site}) kaydedildi!"
				ui.rerun()


# ==============================================================================
# SEKME 2: AYLIK YEMEK PROGRAMI & GEMİNİ KARBONHİDRAT ÇIKARICI + AYLIK PLAN
# ==============================================================================

def _render_tab_monthly_menu(
	ctx_api: dict[str, Any],
	person: sqlite3.Row,
	student_id: int,
	prof: dict[str, Any],
) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_gai_on, gai_key, gai_model = ctx_api["_get_active_gemini_config"]()
	disp_name = ctx_api["mask_kvkk_value"](person["full_name"], "name")

	today = date.today()
	ym_options: list[str] = []
	for offset in range(-2, 7):
		m_tot = today.year * 12 + (today.month - 1) + offset
		y_v, m_v = divmod(m_tot, 12)
		ym_options.append(f"{y_v:04d}-{m_v + 1:02d}")
	cur_ym_default = f"{today.year:04d}-{today.month:02d}"

	def _fmt_ym(ym_s: str) -> str:
		try:
			yy, mm = ym_s.split("-")
			return f"{TR_MONTH_NAMES.get(int(mm), mm)} {yy} ({ym_s})"
		except Exception:
			return ym_s

	ym_c1, ym_c2 = ui.columns([1.6, 2.4])
	with ym_c1:
		sel_ym = ui.selectbox(
			"📅 Çalışılacak Ay (Yemek Programı)",
			ym_options,
			index=ym_options.index(cur_ym_default) if cur_ym_default in ym_options else 2,
			format_func=_fmt_ym,
			key="ins_menu_sel_ym",
		)
	with ym_c2:
		ui.caption("Okulun aylık yemek listesini **PDF, Fotoğraf, Excel, Word veya Metin** olarak yükleyin; Gemini AI her günün yemeklerini ve tahmini karbonhidratlarını otomatik çıkarıp takvime işler.")

	flash_m = ui.session_state.pop("ins_menu_flash", None)
	if flash_m:
		ui.success(flash_m)

	# A. Gemini AI Aylık Menü Yükleyici
	with ui.expander("✨ Gemini AI ile Aylık Yemek Programını Otomatik Oku & Karbonhidrat Çıkar", expanded=True):
		mc1, mc2 = ui.columns([1.6, 1.4])
		with mc1:
			up_menu_file = ui.file_uploader(
				"📎 Aylık Yemek Listesi Dosyası (.pdf, .png, .jpg, .xlsx, .docx)",
				type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "docx", "txt"],
				key="ins_menu_ai_file",
			)
		with mc2:
			def_mtype_ai = ui.selectbox("Varsayılan Öğün Türü", ["Öğle Yemeği", "Kahvaltı", "İkindi / Ara Öğün"], index=0, key="ins_menu_ai_mtype")
		paste_menu_txt = ui.text_area(
			"📋 Veya Aylık Yemek Listesi Metnini Buraya Yapıştırın",
			height=105,
			placeholder="Örn:\n02.11.2026: Mercimek Çorbası, Orman Kebabı, Pirinç Pilavı, Ayran\n03.11.2026: Tarhana Çorbası, Fırın Köfte, Makarna, Mandalina...",
			key="ins_menu_ai_text",
		)
		if ui.button(
			"✨ Gemini AI ile Menüyü Çözümle ve Takvime Kaydet",
			type="primary",
			disabled=not is_gai_on,
			use_container_width=True,
			key="ins_btn_parse_menu_ai",
		):
			fb = up_menu_file.getvalue() if up_menu_file is not None else None
			fn = up_menu_file.name if up_menu_file is not None else ""
			with ui.spinner("Gemini aylık yemek listesini gün gün ayrıştırıyor ve Türk porsiyon karbonhidratlarını hesaplıyor..."):
				ok_pm, parsed_days, used_pm = gai.parse_monthly_school_menu_with_gemini(
					api_key=gai_key,
					model=gai_model,
					year_month=sel_ym,
					file_bytes=fb,
					file_name=fn,
					raw_text=paste_menu_txt,
					default_meal_type=def_mtype_ai,
				)
			if ok_pm and parsed_days:
				for day_obj in parsed_days:
					upsert_school_meal_menu(
						connect_db_fn,
						day_obj["menu_date"],
						day_obj["meal_type"],
						day_obj["items"],
						source_note=f"Gemini AI ({used_pm})",
					)
				ui.session_state["ins_menu_flash"] = f"✅ {len(parsed_days)} günlük yemek menüsü ve karbonhidrat değerleri Gemini ({used_pm}) ile başarıyla kaydedildi!"
				ui.rerun()
			else:
				ui.error(f"Menü ayrıştırma hatası: {used_pm}")

	# B. Manuel Günlük Menü Ekle / Düzenle
	with ui.expander("✏️ Tek Günlük Yemek Menüsü Elle Ekle / Düzenle", expanded=False):
		with ui.form("form_manual_daily_menu"):
			mm1, mm2 = ui.columns(2)
			man_date = mm1.date_input("Menü Tarihi", value=date.today(), key="ins_man_mdate")
			man_type = mm2.selectbox("Öğün Türü", ["Öğle Yemeği", "Kahvaltı", "İkindi / Ara Öğün"], key="ins_man_mtype")
			man_lines = ui.text_area(
				"Yemekler (Her satıra: Yemek Adı | Karbonhidrat gram | Porsiyon — örn: Mercimek Çorbası | 18 | 1 kase)",
				height=115,
				placeholder="Mercimek Çorbası | 18 | 1 kase\nEtli Kuru Fasulye | 22 | 1 porsiyon\nPirinç Pilavı | 30 | 4 kaşık\nAyran | 4 | 1 kutu",
				key="ins_man_mlines",
			)
			if ui.form_submit_button("💾 Günlük Menüyü Kaydet", type="primary", use_container_width=True):
				man_items: list[dict[str, Any]] = []
				for ln in (man_lines or "").splitlines():
					s = ln.strip()
					if not s:
						continue
					parts = [p.strip() for p in s.split("|")]
					nm = parts[0]
					try:
						cg = float(parts[1].replace(",", ".").replace("g", "").strip()) if len(parts) > 1 else 15.0
					except Exception:
						cg = 15.0
					por = parts[2] if len(parts) > 2 else "1 porsiyon"
					man_items.append({"name": nm, "portion": por, "carbs_g": max(0.0, round(cg, 1)), "gi": "Orta", "fat_protein_note": ""})
				if man_items:
					upsert_school_meal_menu(connect_db_fn, man_date.isoformat(), man_type, man_items, source_note="Manuel Giriş")
					ui.session_state["ins_menu_flash"] = f"✅ {man_date.strftime('%d.%m.%Y')} ({man_type}) menüsü kaydedildi!"
					ui.rerun()
				else:
					ui.warning("Lütfen en az bir yemek satırı girin.")

	# C. Seçili Öğrenci İçin Aylık Yemek & Tahmini Öğün İnsülin Planı Tablosu + Word/Excel İndirme
	plan_df, _ = build_monthly_insulin_plan_df(connect_db_fn, student_id, sel_ym)
	ui.markdown(f"#### 📊 {ctx_api['mask_kvkk_value'](person['full_name'], 'name')} — {_fmt_ym(sel_ym)} Aylık Menü & Tahmini İnsülin Planı")
	if plan_df.empty:
		ui.info(f"`{_fmt_ym(sel_ym)}` ayı için henüz kayıtlı okul yemek menüsü yok. Yukarıdan aylık yemek listesini yükleyin veya elle gün ekleyin.")
	else:
		ui.caption("Tablodaki dozlar **yalnızca öğün karbonhidratına** göre öğrencinin profilindeki İKO ile hesaplanmış öngörüdür (yüksek şeker düzeltmesi ve IOB günlük ölçümde ayrıca eklenir).")
		ui.dataframe(plan_df, width="stretch", hide_index=True)

		p_slug = re.sub(r"[^A-Za-z0-9_-]+", "_", str(person["full_name"] or "ogrenci")).strip("_") or "ogrenci"
		plan_title = f"{person['full_name']} ({person['class_name']}) — {_fmt_ym(sel_ym)} Aylık Yemek & İnsülin Planı"
		plan_sub = f"İKO (Sabah/Öğle/Ara): 1/{prof['icr_breakfast']:.0f}g · 1/{prof['icr_lunch']:.0f}g · 1/{prof['icr_snack']:.0f}g | İDF: {prof['isf']:.0f} | Hedef: {prof['target_bg']:.0f} mg/dL | İnsülin: {prof['insulin_brand']}"

		dl1, dl2 = ui.columns(2)
		with dl1:
			ui.download_button(
				"📄 Aylık İnsülin & Yemek Planını Word İndir (.docx)",
				data=lambda _t=plan_title, _s=plan_sub, _df=plan_df.copy(): ctx_api["build_ai_custom_report_docx"](_t, subtitle=_s, df=_df),
				file_name=f"{p_slug}_aylik_insulin_plani_{sel_ym}.docx",
				mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				type="primary",
				use_container_width=True,
				key=f"ins_dl_plan_docx_{student_id}_{sel_ym}",
			)
		with dl2:
			ui.download_button(
				"📊 Aylık İnsülin & Yemek Planını Excel İndir (.xlsx)",
				data=lambda _t=plan_title, _df=plan_df.copy(): ctx_api["build_ai_custom_report_excel"](_t, df=_df),
				file_name=f"{p_slug}_aylik_insulin_plani_{sel_ym}.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				type="secondary",
				use_container_width=True,
				key=f"ins_dl_plan_xlsx_{student_id}_{sel_ym}",
			)

		# Kayıtlı gün silme kolaylığı
		raw_menus = get_school_menus_for_month(connect_db_fn, sel_ym)
		with ui.expander("🗑️ Kayıtlı Menü Günü Sil", expanded=False):
			del_opts = {f"{m['menu_date']} · {m['meal_type']} ({m['total_carbs_g']:.0f}g)": int(m["id"]) for m in raw_menus}
			del_lbl = ui.selectbox("Silinecek Menü Günü", list(del_opts.keys()), key="ins_del_menu_sel")
			if ui.button("🗑️ Seçili Günü Sil", type="secondary", key="ins_del_menu_btn"):
				delete_school_menu_by_id(connect_db_fn, del_opts[del_lbl])
				ui.rerun()


# ==============================================================================
# SEKME 3: ÖĞRENCİ DİYABET & HEKİM REÇETE PARAMETRELERİ
# ==============================================================================

def _render_tab_diabetes_profile(
	ctx_api: dict[str, Any],
	person: sqlite3.Row,
	student_id: int,
	prof: dict[str, Any],
	growth: dict[str, Any],
) -> None:
	connect_db_fn = ctx_api["connect_db"]
	ui.markdown(f"#### ⚙️ {ctx_api['mask_kvkk_value'](person['full_name'], 'name')} — Hekim Diyabet & İnsülin Parametreleri")
	ui.caption("Çocuk endokrinoloji hekiminin okul diyabet bakım planında belirttiği İKO, İDF, hedef kan şekeri ve güvenlik sınırlarını girin.")

	# 500/1800 pediatrik referans yardımcı bilgi kutusu (TDD giriliyse)
	cur_tdd = float(prof.get("tdd_units") or 0.0)
	if cur_tdd > 0:
		ref_icr = round(500.0 / cur_tdd, 1)
		ref_isf = round(1800.0 / cur_tdd, 0)
		kg_txt = f" · TDD/kg: **{cur_tdd / growth['kilo_kg']:.2f} Ü/kg**" if growth.get("kilo_kg") else ""
		ui.info(f"📐 **Pediatrik Referans Kontrolü (TDD={cur_tdd:.1f} Ü{kg_txt}):** 500 Kuralı tahmini İKO ≈ `1/{ref_icr}g` · 1800 Kuralı tahmini İDF ≈ `{ref_isf:.0f} mg/dL/Ü` *(Daima hekimin yazılı planı esas alınır)*.")

	with ui.form(f"form_diab_prof_{student_id}"):
		p1, p2, p3 = ui.columns(3)
		f_is_diab = p1.checkbox("💙 Diyabet Takibi Aktif", value=bool(prof.get("is_diabetic", True) or not prof.get("_exists")), key=f"dp_isd_{student_id}")
		cur_dt = prof.get("diabetes_type") or DIABETES_TYPE_OPTIONS[0]
		dt_idx = DIABETES_TYPE_OPTIONS.index(cur_dt) if cur_dt in DIABETES_TYPE_OPTIONS else 0
		f_dtype = p2.selectbox("Diyabet Türü", DIABETES_TYPE_OPTIONS, index=dt_idx, key=f"dp_dtype_{student_id}")

		cur_br = prof.get("insulin_brand") or INSULIN_BRAND_OPTIONS[0]
		br_opts = INSULIN_BRAND_OPTIONS if cur_br in INSULIN_BRAND_OPTIONS else [cur_br, *INSULIN_BRAND_OPTIONS]
		f_brand = p3.selectbox("Hızlı Etkili (Bolus) İnsülin", br_opts, index=br_opts.index(cur_br), key=f"dp_brand_{student_id}")

		q1, q2, q3 = ui.columns(3)
		cur_deliv = prof.get("delivery_method") or DELIVERY_METHOD_OPTIONS[0]
		dv_opts = DELIVERY_METHOD_OPTIONS if cur_deliv in DELIVERY_METHOD_OPTIONS else [cur_deliv, *DELIVERY_METHOD_OPTIONS]
		f_deliv = q1.selectbox("Uygulama Aracı & Hassasiyet", dv_opts, index=dv_opts.index(cur_deliv), key=f"dp_deliv_{student_id}")
		f_basal = q2.text_input("Bazal (Uzun Etkili) İnsülin", value=str(prof.get("basal_insulin") or ""), placeholder="Örn: Lantus 12 Ü akşam / Tresiba / Pompa bazal", key=f"dp_basal_{student_id}")
		f_tdd = q3.number_input("Günlük Toplam İnsülin — TDD (Ü)", min_value=0.0, max_value=200.0, step=0.5, value=float(prof.get("tdd_units") or 0.0), key=f"dp_tdd_{student_id}")

		ui.markdown("##### 📐 İnsülin / Karbonhidrat Oranları (İKO — 1 Ünite kaç gram karbonhidrat karşılar?)")
		ic1, ic2, ic3 = ui.columns(3)
		f_icr_b = ic1.number_input("🌅 Sabah Kahvaltı İKO (g/Ü)", min_value=2.0, max_value=60.0, step=0.5, value=float(prof.get("icr_breakfast") or 15.0), key=f"dp_icrb_{student_id}")
		f_icr_l = ic2.number_input("☀️ Öğle Yemeği İKO (g/Ü)", min_value=2.0, max_value=60.0, step=0.5, value=float(prof.get("icr_lunch") or 15.0), key=f"dp_icrl_{student_id}")
		f_icr_s = ic3.number_input("🍎 İkindi / Ara Öğün İKO (g/Ü)", min_value=2.0, max_value=60.0, step=0.5, value=float(prof.get("icr_snack") or 18.0), key=f"dp_icrs_{student_id}")

		ui.markdown("##### 🎯 İDF, Hedef Kan Şekeri ve Güvenlik Bariyerleri")
		sf1, sf2, sf3, sf4 = ui.columns(4)
		f_isf = sf1.number_input("İDF (1 Ü kaç mg/dL düşürür?)", min_value=10.0, max_value=250.0, step=1.0, value=float(prof.get("isf") or 50.0), key=f"dp_isf_{student_id}")
		f_tgt = sf2.number_input("Hedef Kan Şekeri (mg/dL)", min_value=80.0, max_value=160.0, step=5.0, value=float(prof.get("target_bg") or 120.0), key=f"dp_tgt_{student_id}")
		f_cth = sf3.number_input("Düzeltme Başlangıç Eşiği", min_value=110.0, max_value=220.0, step=5.0, value=float(prof.get("correction_threshold") or 150.0), key=f"dp_cth_{student_id}")
		f_hypo = sf4.number_input("Hipoglisemi Eşiği (mg/dL)", min_value=55.0, max_value=90.0, step=5.0, value=float(prof.get("hypo_threshold") or 70.0), key=f"dp_hypo_{student_id}")

		sf5, sf6, sf7 = ui.columns(3)
		f_ket_th = sf5.number_input("Keton Kontrol Eşiği (mg/dL)", min_value=200.0, max_value=350.0, step=10.0, value=float(prof.get("hyper_ketone_threshold") or 250.0), key=f"dp_ketth_{student_id}")
		f_dur = sf6.number_input("Aktif İnsülin (IOB) Süresi (Saat)", min_value=2.0, max_value=5.0, step=0.5, value=float(prof.get("insulin_duration_hours") or 3.5), key=f"dp_dur_{student_id}")
		f_maxb = sf7.number_input("Maksimum Tek Bolus Sınırı (Ü)", min_value=1.0, max_value=40.0, step=0.5, value=float(prof.get("max_single_bolus") or 12.0), key=f"dp_maxb_{student_id}")

		ui.markdown("##### 🆘 Sensör, Glukagon & Hekim İletişim Bilgileri")
		ex1, ex2, ex3 = ui.columns(3)
		f_cgm = ex1.text_input("CGM / Sensör Modeli", value=str(prof.get("cgm_device") or ""), placeholder="Örn: Dexcom G7 / FreeStyle Libre 2", key=f"dp_cgm_{student_id}")
		f_gluc = ex2.text_input("Okuldaki Glukagon / Baqsimi Konumu", value=str(prof.get("glucagon_location") or ""), placeholder="Örn: Revir dolabı üst raf + sınıf acil çantası", key=f"dp_gluc_{student_id}")
		f_doc = ex3.text_input("Takip Eden Endokrinoloji Hekimi", value=str(prof.get("physician_name") or ""), placeholder="Örn: Prof. Dr. ... / Şehir Hastanesi", key=f"dp_doc_{student_id}")

		f_proto = ui.text_area(
			"📋 Hekim Özel Talimatları & Acil Eylem Planı",
			value=str(prof.get("emergency_protocol") or ""),
			height=85,
			placeholder="Örn: Yemekten 10 dk önce uygula; KŞ < 80 ise yemek sonrası uygula; beden eğitimi öncesi KŞ < 100 ise 15g meyve suyu ver...",
			key=f"dp_proto_{student_id}",
		)
		sync_ch_chk = ui.checkbox("📋 Sağlık profilindeki 'Kronik Hastalıklar' alanına da diyabet tanısını otomatik işle", value=True, key=f"dp_sync_ch_{student_id}")
		save_dp = ui.form_submit_button("💾 Öğrenci Diyabet & İnsülin Profilini Kaydet", type="primary", use_container_width=True)

	if save_dp:
		save_diabetes_profile(
			connect_db_fn,
			student_id,
			{
				"is_diabetic": f_is_diab,
				"diabetes_type": f_dtype,
				"insulin_brand": f_brand,
				"basal_insulin": f_basal,
				"delivery_method": f_deliv,
				"pen_increment": PEN_INCREMENT_MAP.get(f_deliv, 0.5),
				"icr_breakfast": f_icr_b,
				"icr_lunch": f_icr_l,
				"icr_snack": f_icr_s,
				"isf": f_isf,
				"target_bg": f_tgt,
				"correction_threshold": f_cth,
				"hypo_threshold": f_hypo,
				"hyper_ketone_threshold": f_ket_th,
				"insulin_duration_hours": f_dur,
				"tdd_units": f_tdd,
				"max_single_bolus": f_maxb,
				"cgm_device": f_cgm,
				"glucagon_location": f_gluc,
				"physician_name": f_doc,
				"emergency_protocol": f_proto,
			},
			sync_chronic=sync_ch_chk,
		)
		ctx_api["sync_after_person_mutation"](student_id)
		ui.success("✅ Öğrenci diyabet ve insülin reçete parametreleri kaydedildi!")
		ui.rerun()


# ==============================================================================
# SEKME 4: İNSÜLİN & KAN ŞEKERİ DEFTERİ, TIR TRENDİ, ROTASYON & RAPORLAR
# ==============================================================================

def _render_tab_logs_and_trends(
	ctx_api: dict[str, Any],
	person: sqlite3.Row,
	student_id: int,
	prof: dict[str, Any],
	growth: dict[str, Any],
) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_gai_on, gai_key, gai_model = ctx_api["_get_active_gemini_config"]()
	disp_name = ctx_api["mask_kvkk_value"](person["full_name"], "name")

	p_c1, p_c2 = ui.columns([1.2, 2.8])
	days_back = p_c1.selectbox("İzlem Aralığı", [7, 14, 30, 90, 365], index=2, format_func=lambda d: f"Son {d} Gün", key=f"ins_log_days_{student_id}")
	logs_df = get_student_bolus_logs_df(connect_db_fn, student_id, days_back=int(days_back))
	tir = compute_glycemic_tir_stats(logs_df, hypo_thresh=float(prof.get("hypo_threshold") or 70.0), hyper_high=180.0)

	k1, k2, k3, k4, k5 = ui.columns(5)
	k1.metric("Toplam Uygulama", f"{tir['total_logs']}")
	k2.metric("Ortalama KŞ", f"{tir['avg_bg']} mg/dL" if tir["avg_bg"] is not None else "—")
	k3.metric("🎯 Hedef Aralık (TIR)", f"%{tir['tir_pct']}" if tir["tir_pct"] is not None else "—")
	k4.metric("🔻 Hipoglisemi (<70)", f"{tir['hypo_cnt']}")
	k5.metric("🔺 Hiperglisemi (>180)", f"{tir['hyper_cnt']}")

	if logs_df.empty:
		ui.info("Seçili aralıkta henüz kaydedilmiş insülin veya kan şekeri uygulaması bulunmuyor.")
		return

	# Görsel Kan Şekeri Trend Grafiği (Hedef 70-180 bandı — :T Zaman Ekseni)
	bg_chart_df = logs_df.dropna(subset=["Kan_Sekeri"]).copy()
	if not bg_chart_df.empty:
		try:
			import altair as alt
			bg_chart_df["Zaman_Str"] = bg_chart_df["Tarih"].astype(str) + " " + bg_chart_df["Saat"].astype(str)
			bg_chart_df["Zaman"] = pd.to_datetime(bg_chart_df["Zaman_Str"], errors="coerce")
			bg_chart_df = bg_chart_df.dropna(subset=["Zaman"]).sort_values("Zaman")
			if not bg_chart_df.empty:
				line = alt.Chart(bg_chart_df).mark_line(point=True, color="#0d9488", strokeWidth=2.5).encode(
					x=alt.X("Zaman:T", title="Tarih & Saat"),
					y=alt.Y("Kan_Sekeri:Q", title="Kan Şekeri (mg/dL)"),
					tooltip=["Tarih", "Saat", "Öğün", "Kan_Sekeri", "Toplam_Karb", "Uygulanan_Doz", "Enjeksiyon_Bolgesi"],
				)
				rule_lo = alt.Chart(pd.DataFrame({"y": [float(prof.get("hypo_threshold") or 70)]})).mark_rule(color="#ef4444", strokeDash=[4, 4]).encode(y="y:Q")
				rule_hi = alt.Chart(pd.DataFrame({"y": [180.0]})).mark_rule(color="#f59e0b", strokeDash=[4, 4]).encode(y="y:Q")
				ui.altair_chart(alt.layer(line, rule_lo, rule_hi).properties(height=280), use_container_width=True)
		except Exception:
			pass

	# Enjeksiyon Bölgesi Rotasyon Özeti
	if tir["site_counts"]:
		site_pills = " · ".join(f"`{k}`: **{v}** kez" for k, v in tir["site_counts"].items())
		ui.markdown(f"🔄 **Enjeksiyon Bölgesi Rotasyon Dağılımı:** {site_pills}")

	# Tablo
	disp_df = logs_df.drop(columns=["id"], errors="ignore")
	ui.dataframe(disp_df, width="stretch", hide_index=True)

	# Gemini AI 14-30 Günlük Glisemik Trend & Hekim Kontrol Özeti
	trend_key = f"ins_ai_trend_{student_id}"
	if is_gai_on:
		if ui.button("✨ Gemini AI ile Glisemik Trend & Endokrinoloji Kontrol Özeti Oluştur", type="primary", use_container_width=True, key=f"ins_btn_ai_trend_{student_id}"):
			anon_ctx, _, _ = ctx_api["_build_person_anon_bundle"](student_id)
			logs_rec = disp_df[["Tarih", "Saat", "Öğün", "Kan_Sekeri", "Toplam_Karb", "Aktivite", "Uygulanan_Doz", "Enjeksiyon_Bolgesi"]].head(30).to_dict(orient="records")
			with ui.spinner("Gemini son glisemik trend ve TIR verilerini analiz ediyor..."):
				ok_tr, tr_md, used_tr = gai.analyze_glycemic_trend_with_gemini(
					api_key=gai_key,
					model=gai_model,
					anon_student_context=anon_ctx,
					growth_info=growth,
					tir_stats=tir,
					logs_summary=logs_rec,
				)
			ui.session_state[trend_key] = {"ok": ok_tr, "text": tr_md, "model": used_tr}
			ui.rerun()

	saved_tr = ui.session_state.get(trend_key)
	ai_trend_md = ""
	if isinstance(saved_tr, dict):
		if saved_tr.get("ok"):
			ai_trend_md = str(saved_tr.get("text") or "")
			with ui.container(border=True):
				ui.markdown(f"##### 🤖 Gemini AI Glisemik Trend Özeti <small>({saved_tr.get('model')})</small>", unsafe_allow_html=True)
				ui.markdown(ai_trend_md)
		else:
			ui.error(f"Gemini AI hatası: {saved_tr.get('text')}")

	# Word & Excel Rapor Çıktıları
	p_slug = re.sub(r"[^A-Za-z0-9_-]+", "_", str(person["full_name"] or "ogrenci")).strip("_") or "ogrenci"
	rep_title = f"{person['full_name']} ({person['class_name']}) — Okul İçi İnsülin & Kan Şekeri İzlem Raporu"
	rep_sub = (
		f"Son {days_back} Gün | Ort. KŞ: {tir['avg_bg'] or '—'} mg/dL | TIR (%70-180): %{tir['tir_pct'] if tir['tir_pct'] is not None else '—'} | "
		f"Hipo: {tir['hypo_cnt']} · Hiper: {tir['hyper_cnt']} | İKO (S/Ö/A): 1/{prof['icr_breakfast']:.0f}·1/{prof['icr_lunch']:.0f}·1/{prof['icr_snack']:.0f}"
	)
	secs = [{"heading": "🤖 Glisemik Trend & Klinik Değerlendirme (Gemini AI)", "body": ai_trend_md}] if ai_trend_md else None

	rd1, rd2 = ui.columns(2)
	with rd1:
		ui.download_button(
			"📄 İnsülin & Kan Şekeri İzlem Raporu İndir (.docx)",
			data=lambda _t=rep_title, _s=rep_sub, _sc=secs, _df=disp_df.copy(): ctx_api["build_ai_custom_report_docx"](_t, subtitle=_s, sections=_sc, df=_df),
			file_name=f"{p_slug}_insulin_izlem_raporu.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			type="primary",
			use_container_width=True,
			key=f"ins_dl_logs_docx_{student_id}",
		)
	with rd2:
		ui.download_button(
			"📊 İnsülin & Kan Şekeri Defterini Excel İndir (.xlsx)",
			data=lambda _t=rep_title, _df=disp_df.copy(), _sc=secs: ctx_api["build_ai_custom_report_excel"](_t, df=_df, sections=_sc),
			file_name=f"{p_slug}_insulin_izlem_defteri.xlsx",
			mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			type="secondary",
			use_container_width=True,
			key=f"ins_dl_logs_xlsx_{student_id}",
		)

	with ui.expander("🗑️ Hatalı İnsülin Kaydı Sil", expanded=False):
		del_l_map = {
			f"#{int(r['id'])} · {r['Tarih']} {r['Saat']} — {r['Öğün']} ({r['Uygulanan_Doz']} Ü)": int(r["id"])
			for _, r in logs_df.iterrows()
		}
		sel_del_l = ui.selectbox("Silinecek Kayıt", list(del_l_map.keys()), key=f"ins_del_log_sel_{student_id}")
		if ui.button("🗑️ Seçili Kaydı Kalıcı Olarak Sil", type="secondary", key=f"ins_del_log_btn_{student_id}"):
			with connect_db_fn() as conn:
				conn.execute("DELETE FROM insulin_bolus_logs WHERE id=?", (del_l_map[sel_del_l],))
			ui.rerun()
