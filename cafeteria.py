"""
Revir Defteri - Yemekhane Yönetimi Modülü (cafeteria.py)

Özellikler:
1. Yemekler Bankası & Malzeme / Kalori / Türk Gıda Kodeksi 14 Alerjen Motoru
2. İş Günü & Resmi Tatil Takvimine Bağlı Öğün Bazlı Aylık/Haftalık Menü Planlayıcı
3. İmzalı Resmi A4 (Word .docx & Yazdırılabilir HTML/PDF) Menü Çıktısı + Diyabet Menüsü Senkronizasyonu
4. Gelen Ürünler, Fatura / İrsaliye, Birim & Aylık / Toplam Gider Muhasebesi
5. Yemek Yiyen Abone Takibi, İş Günü × Öğün Ücreti Tahakkuku, Getiren İsimli Tahsilat & Okul Kasa Teslimi
"""
from __future__ import annotations

import base64
import calendar
import html
import io
import json
import re
import sqlite3
from datetime import date, datetime, timedelta
from typing import Any, Callable

import pandas as pd

import gemini_assistant as gai
import insulin_diabetes as idm
import revir_ui as ui


# ==============================================================================
# 1. SABİTLER & STANDART ETİKETLER
# ==============================================================================

TR_MONTH_NAMES: dict[int, str] = {
	1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan",
	5: "Mayıs", 6: "Haziran", 7: "Temmuz", 8: "Ağustos",
	9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık",
}

TR_WEEKDAYS_SHORT: dict[int, str] = {
	0: "Pazartesi", 1: "Salı", 2: "Çarşamba", 3: "Perşembe",
	4: "Cuma", 5: "Cumartesi", 6: "Pazar",
}

MEAL_SLOTS: list[str] = [
	"Sabah Kahvaltısı",
	"Öğle Yemeği",
	"İkindi Ara Öğünü",
]

RAMAZAN_EXTRA_SLOTS: list[str] = ["İftar", "Sahur"]
ALL_MEAL_SLOTS: list[str] = MEAL_SLOTS + RAMAZAN_EXTRA_SLOTS

DISH_CATEGORIES: list[str] = [
	"Çorba",
	"Ana Yemek",
	"Pilav / Makarna / Börek",
	"Salata / Meze / Yoğurt",
	"Kahvaltılık",
	"Tatlı / Meyve",
	"İçecek",
	"Ekmek / Diğer",
]

INGREDIENT_UNITS: list[str] = [
	"g", "ml", "adet", "yemek kaşığı", "çay kaşığı", "su bardağı", "porsiyon",
]

PURCHASE_CATEGORIES: list[str] = [
	"Sebze & Meyve",
	"Et / Tavuk / Balık",
	"Süt & Kahvaltılık",
	"Bakliyat & Kuru Gıda",
	"Yağ & Baharat",
	"Unlu Mamul & Ekmek",
	"İçecek",
	"Temizlik & Mutfak Sarf",
	"Diğer",
]

PURCHASE_UNITS: list[str] = [
	"kg", "g", "litre", "ml", "adet", "koli", "paket", "bağ", "teneke", "çuval",
]

DINER_ROLE_TYPES: list[str] = [
	"Öğrenci",
	"Öğretmen",
	"Personel",
	"Burslu / Ücretsiz",
	"Misafir",
]

TR_ALLERGEN_OPTIONS: list[str] = [
	"",
	"Gluten (Buğday/Arpa)",
	"Süt ve Süt Ürünleri (Laktoz)",
	"Yumurta",
	"Yer Fıstığı",
	"Sert Kabuklu Yemişler",
	"Soya",
	"Balık",
	"Kabuklular",
	"Yumuşakçalar",
	"Susam",
	"Hardal",
	"Kereviz",
	"Acı Bakla",
	"Sülfit",
]

ALLERGEN_BADGE_MAP: dict[str, str] = {
	"Gluten (Buğday/Arpa)": "🌾 Gluten",
	"Süt ve Süt Ürünleri (Laktoz)": "🥛 Süt/Laktoz",
	"Yumurta": "🥚 Yumurta",
	"Yer Fıstığı": "🥜 Yer Fıstığı",
	"Sert Kabuklu Yemişler": "🌰 Kuruyemiş",
	"Soya": "🫘 Soya",
	"Balık": "🐟 Balık",
	"Kabuklular": "🦐 Kabuklu",
	"Yumuşakçalar": "🐚 Yumuşakça",
	"Susam": "⚪ Susam",
	"Hardal": "🟡 Hardal",
	"Kereviz": "🥬 Kereviz",
	"Acı Bakla": "🌱 Acı Bakla",
	"Sülfit": "⚗️ Sülfit",
}

_ALLERGEN_SYNONYMS: dict[str, tuple[str, ...]] = {
	"Gluten (Buğday/Arpa)": ("gluten", "buğday", "bugday", "arpa", "çölyak", "colyak", "un"),
	"Süt ve Süt Ürünleri (Laktoz)": ("süt", "sut", "laktoz", "yoğurt", "yogurt", "peynir", "kazein", "süt ürünü"),
	"Yumurta": ("yumurta", "egg"),
	"Yer Fıstığı": ("yer fıstığı", "yer fistigi", "fıstık", "fistik"),
	"Sert Kabuklu Yemişler": ("ceviz", "fındık", "findik", "badem", "antep", "kuruyemiş", "kaju"),
	"Soya": ("soya",),
	"Balık": ("balık", "balik", "ton", "somon"),
	"Kabuklular": ("karides", "kabuklu"),
	"Yumuşakçalar": ("midye", "yumuşakça"),
	"Susam": ("susam", "tahin"),
	"Hardal": ("hardal",),
	"Kereviz": ("kereviz",),
	"Acı Bakla": ("acı bakla", "lupen", "lüpen"),
	"Sülfit": ("sülfit", "sulfit", "kükürt"),
}


def _tr_fold(val: Any) -> str:
	s = re.sub(r"\s+", " ", str(val or "").replace("\xa0", " ")).strip()
	return s.replace("İ", "i").replace("I", "ı").lower().replace("\u0307", "").translate(str.maketrans("çğıöşü", "cgiosu"))


def _safe_float(val: Any, default: float = 0.0) -> float:
	if val is None or (isinstance(val, float) and pd.isna(val)):
		return default
	s = str(val).strip().replace(",", ".")
	if not s or s.lower() in {"nan", "none", "<na>"}:
		return default
	try:
		return float(s)
	except Exception:
		m = re.search(r"-?\d+(?:\.\d+)?", s)
		return float(m.group(0)) if m else default


def _normalize_allergen_label(raw: Any) -> str:
	txt = str(raw or "").strip()
	if not txt or txt.lower() in {"yok", "none", "nan", "-"}:
		return ""
	if txt in ALLERGEN_BADGE_MAP:
		return txt
	f_txt = _tr_fold(txt)
	for canon, syns in _ALLERGEN_SYNONYMS.items():
		if _tr_fold(canon) == f_txt or any(_tr_fold(sy) in f_txt for sy in syns):
			return canon
	return txt


def format_allergen_badge(allergen: str) -> str:
	canon = _normalize_allergen_label(allergen)
	return ALLERGEN_BADGE_MAP.get(canon, f"⚠️ {canon}" if canon else "")


_MEAL_TIME_SETTING_KEYS: dict[str, tuple[str, str]] = {
	"Sabah Kahvaltısı": ("cafeteria_time_breakfast", "08:00 - 08:45"),
	"Öğle Yemeği": ("cafeteria_time_lunch", "12:00 - 13:00"),
	"İkindi Ara Öğünü": ("cafeteria_time_snack", "15:00 - 15:30"),
}


def resolve_province_name(raw_prov: str | None, fallback: str = "Ankara") -> str:
	"""Verilen il adını 81 il koordinat sözlüğünde Türkçe karakter uyumlu olarak eşler."""
	clean = re.sub(r"\s+", " ", str(raw_prov or "").strip())
	try:
		import school_calendar as scal
		coords = scal.TR_PROVINCE_COORDS
	except Exception:
		return clean or fallback
	if not clean:
		return fallback if fallback in coords else "Ankara"
	if clean in coords:
		return clean
	f_clean = _tr_fold(clean)
	for k in coords:
		if _tr_fold(k) == f_clean:
			return k
	return fallback if fallback in coords else "Ankara"


def get_ramazan_times_for_dates(
	connect_db_fn: Callable[[], sqlite3.Connection] | None,
	province: str,
	date_isos: list[str],
) -> dict[str, dict[str, str]]:
	"""Seçili il ve tarih listesi için Sahur (İmsak) ve İftar (Akşam) saatlerini döndürür."""
	prov = resolve_province_name(province)
	out: dict[str, dict[str, str]] = {}
	if not date_isos:
		return out
	try:
		import school_calendar as scal
	except Exception:
		return out

	use_online_probe = True
	for dt_s in date_isos:
		d_iso = str(dt_s)[:10]
		try:
			d_obj = date.fromisoformat(d_iso)
		except Exception:
			continue
		if use_online_probe:
			pt = scal.get_prayer_times_for_province(prov, d_obj, connect_db_fn)
			if "Yerleşik" in str(pt.get("source") or ""):
				use_online_probe = False
		else:
			pt = scal._compute_offline_prayer_times(prov, d_obj)
		out[d_iso] = {
			"sahur": str(pt.get("İmsak") or ""),
			"iftar": str(pt.get("Akşam") or ""),
			"source": str(pt.get("source") or ""),
			"hijri": scal.gregorian_to_hijri_str(d_obj) if hasattr(scal, "gregorian_to_hijri_str") else "",
		}
	return out


def parse_service_schedule_from_cfg(app_cfg: dict[str, str] | None = None) -> dict[str, Any]:
	"""Yemekhane ayarlarından haftalık hizmet günü, günlük öğün sayısı, aktif öğünler, saatler ve Ramazan modunu çözümler."""
	cfg = app_cfg or {}
	ramazan_mode = str(cfg.get("cafeteria_ramazan_mode") or "0").strip() == "1"
	raw_prov = str(cfg.get("cafeteria_ramazan_province") or "").strip() or str(cfg.get("default_province") or "Ankara").strip()
	ramazan_province = resolve_province_name(raw_prov)

	raw_wd = str(cfg.get("cafeteria_service_weekdays") or "0,1,2,3,4").strip()
	wds: set[int] = set()
	for tok in raw_wd.split(","):
		t = tok.strip()
		if t.isdigit() and 0 <= int(t) <= 6:
			wds.add(int(t))
	if not wds:
		wds = {0, 1, 2, 3, 4}

	try:
		days_pw = max(1, min(7, int(float(str(cfg.get("cafeteria_service_days_per_week") or len(wds))))))
	except Exception:
		days_pw = len(wds)

	valid_pool = ALL_MEAL_SLOTS if ramazan_mode else MEAL_SLOTS
	raw_slots = [s.strip() for s in str(cfg.get("cafeteria_active_meal_slots") or "").split(",") if s.strip() in valid_pool]
	try:
		meals_pd = max(1, min(5, int(float(str(cfg.get("cafeteria_meals_per_day") or (len(raw_slots) or 1))))))
	except Exception:
		meals_pd = len(raw_slots) or 1

	if not raw_slots:
		if meals_pd == 1:
			raw_slots = ["Öğle Yemeği"]
		elif meals_pd == 2:
			raw_slots = ["Sabah Kahvaltısı", "Öğle Yemeği"]
		else:
			raw_slots = list(MEAL_SLOTS)

	meal_times: dict[str, str] = {}
	for sl, (cfg_k, def_v) in _MEAL_TIME_SETTING_KEYS.items():
		meal_times[sl] = str(cfg.get(cfg_k) or def_v).strip() or def_v
	if ramazan_mode:
		meal_times["İftar"] = f"{ramazan_province} İftar Vakti"
		meal_times["Sahur"] = f"{ramazan_province} İmsak Vakti"

	return {
		"days_per_week": days_pw,
		"weekdays_set": wds,
		"weekday_names": [TR_WEEKDAYS_SHORT[i] for i in sorted(wds)],
		"meals_per_day": meals_pd,
		"active_meal_slots": raw_slots,
		"available_slots": valid_pool,
		"meal_times": meal_times,
		"ramazan_mode": ramazan_mode,
		"ramazan_province": ramazan_province,
	}


# ==============================================================================
# 2. VERİTABANI ŞEMASI (8 TABLO — İDEMPOTENT)
# ==============================================================================

def ensure_cafeteria_tables(connection: sqlite3.Connection) -> None:
	"""Yemekhane yemekler, malzemeler, menü planları, alım faturaları ve kasa tablolarını oluşturur."""
	connection.executescript(
		"""
		CREATE TABLE IF NOT EXISTS cafeteria_dishes (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			name TEXT UNIQUE NOT NULL,
			category TEXT NOT NULL DEFAULT 'Ana Yemek',
			portion_g REAL NOT NULL DEFAULT 250.0,
			total_kcal REAL NOT NULL DEFAULT 0.0,
			carbs_g REAL NOT NULL DEFAULT 0.0,
			protein_g REAL NOT NULL DEFAULT 0.0,
			fat_g REAL NOT NULL DEFAULT 0.0,
			allergens_json TEXT NOT NULL DEFAULT '[]',
			recipe_notes TEXT DEFAULT '',
			source TEXT NOT NULL DEFAULT 'manual',
			created_at TEXT NOT NULL,
			updated_at TEXT NOT NULL
		);
		CREATE TABLE IF NOT EXISTS cafeteria_dish_ingredients (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			dish_id INTEGER NOT NULL REFERENCES cafeteria_dishes(id) ON DELETE CASCADE,
			ingredient_name TEXT NOT NULL,
			amount REAL NOT NULL DEFAULT 0.0,
			unit TEXT NOT NULL DEFAULT 'g',
			kcal_portion REAL NOT NULL DEFAULT 0.0,
			carbs_portion_g REAL NOT NULL DEFAULT 0.0,
			allergen_tag TEXT DEFAULT '',
			sort_order INTEGER NOT NULL DEFAULT 0
		);
		CREATE INDEX IF NOT EXISTS idx_caf_ing_dish ON cafeteria_dish_ingredients(dish_id);

		CREATE TABLE IF NOT EXISTS cafeteria_menu_plans (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			year_month TEXT NOT NULL,
			plan_date TEXT NOT NULL,
			meal_slot TEXT NOT NULL DEFAULT 'Öğle Yemeği',
			dish_ids_json TEXT NOT NULL DEFAULT '[]',
			custom_items_json TEXT NOT NULL DEFAULT '[]',
			total_kcal REAL NOT NULL DEFAULT 0.0,
			total_carbs_g REAL NOT NULL DEFAULT 0.0,
			allergens_json TEXT NOT NULL DEFAULT '[]',
			notes TEXT DEFAULT '',
			synced_to_diabetes INTEGER NOT NULL DEFAULT 0,
			updated_at TEXT NOT NULL,
			UNIQUE(plan_date, meal_slot)
		);
		CREATE INDEX IF NOT EXISTS idx_caf_menu_ym_date ON cafeteria_menu_plans(year_month, plan_date);

		CREATE TABLE IF NOT EXISTS cafeteria_purchases (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			invoice_date TEXT NOT NULL,
			year_month TEXT NOT NULL,
			invoice_no TEXT DEFAULT '',
			supplier_name TEXT DEFAULT '',
			category TEXT NOT NULL DEFAULT 'Diğer',
			subtotal_amount REAL NOT NULL DEFAULT 0.0,
			vat_amount REAL NOT NULL DEFAULT 0.0,
			total_amount REAL NOT NULL DEFAULT 0.0,
			payment_status TEXT NOT NULL DEFAULT 'Ödendi',
			document_name TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);
		CREATE INDEX IF NOT EXISTS idx_caf_purch_ym_date ON cafeteria_purchases(year_month, invoice_date);

		CREATE TABLE IF NOT EXISTS cafeteria_purchase_items (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			purchase_id INTEGER NOT NULL REFERENCES cafeteria_purchases(id) ON DELETE CASCADE,
			item_name TEXT NOT NULL,
			category TEXT NOT NULL DEFAULT 'Diğer',
			quantity REAL NOT NULL DEFAULT 1.0,
			unit_type TEXT NOT NULL DEFAULT 'kg',
			unit_price REAL NOT NULL DEFAULT 0.0,
			vat_rate REAL NOT NULL DEFAULT 0.0,
			line_total REAL NOT NULL DEFAULT 0.0,
			expiry_date TEXT DEFAULT '',
			notes TEXT DEFAULT ''
		);
		CREATE INDEX IF NOT EXISTS idx_caf_pitem_purch ON cafeteria_purchase_items(purchase_id);

		CREATE TABLE IF NOT EXISTS cafeteria_diners (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			student_id INTEGER REFERENCES students(id) ON DELETE SET NULL,
			full_name TEXT NOT NULL,
			role_type TEXT NOT NULL DEFAULT 'Öğrenci',
			class_or_dept TEXT DEFAULT '',
			daily_meals_count INTEGER NOT NULL DEFAULT 1,
			fee_exempt INTEGER NOT NULL DEFAULT 0,
			allergies_override TEXT DEFAULT '',
			is_active INTEGER NOT NULL DEFAULT 1,
			start_date TEXT DEFAULT '',
			end_date TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			UNIQUE(student_id)
		);

		CREATE TABLE IF NOT EXISTS cafeteria_monthly_fees (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			year_month TEXT NOT NULL,
			diner_id INTEGER NOT NULL REFERENCES cafeteria_diners(id) ON DELETE CASCADE,
			workdays_in_month INTEGER NOT NULL DEFAULT 20,
			planned_days INTEGER NOT NULL DEFAULT 20,
			meals_per_day INTEGER NOT NULL DEFAULT 1,
			unit_meal_price REAL NOT NULL DEFAULT 0.0,
			discount_or_adjustment REAL NOT NULL DEFAULT 0.0,
			expected_amount REAL NOT NULL DEFAULT 0.0,
			note TEXT DEFAULT '',
			updated_at TEXT NOT NULL,
			UNIQUE(year_month, diner_id)
		);
		CREATE INDEX IF NOT EXISTS idx_caf_fees_ym ON cafeteria_monthly_fees(year_month);

		CREATE TABLE IF NOT EXISTS cafeteria_cash_ledger (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			tx_date TEXT NOT NULL,
			year_month TEXT NOT NULL,
			tx_type TEXT NOT NULL DEFAULT 'IN_PAYMENT',
			diner_id INTEGER REFERENCES cafeteria_diners(id) ON DELETE SET NULL,
			for_person_name TEXT DEFAULT '',
			delivered_by_name TEXT DEFAULT '',
			received_by_name TEXT DEFAULT '',
			amount REAL NOT NULL DEFAULT 0.0,
			payment_method TEXT NOT NULL DEFAULT 'Nakit',
			receipt_no TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);
		CREATE INDEX IF NOT EXISTS idx_caf_cash_ym_date ON cafeteria_cash_ledger(year_month, tx_date);
		"""
	)
	diner_cols = {row["name"] for row in connection.execute("PRAGMA table_info(cafeteria_diners)")}
	if "end_date" not in diner_cols:
		connection.execute("ALTER TABLE cafeteria_diners ADD COLUMN end_date TEXT DEFAULT ''")
	fee_cols = {row["name"] for row in connection.execute("PRAGMA table_info(cafeteria_monthly_fees)")}
	if "carry_forward_debt" not in fee_cols:
		connection.execute("ALTER TABLE cafeteria_monthly_fees ADD COLUMN carry_forward_debt REAL NOT NULL DEFAULT 0.0")


# ==============================================================================
# 3. BÖLÜM 1: YEMEKLER BANKASI, MALZEMELER, KALORİ & ALERJEN MOTORU
# ==============================================================================

def compute_dish_nutrition_from_ingredients(ingredients: list[dict[str, Any]]) -> tuple[float, float, list[str]]:
	"""Malzeme satırlarından porsiyon toplam kalorisini, karbonhidratını ve benzersiz alerjenleri hesaplar."""
	tot_kcal = 0.0
	tot_carb = 0.0
	allergens: list[str] = []
	seen_al: set[str] = set()
	for ing in ingredients or []:
		if not isinstance(ing, dict):
			continue
		name = str(ing.get("ingredient_name") or ing.get("name") or "").strip()
		if not name:
			continue
		tot_kcal += max(0.0, _safe_float(ing.get("kcal_portion"), 0.0))
		tot_carb += max(0.0, _safe_float(ing.get("carbs_portion_g"), 0.0))
		al = _normalize_allergen_label(ing.get("allergen_tag") or "")
		if al and al not in seen_al:
			seen_al.add(al)
			allergens.append(al)
	return round(tot_kcal, 1), round(tot_carb, 1), allergens


def scale_edited_ingredients(
	baseline_rows: list[dict[str, Any]],
	edited_df: pd.DataFrame,
	*,
	auto_scale_nutrition: bool = True,
) -> list[dict[str, Any]]:
	"""
	Düzenlenen malzeme tablosunda kullanıcının yalnızca miktarı (amount) değiştirmesi hâlinde
	baz kalori ve karbonhidrat değerlerini oranlayarak günceller; doğrudan kalori yazılmışsa korur.
	"""
	base_by_idx = {i: r for i, r in enumerate(baseline_rows or [])}
	base_by_name = {_tr_fold(r.get("ingredient_name")): r for r in (baseline_rows or []) if r.get("ingredient_name")}
	out: list[dict[str, Any]] = []
	if edited_df is None or edited_df.empty:
		return out

	for r_idx, row in edited_df.iterrows():
		ing_name = str(row.get("Malzeme") or row.get("ingredient_name") or "").strip()
		if not ing_name or ing_name.lower() in {"nan", "none"}:
			continue
		amt = max(0.0, _safe_float(row.get("Miktar", row.get("amount")), 0.0))
		unit = str(row.get("Birim") or row.get("unit") or "g").strip() or "g"
		kcal = max(0.0, _safe_float(row.get("Kalori (kcal)", row.get("kcal_portion")), 0.0))
		carb = max(0.0, _safe_float(row.get("Karb (g)", row.get("carbs_portion_g")), 0.0))
		al_raw = row.get("Alerjen") if "Alerjen" in row else row.get("allergen_tag")
		al_tag = _normalize_allergen_label(al_raw)

		if auto_scale_nutrition:
			orig = base_by_name.get(_tr_fold(ing_name)) or base_by_idx.get(r_idx)
			if isinstance(orig, dict):
				old_amt = _safe_float(orig.get("amount"), 0.0)
				old_kcal = _safe_float(orig.get("kcal_portion"), 0.0)
				old_carb = _safe_float(orig.get("carbs_portion_g"), 0.0)
				old_unit = str(orig.get("unit") or "g").strip()
				if old_amt > 0 and amt > 0 and unit == old_unit and abs(amt - old_amt) > 1e-6:
					# Kullanıcı kaloriyi ayrıca elle değiştirmediyse oranla
					if abs(kcal - old_kcal) < 0.05 and old_kcal > 0:
						kcal = round(old_kcal * (amt / old_amt), 1)
					if abs(carb - old_carb) < 0.05 and old_carb > 0:
						carb = round(old_carb * (amt / old_amt), 1)

		out.append({
			"ingredient_name": ing_name,
			"amount": round(amt, 2),
			"unit": unit,
			"kcal_portion": round(kcal, 1),
			"carbs_portion_g": round(carb, 1),
			"allergen_tag": al_tag,
		})
	return out


def upsert_dish_with_ingredients(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	dish_id: int | None = None,
	name: str,
	category: str = "Ana Yemek",
	portion_g: float = 250.0,
	ingredients: list[dict[str, Any]] | None = None,
	protein_g: float = 0.0,
	fat_g: float = 0.0,
	manual_kcal: float | None = None,
	manual_carbs_g: float | None = None,
	extra_allergens: list[str] | None = None,
	recipe_notes: str = "",
	source: str = "manual",
) -> int:
	"""Yemeği ve malzeme kalemlerini kaydeder; toplam kalori ve alerjen listesini hesaplar."""
	clean_name = re.sub(r"\s+", " ", str(name or "").strip())
	if not clean_name:
		raise ValueError("Yemek adı boş bırakılamaz.")
	ing_list = ingredients or []
	calc_kcal, calc_carb, calc_allergens = compute_dish_nutrition_from_ingredients(ing_list)
	final_kcal = round(float(manual_kcal), 1) if (manual_kcal is not None and not ing_list) else (calc_kcal if ing_list else round(float(manual_kcal or 0.0), 1))
	final_carb = round(float(manual_carbs_g), 1) if (manual_carbs_g is not None and not ing_list) else (calc_carb if ing_list else round(float(manual_carbs_g or 0.0), 1))

	all_allergens = list(calc_allergens)
	for ea in (extra_allergens or []):
		norm_ea = _normalize_allergen_label(ea)
		if norm_ea and norm_ea not in all_allergens:
			all_allergens.append(norm_ea)

	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		existing = None
		if dish_id is not None:
			existing = conn.execute("SELECT id FROM cafeteria_dishes WHERE id=?", (int(dish_id),)).fetchone()
		if existing is None:
			existing = conn.execute(
				"SELECT id FROM cafeteria_dishes WHERE lower(trim(name))=lower(trim(?))",
				(clean_name,),
			).fetchone()

		if existing:
			target_id = int(existing["id"])
			conn.execute(
				"""UPDATE cafeteria_dishes
				   SET name=?, category=?, portion_g=?, total_kcal=?, carbs_g=?, protein_g=?, fat_g=?,
				       allergens_json=?, recipe_notes=?, source=?, updated_at=?
				   WHERE id=?""",
				(
					clean_name,
					str(category or "Ana Yemek"),
					max(0.0, float(portion_g or 0.0)),
					final_kcal,
					final_carb,
					max(0.0, float(protein_g or 0.0)),
					max(0.0, float(fat_g or 0.0)),
					json.dumps(all_allergens, ensure_ascii=False),
					str(recipe_notes or "").strip(),
					str(source or "manual"),
					now_iso,
					target_id,
				),
			)
			conn.execute("DELETE FROM cafeteria_dish_ingredients WHERE dish_id=?", (target_id,))
		else:
			cur = conn.execute(
				"""INSERT INTO cafeteria_dishes(
					name, category, portion_g, total_kcal, carbs_g, protein_g, fat_g,
					allergens_json, recipe_notes, source, created_at, updated_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
				(
					clean_name,
					str(category or "Ana Yemek"),
					max(0.0, float(portion_g or 0.0)),
					final_kcal,
					final_carb,
					max(0.0, float(protein_g or 0.0)),
					max(0.0, float(fat_g or 0.0)),
					json.dumps(all_allergens, ensure_ascii=False),
					str(recipe_notes or "").strip(),
					str(source or "manual"),
					now_iso,
					now_iso,
				),
			)
			target_id = int(cur.lastrowid)

		for idx, ing in enumerate(ing_list):
			if not isinstance(ing, dict):
				continue
			iname = str(ing.get("ingredient_name") or ing.get("name") or "").strip()
			if not iname:
				continue
			conn.execute(
				"""INSERT INTO cafeteria_dish_ingredients(
					dish_id, ingredient_name, amount, unit, kcal_portion, carbs_portion_g, allergen_tag, sort_order
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
				(
					target_id,
					iname,
					max(0.0, _safe_float(ing.get("amount"), 0.0)),
					str(ing.get("unit") or "g").strip() or "g",
					max(0.0, _safe_float(ing.get("kcal_portion"), 0.0)),
					max(0.0, _safe_float(ing.get("carbs_portion_g"), 0.0)),
					_normalize_allergen_label(ing.get("allergen_tag") or ""),
					idx,
				),
			)
	recalculate_menu_plans_for_dish(connect_db_fn, target_id)
	return target_id


def parse_custom_dish_entries(raw_text: str) -> list[dict[str, Any]]:
	"""
	Serbest metin yemek satırlarını ayrıştırır (köşeli parantez içindeki virgülleri korur).
	Örnek: 'Tam Buğday Ekmek [130 kcal, 24g karb, gluten], Mevsim Salata'
	"""
	out: list[dict[str, Any]] = []
	for part in re.split(r",(?![^\[]*\])", str(raw_text or "")):
		chunk = part.strip()
		if not chunk:
			continue
		kcal_v, carb_v, al_v = 0.0, 0.0, ""
		m = re.search(r"\[([^\]]+)\]\s*$", chunk)
		name_clean = chunk
		if m:
			name_clean = chunk[:m.start()].strip()
			meta_s = m.group(1)
			mk = re.search(r"(\d+(?:[.,]\d+)?)\s*kcal", meta_s, flags=re.IGNORECASE)
			mc = re.search(r"(\d+(?:[.,]\d+)?)\s*g?\s*karb", meta_s, flags=re.IGNORECASE)
			if mk:
				kcal_v = round(_safe_float(mk.group(1), 0.0), 1)
			if mc:
				carb_v = round(_safe_float(mc.group(1), 0.0), 1)
			rem = re.sub(r"\d+(?:[.,]\d+)?\s*(?:kcal|g?\s*karb\w*)", "", meta_s, flags=re.IGNORECASE)
			for tok in re.split(r"[;|/]+", rem):
				tok_n = _normalize_allergen_label(tok.strip(" ,."))
				if tok_n in ALLERGEN_BADGE_MAP:
					al_v = tok_n
					break
		if name_clean:
			out.append({"name": name_clean, "kcal": kcal_v, "carbs_g": carb_v, "allergen": al_v})
	return out


def upsert_quick_dish_from_plan_item(
	connect_db_fn: Callable[[], sqlite3.Connection],
	item: dict[str, Any],
	*,
	category: str = "Ana Yemek",
) -> int | None:
	"""Menü planında veya AI çıktısında geçen yeni yemeği bankada yoksa temel besin değerleriyle ekler."""
	nm = re.sub(r"\s+", " ", str(item.get("name") or "").strip())
	if not nm:
		return None
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		ex = conn.execute("SELECT id FROM cafeteria_dishes WHERE lower(trim(name))=lower(trim(?))", (nm,)).fetchone()
		if ex:
			return int(ex["id"])
	kcal = max(0.0, _safe_float(item.get("kcal"), 0.0))
	carb = max(0.0, _safe_float(item.get("carbs_g"), 0.0))
	al_raw = str(item.get("allergen") or "")
	extras = [_normalize_allergen_label(a) for a in re.split(r"[,;/]+", al_raw) if _normalize_allergen_label(a)]
	return upsert_dish_with_ingredients(
		connect_db_fn,
		name=nm,
		category=category,
		portion_g=250.0,
		ingredients=[],
		manual_kcal=kcal,
		manual_carbs_g=carb,
		extra_allergens=extras,
		recipe_notes="Menü planından otomatik eklendi",
		source="ai",
	)


def get_all_dishes(connect_db_fn: Callable[[], sqlite3.Connection]) -> list[dict[str, Any]]:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute("SELECT * FROM cafeteria_dishes ORDER BY category ASC, name ASC").fetchall()
		ing_rows = conn.execute("SELECT * FROM cafeteria_dish_ingredients ORDER BY dish_id ASC, sort_order ASC, id ASC").fetchall()

	ings_by_dish: dict[int, list[dict[str, Any]]] = {}
	for ir in ing_rows:
		d_ir = dict(ir)
		ings_by_dish.setdefault(int(d_ir["dish_id"]), []).append(d_ir)

	out: list[dict[str, Any]] = []
	for r in rows:
		d = dict(r)
		try:
			d["allergens"] = json.loads(d.get("allergens_json") or "[]")
		except Exception:
			d["allergens"] = []
		d["ingredients"] = ings_by_dish.get(int(d["id"]), [])
		out.append(d)
	return out


def delete_dish_by_id(connect_db_fn: Callable[[], sqlite3.Connection], dish_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute("DELETE FROM cafeteria_dishes WHERE id=?", (int(dish_id),))


def match_allergens_against_school(
	connect_db_fn: Callable[[], sqlite3.Connection],
	allergens: list[str],
) -> list[dict[str, str]]:
	"""Verilen alerjen listesini revir alerji/çölyak/kronik/özel diyet profilleri ve yemekhane aboneleriyle çapraz eşler."""
	norm_als = [_normalize_allergen_label(a) for a in (allergens or []) if _normalize_allergen_label(a)]
	if not norm_als:
		return []

	hits: list[dict[str, str]] = []
	seen_keys: set[tuple[str, str]] = set()
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		try:
			stu_rows = conn.execute(
				"""SELECT s.id, s.full_name, COALESCE(c.name, '—') AS class_name,
				          TRIM(
				              COALESCE(hp.allergies, '') || ' ' ||
				              COALESCE(hp.chronic_conditions, '') || ' ' ||
				              COALESCE(hp.dietary_swallowing_notes, '')
				          ) AS combined_risk,
				          COALESCE(hp.allergies, '') AS allergies,
				          COALESCE(hp.chronic_conditions, '') AS chronic_conditions,
				          COALESCE(hp.dietary_swallowing_notes, '') AS dietary_swallowing_notes
				   FROM students s
				   JOIN health_profiles hp ON hp.student_id = s.id
				   LEFT JOIN classes c ON c.id = s.class_id
				   WHERE TRIM(COALESCE(hp.allergies, '') || COALESCE(hp.chronic_conditions, '') || COALESCE(hp.dietary_swallowing_notes, '')) != ''"""
			).fetchall()
		except Exception:
			stu_rows = []
		try:
			diner_rows = conn.execute(
				"""SELECT id, full_name, COALESCE(class_or_dept, role_type) AS class_name,
				          allergies_override AS combined_risk
				   FROM cafeteria_diners
				   WHERE is_active=1 AND TRIM(COALESCE(allergies_override, '')) != ''"""
			).fetchall()
		except Exception:
			diner_rows = []

	def _scan(rows_in: list[Any], src_label: str) -> None:
		for r in rows_in:
			al_txt = _tr_fold(r["combined_risk"])
			if not al_txt or al_txt in {"yok", "bilinmiyor", "-"}:
				continue
			matched_als: list[str] = []
			for al in norm_als:
				syns = _ALLERGEN_SYNONYMS.get(al, (_tr_fold(al),))
				if any(_tr_fold(sy) in al_txt for sy in syns):
					matched_als.append(al)
			if matched_als:
				key = (str(r["full_name"]), ", ".join(matched_als))
				if key not in seen_keys:
					seen_keys.add(key)
					raw_detail = str(r["combined_risk"]).strip()
					hits.append({
						"full_name": str(r["full_name"]),
						"class_name": str(r["class_name"] or "—"),
						"allergies": raw_detail,
						"matched": ", ".join(matched_als),
						"source_type": src_label,
					})

	_scan(stu_rows, "Revir Profili")
	_scan(diner_rows, "Yemekhane Abone")
	return hits


def get_diabetic_students_for_menu(connect_db_fn: Callable[[], sqlite3.Connection]) -> list[dict[str, Any]]:
	"""Okuldaki aktif Tip 1 / takipli diyabet öğrencilerini döndürür."""
	with connect_db_fn() as conn:
		try:
			rows = conn.execute(
				"""SELECT s.id, s.full_name, COALESCE(c.name, '—') AS class_name,
				          dp.diabetes_type, dp.icr_lunch
				   FROM students s
				   JOIN diabetes_profiles dp ON dp.student_id = s.id
				   LEFT JOIN classes c ON c.id = s.class_id
				   WHERE COALESCE(dp.is_diabetic, 0) = 1
				   ORDER BY c.name ASC, s.full_name ASC"""
			).fetchall()
			return [dict(r) for r in rows]
		except Exception:
			return []


def get_clinical_nutrition_risk_summary(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	mask_fn: Callable[[Any, str], str] | None = None,
) -> dict[str, Any]:
	"""AI güvenlik denetimi ve planlayıcı için KVKK-uyumlu klinik beslenme risk özeti çıkarır."""
	allergic_hits = match_allergens_against_school(connect_db_fn, [a for a in TR_ALLERGEN_OPTIONS if a])
	diabetic_list = get_diabetic_students_for_menu(connect_db_fn)
	swallow_notes: list[dict[str, str]] = []
	with connect_db_fn() as conn:
		try:
			rows = conn.execute(
				"""SELECT s.full_name, COALESCE(c.name, '—') AS class_name, hp.dietary_swallowing_notes
				   FROM students s
				   JOIN health_profiles hp ON hp.student_id = s.id
				   LEFT JOIN classes c ON c.id = s.class_id
				   WHERE TRIM(COALESCE(hp.dietary_swallowing_notes, '')) != ''"""
			).fetchall()
			for r in rows:
				nm = mask_fn(r["full_name"], "name") if callable(mask_fn) else f"{str(r['full_name']).split()[0][0]}***"
				swallow_notes.append({"person": nm, "class": str(r["class_name"]), "note": str(r["dietary_swallowing_notes"])})
		except Exception:
			pass
	return {
		"allergic_or_celiac_count": len(allergic_hits),
		"allergic_or_celiac_groups": sorted({m.strip() for h in allergic_hits for m in h["matched"].split(",") if m.strip()}),
		"allergic_sample_classes": [f"{h['class_name']} ({h['matched']})" for h in allergic_hits[:20]],
		"diabetic_student_count": len(diabetic_list),
		"diabetic_classes": [d["class_name"] for d in diabetic_list[:20]],
		"special_diet_notes": swallow_notes[:15],
	}


# ==============================================================================
# 4. BÖLÜM 2: İŞ GÜNLERİ, ÖĞÜN BAZLI MENÜ PLANLAYICI & A4 ÇIKTI MOTORU
# ==============================================================================

def get_month_workdays_info(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
	*,
	include_weekends: bool = False,
	treat_half_day_as_workday: bool = True,
	service_weekdays: set[int] | None = None,
) -> list[dict[str, Any]]:
	"""Seçili ayın günlerini okul takvimindeki tatillere ve yemekhane haftalık çalışma günlerine göre sınıflandırır."""
	ym_clean = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy, mm = [int(x) for x in ym_clean.split("-")]
	try:
		import school_calendar as scal
		scal.ensure_year_builtin_holidays_seeded(yy, connect_db_fn)
	except Exception:
		pass

	events_by_date: dict[str, list[dict[str, Any]]] = {}
	cfg_map: dict[str, str] = {}
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		try:
			rows = conn.execute(
				"SELECT event_date, title, category, is_civil_holiday, is_half_day FROM calendar_events WHERE event_date LIKE ?",
				(f"{ym_clean}-%",),
			).fetchall()
			for r in rows:
				d_k = str(r["event_date"])[:10]
				events_by_date.setdefault(d_k, []).append(dict(r))
		except Exception:
			pass
		try:
			for s_r in conn.execute("SELECT key, value FROM app_settings WHERE key LIKE 'cafeteria_%'").fetchall():
				cfg_map[str(s_r["key"])] = str(s_r["value"])
		except Exception:
			pass

	active_wds = set(service_weekdays) if service_weekdays is not None else parse_service_schedule_from_cfg(cfg_map)["weekdays_set"]

	_, num_days = calendar.monthrange(yy, mm)
	out: list[dict[str, Any]] = []
	for d_num in range(1, num_days + 1):
		d_obj = date(yy, mm, d_num)
		d_iso = d_obj.isoformat()
		wd = d_obj.weekday()
		is_wknd = wd >= 5
		evs = events_by_date.get(d_iso, [])
		full_hol_titles = [
			str(e.get("title") or "Resmi Tatil")
			for e in evs
			if (bool(e.get("is_civil_holiday")) and not bool(e.get("is_half_day")))
			or str(e.get("category") or "") == "secim_idari_tatil"
		]
		half_titles = [str(e.get("title") or "Yarım Gün") for e in evs if bool(e.get("is_half_day"))]
		is_holiday = bool(full_hol_titles)
		is_half = bool(half_titles)
		is_service_day = include_weekends or (wd in active_wds)
		is_work = is_service_day and (not is_holiday) and (treat_half_day_as_workday or not is_half)
		out.append({
			"date_iso": d_iso,
			"day_num": d_num,
			"weekday": wd,
			"weekday_name": TR_WEEKDAYS_SHORT.get(wd, ""),
			"iso_week": d_obj.isocalendar()[1],
			"is_weekend": is_wknd,
			"is_holiday": is_holiday,
			"is_half_day": is_half,
			"holiday_title": ", ".join(full_hol_titles or half_titles),
			"is_workday": is_work,
		})
	return out


_DIABETES_SLOT_MAP: dict[str, str] = {
	"Sabah Kahvaltısı": "Kahvaltı",
	"Öğle Yemeği": "Öğle Yemeği",
	"İkindi Ara Öğünü": "İkindi / Ara Öğün",
	"İftar": "Akşam Yemeği",
	"Sahur": "Sahur / Gece",
}


def _delete_synced_diabetes_slot(
	connect_db_fn: Callable[[], sqlite3.Connection],
	plan_date: str,
	meal_slot: str | None = None,
) -> None:
	"""Yemekhaneden senkronize edilmiş diyabet okul menüsü kaydını güvenle siler."""
	p_dt = str(plan_date).strip()[:10]
	with connect_db_fn() as conn:
		try:
			idm.ensure_diabetes_tables(conn)
			if meal_slot:
				m_type = _DIABETES_SLOT_MAP.get(meal_slot, meal_slot)
				conn.execute(
					"DELETE FROM school_meal_menus WHERE menu_date=? AND meal_type=? AND source_note LIKE 'Yemekhane%'",
					(p_dt, m_type),
				)
			else:
				conn.execute(
					"DELETE FROM school_meal_menus WHERE menu_date=? AND source_note LIKE 'Yemekhane%'",
					(p_dt,),
				)
		except Exception:
			pass


def _compute_menu_slot_totals(
	conn: sqlite3.Connection,
	d_ids: list[int],
	cust: list[dict[str, Any]],
) -> tuple[float, float, list[str], list[dict[str, Any]]]:
	"""Yemek ID'leri ve serbest öğelerden toplam kalori, karbonhidrat, alerjen ve diyabet öğelerini hesaplar."""
	tot_kcal = 0.0
	tot_carb = 0.0
	allergens: list[str] = []
	seen_al: set[str] = set()
	diabetes_items: list[dict[str, Any]] = []

	if d_ids:
		ph = ",".join("?" for _ in d_ids)
		d_rows = conn.execute(f"SELECT * FROM cafeteria_dishes WHERE id IN ({ph})", d_ids).fetchall()
		by_id = {int(r["id"]): dict(r) for r in d_rows}
		for did in d_ids:
			dr = by_id.get(did)
			if not dr:
				continue
			tot_kcal += _safe_float(dr.get("total_kcal"), 0.0)
			tot_carb += _safe_float(dr.get("carbs_g"), 0.0)
			diabetes_items.append({
				"name": str(dr["name"]),
				"portion": f"{_safe_float(dr.get('portion_g'), 250.0):g}g ({_safe_float(dr.get('total_kcal'), 0.0):g} kcal)",
				"carbs_g": round(_safe_float(dr.get("carbs_g"), 0.0), 1),
			})
			try:
				for a in json.loads(dr.get("allergens_json") or "[]"):
					na = _normalize_allergen_label(a)
					if na and na not in seen_al:
						seen_al.add(na)
						allergens.append(na)
			except Exception:
				pass

	for ci in cust:
		ck = max(0.0, _safe_float(ci.get("kcal"), 0.0))
		cc = max(0.0, _safe_float(ci.get("carbs_g"), 0.0))
		tot_kcal += ck
		tot_carb += cc
		diabetes_items.append({
			"name": str(ci.get("name") or "").strip(),
			"portion": str(ci.get("portion") or "1 porsiyon"),
			"carbs_g": round(cc, 1),
		})
		na = _normalize_allergen_label(ci.get("allergen"))
		if na and na not in seen_al:
			seen_al.add(na)
			allergens.append(na)

	return round(tot_kcal, 1), round(tot_carb, 1), allergens, diabetes_items


def recalculate_menu_plans_for_dish(
	connect_db_fn: Callable[[], sqlite3.Connection],
	dish_id: int,
) -> int:
	"""Bir yemek güncellendiğinde o yemeği içeren tüm planlı menülerin kalori/karb/alerjen özetini yeniler."""
	to_resync: list[tuple[str, str, list[dict[str, Any]]]] = []
	updated = 0
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute("SELECT * FROM cafeteria_menu_plans").fetchall()
		now_iso = datetime.now().isoformat(timespec="minutes")
		for r in rows:
			try:
				d_ids = [int(x) for x in json.loads(r["dish_ids_json"] or "[]")]
			except Exception:
				d_ids = []
			if int(dish_id) not in d_ids:
				continue
			try:
				cust = json.loads(r["custom_items_json"] or "[]")
			except Exception:
				cust = []
			tot_k, tot_c, als, diab_items = _compute_menu_slot_totals(conn, d_ids, cust)
			conn.execute(
				"""UPDATE cafeteria_menu_plans
				   SET total_kcal=?, total_carbs_g=?, allergens_json=?, updated_at=?
				   WHERE id=?""",
				(tot_k, tot_c, json.dumps(als, ensure_ascii=False), now_iso, int(r["id"])),
			)
			updated += 1
			if r["synced_to_diabetes"] and diab_items:
				to_resync.append((str(r["plan_date"])[:10], str(r["meal_slot"]), diab_items))

	for p_date, slot, diab_items in to_resync:
		idm.upsert_school_meal_menu(
			connect_db_fn,
			p_date,
			_DIABETES_SLOT_MAP.get(slot, slot),
			diab_items,
			source_note="Yemekhane Planı Senkronizasyonu",
		)
	return updated


def upsert_menu_plan_slot(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	plan_date: str,
	meal_slot: str = "Öğle Yemeği",
	dish_ids: list[int] | None = None,
	custom_items: list[dict[str, Any]] | None = None,
	notes: str = "",
	sync_to_diabetes: bool = False,
) -> dict[str, Any]:
	"""Belirli tarih ve öğün slotuna yemek listesini kaydeder; boşsa slottaki kaydı ve bağlı diyabet kaydını temizler."""
	p_date = str(plan_date).strip()[:10]
	ym = p_date[:7]
	slot = str(meal_slot or "Öğle Yemeği").strip() or "Öğle Yemeği"
	d_ids = [int(x) for x in (dish_ids or []) if str(x).isdigit()]
	cust = [c for c in (custom_items or []) if isinstance(c, dict) and str(c.get("name") or "").strip()]

	if not d_ids and not cust:
		delete_menu_plan_slot(connect_db_fn, p_date, slot)
		return {"plan_date": p_date, "meal_slot": slot, "total_kcal": 0.0, "total_carbs_g": 0.0, "allergens": [], "items": []}

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		tot_kcal, tot_carb, allergens, diabetes_items = _compute_menu_slot_totals(conn, d_ids, cust)
		now_iso = datetime.now().isoformat(timespec="minutes")
		conn.execute(
			"""INSERT INTO cafeteria_menu_plans(
				year_month, plan_date, meal_slot, dish_ids_json, custom_items_json,
				total_kcal, total_carbs_g, allergens_json, notes, synced_to_diabetes, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(plan_date, meal_slot) DO UPDATE SET
				year_month=excluded.year_month,
				dish_ids_json=excluded.dish_ids_json,
				custom_items_json=excluded.custom_items_json,
				total_kcal=excluded.total_kcal,
				total_carbs_g=excluded.total_carbs_g,
				allergens_json=excluded.allergens_json,
				notes=excluded.notes,
				synced_to_diabetes=excluded.synced_to_diabetes,
				updated_at=excluded.updated_at""",
			(
				ym, p_date, slot,
				json.dumps(d_ids),
				json.dumps(cust, ensure_ascii=False),
				tot_kcal, tot_carb,
				json.dumps(allergens, ensure_ascii=False),
				str(notes or "").strip(),
				1 if sync_to_diabetes else 0,
				now_iso,
			),
		)

	if sync_to_diabetes and diabetes_items:
		idm.upsert_school_meal_menu(
			connect_db_fn,
			p_date,
			_DIABETES_SLOT_MAP.get(slot, slot),
			diabetes_items,
			source_note="Yemekhane Planı Senkronizasyonu",
		)

	return {
		"plan_date": p_date,
		"meal_slot": slot,
		"total_kcal": tot_kcal,
		"total_carbs_g": tot_carb,
		"allergens": allergens,
		"items": diabetes_items,
	}


def get_menu_plans_for_month(connect_db_fn: Callable[[], sqlite3.Connection], year_month: str) -> dict[tuple[str, str], dict[str, Any]]:
	"""Ayın planlanmış tüm öğünlerini {(plan_date, meal_slot): kayıt} sözlüğü olarak döndürür."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute(
			"SELECT * FROM cafeteria_menu_plans WHERE year_month=? ORDER BY plan_date ASC, meal_slot ASC",
			(ym,),
		).fetchall()
		dishes_map = {
			int(dr["id"]): dict(dr)
			for dr in conn.execute("SELECT id, name, category, total_kcal, carbs_g, allergens_json FROM cafeteria_dishes").fetchall()
		}

	out: dict[tuple[str, str], dict[str, Any]] = {}
	for r in rows:
		d = dict(r)
		try:
			d_ids = json.loads(d.get("dish_ids_json") or "[]")
		except Exception:
			d_ids = []
		try:
			cust = json.loads(d.get("custom_items_json") or "[]")
		except Exception:
			cust = []
		try:
			als = json.loads(d.get("allergens_json") or "[]")
		except Exception:
			als = []
		dish_objs = [dishes_map[int(x)] for x in d_ids if int(x) in dishes_map]
		dish_names = [str(x["name"]) for x in dish_objs] + [str(c.get("name")) for c in cust if isinstance(c, dict) and c.get("name")]
		d["dish_ids"] = d_ids
		d["custom_items"] = cust
		d["allergens"] = als
		d["dish_names"] = dish_names
		out[(str(d["plan_date"])[:10], str(d["meal_slot"]))] = d
	return out


def sync_cafeteria_month_to_diabetes(connect_db_fn: Callable[[], sqlite3.Connection], year_month: str) -> int:
	"""Seçili aydaki tüm yemekhane menülerini İnsülin & Diyabet Yönetimi `school_meal_menus` tablosuna aktarır."""
	plans = get_menu_plans_for_month(connect_db_fn, year_month)
	all_dishes = {d["id"]: d for d in get_all_dishes(connect_db_fn)}
	slot_map = {
		"Sabah Kahvaltısı": "Kahvaltı",
		"Öğle Yemeği": "Öğle Yemeği",
		"İkindi Ara Öğünü": "İkindi / Ara Öğün",
	}
	synced_cnt = 0
	for (p_date, slot), rec in plans.items():
		items_for_diab: list[dict[str, Any]] = []
		for did in rec.get("dish_ids") or []:
			dr = all_dishes.get(int(did))
			if dr:
				items_for_diab.append({
					"name": dr["name"],
					"portion": f"{_safe_float(dr.get('portion_g'), 250.0):g}g ({_safe_float(dr.get('total_kcal'), 0.0):g} kcal)",
					"carbs_g": round(_safe_float(dr.get("carbs_g"), 0.0), 1),
				})
		for ci in rec.get("custom_items") or []:
			if isinstance(ci, dict) and ci.get("name"):
				items_for_diab.append({
					"name": str(ci["name"]),
					"portion": str(ci.get("portion") or "1 porsiyon"),
					"carbs_g": round(_safe_float(ci.get("carbs_g"), 0.0), 1),
				})
		if items_for_diab:
			idm.upsert_school_meal_menu(
				connect_db_fn,
				p_date,
				slot_map.get(slot, slot),
				items_for_diab,
				source_note="Yemekhane Planı Senkronizasyonu",
			)
			with connect_db_fn() as conn:
				conn.execute(
					"UPDATE cafeteria_menu_plans SET synced_to_diabetes=1 WHERE plan_date=? AND meal_slot=?",
					(p_date, slot),
				)
			synced_cnt += 1
	return synced_cnt


def delete_menu_plan_slot(connect_db_fn: Callable[[], sqlite3.Connection], plan_date: str, meal_slot: str | None = None) -> int:
	p_dt = str(plan_date).strip()[:10]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		if meal_slot:
			cur = conn.execute("DELETE FROM cafeteria_menu_plans WHERE plan_date=? AND meal_slot=?", (p_dt, meal_slot))
		else:
			cur = conn.execute("DELETE FROM cafeteria_menu_plans WHERE plan_date=?", (p_dt,))
		deleted = int(cur.rowcount or 0)
	_delete_synced_diabetes_slot(connect_db_fn, p_dt, meal_slot)
	return deleted


def copy_menu_plans(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	date_pairs: list[tuple[str, str]],
	meal_slots: list[str] | None = None,
	overwrite: bool = True,
	sync_to_diabetes: bool = False,
) -> int:
	"""Verilen (kaynak_tarih, hedef_tarih) çiftleri arasında planlı öğünleri kopyalar."""
	if not date_pairs:
		return 0
	src_dates = sorted({str(s)[:10] for s, _ in date_pairs if s})
	if not src_dates:
		return 0
	ph = ",".join("?" for _ in src_dates)
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		src_rows = conn.execute(
			f"SELECT * FROM cafeteria_menu_plans WHERE plan_date IN ({ph})",
			src_dates,
		).fetchall()
		by_src: dict[str, list[dict[str, Any]]] = {}
		for r in src_rows:
			by_src.setdefault(str(r["plan_date"])[:10], []).append(dict(r))

		tgt_dates = sorted({str(t)[:10] for _, t in date_pairs if t})
		existing_tgt: set[tuple[str, str]] = set()
		if not overwrite and tgt_dates:
			tph = ",".join("?" for _ in tgt_dates)
			for er in conn.execute(f"SELECT plan_date, meal_slot FROM cafeteria_menu_plans WHERE plan_date IN ({tph})", tgt_dates).fetchall():
				existing_tgt.add((str(er["plan_date"])[:10], str(er["meal_slot"])))

	copied = 0
	for s_dt, t_dt in date_pairs:
		s_iso, t_iso = str(s_dt)[:10], str(t_dt)[:10]
		for src_rec in by_src.get(s_iso, []):
			sl = str(src_rec["meal_slot"])
			if meal_slots and sl not in meal_slots:
				continue
			if (not overwrite) and ((t_iso, sl) in existing_tgt):
				continue
			try:
				d_ids = json.loads(src_rec.get("dish_ids_json") or "[]")
			except Exception:
				d_ids = []
			try:
				cust = json.loads(src_rec.get("custom_items_json") or "[]")
			except Exception:
				cust = []
			upsert_menu_plan_slot(
				connect_db_fn,
				plan_date=t_iso,
				meal_slot=sl,
				dish_ids=d_ids,
				custom_items=cust,
				notes=str(src_rec.get("notes") or ""),
				sync_to_diabetes=sync_to_diabetes,
			)
			copied += 1
	return copied


def _group_month_days_by_week(days_info: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
	weeks: list[list[dict[str, Any]]] = []
	cur: list[dict[str, Any]] = []
	cur_wk: int | None = None
	for d in days_info:
		if not d.get("is_workday"):
			continue
		wk = int(d["iso_week"])
		if cur_wk is None or wk != cur_wk:
			if cur:
				weeks.append(cur)
			cur = [d]
			cur_wk = wk
		else:
			cur.append(d)
	if cur:
		weeks.append(cur)
	return weeks


def _decode_school_logo_bytes(school_meta: dict[str, str] | None) -> bytes | None:
	"""Ayarlardaki base64 okul logosunu çözümler; yoksa None döndürür."""
	raw = str((school_meta or {}).get("cafeteria_school_logo_b64") or "").strip()
	if not raw:
		return None
	if "," in raw and raw.startswith("data:"):
		raw = raw.split(",", 1)[1]
	try:
		b = base64.b64decode(raw, validate=False)
		return b if len(b) > 64 else None
	except Exception:
		return None


def _build_weekly_grid_rows(days_list: list[dict[str, Any]], include_weekends: bool) -> list[dict[int, dict[str, Any]]]:
	"""Gün listesini Pazartesi..Cuma (veya Pazar) sütunlu haftalık satır sözlüklerine dönüştürür."""
	max_wd = 6 if include_weekends else 4
	rows: list[dict[int, dict[str, Any]]] = []
	cur: dict[int, dict[str, Any]] = {}
	cur_wk: int | None = None
	for d in days_list:
		wd = int(d.get("weekday", 0))
		if wd > max_wd:
			continue
		wk = int(d.get("iso_week", 0))
		if cur_wk is None or wk != cur_wk:
			if cur:
				rows.append(cur)
			cur = {wd: d}
			cur_wk = wk
		else:
			cur[wd] = d
	if cur:
		rows.append(cur)
	return rows


DEFAULT_MENU_COLOR_TONE: str = "Çok Açık Pastel (Varsayılan)"

MENU_COLOR_TONE_PRESETS: dict[str, tuple[float, float]] = {
	"Beyaz / Renksiz": (0.0, 0.04),
	"Ultra Açık Pastel": (0.03, 0.10),
	"Çok Açık Pastel (Varsayılan)": (0.05, 0.16),
	"Yumuşak Pastel": (0.09, 0.24),
	"Orta Ton Pastel": (0.14, 0.34),
	"Belirgin Ton": (0.20, 0.45),
}

_MENU_WEEK_BASE_COLORS: list[dict[str, str]] = [
	{"short": "Sarı", "base": "#FACC15", "accent": "#CA8A04", "date_fg": "#713F12"},
	{"short": "Turuncu", "base": "#FB923C", "accent": "#EA580C", "date_fg": "#9A3412"},
	{"short": "Yeşil", "base": "#4ADE80", "accent": "#16A34A", "date_fg": "#14532D"},
	{"short": "Mavi", "base": "#60A5FA", "accent": "#2563EB", "date_fg": "#1E3A8A"},
	{"short": "Mor", "base": "#A78BFA", "accent": "#7C3AED", "date_fg": "#4C1D95"},
	{"short": "Gül", "base": "#FB7185", "accent": "#E11D48", "date_fg": "#881337"},
]


def _hex_to_rgb_tuple(hex_color: str) -> tuple[int, int, int]:
	h = str(hex_color or "0F172A").strip().lstrip("#")
	if len(h) != 6:
		return (15, 23, 42)
	try:
		return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
	except Exception:
		return (15, 23, 42)


def _blend_hex_with_white(base_hex: str, ratio: float) -> str:
	r_clamped = max(0.0, min(1.0, float(ratio)))
	br, bg, bb = _hex_to_rgb_tuple(base_hex)
	nr = int(round(255.0 * (1.0 - r_clamped) + br * r_clamped))
	ng = int(round(255.0 * (1.0 - r_clamped) + bg * r_clamped))
	nb = int(round(255.0 * (1.0 - r_clamped) + bb * r_clamped))
	return f"#{nr:02X}{ng:02X}{nb:02X}"


def get_week_palette_pair(week_idx: int, color_tone: str | None = None) -> dict[str, str]:
	"""Verilen hafta sırası ve renk tonu ayarına göre açık/koyu arka plan + vurgu renklerini üretir."""
	tone_name = str(color_tone or DEFAULT_MENU_COLOR_TONE).strip()
	light_r, dark_r = MENU_COLOR_TONE_PRESETS.get(tone_name, MENU_COLOR_TONE_PRESETS[DEFAULT_MENU_COLOR_TONE])
	base_cfg = _MENU_WEEK_BASE_COLORS[max(0, int(week_idx)) % len(_MENU_WEEK_BASE_COLORS)]
	is_monochrome = (tone_name == "Beyaz / Renksiz")
	if is_monochrome:
		light_hex = "#FFFFFF"
		dark_hex = "#F8FAFC"
		accent_hex = "#94A3B8"
		date_fg_hex = "#0F766E"
	else:
		light_hex = _blend_hex_with_white(base_cfg["base"], light_r)
		dark_hex = _blend_hex_with_white(base_cfg["base"], dark_r)
		accent_hex = base_cfg["accent"]
		date_fg_hex = base_cfg["date_fg"]
	return {
		"short": base_cfg["short"],
		"light": light_hex,
		"dark": dark_hex,
		"accent": accent_hex,
		"date_fg": date_fg_hex,
	}


def _build_week_day_style_map(
	days_list: list[dict[str, Any]],
	*,
	start_week_idx: int = 0,
	color_tone: str | None = None,
) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]]]:
	"""Her hafta için ayrı renk tonu ve ardışık günlerde açık/koyu ton üretir."""
	style_by_iso: dict[str, dict[str, Any]] = {}
	legend_weeks: list[dict[str, Any]] = []
	seen_weeks: set[int] = set()
	prev_wk_key: tuple[int, int] | None = None
	cur_wk_pos = -1
	day_seq_in_week = 0

	ordered_days = sorted(days_list or [], key=lambda x: str(x.get("date_iso") or ""))
	for d in ordered_days:
		d_iso = str(d.get("date_iso") or "")[:10]
		if not d_iso:
			continue
		try:
			iso_cal = date.fromisoformat(d_iso).isocalendar()
			wk_key = (int(iso_cal[0]), int(iso_cal[1]))
		except Exception:
			wk_key = (0, int(d.get("iso_week") or 0))
		if prev_wk_key is None or wk_key != prev_wk_key:
			cur_wk_pos += 1
			prev_wk_key = wk_key
			day_seq_in_week = 0
		else:
			day_seq_in_week += 1

		eff_wk_idx = max(0, int(start_week_idx) + cur_wk_pos)
		pal = get_week_palette_pair(eff_wk_idx, color_tone=color_tone)
		is_dark = (day_seq_in_week % 2 == 1)
		bg_hex = pal["dark"] if is_dark else pal["light"]
		week_num = eff_wk_idx + 1
		style_by_iso[d_iso] = {
			"week_num": week_num,
			"palette_short": pal["short"],
			"tone": "dark" if is_dark else "light",
			"bg_hex": bg_hex,
			"bg_word": bg_hex.lstrip("#").upper(),
			"light_hex": pal["light"],
			"dark_hex": pal["dark"],
			"accent_hex": pal["accent"],
			"accent_word": pal["accent"].lstrip("#").upper(),
			"date_fg_hex": pal["date_fg"],
			"date_fg_rgb": _hex_to_rgb_tuple(pal["date_fg"]),
		}
		if week_num not in seen_weeks:
			seen_weeks.add(week_num)
			legend_weeks.append({
				"week_num": week_num,
				"label": f"{week_num}. Hafta ({pal['short']})",
				"light": pal["light"],
				"dark": pal["dark"],
				"accent": pal["accent"],
				"date_fg": pal["date_fg"],
			})
	return style_by_iso, legend_weeks


def resolve_menu_signer_slots(app_cfg: dict[str, str] | None = None) -> list[tuple[str, str]]:
	"""Yemekhane ayarlarındaki 3 imza yuvasını (Sol, Orta, Sağ) döndürür; kullanıcı özelleştirdikten sonra silinen slotu hemşire/müdür adıyla ezmez."""
	cfg = app_cfg or {}
	s1_r = str(cfg["cafeteria_sig1_role"] if "cafeteria_sig1_role" in cfg else "Hazırlayan (Revir / Diyetisyen)").strip()
	s2_r = str(cfg["cafeteria_sig2_role"] if "cafeteria_sig2_role" in cfg else "Yemekhane Sorumlusu").strip()
	s3_r = str(cfg["cafeteria_sig3_role"] if "cafeteria_sig3_role" in cfg else "Okul Müdürü (Onay)").strip()
	s1_raw = str(cfg.get("cafeteria_sig1_name") or "").strip()
	s2_raw = str(cfg.get("cafeteria_sig2_name") or "").strip()
	s3_raw = str(cfg.get("cafeteria_sig3_name") or "").strip()
	is_custom = (
		str(cfg.get("cafeteria_sigs_customized") or "0").strip() == "1"
		or bool(s1_raw or s2_raw or s3_raw)
	)
	if not is_custom:
		s1_raw = str(cfg.get("nurse_name") or "").strip()
		s3_raw = str(cfg.get("principal_name") or "").strip()
	return [(s1_r, s1_raw), (s2_r, s2_raw), (s3_r, s3_raw)]


def _normalize_signers_for_render(
	signers: list[tuple[str, str]] | None,
	meta: dict[str, str] | None,
) -> tuple[list[tuple[str, str]], bool]:
	"""İmzacı listesini ve varsa yalnızca en sağdaki imzacının sağa yaslanma bilgisini döndürür."""
	raw_slots = list(signers) if signers is not None else resolve_menu_signer_slots(meta)
	indexed = [(idx, str(r or "").strip(), str(p or "").strip()) for idx, (r, p) in enumerate(raw_slots) if (str(r or "").strip() or str(p or "").strip())]
	if not indexed:
		m = meta or {}
		if str(m.get("cafeteria_sigs_customized") or "0").strip() == "1":
			return [], False
		return [
			("Hazırlayan", str(m.get("nurse_name") or "").strip() or "........................"),
			("Onay / Okul Müdürü", str(m.get("principal_name") or "").strip() or "........................"),
		], False
	right_only = (len(raw_slots) >= 3 and len(indexed) == 1 and indexed[0][0] == len(raw_slots) - 1)
	return [(r, p) for _, r, p in indexed], right_only


def resolve_menu_signers_for_print(app_cfg: dict[str, str] | None = None) -> list[tuple[str, str]]:
	"""A4 HTML/Word baskısı için konumunu koruyarak 3 imzacı yuvasını döndürür."""
	return resolve_menu_signer_slots(app_cfg)


def _extract_meal_display_parts(
	rec: dict[str, Any],
	*,
	show_kcal: bool,
	show_carbs: bool,
	show_allergens: bool,
) -> dict[str, Any]:
	"""Öğün yemeklerini yan yana ve alt satırda kalori + karbonhidrat + alerjen bilgilerini döndürür."""
	names = [str(nm).strip() for nm in (rec.get("dish_names") or []) if str(nm).strip()]
	dishes_inline = " · ".join(names)
	kv = _safe_float(rec.get("total_kcal"), 0.0)
	cv = _safe_float(rec.get("total_carbs_g"), 0.0)
	als = [str(a).strip() for a in (rec.get("allergens") or []) if str(a).strip()]
	return {
		"dishes_inline": dishes_inline,
		"kv": kv,
		"cv": cv,
		"als": als,
		"has_kcal": bool(show_kcal and kv > 0),
		"has_carbs": bool(show_carbs and (cv > 0 or kv > 0)),
		"show_allergens": bool(show_allergens),
	}


def _prepare_a4_menu_context(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	scope: str,
	week_index: int,
	active_slots: list[str] | None,
	include_weekends: bool,
	skip_empty_days: bool,
	school_meta: dict[str, str] | None,
	color_tone: str | None = None,
) -> dict[str, Any]:
	"""A4 HTML ve Word çıktıları için gün filtresi (boş günleri otomatik çıkarma), hizmet düzeni, renk tonu ve yemek saatlerini hazırlar."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy, mm = [int(x) for x in ym.split("-")]
	month_title = f"{TR_MONTH_NAMES.get(mm, mm)} {yy}"
	meta = dict(school_meta or {})
	if not meta:
		try:
			with connect_db_fn() as conn:
				for sr in conn.execute("SELECT key, value FROM app_settings").fetchall():
					meta[str(sr["key"])] = str(sr["value"])
		except Exception:
			pass

	sched = parse_service_schedule_from_cfg(meta)
	slots = [s for s in (active_slots or sched["active_meal_slots"] or ["Öğle Yemeği"]) if s]
	meal_times = sched["meal_times"]
	days_info = get_month_workdays_info(connect_db_fn, ym, include_weekends=include_weekends)
	plans = get_menu_plans_for_month(connect_db_fn, ym)

	weeks = _group_month_days_by_week(days_info)
	if scope == "weekly" and weeks:
		safe_idx = max(0, min(week_index, len(weeks) - 1))
		raw_days = weeks[safe_idx]
		range_txt = f"{raw_days[0]['date_iso']} – {raw_days[-1]['date_iso']}"
	else:
		safe_idx = 0
		raw_days = [d for d in days_info if d.get("is_workday")]
		range_txt = ""

	if skip_empty_days:
		target_days = [
			d for d in raw_days
			if any(plans.get((d["date_iso"], sl)) and plans[(d["date_iso"], sl)].get("dish_names") for sl in slots)
		]
	else:
		target_days = raw_days

	planned_meal_count = sum(
		1 for d in target_days for sl in slots
		if plans.get((d["date_iso"], sl)) and plans[(d["date_iso"], sl)].get("dish_names")
	)
	if scope == "weekly" and weeks:
		scope_label = f"{safe_idx + 1}. Hafta Yemek Menüsü ({range_txt} · {planned_meal_count} Öğün)"
	else:
		scope_label = f"{month_title} Aylık Yemek Menüsü ({planned_meal_count} Öğün)"

	days_names_s = ", ".join(nm[:3] for nm in sched["weekday_names"])
	slot_times_parts = [f"{sl}: {meal_times.get(sl, '')}" if meal_times.get(sl) else sl for sl in slots]
	schedule_line = f"Haftalık Hizmet: {sched['days_per_week']} Gün ({days_names_s}) · Günlük {len(slots)} Öğün · " + " | ".join(slot_times_parts)
	ramazan_times: dict[str, dict[str, str]] = {}
	if sched["ramazan_mode"]:
		ramazan_times = get_ramazan_times_for_dates(connect_db_fn, sched["ramazan_province"], [d["date_iso"] for d in target_days])
		schedule_line += f" · 🌙 Ramazan Modu ({sched['ramazan_province']} İftar & Sahur Saatleri)"
	eff_wknds = include_weekends or any(wd >= 5 for wd in sched["weekdays_set"])
	eff_tone = str(color_tone or meta.get("cafeteria_menu_color_tone") or DEFAULT_MENU_COLOR_TONE).strip()
	day_styles, _ = _build_week_day_style_map(
		target_days,
		start_week_idx=safe_idx if (scope == "weekly" and weeks) else 0,
		color_tone=eff_tone,
	)

	return {
		"ym": ym,
		"yy": yy,
		"mm": mm,
		"slots": slots,
		"meal_times": meal_times,
		"plans": plans,
		"target_days": target_days,
		"scope_label": scope_label,
		"schedule_line": schedule_line,
		"meta": meta,
		"eff_wknds": eff_wknds,
		"ramazan_mode": sched["ramazan_mode"],
		"ramazan_province": sched["ramazan_province"],
		"ramazan_times": ramazan_times,
		"day_styles": day_styles,
	}


def _render_html_meal_cell(
	rec: dict[str, Any],
	*,
	show_kcal: bool,
	show_carbs: bool,
	show_allergens: bool,
) -> str:
	parts = _extract_meal_display_parts(
		rec, show_kcal=show_kcal, show_carbs=show_carbs, show_allergens=show_allergens
	)
	dishes_html = f"<div class='dishes'>{html.escape(parts['dishes_inline'])}</div>"
	meta_bits: list[str] = []
	if parts["has_kcal"]:
		meta_bits.append(f"<span class='meta-pill kcal-pill'>🔥 {parts['kv']:g} kcal</span>")
	if parts["has_carbs"]:
		meta_bits.append(f"<span class='meta-pill carb-pill'>🍞 {parts['cv']:g} g Karb.</span>")
	if parts["show_allergens"]:
		if parts["als"]:
			al_str = " · ".join(format_allergen_badge(a) for a in parts["als"])
			meta_bits.append(f"<span class='meta-pill al-pill'>⚠️ {html.escape(al_str)}</span>")
		else:
			meta_bits.append("<span class='meta-pill ok-pill'>✅ Alerjen: Yok</span>")
	meta_div = f"<div class='cell-meta'>{' '.join(meta_bits)}</div>" if meta_bits else ""
	return f"{dishes_html}{meta_div}"


def build_menu_a4_html(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	scope: str = "monthly",
	week_index: int = 0,
	active_slots: list[str] | None = None,
	show_kcal: bool = True,
	show_carbs: bool = True,
	show_allergens: bool = True,
	include_weekends: bool = False,
	layout_mode: str = "standard",
	skip_empty_days: bool = True,
	signers: list[tuple[str, str]] | None = None,
	school_meta: dict[str, str] | None = None,
	color_tone: str | None = None,
) -> str:
	"""Aylık veya haftalık menüyü standart, tek A4 kompakt veya haftalık ızgara düzeninde haftalık renk tonlu imzalı HTML olarak oluşturur."""
	ctx_p = _prepare_a4_menu_context(
		connect_db_fn,
		year_month=year_month,
		scope=scope,
		week_index=week_index,
		active_slots=active_slots,
		include_weekends=include_weekends,
		skip_empty_days=skip_empty_days,
		school_meta=school_meta,
		color_tone=color_tone,
	)
	yy, mm = ctx_p["yy"], ctx_p["mm"]
	slots = ctx_p["slots"]
	meal_times = ctx_p["meal_times"]
	plans = ctx_p["plans"]
	target_days = ctx_p["target_days"]
	scope_label = ctx_p["scope_label"]
	schedule_line = ctx_p["schedule_line"]
	meta = ctx_p["meta"]
	eff_wknds = ctx_p["eff_wknds"]
	ramazan_mode = bool(ctx_p.get("ramazan_mode"))
	ramazan_times = ctx_p.get("ramazan_times") or {}
	day_styles = ctx_p.get("day_styles") or {}

	inst_top = html.escape(meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	school_nm = html.escape(meta.get("school_name") or "Okul / Kurum Yemekhanesi")
	custom_board_note = str(meta.get("cafeteria_board_note") or "").strip()
	footer_note = html.escape(
		custom_board_note
		or "* Gıda alerjisi veya özel beslenme gereksinimi (Çölyak, Tip 1 Diyabet vb.) bulunan öğrencilerimiz için okul reviri ile koordineli alternatif sunum uygulanmaktadır."
	)
	logo_bytes = _decode_school_logo_bytes(meta)
	logo_img_html = ""
	if logo_bytes:
		b64_s = base64.b64encode(logo_bytes).decode("ascii")
		mime_l = "image/png" if logo_bytes[:4] == b"\x89PNG" else "image/jpeg"
		logo_img_html = f"<img src='data:{mime_l};base64,{b64_s}' style='height:46px;max-width:90px;object-fit:contain;margin-bottom:3px;' alt='Okul Logosu'><br>"

	sig_list, sig_right_only = _normalize_signers_for_render(signers, meta)

	is_compact = (layout_mode == "compact_one_page")
	is_grid = (layout_mode == "weekly_grid")

	if is_grid:
		wd_cols = sorted({int(d["weekday"]) for d in target_days}) or [0, 1, 2, 3, 4]
		hdr_cols = [f"<th>{TR_WEEKDAYS_SHORT[wd]}</th>" for wd in wd_cols]
		grid_weeks = _build_weekly_grid_rows(target_days, eff_wknds or any(w >= 5 for w in wd_cols))
		rows_html = []
		for wk_dict in grid_weeks:
			tds = []
			for wd in wd_cols:
				d = wk_dict.get(wd)
				if not d:
					tds.append("<td class='empty-cell'>—</td>")
					continue
				d_iso = d["date_iso"]
				st = day_styles.get(d_iso, {})
				bg_hex = st.get("bg_hex", "#FEF9C3")
				acc_hex = st.get("accent_hex", "#CA8A04")
				fg_hex = st.get("date_fg_hex", "#0F172A")
				wk_num = st.get("week_num", 1)
				rt_d = ramazan_times.get(d_iso) if ramazan_mode else None
				ram_sub = f"<div style='font-size:8.5px;color:#0f766e;font-weight:700;'>🌙 S:{rt_d['sahur']} · 🌅 İ:{rt_d['iftar']}</div>" if rt_d else ""
				cell_parts = [
					f"<div class='grid-dt' style='color:{fg_hex};border-bottom-color:{acc_hex};'>"
					f"{d['day_num']:02d}.{mm:02d}.{yy} <span class='wk-tag' style='border-color:{acc_hex};color:{fg_hex};'>{wk_num}. Hft</span>"
					f"{ram_sub}</div>"
				]
				for sl in slots:
					rec = plans.get((d_iso, sl))
					if rec and rec.get("dish_names"):
						tm_s = f" ({html.escape(meal_times[sl])})" if meal_times.get(sl) else ""
						sl_hdr = f"<div class='grid-sl'>{html.escape(sl)}{tm_s}</div>" if (len(slots) > 1 or tm_s) else ""
						cell_parts.append(
							f"{sl_hdr}"
							+ _render_html_meal_cell(
								rec, show_kcal=show_kcal, show_carbs=show_carbs, show_allergens=show_allergens
							)
						)
					else:
						cell_parts.append("<div class='empty-cell'>—</div>")
				tds.append(f"<td style='background:{bg_hex};border-top:3px solid {acc_hex};'>{''.join(cell_parts)}</td>")
			rows_html.append(f"<tr>{''.join(tds)}</tr>")
	else:
		rows_html = []
		for d in target_days:
			d_iso = d["date_iso"]
			st = day_styles.get(d_iso, {})
			bg_hex = st.get("bg_hex", "#FEF9C3")
			acc_hex = st.get("accent_hex", "#CA8A04")
			fg_hex = st.get("date_fg_hex", "#0F172A")
			wk_num = st.get("week_num", 1)
			rt_d = ramazan_times.get(d_iso) if ramazan_mode else None
			ram_line = (
				f"<br><small style='color:#0f766e;font-weight:700;'>🌙 S:{rt_d['sahur']} · 🌅 İ:{rt_d['iftar']}</small>"
				if (rt_d and is_compact)
				else (f"<br><small style='color:#0f766e;font-weight:700;'>🌙 Sahur: {rt_d['sahur']}<br>🌅 İftar: {rt_d['iftar']}</small>" if rt_d else "")
			)
			wk_badge = f" <span class='wk-tag' style='border-color:{acc_hex};color:{fg_hex};'>{wk_num}. Hft</span>"
			dt_tr = f"{d['day_num']:02d}.{mm:02d}.{yy}{wk_badge}<br><small style='color:{fg_hex};'>{html.escape(d['weekday_name'])}</small>{ram_line}"
			day_kcal = 0.0
			day_carbs = 0.0
			cells = [f"<td class='col-date' style='background:{bg_hex};color:{fg_hex};border-left:4px solid {acc_hex};'>{dt_tr}</td>"]
			for sl in slots:
				rec = plans.get((d_iso, sl))
				if rec and rec.get("dish_names"):
					day_kcal += _safe_float(rec.get("total_kcal"), 0.0)
					day_carbs += _safe_float(rec.get("total_carbs_g"), 0.0)
					inner_h = _render_html_meal_cell(
						rec, show_kcal=show_kcal, show_carbs=show_carbs, show_allergens=show_allergens
					)
					cells.append(f"<td style='background:{bg_hex};'>{inner_h}</td>")
				else:
					cells.append(f"<td class='empty-cell' style='background:{bg_hex};'>—</td>")
			if (show_kcal or show_carbs) and len(slots) > 1:
				tot_parts = []
				if show_kcal and day_kcal > 0:
					tot_parts.append(f"<b>{day_kcal:g} kcal</b>")
				if show_carbs and (day_carbs > 0 or day_kcal > 0):
					tot_parts.append(f"<small style='color:#1e40af;font-weight:700;'>{day_carbs:g} g Karb.</small>")
				tot_txt = "<br>".join(tot_parts) if tot_parts else "—"
				cells.append(f"<td class='col-tot' style='background:{bg_hex};color:{fg_hex};'>{tot_txt}</td>")
			rows_html.append(f"<tr>{''.join(cells)}</tr>")

		hdr_cols = ["<th>Tarih / Gün</th>"] + [
			f"<th>{html.escape(sl)}" + (f"<br><small style='font-weight:500;opacity:.92'>{html.escape(meal_times[sl])}</small>" if meal_times.get(sl) else "") + "</th>"
			for sl in slots
		]
		if (show_kcal or show_carbs) and len(slots) > 1:
			hdr_cols.append("<th>Günlük Toplam</th>")

	sig_box_style = "margin-left:auto;flex:0 0 34%;max-width:34%;" if sig_right_only else ""
	sig_html = "".join(
		f"<div class='sig-box' style='{sig_box_style}'><div class='sig-role'>{html.escape(str(role))}</div>"
		f"<div class='sig-name'>{html.escape(str(person or '................................'))}</div>"
		f"<div class='sig-line'>İmza / Tarih</div></div>"
		for role, person in sig_list
	)

	page_orient = "landscape" if (is_grid or (len(slots) >= 2 and scope != "weekly")) else "portrait"
	page_w = "297mm" if page_orient == "landscape" else "210mm"
	tbl_font = "9.2px" if is_compact else ("10px" if is_grid else "11px")
	cell_pad = "2.8px 5px" if is_compact else "5px 7px"
	sig_top = "10px" if is_compact else "22px"

	return f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{html.escape(scope_label)}</title>
<style>
@page {{ size: A4 {page_orient}; margin: {'6mm 8mm' if is_compact else '10mm 12mm'}; }}
* {{ box-sizing: border-box; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color: #0f172a; background: #cbd5e1; margin: 0; padding: 14px; }}
.topbar {{ max-width: {page_w}; margin: 0 auto 10px auto; display: flex; justify-content: space-between; align-items: center; }}
.btn-print {{ padding: 8px 16px; background: #0f766e; color: #fff; border: none; border-radius: 7px; font-weight: 700; cursor: pointer; }}
.a4-sheet {{ width: min(100%, {page_w}); margin: 0 auto; padding: {'6mm 8mm' if is_compact else '11mm 13mm'}; background: #fff; border: 1px solid #94a3b8; box-shadow: 0 8px 24px rgba(15,23,42,.18); }}
.antet {{ text-align: center; border-bottom: 2px solid #0f766e; padding-bottom: 6px; margin-bottom: 8px; }}
.antet-top {{ font-size: 10.5px; color: #64748b; letter-spacing: .5px; }}
.antet-school {{ font-size: 15px; font-weight: 800; color: #0f766e; margin-top: 2px; }}
.antet-title {{ font-size: 12.5px; font-weight: 700; color: #1e293b; margin-top: 2px; }}
.antet-sched {{ font-size: 10px; color: #475569; margin-top: 3px; }}
.wk-tag {{ display: inline-block; font-size: {'8px' if is_compact else '8.5px'}; font-weight: 700; background: rgba(255,255,255,0.82); border: 1px solid #94a3b8; border-radius: 4px; padding: 0 3px; margin-left: 3px; vertical-align: middle; }}
table {{ width: 100%; border-collapse: collapse; font-size: {tbl_font}; table-layout: {'fixed' if is_grid else 'auto'}; }}
th, td {{ border: 1px solid #94a3b8; padding: {cell_pad}; vertical-align: top; text-align: left; color: #0f172a; }}
th {{ background: #0f766e; color: #fff; font-weight: 700; text-align: center; }}
.col-date {{ width: 108px; font-weight: 700; white-space: nowrap; }}
.col-tot {{ width: 88px; text-align: center; vertical-align: middle; }}
.grid-dt {{ font-weight: 800; border-bottom: 1px solid #94a3b8; margin-bottom: 4px; padding-bottom: 2px; font-size: 10px; }}
.grid-sl {{ font-weight: 700; color: #0f172a; font-size: 9.5px; margin-top: 3px; }}
.dishes {{ font-weight: 600; color: #0f172a; line-height: {'1.22' if is_compact else '1.35'}; }}
.cell-meta {{ margin-top: 3px; display: flex; flex-wrap: wrap; gap: 3px; align-items: center; }}
.meta-pill {{ display: inline-block; background: rgba(255,255,255,0.92); border: 1px solid rgba(15,23,42,0.18); border-radius: 4px; padding: 1px 5px; font-size: {'8.2px' if is_compact else '9px'}; font-weight: 700; line-height: 1.25; }}
.kcal-pill {{ color: #92400e; border-color: #f59e0b; }}
.carb-pill {{ color: #1e40af; border-color: #60a5fa; }}
.al-pill {{ color: #991b1b; border-color: #f87171; background: rgba(255,245,245,0.95); }}
.ok-pill {{ color: #166534; border-color: #86efac; font-weight: 600; }}
.empty-cell {{ text-align: center; color: #64748b; background: #f8fafc; }}
.note-line {{ margin-top: 6px; font-size: 9.5px; color: #475569; }}
.sig-row {{ margin-top: {sig_top}; display: flex; justify-content: space-around; gap: 16px; page-break-inside: avoid; }}
.sig-box {{ flex: 1; text-align: center; border-top: 1px solid #94a3b8; padding-top: 5px; font-size: 10.5px; }}
.sig-role {{ font-weight: 700; color: #1e293b; }}
.sig-name {{ font-weight: 600; color: #0f766e; margin-top: 2px; }}
.sig-line {{ font-size: 9.5px; color: #64748b; margin-top: 14px; }}
@media print {{
  .topbar {{ display: none !important; }}
  body {{ background: #fff; padding: 0; }}
  .a4-sheet {{ width: 100%; border: none; box-shadow: none; padding: 0; }}
}}
</style>
</head>
<body>
<div class='topbar'>
  <span style='font-weight:700;font-size:12px;color:#334155;'>🍽️ A4 Menü Önizleme ({html.escape(scope_label)})</span>
  <button class='btn-print' onclick='window.print()'>🖨️ A4 Yazdır / PDF Kaydet</button>
</div>
<div class='a4-sheet'>
  <div class='antet'>
    {logo_img_html}
    <div class='antet-top'>{inst_top}</div>
    <div class='antet-school'>{school_nm}</div>
    <div class='antet-title'>{html.escape(scope_label)}</div>
    <div class='antet-sched'>{html.escape(schedule_line)}</div>
  </div>
  <table>
    <thead><tr>{''.join(hdr_cols)}</tr></thead>
    <tbody>{''.join(rows_html) if rows_html else f'<tr><td colspan="{max(2, len(hdr_cols))}" class="empty-cell">Bu dönem için öğün girilmiş gün bulunamadı.</td></tr>'}</tbody>
  </table>
  <div class='note-line'>{footer_note}</div>
  <div class='sig-row'>{sig_html}</div>
</div>
</body>
</html>"""


def _ensure_docx_mod():
	import docx
	from docx.enum.section import WD_ORIENT
	from docx.enum.table import WD_TABLE_ALIGNMENT
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.oxml import OxmlElement, parse_xml
	from docx.oxml.ns import nsdecls, qn
	from docx.shared import Cm, Pt, RGBColor
	return {
		"docx": docx,
		"WD_ORIENT": WD_ORIENT,
		"WD_TABLE_ALIGNMENT": WD_TABLE_ALIGNMENT,
		"WD_ALIGN_PARAGRAPH": WD_ALIGN_PARAGRAPH,
		"OxmlElement": OxmlElement,
		"parse_xml": parse_xml,
		"nsdecls": nsdecls,
		"qn": qn,
		"Cm": Cm,
		"Pt": Pt,
		"RGBColor": RGBColor,
	}


def _style_docx_cell(
	dx: dict[str, Any],
	cell: Any,
	bg_hex: str,
	compact: bool = False,
	left_accent_hex: str = "",
	top_accent_hex: str = "",
) -> None:
	tcPr = cell._tc.get_or_add_tcPr()
	if left_accent_hex or top_accent_hex:
		b_parts = []
		if top_accent_hex:
			b_parts.append(f'<w:top w:val="single" w:sz="18" w:space="0" w:color="{top_accent_hex}"/>')
		if left_accent_hex:
			b_parts.append(f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{left_accent_hex}"/>')
		tcPr.append(dx["parse_xml"](f'<w:tcBorders {dx["nsdecls"]("w")}>{"".join(b_parts)}</w:tcBorders>'))
	shading = dx["parse_xml"](f'<w:shd {dx["nsdecls"]("w")} w:fill="{bg_hex}"/>')
	tcPr.append(shading)
	if compact:
		mar = dx["parse_xml"](
			f'<w:tcMar {dx["nsdecls"]("w")}>'
			'<w:top w:w="28" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/>'
			'<w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/>'
			'</w:tcMar>'
		)
		tcPr.append(mar)


def build_menu_a4_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	scope: str = "monthly",
	week_index: int = 0,
	active_slots: list[str] | None = None,
	show_kcal: bool = True,
	show_carbs: bool = True,
	show_allergens: bool = True,
	include_weekends: bool = False,
	layout_mode: str = "standard",
	skip_empty_days: bool = True,
	signers: list[tuple[str, str]] | None = None,
	school_meta: dict[str, str] | None = None,
	color_tone: str | None = None,
) -> bytes:
	"""Aylık veya haftalık yemek menüsünü standart, tek A4 kompakt veya haftalık ızgara düzeninde haftalık renk tonlu Word (.docx) belgesi olarak üretir."""
	dx = _ensure_docx_mod()
	docx = dx["docx"]
	Cm, Pt, RGBColor = dx["Cm"], dx["Pt"], dx["RGBColor"]

	ctx_p = _prepare_a4_menu_context(
		connect_db_fn,
		year_month=year_month,
		scope=scope,
		week_index=week_index,
		active_slots=active_slots,
		include_weekends=include_weekends,
		skip_empty_days=skip_empty_days,
		school_meta=school_meta,
		color_tone=color_tone,
	)
	yy, mm = ctx_p["yy"], ctx_p["mm"]
	slots = ctx_p["slots"]
	meal_times = ctx_p["meal_times"]
	plans = ctx_p["plans"]
	target_days = ctx_p["target_days"]
	scope_label = ctx_p["scope_label"]
	schedule_line = ctx_p["schedule_line"]
	meta = ctx_p["meta"]
	eff_wknds = ctx_p["eff_wknds"]
	ramazan_mode = bool(ctx_p.get("ramazan_mode"))
	ramazan_times = ctx_p.get("ramazan_times") or {}
	day_styles = ctx_p.get("day_styles") or {}

	inst_top = meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI"
	school_nm = meta.get("school_name") or "Okul / Kurum Yemekhanesi"
	custom_board_note = str(meta.get("cafeteria_board_note") or "").strip()

	is_compact = (layout_mode == "compact_one_page")
	is_grid = (layout_mode == "weekly_grid")

	doc = docx.Document()
	sec = doc.sections[0]
	is_landscape = is_grid or (len(slots) >= 2 and scope != "weekly")
	sec.orientation = dx["WD_ORIENT"].LANDSCAPE if is_landscape else dx["WD_ORIENT"].PORTRAIT
	sec.page_width = Cm(29.7 if is_landscape else 21.0)
	sec.page_height = Cm(21.0 if is_landscape else 29.7)
	sec.top_margin = Cm(0.65 if is_compact else 1.0)
	sec.bottom_margin = Cm(0.65 if is_compact else 1.0)
	sec.left_margin = Cm(0.85 if is_compact else 1.1)
	sec.right_margin = Cm(0.85 if is_compact else 1.1)

	logo_bytes = _decode_school_logo_bytes(meta)
	if logo_bytes:
		try:
			p_lg = doc.add_paragraph()
			p_lg.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
			p_lg.paragraph_format.space_before = Pt(0)
			p_lg.paragraph_format.space_after = Pt(2)
			p_lg.add_run().add_picture(io.BytesIO(logo_bytes), height=Cm(1.3))
		except Exception:
			pass

	for txt_line, fsz, bold_f, rgb_c in [
		(inst_top, 8.5 if is_compact else 9.5, False, RGBColor(100, 116, 139)),
		(school_nm.upper(), 11.0 if is_compact else 12.5, True, RGBColor(15, 118, 110)),
		(scope_label, 10.0 if is_compact else 11.0, True, RGBColor(30, 41, 59)),
		(schedule_line, 7.8 if is_compact else 8.5, False, RGBColor(71, 85, 105)),
	]:
		p = doc.add_paragraph()
		p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		p.paragraph_format.space_before = Pt(0)
		p.paragraph_format.space_after = Pt(1.5 if is_compact else 2)
		r = p.add_run(txt_line)
		r.font.name = "Calibri"
		r.font.size = Pt(fsz)
		r.bold = bold_f
		r.font.color.rgb = rgb_c

	borders = dx["parse_xml"](
		f'<w:tblBorders {dx["nsdecls"]("w")}>'
		'<w:top w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'<w:left w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'<w:right w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
		'</w:tblBorders>'
	)

	def _render_hdr(tbl_obj: Any, headers: list[str]) -> None:
		tbl_obj._tbl.tblPr.append(borders)
		for ci, h_txt in enumerate(headers):
			c = tbl_obj.cell(0, ci)
			_style_docx_cell(dx, c, "0F766E", compact=is_compact)
			p = c.paragraphs[0]
			p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
			p.paragraph_format.space_after = Pt(0)
			rr = p.add_run(h_txt)
			rr.font.name = "Calibri"
			rr.font.size = Pt(8.0 if is_compact else 9.0)
			rr.bold = True
			rr.font.color.rgb = RGBColor(255, 255, 255)

	body_fsz = 7.3 if is_compact else 8.5
	meta_fsz = 6.6 if is_compact else 7.6

	def _append_docx_meal_runs(cp: Any, rec: dict[str, Any], *, trailing_nl: bool = False) -> None:
		parts = _extract_meal_display_parts(
			rec, show_kcal=show_kcal, show_carbs=show_carbs, show_allergens=show_allergens
		)
		r_d = cp.add_run(parts["dishes_inline"])
		r_d.font.name = "Calibri"
		r_d.font.size = Pt(body_fsz)
		r_d.bold = True
		r_d.font.color.rgb = RGBColor(15, 23, 42)

		seg_runs: list[tuple[str, Any, bool]] = []
		if parts["has_kcal"]:
			seg_runs.append((f"Kalori: {parts['kv']:g} kcal", RGBColor(146, 64, 14), True))
		if parts["has_carbs"]:
			seg_runs.append((f"Karb.: {parts['cv']:g} g", RGBColor(30, 64, 175), True))
		if parts["show_allergens"]:
			if parts["als"]:
				seg_runs.append(("Alerjen: " + ", ".join(parts["als"]), RGBColor(153, 27, 27), True))
			else:
				seg_runs.append(("Alerjen: Yok", RGBColor(22, 101, 52), False))
		if seg_runs:
			cp.add_run("\n")
			for si_r, (txt_r, col_r, b_r) in enumerate(seg_runs):
				if si_r > 0:
					r_sep = cp.add_run("  |  ")
					r_sep.font.name = "Calibri"
					r_sep.font.size = Pt(meta_fsz)
					r_sep.font.color.rgb = RGBColor(71, 85, 105)
				rm = cp.add_run(txt_r)
				rm.font.name = "Calibri"
				rm.font.size = Pt(meta_fsz)
				rm.bold = b_r
				rm.font.color.rgb = col_r
		if trailing_nl:
			cp.add_run("\n")

	if not target_days:
		p_emp = doc.add_paragraph()
		p_emp.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		r_emp = p_emp.add_run("Bu dönem için henüz öğün girilmiş gün bulunmamaktadır.")
		r_emp.font.size = Pt(9.5)
		r_emp.font.color.rgb = RGBColor(100, 116, 139)
	elif is_grid:
		wd_cols = sorted({int(d["weekday"]) for d in target_days}) or [0, 1, 2, 3, 4]
		cols_hdr = [TR_WEEKDAYS_SHORT[wd] for wd in wd_cols]
		grid_weeks = _build_weekly_grid_rows(target_days, eff_wknds or any(w >= 5 for w in wd_cols))
		tbl = doc.add_table(rows=max(1, len(grid_weeks)) + 1, cols=len(cols_hdr))
		tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
		tbl.autofit = True
		_render_hdr(tbl, cols_hdr)
		for ri, wk_dict in enumerate(grid_weeks, start=1):
			for ci, wd in enumerate(wd_cols):
				cell = tbl.cell(ri, ci)
				cp = cell.paragraphs[0]
				cp.paragraph_format.space_after = Pt(0)
				d = wk_dict.get(wd)
				if not d:
					_style_docx_cell(dx, cell, "F8FAFC", compact=is_compact)
					cp.add_run("—").font.size = Pt(body_fsz)
					continue
				d_iso = d["date_iso"]
				st = day_styles.get(d_iso, {})
				bg_w = st.get("bg_word", "FEF9C3")
				acc_w = st.get("accent_word", "CA8A04")
				fg_rgb = RGBColor(*st.get("date_fg_rgb", (15, 23, 42)))
				wk_num = st.get("week_num", 1)
				_style_docx_cell(dx, cell, bg_w, compact=is_compact, top_accent_hex=acc_w)
				rt_d = ramazan_times.get(d_iso) if ramazan_mode else None
				ram_g = f"\n🌙 S:{rt_d['sahur']} · İ:{rt_d['iftar']}\n" if rt_d else "\n"
				rd = cp.add_run(f"{d['day_num']:02d}.{mm:02d}.{yy} ({wk_num}. Hft){ram_g}")
				rd.font.name = "Calibri"
				rd.font.size = Pt(body_fsz)
				rd.bold = True
				rd.font.color.rgb = fg_rgb
				active_recs = [(sl, plans.get((d_iso, sl))) for sl in slots]
				for idx_sl, (sl, rec) in enumerate(active_recs):
					if rec and rec.get("dish_names"):
						tm_s = f" · {meal_times[sl]}" if meal_times.get(sl) else ""
						if len(slots) > 1 or tm_s:
							rs = cp.add_run(f"[{sl}{tm_s}]\n")
							rs.font.size = Pt(meta_fsz)
							rs.bold = True
							rs.font.color.rgb = fg_rgb
						_append_docx_meal_runs(cp, rec, trailing_nl=(idx_sl < len(active_recs) - 1))
	else:
		slot_hdrs = [
			f"{sl}\n({meal_times[sl]})" if meal_times.get(sl) else sl
			for sl in slots
		]
		show_tot_col = (show_kcal or show_carbs) and len(slots) > 1
		cols_hdr = ["Tarih / Gün"] + slot_hdrs + (["Günlük Toplam"] if show_tot_col else [])
		tbl = doc.add_table(rows=len(target_days) + 1, cols=len(cols_hdr))
		tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
		tbl.autofit = True
		_render_hdr(tbl, cols_hdr)

		for ri, d in enumerate(target_days, start=1):
			d_iso = d["date_iso"]
			st = day_styles.get(d_iso, {})
			bg_w = st.get("bg_word", "FEF9C3")
			acc_w = st.get("accent_word", "CA8A04")
			fg_rgb = RGBColor(*st.get("date_fg_rgb", (15, 23, 42)))
			wk_num = st.get("week_num", 1)
			rt_d = ramazan_times.get(d_iso) if ramazan_mode else None

			c0 = tbl.cell(ri, 0)
			_style_docx_cell(dx, c0, bg_w, compact=is_compact, left_accent_hex=acc_w)
			p0 = c0.paragraphs[0]
			p0.paragraph_format.space_after = Pt(0)
			ram_txt = (
				f"\nS:{rt_d['sahur']} İ:{rt_d['iftar']}"
				if (rt_d and is_compact)
				else (f"\nSahur: {rt_d['sahur']}\nİftar: {rt_d['iftar']}" if rt_d else "")
			)
			dt_lbl = (
				f"{d['day_num']:02d}.{mm:02d} {d['weekday_name'][:3]} ({wk_num}.H)"
				if is_compact
				else f"{d['day_num']:02d}.{mm:02d}.{yy} ({wk_num}. Hft)\n{d['weekday_name']}"
			) + ram_txt
			r0 = p0.add_run(dt_lbl)
			r0.font.name = "Calibri"
			r0.font.size = Pt(body_fsz)
			r0.bold = True
			r0.font.color.rgb = fg_rgb

			day_kcal = 0.0
			day_carbs = 0.0
			for si, sl in enumerate(slots, start=1):
				cell = tbl.cell(ri, si)
				_style_docx_cell(dx, cell, bg_w, compact=is_compact)
				cp = cell.paragraphs[0]
				cp.paragraph_format.space_after = Pt(0)
				rec = plans.get((d_iso, sl))
				if rec and rec.get("dish_names"):
					day_kcal += _safe_float(rec.get("total_kcal"), 0.0)
					day_carbs += _safe_float(rec.get("total_carbs_g"), 0.0)
					_append_docx_meal_runs(cp, rec, trailing_nl=False)
				else:
					r_e = cp.add_run("—")
					r_e.font.size = Pt(body_fsz)

			if show_tot_col:
				ct = tbl.cell(ri, len(slots) + 1)
				_style_docx_cell(dx, ct, bg_w, compact=is_compact)
				pt = ct.paragraphs[0]
				pt.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
				pt.paragraph_format.space_after = Pt(0)
				tot_lines: list[str] = []
				if show_kcal and day_kcal > 0:
					tot_lines.append(f"{day_kcal:g} kcal")
				if show_carbs and (day_carbs > 0 or day_kcal > 0):
					tot_lines.append(f"Karb: {day_carbs:g} g")
				rt = pt.add_run("\n".join(tot_lines) if tot_lines else "—")
				rt.font.name = "Calibri"
				rt.font.size = Pt(body_fsz)
				rt.bold = True
				rt.font.color.rgb = fg_rgb

	if custom_board_note:
		p_note = doc.add_paragraph()
		p_note.paragraph_format.space_before = Pt(4)
		p_note.paragraph_format.space_after = Pt(2)
		rn = p_note.add_run(custom_board_note)
		rn.font.size = Pt(7.5 if is_compact else 8.5)
		rn.font.color.rgb = RGBColor(71, 85, 105)

	sig_list, sig_right_only = _normalize_signers_for_render(signers, meta)
	if sig_list:
		p_sp = doc.add_paragraph()
		p_sp.paragraph_format.space_before = Pt(6 if is_compact else 12)
		p_sp.paragraph_format.keep_with_next = True
		sig_tbl = doc.add_table(rows=1, cols=2 if sig_right_only else len(sig_list))
		sig_tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
		render_pairs = [("", ""), sig_list[0]] if sig_right_only else sig_list
		for idx, (role_lbl, p_name) in enumerate(render_pairs):
			if not (role_lbl or p_name):
				continue
			sc = sig_tbl.cell(0, idx)
			sp = sc.paragraphs[0]
			sp.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
			sr1 = sp.add_run(f"{role_lbl}\n")
			sr1.font.name = "Calibri"
			sr1.font.size = Pt(8.5 if is_compact else 9.5)
			sr1.bold = True
			sr2 = sp.add_run(f"{p_name or '........................'}\n\nİmza / Tarih")
			sr2.font.name = "Calibri"
			sr2.font.size = Pt(8.0 if is_compact else 8.5)
			sr2.font.color.rgb = RGBColor(71, 85, 105)

	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


# ==============================================================================
# 5. BÖLÜM 3: GELEN ÜRÜN, FATURA & GİDER MUHASEBESİ
# ==============================================================================

def upsert_purchase_invoice(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	purchase_id: int | None = None,
	invoice_date: str,
	invoice_no: str = "",
	supplier_name: str = "",
	category: str = "Diğer",
	items: list[dict[str, Any]] | None = None,
	payment_status: str = "Ödendi",
	document_name: str = "",
	notes: str = "",
) -> int:
	"""Fatura başlığını ve alt ürün kalemlerini kaydeder; aylık/toplam gider hesaplamalarını tutarlı kılar."""
	inv_dt = str(invoice_date or date.today().isoformat()).strip()[:10]
	ym = inv_dt[:7]
	clean_items: list[dict[str, Any]] = []
	subtotal = 0.0
	vat_tot = 0.0
	for it in (items or []):
		if not isinstance(it, dict):
			continue
		nm = str(it.get("item_name") or it.get("Ürün Adı") or "").strip()
		if not nm:
			continue
		qty = max(0.0, _safe_float(it.get("quantity", it.get("Miktar")), 1.0))
		unit = str(it.get("unit_type") or it.get("Birim") or it.get("Miktar Çeşidi") or "kg").strip() or "kg"
		u_price = max(0.0, _safe_float(it.get("unit_price", it.get("Birim Fiyat (₺)")), 0.0))
		v_rate = max(0.0, _safe_float(it.get("vat_rate", it.get("KDV (%)")), 0.0))
		base_line = qty * u_price
		explicit_line = _safe_float(it.get("line_total", it.get("Tutar (₺)")), -1.0)
		if base_line > 0:
			v_amt = base_line * (v_rate / 100.0)
			line_t = round(base_line + v_amt, 2)
			subtotal += base_line
			vat_tot += v_amt
		elif explicit_line >= 0:
			line_t = round(explicit_line, 2)
			base_est = line_t / (1.0 + v_rate / 100.0) if v_rate > 0 else line_t
			subtotal += base_est
			vat_tot += max(0.0, line_t - base_est)
		else:
			line_t = 0.0
		clean_items.append({
			"item_name": nm,
			"category": str(it.get("category") or it.get("Kategori") or category or "Diğer").strip(),
			"quantity": round(qty, 3),
			"unit_type": unit,
			"unit_price": round(u_price, 2),
			"vat_rate": round(v_rate, 1),
			"line_total": line_t,
			"expiry_date": str(it.get("expiry_date") or it.get("SKT") or "").strip()[:10],
			"notes": str(it.get("notes") or it.get("Not") or "").strip(),
		})

	tot_amt = round(subtotal + vat_tot, 2)
	subtotal = round(subtotal, 2)
	vat_tot = round(vat_tot, 2)
	now_iso = datetime.now().isoformat(timespec="minutes")

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		if purchase_id is not None and conn.execute("SELECT 1 FROM cafeteria_purchases WHERE id=?", (int(purchase_id),)).fetchone():
			pid = int(purchase_id)
			conn.execute(
				"""UPDATE cafeteria_purchases
				   SET invoice_date=?, year_month=?, invoice_no=?, supplier_name=?, category=?,
				       subtotal_amount=?, vat_amount=?, total_amount=?, payment_status=?, document_name=?, notes=?
				   WHERE id=?""",
				(inv_dt, ym, invoice_no.strip(), supplier_name.strip(), category, subtotal, vat_tot, tot_amt, payment_status, document_name, notes.strip(), pid),
			)
			conn.execute("DELETE FROM cafeteria_purchase_items WHERE purchase_id=?", (pid,))
		else:
			cur = conn.execute(
				"""INSERT INTO cafeteria_purchases(
					invoice_date, year_month, invoice_no, supplier_name, category,
					subtotal_amount, vat_amount, total_amount, payment_status, document_name, notes, created_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
				(inv_dt, ym, invoice_no.strip(), supplier_name.strip(), category, subtotal, vat_tot, tot_amt, payment_status, document_name, notes.strip(), now_iso),
			)
			pid = int(cur.lastrowid)

		for ci in clean_items:
			conn.execute(
				"""INSERT INTO cafeteria_purchase_items(
					purchase_id, item_name, category, quantity, unit_type, unit_price, vat_rate, line_total, expiry_date, notes
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
				(pid, ci["item_name"], ci["category"], ci["quantity"], ci["unit_type"], ci["unit_price"], ci["vat_rate"], ci["line_total"], ci["expiry_date"], ci["notes"]),
			)
	return pid


def get_purchases_and_items(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str | None = None,
	start_date: str | None = None,
	end_date: str | None = None,
	category: str | None = None,
	query_text: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
	"""Fatura tarihine göre sıralı faturalar ve alt kalem listesini döndürür."""
	where_bits = ["1=1"]
	params: list[Any] = []
	if year_month and year_month != "Tümü":
		where_bits.append("p.year_month = ?")
		params.append(str(year_month)[:7])
	if start_date:
		where_bits.append("p.invoice_date >= ?")
		params.append(str(start_date)[:10])
	if end_date:
		where_bits.append("p.invoice_date <= ?")
		params.append(str(end_date)[:10])
	if category and category != "Tümü":
		where_bits.append("p.category = ?")
		params.append(category)

	w_sql = " AND ".join(where_bits)
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		p_rows = conn.execute(
			f"SELECT * FROM cafeteria_purchases p WHERE {w_sql} ORDER BY p.invoice_date DESC, p.id DESC",
			params,
		).fetchall()
		i_rows = conn.execute(
			f"""SELECT pi.*, p.invoice_date, p.invoice_no, p.supplier_name
				FROM cafeteria_purchase_items pi
				JOIN cafeteria_purchases p ON p.id = pi.purchase_id
				WHERE {w_sql}
				ORDER BY p.invoice_date DESC, pi.id DESC""",
			params,
		).fetchall()

	purchases = [dict(r) for r in p_rows]
	items = [dict(r) for r in i_rows]
	all_items_by_pid: dict[int, list[dict[str, Any]]] = {}
	for it in items:
		all_items_by_pid.setdefault(int(it["purchase_id"]), []).append(it)

	qf = _tr_fold(query_text)
	if qf:
		items = [
			it for it in items
			if qf in _tr_fold(it.get("item_name")) or qf in _tr_fold(it.get("supplier_name")) or qf in _tr_fold(it.get("invoice_no"))
		]
		matching_pids = {it["purchase_id"] for it in items}
		purchases = [
			pr for pr in purchases
			if pr["id"] in matching_pids or qf in _tr_fold(pr.get("supplier_name")) or qf in _tr_fold(pr.get("invoice_no"))
		]

	for pr in purchases:
		pid_i = int(pr["id"])
		pr_items = sorted(all_items_by_pid.get(pid_i, []), key=lambda x: int(x.get("id") or 0))
		pr["items"] = pr_items
		pr["item_count"] = len(pr_items)
		parts_s = []
		for pit in pr_items[:4]:
			q_v = _safe_float(pit.get("quantity"), 0.0)
			u_v = str(pit.get("unit_type") or "")
			up_v = _safe_float(pit.get("unit_price"), 0.0)
			parts_s.append(f"{pit['item_name']} ({q_v:g} {u_v} × {up_v:g} ₺)")
		if len(pr_items) > 4:
			parts_s.append(f"+{len(pr_items) - 4} kalem")
		pr["items_summary"] = ", ".join(parts_s) if parts_s else "—"
	return purchases, items


def get_expense_summary_stats(connect_db_fn: Callable[[], sqlite3.Connection], year_month: str) -> dict[str, Any]:
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy = ym[:4]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		m_row = conn.execute(
			"""SELECT COALESCE(SUM(total_amount),0) AS tot,
			          COUNT(*) AS cnt,
			          COALESCE(SUM(CASE WHEN payment_status='Vadeli / Bekliyor' THEN total_amount ELSE 0 END), 0) AS pending_tot,
			          COALESCE(SUM(CASE WHEN payment_status='Vadeli / Bekliyor' THEN 1 ELSE 0 END), 0) AS pending_cnt
			   FROM cafeteria_purchases WHERE year_month=?""",
			(ym,),
		).fetchone()
		y_row = conn.execute(
			"SELECT COALESCE(SUM(total_amount),0) AS tot FROM cafeteria_purchases WHERE substr(year_month,1,4)=?",
			(yy,),
		).fetchone()
		all_row = conn.execute(
			"""SELECT COALESCE(SUM(total_amount),0) AS tot,
			          COALESCE(SUM(CASE WHEN payment_status='Vadeli / Bekliyor' THEN total_amount ELSE 0 END), 0) AS all_pending
			   FROM cafeteria_purchases"""
		).fetchone()
		cat_rows = conn.execute(
			"""SELECT pi.category AS cat, COALESCE(SUM(pi.line_total),0) AS tot
			   FROM cafeteria_purchase_items pi
			   JOIN cafeteria_purchases p ON p.id = pi.purchase_id
			   WHERE p.year_month=?
			   GROUP BY pi.category
			   ORDER BY tot DESC""",
			(ym,),
		).fetchall()
	m_tot = round(float(m_row["tot"] or 0.0), 2)
	m_pend = round(float(m_row["pending_tot"] or 0.0), 2)
	return {
		"month_total": m_tot,
		"month_paid_total": round(max(0.0, m_tot - m_pend), 2),
		"month_pending_total": m_pend,
		"month_pending_count": int(m_row["pending_cnt"] or 0),
		"month_invoice_count": int(m_row["cnt"] or 0),
		"year_total": round(float(y_row["tot"] or 0.0), 2),
		"all_time_total": round(float(all_row["tot"] or 0.0), 2),
		"all_pending_total": round(float(all_row["all_pending"] or 0.0), 2),
		"by_category": {str(r["cat"]): round(float(r["tot"] or 0.0), 2) for r in cat_rows},
	}


def get_purchase_by_id(connect_db_fn: Callable[[], sqlite3.Connection], purchase_id: int) -> dict[str, Any] | None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		p = conn.execute("SELECT * FROM cafeteria_purchases WHERE id=?", (int(purchase_id),)).fetchone()
		if not p:
			return None
		d = dict(p)
		d["items"] = [
			dict(r)
			for r in conn.execute("SELECT * FROM cafeteria_purchase_items WHERE purchase_id=? ORDER BY id ASC", (int(purchase_id),)).fetchall()
		]
		return d


def update_purchase_payment_status(
	connect_db_fn: Callable[[], sqlite3.Connection],
	purchase_id: int,
	payment_status: str = "Ödendi",
) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute(
			"UPDATE cafeteria_purchases SET payment_status=? WHERE id=?",
			(str(payment_status or "Ödendi"), int(purchase_id)),
		)


def get_expiring_purchase_items(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	within_days: int = 30,
	include_expired: bool = True,
) -> list[dict[str, Any]]:
	"""Fatura kalemlerinde SKT (expiry_date) yaklaşan veya süresi geçmiş mutfak ürünlerini listeler."""
	today_d = date.today()
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute(
			"""SELECT pi.id, pi.item_name, pi.category, pi.quantity, pi.unit_type, pi.expiry_date,
			          p.invoice_date, p.supplier_name
			   FROM cafeteria_purchase_items pi
			   JOIN cafeteria_purchases p ON p.id = pi.purchase_id
			   WHERE TRIM(COALESCE(pi.expiry_date, '')) != ''
			   ORDER BY pi.expiry_date ASC"""
		).fetchall()
	out: list[dict[str, Any]] = []
	for r in rows:
		exp_s = str(r["expiry_date"] or "")[:10]
		try:
			exp_d = date.fromisoformat(exp_s)
		except Exception:
			continue
		days_left = (exp_d - today_d).days
		if days_left < 0 and not include_expired:
			continue
		if days_left <= within_days:
			d = dict(r)
			d["days_left"] = days_left
			d["urgency"] = "🔴 Tarihi Geçmiş" if days_left < 0 else ("🟠 Kritik (<=7 gün)" if days_left <= 7 else "🟡 Yaklaşıyor")
			out.append(d)
	return out


def _normalize_to_purchase_unit(amount: float, ing_unit: str, purch_unit: str) -> float | None:
	"""Malzeme miktarını faturadaki birim cinsine ölçekler; uyumsuz birimse None döndürür."""
	iu = _tr_fold(ing_unit)
	pu = _tr_fold(purch_unit)
	if iu == pu:
		return amount
	weight_to_kg = {"g": 0.001, "gr": 0.001, "kg": 1.0, "yemek kasigi": 0.015, "cay kasigi": 0.005, "su bardagi": 0.2}
	vol_to_l = {"ml": 0.001, "l": 1.0, "lt": 1.0, "litre": 1.0, "su bardagi": 0.2, "yemek kasigi": 0.015}
	if pu == "kg" and iu in weight_to_kg:
		return amount * weight_to_kg[iu]
	if pu == "g" and iu in weight_to_kg:
		return amount * weight_to_kg[iu] * 1000.0
	if pu in ("litre", "lt", "l") and iu in vol_to_l:
		return amount * vol_to_l[iu]
	if pu == "ml" and iu in vol_to_l:
		return amount * vol_to_l[iu] * 1000.0
	return None


def _get_latest_purchase_prices_map(connect_db_fn: Callable[[], sqlite3.Connection]) -> dict[str, dict[str, Any]]:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute(
			"""SELECT pi.item_name, pi.unit_type, pi.unit_price, pi.vat_rate, p.invoice_date
			   FROM cafeteria_purchase_items pi
			   JOIN cafeteria_purchases p ON p.id = pi.purchase_id
			   WHERE pi.unit_price > 0
			   ORDER BY p.invoice_date DESC, pi.id DESC"""
		).fetchall()
	out: dict[str, dict[str, Any]] = {}
	for r in rows:
		k = _tr_fold(r["item_name"])
		if k and k not in out:
			d = dict(r)
			d["eff_unit_price"] = round(float(d["unit_price"]) * (1.0 + float(d.get("vat_rate") or 0.0) / 100.0), 4)
			out[k] = d
	return out


def get_ingredient_price_inflation(connect_db_fn: Callable[[], sqlite3.Connection]) -> list[dict[str, Any]]:
	"""Aynı ürünün son fatura fiyatı ile bir önceki faturadaki fiyatını karşılaştırır."""
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute(
			"""SELECT pi.item_name, pi.unit_type, pi.unit_price, p.invoice_date, p.supplier_name
			   FROM cafeteria_purchase_items pi
			   JOIN cafeteria_purchases p ON p.id = pi.purchase_id
			   WHERE pi.unit_price > 0
			   ORDER BY p.invoice_date DESC, pi.id DESC"""
		).fetchall()
	grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
	for r in rows:
		key = (_tr_fold(r["item_name"]), _tr_fold(r["unit_type"]))
		grouped.setdefault(key, []).append(dict(r))

	changes: list[dict[str, Any]] = []
	for (_, _), lst in grouped.items():
		if len(lst) < 2:
			continue
		latest, prev = lst[0], lst[1]
		p_new, p_old = float(latest["unit_price"]), float(prev["unit_price"])
		if p_old <= 0 or abs(p_new - p_old) < 0.01:
			continue
		pct = round(((p_new - p_old) / p_old) * 100.0, 1)
		changes.append({
			"item_name": latest["item_name"],
			"unit_type": latest["unit_type"],
			"prev_price": round(p_old, 2),
			"prev_date": prev["invoice_date"],
			"latest_price": round(p_new, 2),
			"latest_date": latest["invoice_date"],
			"pct_change": pct,
		})
	changes.sort(key=lambda x: abs(x["pct_change"]), reverse=True)
	return changes


def estimate_dish_portion_cost(
	connect_db_fn: Callable[[], sqlite3.Connection],
	dish_id: int,
) -> dict[str, Any]:
	"""Son ürün faturalarındaki birim fiyatlara göre 1 öğrenci porsiyonunun tahmini maliyetini (₺) hesaplar."""
	prices = _get_latest_purchase_prices_map(connect_db_fn)
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		ings = conn.execute(
			"SELECT * FROM cafeteria_dish_ingredients WHERE dish_id=? ORDER BY sort_order ASC",
			(int(dish_id),),
		).fetchall()
	total_cost = 0.0
	matched = 0
	unmatched: list[str] = []
	for ing in ings:
		nm = str(ing["ingredient_name"] or "")
		nf = _tr_fold(nm)
		p_rec = prices.get(nf)
		if not p_rec:
			for pk, pv in prices.items():
				if (nf in pk or pk in nf) and len(nf) >= 3:
					p_rec = pv
					break
		if not p_rec:
			unmatched.append(nm)
			continue
		conv = _normalize_to_purchase_unit(float(ing["amount"] or 0.0), str(ing["unit"] or "g"), str(p_rec["unit_type"]))
		if conv is None:
			unmatched.append(f"{nm} ({ing['unit']}/{p_rec['unit_type']})")
			continue
		total_cost += conv * float(p_rec["eff_unit_price"])
		matched += 1
	return {
		"estimated_cost_tl": round(total_cost, 2),
		"matched_count": matched,
		"total_ingredients": len(ings),
		"unmatched": unmatched,
	}


def build_menu_shopping_list(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
	*,
	scope: str = "monthly",
	week_index: int = 0,
	diner_count: int | None = None,
) -> dict[str, Any]:
	"""Planlı menü × kişi sayısı formülüyle toplam malzeme ihtiyaç / satın alma sipariş listesi oluşturur."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	days_info = get_month_workdays_info(connect_db_fn, ym)
	plans = get_menu_plans_for_month(connect_db_fn, ym)
	dishes_map = {d["id"]: d for d in get_all_dishes(connect_db_fn)}
	prices = _get_latest_purchase_prices_map(connect_db_fn)

	if diner_count is None:
		d_cnt = max(1, len(get_diners(connect_db_fn, active_only=True)))
	else:
		d_cnt = max(1, int(diner_count))

	if scope == "weekly":
		wks = _group_month_days_by_week(days_info)
		safe_i = max(0, min(week_index, len(wks) - 1)) if wks else 0
		target_dates = {d["date_iso"] for d in wks[safe_i]} if wks else set()
	else:
		target_dates = {d["date_iso"] for d in days_info if d.get("is_workday")}

	agg: dict[tuple[str, str], dict[str, Any]] = {}
	for (p_dt, _), rec in plans.items():
		if p_dt not in target_dates:
			continue
		for did in (rec.get("dish_ids") or []):
			dish = dishes_map.get(int(did))
			if not dish:
				continue
			for ing in (dish.get("ingredients") or []):
				nm = str(ing.get("ingredient_name") or "").strip()
				u = str(ing.get("unit") or "g").strip()
				amt_tot = _safe_float(ing.get("amount"), 0.0) * d_cnt
				if not nm or amt_tot <= 0:
					continue
				target_u = "kg" if _tr_fold(u) in {"g", "gr", "kg"} else ("litre" if _tr_fold(u) in {"ml", "l", "lt", "litre"} else u)
				norm_amt = _normalize_to_purchase_unit(amt_tot, u, target_u)
				if norm_amt is None:
					norm_amt = amt_tot
					target_u = u
				k = (_tr_fold(nm), _tr_fold(target_u))
				if k not in agg:
					agg[k] = {"Malzeme / Ürün": nm, "Toplam Miktar": 0.0, "Birim": target_u}
				agg[k]["Toplam Miktar"] += norm_amt

	rows_out: list[dict[str, Any]] = []
	est_total_tl = 0.0
	for (_, _), item in sorted(agg.items(), key=lambda x: x[1]["Malzeme / Ürün"]):
		qty = round(float(item["Toplam Miktar"]), 2)
		nf = _tr_fold(item["Malzeme / Ürün"])
		p_rec = prices.get(nf)
		if not p_rec:
			for pk, pv in prices.items():
				if (nf in pk or pk in nf) and len(nf) >= 3:
					p_rec = pv
					break
		unit_p = 0.0
		line_est = 0.0
		if p_rec:
			conv_p = _normalize_to_purchase_unit(qty, item["Birim"], str(p_rec["unit_type"]))
			if conv_p is not None:
				unit_p = round(float(p_rec["eff_unit_price"]), 2)
				line_est = round(conv_p * float(p_rec["eff_unit_price"]), 2)
				est_total_tl += line_est
		rows_out.append({
			"Malzeme / Ürün": item["Malzeme / Ürün"],
			"Toplam İhtiyaç": qty,
			"Birim": item["Birim"],
			"Son Fatura Birim Fiyatı (₺)": unit_p if unit_p > 0 else None,
			"Tahmini Maliyet (₺)": line_est if line_est > 0 else None,
		})
	return {
		"diner_count": d_cnt,
		"items": rows_out,
		"estimated_total_tl": round(est_total_tl, 2),
	}


def build_shopping_list_excel(shopping_data: dict[str, Any], year_month: str) -> bytes:
	df = pd.DataFrame(shopping_data.get("items") or [{"Bilgi": "Seçili dönemde malzemeli yemek planı bulunamadı"}])
	buf = io.BytesIO()
	with pd.ExcelWriter(buf, engine="openpyxl") as writer:
		df.to_excel(writer, sheet_name=f"Ihtiyac_{year_month}"[:31], index=False)
	return buf.getvalue()


def delete_purchase_by_id(connect_db_fn: Callable[[], sqlite3.Connection], purchase_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute("DELETE FROM cafeteria_purchases WHERE id=?", (int(purchase_id),))


PURCHASE_STAT_VIEWS: list[str] = [
	"📦 Ürün Bazlı Harcama, Miktar & Fiyat Özeti",
	"🗂️ Kategori Bazlı Gider Dağılımı",
	"🏪 Tedarikçi / Firma Karşılaştırması",
	"📅 Aylık Harcama & Gider Trendi",
	"📈 Ürün Birim Fiyat Değişimi (Enflasyon)",
	"🔎 Seçili Ürün Alım & Fiyat Geçmişi",
]


def compute_purchase_analytics(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str | None = "Tümü",
	start_date: str | None = None,
	end_date: str | None = None,
	category: str | None = "Tümü",
	query_text: str = "",
	product_name: str = "",
) -> dict[str, Any]:
	"""Fatura ve kalem verilerinden seçilebilir ürün, kategori, tedarikçi, aylık trend ve birim fiyat istatistiklerini hesaplar."""
	purchases, items = get_purchases_and_items(
		connect_db_fn,
		year_month=year_month,
		start_date=start_date,
		end_date=end_date,
		category=category,
		query_text=query_text,
	)

	total_spend = round(sum(_safe_float(pr.get("total_amount"), 0.0) for pr in purchases), 2)
	subtotal_tl = round(sum(_safe_float(pr.get("subtotal_amount"), 0.0) for pr in purchases), 2)
	vat_tl = round(sum(_safe_float(pr.get("vat_amount"), 0.0) for pr in purchases), 2)
	pending_tl = round(sum(_safe_float(pr.get("total_amount"), 0.0) for pr in purchases if pr.get("payment_status") == "Vadeli / Bekliyor"), 2)
	paid_tl = round(max(0.0, total_spend - pending_tl), 2)
	inv_cnt = len(purchases)
	avg_inv_tl = round(total_spend / inv_cnt, 2) if inv_cnt > 0 else 0.0

	# 1. Ürün Bazlı Gruplama
	prod_groups: dict[tuple[str, str], dict[str, Any]] = {}
	cat_groups: dict[str, dict[str, Any]] = {}
	sup_groups: dict[str, dict[str, Any]] = {}
	month_groups: dict[str, dict[str, Any]] = {}
	inf_groups: dict[tuple[str, str], list[dict[str, Any]]] = {}

	total_items_spend = 0.0
	for it in items:
		nm = str(it.get("item_name") or "").strip()
		if not nm:
			continue
		u_tp = str(it.get("unit_type") or "kg").strip() or "kg"
		cat_s = str(it.get("category") or "Diğer").strip() or "Diğer"
		qty = _safe_float(it.get("quantity"), 0.0)
		u_pr = _safe_float(it.get("unit_price"), 0.0)
		l_tot = _safe_float(it.get("line_total"), 0.0)
		pid = int(it.get("purchase_id") or 0)
		inv_dt = str(it.get("invoice_date") or "")[:10]
		total_items_spend += l_tot

		pk = (_tr_fold(nm), _tr_fold(u_tp))
		if pk not in prod_groups:
			prod_groups[pk] = {
				"item_name": nm,
				"category": cat_s,
				"unit_type": u_tp,
				"pids": set(),
				"total_qty": 0.0,
				"total_spend": 0.0,
				"weighted_num": 0.0,
				"unit_prices": [],
				"latest_date": inv_dt,
				"latest_price": u_pr,
			}
		pg = prod_groups[pk]
		pg["pids"].add(pid)
		pg["total_qty"] += qty
		pg["total_spend"] += l_tot
		if u_pr > 0:
			pg["weighted_num"] += qty * u_pr
			pg["unit_prices"].append(u_pr)
			inf_groups.setdefault(pk, []).append(it)
		if (inv_dt >= pg["latest_date"]) and u_pr > 0:
			pg["latest_date"] = inv_dt
			pg["latest_price"] = u_pr

		if cat_s not in cat_groups:
			cat_groups[cat_s] = {"count": 0, "products": set(), "spend": 0.0}
		cat_groups[cat_s]["count"] += 1
		cat_groups[cat_s]["products"].add(_tr_fold(nm))
		cat_groups[cat_s]["spend"] += l_tot

	product_summary: list[dict[str, Any]] = []
	for pg in prod_groups.values():
		up_list = pg["unit_prices"]
		tq = pg["total_qty"]
		avg_p = round(pg["weighted_num"] / tq, 2) if (tq > 0 and pg["weighted_num"] > 0) else (round(sum(up_list) / len(up_list), 2) if up_list else 0.0)
		ts = round(pg["total_spend"], 2)
		product_summary.append({
			"Ürün Adı": pg["item_name"],
			"Kategori": pg["category"],
			"Birim": pg["unit_type"],
			"Alım Sayısı": len(pg["pids"]),
			"Toplam Miktar": round(tq, 2),
			"Ort. Birim Fiyat (₺)": avg_p,
			"Min Fiyat (₺)": round(min(up_list), 2) if up_list else 0.0,
			"Max Fiyat (₺)": round(max(up_list), 2) if up_list else 0.0,
			"Son Birim Fiyat (₺)": round(pg["latest_price"], 2),
			"Toplam Harcama (₺)": ts,
			"Pay (%)": round((ts / total_items_spend) * 100.0, 1) if total_items_spend > 0 else 0.0,
		})
	product_summary.sort(key=lambda x: (-x["Toplam Harcama (₺)"], -x["Alım Sayısı"], x["Ürün Adı"]))

	category_summary = [
		{
			"Kategori": c_nm,
			"Kalem Sayısı": cd["count"],
			"Farklı Ürün": len(cd["products"]),
			"Toplam Harcama (₺)": round(cd["spend"], 2),
			"Pay (%)": round((cd["spend"] / total_items_spend) * 100.0, 1) if total_items_spend > 0 else 0.0,
		}
		for c_nm, cd in cat_groups.items()
	]
	category_summary.sort(key=lambda x: (-x["Toplam Harcama (₺)"], x["Kategori"]))

	for pr in purchases:
		sup_nm = str(pr.get("supplier_name") or "").strip() or "Tedarikçi Belirtilmemiş"
		sup_k = _tr_fold(sup_nm)
		if sup_k not in sup_groups:
			sup_groups[sup_k] = {"name": sup_nm, "inv_cnt": 0, "item_cnt": 0, "paid": 0.0, "pend": 0.0, "tot": 0.0}
		sg = sup_groups[sup_k]
		amt = _safe_float(pr.get("total_amount"), 0.0)
		sg["inv_cnt"] += 1
		sg["item_cnt"] += int(pr.get("item_count") or len(pr.get("items") or []))
		sg["tot"] += amt
		if pr.get("payment_status") == "Vadeli / Bekliyor":
			sg["pend"] += amt
		else:
			sg["paid"] += amt

		ym_k = str(pr.get("year_month") or str(pr.get("invoice_date") or "")[:7])[:7]
		if ym_k:
			if ym_k not in month_groups:
				month_groups[ym_k] = {"inv_cnt": 0, "item_cnt": 0, "paid": 0.0, "pend": 0.0, "tot": 0.0}
			mg = month_groups[ym_k]
			mg["inv_cnt"] += 1
			mg["item_cnt"] += int(pr.get("item_count") or len(pr.get("items") or []))
			mg["tot"] += amt
			if pr.get("payment_status") == "Vadeli / Bekliyor":
				mg["pend"] += amt
			else:
				mg["paid"] += amt

	supplier_summary = [
		{
			"Tedarikçi / Firma": sg["name"],
			"Fatura Sayısı": sg["inv_cnt"],
			"Kalem Sayısı": sg["item_cnt"],
			"Ödenen (₺)": round(sg["paid"], 2),
			"Vadeli (₺)": round(sg["pend"], 2),
			"Toplam Harcama (₺)": round(sg["tot"], 2),
			"Ort. Fatura (₺)": round(sg["tot"] / sg["inv_cnt"], 2) if sg["inv_cnt"] > 0 else 0.0,
		}
		for sg in sup_groups.values()
	]
	supplier_summary.sort(key=lambda x: (-x["Toplam Harcama (₺)"], x["Tedarikçi / Firma"]))

	monthly_trend = [
		{
			"Ay": ym_k,
			"Dönem": _fmt_ym(ym_k),
			"Fatura Sayısı": mg["inv_cnt"],
			"Kalem Sayısı": mg["item_cnt"],
			"Ödenen (₺)": round(mg["paid"], 2),
			"Vadeli (₺)": round(mg["pend"], 2),
			"Toplam Harcama (₺)": round(mg["tot"], 2),
		}
		for ym_k, mg in sorted(month_groups.items())
	]

	price_changes: list[dict[str, Any]] = []
	for (_, _), lst in inf_groups.items():
		if len(lst) < 2:
			continue
		lst_ord = sorted(lst, key=lambda x: (str(x.get("invoice_date") or ""), int(x.get("id") or 0)), reverse=True)
		latest, prev = lst_ord[0], lst_ord[1]
		p_new, p_old = _safe_float(latest.get("unit_price"), 0.0), _safe_float(prev.get("unit_price"), 0.0)
		if p_old <= 0 or abs(p_new - p_old) < 0.01:
			continue
		pct = round(((p_new - p_old) / p_old) * 100.0, 1)
		price_changes.append({
			"Ürün": latest["item_name"],
			"Birim": latest["unit_type"],
			"Önceki Fiyat (₺)": round(p_old, 2),
			"Önceki Tarih": prev.get("invoice_date") or "",
			"Son Fiyat (₺)": round(p_new, 2),
			"Son Tarih": latest.get("invoice_date") or "",
			"Değişim (%)": pct,
		})
	price_changes.sort(key=lambda x: abs(x["Değişim (%)"]), reverse=True)

	product_options = sorted({str(it.get("item_name") or "").strip() for it in items if str(it.get("item_name") or "").strip()})
	chosen_prod = str(product_name or "").strip() or (product_summary[0]["Ürün Adı"] if product_summary else "")
	chosen_fold = _tr_fold(chosen_prod)
	product_history = [
		{
			"Tarih": it.get("invoice_date") or "",
			"Fatura No": it.get("invoice_no") or "—",
			"Tedarikçi": it.get("supplier_name") or "—",
			"Miktar": it.get("quantity") or 0.0,
			"Birim": it.get("unit_type") or "",
			"Birim Fiyat (₺)": it.get("unit_price") or 0.0,
			"KDV (%)": it.get("vat_rate") or 0.0,
			"Toplam (₺)": it.get("line_total") or 0.0,
		}
		for it in sorted(items, key=lambda x: (str(x.get("invoice_date") or ""), int(x.get("id") or 0)))
		if chosen_fold and _tr_fold(it.get("item_name")) == chosen_fold
	]

	return {
		"summary": {
			"invoice_count": inv_cnt,
			"item_row_count": len(items),
			"distinct_products": len({r["Ürün Adı"] for r in product_summary}),
			"distinct_suppliers": len(supplier_summary),
			"subtotal_tl": subtotal_tl,
			"vat_tl": vat_tl,
			"total_spend_tl": total_spend,
			"paid_total_tl": paid_tl,
			"pending_total_tl": pending_tl,
			"avg_invoice_tl": avg_inv_tl,
			"top_product_name": product_summary[0]["Ürün Adı"] if product_summary else "—",
			"top_product_spend": product_summary[0]["Toplam Harcama (₺)"] if product_summary else 0.0,
			"top_supplier_name": supplier_summary[0]["Tedarikçi / Firma"] if supplier_summary else "—",
			"top_supplier_spend": supplier_summary[0]["Toplam Harcama (₺)"] if supplier_summary else 0.0,
		},
		"product_summary": product_summary,
		"category_summary": category_summary,
		"supplier_summary": supplier_summary,
		"monthly_trend": monthly_trend,
		"price_changes": price_changes,
		"product_options": product_options,
		"selected_product": chosen_prod,
		"product_history": product_history,
	}


def build_purchase_analytics_excel(analytics: dict[str, Any], period_label: str = "") -> bytes:
	"""Seçilen dönemdeki fatura ve ürün istatistiklerini çok sekmeli Excel (.xlsx) dosyası olarak üretir."""
	sm = analytics.get("summary") or {}
	df_kpi = pd.DataFrame([
		{"Gösterge": "Analiz Dönemi", "Değer": period_label or "Tüm Tarihler"},
		{"Gösterge": "Toplam Fatura Sayısı", "Değer": sm.get("invoice_count", 0)},
		{"Gösterge": "Toplam Ürün Kalemi Satırı", "Değer": sm.get("item_row_count", 0)},
		{"Gösterge": "Farklı Ürün Çeşidi", "Değer": sm.get("distinct_products", 0)},
		{"Gösterge": "Tedarikçi Sayısı", "Değer": sm.get("distinct_suppliers", 0)},
		{"Gösterge": "Toplam Harcama (₺)", "Değer": sm.get("total_spend_tl", 0.0)},
		{"Gösterge": "Ortalama Fatura Tutarı (₺)", "Değer": sm.get("avg_invoice_tl", 0.0)},
		{"Gösterge": "En Çok Harcanan Ürün", "Değer": f"{sm.get('top_product_name', '—')} ({sm.get('top_product_spend', 0.0):,.2f} ₺)"},
	])
	df_prod = pd.DataFrame(analytics.get("product_summary") or [{"Bilgi": "Ürün verisi yok"}])
	df_cat = pd.DataFrame(analytics.get("category_summary") or [{"Bilgi": "Kategori verisi yok"}])
	df_sup = pd.DataFrame(analytics.get("supplier_summary") or [{"Bilgi": "Tedarikçi verisi yok"}])
	df_mon = pd.DataFrame(analytics.get("monthly_trend") or [{"Bilgi": "Aylık trend verisi yok"}])
	df_inf = pd.DataFrame(analytics.get("price_changes") or [{"Bilgi": "Fiyat değişimi verisi yok"}])

	buf = io.BytesIO()
	with pd.ExcelWriter(buf, engine="openpyxl") as writer:
		df_kpi.to_excel(writer, sheet_name="Ozet_KPI", index=False)
		df_prod.to_excel(writer, sheet_name="Urun_Bazli", index=False)
		df_cat.to_excel(writer, sheet_name="Kategori_Bazli", index=False)
		df_sup.to_excel(writer, sheet_name="Tedarikci_Analizi", index=False)
		df_mon.to_excel(writer, sheet_name="Aylik_Trend", index=False)
		df_inf.to_excel(writer, sheet_name="Fiyat_Degisimi", index=False)
	return buf.getvalue()


# ==============================================================================
# 6. BÖLÜM 4: YEMEK YİYENLER, ÖĞÜN ÜCRETİ TAHAKKUKU & KASA MUHASEBESİ
# ==============================================================================

def upsert_diner(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	diner_id: int | None = None,
	student_id: int | None = None,
	full_name: str,
	role_type: str = "Öğrenci",
	class_or_dept: str = "",
	daily_meals_count: int = 1,
	fee_exempt: bool = False,
	allergies_override: str = "",
	is_active: bool = True,
	notes: str = "",
) -> int:
	fn = re.sub(r"\s+", " ", str(full_name or "").strip())
	if not fn:
		raise ValueError("Kişi adı boş bırakılamaz.")
	ex_int = 1 if (fee_exempt or role_type == "Burslu / Ücretsiz") else 0
	act_int = 1 if is_active else 0
	meals_n = max(1, min(3, int(daily_meals_count or 1)))
	sid_val = int(student_id) if student_id else None

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		existing = None
		if diner_id is not None:
			existing = conn.execute("SELECT id FROM cafeteria_diners WHERE id=?", (int(diner_id),)).fetchone()
		elif sid_val is not None:
			existing = conn.execute("SELECT id FROM cafeteria_diners WHERE student_id=?", (sid_val,)).fetchone()
		if existing:
			did = int(existing["id"])
			conn.execute(
				"""UPDATE cafeteria_diners
				   SET student_id=?, full_name=?, role_type=?, class_or_dept=?, daily_meals_count=?,
				       fee_exempt=?, allergies_override=?, is_active=?, notes=?
				   WHERE id=?""",
				(sid_val, fn, role_type, class_or_dept, meals_n, ex_int, allergies_override.strip(), act_int, notes.strip(), did),
			)
			return did
		cur = conn.execute(
			"""INSERT INTO cafeteria_diners(
				student_id, full_name, role_type, class_or_dept, daily_meals_count,
				fee_exempt, allergies_override, is_active, start_date, notes
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(sid_val, fn, role_type, class_or_dept, meals_n, ex_int, allergies_override.strip(), act_int, date.today().isoformat(), notes.strip()),
		)
		return int(cur.lastrowid)


def import_school_people_as_diners(
	connect_db_fn: Callable[[], sqlite3.Connection],
	person_ids: list[int],
	*,
	daily_meals_count: int = 1,
) -> int:
	"""Revir veritabanındaki öğrenci/personel kayıtlarını yemek yiyenler listesine bağlar."""
	added = 0
	if not person_ids:
		return 0
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		ph = ",".join("?" for _ in person_ids)
		rows = conn.execute(
			f"""SELECT s.id, s.full_name, s.person_type, COALESCE(c.name, '—') AS class_name,
					   COALESCE(hp.allergies, '') AS allergies
				FROM students s
				LEFT JOIN classes c ON c.id = s.class_id
				LEFT JOIN health_profiles hp ON hp.student_id = s.id
				WHERE s.id IN ({ph})""",
			[int(x) for x in person_ids],
		).fetchall()
		for r in rows:
			exists = conn.execute("SELECT id FROM cafeteria_diners WHERE student_id=?", (int(r["id"]),)).fetchone()
			if exists:
				conn.execute("UPDATE cafeteria_diners SET is_active=1 WHERE id=?", (int(exists["id"]),))
				continue
			ptype = str(r["person_type"] or "Öğrenci")
			role = "Öğrenci" if ptype == "Öğrenci" else "Personel"
			conn.execute(
				"""INSERT INTO cafeteria_diners(
					student_id, full_name, role_type, class_or_dept, daily_meals_count,
					fee_exempt, allergies_override, is_active, start_date, notes
				) VALUES (?, ?, ?, ?, ?, 0, ?, 1, ?, '')""",
				(int(r["id"]), r["full_name"], role, r["class_name"], max(1, int(daily_meals_count)), r["allergies"], date.today().isoformat()),
			)
			added += 1
	return added


def get_diners(connect_db_fn: Callable[[], sqlite3.Connection], active_only: bool = True) -> list[dict[str, Any]]:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		sql = """SELECT d.*, COALESCE(hp.allergies, '') AS revir_allergies
				 FROM cafeteria_diners d
				 LEFT JOIN health_profiles hp ON hp.student_id = d.student_id"""
		if active_only:
			sql += " WHERE d.is_active=1"
		sql += " ORDER BY d.class_or_dept ASC, d.full_name ASC"
		return [dict(r) for r in conn.execute(sql).fetchall()]


def deactivate_diner_by_id(connect_db_fn: Callable[[], sqlite3.Connection], diner_id: int, end_date: str = "") -> None:
	"""Aboneyi geçmiş tahakkuk ve kasa kayıtlarını silmeden pasife alır."""
	e_dt = str(end_date or date.today().isoformat()).strip()[:10]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute(
			"UPDATE cafeteria_diners SET is_active=0, end_date=? WHERE id=?",
			(e_dt, int(diner_id)),
		)


def reactivate_diner_by_id(connect_db_fn: Callable[[], sqlite3.Connection], diner_id: int) -> None:
	"""Pasif yemekhane abonesini yeniden aktif listeye alır."""
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute(
			"UPDATE cafeteria_diners SET is_active=1, end_date='' WHERE id=?",
			(int(diner_id),),
		)


def delete_diner_by_id(connect_db_fn: Callable[[], sqlite3.Connection], diner_id: int, *, force: bool = False) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		if not force:
			has_fees = conn.execute("SELECT 1 FROM cafeteria_monthly_fees WHERE diner_id=? LIMIT 1", (int(diner_id),)).fetchone()
			has_cash = conn.execute("SELECT 1 FROM cafeteria_cash_ledger WHERE diner_id=? LIMIT 1", (int(diner_id),)).fetchone()
			if has_fees or has_cash:
				raise ValueError("Bu abonenin geçmiş tahakkuk veya tahsilat kayıtları bulunuyor; geçmişi korumak için 'Pasife Al' kullanın.")
		conn.execute("DELETE FROM cafeteria_diners WHERE id=?", (int(diner_id),))


def generate_or_refresh_monthly_fees(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
	*,
	student_meal_price: float,
	staff_meal_price: float | None = None,
	workdays_override: int | None = None,
	reset_planned_days: bool = False,
) -> int:
	"""
	Seçili ayın iş günü sayısı × günlük öğün × öğün ücreti + devreden bakiye - indirim formülüyle
	aktif yemek yiyenlerin aylık ücret tahakkuklarını hesaplar ve günceller.
	"""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	if workdays_override is not None:
		workdays = max(0, int(workdays_override))
	else:
		days_info = get_month_workdays_info(connect_db_fn, ym)
		workdays = sum(1 for d in days_info if d.get("is_workday"))

	stu_p = max(0.0, float(student_meal_price or 0.0))
	stf_p = max(0.0, float(staff_meal_price)) if staff_meal_price is not None else stu_p
	diners = get_diners(connect_db_fn, active_only=True)
	now_iso = datetime.now().isoformat(timespec="minutes")
	count = 0

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		for dn in diners:
			did = int(dn["id"])
			is_exempt = bool(dn.get("fee_exempt")) or dn.get("role_type") == "Burslu / Ücretsiz"
			unit_p = 0.0 if is_exempt else (stu_p if dn.get("role_type") in ("Öğrenci", "Misafir") else stf_p)
			cur_row = conn.execute(
				"SELECT * FROM cafeteria_monthly_fees WHERE year_month=? AND diner_id=?",
				(ym, did),
			).fetchone()
			meals_pd = int(cur_row["meals_per_day"]) if (cur_row and not reset_planned_days) else int(dn.get("daily_meals_count") or 1)
			p_days = workdays if (cur_row is None or reset_planned_days) else int(cur_row["planned_days"])
			disc = float(cur_row["discount_or_adjustment"]) if cur_row else 0.0
			carry = float(cur_row["carry_forward_debt"] or 0.0) if cur_row and "carry_forward_debt" in cur_row.keys() else 0.0
			exp_amt = 0.0 if is_exempt else max(0.0, round(p_days * meals_pd * unit_p + carry - disc, 2))
			conn.execute(
				"""INSERT INTO cafeteria_monthly_fees(
					year_month, diner_id, workdays_in_month, planned_days, meals_per_day,
					unit_meal_price, discount_or_adjustment, carry_forward_debt, expected_amount, note, updated_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				ON CONFLICT(year_month, diner_id) DO UPDATE SET
					workdays_in_month=excluded.workdays_in_month,
					planned_days=excluded.planned_days,
					meals_per_day=excluded.meals_per_day,
					unit_meal_price=excluded.unit_meal_price,
					discount_or_adjustment=excluded.discount_or_adjustment,
					carry_forward_debt=excluded.carry_forward_debt,
					expected_amount=excluded.expected_amount,
					updated_at=excluded.updated_at""",
				(ym, did, workdays, p_days, meals_pd, unit_p, disc, carry, exp_amt, (cur_row["note"] if cur_row else ""), now_iso),
			)
			count += 1
	return count


def update_monthly_fee_row(
	connect_db_fn: Callable[[], sqlite3.Connection],
	fee_id: int,
	*,
	planned_days: int,
	meals_per_day: int,
	unit_meal_price: float,
	discount_or_adjustment: float = 0.0,
	carry_forward_debt: float | None = None,
	note: str = "",
) -> None:
	p_days = max(0, int(planned_days))
	mpd = max(1, min(3, int(meals_per_day)))
	up = max(0.0, float(unit_meal_price))
	disc = float(discount_or_adjustment or 0.0)
	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		if carry_forward_debt is None:
			r_cf = conn.execute("SELECT carry_forward_debt, note FROM cafeteria_monthly_fees WHERE id=?", (int(fee_id),)).fetchone()
			cf_val = float(r_cf["carry_forward_debt"] or 0.0) if r_cf else 0.0
			cur_note = note.strip() or (str(r_cf["note"] or "") if r_cf else "")
		else:
			cf_val = float(carry_forward_debt or 0.0)
			cur_note = note.strip()
		exp_amt = max(0.0, round(p_days * mpd * up + cf_val - disc, 2))
		conn.execute(
			"""UPDATE cafeteria_monthly_fees
			   SET planned_days=?, meals_per_day=?, unit_meal_price=?, discount_or_adjustment=?,
			       carry_forward_debt=?, expected_amount=?, note=?, updated_at=?
			   WHERE id=?""",
			(p_days, mpd, up, disc, cf_val, exp_amt, cur_note, now_iso, int(fee_id)),
		)


def carry_forward_unpaid_balances(
	connect_db_fn: Callable[[], sqlite3.Connection],
	from_year_month: str,
	to_year_month: str,
) -> int:
	"""Önceki ayda ödenmeyen kalan borcu (veya fazla ödemeyi) hedef ayın `carry_forward_debt` sütununa aktarır."""
	prev_summary = get_cash_summary_and_diner_balances(connect_db_fn, from_year_month)
	prev_bal_by_did = {
		int(b["diner_id"]): round(float(b["expected_amount"]) - float(b["paid_amount"]), 2)
		for b in prev_summary["balances"]
	}
	if not prev_bal_by_did:
		return 0
	to_ym = str(to_year_month)[:7]
	now_iso = datetime.now().isoformat(timespec="minutes")
	updated = 0
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute("SELECT * FROM cafeteria_monthly_fees WHERE year_month=?", (to_ym,)).fetchall()
		for r in rows:
			did = int(r["diner_id"])
			if did not in prev_bal_by_did:
				continue
			cf = prev_bal_by_did[did]
			exp_amt = max(0.0, round(float(r["planned_days"]) * float(r["meals_per_day"]) * float(r["unit_meal_price"]) + cf - float(r["discount_or_adjustment"]), 2))
			conn.execute(
				"UPDATE cafeteria_monthly_fees SET carry_forward_debt=?, expected_amount=?, updated_at=? WHERE id=?",
				(cf, exp_amt, now_iso, int(r["id"])),
			)
			updated += 1
	return updated


def record_cash_transaction(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	tx_date: str,
	tx_type: str,
	amount: float,
	diner_id: int | None = None,
	for_person_name: str = "",
	delivered_by_name: str = "",
	received_by_name: str = "",
	payment_method: str = "Nakit",
	receipt_no: str = "",
	notes: str = "",
	year_month: str | None = None,
) -> int:
	"""
	Kasa defterine hareket ekler:
	- `IN_PAYMENT`: Tahsilat (getiren isim + öğrenci/kişi + miktar kasaya girer)
	- `OUT_SCHOOL_HANDOVER`: Kasadan okula teslim edilen para (kasadan çıkar)
	- `OUT_REFUND_OR_PETTY`: İade veya kasadan harcama (kasadan çıkar)
	"""
	amt = round(max(0.0, float(amount or 0.0)), 2)
	if amt <= 0:
		raise ValueError("Tutar 0'dan büyük olmalıdır.")
	dt_s = str(tx_date or date.today().isoformat()).strip()[:10]
	ym = str(year_month or dt_s[:7])[:7]
	valid_types = {"IN_PAYMENT", "OUT_SCHOOL_HANDOVER", "OUT_REFUND_OR_PETTY"}
	t_type = tx_type if tx_type in valid_types else "IN_PAYMENT"
	now_iso = datetime.now().isoformat(timespec="minutes")

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		for_nm = str(for_person_name or "").strip()
		if diner_id and not for_nm:
			dn = conn.execute("SELECT full_name FROM cafeteria_diners WHERE id=?", (int(diner_id),)).fetchone()
			if dn:
				for_nm = str(dn["full_name"])
		cur = conn.execute(
			"""INSERT INTO cafeteria_cash_ledger(
				tx_date, year_month, tx_type, diner_id, for_person_name,
				delivered_by_name, received_by_name, amount, payment_method, receipt_no, notes, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				dt_s, ym, t_type, int(diner_id) if diner_id else None, for_nm,
				str(delivered_by_name or "").strip(), str(received_by_name or "").strip(),
				amt, str(payment_method or "Nakit").strip(), str(receipt_no or "").strip(),
				str(notes or "").strip(), now_iso,
			),
		)
		return int(cur.lastrowid)


def delete_cash_transaction(connect_db_fn: Callable[[], sqlite3.Connection], tx_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute("DELETE FROM cafeteria_cash_ledger WHERE id=?", (int(tx_id),))


def record_bulk_sibling_payments(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	tx_date: str,
	delivered_by_name: str,
	allocations: list[dict[str, Any]],
	payment_method: str = "Nakit",
	receipt_no: str = "",
	notes: str = "",
	year_month: str | None = None,
) -> list[int]:
	"""Tek kişinin (ör. veli) birden fazla öğrenci için getirdiği ödemeyi kasa defterine kalem kalem işler."""
	created_ids: list[int] = []
	for alloc in (allocations or []):
		if not isinstance(alloc, dict):
			continue
		amt = round(max(0.0, _safe_float(alloc.get("amount"), 0.0)), 2)
		if amt <= 0:
			continue
		did = alloc.get("diner_id")
		tx_id = record_cash_transaction(
			connect_db_fn,
			tx_date=tx_date,
			tx_type="IN_PAYMENT",
			amount=amt,
			diner_id=int(did) if did else None,
			for_person_name=str(alloc.get("for_person_name") or ""),
			delivered_by_name=delivered_by_name,
			payment_method=payment_method,
			receipt_no=receipt_no,
			notes=str(alloc.get("notes") or notes or "Toplu / Kardeş Tahsilatı").strip(),
			year_month=year_month,
		)
		created_ids.append(tx_id)
	return created_ids


def get_cash_summary_and_diner_balances(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
) -> dict[str, Any]:
	"""Seçili ay tahakkuk/bakiye çizelgesi (son getiren dahil) ile hem aylık hem tüm zamanlar kasa net bakiyesini döndürür."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		fee_rows = conn.execute(
			"""SELECT f.*, d.full_name, d.role_type, d.class_or_dept, d.fee_exempt
			   FROM cafeteria_monthly_fees f
			   JOIN cafeteria_diners d ON d.id = f.diner_id
			   WHERE f.year_month=?
			   ORDER BY d.class_or_dept ASC, d.full_name ASC""",
			(ym,),
		).fetchall()
		paid_rows = conn.execute(
			"""SELECT diner_id, COALESCE(SUM(amount), 0) AS paid
			   FROM cafeteria_cash_ledger
			   WHERE year_month=? AND tx_type='IN_PAYMENT' AND diner_id IS NOT NULL
			   GROUP BY diner_id""",
			(ym,),
		).fetchall()
		paid_map = {int(r["diner_id"]): float(r["paid"] or 0.0) for r in paid_rows}

		month_tx = [
			dict(r)
			for r in conn.execute(
				"SELECT * FROM cafeteria_cash_ledger WHERE year_month=? ORDER BY tx_date DESC, id DESC",
				(ym,),
			).fetchall()
		]
		all_tx_sums = conn.execute(
			"""SELECT
				COALESCE(SUM(CASE WHEN tx_type='IN_PAYMENT' THEN amount ELSE 0 END), 0) AS tot_in,
				COALESCE(SUM(CASE WHEN tx_type='OUT_SCHOOL_HANDOVER' THEN amount ELSE 0 END), 0) AS tot_handover,
				COALESCE(SUM(CASE WHEN tx_type='OUT_REFUND_OR_PETTY' THEN amount ELSE 0 END), 0) AS tot_other_out
			   FROM cafeteria_cash_ledger"""
		).fetchone()

	last_pay_info: dict[int, tuple[str, str]] = {}
	for t in month_tx:
		if t.get("tx_type") == "IN_PAYMENT" and t.get("diner_id"):
			did_k = int(t["diner_id"])
			if did_k not in last_pay_info:
				last_pay_info[did_k] = (str(t.get("delivered_by_name") or ""), str(t.get("tx_date") or ""))

	balances: list[dict[str, Any]] = []
	month_expected = 0.0
	month_paid_matched = 0.0
	for fr in fee_rows:
		d = dict(fr)
		did = int(d["diner_id"])
		exp = round(float(d.get("expected_amount") or 0.0), 2)
		paid = round(paid_map.get(did, 0.0), 2)
		rem = round(exp - paid, 2)
		if d.get("fee_exempt") or d.get("role_type") == "Burslu / Ücretsiz" or exp == 0:
			st = "Muaf / Burslu"
		elif paid >= exp:
			st = "🟢 Ödendi"
		elif paid > 0:
			st = "🟡 Kısmi Ödeme"
		else:
			st = "🔴 Ödenmedi"
		d["paid_amount"] = paid
		d["remaining_amount"] = max(0.0, rem)
		d["status_label"] = st
		d["carry_forward_debt"] = round(float(d.get("carry_forward_debt") or 0.0), 2)
		lp = last_pay_info.get(did, ("", ""))
		d["last_delivered_by"] = lp[0]
		d["last_payment_date"] = lp[1]
		balances.append(d)
		month_expected += exp
		month_paid_matched += paid

	m_in = sum(float(t["amount"] or 0.0) for t in month_tx if t["tx_type"] == "IN_PAYMENT")
	m_hand = sum(float(t["amount"] or 0.0) for t in month_tx if t["tx_type"] == "OUT_SCHOOL_HANDOVER")
	m_out_other = sum(float(t["amount"] or 0.0) for t in month_tx if t["tx_type"] == "OUT_REFUND_OR_PETTY")

	tot_in = float(all_tx_sums["tot_in"] or 0.0)
	tot_hand = float(all_tx_sums["tot_handover"] or 0.0)
	tot_other = float(all_tx_sums["tot_other_out"] or 0.0)
	vault_balance = round(tot_in - tot_hand - tot_other, 2)

	return {
		"balances": balances,
		"month_transactions": month_tx,
		"month_expected": round(month_expected, 2),
		"month_collected": round(m_in, 2),
		"month_handover": round(m_hand, 2),
		"month_other_out": round(m_out_other, 2),
		"month_net": round(m_in - m_hand - m_out_other, 2),
		"month_remaining": round(max(0.0, month_expected - month_paid_matched), 2),
		"vault_total_in": round(tot_in, 2),
		"vault_total_handover": round(tot_hand, 2),
		"vault_total_other_out": round(tot_other, 2),
		"vault_balance": vault_balance,
	}


def build_payment_receipt_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	tx_id: int,
	school_meta: dict[str, str] | None = None,
) -> bytes:
	"""Tekil bir ücret tahsilatı (`IN_PAYMENT`) için resmi Yemekhane Tahsilat Alındı Makbuzu (.docx) üretir."""
	dx = _ensure_docx_mod()
	docx = dx["docx"]
	Cm, Pt, RGBColor = dx["Cm"], dx["Pt"], dx["RGBColor"]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		tx_r = conn.execute("SELECT * FROM cafeteria_cash_ledger WHERE id=?", (int(tx_id),)).fetchone()
		if not tx_r:
			raise ValueError(f"Tahsilat kaydı bulunamadı (#{tx_id}).")
		tx = dict(tx_r)
	ym = str(tx.get("year_month") or date.today().strftime("%Y-%m"))[:7]
	cash = get_cash_summary_and_diner_balances(connect_db_fn, ym)
	rem_bal = None
	if tx.get("diner_id"):
		for b in cash["balances"]:
			if int(b["diner_id"]) == int(tx["diner_id"]):
				rem_bal = b["remaining_amount"]
				break

	meta = school_meta or {}
	doc = docx.Document()
	for s in doc.sections:
		s.top_margin = Cm(1.4)
		s.bottom_margin = Cm(1.4)
		s.left_margin = Cm(1.6)
		s.right_margin = Cm(1.6)

	for txt_l, fsz, bf, col in [
		(meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI", 9.5, False, RGBColor(100, 116, 139)),
		((meta.get("school_name") or "OKUL YEMEKHANESİ").upper(), 12.5, True, RGBColor(15, 118, 110)),
		(f"YEMEKHANE ÜCRET TAHSİLAT ALINDI MAKBUZU (Kayıt #{int(tx_id)})", 11.0, True, RGBColor(30, 41, 59)),
	]:
		p = doc.add_paragraph()
		p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		p.paragraph_format.space_after = Pt(2)
		rr = p.add_run(txt_l)
		rr.font.name = "Calibri"
		rr.font.size = Pt(fsz)
		rr.bold = bf
		rr.font.color.rgb = col

	kv = [
		("Makbuz / Kayıt Tarihi", str(tx.get("tx_date") or "")),
		("Dönem", _fmt_ym(ym)),
		("Kimin Adına (Öğrenci / Abone)", str(tx.get("for_person_name") or "—")),
		("Ücreti Getiren / Ödeyen İsim", str(tx.get("delivered_by_name") or "—")),
		("Tahsil Edilen Tutar", f"{float(tx.get('amount') or 0.0):,.2f} ₺"),
		("Ödeme Yöntemi", str(tx.get("payment_method") or "Nakit")),
		("Makbuz / Açıklama", f"{tx.get('receipt_no') or ''} {tx.get('notes') or ''}".strip() or "—"),
	]
	if rem_bal is not None:
		kv.append(("Bu Ödeme Sonrası Dönem Kalanı", f"{float(rem_bal):,.2f} ₺"))

	tbl = doc.add_table(rows=len(kv), cols=2)
	tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	for ri, (k, v) in enumerate(kv):
		c1, c2 = tbl.cell(ri, 0), tbl.cell(ri, 1)
		_style_docx_cell(dx, c1, "F1F5F9")
		_style_docx_cell(dx, c2, "FFFFFF")
		r1 = c1.paragraphs[0].add_run(k)
		r1.bold = True
		r1.font.size = Pt(9.5)
		r2 = c2.paragraphs[0].add_run(v)
		r2.font.size = Pt(9.5)

	p_sp = doc.add_paragraph()
	p_sp.paragraph_format.space_before = Pt(18)
	stbl = doc.add_table(rows=1, cols=2)
	stbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	for ci, (r_l, nm_l) in enumerate([
		("Ödemeyi Yapan / Getiren", str(tx.get("delivered_by_name") or "........................")),
		("Tahsil Eden Yetkili", str(tx.get("received_by_name") or meta.get("nurse_name") or "........................")),
	]):
		sc = stbl.cell(0, ci)
		sp = sc.paragraphs[0]
		sp.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		sr1 = sp.add_run(f"{r_l}\n")
		sr1.bold = True
		sr1.font.size = Pt(9.5)
		sr2 = sp.add_run(f"{nm_l}\n\nİmza / Tarih")
		sr2.font.size = Pt(8.5)

	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def get_budget_and_pnl_summary(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
) -> dict[str, Any]:
	"""Seçili ay ve tüm zamanlar için yemekhane Aidat Geliri vs Fatura/Mutfak Gideri bütçe denkliğini hesaplar."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	cash = get_cash_summary_and_diner_balances(connect_db_fn, ym)
	exp = get_expense_summary_stats(connect_db_fn, ym)
	month_net_cash_vs_exp = round(cash["month_collected"] - exp["month_total"], 2)
	month_net_accrual_vs_exp = round(cash["month_expected"] - exp["month_total"], 2)
	all_net_vs_exp = round(cash["vault_total_in"] - exp["all_time_total"], 2)
	coll_rate = round((cash["month_collected"] / cash["month_expected"] * 100.0), 1) if cash["month_expected"] > 0 else 0.0
	return {
		"year_month": ym,
		"month_expected": cash["month_expected"],
		"month_collected": cash["month_collected"],
		"collection_rate_pct": coll_rate,
		"month_expenses": exp["month_total"],
		"month_paid_expenses": exp["month_paid_total"],
		"month_pending_expenses": exp["month_pending_total"],
		"month_net_cash_vs_exp": month_net_cash_vs_exp,
		"month_net_accrual_vs_exp": month_net_accrual_vs_exp,
		"all_time_collected": cash["vault_total_in"],
		"all_time_expenses": exp["all_time_total"],
		"all_time_net_vs_exp": all_net_vs_exp,
		"vault_balance": cash["vault_balance"],
		"by_expense_category": exp["by_category"],
	}


def build_cash_handover_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	school_meta: dict[str, str] | None = None,
	deliverer_name: str = "",
	receiver_name: str = "",
) -> bytes:
	"""Seçili ayın yemekhane kasa tahsilatı ve okula teslim tutanağını Word (.docx) olarak üretir."""
	dx = _ensure_docx_mod()
	docx = dx["docx"]
	Cm, Pt, RGBColor = dx["Cm"], dx["Pt"], dx["RGBColor"]
	summary = get_cash_summary_and_diner_balances(connect_db_fn, year_month)
	yy, mm = [int(x) for x in str(year_month)[:7].split("-")]
	m_lbl = f"{TR_MONTH_NAMES.get(mm, mm)} {yy}"
	meta = school_meta or {}

	doc = docx.Document()
	for s in doc.sections:
		s.top_margin = Cm(1.3)
		s.bottom_margin = Cm(1.3)
		s.left_margin = Cm(1.4)
		s.right_margin = Cm(1.4)

	for txt_l, fsz, bf, col in [
		(meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI", 10, False, RGBColor(100, 116, 139)),
		((meta.get("school_name") or "OKUL YEMEKHANESİ").upper(), 13, True, RGBColor(15, 118, 110)),
		(f"{m_lbl} YEMEKHANE TAHSİLAT VE OKULA KASA TESLİM TUTANAĞI", 11.5, True, RGBColor(30, 41, 59)),
	]:
		p = doc.add_paragraph()
		p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		p.paragraph_format.space_after = Pt(2)
		r = p.add_run(txt_l)
		r.font.name = "Calibri"
		r.font.size = Pt(fsz)
		r.bold = bf
		r.font.color.rgb = col

	kv_rows = [
		("Dönem", m_lbl),
		("Aylık Beklenen Toplam Aidat", f"{summary['month_expected']:,.2f} ₺"),
		("Dönem Tahsilatı (Kasaya Giren)", f"{summary['month_collected']:,.2f} ₺"),
		("Dönem Okula Teslim Edilen", f"{summary['month_handover']:,.2f} ₺"),
		("Kalan Bekleyen Aidat", f"{summary['month_remaining']:,.2f} ₺"),
		("Güncel Kasa Mevcudu (Tüm Dönemler Net)", f"{summary['vault_balance']:,.2f} ₺"),
	]
	kv_tbl = doc.add_table(rows=len(kv_rows), cols=2)
	kv_tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	for idx, (k, v) in enumerate(kv_rows):
		c1, c2 = kv_tbl.cell(idx, 0), kv_tbl.cell(idx, 1)
		_style_docx_cell(dx, c1, "F1F5F9")
		_style_docx_cell(dx, c2, "FFFFFF")
		r1 = c1.paragraphs[0].add_run(k)
		r1.bold = True
		r1.font.size = Pt(9.5)
		r2 = c2.paragraphs[0].add_run(v)
		r2.font.size = Pt(9.5)

	txs = summary["month_transactions"]
	if txs:
		p_t = doc.add_paragraph()
		p_t.paragraph_format.space_before = Pt(10)
		rt = p_t.add_run("Dönem İçi Kasa ve Teslim Hareketleri")
		rt.bold = True
		rt.font.color.rgb = RGBColor(15, 118, 110)

		t_tbl = doc.add_table(rows=len(txs) + 1, cols=6)
		hdrs = ["Tarih", "İşlem Türü", "Kimin Adına", "Getiren / Teslim Alan", "Yöntem / Makbuz", "Tutar (₺)"]
		for ci, h_txt in enumerate(hdrs):
			hc = t_tbl.cell(0, ci)
			_style_docx_cell(dx, hc, "0F766E")
			rr = hc.paragraphs[0].add_run(h_txt)
			rr.bold = True
			rr.font.size = Pt(8.5)
			rr.font.color.rgb = RGBColor(255, 255, 255)
		kind_lbls = {
			"IN_PAYMENT": "Tahsilat (+)",
			"OUT_SCHOOL_HANDOVER": "Okula Teslim (-)",
			"OUT_REFUND_OR_PETTY": "İade/Harcama (-)",
		}
		for ri, tx in enumerate(txs, start=1):
			vals = [
				str(tx.get("tx_date") or ""),
				kind_lbls.get(tx.get("tx_type"), tx.get("tx_type")),
				str(tx.get("for_person_name") or "—"),
				str(tx.get("delivered_by_name") or tx.get("received_by_name") or "—"),
				f"{tx.get('payment_method') or 'Nakit'} {('· ' + tx['receipt_no']) if tx.get('receipt_no') else ''}".strip(),
				f"{float(tx.get('amount') or 0.0):,.2f} ₺",
			]
			for ci, vv in enumerate(vals):
				cc = t_tbl.cell(ri, ci)
				_style_docx_cell(dx, cc, "FFFFFF" if ri % 2 == 1 else "F8FAFC")
				rr = cc.paragraphs[0].add_run(vv)
				rr.font.size = Pt(8.5)

	p_sp = doc.add_paragraph()
	p_sp.paragraph_format.space_before = Pt(16)
	stbl = doc.add_table(rows=1, cols=2)
	stbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	for ci, (r_lbl, p_nm) in enumerate([
		("Teslim Eden (Yemekhane / Revir Sorumlusu)", deliverer_name or meta.get("nurse_name") or "........................"),
		("Teslim Alan (Okul İdaresi / Müdürü)", receiver_name or meta.get("principal_name") or "........................"),
	]):
		sc = stbl.cell(0, ci)
		sp = sc.paragraphs[0]
		sp.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		sr1 = sp.add_run(f"{r_lbl}\n")
		sr1.bold = True
		sr1.font.size = Pt(9.5)
		sr2 = sp.add_run(f"{p_nm}\n\nİmza / Tarih")
		sr2.font.size = Pt(8.5)

	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_cafeteria_full_workbook_excel(connect_db_fn: Callable[[], sqlite3.Connection], year_month: str) -> bytes:
	"""Seçili ay için Yemekler, Aylık Menü, Faturalar/Giderler ve Aidat/Kasa çalışma kitabını (.xlsx) üretir."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	dishes = get_all_dishes(connect_db_fn)
	plans = get_menu_plans_for_month(connect_db_fn, ym)
	purchases, items = get_purchases_and_items(connect_db_fn, year_month=ym)
	cash = get_cash_summary_and_diner_balances(connect_db_fn, ym)
	app_cfg_xl: dict[str, str] = {}
	try:
		with connect_db_fn() as conn:
			for sr in conn.execute("SELECT key, value FROM app_settings").fetchall():
				app_cfg_xl[str(sr["key"])] = str(sr["value"])
	except Exception:
		pass
	sched_xl = parse_service_schedule_from_cfg(app_cfg_xl)
	ram_xl = get_ramazan_times_for_dates(connect_db_fn, sched_xl["ramazan_province"], [r["plan_date"] for r in plans.values()]) if sched_xl["ramazan_mode"] else {}

	df_dishes = pd.DataFrame([
		{
			"Yemek Adı": d["name"],
			"Kategori": d["category"],
			"Porsiyon (g)": d["portion_g"],
			"Kalori (kcal)": d["total_kcal"],
			"Karbonhidrat (g)": d["carbs_g"],
			"Alerjenler": ", ".join(d.get("allergens") or []),
		}
		for d in dishes
	]) if dishes else pd.DataFrame([{"Bilgi": "Kayıtlı yemek yok"}])

	df_menu = pd.DataFrame([
		{
			"Tarih": rec["plan_date"],
			"Öğün": rec["meal_slot"],
			**({
				"Sahur (İmsak)": ram_xl.get(rec["plan_date"], {}).get("sahur", ""),
				"İftar": ram_xl.get(rec["plan_date"], {}).get("iftar", ""),
			} if sched_xl["ramazan_mode"] else {}),
			"Yemekler": " · ".join(rec.get("dish_names") or []),
			"Kalori (kcal)": rec["total_kcal"],
			"Karbonhidrat (g)": rec["total_carbs_g"],
			"Alerjen Uyarısı": ", ".join(rec.get("allergens") or []) or "Alerjen: Yok",
		}
		for rec in plans.values()
	]) if plans else pd.DataFrame([{"Bilgi": f"{ym} ayı menüsü boş"}])

	df_items = pd.DataFrame([
		{
			"Fatura Tarihi": it["invoice_date"],
			"Fatura No": it.get("invoice_no") or "",
			"Tedarikçi": it.get("supplier_name") or "",
			"Ürün Adı": it["item_name"],
			"Kategori": it["category"],
			"Miktar": it["quantity"],
			"Birim": it["unit_type"],
			"Birim Fiyat (₺)": it["unit_price"],
			"KDV (%)": it["vat_rate"],
			"Tutar (₺)": it["line_total"],
		}
		for it in items
	]) if items else pd.DataFrame([{"Bilgi": f"{ym} ayı ürün/fatura kaydı yok"}])

	df_fees = pd.DataFrame([
		{
			"Ad Soyad": b["full_name"],
			"Statü": b["role_type"],
			"Sınıf/Birim": b.get("class_or_dept") or "",
			"İş Günü": b["planned_days"],
			"Günlük Öğün": b["meals_per_day"],
			"Öğün Ücreti (₺)": b["unit_meal_price"],
			"Devreden (₺)": b.get("carry_forward_debt", 0.0),
			"Beklenen (₺)": b["expected_amount"],
			"Ödenen (₺)": b["paid_amount"],
			"Kalan (₺)": b["remaining_amount"],
			"Son Getiren": b.get("last_delivered_by") or "",
			"Durum": b["status_label"],
		}
		for b in cash["balances"]
	]) if cash["balances"] else pd.DataFrame([{"Bilgi": f"{ym} ayı abone aidat kaydı yok"}])

	df_cash = pd.DataFrame([
		{
			"Tarih": t["tx_date"],
			"Tür": t["tx_type"],
			"Kimin Adına": t.get("for_person_name") or "",
			"Getiren / Teslim Alan": t.get("delivered_by_name") or t.get("received_by_name") or "",
			"Tutar (₺)": t["amount"],
			"Yöntem": t["payment_method"],
			"Makbuz/Not": f"{t.get('receipt_no') or ''} {t.get('notes') or ''}".strip(),
		}
		for t in cash["month_transactions"]
	]) if cash["month_transactions"] else pd.DataFrame([{"Bilgi": f"{ym} ayı kasa hareketi yok"}])

	buf = io.BytesIO()
	with pd.ExcelWriter(buf, engine="openpyxl") as writer:
		df_menu.to_excel(writer, sheet_name="Aylik_Menu", index=False)
		df_dishes.to_excel(writer, sheet_name="Yemekler_Bankasi", index=False)
		df_items.to_excel(writer, sheet_name="Gelen_Urun_Gider", index=False)
		df_fees.to_excel(writer, sheet_name="Abone_Aidat", index=False)
		df_cash.to_excel(writer, sheet_name="Kasa_Defteri", index=False)
		if plans:
			try:
				from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
				ws_m = writer.sheets["Aylik_Menu"]
				unique_dates = [{"date_iso": d_i} for d_i in sorted({str(r["plan_date"])[:10] for r in plans.values()})]
				xl_styles, _ = _build_week_day_style_map(
					unique_dates,
					color_tone=app_cfg_xl.get("cafeteria_menu_color_tone"),
				)
				hdr_fill = PatternFill("solid", fgColor="0F766E")
				hdr_font = Font(color="FFFFFF", bold=True)
				thin_side = Side(style="thin", color="94A3B8")
				cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
				for ci in range(1, ws_m.max_column + 1):
					hc = ws_m.cell(row=1, column=ci)
					hc.fill = hdr_fill
					hc.font = hdr_font
					hc.border = cell_border
					hc.alignment = Alignment(horizontal="center", vertical="center")
				for ri in range(2, ws_m.max_row + 1):
					d_val = str(ws_m.cell(row=ri, column=1).value or "")[:10]
					st_xl = xl_styles.get(d_val)
					if not st_xl:
						continue
					r_fill = PatternFill("solid", fgColor=st_xl["bg_word"])
					r_font = Font(color="0F172A")
					for ci in range(1, ws_m.max_column + 1):
						cc = ws_m.cell(row=ri, column=ci)
						cc.fill = r_fill
						cc.font = Font(color=st_xl["date_fg_hex"].lstrip("#"), bold=True) if ci == 1 else r_font
						cc.border = cell_border
			except Exception:
				pass
	return buf.getvalue()


# ==============================================================================
# 7. UI SAYFA BİLEŞENLERİ (5 SEKME)
# ==============================================================================

def _month_options(connect_db_fn: Callable[[], sqlite3.Connection] | None = None) -> tuple[list[str], str]:
	today = date.today()
	opts_set: set[str] = set()
	for off in range(-4, 8):
		tot = today.year * 12 + (today.month - 1) + off
		yy, mm = divmod(tot, 12)
		opts_set.add(f"{yy:04d}-{mm + 1:02d}")
	if callable(connect_db_fn):
		try:
			with connect_db_fn() as conn:
				for tbl in ("cafeteria_menu_plans", "cafeteria_purchases", "cafeteria_monthly_fees", "cafeteria_cash_ledger"):
					for r in conn.execute(f"SELECT DISTINCT year_month FROM {tbl} WHERE TRIM(COALESCE(year_month, '')) != ''").fetchall():
						ym_s = str(r[0] or "")[:7]
						if re.match(r"^\d{4}-\d{2}$", ym_s):
							opts_set.add(ym_s)
		except Exception:
			pass
	cur = f"{today.year:04d}-{today.month:02d}"
	return sorted(opts_set), cur


def _fmt_ym(ym_s: str) -> str:
	try:
		yy, mm = ym_s.split("-")
		return f"{TR_MONTH_NAMES.get(int(mm), mm)} {yy}"
	except Exception:
		return ym_s


def _get_dish_planned_usages(connect_db_fn: Callable[[], sqlite3.Connection], dish_id: int) -> list[dict[str, str]]:
	"""Verilen yemeğin planlı menülerde geçtiği tarih ve öğünleri döndürür."""
	out: list[dict[str, str]] = []
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		rows = conn.execute(
			"SELECT plan_date, meal_slot, year_month, dish_ids_json FROM cafeteria_menu_plans ORDER BY plan_date ASC, meal_slot ASC"
		).fetchall()
		for r in rows:
			try:
				d_ids = [int(x) for x in json.loads(r["dish_ids_json"] or "[]")]
			except Exception:
				d_ids = []
			if int(dish_id) in d_ids:
				out.append({"Tarih": str(r["plan_date"]), "Öğün": str(r["meal_slot"]), "Dönem": str(r["year_month"])})
	return out


def _render_dish_editor_body(
	ctx_api: dict[str, Any],
	*,
	editing_dish: dict[str, Any] | None,
	key_prefix: str,
	in_dialog: bool = False,
) -> None:
	"""Yemek adı, makro besinler, AI malzeme çözücü, ölçeklenebilir malzeme tablosu ve alerji/menü etki bölümlerini çizer."""
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()
	mask_fn = ctx_api["mask_kvkk_value"]

	sel_sig = f"id:{editing_dish['id']}" if editing_dish else "new"
	sig_k = f"{key_prefix}_loaded_sig"
	draft_k = f"{key_prefix}_edit_draft"
	if ui.session_state.get(sig_k) != sel_sig:
		ui.session_state[sig_k] = sel_sig
		if editing_dish:
			ui.session_state[f"{key_prefix}_name"] = editing_dish["name"]
			ui.session_state[f"{key_prefix}_cat"] = editing_dish["category"] if editing_dish["category"] in DISH_CATEGORIES else DISH_CATEGORIES[1]
			ui.session_state[f"{key_prefix}_port"] = float(editing_dish.get("portion_g") or 250.0)
			ui.session_state[f"{key_prefix}_prot"] = float(editing_dish.get("protein_g") or 0.0)
			ui.session_state[f"{key_prefix}_fat"] = float(editing_dish.get("fat_g") or 0.0)
			ui.session_state[f"{key_prefix}_rec"] = str(editing_dish.get("recipe_notes") or "")
			ui.session_state[f"{key_prefix}_ex_als"] = [a for a in (editing_dish.get("allergens") or []) if a in ALLERGEN_BADGE_MAP]
			ui.session_state[draft_k] = [dict(x) for x in (editing_dish.get("ingredients") or [])]
		else:
			ui.session_state[f"{key_prefix}_name"] = ""
			ui.session_state[f"{key_prefix}_cat"] = DISH_CATEGORIES[1]
			ui.session_state[f"{key_prefix}_port"] = 250.0
			ui.session_state[f"{key_prefix}_prot"] = 0.0
			ui.session_state[f"{key_prefix}_fat"] = 0.0
			ui.session_state[f"{key_prefix}_rec"] = ""
			ui.session_state[f"{key_prefix}_ex_als"] = []
			ui.session_state[draft_k] = []

	# Bölüm 1: Temel Kimlik & Porsiyon
	if in_dialog:
		ui.markdown("##### 1️⃣ Temel Yemek Bilgileri & Porsiyon")
	dc1, dc2, dc3, dc4, dc5 = ui.columns([2.0, 1.3, 0.85, 0.85, 0.85])
	with dc1:
		d_name = ui.text_input("Yemek Adı*", placeholder="Örn: Mercimek Çorbası, İzmir Köfte", key=f"{key_prefix}_name")
	with dc2:
		d_cat = ui.selectbox("Kategori", DISH_CATEGORIES, key=f"{key_prefix}_cat")
	with dc3:
		d_port = ui.number_input("Porsiyon (g)", min_value=20.0, step=10.0, key=f"{key_prefix}_port")
	with dc4:
		d_prot = ui.number_input("Protein (g)", min_value=0.0, step=1.0, key=f"{key_prefix}_prot")
	with dc5:
		d_fat = ui.number_input("Yağ (g)", min_value=0.0, step=1.0, key=f"{key_prefix}_fat")

	# Bölüm 2: AI Malzeme & Besin Çözümleyici
	if in_dialog:
		ui.markdown("##### 2️⃣ Yapay Zekâ Reçete & Malzeme Çözümleyici")
	ai_c1, ai_c2 = ui.columns([2.2, 1.3], vertical_alignment="center")
	with ai_c1:
		ai_note = ui.text_input("AI Reçete İpucu (Opsiyonel)", placeholder="Örn: İlkokul porsiyonu, tereyağlı, glütensiz", key=f"{key_prefix}_ai_note")
	with ai_c2:
		if ui.button("✨ Yapay Zekâ ile Malzeme, Kalori & Alerjen Çıkar", disabled=not is_on, type="primary" if is_on else "secondary", use_container_width=True, key=f"{key_prefix}_btn_ai"):
			if not str(d_name or "").strip():
				ui.warning("Lütfen önce yemek adını yazın.")
			else:
				with ui.spinner(f"Gemini AI '{d_name}' için 1 porsiyon malzeme ve besin analizini çıkarıyor..."):
					ok_d, res_d, used_m = gai.generate_dish_ingredients_and_nutrition_with_gemini(
						api_key=api_key, model=model, dish_name=d_name, category=d_cat, extra_notes=ai_note
					)
				if ok_d:
					ui.session_state[draft_k] = res_d.get("ingredients") or []
					if res_d.get("portion_g"):
						ui.session_state[f"{key_prefix}_port"] = float(res_d["portion_g"])
					if res_d.get("protein_g") is not None:
						ui.session_state[f"{key_prefix}_prot"] = float(res_d.get("protein_g") or 0.0)
					if res_d.get("fat_g") is not None:
						ui.session_state[f"{key_prefix}_fat"] = float(res_d.get("fat_g") or 0.0)
					if res_d.get("recipe_notes"):
						ui.session_state[f"{key_prefix}_rec"] = str(res_d["recipe_notes"])
					ui.success(f"✅ {len(ui.session_state[draft_k])} malzeme Gemini ({used_m}) ile hazırlandı! Tablodan dilediğiniz gibi düzenleyebilirsiniz.")
					ui.rerun()
				else:
					ui.error(f"AI hatası: {(res_d or {}).get('error', 'Yanıt alınamadı')}")

	# Bölüm 3: Malzeme Listesi & Otomatik Miktar Oranlama
	if in_dialog:
		ui.markdown("##### 3️⃣ Malzeme Listesi & Otomatik Kalori Oranlama")
	draft_ings: list[dict[str, Any]] = ui.session_state.get(draft_k) or []
	auto_scale = ui.checkbox("⚖️ Malzeme miktarını değiştirdiğimde kaloriyi oranla (AI kapalıyken de çalışır)", value=True, key=f"{key_prefix}_autoscale")
	df_ing_in = pd.DataFrame([
		{
			"Malzeme": x.get("ingredient_name") or "",
			"Miktar": _safe_float(x.get("amount"), 0.0),
			"Birim": x.get("unit") if x.get("unit") in INGREDIENT_UNITS else "g",
			"Kalori (kcal)": _safe_float(x.get("kcal_portion"), 0.0),
			"Karb (g)": _safe_float(x.get("carbs_portion_g"), 0.0),
			"Alerjen": _normalize_allergen_label(x.get("allergen_tag") or ""),
		}
		for x in draft_ings
	]) if draft_ings else pd.DataFrame(columns=["Malzeme", "Miktar", "Birim", "Kalori (kcal)", "Karb (g)", "Alerjen"])

	ed_ing_df = ui.data_editor(
		df_ing_in,
		num_rows="dynamic",
		use_container_width=True,
		hide_index=True,
		key=f"{key_prefix}_ing_editor_{sel_sig}",
		column_config={
			"Malzeme": ui.column_config.TextColumn("Malzeme Adı", required=True),
			"Miktar": ui.column_config.NumberColumn("Miktar", min_value=0.0, step=5.0),
			"Birim": ui.column_config.SelectboxColumn("Birim", options=INGREDIENT_UNITS),
			"Kalori (kcal)": ui.column_config.NumberColumn("Kalori (kcal)", min_value=0.0, step=1.0),
			"Karb (g)": ui.column_config.NumberColumn("Karb (g)", min_value=0.0, step=0.5),
			"Alerjen": ui.column_config.SelectboxColumn("Alerjen Uyarısı", options=TR_ALLERGEN_OPTIONS),
		},
	)

	scaled_ings = scale_edited_ingredients(draft_ings, ed_ing_df, auto_scale_nutrition=auto_scale)
	tot_k, tot_c, dish_allergens = compute_dish_nutrition_from_ingredients(scaled_ings)
	extra_als = ui.multiselect(
		"Ek Manuel Alerjen Etiketleri (Malzeme dışı / Eser miktar)",
		[a for a in TR_ALLERGEN_OPTIONS if a],
		key=f"{key_prefix}_ex_als",
	)
	merged_allergens = list(dish_allergens)
	for ea in (extra_als or []):
		if ea and ea not in merged_allergens:
			merged_allergens.append(ea)
	al_badges = " · ".join(format_allergen_badge(a) for a in merged_allergens) if merged_allergens else "Belirgin 14 temel alerjen yok"

	# Bölüm 4: Besin Özeti, Revir Çapraz Güvenliği & Menü Kullanımı
	if in_dialog:
		ui.markdown("##### 4️⃣ Besin Özeti, Alerjen Çakışması & Menü Kullanımı")
	m1, m2, m3, m4 = ui.columns(4)
	m1.metric("Porsiyon Kalorisi", f"{tot_k:g} kcal")
	m2.metric("Karbonhidrat", f"{tot_c:g} g")
	m3.metric("Protein / Yağ", f"{d_prot:g}g / {d_fat:g}g")
	m4.metric("Alerjen Türü", f"{len(merged_allergens)} tür")
	if merged_allergens:
		ui.warning(f"⚠️ **Yemek Alerjen Uyarısı:** {al_badges}")
		stu_hits = match_allergens_against_school(connect_db_fn, merged_allergens)
		if stu_hits:
			hit_txt = ", ".join(f"{mask_fn(h['full_name'], 'name')} ({h['class_name']} - {h['matched']})" for h in stu_hits[:8])
			ui.error(f"🚨 **Okulda Çakışan Alerjik / Çölyaklı / Özel Diyet Kişi ({len(stu_hits)}):** {hit_txt}")
	else:
		ui.info(f"🟢 {al_badges}")

	if in_dialog and editing_dish:
		usages = _get_dish_planned_usages(connect_db_fn, int(editing_dish["id"]))
		if usages:
			with ui.expander(f"📅 Bu Yemeğin Kullanıldığı Planlı Öğünler ({len(usages)}) — Kaydedince Kalori/Alerjen Otomatik Güncellenir"):
				ui.dataframe(pd.DataFrame(usages), hide_index=True, use_container_width=True)

	# Bölüm 5: Tarif / Pişirme Notu & İşlemler
	if in_dialog:
		ui.markdown("##### 5️⃣ Tarif / Servis Notu & Kayıt İşlemleri")
	d_rec = ui.text_input("Tarif / Pişirme Notu", key=f"{key_prefix}_rec", placeholder="Kısa hazırlık veya servis notu...")

	if in_dialog and editing_dish:
		b1, b2, b3, b4 = ui.columns([1.8, 1.3, 1.0, 0.8])
		with b1:
			if ui.button("💾 Değişiklikleri Kaydet", type="primary", use_container_width=True, key=f"{key_prefix}_btn_save"):
				try:
					saved_id = upsert_dish_with_ingredients(
						connect_db_fn,
						dish_id=int(editing_dish["id"]),
						name=d_name,
						category=d_cat,
						portion_g=d_port,
						ingredients=scaled_ings,
						protein_g=d_prot,
						fat_g=d_fat,
						extra_allergens=extra_als,
						recipe_notes=d_rec,
						source="ai" if is_on else str(editing_dish.get("source") or "manual"),
					)
					ui.session_state[sig_k] = None
					ui.session_state["caf_dish_loaded_sig"] = None
					ui.session_state["caf_dish_toast"] = f"✅ '{d_name}' (#{saved_id}) güncellendi ve bağlı menüler yenilendi ({tot_k:g} kcal)."
					ui.close_dialog(all_dialogs=True)
					ui.rerun()
				except Exception as exc:
					ui.error(str(exc))
		with b2:
			if ui.button("📑 Kopyasını Oluştur", type="secondary", use_container_width=True, key=f"{key_prefix}_btn_clone"):
				try:
					clone_name = f"{str(d_name).strip()} (Kopya)"
					new_id = upsert_dish_with_ingredients(
						connect_db_fn,
						name=clone_name,
						category=d_cat,
						portion_g=d_port,
						ingredients=scaled_ings,
						protein_g=d_prot,
						fat_g=d_fat,
						extra_allergens=extra_als,
						recipe_notes=d_rec,
						source="manual",
					)
					ui.session_state["caf_dish_toast"] = f"✅ '{clone_name}' (#{new_id}) yeni yemek olarak çoğaltıldı."
					ui.close_dialog(all_dialogs=True)
					ui.rerun()
				except Exception as exc:
					ui.error(str(exc))
		with b3:
			if ui.button("🗑️ Sil", type="secondary", use_container_width=True, key=f"{key_prefix}_btn_del"):
				delete_dish_by_id(connect_db_fn, int(editing_dish["id"]))
				ui.session_state[sig_k] = None
				ui.session_state["caf_dish_toast"] = f"🗑️ '{editing_dish['name']}' yemekler bankasından silindi."
				ui.close_dialog(all_dialogs=True)
				ui.rerun()
		with b4:
			if ui.button("Kapat", use_container_width=True, key=f"{key_prefix}_btn_close"):
				ui.close_dialog(all_dialogs=True)
				ui.rerun()
	else:
		if ui.button("💾 Yemeği ve Malzemelerini Kaydet / Güncelle", type="primary", use_container_width=True, key=f"{key_prefix}_btn_save"):
			try:
				saved_id = upsert_dish_with_ingredients(
					connect_db_fn,
					dish_id=int(editing_dish["id"]) if editing_dish else None,
					name=d_name,
					category=d_cat,
					portion_g=d_port,
					ingredients=scaled_ings,
					protein_g=d_prot,
					fat_g=d_fat,
					extra_allergens=extra_als,
					recipe_notes=d_rec,
					source="ai" if is_on else "manual",
				)
				ui.session_state[draft_k] = scaled_ings
				ui.session_state[sig_k] = None
				ui.success(f"✅ '{d_name}' (#{saved_id}) yemekler bankasına kaydedildi ve bağlı menüler güncellendi ({tot_k:g} kcal).")
				ui.rerun()
			except Exception as exc:
				ui.error(str(exc))


@ui.dialog("🍽️ Yemek Detayı, Malzeme & Besin Düzenleyici", width="large")
def open_dish_editor_dialog(ctx_api: dict[str, Any], dish_id: int) -> None:
	"""Tablodan adına tıklanan yemeğin tüm bilgilerini düzenlemek için açılan 5 bölümlü pop-up pencere."""
	dishes = get_all_dishes(ctx_api["connect_db"])
	target = next((d for d in dishes if int(d["id"]) == int(dish_id)), None)
	if not target:
		ui.warning("Seçilen yemek bulunamadı.")
		if ui.button("Kapat", key="caf_dlg_missing_close"):
			ui.close_dialog(all_dialogs=True)
		return
	src_lbl = "🤖 Gemini AI Destekli" if target.get("source") == "ai" else "✍️ Manuel / Kurum Reçetesi"
	ui.caption(f"Kayıt #{int(target['id'])} · Kaynak: {src_lbl} · Son Güncelleme: {target.get('updated_at') or '—'}")
	_render_dish_editor_body(ctx_api, editing_dish=target, key_prefix="caf_dlg_d", in_dialog=True)


def _render_tab_dishes(ctx_api: dict[str, Any]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()

	ui.markdown("### 🍲 Kayıtlı Yemekler Bankası, AI Malzeme/Kalori & Alerjen Çözümleyici")
	ui.caption("Yeni yemek ekleyin veya aşağıdaki **Kayıtlı Yemekler Listesi** tablosunda yemek adına tıklayarak pop-up düzenleme ekranını açın.")
	if ui.session_state.get("caf_dish_toast"):
		ui.success(ui.session_state.pop("caf_dish_toast"))

	dishes = get_all_dishes(connect_db_fn)
	dish_map = {f"{d['name']} ({d['category']} · {d['total_kcal']:g} kcal)": d for d in dishes}

	top_c1, top_c2 = ui.columns([2.2, 1.0], vertical_alignment="center")
	with top_c1:
		edit_choices = ["➕ Yeni Yemek Ekle"] + list(dish_map.keys())
		picked_lbl = ui.selectbox("Düzenlenecek veya Eklenecek Yemek", edit_choices, key="caf_dish_picker")
	editing_dish = dish_map.get(picked_lbl) if picked_lbl != "➕ Yeni Yemek Ekle" else None
	with top_c2:
		if editing_dish and ui.button("🪟 Seçili Yemeği Pop-up'ta Aç", type="secondary", use_container_width=True, key="caf_btn_open_popup_dish"):
			open_dish_editor_dialog(ctx_api, int(editing_dish["id"]))

	with ui.container(border=True):
		_render_dish_editor_body(ctx_api, editing_dish=editing_dish, key_prefix="caf_d", in_dialog=False)

	if dishes:
		missing_dishes = [d for d in dishes if not (d.get("ingredients") or []) or float(d.get("total_kcal") or 0) <= 0]
		if is_on and missing_dishes:
			with ui.expander(f"✨ Malzeme / Kalorisi Eksik {len(missing_dishes)} Yemeği AI ile Toplu Tamamla (Önizlemeli)"):
				if ui.button("🤖 Eksik Yemekleri Analiz Et", key="caf_btn_batch_enrich"):
					payload_in = [{"id": d["id"], "name": d["name"], "category": d["category"]} for d in missing_dishes[:15]]
					with ui.spinner("Gemini AI eksik yemeklerin malzemelerini ve kalorilerini çıkarıyor..."):
						ok_be, res_be, used_be = gai.batch_enrich_dishes_with_gemini(api_key=api_key, model=model, dishes_to_enrich=payload_in)
					if ok_be and res_be.get("enriched_dishes"):
						ui.session_state["caf_ai_batch_dishes_prev"] = res_be["enriched_dishes"]
						ui.success(f"✅ {len(res_be['enriched_dishes'])} yemek analiz edildi ({used_be}).")
					else:
						ui.error(f"Toplu AI hatası: {(res_be or {}).get('error', 'Yanıt alınamadı')}")
				b_prev = ui.session_state.get("caf_ai_batch_dishes_prev")
				if isinstance(b_prev, list) and b_prev:
					ui.dataframe(pd.DataFrame([
						{"Yemek": x.get("name"), "Porsiyon (g)": x.get("portion_g"), "Malzeme Sayısı": len(x.get("ingredients") or [])}
						for x in b_prev
					]), hide_index=True, use_container_width=True)
					if ui.button("✅ Onayla ve Tümünü Yemekler Bankasına Kaydet", type="primary", key="caf_btn_commit_batch_dishes"):
						for x in b_prev:
							if x.get("name"):
								upsert_dish_with_ingredients(
									connect_db_fn,
									dish_id=int(x["id"]) if str(x.get("id") or "").isdigit() else None,
									name=str(x["name"]),
									category=str(x.get("category") or "Ana Yemek"),
									portion_g=_safe_float(x.get("portion_g"), 250.0),
									ingredients=x.get("ingredients") or [],
									recipe_notes=str(x.get("recipe_notes") or ""),
									source="ai",
								)
						ui.session_state.pop("caf_ai_batch_dishes_prev", None)
						ui.success(f"✅ {len(b_prev)} yemek bankaya işlendi.")
						ui.rerun()

		ui.markdown("---")
		ui.markdown(f"##### 📚 Kayıtlı Yemekler Listesi ({len(dishes)})")
		ui.caption("💡 Pop-up düzenleme ekranını açmak için tablodaki **Yemek Adı** hücresine tıklayın.")
		df_list = pd.DataFrame([
			{
				"id": int(d["id"]),
				"Yemek Adı": d["name"],
				"Kategori": d["category"],
				"Porsiyon (g)": d["portion_g"],
				"Kalori (kcal)": d["total_kcal"],
				"Karb (g)": d["carbs_g"],
				"Malzeme Sayısı": len(d.get("ingredients") or []),
				"Alerjen Uyarısı": " · ".join(format_allergen_badge(a) for a in (d.get("allergens") or [])) or "—",
			}
			for d in dishes
		])
		tbl_ev = ui.dataframe(
			df_list,
			hide_index=True,
			use_container_width=True,
			on_select="rerun",
			selection_mode=["single-row", "single-cell"],
			key="caf_dishes_list_tbl",
			column_config={
				"id": None,
				"Yemek Adı": ui.column_config.TextColumn("Yemek Adı", help="Detay, malzeme, kalori ve alerjen düzenlemek için tıklayın"),
			},
		)
		sel_cells = getattr(tbl_ev.selection, "cells", []) or []
		clicked_dish_row = None
		for c in sel_cells:
			c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
			c_col = c[1] if isinstance(c, (tuple, list)) else (c.get("column") if isinstance(c, dict) else None)
			if c_col == "Yemek Adı" and c_row is not None and 0 <= int(c_row) < len(df_list):
				clicked_dish_row = int(c_row)
				break
		if clicked_dish_row is not None:
			clicked_did = int(df_list.iloc[clicked_dish_row]["id"])
			ui.clear_dataframe_selection("caf_dishes_list_tbl")
			open_dish_editor_dialog(ctx_api, clicked_did)


def _render_tab_menu_planner(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()
	mask_fn = ctx_api["mask_kvkk_value"]

	ui.markdown("### 🗓️ İş Günü Bazlı Öğün Menü Planlayıcı & A4 Çıktı")
	dishes = get_all_dishes(connect_db_fn)
	dish_lbl_to_id = {f"{d['name']} ({d['total_kcal']:g} kcal)": int(d["id"]) for d in dishes}
	id_to_dish_lbl = {v: k for k, v in dish_lbl_to_id.items()}

	sched = parse_service_schedule_from_cfg(app_cfg)
	cfg_c1, cfg_c2, cfg_c3 = ui.columns([1.8, 1.1, 1.1])
	with cfg_c1:
		active_slots = ui.multiselect(
			"Planlanacak Öğünler",
			sched["available_slots"],
			default=[s for s in sched["active_meal_slots"] if s in sched["available_slots"]] or ["Öğle Yemeği"],
			key="caf_menu_active_slots",
		) or sched["active_meal_slots"]
	with cfg_c2:
		inc_wknd = ui.checkbox("Hafta Sonlarını da Göster", value=any(wd >= 5 for wd in sched["weekdays_set"]), key="caf_menu_inc_wknd")
	with cfg_c3:
		auto_sync_diab = ui.checkbox("💉 Diyabete Oto-Eşitle", value=True, key="caf_menu_auto_sync_diab")

	days_info = get_month_workdays_info(connect_db_fn, sel_ym, include_weekends=inc_wknd)
	workdays = [d for d in days_info if d.get("is_workday")]
	holidays_in_month = [d for d in days_info if d.get("is_holiday")]
	plans = get_menu_plans_for_month(connect_db_fn, sel_ym)
	ram_times_planner = (
		get_ramazan_times_for_dates(connect_db_fn, sched["ramazan_province"], [d["date_iso"] for d in workdays])
		if sched["ramazan_mode"]
		else {}
	)

	tot_planned_days = len({p[0] for p in plans.keys()})
	slot_times_disp = " · ".join(f"{sl} ({sched['meal_times'][sl]})" for sl in active_slots if sched["meal_times"].get(sl))
	ram_info_bit = f" · 🌙 **Ramazan Modu ({sched['ramazan_province']})**" if sched["ramazan_mode"] else ""
	ui.info(
		f"📅 **{_fmt_ym(sel_ym)} Hizmet & İş Günü Özeti:** Haftada **{sched['days_per_week']} gün**, günde **{len(active_slots)} öğün** "
		f"({slot_times_disp}) · Ay toplamı **{len(workdays)} hizmet günü** · Planlanmış gün: **{tot_planned_days}**{ram_info_bit}"
		+ (f" · Resmi/İdari Tatil ({len(holidays_in_month)}): " + ", ".join(f"{h['day_num']} ({h['holiday_title']})" for h in holidays_in_month[:4]) if holidays_in_month else "")
	)

	# AI Akıllı Menü Oluşturucu & Dosyadan Aktarıcı
	with ui.expander("🤖 Gemini AI ile Aylık/Haftalık Dengeli Menü Tasarla veya Dosyadan Aktar"):
		if not is_on:
			ui.caption("AI kapalıyken aşağıdaki iş günü kutularından manuel planlama yapabilirsiniz.")
		else:
			ac1, ac2 = ui.columns(2)
			with ac1:
				ai_menu_inst = ui.text_input("AI Menü Tercihi", placeholder="Örn: Haftada 2 gün kurubaklagil, 1 gün tavuk/balık, dengeli öğrenci kalorisi", key="caf_ai_menu_inst")
				if ui.button("✨ Seçili Ayın İş Günlerine AI Menü Tasarla (Önizlemeli)", type="primary", use_container_width=True, key="caf_btn_ai_gen_menu"):
					bank_compact = [
						{"id": d["id"], "name": d["name"], "category": d["category"], "kcal": d["total_kcal"], "allergens": d.get("allergens") or []}
						for d in dishes
					]
					with ui.spinner("Gemini AI iş günleri için kalori ve alerjen dengeli menü oluşturuyor..."):
						ok_gm, res_gm, used_gm = gai.generate_monthly_cafeteria_menu_with_gemini(
							api_key=api_key,
							model=model,
							year_month=sel_ym,
							workdays_list=[{"date": w["date_iso"], "weekday": w["weekday_name"]} for w in workdays],
							meal_slots=active_slots,
							available_dishes=bank_compact,
							custom_instruction=ai_menu_inst,
						)
					if ok_gm and res_gm.get("plans"):
						ui.session_state["caf_ai_menu_preview"] = {"plans": res_gm["plans"], "note": f"{len(res_gm['plans'])} öğün planı ({used_gm}) hazırlandı."}
					else:
						ui.error(f"AI menü hatası: {(res_gm or {}).get('error', 'Yanıt alınamadı')}")
			with ac2:
				up_mfile = ui.file_uploader("📎 Mevcut Yemek Listesi Yükle (.pdf, .docx, .xlsx, .png, .jpg)", type=["pdf", "docx", "xlsx", "xls", "png", "jpg", "jpeg", "txt"], key="caf_ai_menu_file")
				paste_mtxt = ui.text_area("Veya Menü Metni Yapıştır", height=68, key="caf_ai_menu_paste", placeholder="02.11.2026 Öğle: Mercimek Çorbası, Etli Nohut...")
				if ui.button("📥 Belgedeki Menüyü Oku (Önizlemeli)", use_container_width=True, key="caf_btn_ai_parse_menu"):
					fb = up_mfile.getvalue() if up_mfile is not None else None
					fn = up_mfile.name if up_mfile is not None else ""
					strict_pdf = str(app_cfg.get("gemini_strict_kvkk_pdf", "0")).strip() == "1"
					with ui.spinner("Gemini AI belgeyi iş günlerine ayrıştırıyor..."):
						ok_pm, res_pm, used_pm = gai.parse_cafeteria_menu_or_payments_text_with_gemini(
							api_key=api_key,
							model=model,
							task_mode="menu",
							workdays_list=[w["date_iso"] for w in workdays],
							raw_text=paste_mtxt,
							file_bytes=fb,
							file_name=fn,
							strict_kvkk_pdf=strict_pdf,
						)
					if ok_pm and res_pm.get("plans"):
						ui.session_state["caf_ai_menu_preview"] = {"plans": res_pm["plans"], "note": f"{len(res_pm['plans'])} öğün belgeden çıkarıldı ({used_pm})."}
					else:
						ui.error(f"Menü okuma hatası: {(res_pm or {}).get('error', 'Satır bulunamadı')}")

			prev_m = ui.session_state.get("caf_ai_menu_preview")
			if isinstance(prev_m, dict) and prev_m.get("plans"):
				ui.success(f"🤖 **Önizleme:** {prev_m.get('note')}")
				ui.dataframe(pd.DataFrame(prev_m["plans"]).head(25), hide_index=True, use_container_width=True)
				save_new_bank = ui.checkbox("✨ Bankada olmayan yeni yemekleri Yemekler Bankasına da otomatik kaydet", value=True, key="caf_ai_menu_save_bank")
				pm1, pm2 = ui.columns(2)
				with pm1:
					if ui.button("✅ Onayla ve AI Menüsünü Takvime İşle", type="primary", use_container_width=True, key="caf_btn_commit_ai_menu"):
						by_name = {_tr_fold(d["name"]): int(d["id"]) for d in dishes}
						applied_n = 0
						for item in prev_m["plans"]:
							p_dt = str(item.get("plan_date") or "")[:10]
							sl_nm = str(item.get("meal_slot") or active_slots[0])
							d_ids_p: list[int] = []
							cust_p: list[dict[str, Any]] = []
							for raw_dnm in (item.get("dishes") or []):
								if isinstance(raw_dnm, dict):
									dnm_s = str(raw_dnm.get("name") or "").strip()
									k_c = _safe_float(raw_dnm.get("kcal"), 0.0)
									c_c = _safe_float(raw_dnm.get("carbs_g"), 0.0)
									al_c = str(raw_dnm.get("allergen") or "")
								else:
									dnm_s = str(raw_dnm or "").strip()
									k_c, c_c, al_c = 0.0, 0.0, ""
								if not dnm_s:
									continue
								found_id = by_name.get(_tr_fold(dnm_s))
								if found_id:
									d_ids_p.append(found_id)
								elif save_new_bank:
									new_id = upsert_quick_dish_from_plan_item(
										connect_db_fn,
										{"name": dnm_s, "kcal": k_c, "carbs_g": c_c, "allergen": al_c},
									)
									if new_id:
										by_name[_tr_fold(dnm_s)] = new_id
										d_ids_p.append(new_id)
								else:
									cust_p.append({"name": dnm_s, "kcal": k_c, "carbs_g": c_c, "allergen": al_c})
							if p_dt and (d_ids_p or cust_p):
								upsert_menu_plan_slot(
									connect_db_fn,
									plan_date=p_dt,
									meal_slot=sl_nm,
									dish_ids=d_ids_p,
									custom_items=cust_p,
									sync_to_diabetes=auto_sync_diab,
								)
								applied_n += 1
						ui.session_state.pop("caf_ai_menu_preview", None)
						ui.success(f"✅ {applied_n} öğün planı kaydedildi.")
						ui.rerun()
				with pm2:
					if ui.button("✖ Önizlemeyi İptal Et", use_container_width=True, key="caf_btn_cancel_ai_menu"):
						ui.session_state.pop("caf_ai_menu_preview", None)
						ui.rerun()

	# AI Klinik & Beslenme Güvenlik Denetçisi
	if is_on and plans:
		with ui.expander("🛡️ Gemini AI Menü Güvenlik & Beslenme Denetçisi (Alerji · Çölyak · Tip 1 Diyabet)"):
			if ui.button("🔍 Seçili Ay Menüsünü Klinik Risklere Göre Denetle", key="caf_btn_ai_audit_menu"):
				risk_ctx = get_clinical_nutrition_risk_summary(connect_db_fn, mask_fn=mask_fn)
				plan_sum = [
					{
						"plan_date": r["plan_date"],
						"meal_slot": r["meal_slot"],
						"dishes": r.get("dish_names") or [],
						"total_kcal": r.get("total_kcal"),
						"total_carbs_g": r.get("total_carbs_g"),
						"allergens": r.get("allergens") or [],
					}
					for r in plans.values()
				]
				dish_comp = [{"name": d["name"], "category": d["category"], "kcal": d["total_kcal"], "allergens": d.get("allergens") or []} for d in dishes]
				with ui.spinner("Gemini AI planlı menüyü alerji, çölyak ve diyabet güvenliği açısından denetliyor..."):
					ok_au, res_au, used_au = gai.audit_cafeteria_menu_with_gemini(
						api_key=api_key,
						model=model,
						year_month=sel_ym,
						planned_menu_summary=plan_sum,
						clinical_risk_summary=risk_ctx,
						available_dishes=dish_comp,
					)
				if ok_au:
					ui.session_state["caf_ai_menu_audit"] = res_au
					ui.success(f"✅ Güvenlik denetimi tamamlandı ({used_au}).")
				else:
					ui.error(f"Denetim hatası: {(res_au or {}).get('error', 'Yanıt alınamadı')}")

			au_res = ui.session_state.get("caf_ai_menu_audit")
			if isinstance(au_res, dict) and au_res.get("overall_summary"):
				ui.metric("Beslenme & Klinik Uyum Skoru", f"{au_res.get('balance_score', '—')} / 100")
				ui.info(str(au_res["overall_summary"]))
				if au_res.get("risk_findings"):
					ui.dataframe(pd.DataFrame(au_res["risk_findings"]), hide_index=True, use_container_width=True)
				for wn in (au_res.get("weekly_nutrition_notes") or []):
					ui.caption(f"• {wn}")

	# Hafta Bazlı Menü Düzenleyici (Seçili ayın iş günleri hafta hafta düzenli kartlar)
	weeks = _group_month_days_by_week(days_info)
	if not weeks:
		ui.warning("Seçilen ay için aktif iş günü bulunamadı.")
		return

	wk_labels = [f"{i + 1}. Hafta ({w[0]['day_num']} – {w[-1]['day_num']} {TR_MONTH_NAMES[int(sel_ym[5:7])]})" for i, w in enumerate(weeks)]
	w_col1, w_col2, w_col3, w_col4 = ui.columns([1.8, 1.1, 1.1, 1.2], vertical_alignment="center")
	with w_col1:
		picked_wk_lbl = ui.selectbox("Düzenlenecek Haftayı Seçin (İş Günleri)", wk_labels, key="caf_menu_wk_pick")
	wk_idx = wk_labels.index(picked_wk_lbl) if picked_wk_lbl in wk_labels else 0
	with w_col2:
		if ui.button("💾 Bu Haftanın Tümünü Kaydet", type="primary", use_container_width=True, key="caf_btn_save_wk_all"):
			cnt_w = 0
			for d in weeks[wk_idx]:
				d_iso = d["date_iso"]
				for sl in active_slots:
					sel_lbls_w = ui.session_state.get(f"caf_ms_{d_iso}_{sl}") or []
					ex_raw_w = ui.session_state.get(f"caf_ex_{d_iso}_{sl}") or ""
					c_ids_w = [dish_lbl_to_id[lbl] for lbl in sel_lbls_w if lbl in dish_lbl_to_id]
					c_cust_w = parse_custom_dish_entries(ex_raw_w)
					if c_ids_w or c_cust_w or ((d_iso, sl) in plans):
						upsert_menu_plan_slot(
							connect_db_fn,
							plan_date=d_iso,
							meal_slot=sl,
							dish_ids=c_ids_w,
							custom_items=c_cust_w,
							sync_to_diabetes=auto_sync_diab,
						)
						cnt_w += 1
			ui.success(f"✅ Seçili haftada {cnt_w} öğün güncellendi.")
			ui.rerun()
	with w_col3:
		if wk_idx > 0 and ui.button("📋 Önceki Haftayı Kopyala", type="secondary", use_container_width=True, key="caf_btn_copy_prev_wk"):
			pairs = [(s["date_iso"], t["date_iso"]) for s, t in zip(weeks[wk_idx - 1], weeks[wk_idx])]
			c_cnt = copy_menu_plans(connect_db_fn, date_pairs=pairs, meal_slots=active_slots, overwrite=True, sync_to_diabetes=auto_sync_diab)
			ui.success(f"✅ {c_cnt} öğün önceki haftadan kopyalandı.")
			ui.rerun()
		elif wk_idx == 0 and ui.button("📆 Geçen Aydan Kopyala", type="secondary", use_container_width=True, key="caf_btn_copy_prev_mo"):
			yy_i, mm_i = int(sel_ym[:4]), int(sel_ym[5:7])
			pm_y, pm_m = (yy_i - 1, 12) if mm_i == 1 else (yy_i, mm_i - 1)
			prev_wd = [d["date_iso"] for d in get_month_workdays_info(connect_db_fn, f"{pm_y:04d}-{pm_m:02d}", include_weekends=inc_wknd) if d.get("is_workday")]
			cur_wd = [d["date_iso"] for d in workdays]
			pairs = list(zip(prev_wd, cur_wd))
			c_cnt = copy_menu_plans(connect_db_fn, date_pairs=pairs, meal_slots=active_slots, overwrite=False, sync_to_diabetes=auto_sync_diab)
			ui.success(f"✅ Geçen aydan {c_cnt} öğün bu aya dağıtıldı.")
			ui.rerun()
	with w_col4:
		if ui.button("💉 Tüm Ayı Diyabete Eşitle", type="secondary", use_container_width=True, key="caf_btn_sync_diab_all"):
			cnt_s = sync_cafeteria_month_to_diabetes(connect_db_fn, sel_ym)
			ui.success(f"✅ {cnt_s} öğün İnsülin & Diyabet Yönetimi modülüyle eşitlendi.")

	def _render_week_nav_bar(pos_tag: str) -> None:
		nav_l, nav_c, nav_r = ui.columns([1.05, 2.3, 1.05], vertical_alignment="center")
		with nav_l:
			if ui.button(
				"◀ Önceki Hafta",
				key=f"caf_btn_prev_wk_{pos_tag}",
				disabled=(wk_idx <= 0),
				use_container_width=True,
				help="Bir önceki haftaya geç",
			):
				ui.session_state["caf_menu_wk_pick"] = wk_labels[wk_idx - 1]
				ui.rerun()
		with nav_c:
			ui.markdown(
				f"<div style='text-align:center;font-weight:700;font-size:13px;padding:6px 10px;border-radius:8px;"
				f"background:rgba(15,118,110,0.08);border:1px solid rgba(15,118,110,0.25);'>"
				f"🗓️ {html.escape(wk_labels[wk_idx])} &nbsp;·&nbsp; ({wk_idx + 1} / {len(weeks)})</div>",
				unsafe_allow_html=True,
			)
		with nav_r:
			if ui.button(
				"Sonraki Hafta ▶",
				key=f"caf_btn_next_wk_{pos_tag}",
				disabled=(wk_idx >= len(weeks) - 1),
				use_container_width=True,
				help="Bir sonraki haftaya geç",
			):
				ui.session_state["caf_menu_wk_pick"] = wk_labels[wk_idx + 1]
				ui.rerun()

	_render_week_nav_bar("top")

	diab_stus = get_diabetic_students_for_menu(connect_db_fn)
	for d in weeks[wk_idx]:
		d_iso = d["date_iso"]
		half_badge = f" · ⏳ {d['holiday_title']}" if d.get("is_half_day") else ""
		rt_day = ram_times_planner.get(d_iso)
		ram_badge = f" · 🌙 Sahur: {rt_day['sahur']} · 🌅 İftar: {rt_day['iftar']}" if rt_day else ""
		with ui.container(border=True):
			ui.markdown(f"**📆 {d['day_num']:02d}.{sel_ym[5:7]}.{sel_ym[:4]} — {d['weekday_name']}**{html.escape(half_badge)}{html.escape(ram_badge)}")
			day_tot_kcal = 0.0
			for sl in active_slots:
				rec = plans.get((d_iso, sl))
				cur_ids = rec.get("dish_ids") or [] if rec else []
				cur_lbls = [id_to_dish_lbl[int(x)] for x in cur_ids if int(x) in id_to_dish_lbl]
				if rec:
					cust_parts = []
					for c in (rec.get("custom_items") or []):
						if c.get("name"):
							meta_b = []
							if _safe_float(c.get("kcal"), 0.0) > 0:
								meta_b.append(f"{_safe_float(c['kcal']):g} kcal")
							if _safe_float(c.get("carbs_g"), 0.0) > 0:
								meta_b.append(f"{_safe_float(c['carbs_g']):g}g karb")
							if c.get("allergen"):
								meta_b.append(str(c["allergen"]))
							cust_parts.append(f"{c['name']} [{', '.join(meta_b)}]" if meta_b else str(c["name"]))
					cur_cust_str = ", ".join(cust_parts)
				else:
					cur_cust_str = ""

				sc1, sc2, sc3 = ui.columns([2.2, 1.4, 0.8], vertical_alignment="center")
				with sc1:
					sel_dish_lbls = ui.multiselect(
						f"🍽️ {sl}",
						list(dish_lbl_to_id.keys()),
						default=cur_lbls,
						key=f"caf_ms_{d_iso}_{sl}",
					)
				with sc2:
					extra_free = ui.text_input(
						f"Ek Serbest Kalem ({sl})",
						value=cur_cust_str,
						placeholder="Örn: Tam Buğday Ekmek [130 kcal, 24g karb, gluten]",
						key=f"caf_ex_{d_iso}_{sl}",
					)
				with sc3:
					if ui.button("💾 Kaydet", key=f"caf_save_slot_{d_iso}_{sl}", use_container_width=True):
						chosen_ids = [dish_lbl_to_id[lbl] for lbl in sel_dish_lbls if lbl in dish_lbl_to_id]
						cust_list = parse_custom_dish_entries(extra_free)
						upsert_menu_plan_slot(
							connect_db_fn,
							plan_date=d_iso,
							meal_slot=sl,
							dish_ids=chosen_ids,
							custom_items=cust_list,
							sync_to_diabetes=auto_sync_diab,
						)
						ui.success(f"{d_iso} ({sl}) kaydedildi.")
						ui.rerun()

				if rec and rec.get("dish_names"):
					kv = _safe_float(rec.get("total_kcal"), 0.0)
					cgv = _safe_float(rec.get("total_carbs_g"), 0.0)
					day_tot_kcal += kv
					als = rec.get("allergens") or []
					al_disp = " · ".join(format_allergen_badge(a) for a in als) if als else "Alerjen uyarısı yok"
					hits = match_allergens_against_school(connect_db_fn, als) if als else []
					hit_s = f" | 🚨 **{len(hits)} alerjik/çölyaklı kişi çakışması:** " + ", ".join(mask_fn(h["full_name"], "name") for h in hits[:4]) if hits else ""
					diab_s = f" | 💙 **Yüksek Karbonhidrat ({len(diab_stus)} diyabetli öğrenci için dikkat)**" if (diab_stus and cgv >= 85.0) else ""
					ui.caption(f"🔥 **{sl} Kalorisi:** `{kv:g} kcal` · **Karb:** `{cgv:g} g` · ⚠️ **Alerjenler:** {al_disp}{hit_s}{diab_s}")

			if len(active_slots) > 1 and day_tot_kcal > 0:
				ui.caption(f"📊 **Gün Toplamı:** `{day_tot_kcal:g} kcal`")

	_render_week_nav_bar("bot")

	# Yalnızca Öğün Girilmiş Günler Özeti (Boş Günler Otomatik Çıkarılır)
	planned_only_rows: list[dict[str, Any]] = []
	for d in workdays:
		d_iso = d["date_iso"]
		if not any(plans.get((d_iso, sl)) and plans[(d_iso, sl)].get("dish_names") for sl in active_slots):
			continue
		rt_d = ram_times_planner.get(d_iso)
		for sl in active_slots:
			rec = plans.get((d_iso, sl))
			if rec and rec.get("dish_names"):
				t_s = sched["meal_times"].get(sl, "")
				planned_only_rows.append({
					"Tarih": f"{d['day_num']:02d}.{sel_ym[5:7]}.{sel_ym[:4]} ({d['weekday_name']})",
					**({"🌙 Sahur (İmsak)": rt_d["sahur"], "🌅 İftar": rt_d["iftar"]} if (sched["ramazan_mode"] and rt_d) else {}),
					"Öğün": f"{sl} ({t_s})" if t_s else sl,
					"Yemekler": " · ".join(rec["dish_names"]),
					"Kalori (kcal)": rec.get("total_kcal") or 0.0,
					"Karbonhidrat (g)": rec.get("total_carbs_g") or 0.0,
					"Alerjen": ", ".join(rec.get("allergens") or []) or "Alerjen: Yok",
				})
	ui.markdown("---")
	ui.markdown(f"#### 📋 Planlanmış Menü Listesi ({len(planned_only_rows)} Öğün — Öğün Girilmemiş Günler Otomatik Çıkarılır)")
	if planned_only_rows:
		ui.dataframe(pd.DataFrame(planned_only_rows), hide_index=True, use_container_width=True)
	else:
		ui.caption("Seçili ay için henüz öğün girilmiş gün bulunmuyor.")

	# A4 Çıktı ve İmzalı Belge Oluşturucu
	ui.markdown("---")
	ui.markdown("#### 🖨️ Resmi A4 Menü Çıktısı (Haftalık Renk Tonlu, Kalori / Karb. / Alerjen & İmzalı)")
	ui.caption("ℹ️ A4 çıktıda yalnızca menü/öğün girilen günler basılır; her hafta ayrı renk tonuyla ve ardışık günler açık/koyu ton farkıyla renklendirilir.")
	p_c1, p_c2, p_c3, p_c4, p_c5 = ui.columns([1.2, 1.3, 0.75, 0.85, 0.85])
	with p_c1:
		pr_scope = ui.radio("Çıktı Kapsamı", ["📅 Tüm Ay", "🗓️ Seçili Hafta"], horizontal=True, key="caf_print_scope")
	with p_c2:
		pr_layout_lbl = ui.selectbox(
			"A4 Sayfa Düzeni",
			["Standart Liste", "Tek A4 Sayfaya Sığdır (Kompakt)", "Haftalık Tablo Izgarası (Pzt-Cum)"],
			key="caf_print_layout",
		)
	with p_c3:
		pr_kcal = ui.checkbox("Kalori (kcal) Göster", value=True, key="caf_print_kcal")
	with p_c4:
		pr_carb = ui.checkbox("Karbonhidrat (g) Göster", value=True, key="caf_print_carbs")
	with p_c5:
		pr_al = ui.checkbox("Alerjen Uyarısı Göster", value=True, key="caf_print_al")
	scope_code = "weekly" if "Hafta" in pr_scope else "monthly"
	layout_code = "compact_one_page" if "Kompakt" in pr_layout_lbl else ("weekly_grid" if "Izgarası" in pr_layout_lbl else "standard")

	signers = resolve_menu_signers_for_print(app_cfg)
	a4_html = build_menu_a4_html(
		connect_db_fn,
		year_month=sel_ym,
		scope=scope_code,
		week_index=wk_idx,
		active_slots=active_slots,
		show_kcal=pr_kcal,
		show_carbs=pr_carb,
		show_allergens=pr_al,
		include_weekends=inc_wknd,
		layout_mode=layout_code,
		signers=signers,
		school_meta=app_cfg,
	)
	a4_docx = build_menu_a4_docx(
		connect_db_fn,
		year_month=sel_ym,
		scope=scope_code,
		week_index=wk_idx,
		active_slots=active_slots,
		show_kcal=pr_kcal,
		show_carbs=pr_carb,
		show_allergens=pr_al,
		include_weekends=inc_wknd,
		layout_mode=layout_code,
		signers=signers,
		school_meta=app_cfg,
	)

	dl1, dl2 = ui.columns(2)
	with dl1:
		ui.download_button(
			"📝 A4 Resmi Menü İndir (.docx)",
			data=a4_docx,
			file_name=f"Yemekhane_Menu_{sel_ym}_{scope_code}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			type="primary",
			use_container_width=True,
			key="caf_dl_menu_docx",
		)
	with dl2:
		ui.download_button(
			"🖨️ Yazdırılabilir A4 HTML İndir (.html)",
			data=a4_html.encode("utf-8"),
			file_name=f"Yemekhane_Menu_{sel_ym}_{scope_code}.html",
			mime="text/html",
			use_container_width=True,
			key="caf_dl_menu_html",
		)
	with ui.expander("👁️ A4 Kağıdı Önizleme ve Doğrudan Yazdırma", expanded=False):
		ui.components.v1.html(a4_html, height=620, scrolling=True)


def _calc_live_invoice_totals(rows_in: list[dict[str, Any]]) -> tuple[float, float, float, int]:
	sub_v, vat_v, valid_n = 0.0, 0.0, 0
	for x in rows_in:
		nm_x = str(x.get("Ürün Adı") or "").strip()
		if not nm_x:
			continue
		q_x = max(0.0, _safe_float(x.get("Miktar"), 0.0))
		p_x = max(0.0, _safe_float(x.get("Birim Fiyat (₺)"), 0.0))
		vr_x = max(0.0, _safe_float(x.get("KDV (%)"), 0.0))
		b_x = q_x * p_x
		sub_v += b_x
		vat_v += b_x * (vr_x / 100.0)
		valid_n += 1
	return round(sub_v, 2), round(vat_v, 2), round(sub_v + vat_v, 2), valid_n


def _render_invoice_editor_body(
	ctx_api: dict[str, Any],
	*,
	editing_purchase: dict[str, Any] | None,
	key_prefix: str,
	in_dialog: bool,
) -> None:
	"""Fatura başlığını ve içindeki ürün/fiyat kalemlerini görüntüleme, ekleme, çıkarma ve düzenleme formu."""
	connect_db_fn = ctx_api["connect_db"]
	editing_pid = int(editing_purchase["id"]) if editing_purchase else None
	p_sig = f"pid:{editing_pid}" if editing_pid else "new"
	sig_key = f"{key_prefix}_loaded_sig"

	if ui.session_state.get(sig_key) != p_sig:
		ui.session_state[sig_key] = p_sig
		ui.session_state.pop(f"{key_prefix}_items_editor", None)
		if editing_purchase:
			try:
				ui.session_state[f"{key_prefix}_date"] = date.fromisoformat(str(editing_purchase["invoice_date"])[:10])
			except Exception:
				ui.session_state[f"{key_prefix}_date"] = date.today()
			ui.session_state[f"{key_prefix}_no"] = str(editing_purchase.get("invoice_no") or "")
			ui.session_state[f"{key_prefix}_sup"] = str(editing_purchase.get("supplier_name") or "")
			if editing_purchase.get("category") in PURCHASE_CATEGORIES:
				ui.session_state[f"{key_prefix}_cat"] = editing_purchase["category"]
			ui.session_state[f"{key_prefix}_st"] = (
				editing_purchase.get("payment_status")
				if editing_purchase.get("payment_status") in ("Ödendi", "Vadeli / Bekliyor")
				else "Ödendi"
			)
			ui.session_state[f"{key_prefix}_notes"] = str(editing_purchase.get("notes") or "")
			ui.session_state[f"{key_prefix}_items_draft"] = [dict(r) for r in (editing_purchase.get("items") or [])]
		else:
			ui.session_state[f"{key_prefix}_no"] = ""
			ui.session_state[f"{key_prefix}_sup"] = ""
			ui.session_state[f"{key_prefix}_notes"] = ""
			ui.session_state[f"{key_prefix}_st"] = "Ödendi"
			ui.session_state.pop(f"{key_prefix}_items_draft", None)

	fc1, fc2, fc3, fc4, fc5 = ui.columns([1.1, 1.0, 1.4, 1.3, 1.0])
	inv_dt = fc1.date_input("Fatura Tarihi", value=date.today(), key=f"{key_prefix}_date")
	inv_no = fc2.text_input("Fatura / Fiş No", key=f"{key_prefix}_no", placeholder="Örn: FTR-2026-041")
	inv_sup = fc3.text_input("Tedarikçi / Firma", key=f"{key_prefix}_sup", placeholder="Örn: Merkez Hal / Kasap")
	inv_cat = fc4.selectbox("Fatura Kategorisi", PURCHASE_CATEGORIES, key=f"{key_prefix}_cat")
	inv_st = fc5.selectbox("Ödeme Durumu", ["Ödendi", "Vadeli / Bekliyor"], key=f"{key_prefix}_st")

	draft_pi = ui.session_state.get(f"{key_prefix}_items_draft") or [
		{"item_name": "", "category": inv_cat, "quantity": 1.0, "unit_type": "kg", "unit_price": 0.0, "vat_rate": 1.0, "expiry_date": ""}
	]
	df_pi = pd.DataFrame([
		{
			"Ürün Adı": str(r.get("item_name") or ""),
			"Kategori": r.get("category") if r.get("category") in PURCHASE_CATEGORIES else inv_cat,
			"Miktar": _safe_float(r.get("quantity"), 1.0),
			"Miktar Çeşidi": r.get("unit_type") if r.get("unit_type") in PURCHASE_UNITS else "kg",
			"Birim Fiyat (₺)": _safe_float(r.get("unit_price"), 0.0),
			"KDV (%)": _safe_float(r.get("vat_rate"), 1.0),
			"SKT (Ops.)": str(r.get("expiry_date") or ""),
		}
		for r in draft_pi
	])
	ui.caption("💡 Ürün adı, miktar, birim, birim fiyat veya KDV hücrelerini tıklayarak düzeltebilir; tablonun altından (`+`) veya **➕ Ürün Satırı Ekle** düğmesiyle yeni kalem ekleyip satır silebilirsiniz.")
	ed_pi = ui.data_editor(
		df_pi,
		num_rows="dynamic",
		hide_index=True,
		use_container_width=True,
		key=f"{key_prefix}_items_editor",
		column_config={
			"Ürün Adı": ui.column_config.TextColumn("Ürün Adı", required=True),
			"Kategori": ui.column_config.SelectboxColumn("Kategori", options=PURCHASE_CATEGORIES),
			"Miktar": ui.column_config.NumberColumn("Miktar", min_value=0.0, step=1.0),
			"Miktar Çeşidi": ui.column_config.SelectboxColumn("Miktar Çeşidi / Birim", options=PURCHASE_UNITS),
			"Birim Fiyat (₺)": ui.column_config.NumberColumn("Birim Fiyat (₺)", min_value=0.0, step=1.0),
			"KDV (%)": ui.column_config.NumberColumn("KDV (%)", min_value=0.0, step=1.0),
			"SKT (Ops.)": ui.column_config.TextColumn("SKT (YYYY-AA-GG)"),
		},
	)
	inv_notes = ui.text_input("Fatura Notu (Opsiyonel)", key=f"{key_prefix}_notes")

	live_rows = ed_pi.to_dict("records")
	sub_v, vat_v, tot_v, cnt_v = _calc_live_invoice_totals(live_rows)
	m1, m2, m3, m4 = ui.columns(4)
	m1.metric("Geçerli Kalem", f"{cnt_v} ürün")
	m2.metric("Ara Toplam", f"{sub_v:,.2f} ₺")
	m3.metric("Toplam KDV", f"{vat_v:,.2f} ₺")
	m4.metric("Fatura Genel Toplamı", f"{tot_v:,.2f} ₺")

	if in_dialog:
		ab1, ab2, ab3, ab4 = ui.columns([1.2, 1.5, 0.9, 0.8], vertical_alignment="center")
	else:
		ab1, ab2 = ui.columns([1.1, 2.2], vertical_alignment="center")

	with ab1:
		if ui.button("➕ Yeni Ürün Satırı Ekle", key=f"{key_prefix}_btn_add_row", use_container_width=True):
			cur_draft = [
				{
					"item_name": str(x.get("Ürün Adı") or ""),
					"category": x.get("Kategori") or inv_cat,
					"quantity": _safe_float(x.get("Miktar"), 1.0),
					"unit_type": x.get("Miktar Çeşidi") or "kg",
					"unit_price": _safe_float(x.get("Birim Fiyat (₺)"), 0.0),
					"vat_rate": _safe_float(x.get("KDV (%)"), 1.0),
					"expiry_date": str(x.get("SKT (Ops.)") or ""),
				}
				for x in live_rows
			]
			cur_draft.append({"item_name": "", "category": inv_cat, "quantity": 1.0, "unit_type": "kg", "unit_price": 0.0, "vat_rate": 1.0, "expiry_date": ""})
			ui.session_state[f"{key_prefix}_items_draft"] = cur_draft
			ui.session_state.pop(f"{key_prefix}_items_editor", None)
			ui.rerun(scope="fragment" if in_dialog else "app")

	with ab2:
		btn_lbl_p = "💾 Fatura Değişikliklerini Güncelle" if editing_pid else "💾 Faturayı ve Ürün Kalemlerini Kaydet"
		if ui.button(btn_lbl_p, type="primary", use_container_width=True, key=f"{key_prefix}_btn_save_inv"):
			valid_items = [x for x in live_rows if str(x.get("Ürün Adı") or "").strip()]
			if not valid_items:
				ui.warning("Lütfen tabloya en az 1 ürün satırı girin.")
			else:
				pid = upsert_purchase_invoice(
					connect_db_fn,
					purchase_id=editing_pid,
					invoice_date=inv_dt.isoformat(),
					invoice_no=inv_no,
					supplier_name=inv_sup,
					category=inv_cat,
					items=valid_items,
					payment_status=inv_st,
					notes=inv_notes,
				)
				ui.session_state.pop(f"{key_prefix}_items_draft", None)
				ui.session_state.pop(f"{key_prefix}_items_editor", None)
				ui.session_state[sig_key] = None
				ui.session_state["caf_inv_loaded_sig"] = None
				ui.session_state["caf_inv_toast"] = f"✅ Fatura (#{pid}) ve {len(valid_items)} ürün kalemi güncellendi."
				if in_dialog:
					ui.close_dialog(all_dialogs=True)
				ui.rerun()

	if in_dialog and editing_pid:
		with ab3:
			if ui.button("🗑️ Faturayı Sil", key=f"{key_prefix}_btn_del_inv", use_container_width=True):
				delete_purchase_by_id(connect_db_fn, editing_pid)
				ui.session_state[sig_key] = None
				ui.session_state["caf_inv_loaded_sig"] = None
				ui.session_state["caf_inv_toast"] = f"🗑️ Fatura (#{editing_pid}) ve bağlı kalemleri silindi."
				ui.close_dialog(all_dialogs=True)
				ui.rerun()
		with ab4:
			if ui.button("✖ Kapat", key=f"{key_prefix}_btn_close_inv", use_container_width=True):
				ui.close_dialog(all_dialogs=True)


@ui.dialog("🧾 Fatura Detayı, Ürünler & Fiyat Düzenleyici", width="large")
def open_invoice_editor_dialog(ctx_api: dict[str, Any], purchase_id: int) -> None:
	"""Tablodan tıklanan faturanın içindeki ürün/fiyat listesini görüntüleme ve düzenleme pop-up penceresi."""
	target = get_purchase_by_id(ctx_api["connect_db"], int(purchase_id))
	if not target:
		ui.warning("Seçilen fatura bulunamadı.")
		if ui.button("Kapat", key="caf_dlg_inv_missing_close"):
			ui.close_dialog(all_dialogs=True)
		return
	ui.caption(
		f"Fatura #{int(target['id'])} · Tarih: {target.get('invoice_date') or '—'} · "
		f"Fatura No: {target.get('invoice_no') or '—'} · Kayıtlı Kalem: {len(target.get('items') or [])} · "
		f"Kayıtlı Tutar: {float(target.get('total_amount') or 0):,.2f} ₺"
	)
	_render_invoice_editor_body(ctx_api, editing_purchase=target, key_prefix="caf_dlg_i", in_dialog=True)


def _render_purchase_statistics_section(connect_db_fn: Callable[[], sqlite3.Connection], sel_ym: str) -> None:
	"""Fatura ve alınan ürün verilerinden kullanıcı tarafından seçilebilen istatistik görünümlerini çizer."""
	ui.markdown("---")
	ui.markdown("#### 📊 Fatura ve Alınan Ürün İstatistikleri (Seçilebilir Analizler)")
	ui.caption("Aşağıdan dönem/filtre belirleyip görüntülemek istediğiniz istatistik başlıklarını seçebilirsiniz.")

	sc1, sc2, sc3 = ui.columns([1.15, 1.15, 1.3])
	stat_scope = sc1.selectbox(
		"Analiz Dönemi",
		["Tüm Tarihler", f"Seçili Ay ({_fmt_ym(sel_ym)})", "Özel Tarih Aralığı"],
		key="caf_stat_scope",
	)
	stat_cat = sc2.selectbox("Kategori", ["Tümü"] + PURCHASE_CATEGORIES, key="caf_stat_cat")
	stat_q = sc3.text_input("Ürün veya Tedarikçi Filtresi", key="caf_stat_q")

	sd_arg: str | None = None
	ed_arg: str | None = None
	ym_arg = sel_ym if "Seçili Ay" in stat_scope else "Tümü"
	if stat_scope == "Özel Tarih Aralığı":
		ym_arg = "Tümü"
		dc1, dc2 = ui.columns(2)
		sd_d = dc1.date_input("Başlangıç Tarihi", value=date.today().replace(day=1), key="caf_stat_sd")
		ed_d = dc2.date_input("Bitiş Tarihi", value=date.today(), key="caf_stat_ed")
		sd_arg = sd_d.isoformat()
		ed_arg = ed_d.isoformat()

	selected_views = ui.multiselect(
		"Görüntülenecek İstatistikler",
		PURCHASE_STAT_VIEWS,
		default=PURCHASE_STAT_VIEWS[:4],
		key="caf_stat_views",
	)
	sel_prod_name = str(ui.session_state.get("caf_stat_prod_pick") or "")
	analytics = compute_purchase_analytics(
		connect_db_fn,
		year_month=ym_arg,
		start_date=sd_arg,
		end_date=ed_arg,
		category=stat_cat,
		query_text=stat_q,
		product_name=sel_prod_name,
	)
	sm = analytics["summary"]
	if sm["invoice_count"] == 0:
		ui.info("Seçilen analiz döneminde henüz fatura verisi bulunmuyor.")
		return

	sk1, sk2, sk3, sk4, sk5 = ui.columns(5)
	sk1.metric("Fatura Sayısı", f"{sm['invoice_count']}")
	sk2.metric("Toplam Kalem", f"{sm['item_row_count']}")
	sk3.metric("Farklı Ürün Çeşidi", f"{sm['distinct_products']}")
	sk4.metric("Toplam Harcama", f"{sm['total_spend_tl']:,.2f} ₺")
	sk5.metric("Ortalama Fatura", f"{sm['avg_invoice_tl']:,.2f} ₺")
	ui.caption(
		f"🏆 **En Yüksek Harcamalı Ürün:** `{sm['top_product_name']}` ({sm['top_product_spend']:,.2f} ₺)  ·  "
		f"🏪 **En Çok Çalışılan Tedarikçi:** `{sm['top_supplier_name']}` ({sm['top_supplier_spend']:,.2f} ₺)"
	)

	if not selected_views:
		ui.caption("Görüntülemek için yukarıdan en az bir istatistik başlığı seçin.")

	if PURCHASE_STAT_VIEWS[0] in selected_views and analytics["product_summary"]:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[0]}")
		df_ps = pd.DataFrame(analytics["product_summary"])
		ui.dataframe(df_ps, hide_index=True, use_container_width=True)
		top_prod_chart = {f"{r['Ürün Adı']} ({r['Birim']})": r["Toplam Harcama (₺)"] for r in analytics["product_summary"][:10]}
		if top_prod_chart:
			ui.bar_chart(top_prod_chart, height=220, color="#0f766e")

	if PURCHASE_STAT_VIEWS[1] in selected_views and analytics["category_summary"]:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[1]}")
		cs1, cs2 = ui.columns([1.3, 1.0], vertical_alignment="center")
		with cs1:
			ui.dataframe(pd.DataFrame(analytics["category_summary"]), hide_index=True, use_container_width=True)
		with cs2:
			cat_chart = {r["Kategori"]: r["Toplam Harcama (₺)"] for r in analytics["category_summary"]}
			if cat_chart:
				ui.bar_chart(cat_chart, height=220, color="#ea580c")

	if PURCHASE_STAT_VIEWS[2] in selected_views and analytics["supplier_summary"]:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[2]}")
		ss1, ss2 = ui.columns([1.3, 1.0], vertical_alignment="center")
		with ss1:
			ui.dataframe(pd.DataFrame(analytics["supplier_summary"]), hide_index=True, use_container_width=True)
		with ss2:
			sup_chart = {r["Tedarikçi / Firma"]: r["Toplam Harcama (₺)"] for r in analytics["supplier_summary"][:8]}
			if sup_chart:
				ui.bar_chart(sup_chart, height=220, color="#2563eb")

	if PURCHASE_STAT_VIEWS[3] in selected_views and analytics["monthly_trend"]:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[3]}")
		ms1, ms2 = ui.columns([1.3, 1.0], vertical_alignment="center")
		with ms1:
			ui.dataframe(pd.DataFrame(analytics["monthly_trend"]), hide_index=True, use_container_width=True)
		with ms2:
			mon_chart = {r["Dönem"]: r["Toplam Harcama (₺)"] for r in analytics["monthly_trend"]}
			if mon_chart:
				ui.bar_chart(mon_chart, height=220, color="#7c3aed")

	if PURCHASE_STAT_VIEWS[4] in selected_views:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[4]}")
		if analytics["price_changes"]:
			ui.dataframe(
				pd.DataFrame([
					{**c, "Değişim (%)": f"{'+' if c['Değişim (%)'] > 0 else ''}{c['Değişim (%)']}%"}
					for c in analytics["price_changes"]
				]),
				hide_index=True,
				use_container_width=True,
			)
		else:
			ui.caption("Birim fiyat değişimi hesaplamak için aynı ürün adına ait en az 2 fatura girişi gerekir.")

	if PURCHASE_STAT_VIEWS[5] in selected_views and analytics["product_options"]:
		ui.markdown(f"##### {PURCHASE_STAT_VIEWS[5]}")
		picked_p = ui.selectbox("Alım ve Fiyat Geçmişi İncelenecek Ürün", analytics["product_options"], key="caf_stat_prod_pick")
		if picked_p != analytics["selected_product"]:
			analytics = compute_purchase_analytics(
				connect_db_fn,
				year_month=ym_arg,
				start_date=sd_arg,
				end_date=ed_arg,
				category=stat_cat,
				query_text=stat_q,
				product_name=picked_p,
			)
		hist_rows = analytics["product_history"]
		if hist_rows:
			ui.dataframe(pd.DataFrame(hist_rows), hide_index=True, use_container_width=True)
			hist_chart = {f"{r['Tarih']} ({r['Fatura No']})": r["Birim Fiyat (₺)"] for r in hist_rows}
			ui.bar_chart(hist_chart, height=200, color="#ca8a04")

	ui.download_button(
		"📥 Seçili Dönem Fatura & Ürün İstatistiklerini İndir (.xlsx)",
		data=lambda: build_purchase_analytics_excel(analytics, stat_scope),
		file_name=f"Yemekhane_Fatura_Istatistikleri_{sel_ym}.xlsx",
		mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		key="caf_dl_purch_analytics_xlsx",
	)


def _render_tab_purchases_and_expenses(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()

	ui.markdown("### 🚚 Gelen Ürünler, Fatura Kayıtları, SKT & Aylık / Toplam Gider Takibi")
	if ui.session_state.get("caf_inv_toast"):
		ui.success(ui.session_state.pop("caf_inv_toast"))

	stats = get_expense_summary_stats(connect_db_fn, sel_ym)

	k1, k2, k3, k4, k5 = ui.columns(5)
	k1.metric(f"{_fmt_ym(sel_ym)} Gideri", f"{stats['month_total']:,.2f} ₺", delta=f"{stats['month_invoice_count']} fatura")
	k2.metric("Vadeli / Bekleyen Borç", f"{stats['month_pending_total']:,.2f} ₺", delta=f"{stats['month_pending_count']} vadeli" if stats["month_pending_count"] else "Borç Yok")
	k3.metric(f"{sel_ym[:4]} Yılı Toplam Gider", f"{stats['year_total']:,.2f} ₺")
	k4.metric("Genel Toplam Gider", f"{stats['all_time_total']:,.2f} ₺")
	top_cat = max(stats["by_category"].items(), key=lambda x: x[1])[0] if stats["by_category"] else "—"
	k5.metric("En Yüksek Gider Kalemi", top_cat)

	# SKT Yaklaşan / Tarihi Geçen Gıda Uyarı Bandı
	expiring = get_expiring_purchase_items(connect_db_fn, within_days=30, include_expired=True)
	if expiring:
		with ui.expander(f"⏳ SKT Takip Uyarısı: {len(expiring)} Gelen Mutfak Ürününde Son Kullanma Tarihi Yaklaşıyor/Geçmiş", expanded=True):
			df_skt = pd.DataFrame([
				{
					"Durum": e["urgency"],
					"SKT": e["expiry_date"],
					"Kalan Gün": e["days_left"],
					"Ürün Adı": e["item_name"],
					"Miktar": f"{e['quantity']} {e['unit_type']}",
					"Tedarikçi": e.get("supplier_name") or "—",
					"Fatura Tarihi": e["invoice_date"],
				}
				for e in expiring
			])
			ui.dataframe(df_skt, hide_index=True, use_container_width=True)

	# AI Fatura / Fiş Okuyucu
	if is_on:
		with ui.expander("📸 Gemini AI Fiş / Fatura / İrsaliye Okuyucu (PDF, Fotoğraf, Excel, Word)"):
			up_inv = ui.file_uploader("Fatura / Fiş Görseli veya Belgesi", type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "docx", "txt"], key="caf_ai_inv_file")
			txt_inv = ui.text_input("Veya Fatura Metnini Yapıştırın", placeholder="Örn: 15.11.2026 Özkan Gıda No:124 - 20 kg pirinç 55 TL/kg, 10 L sıvıyağ 65 TL/L", key="caf_ai_inv_txt")
			if ui.button("✨ Faturayı AI ile Oku (Önizlemeden Onayla)", type="primary", key="caf_btn_ai_parse_inv"):
				fb = up_inv.getvalue() if up_inv is not None else None
				fn = up_inv.name if up_inv is not None else ""
				strict_pdf = str(app_cfg.get("gemini_strict_kvkk_pdf", "0")).strip() == "1"
				with ui.spinner("Gemini AI fatura kalemlerini, miktar çeşitlerini ve fiyatları çıkarıyor..."):
					ok_iv, iv_data, used_iv = gai.parse_cafeteria_invoice_with_gemini(
						api_key=api_key, model=model, file_bytes=fb, file_name=fn, raw_text=txt_inv, strict_kvkk_pdf=strict_pdf
					)
				if ok_iv and iv_data.get("items"):
					ui.session_state["caf_inv_picker"] = "➕ Yeni Fatura Girişi"
					ui.session_state["caf_inv_loaded_sig"] = "new"
					ui.session_state.pop("caf_inv_items_editor", None)
					ui.session_state["caf_inv_date"] = date.fromisoformat(str(iv_data.get("invoice_date") or date.today().isoformat())[:10]) if re.match(r"^\d{4}-\d{2}-\d{2}$", str(iv_data.get("invoice_date") or "")) else date.today()
					ui.session_state["caf_inv_no"] = str(iv_data.get("invoice_no") or "")
					ui.session_state["caf_inv_sup"] = str(iv_data.get("supplier_name") or "")
					if iv_data.get("category") in PURCHASE_CATEGORIES:
						ui.session_state["caf_inv_cat"] = iv_data["category"]
					ui.session_state["caf_inv_items_draft"] = iv_data["items"]
					ui.success(f"✅ {len(iv_data['items'])} ürün kalemi Gemini ({used_iv}) ile aşağıdaki forma aktarıldı! Kontrol edip Kaydet'e basın.")
					ui.rerun()
				else:
					ui.error(f"Fatura okunamadı: {(iv_data or {}).get('error', 'Kalem bulunamadı')}")

	# Yeni veya Mevcut Fatura Seçimi / Düzenleme
	all_purchases_for_edit, _ = get_purchases_and_items(connect_db_fn, year_month="Tümü")
	edit_p_map = {
		f"#{pr['id']} — {pr['invoice_date']} · {pr.get('supplier_name') or 'Tedarikçi Yok'} · {float(pr.get('total_amount') or 0):,.2f} ₺ ({pr['payment_status']})": int(pr["id"])
		for pr in all_purchases_for_edit[:80]
	}
	ep1, ep2, ep3 = ui.columns([2.0, 1.0, 1.0], vertical_alignment="center")
	with ep1:
		picked_p_lbl = ui.selectbox("Fatura İşlemi (Yeni Giriş veya Mevcut Fatura Düzenle)", ["➕ Yeni Fatura Girişi"] + list(edit_p_map.keys()), key="caf_inv_picker")
	editing_pid = edit_p_map.get(picked_p_lbl)
	editing_p_obj = get_purchase_by_id(connect_db_fn, editing_pid) if editing_pid else None
	with ep2:
		if editing_p_obj and ui.button("🪟 Faturayı Pop-up'ta Aç", type="secondary", use_container_width=True, key="caf_btn_open_popup_inv"):
			open_invoice_editor_dialog(ctx_api, int(editing_pid))
	with ep3:
		if editing_p_obj and editing_p_obj.get("payment_status") == "Vadeli / Bekliyor":
			if ui.button("✅ Vadeli Faturayı Ödendi İşaretle", type="primary", use_container_width=True, key="caf_btn_mark_paid"):
				update_purchase_payment_status(connect_db_fn, int(editing_pid), "Ödendi")
				ui.session_state["caf_inv_loaded_sig"] = None
				ui.success(f"✅ Fatura (#{editing_pid}) 'Ödendi' olarak işaretlendi.")
				ui.rerun()

	with ui.container(border=True):
		ui.markdown("##### ✏️ Fatura / İrsaliye Bilgileri ve Ürün Kalemleri" if editing_p_obj else "##### ➕ Yeni Gelen Ürün / Fatura Girişi")
		_render_invoice_editor_body(ctx_api, editing_purchase=editing_p_obj, key_prefix="caf_inv", in_dialog=False)

	# Fatura Tarihine Göre Filtrelenmiş Faturalar & Ürünler
	ui.markdown("---")
	fl1, fl2, fl3 = ui.columns([1.2, 1.3, 1.5])
	scope_opt = fl1.selectbox("Dönem Filtresi", [f"Seçili Ay ({_fmt_ym(sel_ym)})", "Tüm Tarihler"], key="caf_purch_scope_flt")
	cat_flt = fl2.selectbox("Kategori Filtresi", ["Tümü"] + PURCHASE_CATEGORIES, key="caf_purch_cat_flt")
	q_flt = fl3.text_input("Ürün / Tedarikçi / Fatura No Ara", key="caf_purch_q_flt")

	ym_arg = sel_ym if "Seçili Ay" in scope_opt else "Tümü"
	purchases, items = get_purchases_and_items(connect_db_fn, year_month=ym_arg, category=cat_flt, query_text=q_flt)

	if purchases:
		ui.markdown(f"##### 🧾 Kayıtlı Faturalar Listesi ({len(purchases)} fatura · Görüntülemek & Düzenlemek İçin Faturaya Tıklayın)")
		ui.caption("💡 Fatura satırına tıkladığınızda o faturadaki tüm alınan ürünler ve fiyatları pop-up pencerede açılır; ürün ekleyip çıkarabilir veya isim/fiyat düzeltebilirsiniz.")
		df_show_invs = pd.DataFrame([
			{
				"id": int(pr["id"]),
				"Fatura Tarihi": pr["invoice_date"],
				"Fatura No": pr.get("invoice_no") or "—",
				"Tedarikçi": pr.get("supplier_name") or "—",
				"Kategori": pr.get("category") or "Diğer",
				"Kalem Sayısı": int(pr.get("item_count") or len(pr.get("items") or [])),
				"Alınan Ürünler & Fiyatlar": pr.get("items_summary") or "—",
				"Toplam Tutar (₺)": round(_safe_float(pr.get("total_amount"), 0.0), 2),
				"Durum": pr.get("payment_status") or "Ödendi",
			}
			for pr in purchases
		])
		tbl_inv_ev = ui.dataframe(
			df_show_invs,
			hide_index=True,
			use_container_width=True,
			on_select="rerun",
			selection_mode=["single-row", "single-cell"],
			key="caf_invoices_list_tbl",
			column_config={
				"id": None,
				"Alınan Ürünler & Fiyatlar": ui.column_config.TextColumn("Alınan Ürünler & Fiyatlar", help="Fatura içindeki ürünleri görüntülemek ve düzenlemek için satıra tıklayın"),
			},
		)
		inv_rows_sel = getattr(tbl_inv_ev.selection, "rows", []) or []
		inv_cells_sel = getattr(tbl_inv_ev.selection, "cells", []) or []
		clicked_inv_row = None
		if inv_rows_sel:
			try:
				r0 = int(inv_rows_sel[0])
				if 0 <= r0 < len(df_show_invs):
					clicked_inv_row = r0
			except Exception:
				pass
		if clicked_inv_row is None:
			for c in inv_cells_sel:
				c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
				if c_row is not None and 0 <= int(c_row) < len(df_show_invs):
					clicked_inv_row = int(c_row)
					break
		if clicked_inv_row is not None:
			clicked_pid = int(df_show_invs.iloc[clicked_inv_row]["id"])
			ui.clear_dataframe_selection("caf_invoices_list_tbl")
			open_invoice_editor_dialog(ctx_api, clicked_pid)

	if items:
		df_show_items = pd.DataFrame([
			{
				"Fatura Tarihi": it["invoice_date"],
				"Fatura No": it.get("invoice_no") or "—",
				"Tedarikçi": it.get("supplier_name") or "—",
				"Ürün Adı": it["item_name"],
				"Kategori": it["category"],
				"Miktar": it["quantity"],
				"Çeşidi": it["unit_type"],
				"Birim Fiyat (₺)": it["unit_price"],
				"KDV (%)": it["vat_rate"],
				"SKT": it.get("expiry_date") or "—",
				"Toplam Tutar (₺)": it["line_total"],
			}
			for it in items
		])
		ui.markdown(f"##### 📋 Fatura Tarihine Göre Gelen Ürün Kalemleri ({len(items)} kalem · Filtrelenmiş Toplam: **{sum(float(i['line_total'] or 0) for i in items):,.2f} ₺**)")
		ui.dataframe(df_show_items, hide_index=True, use_container_width=True)
		_exp_buf = io.BytesIO()
		df_show_items.to_excel(_exp_buf, index=False, engine="openpyxl")
		ui.download_button(
			"📥 Filtrelenen Ürün / Gider Listesini İndir (.xlsx)",
			data=_exp_buf.getvalue(),
			file_name=f"Yemekhane_Giderler_{sel_ym}.xlsx",
			mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			key="caf_dl_filtered_expenses_xlsx",
		)
	else:
		ui.info("Seçilen filtrede gelen ürün / fatura kalemi bulunamadı.")

	# Seçilebilir Fatura ve Alınan Ürün İstatistikleri Bölümü
	_render_purchase_statistics_section(connect_db_fn, sel_ym)

	# Menüye Göre Satın Alma İhtiyaç Listesi
	ui.markdown("---")
	with ui.expander("🧺 Planlı Menüye Göre Otomatik Market / Malzeme İhtiyaç Listesi Hesaplayıcı", expanded=False):
		sh1, sh2 = ui.columns(2)
		sh_scope = sh1.radio("İhtiyaç Kapsamı", ["Tüm Ay Menüsü", "1. Hafta Menüsü"], horizontal=True, key="caf_shop_scope")
		cur_diners_cnt = max(1, len(get_diners(connect_db_fn, active_only=True)))
		sh_cnt = sh2.number_input("Hesaplanacak Günlük Kişi / Porsiyon Sayısı", min_value=1, step=5, value=cur_diners_cnt, key="caf_shop_cnt")
		shop_res = build_menu_shopping_list(
			connect_db_fn,
			sel_ym,
			scope="weekly" if "Hafta" in sh_scope else "monthly",
			week_index=0,
			diner_count=int(sh_cnt),
		)
		if shop_res["items"]:
			ui.metric(f"{shop_res['diner_count']} Kişilik Tahmini Malzeme Maliyeti", f"{shop_res['estimated_total_tl']:,.2f} ₺")
			ui.dataframe(pd.DataFrame(shop_res["items"]), hide_index=True, use_container_width=True)
			ui.download_button(
				"📥 İhtiyaç / Sipariş Listesini İndir (.xlsx)",
				data=lambda: build_shopping_list_excel(shop_res, sel_ym),
				file_name=f"Yemekhane_Ihtiyac_Listesi_{sel_ym}.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key="caf_dl_shop_xlsx",
			)
		else:
			ui.caption("Seçili dönem menüsünde malzemeli kayıtlı yemek bulunamadı.")

	if purchases:
		with ui.expander(f"🗑️ Fatura Kaydı Sil ({len(purchases)} fatura)"):
			p_opts = {
				f"#{pr['id']} — {pr['invoice_date']} · {pr.get('supplier_name') or 'Tedarikçi Yok'} · {float(pr.get('total_amount') or 0):,.2f} ₺": int(pr["id"])
				for pr in purchases
			}
			sel_del_p = ui.selectbox("Silinecek Fatura", list(p_opts.keys()), key="caf_del_p_sel")
			if ui.button("Seçili Faturayı Sil", key="caf_del_p_btn", type="secondary"):
				delete_purchase_by_id(connect_db_fn, p_opts[sel_del_p])
				ui.session_state["caf_inv_loaded_sig"] = None
				ui.success("Fatura ve bağlı ürün kalemleri silindi.")
				ui.rerun()


def _render_tab_diners_and_cash(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api["mask_kvkk_value"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()

	ui.markdown("### 👥 Yemek Yiyenler, Aylık Öğün Ücreti Hesabı & Kasa Muhasebesi")
	cash = get_cash_summary_and_diner_balances(connect_db_fn, sel_ym)

	# Üst Kasa & Tahakkuk KPI Kartları
	c1, c2, c3, c4, c5 = ui.columns(5)
	c1.metric("Aylık Beklenen", f"{cash['month_expected']:,.2f} ₺")
	c2.metric("Ay İçi Tahsilat (+)", f"{cash['month_collected']:,.2f} ₺")
	c3.metric("Kalan Aidat", f"{cash['month_remaining']:,.2f} ₺")
	c4.metric("Okula Teslim (-)", f"{cash['vault_total_handover']:,.2f} ₺")
	c5.metric("💵 Şu An Kasadaki Para", f"{cash['vault_balance']:,.2f} ₺", delta=f"Toplam Giren: {cash['vault_total_in']:,.2f} ₺")

	# 1. Yemek Yiyen Ekleme, Tahakkuk, Kasa & Bütçe Alt Sekmeleri
	t_diners, t_fees, t_cash, t_budget = ui.tabs([
		"🧑‍🎓 Yemek Yiyen Kişiler",
		"🧮 İş Günü × Öğün Ücreti & Devir",
		"💵 Ücret Tahsilatı (Getiren İsim) & Kasa",
		"📊 Gelir–Gider Bütçe Dengesi & AI",
	])

	diners = get_diners(connect_db_fn, active_only=True)
	all_diners_inc_inactive = get_diners(connect_db_fn, active_only=False)
	inactive_diners = [d for d in all_diners_inc_inactive if not d.get("is_active")]

	with t_diners:
		with ui.expander("➕ Revirden Öğrenci/Personel Bağla veya Yeni Yemek Yiyen Ekle", expanded=not bool(diners)):
			ac1, ac2 = ui.columns(2)
			with ac1:
				with connect_db_fn() as conn:
					stu_all = conn.execute(
						"SELECT s.id, s.full_name, COALESCE(c.name, '—') AS cls FROM students s LEFT JOIN classes c ON c.id=s.class_id ORDER BY c.name, s.full_name"
					).fetchall()
				linked_sids = {int(d["student_id"]) for d in all_diners_inc_inactive if d.get("student_id")}
				link_opts = {
					f"{mask_fn(r['full_name'], 'name')} · {r['cls']}": int(r["id"])
					for r in stu_all
					if int(r["id"]) not in linked_sids
				}
				def_meals_cnt = parse_service_schedule_from_cfg(app_cfg)["meals_per_day"]
				picked_links = ui.multiselect("Kayıtlı Öğrenci / Personel Seç", list(link_opts.keys()), key="caf_link_stu_ms")
				link_meals = ui.number_input("Günlük Öğün Sayısı", min_value=1, max_value=3, value=def_meals_cnt, key="caf_link_meals")
				if ui.button("📥 Seçilenleri Yemek Listesine Ekle", type="primary", use_container_width=True, disabled=not picked_links, key="caf_btn_link_stu"):
					cnt_add = import_school_people_as_diners(connect_db_fn, [link_opts[k] for k in picked_links], daily_meals_count=int(link_meals))
					ui.success(f"✅ {cnt_add} kişi yemek yiyenler listesine eklendi.")
					ui.rerun()
			with ac2:
				with ui.form("caf_form_add_diner", clear_on_submit=True):
					fn_new = ui.text_input("Serbest Kişi Adı Soyadı*")
					r_new = ui.selectbox("Statü", DINER_ROLE_TYPES)
					d_new = ui.text_input("Sınıf / Bölüm", placeholder="Örn: 3-A veya Mutfak")
					m_new = ui.number_input("Günlük Öğün", min_value=1, max_value=3, value=def_meals_cnt)
					ex_new = ui.checkbox("Ücretsiz / Burslu")
					if ui.form_submit_button("➕ Listeye Ekle", use_container_width=True):
						if not clean_name_valid(fn_new):
							ui.warning("Ad soyad yazın.")
						else:
							upsert_diner(connect_db_fn, full_name=fn_new, role_type=r_new, class_or_dept=d_new, daily_meals_count=int(m_new), fee_exempt=ex_new)
							ui.success(f"✅ {fn_new} yemek listesine eklendi.")
							ui.rerun()

		if diners:
			df_din = pd.DataFrame([
				{
					"Ad Soyad": mask_fn(d["full_name"], "name"),
					"Statü": d["role_type"],
					"Sınıf/Birim": d.get("class_or_dept") or "—",
					"Günlük Öğün": d["daily_meals_count"],
					"Ücret Muaf": "Evet" if d.get("fee_exempt") else "Hayır",
					"Bilinen Alerji": d.get("revir_allergies") or d.get("allergies_override") or "—",
				}
				for d in diners
			])
			ui.dataframe(df_din, hide_index=True, use_container_width=True)
			with ui.expander("⏸️ Aboneyi Pasife Al (Geçmişi Korur) / Kalıcı Sil"):
				d_del_map = {f"#{d['id']} — {mask_fn(d['full_name'], 'name')} ({d.get('class_or_dept') or d['role_type']})": int(d["id"]) for d in diners}
				sel_dd = ui.selectbox("İşlem Yapılacak Abone", list(d_del_map.keys()), key="caf_del_diner_sel")
				d_act1, d_act2 = ui.columns(2)
				with d_act1:
					if ui.button("⏸️ Aboneliği Pasife Al (Önerilen)", type="primary", use_container_width=True, key="caf_deact_diner_btn"):
						deactivate_diner_by_id(connect_db_fn, d_del_map[sel_dd])
						ui.success("Abone pasife alındı; önceki ay tahakkuk ve tahsilat geçmişi korundu.")
						ui.rerun()
				with d_act2:
					if ui.button("🗑️ Kaydı Kalıcı Sil (Hareketsizler)", type="secondary", use_container_width=True, key="caf_del_diner_btn"):
						try:
							delete_diner_by_id(connect_db_fn, d_del_map[sel_dd], force=False)
							ui.success("Abone kalıcı olarak silindi.")
							ui.rerun()
						except Exception as exc:
							ui.warning(str(exc))

		if inactive_diners:
			with ui.expander(f"♻️ Pasif / Ayrılan Aboneler ({len(inactive_diners)})"):
				in_opts = {f"#{d['id']} — {mask_fn(d['full_name'], 'name')} (Bitiş: {d.get('end_date') or '—'})": int(d["id"]) for d in inactive_diners}
				sel_react = ui.selectbox("Tekrar Aktifleştirilecek Kişi", list(in_opts.keys()), key="caf_react_sel")
				if ui.button("▶ Tekrar Aktif Listeye Al", key="caf_react_btn"):
					reactivate_diner_by_id(connect_db_fn, in_opts[sel_react])
					ui.success("Abone yeniden aktifleştirildi.")
					ui.rerun()

	with t_fees:
		days_info = get_month_workdays_info(connect_db_fn, sel_ym)
		auto_wd = sum(1 for d in days_info if d.get("is_workday"))
		fc1, fc2, fc3, fc4, fc5 = ui.columns([1.2, 1.2, 1.0, 1.2, 1.2])
		stu_price = fc1.number_input("Öğrenci 1 Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_meal_price"), 75.0), key="caf_fee_stu_p")
		stf_price = fc2.number_input("Öğretmen/Personel Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_staff_price"), 90.0), key="caf_fee_stf_p")
		wd_override = fc3.number_input("Ayın İş Günü Sayısı", min_value=0, max_value=31, value=auto_wd, key="caf_fee_wd_inp")
		with fc4:
			ui.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
			if ui.button("🧮 Ücretleri Hesapla", type="primary", use_container_width=True, key="caf_btn_calc_fees"):
				n_f = generate_or_refresh_monthly_fees(
					connect_db_fn,
					sel_ym,
					student_meal_price=float(stu_price),
					staff_meal_price=float(stf_price),
					workdays_override=int(wd_override),
				)
				ui.success(f"✅ {n_f} kişi için {int(wd_override)} iş günü üzerinden {_fmt_ym(sel_ym)} ücretleri hesaplandı.")
				ui.rerun()
		with fc5:
			ui.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
			if ui.button("📲 Önceki Aydan Devir", type="secondary", use_container_width=True, key="caf_btn_carry_fwd", disabled=not bool(cash["balances"])):
				yy_f, mm_f = int(sel_ym[:4]), int(sel_ym[5:7])
				py_f, pm_f = (yy_f - 1, 12) if mm_f == 1 else (yy_f, mm_f - 1)
				prev_ym_s = f"{py_f:04d}-{pm_f:02d}"
				cf_n = carry_forward_unpaid_balances(connect_db_fn, prev_ym_s, sel_ym)
				ui.success(f"✅ {_fmt_ym(prev_ym_s)} döneminden {cf_n} kişinin kalan bakiyesi devredildi.")
				ui.rerun()

		if cash["balances"]:
			df_bal = pd.DataFrame([
				{
					"fee_id": int(b["id"]),
					"Ad Soyad": mask_fn(b["full_name"], "name"),
					"Statü": b["role_type"],
					"Sınıf/Birim": b.get("class_or_dept") or "—",
					"Planlı Gün": int(b["planned_days"]),
					"Günlük Öğün": int(b["meals_per_day"]),
					"Öğün (₺)": float(b["unit_meal_price"]),
					"Devreden (₺)": float(b.get("carry_forward_debt") or 0.0),
					"İndirim (₺)": float(b["discount_or_adjustment"]),
					"Beklenen (₺)": float(b["expected_amount"]),
					"Ödenen (₺)": float(b["paid_amount"]),
					"Kalan (₺)": float(b["remaining_amount"]),
					"Son Getiren": mask_fn(b.get("last_delivered_by") or "—", "name") if b.get("last_delivered_by") else "—",
					"Durum": b["status_label"],
				}
				for b in cash["balances"]
			])
			ed_bal = ui.data_editor(
				df_bal,
				hide_index=True,
				use_container_width=True,
				key="caf_fees_editor",
				column_config={
					"fee_id": None,
					"Ad Soyad": ui.column_config.TextColumn("Ad Soyad", disabled=True),
					"Statü": ui.column_config.TextColumn("Statü", disabled=True),
					"Sınıf/Birim": ui.column_config.TextColumn("Sınıf/Birim", disabled=True),
					"Planlı Gün": ui.column_config.NumberColumn("Katılım Günü", min_value=0, max_value=31, step=1),
					"Günlük Öğün": ui.column_config.NumberColumn("Öğün", min_value=1, max_value=3, step=1),
					"Öğün (₺)": ui.column_config.NumberColumn("Öğün (₺)", min_value=0.0, step=5.0),
					"Devreden (₺)": ui.column_config.NumberColumn("Devreden (₺)", step=10.0),
					"İndirim (₺)": ui.column_config.NumberColumn("İndirim (₺)", step=10.0),
					"Beklenen (₺)": ui.column_config.NumberColumn("Beklenen (₺)", disabled=True),
					"Ödenen (₺)": ui.column_config.NumberColumn("Ödenen (₺)", disabled=True),
					"Kalan (₺)": ui.column_config.NumberColumn("Kalan (₺)", disabled=True),
					"Son Getiren": ui.column_config.TextColumn("Son Getiren", disabled=True),
					"Durum": ui.column_config.TextColumn("Durum", disabled=True),
				},
			)
			if ui.button("💾 Gün / Devir / Ücret Düzeltmelerini Kaydet", type="secondary", use_container_width=True, key="caf_btn_save_fee_edits"):
				for _, r_f in ed_bal.iterrows():
					update_monthly_fee_row(
						connect_db_fn,
						int(r_f["fee_id"]),
						planned_days=int(r_f["Planlı Gün"]),
						meals_per_day=int(r_f["Günlük Öğün"]),
						unit_meal_price=float(r_f["Öğün (₺)"]),
						discount_or_adjustment=float(r_f["İndirim (₺)"]),
						carry_forward_debt=float(r_f["Devreden (₺)"]),
					)
				ui.success("✅ Kişi bazlı katılım günü, devir ve ücret düzeltmeleri güncellendi.")
				ui.rerun()
		else:
			ui.info(f"{_fmt_ym(sel_ym)} için henüz aylık aidat oluşturulmamış. Önce kişi ekleyin ve '🧮 Ücretleri Hesapla' butonuna basın.")

	with t_cash:
		pay_col, out_col = ui.columns(2)
		diner_opts = {f"{mask_fn(d['full_name'], 'name')} ({d.get('class_or_dept') or d['role_type']})": int(d["id"]) for d in diners}
		with pay_col:
			with ui.container(border=True):
				ui.markdown("##### 💰 Ücret Tahsilatı Al (Kasaya Giren)")
				with ui.form("caf_form_in_pay", clear_on_submit=True):
					p_diner_lbl = ui.selectbox("Kimin Adına Ödeniyor*", list(diner_opts.keys()) or ["— Önce Kişi Ekleyin —"], key="caf_pay_diner")
					p_deliv = ui.text_input("Ücreti Getiren İsim*", placeholder="Örn: Veli Fatma Demir veya Öğrt. Ayşe", key="caf_pay_deliv")
					pc1, pc2, pc3 = ui.columns(3)
					p_amt = pc1.number_input("Verilen Miktar (₺)*", min_value=0.0, step=50.0, value=0.0, key="caf_pay_amt")
					p_meth = pc2.selectbox("Yöntem", ["Nakit", "IBAN / Havale", "Kart"], key="caf_pay_meth")
					p_dt = pc3.date_input("Tarih", value=date.today(), key="caf_pay_dt")
					p_note = ui.text_input("Makbuz No / Not", key="caf_pay_note")
					if ui.form_submit_button("✅ Tahsilatı Kasaya İşle", type="primary", use_container_width=True, disabled=not diner_opts):
						try:
							did_p = diner_opts.get(p_diner_lbl)
							record_cash_transaction(
								connect_db_fn,
								tx_date=p_dt.isoformat(),
								tx_type="IN_PAYMENT",
								amount=float(p_amt),
								diner_id=did_p,
								delivered_by_name=p_deliv,
								payment_method=p_meth,
								notes=p_note,
								year_month=sel_ym,
							)
							ui.success(f"✅ {float(p_amt):,.2f} ₺ tahsilat kasaya kaydedildi.")
							ui.rerun()
						except Exception as exc:
							ui.error(str(exc))

		with out_col:
			with ui.container(border=True):
				ui.markdown("##### 🏛️ Okula Teslim / Kasadan Çıkan Para")
				with ui.form("caf_form_out_hand", clear_on_submit=True):
					o_kind = ui.selectbox("Çıkış Türü", ["Okula Teslim Edilen (OUT_SCHOOL_HANDOVER)", "İade / Kasa Harcaması (OUT_REFUND_OR_PETTY)"], key="caf_out_kind")
					o_recv = ui.text_input("Teslim Alan Yetkili / Açıklama*", placeholder="Örn: Okul Müdürü / Müd. Yrd.", key="caf_out_recv")
					oc1, oc2, oc3 = ui.columns(3)
					o_amt = oc1.number_input("Teslim / Çıkan Tutar (₺)*", min_value=0.0, step=100.0, value=0.0, key="caf_out_amt")
					o_rec_no = oc2.text_input("Tutanak / Makbuz No", key="caf_out_rec")
					o_dt = oc3.date_input("Tarih", value=date.today(), key="caf_out_dt")
					o_note = ui.text_input("Not", key="caf_out_note")
					if ui.form_submit_button("📤 Kasadan Çıkışı Kaydet", type="secondary", use_container_width=True):
						try:
							t_code = "OUT_SCHOOL_HANDOVER" if "OUT_SCHOOL_HANDOVER" in o_kind else "OUT_REFUND_OR_PETTY"
							record_cash_transaction(
								connect_db_fn,
								tx_date=o_dt.isoformat(),
								tx_type=t_code,
								amount=float(o_amt),
								for_person_name="Okul İdaresi / Kasa Teslimi" if t_code == "OUT_SCHOOL_HANDOVER" else "İade / Harcama",
								received_by_name=o_recv,
								receipt_no=o_rec_no,
								notes=o_note,
								year_month=sel_ym,
							)
							ui.success(f"✅ {float(o_amt):,.2f} ₺ kasadan çıkış kaydedildi.")
							ui.rerun()
						except Exception as exc:
							ui.error(str(exc))

		if diners:
			with ui.expander("👨‍👩‍👧‍👦 Toplu / Kardeş Ödemesi Dağıt (Tek Getiren → Birden Fazla Öğrenci)"):
				bc1, bc2, bc3 = ui.columns(3)
				b_deliv = bc1.text_input("Ücreti Getiren Veli / İsim*", key="caf_bulk_deliv", placeholder="Örn: Veli Ahmet Kaya")
				b_meth = bc2.selectbox("Ödeme Yöntemi", ["Nakit", "IBAN / Havale", "Kart"], key="caf_bulk_meth")
				b_dt = bc3.date_input("Ödeme Tarihi", value=date.today(), key="caf_bulk_dt")
				rem_map = {int(b["diner_id"]): float(b["remaining_amount"]) for b in cash["balances"]}
				b_sel = ui.multiselect("Ödeme Dağıtılacak Kişiler*", list(diner_opts.keys()), key="caf_bulk_diners")
				if b_sel:
					df_alloc = pd.DataFrame([
						{
							"diner_id": diner_opts[lbl],
							"Kişi": lbl,
							"Dağıtılacak Tutar (₺)": rem_map.get(diner_opts[lbl], 0.0),
							"Not": "",
						}
						for lbl in b_sel
					])
					ed_alloc = ui.data_editor(
						df_alloc,
						hide_index=True,
						use_container_width=True,
						key="caf_bulk_alloc_ed",
						column_config={"diner_id": None, "Kişi": ui.column_config.TextColumn("Kişi", disabled=True)},
					)
					if ui.button("✅ Toplu / Kardeş Tahsilatlarını Kaydet", type="primary", key="caf_btn_bulk_pay"):
						alloc_rows = [
							{"diner_id": int(r["diner_id"]), "amount": float(r["Dağıtılacak Tutar (₺)"] or 0), "notes": str(r.get("Not") or "")}
							for _, r in ed_alloc.iterrows()
						]
						ids_new = record_bulk_sibling_payments(
							connect_db_fn,
							tx_date=b_dt.isoformat(),
							delivered_by_name=b_deliv,
							allocations=alloc_rows,
							payment_method=b_meth,
							year_month=sel_ym,
						)
						ui.success(f"✅ {len(ids_new)} kişi için toplu tahsilat kaydedildi.")
						ui.rerun()

		if is_on and diners:
			with ui.expander("🤖 Gemini AI ile Toplu Ödeme Mesajı / Listesi Oku (Önizlemeli)"):
				pay_msg_txt = ui.text_area("Öğretmen / Veli Ödeme Mesajı veya Listesi", height=80, placeholder="Örn: Ali Yılmaz'ın annesi Fatma 1500 TL nakit getirdi, Ece Demir IBAN 1500 TL...", key="caf_ai_pay_txt")
				if ui.button("✨ Ödeme Listesini Ayrıştır", key="caf_btn_ai_parse_pay"):
					strict_pdf = str(app_cfg.get("gemini_strict_kvkk_pdf", "0")).strip() == "1"
					with ui.spinner("Gemini AI ödeme satırlarını çözümlüyor..."):
						ok_p, res_p, used_p = gai.parse_cafeteria_menu_or_payments_text_with_gemini(
							api_key=api_key,
							model=model,
							task_mode="payments",
							workdays_list=[],
							raw_text=pay_msg_txt,
							strict_kvkk_pdf=strict_pdf,
						)
					if ok_p and res_p.get("payments"):
						ui.session_state["caf_ai_pay_preview"] = res_p["payments"]
						ui.success(f"✅ {len(res_p['payments'])} ödeme satırı bulundu ({used_p}).")
					else:
						ui.error(f"Ödeme metni okunamadı: {(res_p or {}).get('error', 'Satır yok')}")

				pay_prev = ui.session_state.get("caf_ai_pay_preview")
				if isinstance(pay_prev, list) and pay_prev:
					ui.dataframe(pd.DataFrame(pay_prev), hide_index=True, use_container_width=True)
					if ui.button("✅ Onayla ve Ayrıştırılan Ödemeleri Kasaya İşle", type="primary", key="caf_btn_commit_ai_pay"):
						by_nm = {_tr_fold(d["full_name"]): int(d["id"]) for d in diners}
						ins_c = 0
						for it in pay_prev:
							p_nm = str(it.get("for_person_name") or "").strip()
							did_m = by_nm.get(_tr_fold(p_nm))
							amt_m = _safe_float(it.get("amount"), 0.0)
							if amt_m > 0:
								record_cash_transaction(
									connect_db_fn,
									tx_date=str(it.get("tx_date") or date.today().isoformat())[:10],
									tx_type="IN_PAYMENT",
									amount=amt_m,
									diner_id=did_m,
									for_person_name=p_nm,
									delivered_by_name=str(it.get("delivered_by_name") or ""),
									payment_method=str(it.get("payment_method") or "Nakit"),
									notes=str(it.get("notes") or "AI Liste Aktarımı"),
									year_month=sel_ym,
								)
								ins_c += 1
						ui.session_state.pop("caf_ai_pay_preview", None)
						ui.success(f"✅ {ins_c} tahsilat kasaya işlendi.")
						ui.rerun()

		# Kasa Hareketleri Filtreleme, Makbuz & Teslim Tutanağı
		txs = cash["month_transactions"]
		ui.markdown(f"##### 📜 {_fmt_ym(sel_ym)} Kasa Hareketleri ({len(txs)})")
		if txs:
			kf1, kf2 = ui.columns([1.2, 2.0])
			tx_kind_flt = kf1.selectbox("İşlem Türü Filtresi", ["Tümü", "🟢 Yalnızca Tahsilatlar", "🔵 Okula Teslimler", "🟠 İade/Harcamalar"], key="caf_tx_kind_flt")
			tx_q_flt = _tr_fold(kf2.text_input("Getiren / Öğrenci / Makbuz Ara", key="caf_tx_q_flt"))
			txs_flt = []
			for t in txs:
				if "Tahsilat" in tx_kind_flt and t["tx_type"] != "IN_PAYMENT":
					continue
				if "Teslim" in tx_kind_flt and t["tx_type"] != "OUT_SCHOOL_HANDOVER":
					continue
				if "İade" in tx_kind_flt and t["tx_type"] != "OUT_REFUND_OR_PETTY":
					continue
				if tx_q_flt and not (
					tx_q_flt in _tr_fold(t.get("for_person_name"))
					or tx_q_flt in _tr_fold(t.get("delivered_by_name"))
					or tx_q_flt in _tr_fold(t.get("received_by_name"))
					or tx_q_flt in _tr_fold(t.get("receipt_no"))
					or tx_q_flt in _tr_fold(t.get("notes"))
				):
					continue
				txs_flt.append(t)

			kind_tr = {
				"IN_PAYMENT": "🟢 Tahsilat (+Kasaya)",
				"OUT_SCHOOL_HANDOVER": "🔵 Okula Teslim (-Kasadan)",
				"OUT_REFUND_OR_PETTY": "🟠 İade/Harcama (-Kasadan)",
			}
			df_tx = pd.DataFrame([
				{
					"ID": int(t["id"]),
					"Tarih": t["tx_date"],
					"İşlem": kind_tr.get(t["tx_type"], t["tx_type"]),
					"Kimin Adına": mask_fn(t.get("for_person_name") or "—", "name"),
					"Ücreti Getiren / Teslim Alan": mask_fn(t.get("delivered_by_name") or t.get("received_by_name") or "—", "name"),
					"Tutar (₺)": float(t.get("amount") or 0.0),
					"Yöntem": t.get("payment_method") or "Nakit",
					"Makbuz/Not": f"{t.get('receipt_no') or ''} {t.get('notes') or ''}".strip() or "—",
				}
				for t in txs_flt
			]) if txs_flt else pd.DataFrame([{"Bilgi": "Filtreye uyan kasa hareketi bulunamadı."}])
			ui.dataframe(df_tx, hide_index=True, use_container_width=True)

			tx_dl1, tx_dl2, tx_dl3 = ui.columns(3)
			with tx_dl1:
				ui.download_button(
					"📝 Okula Kasa Teslim Tutanağı (.docx)",
					data=lambda: build_cash_handover_docx(connect_db_fn, year_month=sel_ym, school_meta=app_cfg),
					file_name=f"Yemekhane_Kasa_Teslim_Tutanagi_{sel_ym}.docx",
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					type="primary",
					use_container_width=True,
					key="caf_dl_cash_docx",
				)
			with tx_dl2:
				in_txs = [t for t in txs if t["tx_type"] == "IN_PAYMENT"]
				if in_txs:
					rcpt_map = {f"#{int(t['id'])} — {mask_fn(t.get('for_person_name') or '', 'name')} · {float(t['amount']):,.0f} ₺": int(t["id"]) for t in in_txs}
					sel_rcpt = ui.selectbox("Tahsilat Makbuzu Seç", list(rcpt_map.keys()), key="caf_rcpt_sel")
					sel_rcpt_id = rcpt_map[sel_rcpt]
					ui.download_button(
						"🧾 Bireysel Tahsilat Makbuzu İndir (.docx)",
						data=lambda tid=sel_rcpt_id: build_payment_receipt_docx(connect_db_fn, tx_id=tid, school_meta=app_cfg),
						file_name=f"Yemekhane_Makbuz_{sel_rcpt_id}.docx",
						mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
						use_container_width=True,
						key="caf_dl_rcpt_docx",
					)
			with tx_dl3:
				with ui.expander("🗑️ Yanlış Girilen Kasa Hareketini Sil"):
					tx_opts = {f"#{int(t['id'])} — {t['tx_date']} · {float(t['amount']):,.2f} ₺ ({t.get('for_person_name') or ''})": int(t["id"]) for t in txs}
					s_tx = ui.selectbox("Silinecek Kasa Kaydı", list(tx_opts.keys()), key="caf_del_tx_sel")
					if ui.button("Kasa Kaydını Sil", key="caf_del_tx_btn"):
						delete_cash_transaction(connect_db_fn, tx_opts[s_tx])
						ui.success("Kasa kaydı silindi.")
						ui.rerun()
		else:
			ui.caption("Bu dönem için henüz tahsilat veya teslim kaydı yok.")

	with t_budget:
		pnl = get_budget_and_pnl_summary(connect_db_fn, sel_ym)
		ui.markdown(f"#### 📊 {_fmt_ym(sel_ym)} ve Genel Yemekhane Gelir – Gider Bütçe Dengesi")
		b1, b2, b3, b4 = ui.columns(4)
		b1.metric("Tahsil Edilen Gelir", f"{pnl['month_collected']:,.2f} ₺", delta=f"Tahsilat Oranı: %{pnl['collection_rate_pct']}")
		b2.metric("Mutfak / Ürün Gideri", f"{pnl['month_expenses']:,.2f} ₺", delta=f"Vadeli: {pnl['month_pending_expenses']:,.2f} ₺" if pnl["month_pending_expenses"] else "Tamamı Ödendi")
		b3.metric("Aylık Net Fark (Gelir - Gider)", f"{pnl['month_net_cash_vs_exp']:,.2f} ₺")
		b4.metric("Genel Tüm Zamanlar Net Bütçe", f"{pnl['all_time_net_vs_exp']:,.2f} ₺")

		ui.bar_chart({
			"Aylık Beklenen Aidat": pnl["month_expected"],
			"Aylık Tahsilat": pnl["month_collected"],
			"Aylık Fatura Gideri": pnl["month_expenses"],
		})

		if is_on:
			with ui.expander("🤖 Gemini AI Mali Analiz, Enflasyon Uyarısı & KVKK-Uyumlu Veli Hatırlatma Şablonu"):
				if ui.button("✨ Aylık Bütçe & Fiyat Artışlarını AI ile Analiz Et", type="primary", key="caf_btn_ai_fin"):
					unpaid_s = [
						{"class_or_dept": b.get("class_or_dept"), "role": b["role_type"], "remaining_tl": b["remaining_amount"], "status": b["status_label"]}
						for b in cash["balances"]
						if float(b.get("remaining_amount") or 0) > 0
					]
					inf_s = get_ingredient_price_inflation(connect_db_fn)
					with ui.spinner("Gemini AI gelir-gider dengesi ve fiyat trendlerini yorumluyor..."):
						ok_f, res_f, used_f = gai.analyze_cafeteria_finances_and_inflation_with_gemini(
							api_key=api_key,
							model=model,
							year_month=sel_ym,
							budget_summary=pnl,
							price_changes=inf_s,
							unpaid_summary=unpaid_s,
						)
					if ok_f:
						ui.session_state["caf_ai_fin_res"] = res_f
						ui.success(f"✅ Mali analiz tamamlandı ({used_f}).")
					else:
						ui.error(f"AI analiz hatası: {(res_f or {}).get('error', 'Yanıt alınamadı')}")

				f_res = ui.session_state.get("caf_ai_fin_res")
				if isinstance(f_res, dict) and f_res.get("executive_summary"):
					ui.info(str(f_res["executive_summary"]))
					for al_f in (f_res.get("inflation_alerts") or []):
						ui.warning(f"📈 {al_f}")
					for cs_f in (f_res.get("cost_saving_actions") or []):
						ui.caption(f"💡 {cs_f}")
					if f_res.get("parent_reminder_template"):
						ui.text_area("📩 Örnek Veli Aidat Hatırlatma Mesajı (Kopyalanabilir)", value=str(f_res["parent_reminder_template"]), height=100, key="caf_ai_rem_tmpl")


def _render_tab_settings_and_exports(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	save_cfg_fn = ctx_api.get("save_app_settings")
	sched = parse_service_schedule_from_cfg(app_cfg)
	ui.markdown("### ⚙️ Yemekhane Ayarları, Haftalık Çalışma / Öğün Saatleri, İmza & Toplu Excel Çıktısı")

	with ui.container(border=True):
		ui.markdown("##### 🗓️ Haftalık Yemek Günü, Günlük Öğün Sayısı, Yemek Saatleri, Ramazan Modu & İmza Ayarları")
		try:
			import school_calendar as scal
			all_prov_names = sorted(scal.TR_PROVINCE_COORDS.keys())
		except Exception:
			all_prov_names = ["Ankara", "İstanbul", "İzmir"]
		auto_prov_lbl = f"Kurum İli (Otomatik - {resolve_province_name(app_cfg.get('default_province') or 'Ankara')})"
		prov_choices = [auto_prov_lbl] + all_prov_names
		cur_ram_prov = str(app_cfg.get("cafeteria_ramazan_province") or "").strip()
		cur_prov_val = cur_ram_prov if cur_ram_prov in all_prov_names else auto_prov_lbl

		with ui.form("caf_form_settings"):
			wd_name_to_idx = {TR_WEEKDAYS_SHORT[i]: i for i in range(7)}
			default_wd_names = [TR_WEEKDAYS_SHORT[i] for i in sorted(sched["weekdays_set"])]

			hs1, hs2 = ui.columns([1.0, 2.2])
			days_pw_in = hs1.number_input("Haftada Kaç Gün Yemek Çıkıyor?", min_value=1, max_value=7, step=1, value=sched["days_per_week"])
			sel_wds_in = hs2.multiselect("Yemek Çıkan Günler", list(wd_name_to_idx.keys()), default=default_wd_names)

			hs3, hs4 = ui.columns([1.0, 2.2])
			meals_pd_in = hs3.number_input("Günde Kaç Öğün Yemek Çıkıyor?", min_value=1, max_value=5, step=1, value=sched["meals_per_day"])
			sel_slots_in = hs4.multiselect("Aktif Yemek Öğünleri (Ramazan Modu için İftar / Sahur da seçilebilir)", ALL_MEAL_SLOTS, default=sched["active_meal_slots"])

			tm1, tm2, tm3 = ui.columns(3)
			t_bkf_in = tm1.text_input("Sabah Kahvaltısı Saati", value=sched["meal_times"]["Sabah Kahvaltısı"], placeholder="08:00 - 08:45")
			t_lun_in = tm2.text_input("Öğle Yemeği Saati", value=sched["meal_times"]["Öğle Yemeği"], placeholder="12:00 - 13:00")
			t_snk_in = tm3.text_input("İkindi Ara Öğünü Saati", value=sched["meal_times"]["İkindi Ara Öğünü"], placeholder="15:00 - 15:30")

			rm1, rm2 = ui.columns([1.5, 1.5], vertical_alignment="center")
			ram_on_in = rm1.checkbox("🌙 Ramazan Modu Aktif (Menüye İl Bazlı İftar ve Sahur Saatlerini İşle)", value=sched["ramazan_mode"])
			ram_prov_in = rm2.selectbox("İftar & Sahur İçin İl Seçimi", prov_choices, index=prov_choices.index(cur_prov_val) if cur_prov_val in prov_choices else 0)

			if sched["ramazan_mode"]:
				today_rt = get_ramazan_times_for_dates(connect_db_fn, sched["ramazan_province"], [date.today().isoformat()]).get(date.today().isoformat(), {})
				if today_rt:
					ui.info(f"🌙 **{sched['ramazan_province']} Bugünkü Vakitler ({date.today().isoformat()}):** Sahur (İmsak): **{today_rt.get('sahur')}** · İftar (Akşam): **{today_rt.get('iftar')}** _({today_rt.get('source')})_")

			ui.markdown("---")
			(s1_def_r, s1_def_n), (s2_def_r, s2_def_n), (s3_def_r, s3_def_n) = resolve_menu_signer_slots(app_cfg)
			tone_choices = list(MENU_COLOR_TONE_PRESETS.keys())
			cur_tone = str(app_cfg.get("cafeteria_menu_color_tone") or DEFAULT_MENU_COLOR_TONE).strip()
			if cur_tone not in MENU_COLOR_TONE_PRESETS:
				cur_tone = DEFAULT_MENU_COLOR_TONE

			cfg_fp = json.dumps([s1_def_r, s1_def_n, s2_def_r, s2_def_n, s3_def_r, s3_def_n, cur_tone], ensure_ascii=False)
			if ui.session_state.get("caf_cfg_form_fp") != cfg_fp:
				ui.session_state["caf_cfg_form_fp"] = cfg_fp
				ui.session_state["caf_cfg_s1_r"] = s1_def_r
				ui.session_state["caf_cfg_s1_n"] = s1_def_n
				ui.session_state["caf_cfg_s2_r"] = s2_def_r
				ui.session_state["caf_cfg_s2_n"] = s2_def_n
				ui.session_state["caf_cfg_s3_r"] = s3_def_r
				ui.session_state["caf_cfg_s3_n"] = s3_def_n
				ui.session_state["caf_cfg_menu_tone"] = cur_tone

			ct1, ct2, ct3 = ui.columns([1.2, 0.9, 0.9], vertical_alignment="center")
			tone_in = ct1.selectbox("🎨 Menü Çıktısı Renk Tonu", tone_choices, index=tone_choices.index(cur_tone), key="caf_cfg_menu_tone")
			def_stu_p = ct2.number_input("Varsayılan Öğrenci Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_meal_price"), 75.0))
			def_stf_p = ct3.number_input("Varsayılan Öğretmen/Personel Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_staff_price"), 90.0))

			# Seçili renk tonunun canlı mini önizlemesi
			sw_items = []
			for wk_i in range(4):
				pal_w = get_week_palette_pair(wk_i, color_tone=tone_in)
				sw_items.append(
					f"<span style='display:inline-flex;align-items:center;gap:3px;font-size:11px;padding:2px 6px;border:1px solid #cbd5e1;border-radius:5px;'>"
					f"<span style='width:12px;height:12px;border-radius:2px;background:{pal_w['light']};border:1px solid #94a3b8;'></span>"
					f"<span style='width:12px;height:12px;border-radius:2px;background:{pal_w['dark']};border:1px solid #94a3b8;'></span>"
					f"{wk_i + 1}. Hft</span>"
				)
			ui.markdown(f"<div style='display:flex;flex-wrap:wrap;gap:6px;margin:-4px 0 6px 0;'>{''.join(sw_items)}</div>", unsafe_allow_html=True)

			sg1, sg2, sg3 = ui.columns(3)
			s1_r = sg1.text_input("1. İmzacı Ünvan (Sol)", value=s1_def_r, key="caf_cfg_s1_r")
			s1_n = sg1.text_input("1. İmzacı Ad Soyad (Sol)", value=s1_def_n, key="caf_cfg_s1_n")
			s2_r = sg2.text_input("2. İmzacı Ünvan (Orta — boş bırakılırsa 2 imza)", value=s2_def_r, key="caf_cfg_s2_r")
			s2_n = sg2.text_input("2. İmzacı Ad Soyad (Orta)", value=s2_def_n, key="caf_cfg_s2_n")
			s3_r = sg3.text_input("3. İmzacı Ünvan (Sağ)", value=s3_def_r, key="caf_cfg_s3_r")
			s3_n = sg3.text_input("3. İmzacı Ad Soyad (Sağ)", value=s3_def_n, key="caf_cfg_s3_n")
			b_note_in = ui.text_input(
				"A4 Menü Alt Bilgi Notu (Opsiyonel — boş bırakılırsa standart alerji/revir uyarısı yazılır)",
				value=app_cfg.get("cafeteria_board_note", ""),
				placeholder="* Gıda alerjisi / özel diyeti olan öğrencilerimiz için revirle koordineli sunum yapılır.",
			)

			if ui.form_submit_button("💾 Yemekhane Ayarlarını Kaydet", type="primary", use_container_width=True) and callable(save_cfg_fn):
				picked_wd_idxs = sorted(wd_name_to_idx[n] for n in sel_wds_in if n in wd_name_to_idx)
				if set(picked_wd_idxs) != sched["weekdays_set"] and picked_wd_idxs:
					chosen_wd_idxs = picked_wd_idxs
					final_days_pw = len(chosen_wd_idxs)
				elif int(days_pw_in) != sched["days_per_week"] or not picked_wd_idxs:
					final_days_pw = max(1, min(7, int(days_pw_in)))
					chosen_wd_idxs = list(range(final_days_pw))
				else:
					chosen_wd_idxs = picked_wd_idxs
					final_days_pw = len(chosen_wd_idxs)

				picked_slots = [s for s in sel_slots_in if s in ALL_MEAL_SLOTS]
				if picked_slots != sched["active_meal_slots"] and picked_slots:
					chosen_slots = picked_slots
					final_meals_pd = len(chosen_slots)
				elif int(meals_pd_in) != sched["meals_per_day"] or not picked_slots:
					final_meals_pd = max(1, min(5, int(meals_pd_in)))
					chosen_slots = ["Öğle Yemeği"] if final_meals_pd == 1 else (["Sabah Kahvaltısı", "Öğle Yemeği"] if final_meals_pd == 2 else ALL_MEAL_SLOTS[:final_meals_pd])
				else:
					chosen_slots = picked_slots
					final_meals_pd = len(chosen_slots)

				prov_to_save = "" if ram_prov_in == auto_prov_lbl else ram_prov_in
				save_cfg_fn({
					"cafeteria_service_days_per_week": str(final_days_pw),
					"cafeteria_service_weekdays": ",".join(str(i) for i in chosen_wd_idxs),
					"cafeteria_meals_per_day": str(final_meals_pd),
					"cafeteria_active_meal_slots": ",".join(chosen_slots),
					"cafeteria_time_breakfast": t_bkf_in.strip() or "08:00 - 08:45",
					"cafeteria_time_lunch": t_lun_in.strip() or "12:00 - 13:00",
					"cafeteria_time_snack": t_snk_in.strip() or "15:00 - 15:30",
					"cafeteria_ramazan_mode": "1" if ram_on_in else "0",
					"cafeteria_ramazan_province": prov_to_save,
					"cafeteria_menu_color_tone": tone_in,
					"cafeteria_default_meal_price": str(def_stu_p),
					"cafeteria_default_staff_price": str(def_stf_p),
					"cafeteria_sig1_role": s1_r.strip(),
					"cafeteria_sig1_name": s1_n.strip(),
					"cafeteria_sig2_role": s2_r.strip(),
					"cafeteria_sig2_name": s2_n.strip(),
					"cafeteria_sig3_role": s3_r.strip(),
					"cafeteria_sig3_name": s3_n.strip(),
					"cafeteria_sigs_customized": "1",
					"cafeteria_board_note": b_note_in.strip(),
				})
				ui.session_state.pop("caf_menu_active_slots", None)
				ui.session_state.pop("caf_cfg_form_fp", None)
				if callable(ctx_api.get("clear_all_data_caches")):
					ctx_api["clear_all_data_caches"]()
				ui.success("✅ Yemekhane çalışma düzeni, Ramazan modu, renk tonu, yemek saatleri ve imza ayarları kaydedildi.")
				ui.rerun()

	with ui.expander("🏫 A4 Menü İçin Resmi Okul Logosu / Mühür Yükle (PNG/JPG)", expanded=False):
		has_logo = bool(str(app_cfg.get("cafeteria_school_logo_b64") or "").strip())
		if has_logo:
			ui.info("✅ Kayıtlı okul logosu mevcut (A4 Word ve HTML çıktılarında otomatik gösterilir).")
		up_logo = ui.file_uploader("Logo Dosyası Seç (.png, .jpg)", type=["png", "jpg", "jpeg"], key="caf_logo_up")
		lg_c1, lg_c2 = ui.columns(2)
		with lg_c1:
			if ui.button("🖼️ Logoyu Kaydet", disabled=up_logo is None, use_container_width=True, key="caf_btn_save_logo") and callable(save_cfg_fn) and up_logo is not None:
				b64_logo = base64.b64encode(up_logo.getvalue()).decode("ascii")
				save_cfg_fn({"cafeteria_school_logo_b64": b64_logo})
				if callable(ctx_api.get("clear_all_data_caches")):
					ctx_api["clear_all_data_caches"]()
				ui.success("✅ Okul logosu kaydedildi.")
				ui.rerun()
		with lg_c2:
			if has_logo and ui.button("🗑️ Logoyu Kaldır", type="secondary", use_container_width=True, key="caf_btn_del_logo") and callable(save_cfg_fn):
				save_cfg_fn({"cafeteria_school_logo_b64": ""})
				if callable(ctx_api.get("clear_all_data_caches")):
					ctx_api["clear_all_data_caches"]()
				ui.success("Logo kaldırıldı.")
				ui.rerun()

	ui.markdown("---")
	ui.download_button(
		f"📊 {_fmt_ym(sel_ym)} Tüm Yemekhane Çalışma Kitabını İndir (.xlsx)",
		data=lambda: build_cafeteria_full_workbook_excel(connect_db_fn, sel_ym),
		file_name=f"Yemekhane_Tam_Rapor_{sel_ym}.xlsx",
		mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		type="primary",
		use_container_width=True,
		key="caf_dl_full_xlsx",
	)


def clean_name_valid(val: Any) -> bool:
	return bool(str(val or "").strip())


def render_cafeteria_management_page(ctx_api: dict[str, Any]) -> None:
	"""🍽️ Yemekhane Yönetimi ana sayfasını 5 sekme ile oluşturur."""
	connect_db_fn = ctx_api["connect_db"]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)

	app_cfg = ctx_api["get_app_settings"]()
	ym_opts, cur_ym = _month_options(connect_db_fn)

	hdr_l, hdr_r = ui.columns([2.6, 1.2], vertical_alignment="center")
	with hdr_l:
		ui.markdown("## 🍽️ Yemekhane, Menü Planlama, Gider & Kasa Yönetimi")
		ui.caption("Yemekler bankası (AI malzeme/kalori/alerjen), iş günü bazlı öğün planı, imzalı A4 çıktı, gelen ürün/fatura giderleri ve öğrenci/kasa muhasebesi.")
	with hdr_r:
		sel_ym = ui.selectbox(
			"Aktif Ay",
			ym_opts,
			index=ym_opts.index(cur_ym) if cur_ym in ym_opts else 0,
			format_func=_fmt_ym,
			key="caf_global_sel_ym",
		)

	tab_menu, tab_dishes, tab_purch, tab_cash, tab_cfg = ui.tabs([
		"🗓️ Menü Planlayıcı & A4 Çıktı",
		"🍲 Yemekler Bankası & AI Malzeme",
		"🚚 Gelen Ürün & Giderler",
		"💵 Yemek Yiyenler & Kasa",
		"⚙️ Yemekhane Ayarları & Excel",
	])
	with tab_menu:
		_render_tab_menu_planner(ctx_api, sel_ym, app_cfg)
	with tab_dishes:
		_render_tab_dishes(ctx_api)
	with tab_purch:
		_render_tab_purchases_and_expenses(ctx_api, sel_ym, app_cfg)
	with tab_cash:
		_render_tab_diners_and_cash(ctx_api, sel_ym, app_cfg)
	with tab_cfg:
		_render_tab_settings_and_exports(ctx_api, sel_ym, app_cfg)
