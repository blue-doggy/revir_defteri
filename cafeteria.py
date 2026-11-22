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

import calendar
import html
import io
import json
import re
import sqlite3
from datetime import date, datetime
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
	return target_id


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
	"""Verilen alerjen listesini revir sağlık profilleri ve yemekhane aboneleriyle çapraz eşler."""
	norm_als = [_normalize_allergen_label(a) for a in (allergens or []) if _normalize_allergen_label(a)]
	if not norm_als:
		return []

	hits: list[dict[str, str]] = []
	seen_keys: set[tuple[str, str]] = set()
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		try:
			stu_rows = conn.execute(
				"""SELECT s.id, s.full_name, COALESCE(c.name, '—') AS class_name, hp.allergies
				   FROM students s
				   JOIN health_profiles hp ON hp.student_id = s.id
				   LEFT JOIN classes c ON c.id = s.class_id
				   WHERE TRIM(COALESCE(hp.allergies, '')) != ''"""
			).fetchall()
		except Exception:
			stu_rows = []
		try:
			diner_rows = conn.execute(
				"""SELECT id, full_name, COALESCE(class_or_dept, role_type) AS class_name, allergies_override
				   FROM cafeteria_diners
				   WHERE is_active=1 AND TRIM(COALESCE(allergies_override, '')) != ''"""
			).fetchall()
		except Exception:
			diner_rows = []

	def _scan(rows_in: list[Any], col_name: str) -> None:
		for r in rows_in:
			al_txt = _tr_fold(r[col_name])
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
					hits.append({
						"full_name": str(r["full_name"]),
						"class_name": str(r["class_name"] or "—"),
						"allergies": str(r[col_name]),
						"matched": ", ".join(matched_als),
					})

	_scan(stu_rows, "allergies")
	_scan(diner_rows, "allergies_override")
	return hits


# ==============================================================================
# 4. BÖLÜM 2: İŞ GÜNLERİ, ÖĞÜN BAZLI MENÜ PLANLAYICI & A4 ÇIKTI MOTORU
# ==============================================================================

def get_month_workdays_info(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
	*,
	include_weekends: bool = False,
	treat_half_day_as_workday: bool = True,
) -> list[dict[str, Any]]:
	"""Seçili ayın günlerini okul takvimindeki tatillere göre iş günü / tatil olarak sınıflandırır."""
	ym_clean = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy, mm = [int(x) for x in ym_clean.split("-")]
	try:
		import school_calendar as scal
		scal.ensure_year_builtin_holidays_seeded(yy, connect_db_fn)
	except Exception:
		pass

	events_by_date: dict[str, list[dict[str, Any]]] = {}
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
		is_work = (include_weekends or not is_wknd) and (not is_holiday) and (treat_half_day_as_workday or not is_half)
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
	"""Belirli tarih ve öğün slotuna yemek listesini kaydeder; kalori, karbonhidrat ve alerjenleri toplar."""
	p_date = str(plan_date).strip()[:10]
	ym = p_date[:7]
	slot = str(meal_slot or "Öğle Yemeği").strip() or "Öğle Yemeği"
	d_ids = [int(x) for x in (dish_ids or []) if str(x).isdigit()]
	cust = [c for c in (custom_items or []) if isinstance(c, dict) and str(c.get("name") or "").strip()]

	tot_kcal = 0.0
	tot_carb = 0.0
	allergens: list[str] = []
	seen_al: set[str] = set()
	diabetes_items: list[dict[str, Any]] = []

	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
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

		tot_kcal = round(tot_kcal, 1)
		tot_carb = round(tot_carb, 1)
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
		slot_map = {
			"Sabah Kahvaltısı": "Kahvaltı",
			"Öğle Yemeği": "Öğle Yemeği",
			"İkindi Ara Öğünü": "İkindi / Ara Öğün",
		}
		idm.upsert_school_meal_menu(
			connect_db_fn,
			p_date,
			slot_map.get(slot, slot),
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
		return int(cur.rowcount or 0)


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


def build_menu_a4_html(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	scope: str = "monthly",
	week_index: int = 0,
	active_slots: list[str] | None = None,
	show_kcal: bool = True,
	show_allergens: bool = True,
	include_weekends: bool = False,
	signers: list[tuple[str, str]] | None = None,
	school_meta: dict[str, str] | None = None,
) -> str:
	"""Aylık veya haftalık menüyü doğrudan A4 kağıdına basılabilir antetli & imzalı HTML formatında oluşturur."""
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy, mm = [int(x) for x in ym.split("-")]
	month_title = f"{TR_MONTH_NAMES.get(mm, mm)} {yy}"
	slots = [s for s in (active_slots or ["Öğle Yemeği"]) if s]
	days_info = get_month_workdays_info(connect_db_fn, ym, include_weekends=include_weekends)
	plans = get_menu_plans_for_month(connect_db_fn, ym)

	weeks = _group_month_days_by_week(days_info)
	if scope == "weekly" and weeks:
		safe_idx = max(0, min(week_index, len(weeks) - 1))
		target_days = weeks[safe_idx]
		scope_label = f"{safe_idx + 1}. Hafta Menüsü ({target_days[0]['date_iso']} – {target_days[-1]['date_iso']})"
	else:
		target_days = [d for d in days_info if d.get("is_workday")]
		scope_label = f"{month_title} Aylık Yemek Menüsü ({len(target_days)} İş Günü)"

	meta = school_meta or {}
	inst_top = html.escape(meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	school_nm = html.escape(meta.get("school_name") or "Okul / Kurum Yemekhanesi")
	sig_list = [s for s in (signers or []) if (s[0] or s[1])]
	if not sig_list:
		sig_list = [
			("Hazırlayan", meta.get("nurse_name") or "........................"),
			("Onay / Okul Müdürü", meta.get("principal_name") or "........................"),
		]

	rows_html = []
	for d in target_days:
		d_iso = d["date_iso"]
		dt_tr = f"{d['day_num']:02d}.{mm:02d}.{yy}<br><small>{html.escape(d['weekday_name'])}</small>"
		day_kcal = 0.0
		cells = [f"<td class='col-date'>{dt_tr}</td>"]
		for sl in slots:
			rec = plans.get((d_iso, sl))
			if rec and rec.get("dish_names"):
				dishes_str = "<br>".join(f"• {html.escape(str(nm))}" for nm in rec["dish_names"])
				meta_bits = []
				k_val = _safe_float(rec.get("total_kcal"), 0.0)
				day_kcal += k_val
				if show_kcal and k_val > 0:
					meta_bits.append(f"<span class='kcal-pill'>🔥 {k_val:g} kcal</span>")
				if show_allergens and rec.get("allergens"):
					al_str = " · ".join(format_allergen_badge(a) for a in rec["allergens"])
					meta_bits.append(f"<span class='al-pill'>{html.escape(al_str)}</span>")
				meta_div = f"<div class='cell-meta'>{' '.join(meta_bits)}</div>" if meta_bits else ""
				cells.append(f"<td><div class='dishes'>{dishes_str}</div>{meta_div}</td>")
			else:
				cells.append("<td class='empty-cell'>—</td>")
		if show_kcal and len(slots) > 1:
			cells.append(f"<td class='col-tot'><b>{day_kcal:g} kcal</b></td>" if day_kcal > 0 else "<td class='col-tot'>—</td>")
		rows_html.append(f"<tr>{''.join(cells)}</tr>")

	hdr_cols = ["<th>Tarih / Gün</th>"] + [f"<th>{html.escape(sl)}</th>" for sl in slots]
	if show_kcal and len(slots) > 1:
		hdr_cols.append("<th>Günlük Toplam</th>")

	sig_html = "".join(
		f"<div class='sig-box'><div class='sig-role'>{html.escape(str(role))}</div>"
		f"<div class='sig-name'>{html.escape(str(person or '................................'))}</div>"
		f"<div class='sig-line'>İmza / Tarih</div></div>"
		for role, person in sig_list
	)

	page_orient = "landscape" if (len(slots) >= 2 and scope != "weekly") else "portrait"
	page_w = "297mm" if page_orient == "landscape" else "210mm"
	return f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{html.escape(scope_label)}</title>
<style>
@page {{ size: A4 {page_orient}; margin: 10mm 12mm; }}
* {{ box-sizing: border-box; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; color: #0f172a; background: #cbd5e1; margin: 0; padding: 14px; }}
.topbar {{ max-width: {page_w}; margin: 0 auto 10px auto; display: flex; justify-content: space-between; align-items: center; }}
.btn-print {{ padding: 8px 16px; background: #0f766e; color: #fff; border: none; border-radius: 7px; font-weight: 700; cursor: pointer; }}
.a4-sheet {{ width: min(100%, {page_w}); margin: 0 auto; padding: 11mm 13mm; background: #fff; border: 1px solid #94a3b8; box-shadow: 0 8px 24px rgba(15,23,42,.18); }}
.antet {{ text-align: center; border-bottom: 2px solid #0f766e; padding-bottom: 8px; margin-bottom: 10px; }}
.antet-top {{ font-size: 11px; color: #64748b; letter-spacing: .5px; }}
.antet-school {{ font-size: 16px; font-weight: 800; color: #0f766e; margin-top: 2px; }}
.antet-title {{ font-size: 13.5px; font-weight: 700; color: #1e293b; margin-top: 3px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th, td {{ border: 1px solid #cbd5e1; padding: 5px 7px; vertical-align: top; text-align: left; }}
th {{ background: #0f766e; color: #fff; font-weight: 700; text-align: center; }}
tr:nth-child(even) td {{ background: #f8fafc; }}
.col-date {{ width: 105px; font-weight: 700; color: #0f766e; white-space: nowrap; }}
.col-tot {{ width: 90px; text-align: center; vertical-align: middle; color: #0f766e; }}
.dishes {{ line-height: 1.38; }}
.cell-meta {{ margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px; }}
.kcal-pill {{ background: #fef3c7; color: #92400e; border-radius: 4px; padding: 1px 5px; font-size: 9.5px; font-weight: 700; }}
.al-pill {{ background: #fee2e2; color: #991b1b; border-radius: 4px; padding: 1px 5px; font-size: 9.5px; font-weight: 600; }}
.empty-cell {{ text-align: center; color: #94a3b8; }}
.note-line {{ margin-top: 8px; font-size: 10px; color: #475569; }}
.sig-row {{ margin-top: 22px; display: flex; justify-content: space-around; gap: 16px; page-break-inside: avoid; }}
.sig-box {{ flex: 1; text-align: center; border-top: 1px solid #94a3b8; padding-top: 6px; font-size: 11px; }}
.sig-role {{ font-weight: 700; color: #1e293b; }}
.sig-name {{ font-weight: 600; color: #0f766e; margin-top: 3px; }}
.sig-line {{ font-size: 10px; color: #64748b; margin-top: 18px; }}
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
    <div class='antet-top'>{inst_top}</div>
    <div class='antet-school'>{school_nm}</div>
    <div class='antet-title'>{html.escape(scope_label)}</div>
  </div>
  <table>
    <thead><tr>{''.join(hdr_cols)}</tr></thead>
    <tbody>{''.join(rows_html) if rows_html else '<tr><td colspan="4" class="empty-cell">Bu dönem için iş günü bulunamadı.</td></tr>'}</tbody>
  </table>
  <div class='note-line'>* Gıda alerjisi veya özel beslenme gereksinimi (Çölyak, Tip 1 Diyabet vb.) bulunan öğrencilerimiz için okul reviri ile koordineli alternatif sunum uygulanmaktadır.</div>
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


def _style_docx_cell(dx: dict[str, Any], cell: Any, bg_hex: str) -> None:
	shading = dx["parse_xml"](f'<w:shd {dx["nsdecls"]("w")} w:fill="{bg_hex}"/>')
	cell._tc.get_or_add_tcPr().append(shading)


def build_menu_a4_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	year_month: str,
	scope: str = "monthly",
	week_index: int = 0,
	active_slots: list[str] | None = None,
	show_kcal: bool = True,
	show_allergens: bool = True,
	include_weekends: bool = False,
	signers: list[tuple[str, str]] | None = None,
	school_meta: dict[str, str] | None = None,
) -> bytes:
	"""Aylık veya haftalık yemek menüsünü imzalı resmi A4 Word (.docx) belgesi olarak üretir."""
	dx = _ensure_docx_mod()
	docx = dx["docx"]
	Cm, Pt, RGBColor = dx["Cm"], dx["Pt"], dx["RGBColor"]

	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy, mm = [int(x) for x in ym.split("-")]
	month_title = f"{TR_MONTH_NAMES.get(mm, mm)} {yy}"
	slots = [s for s in (active_slots or ["Öğle Yemeği"]) if s]
	days_info = get_month_workdays_info(connect_db_fn, ym, include_weekends=include_weekends)
	plans = get_menu_plans_for_month(connect_db_fn, ym)

	weeks = _group_month_days_by_week(days_info)
	if scope == "weekly" and weeks:
		safe_idx = max(0, min(week_index, len(weeks) - 1))
		target_days = weeks[safe_idx]
		scope_label = f"{safe_idx + 1}. Hafta Yemek Menüsü ({target_days[0]['date_iso']} – {target_days[-1]['date_iso']})"
	else:
		target_days = [d for d in days_info if d.get("is_workday")]
		scope_label = f"{month_title} Aylık Yemek Menüsü ({len(target_days)} İş Günü)"

	meta = school_meta or {}
	inst_top = meta.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI"
	school_nm = meta.get("school_name") or "Okul / Kurum Yemekhanesi"

	doc = docx.Document()
	sec = doc.sections[0]
	is_landscape = (len(slots) >= 2 and scope != "weekly")
	sec.orientation = dx["WD_ORIENT"].LANDSCAPE if is_landscape else dx["WD_ORIENT"].PORTRAIT
	sec.page_width = Cm(29.7 if is_landscape else 21.0)
	sec.page_height = Cm(21.0 if is_landscape else 29.7)
	sec.top_margin = Cm(1.0)
	sec.bottom_margin = Cm(1.0)
	sec.left_margin = Cm(1.1)
	sec.right_margin = Cm(1.1)

	for txt_line, fsz, bold_f, rgb_c in [
		(inst_top, 9.5, False, RGBColor(100, 116, 139)),
		(school_nm.upper(), 12.5, True, RGBColor(15, 118, 110)),
		(scope_label, 11.0, True, RGBColor(30, 41, 59)),
	]:
		p = doc.add_paragraph()
		p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		p.paragraph_format.space_before = Pt(0)
		p.paragraph_format.space_after = Pt(2)
		r = p.add_run(txt_line)
		r.font.name = "Calibri"
		r.font.size = Pt(fsz)
		r.bold = bold_f
		r.font.color.rgb = rgb_c

	cols_hdr = ["Tarih / Gün"] + slots + (["Toplam Kalori"] if (show_kcal and len(slots) > 1) else [])
	tbl = doc.add_table(rows=max(1, len(target_days)) + 1, cols=len(cols_hdr))
	tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	tbl.autofit = True

	borders = dx["parse_xml"](
		f'<w:tblBorders {dx["nsdecls"]("w")}>'
		'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'<w:left w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
		'</w:tblBorders>'
	)
	tbl._tbl.tblPr.append(borders)

	for ci, h_txt in enumerate(cols_hdr):
		c = tbl.cell(0, ci)
		_style_docx_cell(dx, c, "0F766E")
		p = c.paragraphs[0]
		p.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		p.paragraph_format.space_after = Pt(0)
		rr = p.add_run(h_txt)
		rr.font.name = "Calibri"
		rr.font.size = Pt(9)
		rr.bold = True
		rr.font.color.rgb = RGBColor(255, 255, 255)

	for ri, d in enumerate(target_days, start=1):
		d_iso = d["date_iso"]
		bg = "FFFFFF" if ri % 2 == 1 else "F8FAFC"
		c0 = tbl.cell(ri, 0)
		_style_docx_cell(dx, c0, bg)
		p0 = c0.paragraphs[0]
		p0.paragraph_format.space_after = Pt(0)
		r0 = p0.add_run(f"{d['day_num']:02d}.{mm:02d}.{yy}\n{d['weekday_name']}")
		r0.font.name = "Calibri"
		r0.font.size = Pt(8.5)
		r0.bold = True
		r0.font.color.rgb = RGBColor(15, 118, 110)

		day_kcal = 0.0
		for si, sl in enumerate(slots, start=1):
			cell = tbl.cell(ri, si)
			_style_docx_cell(dx, cell, bg)
			cp = cell.paragraphs[0]
			cp.paragraph_format.space_after = Pt(0)
			rec = plans.get((d_iso, sl))
			if rec and rec.get("dish_names"):
				lines_txt = "\n".join(f"• {nm}" for nm in rec["dish_names"])
				r_d = cp.add_run(lines_txt)
				r_d.font.name = "Calibri"
				r_d.font.size = Pt(8.5)
				kv = _safe_float(rec.get("total_kcal"), 0.0)
				day_kcal += kv
				extra_bits = []
				if show_kcal and kv > 0:
					extra_bits.append(f"{kv:g} kcal")
				if show_allergens and rec.get("allergens"):
					extra_bits.append("Alerjen: " + ", ".join(rec["allergens"]))
				if extra_bits:
					r_m = cp.add_run("\n" + " | ".join(extra_bits))
					r_m.font.name = "Calibri"
					r_m.font.size = Pt(7.5)
					r_m.font.color.rgb = RGBColor(153, 27, 27) if rec.get("allergens") else RGBColor(146, 64, 14)
			else:
				r_e = cp.add_run("—")
				r_e.font.size = Pt(8.5)

		if show_kcal and len(slots) > 1:
			ct = tbl.cell(ri, len(slots) + 1)
			_style_docx_cell(dx, ct, bg)
			pt = ct.paragraphs[0]
			pt.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
			rt = pt.add_run(f"{day_kcal:g} kcal" if day_kcal > 0 else "—")
			rt.font.name = "Calibri"
			rt.font.size = Pt(8.5)
			rt.bold = True

	sig_list = [s for s in (signers or []) if (s[0] or s[1])]
	if not sig_list:
		sig_list = [
			("Hazırlayan", meta.get("nurse_name") or "........................"),
			("Okul Müdürü (Onay)", meta.get("principal_name") or "........................"),
		]
	p_sp = doc.add_paragraph()
	p_sp.paragraph_format.space_before = Pt(12)
	sig_tbl = doc.add_table(rows=1, cols=len(sig_list))
	sig_tbl.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
	for idx, (role_lbl, p_name) in enumerate(sig_list):
		sc = sig_tbl.cell(0, idx)
		sp = sc.paragraphs[0]
		sp.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
		sr1 = sp.add_run(f"{role_lbl}\n")
		sr1.font.name = "Calibri"
		sr1.font.size = Pt(9.5)
		sr1.bold = True
		sr2 = sp.add_run(f"{p_name or '........................'}\n\nİmza / Tarih")
		sr2.font.name = "Calibri"
		sr2.font.size = Pt(8.5)
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
	return purchases, items


def get_expense_summary_stats(connect_db_fn: Callable[[], sqlite3.Connection], year_month: str) -> dict[str, Any]:
	ym = str(year_month or date.today().strftime("%Y-%m"))[:7]
	yy = ym[:4]
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		m_row = conn.execute(
			"SELECT COALESCE(SUM(total_amount),0) AS tot, COUNT(*) AS cnt FROM cafeteria_purchases WHERE year_month=?",
			(ym,),
		).fetchone()
		y_row = conn.execute(
			"SELECT COALESCE(SUM(total_amount),0) AS tot FROM cafeteria_purchases WHERE substr(year_month,1,4)=?",
			(yy,),
		).fetchone()
		all_row = conn.execute("SELECT COALESCE(SUM(total_amount),0) AS tot FROM cafeteria_purchases").fetchone()
		cat_rows = conn.execute(
			"""SELECT pi.category AS cat, COALESCE(SUM(pi.line_total),0) AS tot
			   FROM cafeteria_purchase_items pi
			   JOIN cafeteria_purchases p ON p.id = pi.purchase_id
			   WHERE p.year_month=?
			   GROUP BY pi.category
			   ORDER BY tot DESC""",
			(ym,),
		).fetchall()
	return {
		"month_total": round(float(m_row["tot"] or 0.0), 2),
		"month_invoice_count": int(m_row["cnt"] or 0),
		"year_total": round(float(y_row["tot"] or 0.0), 2),
		"all_time_total": round(float(all_row["tot"] or 0.0), 2),
		"by_category": {str(r["cat"]): round(float(r["tot"] or 0.0), 2) for r in cat_rows},
	}


def delete_purchase_by_id(connect_db_fn: Callable[[], sqlite3.Connection], purchase_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute("DELETE FROM cafeteria_purchases WHERE id=?", (int(purchase_id),))


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


def delete_diner_by_id(connect_db_fn: Callable[[], sqlite3.Connection], diner_id: int) -> None:
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
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
	Seçili ayın iş günü sayısı × günlük öğün × öğün ücreti formülüyle
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
			exp_amt = 0.0 if is_exempt else max(0.0, round(p_days * meals_pd * unit_p - disc, 2))
			conn.execute(
				"""INSERT INTO cafeteria_monthly_fees(
					year_month, diner_id, workdays_in_month, planned_days, meals_per_day,
					unit_meal_price, discount_or_adjustment, expected_amount, note, updated_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				ON CONFLICT(year_month, diner_id) DO UPDATE SET
					workdays_in_month=excluded.workdays_in_month,
					planned_days=excluded.planned_days,
					meals_per_day=excluded.meals_per_day,
					unit_meal_price=excluded.unit_meal_price,
					discount_or_adjustment=excluded.discount_or_adjustment,
					expected_amount=excluded.expected_amount,
					updated_at=excluded.updated_at""",
				(ym, did, workdays, p_days, meals_pd, unit_p, disc, exp_amt, (cur_row["note"] if cur_row else ""), now_iso),
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
	note: str = "",
) -> None:
	p_days = max(0, int(planned_days))
	mpd = max(1, min(3, int(meals_per_day)))
	up = max(0.0, float(unit_meal_price))
	disc = float(discount_or_adjustment or 0.0)
	exp_amt = max(0.0, round(p_days * mpd * up - disc, 2))
	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db_fn() as conn:
		ensure_cafeteria_tables(conn)
		conn.execute(
			"""UPDATE cafeteria_monthly_fees
			   SET planned_days=?, meals_per_day=?, unit_meal_price=?, discount_or_adjustment=?, expected_amount=?, note=?, updated_at=?
			   WHERE id=?""",
			(p_days, mpd, up, disc, exp_amt, note.strip(), now_iso, int(fee_id)),
		)


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


def get_cash_summary_and_diner_balances(
	connect_db_fn: Callable[[], sqlite3.Connection],
	year_month: str,
) -> dict[str, Any]:
	"""Seçili ay tahakkuk/bakiye çizelgesi ile hem aylık hem tüm zamanlar kasa net bakiyesini döndürür."""
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
			"Yemekler": ", ".join(rec.get("dish_names") or []),
			"Kalori (kcal)": rec["total_kcal"],
			"Karbonhidrat (g)": rec["total_carbs_g"],
			"Alerjen Uyarısı": ", ".join(rec.get("allergens") or []),
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
			"Beklenen (₺)": b["expected_amount"],
			"Ödenen (₺)": b["paid_amount"],
			"Kalan (₺)": b["remaining_amount"],
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
	return buf.getvalue()


# ==============================================================================
# 7. UI SAYFA BİLEŞENLERİ (5 SEKME)
# ==============================================================================

def _month_options() -> tuple[list[str], str]:
	today = date.today()
	opts: list[str] = []
	for off in range(-4, 8):
		tot = today.year * 12 + (today.month - 1) + off
		yy, mm = divmod(tot, 12)
		opts.append(f"{yy:04d}-{mm + 1:02d}")
	cur = f"{today.year:04d}-{today.month:02d}"
	return opts, cur


def _fmt_ym(ym_s: str) -> str:
	try:
		yy, mm = ym_s.split("-")
		return f"{TR_MONTH_NAMES.get(int(mm), mm)} {yy}"
	except Exception:
		return ym_s


def _render_tab_dishes(ctx_api: dict[str, Any]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()
	mask_fn = ctx_api["mask_kvkk_value"]

	ui.markdown("### 🍲 Kayıtlı Yemekler Bankası, AI Malzeme/Kalori & Alerjen Çözümleyici")
	ui.caption("Daha önce pişirdiğiniz veya planladığınız yemekleri kaydedin; dilediğiniz yemeğin içindeki malzemeleri Gemini AI ile otomatik çıkarıp tablodan miktar ve malzemeleri özgürce düzenleyin.")

	dishes = get_all_dishes(connect_db_fn)
	dish_map = {f"{d['name']} ({d['category']} · {d['total_kcal']:g} kcal)": d for d in dishes}

	top_c1, top_c2 = ui.columns([2.2, 1.0], vertical_alignment="center")
	with top_c1:
		edit_choices = ["➕ Yeni Yemek Ekle"] + list(dish_map.keys())
		picked_lbl = ui.selectbox("Düzenlenecek veya İncelenecek Yemek", edit_choices, key="caf_dish_picker")
	with top_c2:
		if picked_lbl != "➕ Yeni Yemek Ekle" and ui.button("🗑️ Seçili Yemeği Sil", type="secondary", use_container_width=True, key="caf_btn_del_dish"):
			del_d = dish_map[picked_lbl]
			delete_dish_by_id(connect_db_fn, int(del_d["id"]))
			ui.session_state.pop("caf_dish_edit_draft", None)
			ui.success(f"'{del_d['name']}' yemekler bankasından silindi.")
			ui.rerun()

	editing_dish = dish_map.get(picked_lbl) if picked_lbl != "➕ Yeni Yemek Ekle" else None
	sel_sig = f"id:{editing_dish['id']}" if editing_dish else "new"
	if ui.session_state.get("caf_dish_loaded_sig") != sel_sig:
		ui.session_state["caf_dish_loaded_sig"] = sel_sig
		if editing_dish:
			ui.session_state["caf_d_name"] = editing_dish["name"]
			ui.session_state["caf_d_cat"] = editing_dish["category"] if editing_dish["category"] in DISH_CATEGORIES else DISH_CATEGORIES[1]
			ui.session_state["caf_d_port"] = float(editing_dish.get("portion_g") or 250.0)
			ui.session_state["caf_d_rec"] = str(editing_dish.get("recipe_notes") or "")
			ui.session_state["caf_dish_edit_draft"] = [dict(x) for x in (editing_dish.get("ingredients") or [])]
		else:
			ui.session_state["caf_d_name"] = ""
			ui.session_state["caf_d_cat"] = DISH_CATEGORIES[1]
			ui.session_state["caf_d_port"] = 250.0
			ui.session_state["caf_d_rec"] = ""
			ui.session_state["caf_dish_edit_draft"] = []

	with ui.container(border=True):
		dc1, dc2, dc3 = ui.columns([2.0, 1.3, 0.9])
		with dc1:
			d_name = ui.text_input("Yemek Adı*", placeholder="Örn: Mercimek Çorbası, İzmir Köfte, Fırın Makarna", key="caf_d_name")
		with dc2:
			d_cat = ui.selectbox("Kategori", DISH_CATEGORIES, key="caf_d_cat")
		with dc3:
			d_port = ui.number_input("Porsiyon (g)", min_value=20.0, step=10.0, key="caf_d_port")

		ai_c1, ai_c2 = ui.columns([2.2, 1.3], vertical_alignment="center")
		with ai_c1:
			ai_note = ui.text_input("AI Reçete İpucu (Opsiyonel)", placeholder="Örn: İlkokul porsiyonu, tereyağlı, glütensiz", key="caf_d_ai_note")
		with ai_c2:
			if ui.button("✨ Yapay Zekâ ile Malzeme, Kalori & Alerjen Çıkar", disabled=not is_on, type="primary" if is_on else "secondary", use_container_width=True, key="caf_btn_ai_dish"):
				if not str(d_name or "").strip():
					ui.warning("Lütfen önce yemek adını yazın.")
				else:
					with ui.spinner(f"Gemini AI '{d_name}' için 1 porsiyon malzeme ve besin analizini çıkarıyor..."):
						ok_d, res_d, used_m = gai.generate_dish_ingredients_and_nutrition_with_gemini(
							api_key=api_key, model=model, dish_name=d_name, category=d_cat, extra_notes=ai_note
						)
					if ok_d:
						ui.session_state["caf_dish_edit_draft"] = res_d.get("ingredients") or []
						if res_d.get("portion_g"):
							ui.session_state["caf_d_port"] = float(res_d["portion_g"])
						if res_d.get("recipe_notes"):
							ui.session_state["caf_d_rec"] = str(res_d["recipe_notes"])
						ui.success(f"✅ {len(ui.session_state['caf_dish_edit_draft'])} malzeme Gemini ({used_m}) ile hazırlandı! Aşağıdaki tablodan dilediğiniz gibi değiştirebilirsiniz.")
						ui.rerun()
					else:
						ui.error(f"AI hatası: {(res_d or {}).get('error', 'Yanıt alınamadı')}")

		draft_ings: list[dict[str, Any]] = ui.session_state.get("caf_dish_edit_draft") or []
		auto_scale = ui.checkbox("⚖️ Malzeme miktarını değiştirdiğimde kaloriyi oranla (AI kapalıyken de çalışır)", value=True, key="caf_d_autoscale")
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
			key="caf_dish_ing_editor",
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
		al_badges = " · ".join(format_allergen_badge(a) for a in dish_allergens) if dish_allergens else "Belirgin 14 temel alerjen yok"

		m1, m2, m3 = ui.columns(3)
		m1.metric("Porsiyon Kalorisi", f"{tot_k:g} kcal")
		m2.metric("Karbonhidrat (Diyabet Uyumu)", f"{tot_c:g} g")
		m3.metric("Tespit Edilen Alerjen", f"{len(dish_allergens)} tür")
		if dish_allergens:
			ui.warning(f"⚠️ **Yemek Alerjen Uyarısı:** {al_badges}")
			stu_hits = match_allergens_against_school(connect_db_fn, dish_allergens)
			if stu_hits:
				hit_txt = ", ".join(f"{mask_fn(h['full_name'], 'name')} ({h['class_name']} - {h['matched']})" for h in stu_hits[:8])
				ui.error(f"🚨 **Okulda Çakışan Alerjik Öğrenci/Kişi ({len(stu_hits)}):** {hit_txt}")
		else:
			ui.info(f"🟢 {al_badges}")

		d_rec = ui.text_input("Tarif / Pişirme Notu", key="caf_d_rec", placeholder="Kısa hazırlık veya servis notu...")
		if ui.button("💾 Yemeği ve Malzemelerini Kaydet / Güncelle", type="primary", use_container_width=True, key="caf_btn_save_dish"):
			try:
				saved_id = upsert_dish_with_ingredients(
					connect_db_fn,
					dish_id=int(editing_dish["id"]) if editing_dish else None,
					name=d_name,
					category=d_cat,
					portion_g=d_port,
					ingredients=scaled_ings,
					recipe_notes=d_rec,
					source="ai" if is_on else "manual",
				)
				ui.session_state["caf_dish_edit_draft"] = scaled_ings
				ui.session_state["caf_dish_loaded_sig"] = None
				ui.success(f"✅ '{d_name}' (#{saved_id}) yemekler bankasına kaydedildi ({tot_k:g} kcal).")
				ui.rerun()
			except Exception as exc:
				ui.error(str(exc))

	if dishes:
		ui.markdown("---")
		ui.markdown(f"##### 📚 Kayıtlı Yemekler Listesi ({len(dishes)})")
		df_list = pd.DataFrame([
			{
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
		ui.dataframe(df_list, hide_index=True, use_container_width=True)


def _render_tab_menu_planner(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()
	mask_fn = ctx_api["mask_kvkk_value"]

	ui.markdown("### 🗓️ İş Günü Bazlı Öğün Menü Planlayıcı & A4 Çıktı")
	dishes = get_all_dishes(connect_db_fn)
	dish_lbl_to_id = {f"{d['name']} ({d['total_kcal']:g} kcal)": int(d["id"]) for d in dishes}
	id_to_dish_lbl = {v: k for k, v in dish_lbl_to_id.items()}

	cfg_c1, cfg_c2, cfg_c3 = ui.columns([1.8, 1.1, 1.1])
	with cfg_c1:
		active_slots = ui.multiselect(
			"Planlanacak Öğünler",
			MEAL_SLOTS,
			default=["Öğle Yemeği"],
			key="caf_menu_active_slots",
		) or ["Öğle Yemeği"]
	with cfg_c2:
		inc_wknd = ui.checkbox("Hafta Sonlarını da Göster", value=False, key="caf_menu_inc_wknd")
	with cfg_c3:
		auto_sync_diab = ui.checkbox("💉 Diyabete Oto-Eşitle", value=True, key="caf_menu_auto_sync_diab")

	days_info = get_month_workdays_info(connect_db_fn, sel_ym, include_weekends=inc_wknd)
	workdays = [d for d in days_info if d.get("is_workday")]
	holidays_in_month = [d for d in days_info if d.get("is_holiday")]
	plans = get_menu_plans_for_month(connect_db_fn, sel_ym)

	tot_planned_days = len({p[0] for p in plans.keys()})
	ui.info(
		f"📅 **{_fmt_ym(sel_ym)} İş Günü Özeti:** Toplam **{len(workdays)} iş günü** · "
		f"Planlanmış gün: **{tot_planned_days}**"
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
								else:
									# Yeni yemeği bankaya otomatik veya custom_items olarak kaydet
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

	# Hafta Bazlı Menü Düzenleyici (Seçili ayın iş günleri hafta hafta düzenli kartlar)
	weeks = _group_month_days_by_week(days_info)
	if not weeks:
		ui.warning("Seçilen ay için aktif iş günü bulunamadı.")
		return

	wk_labels = [f"{i + 1}. Hafta ({w[0]['day_num']} – {w[-1]['day_num']} {TR_MONTH_NAMES[int(sel_ym[5:7])]})" for i, w in enumerate(weeks)]
	w_col1, w_col2 = ui.columns([2.0, 1.5], vertical_alignment="center")
	with w_col1:
		picked_wk_lbl = ui.selectbox("Düzenlenecek Haftayı Seçin (İş Günleri)", wk_labels, key="caf_menu_wk_pick")
	wk_idx = wk_labels.index(picked_wk_lbl) if picked_wk_lbl in wk_labels else 0
	with w_col2:
		if ui.button("💉 Tüm Ayı Diyabet Modülüne Senkronize Et", type="secondary", use_container_width=True, key="caf_btn_sync_diab_all"):
			cnt_s = sync_cafeteria_month_to_diabetes(connect_db_fn, sel_ym)
			ui.success(f"✅ {cnt_s} öğün İnsülin & Diyabet Yönetimi modülüyle eşitlendi.")

	for d in weeks[wk_idx]:
		d_iso = d["date_iso"]
		half_badge = f" · ⏳ {d['holiday_title']}" if d.get("is_half_day") else ""
		with ui.container(border=True):
			ui.markdown(f"**📆 {d['day_num']:02d}.{sel_ym[5:7]}.{sel_ym[:4]} — {d['weekday_name']}**{html.escape(half_badge)}")
			day_tot_kcal = 0.0
			for sl in active_slots:
				rec = plans.get((d_iso, sl))
				cur_ids = rec.get("dish_ids") or [] if rec else []
				cur_lbls = [id_to_dish_lbl[int(x)] for x in cur_ids if int(x) in id_to_dish_lbl]
				cur_cust_str = ", ".join(c.get("name", "") for c in (rec.get("custom_items") or []) if c.get("name")) if rec else ""

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
						placeholder="Örn: Tam Buğday Ekmek, Meyve",
						key=f"caf_ex_{d_iso}_{sl}",
					)
				with sc3:
					if ui.button("💾 Kaydet", key=f"caf_save_slot_{d_iso}_{sl}", use_container_width=True):
						chosen_ids = [dish_lbl_to_id[lbl] for lbl in sel_dish_lbls if lbl in dish_lbl_to_id]
						cust_list = [{"name": p.strip(), "kcal": 0.0, "carbs_g": 0.0} for p in str(extra_free or "").split(",") if p.strip()]
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
					day_tot_kcal += kv
					als = rec.get("allergens") or []
					al_disp = " · ".join(format_allergen_badge(a) for a in als) if als else "Alerjen uyarısı yok"
					hits = match_allergens_against_school(connect_db_fn, als) if als else []
					hit_s = f" | 🚨 **{len(hits)} alerjik öğrenci çakışması:** " + ", ".join(mask_fn(h["full_name"], "name") for h in hits[:4]) if hits else ""
					ui.caption(f"🔥 **{sl} Kalorisi:** `{kv:g} kcal` · **Karb:** `{_safe_float(rec.get('total_carbs_g')):g} g` · ⚠️ **Alerjenler:** {al_disp}{hit_s}")

			if len(active_slots) > 1 and day_tot_kcal > 0:
				ui.caption(f"📊 **Gün Toplamı:** `{day_tot_kcal:g} kcal`")

	# A4 Çıktı ve İmzalı Belge Oluşturucu
	ui.markdown("---")
	ui.markdown("#### 🖨️ Resmi A4 Menü Çıktısı (Aylık / Haftalık & İsim-Ünvan-İmza)")
	p_c1, p_c2, p_c3 = ui.columns([1.6, 1.0, 1.0])
	with p_c1:
		pr_scope = ui.radio("Çıktı Kapsamı", ["📅 Tüm Ay", "🗓️ Seçili Hafta"], horizontal=True, key="caf_print_scope")
	with p_c2:
		pr_kcal = ui.checkbox("Kalori (kcal) Göster", value=True, key="caf_print_kcal")
	with p_c3:
		pr_al = ui.checkbox("Alerjen Uyarısı Göster", value=True, key="caf_print_al")
	scope_code = "weekly" if "Hafta" in pr_scope else "monthly"

	signers = [
		(app_cfg.get("cafeteria_sig1_role", "Hazırlayan (Revir / Diyetisyen)"), app_cfg.get("cafeteria_sig1_name") or app_cfg.get("nurse_name", "")),
		(app_cfg.get("cafeteria_sig2_role", "Yemekhane Sorumlusu"), app_cfg.get("cafeteria_sig2_name", "")),
		(app_cfg.get("cafeteria_sig3_role", "Okul Müdürü (Onay)"), app_cfg.get("cafeteria_sig3_name") or app_cfg.get("principal_name", "")),
	]
	a4_html = build_menu_a4_html(
		connect_db_fn,
		year_month=sel_ym,
		scope=scope_code,
		week_index=wk_idx,
		active_slots=active_slots,
		show_kcal=pr_kcal,
		show_allergens=pr_al,
		include_weekends=inc_wknd,
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
		show_allergens=pr_al,
		include_weekends=inc_wknd,
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


def _render_tab_purchases_and_expenses(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	is_on, api_key, model = ctx_api["_get_active_gemini_config"]()

	ui.markdown("### 🚚 Gelen Ürünler, Fatura Kayıtları & Aylık / Toplam Gider Takibi")
	stats = get_expense_summary_stats(connect_db_fn, sel_ym)

	k1, k2, k3, k4 = ui.columns(4)
	k1.metric(f"{_fmt_ym(sel_ym)} Gideri", f"{stats['month_total']:,.2f} ₺", delta=f"{stats['month_invoice_count']} fatura")
	k2.metric(f"{sel_ym[:4]} Yılı Toplam Gider", f"{stats['year_total']:,.2f} ₺")
	k3.metric("Genel Toplam Gider", f"{stats['all_time_total']:,.2f} ₺")
	top_cat = max(stats["by_category"].items(), key=lambda x: x[1])[0] if stats["by_category"] else "—"
	k4.metric("En Yüksek Gider Kalemi", top_cat)

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

	# Yeni Fatura Girişi
	with ui.container(border=True):
		ui.markdown("##### ➕ Yeni Gelen Ürün / Fatura Girişi")
		fc1, fc2, fc3, fc4, fc5 = ui.columns([1.1, 1.0, 1.4, 1.3, 1.0])
		inv_dt = fc1.date_input("Fatura Tarihi", value=date.today(), key="caf_inv_date")
		inv_no = fc2.text_input("Fatura / Fiş No", key="caf_inv_no", placeholder="Örn: FTR-2026-041")
		inv_sup = fc3.text_input("Tedarikçi / Firma", key="caf_inv_sup", placeholder="Örn: Merkez Hal / Kasap")
		inv_cat = fc4.selectbox("Fatura Kategorisi", PURCHASE_CATEGORIES, key="caf_inv_cat")
		inv_st = fc5.selectbox("Ödeme Durumu", ["Ödendi", "Vadeli / Bekliyor"], key="caf_inv_st")

		draft_pi = ui.session_state.get("caf_inv_items_draft") or [
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
		ed_pi = ui.data_editor(
			df_pi,
			num_rows="dynamic",
			hide_index=True,
			use_container_width=True,
			key="caf_inv_items_editor",
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
		inv_notes = ui.text_input("Fatura Notu (Opsiyonel)", key="caf_inv_notes")
		if ui.button("💾 Faturayı ve Ürün Kalemlerini Kaydet", type="primary", use_container_width=True, key="caf_btn_save_inv"):
			items_to_save = ed_pi.to_dict("records")
			valid_items = [x for x in items_to_save if str(x.get("Ürün Adı") or "").strip()]
			if not valid_items:
				ui.warning("Lütfen tabloya en az 1 ürün satırı girin.")
			else:
				pid = upsert_purchase_invoice(
					connect_db_fn,
					invoice_date=inv_dt.isoformat(),
					invoice_no=inv_no,
					supplier_name=inv_sup,
					category=inv_cat,
					items=valid_items,
					payment_status=inv_st,
					notes=inv_notes,
				)
				ui.session_state.pop("caf_inv_items_draft", None)
				ui.success(f"✅ Fatura (#{pid}) ve {len(valid_items)} ürün kalemi giderlere işlendi.")
				ui.rerun()

	# Fatura Tarihine Göre Filtrelenmiş Liste
	ui.markdown("---")
	fl1, fl2, fl3 = ui.columns([1.2, 1.3, 1.5])
	scope_opt = fl1.selectbox("Dönem Filtresi", [f"Seçili Ay ({_fmt_ym(sel_ym)})", "Tüm Tarihler"], key="caf_purch_scope_flt")
	cat_flt = fl2.selectbox("Kategori Filtresi", ["Tümü"] + PURCHASE_CATEGORIES, key="caf_purch_cat_flt")
	q_flt = fl3.text_input("Ürün / Tedarikçi / Fatura No Ara", key="caf_purch_q_flt")

	ym_arg = sel_ym if "Seçili Ay" in scope_opt else "Tümü"
	purchases, items = get_purchases_and_items(connect_db_fn, year_month=ym_arg, category=cat_flt, query_text=q_flt)

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
				"Toplam Tutar (₺)": it["line_total"],
			}
			for it in items
		])
		ui.markdown(f"##### 🧾 Fatura Tarihine Göre Gelen Ürünler ({len(items)} kalem · Filtrelenmiş Toplam: **{sum(float(i['line_total'] or 0) for i in items):,.2f} ₺**)")
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

	if purchases:
		with ui.expander(f"🗑️ Fatura Kaydı Sil ({len(purchases)} fatura)"):
			p_opts = {
				f"#{pr['id']} — {pr['invoice_date']} · {pr.get('supplier_name') or 'Tedarikçi Yok'} · {float(pr.get('total_amount') or 0):,.2f} ₺": int(pr["id"])
				for pr in purchases
			}
			sel_del_p = ui.selectbox("Silinecek Fatura", list(p_opts.keys()), key="caf_del_p_sel")
			if ui.button("Seçili Faturayı Sil", key="caf_del_p_btn", type="secondary"):
				delete_purchase_by_id(connect_db_fn, p_opts[sel_del_p])
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

	# 1. Yemek Yiyen Ekleme & Aylık Öğün Ücreti Hesaplayıcı
	t_diners, t_fees, t_cash = ui.tabs([
		"🧑‍🎓 Yemek Yiyen Kişiler",
		"🧮 İş Günü × Öğün Ücreti & Tahakkuk",
		"💵 Ücret Tahsilatı (Getiren İsim) & Okul Kasa Teslimi",
	])

	diners = get_diners(connect_db_fn, active_only=True)
	with t_diners:
		with ui.expander("➕ Revirden Öğrenci/Personel Bağla veya Yeni Yemek Yiyen Ekle", expanded=not bool(diners)):
			ac1, ac2 = ui.columns(2)
			with ac1:
				with connect_db_fn() as conn:
					stu_all = conn.execute(
						"SELECT s.id, s.full_name, COALESCE(c.name, '—') AS cls FROM students s LEFT JOIN classes c ON c.id=s.class_id ORDER BY c.name, s.full_name"
					).fetchall()
				linked_sids = {int(d["student_id"]) for d in diners if d.get("student_id")}
				link_opts = {
					f"{mask_fn(r['full_name'], 'name')} · {r['cls']}": int(r["id"])
					for r in stu_all
					if int(r["id"]) not in linked_sids
				}
				picked_links = ui.multiselect("Kayıtlı Öğrenci / Personel Seç", list(link_opts.keys()), key="caf_link_stu_ms")
				link_meals = ui.number_input("Günlük Öğün Sayısı", min_value=1, max_value=3, value=1, key="caf_link_meals")
				if ui.button("📥 Seçilenleri Yemek Listesine Ekle", type="primary", use_container_width=True, disabled=not picked_links, key="caf_btn_link_stu"):
					cnt_add = import_school_people_as_diners(connect_db_fn, [link_opts[k] for k in picked_links], daily_meals_count=int(link_meals))
					ui.success(f"✅ {cnt_add} kişi yemek yiyenler listesine eklendi.")
					ui.rerun()
			with ac2:
				with ui.form("caf_form_add_diner", clear_on_submit=True):
					fn_new = ui.text_input("Serbest Kişi Adı Soyadı*")
					r_new = ui.selectbox("Statü", DINER_ROLE_TYPES)
					d_new = ui.text_input("Sınıf / Bölüm", placeholder="Örn: 3-A veya Mutfak")
					m_new = ui.number_input("Günlük Öğün", min_value=1, max_value=3, value=1)
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
			with ui.expander("🗑️ Yemek Listesinden Kişi Çıkar"):
				d_del_map = {f"#{d['id']} — {mask_fn(d['full_name'], 'name')} ({d.get('class_or_dept') or d['role_type']})": int(d["id"]) for d in diners}
				sel_dd = ui.selectbox("Çıkarılacak Kişi", list(d_del_map.keys()), key="caf_del_diner_sel")
				if ui.button("Kişiyi Yemek Listesinden Sil", key="caf_del_diner_btn"):
					delete_diner_by_id(connect_db_fn, d_del_map[sel_dd])
					ui.success("Kişi yemek listesinden kaldırıldı.")
					ui.rerun()

	with t_fees:
		days_info = get_month_workdays_info(connect_db_fn, sel_ym)
		auto_wd = sum(1 for d in days_info if d.get("is_workday"))
		fc1, fc2, fc3, fc4 = ui.columns(4)
		stu_price = fc1.number_input("Öğrenci 1 Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_meal_price"), 75.0), key="caf_fee_stu_p")
		stf_price = fc2.number_input("Öğretmen/Personel Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_staff_price"), 90.0), key="caf_fee_stf_p")
		wd_override = fc3.number_input("Ayın İş Günü Sayısı", min_value=0, max_value=31, value=auto_wd, key="caf_fee_wd_inp")
		with fc4:
			ui.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
			if ui.button("🧮 Aylık Ücretleri Hesapla / Güncelle", type="primary", use_container_width=True, key="caf_btn_calc_fees"):
				n_f = generate_or_refresh_monthly_fees(
					connect_db_fn,
					sel_ym,
					student_meal_price=float(stu_price),
					staff_meal_price=float(stf_price),
					workdays_override=int(wd_override),
				)
				ui.success(f"✅ {n_f} kişi için {int(wd_override)} iş günü üzerinden {_fmt_ym(sel_ym)} ücretleri hesaplandı.")
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
					"İndirim (₺)": float(b["discount_or_adjustment"]),
					"Beklenen (₺)": float(b["expected_amount"]),
					"Ödenen (₺)": float(b["paid_amount"]),
					"Kalan (₺)": float(b["remaining_amount"]),
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
					"Planlı Gün": ui.column_config.NumberColumn("Katılım Günü (Düzenlenebilir)", min_value=0, max_value=31, step=1),
					"Günlük Öğün": ui.column_config.NumberColumn("Günlük Öğün", min_value=1, max_value=3, step=1),
					"Öğün (₺)": ui.column_config.NumberColumn("Öğün (₺)", min_value=0.0, step=5.0),
					"İndirim (₺)": ui.column_config.NumberColumn("İndirim / Devir (₺)", step=10.0),
					"Beklenen (₺)": ui.column_config.NumberColumn("Beklenen (₺)", disabled=True),
					"Ödenen (₺)": ui.column_config.NumberColumn("Ödenen (₺)", disabled=True),
					"Kalan (₺)": ui.column_config.NumberColumn("Kalan (₺)", disabled=True),
					"Durum": ui.column_config.TextColumn("Durum", disabled=True),
				},
			)
			if ui.button("💾 Gün / Ücret Düzeltmelerini Kaydet", type="secondary", use_container_width=True, key="caf_btn_save_fee_edits"):
				for _, r_f in ed_bal.iterrows():
					update_monthly_fee_row(
						connect_db_fn,
						int(r_f["fee_id"]),
						planned_days=int(r_f["Planlı Gün"]),
						meals_per_day=int(r_f["Günlük Öğün"]),
						unit_meal_price=float(r_f["Öğün (₺)"]),
						discount_or_adjustment=float(r_f["İndirim (₺)"]),
					)
				ui.success("✅ Kişi bazlı katılım günü ve ücret düzeltmeleri güncellendi.")
				ui.rerun()
		else:
			ui.info(f"{_fmt_ym(sel_ym)} için henüz aylık aidat oluşturulmamış. Önce kişi ekleyin ve '🧮 Aylık Ücretleri Hesapla' butonuna basın.")

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

		# Kasa Hareketleri Tablosu & Teslim Tutanağı İndir
		txs = cash["month_transactions"]
		ui.markdown(f"##### 📜 {_fmt_ym(sel_ym)} Kasa Hareketleri ({len(txs)})")
		if txs:
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
				for t in txs
			])
			ui.dataframe(df_tx, hide_index=True, use_container_width=True)
			tx_dl1, tx_dl2 = ui.columns(2)
			with tx_dl1:
				ui.download_button(
					"📝 Okula Kasa Teslim Tutanağı İndir (.docx)",
					data=lambda: build_cash_handover_docx(connect_db_fn, year_month=sel_ym, school_meta=app_cfg),
					file_name=f"Yemekhane_Kasa_Teslim_Tutanagi_{sel_ym}.docx",
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					type="primary",
					use_container_width=True,
					key="caf_dl_cash_docx",
				)
			with tx_dl2:
				with ui.expander("🗑️ Yanlış Girilen Kasa Hareketini Sil"):
					tx_opts = {f"#{int(t['id'])} — {t['tx_date']} · {float(t['amount']):,.2f} ₺ ({t.get('for_person_name') or ''})": int(t["id"]) for t in txs}
					s_tx = ui.selectbox("Silinecek Kasa Kaydı", list(tx_opts.keys()), key="caf_del_tx_sel")
					if ui.button("Kasa Kaydını Sil", key="caf_del_tx_btn"):
						delete_cash_transaction(connect_db_fn, tx_opts[s_tx])
						ui.success("Kasa kaydı silindi.")
						ui.rerun()
		else:
			ui.caption("Bu dönem için henüz tahsilat veya teslim kaydı yok.")


def _render_tab_settings_and_exports(ctx_api: dict[str, Any], sel_ym: str, app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	save_cfg_fn = ctx_api.get("save_app_settings")
	ui.markdown("### ⚙️ Yemekhane Ayarları, A4 İmza Ünvanları & Toplu Excel Çıktısı")

	with ui.container(border=True):
		ui.markdown("##### ✍️ A4 Menü Çıktısı Alt İmza & Varsayılan Öğün Tarifesi")
		with ui.form("caf_form_settings"):
			sc1, sc2 = ui.columns(2)
			def_stu_p = sc1.number_input("Varsayılan Öğrenci Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_meal_price"), 75.0))
			def_stf_p = sc2.number_input("Varsayılan Öğretmen/Personel Öğün Ücreti (₺)", min_value=0.0, step=5.0, value=_safe_float(app_cfg.get("cafeteria_default_staff_price"), 90.0))

			sg1, sg2, sg3 = ui.columns(3)
			s1_r = sg1.text_input("1. İmzacı Ünvan", value=app_cfg.get("cafeteria_sig1_role", "Hazırlayan (Revir / Diyetisyen)"))
			s1_n = sg1.text_input("1. İmzacı Ad Soyad", value=app_cfg.get("cafeteria_sig1_name") or app_cfg.get("nurse_name", ""))
			s2_r = sg2.text_input("2. İmzacı Ünvan (Boş bırakılırsa 2 imza olur)", value=app_cfg.get("cafeteria_sig2_role", "Yemekhane Sorumlusu"))
			s2_n = sg2.text_input("2. İmzacı Ad Soyad", value=app_cfg.get("cafeteria_sig2_name", ""))
			s3_r = sg3.text_input("3. İmzacı Ünvan", value=app_cfg.get("cafeteria_sig3_role", "Okul Müdürü (Onay)"))
			s3_n = sg3.text_input("3. İmzacı Ad Soyad", value=app_cfg.get("cafeteria_sig3_name") or app_cfg.get("principal_name", ""))

			if ui.form_submit_button("💾 Yemekhane Ayarlarını Kaydet", type="primary", use_container_width=True) and callable(save_cfg_fn):
				save_cfg_fn({
					"cafeteria_default_meal_price": str(def_stu_p),
					"cafeteria_default_staff_price": str(def_stf_p),
					"cafeteria_sig1_role": s1_r.strip(),
					"cafeteria_sig1_name": s1_n.strip(),
					"cafeteria_sig2_role": s2_r.strip(),
					"cafeteria_sig2_name": s2_n.strip(),
					"cafeteria_sig3_role": s3_r.strip(),
					"cafeteria_sig3_name": s3_n.strip(),
				})
				if callable(ctx_api.get("clear_all_data_caches")):
					ctx_api["clear_all_data_caches"]()
				ui.success("✅ Yemekhane tarife ve imza ayarları kaydedildi.")
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
	ym_opts, cur_ym = _month_options()

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
