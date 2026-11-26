"""
Revir Defteri - Okul Yurdu / Pansiyon Sağlık ve Revir Yönetim Modülü (school_dormitory.py)

Kapsam:
1. Pansiyon Yerleşim & Oda-Yatak Haritası (Kurum öğrencisi seçimi + Başka okuldan gelen yatılı öğrenci kaydı,
   Blok/Kat/Oda/Yatak künyesi, bulaşıcı hastalıkta otomatik aynı oda ve komşu oda/kat yakın temaslı taraması)
2. Revir Yatış (Müşahede Yatağı, Saatlik Gece Vital Takibi), Gece Nöbet Devir Teslim Defteri & 112 Acil Sevk/Refakat Tutanağı
3. 7/24 Akşam/Sabah Aç/Hafta Sonu İlaç Çizelgesi, Kilitli İlaç Dolabı & Oda İlaç Denetimi,
   Kronik Hastalık Oda Bazlı Belletmen Gece Acil Kartı & 23:00 / 03:00 Gece Kan Şekeri İzlemi
4. Yurt İzolasyon (Karantina) Odası, Evci Çıkış/Dönüş Sağlık Taraması, Revir İstirahat/Etüt Muafiyeti,
   Rapor Bitiş Dönüş Uyarısı, Geçici Spor Muafiyeti & Turnuva Sağlık Çantası Tutanağı
5. Yemekhane 72 Saatlik Şahit Numune, Personel Portör/Hijyen, Toplu Gıda Zehirlenmesi Kırmızı Erken Uyarısı,
   Kız Pansiyonu Dismenore & Hijyenik Ped Destek Stoğu, Oda Çarşaf/Kişisel Bakım & Özel Eğitim Gece Cihaz Teslimi
6. Su Deposu Klor/Analiz, Ortak Alan Dezenfeksiyon, Gizlilik Esaslı Uyku/Enürezis/Ev Özlemi (PDR) Takibi,
   Hafta Sonu Veli Ziyaret Görüşmeleri & MEB Müfettiş / İlçe Sağlık Otomatik Puanlamalı Pansiyon Denetim Dosyası
"""
from __future__ import annotations

import html
import io
import json
import re
import sqlite3
from datetime import date, datetime, timedelta
from typing import Any, Callable

import pandas as pd

import revir_ui as ui


# ==============================================================================
# 1. SABİTLER VE KLİNİK SÖZLÜKLER
# ==============================================================================

BOARDING_STATUS_OPTIONS: list[str] = [
	"Yatılı",
	"Gündüzlü",
	"Geçici Misafir Yatılı",
]

RESIDENT_SOURCE_OPTIONS: list[str] = [
	"Kurum Öğrencisi",
	"Başka Okul Öğrencisi",
]

ENTRY_REPORT_STATUS_OPTIONS: list[str] = [
	"Tam (Yurtta Kalabilir)",
	"Eksik / Bekleniyor",
	"Yenilenecek / Süresi Doldu",
]

CONTAGIOUS_KEYWORDS: tuple[str, ...] = (
	"uyuz", "skabiyez", "scabies", "bit", "pediküloz", "pedikuloz",
	"suçiçeği", "sucicegi", "varisella", "grip", "influenza", "covid",
	"gastroenterit", "ishal", "kusma", "rotavirüs", "norovirüs",
	"konjonktivit", "kızamık", "kizamik", "kabakulak", "el ayak ağız",
	"streptokok", "beta", "tonsillit", "impetigo", "mantar",
)

GI_POISONING_KEYWORDS: tuple[str, ...] = (
	"bulantı", "bulanti", "kusma", "ishal", "karın ağrısı", "karin agrisi",
	"gıda zehirlenmesi", "gida zehirlenmesi", "gastroenterit", "kramp",
)

DORM_MED_SLOTS_24H: list[str] = [
	"07:00 (Sabah Aç Karnına / Kahvaltı Öncesi)",
	"08:30 (Kahvaltı Sonrası)",
	"12:30 (Öğle)",
	"16:00 (İkindi / Okul Dönüşü)",
	"19:00 (Akşam Yemeği)",
	"21:00 (Gece Yatmadan Önce)",
	"23:00 (Gece Özel Doz)",
]

OBSERVATION_OUTCOMES: list[str] = [
	"Müşahedede (Aktif Yatış)",
	"İyileşti — Odasına Döndü",
	"Sabah Poliklinik / Aile Hekimi Sevk",
	"112 / Acil Hastane Sevk",
	"Veliye Teslim / Ev İstirahati",
]

EXEMPTION_TYPES: list[str] = [
	"Sabah Derslerinden Muaf (Revir İstirahatli)",
	"Akşam Etüdünden Muaf",
	"Tam Gün Revir İstirahatli",
	"Hastane İstirahat Raporlu (Ev / Memleket)",
	"Yurt İzolasyon İstirahati",
]

SLEEP_PSYCH_CATEGORIES: list[str] = [
	"Enürezis (Gece Alt Islatma)",
	"Uyurgezerlik (Somnambulizm)",
	"Ev Özlemi / Uyum Güçlüğü (Homesickness)",
	"Gece Panik Atağı / Anksiyete",
	"İnsomnia / Uyku Düzeni Bozukluğu",
	"Diğer Psikososyal Gözlem",
]

FACILITY_CHECK_TYPES: list[str] = [
	"Su Deposu Serbest Klor Ölçümü (0.2–0.5 ppm)",
	"Su Deposu Mikrobiyolojik / Kimyasal Analiz",
	"Ortak Banyo / WC Periyodik Dezenfeksiyon",
	"Yatakhane / Koğuş Havalandırma & Dezenfeksiyon",
	"Haşere ve Kemirgen İlaçlama Kontrolü",
]

HYGIENE_SUPPORT_ITEMS: list[str] = [
	"Hijyenik Ped (Gece / Normal)",
	"Yedek İç Çamaşırı",
	"Koruyucu Yatak Alezi (Tek Kullanımlık)",
	"Sıcak Su Torbası & Koruyucu Kılıf",
	"Bireysel Hijyen / Bit Şampuanı Seti",
]

EXTERNAL_DORM_CLASS_NAME = "Pansiyon (Dış Okul)"


def _tr_fold(text: Any) -> str:
	s = str(text or "").replace("İ", "i").replace("I", "ı").lower()
	return (
		s.replace("ğ", "g")
		.replace("ü", "u")
		.replace("ş", "s")
		.replace("ö", "o")
		.replace("ç", "c")
		.replace("ı", "i")
	)


def _now_iso() -> str:
	return datetime.now().isoformat(timespec="minutes")


def _safe_float(val: Any, default: float | None = None) -> float | None:
	if val is None or str(val).strip() == "":
		return default
	try:
		return float(str(val).replace(",", ".").strip())
	except Exception:
		return default


def _parse_room_int(room_no: str) -> int | None:
	m = re.search(r"\d+", str(room_no or ""))
	return int(m.group(0)) if m else None


# ==============================================================================
# 2. VERİTABANI ŞEMASI (İDEMPOTENT — TÜM YURT ALT MODÜLLERİ)
# ==============================================================================

def ensure_dormitory_tables(connection: sqlite3.Connection) -> None:
	"""Okul yurdu / pansiyon modülünün tüm tablolarını idempotent olarak oluşturur."""
	connection.executescript(
		"""
		CREATE TABLE IF NOT EXISTS dorm_residents (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			student_id INTEGER REFERENCES students(id) ON DELETE SET NULL,
			resident_source TEXT NOT NULL DEFAULT 'Kurum Öğrencisi',
			origin_school_name TEXT DEFAULT '',
			full_name TEXT NOT NULL,
			sex TEXT DEFAULT 'E',
			student_no TEXT DEFAULT '',
			class_name TEXT DEFAULT '',
			birth_date TEXT DEFAULT '',
			guardian_name TEXT DEFAULT '',
			guardian_phone TEXT DEFAULT '',
			home_location TEXT DEFAULT '',
			boarding_status TEXT NOT NULL DEFAULT 'Yatılı',
			block_name TEXT NOT NULL DEFAULT 'A Blok',
			floor_no TEXT NOT NULL DEFAULT '1. Kat',
			room_no TEXT NOT NULL DEFAULT '101',
			bed_bunk_no TEXT DEFAULT 'Yatak-1',
			entry_health_report_status TEXT NOT NULL DEFAULT 'Tam (Yurtta Kalabilir)',
			entry_report_date TEXT DEFAULT '',
			uses_devices TEXT DEFAULT '',
			chronic_or_allergy_note TEXT DEFAULT '',
			is_active INTEGER NOT NULL DEFAULT 1,
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL,
			updated_at TEXT NOT NULL,
			UNIQUE(student_id)
		);
		CREATE INDEX IF NOT EXISTS idx_dorm_res_room ON dorm_residents(is_active, block_name, floor_no, room_no);

		CREATE TABLE IF NOT EXISTS dorm_observation_stays (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			bed_label TEXT NOT NULL DEFAULT 'Yatak 1',
			admitted_date TEXT NOT NULL,
			admitted_time TEXT NOT NULL,
			complaint TEXT NOT NULL,
			initial_intervention TEXT DEFAULT '',
			outcome_status TEXT NOT NULL DEFAULT 'Müşahedede (Aktif Yatış)',
			discharged_date TEXT DEFAULT '',
			discharged_time TEXT DEFAULT '',
			discharge_note TEXT DEFAULT '',
			admitted_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);
		CREATE INDEX IF NOT EXISTS idx_dorm_obs_status ON dorm_observation_stays(outcome_status, admitted_date);

		CREATE TABLE IF NOT EXISTS dorm_observation_vitals (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			stay_id INTEGER NOT NULL REFERENCES dorm_observation_stays(id) ON DELETE CASCADE,
			recorded_at TEXT NOT NULL,
			temperature_c REAL,
			systolic_mmhg REAL,
			diastolic_mmhg REAL,
			pulse_bpm REAL,
			spo2_pct REAL,
			glucose_mg_dl REAL,
			care_note TEXT DEFAULT '',
			recorded_by TEXT DEFAULT ''
		);

		CREATE TABLE IF NOT EXISTS dorm_shift_handovers (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			shift_date TEXT NOT NULL,
			shift_type TEXT NOT NULL DEFAULT '17:00 - 08:00 Gece Nöbeti',
			from_staff TEXT NOT NULL,
			to_staff TEXT NOT NULL,
			summary_text TEXT NOT NULL,
			critical_students_note TEXT DEFAULT '',
			admin_notified INTEGER NOT NULL DEFAULT 1,
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_emergency_transfers (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			transfer_date TEXT NOT NULL,
			transfer_time TEXT NOT NULL,
			reason_complaint TEXT NOT NULL,
			transport_mode TEXT NOT NULL DEFAULT '112 Ambulans',
			hospital_name TEXT NOT NULL,
			escort_staff_name TEXT NOT NULL,
			guardian_notified_time TEXT DEFAULT '',
			guardian_contact_person TEXT DEFAULT '',
			return_time TEXT DEFAULT '',
			epicrisis_summary TEXT DEFAULT '',
			rest_days_given INTEGER NOT NULL DEFAULT 0,
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_med_schedules (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			medication_name TEXT NOT NULL,
			dose_text TEXT NOT NULL DEFAULT '1 ölçek / 1 tb',
			slot_label TEXT NOT NULL,
			include_weekends INTEGER NOT NULL DEFAULT 1,
			cabinet_no TEXT DEFAULT 'Dolap-1',
			shelf_no TEXT DEFAULT 'Raf-A',
			delivered_by_guardian TEXT DEFAULT '',
			box_qty REAL NOT NULL DEFAULT 1.0,
			expiry_date TEXT DEFAULT '',
			is_active INTEGER NOT NULL DEFAULT 1,
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_med_dose_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			schedule_id INTEGER NOT NULL REFERENCES dorm_med_schedules(id) ON DELETE CASCADE,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			dose_date TEXT NOT NULL,
			slot_label TEXT NOT NULL,
			given_time TEXT NOT NULL,
			status TEXT NOT NULL DEFAULT 'İمza Karşılığı İçirildi',
			administered_by TEXT DEFAULT '',
			student_signed INTEGER NOT NULL DEFAULT 1,
			notes TEXT DEFAULT '',
			UNIQUE(schedule_id, dose_date, slot_label)
		);

		CREATE TABLE IF NOT EXISTS dorm_room_med_inspections (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			inspection_date TEXT NOT NULL,
			block_name TEXT NOT NULL,
			room_no TEXT NOT NULL,
			unauthorized_med_found INTEGER NOT NULL DEFAULT 0,
			confiscated_items TEXT DEFAULT '',
			action_taken TEXT DEFAULT '',
			inspected_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_night_glucose_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			log_date TEXT NOT NULL,
			target_slot TEXT NOT NULL DEFAULT '03:00 Gece Ölçümü',
			measured_time TEXT NOT NULL,
			glucose_mg_dl REAL NOT NULL,
			ketone_note TEXT DEFAULT '',
			intervention_note TEXT DEFAULT '',
			measured_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_isolation_stays (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			isolation_room_no TEXT NOT NULL DEFAULT 'İzolasyon-1',
			disease_or_reason TEXT NOT NULL,
			start_date TEXT NOT NULL,
			expected_end_date TEXT NOT NULL,
			actual_end_date TEXT DEFAULT '',
			room_meal_service INTEGER NOT NULL DEFAULT 1,
			diet_type TEXT DEFAULT 'Normal / Bol Sıvı',
			status TEXT NOT NULL DEFAULT 'İzolasyonda',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_leave_returns (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			leave_type TEXT NOT NULL DEFAULT 'Hafta Sonu Evci',
			depart_date TEXT NOT NULL,
			expected_return_date TEXT NOT NULL,
			actual_return_date TEXT DEFAULT '',
			return_screening_status TEXT NOT NULL DEFAULT 'Dönüş Bekleniyor',
			fever_or_symptom_note TEXT DEFAULT '',
			medical_report_checked INTEGER NOT NULL DEFAULT 0,
			screened_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_rest_and_exemptions (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			exemption_kind TEXT NOT NULL,
			start_date TEXT NOT NULL,
			end_date TEXT NOT NULL,
			reason_diagnosis TEXT NOT NULL,
			returned_confirmed INTEGER NOT NULL DEFAULT 0,
			issued_by TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_sports_exemptions (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			injury_or_reason TEXT NOT NULL,
			start_date TEXT NOT NULL,
			end_date TEXT NOT NULL,
			scope_text TEXT NOT NULL DEFAULT 'Beden Eğitimi Dersi ve Yurt Sportif Antrenmanları',
			doctor_or_hospital TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_tournament_kits (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			event_name TEXT NOT NULL,
			destination TEXT NOT NULL,
			coach_or_teacher TEXT NOT NULL,
			given_date TEXT NOT NULL,
			expected_return_date TEXT DEFAULT '',
			returned_date TEXT DEFAULT '',
			kit_name TEXT NOT NULL DEFAULT 'Turnuva Acil Müdahale Çantası #1',
			contents_summary TEXT DEFAULT 'Elastik bandaj, soğuk sprey, buz aküsü, steril gazlı bez, antiseptik, flaster, makas, eldiven',
			return_missing_notes TEXT DEFAULT '',
			status TEXT NOT NULL DEFAULT 'Teslim Edildi (Sahada)',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_food_samples (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			sample_date TEXT NOT NULL,
			meal_slot TEXT NOT NULL DEFAULT 'Öğle Yemeği',
			sampled_dishes TEXT NOT NULL,
			taken_time TEXT NOT NULL DEFAULT '12:15',
			fridge_temp_c REAL DEFAULT 4.0,
			retention_until_iso TEXT NOT NULL,
			taken_by TEXT DEFAULT '',
			hygiene_check_ok INTEGER NOT NULL DEFAULT 1,
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_staff_health_checks (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			staff_name TEXT NOT NULL,
			role_title TEXT NOT NULL DEFAULT 'Aşçı / Mutfak Personeli',
			check_type TEXT NOT NULL DEFAULT 'Periyodik Portör / Sağlık Muayenesi & Hijyen Belgesi',
			exam_date TEXT NOT NULL,
			valid_until_date TEXT NOT NULL,
			result_status TEXT NOT NULL DEFAULT 'Uygun (Çalışabilir)',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_dysmenorrhea_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			log_date TEXT NOT NULL,
			pain_severity INTEGER NOT NULL DEFAULT 5,
			hot_pack_applied INTEGER NOT NULL DEFAULT 1,
			rest_given INTEGER NOT NULL DEFAULT 1,
			medication_note TEXT DEFAULT '',
			referral_advised INTEGER NOT NULL DEFAULT 0,
			notes TEXT DEFAULT '',
			recorded_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_hygiene_support_stock (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			item_name TEXT UNIQUE NOT NULL,
			quantity REAL NOT NULL DEFAULT 0.0,
			min_threshold REAL NOT NULL DEFAULT 10.0,
			unit TEXT NOT NULL DEFAULT 'adet/paket',
			updated_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_personal_hygiene_checks (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			check_date TEXT NOT NULL,
			block_name TEXT NOT NULL,
			room_no TEXT NOT NULL,
			linen_changed INTEGER NOT NULL DEFAULT 1,
			bath_completed INTEGER NOT NULL DEFAULT 1,
			nail_hair_clean INTEGER NOT NULL DEFAULT 1,
			room_ventilated INTEGER NOT NULL DEFAULT 1,
			checked_by TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_device_night_custody (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			custody_date TEXT NOT NULL,
			device_type TEXT NOT NULL DEFAULT 'Gözlük / İşitme Cihazı',
			received_night_time TEXT NOT NULL DEFAULT '22:00',
			storage_box_no TEXT DEFAULT 'Kutu-1',
			returned_morning_time TEXT DEFAULT '',
			condition_ok INTEGER NOT NULL DEFAULT 1,
			staff_name TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_facility_sanitation_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			check_date TEXT NOT NULL,
			check_type TEXT NOT NULL,
			location_area TEXT NOT NULL DEFAULT 'Ana Su Deposu / Ortak WC-Banyo',
			measured_value TEXT DEFAULT '',
			is_compliant INTEGER NOT NULL DEFAULT 1,
			next_due_date TEXT DEFAULT '',
			performed_by TEXT DEFAULT '',
			notes TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_psychosocial_sleep_logs (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			log_date TEXT NOT NULL,
			category TEXT NOT NULL,
			confidential_observation TEXT NOT NULL,
			protective_plan TEXT DEFAULT '',
			pdr_referred INTEGER NOT NULL DEFAULT 0,
			pdr_feedback_note TEXT DEFAULT '',
			recorded_by TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);

		CREATE TABLE IF NOT EXISTS dorm_parent_visit_notes (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			resident_id INTEGER NOT NULL REFERENCES dorm_residents(id) ON DELETE CASCADE,
			visit_date TEXT NOT NULL,
			visit_time TEXT DEFAULT '14:00',
			visitor_name TEXT NOT NULL,
			visitor_relation TEXT DEFAULT 'Veli',
			topics_discussed TEXT NOT NULL,
			agreed_actions TEXT DEFAULT '',
			staff_name TEXT DEFAULT '',
			created_at TEXT NOT NULL
		);
		"""
	)
	# Varsayılan kız/pansiyon destek stok kalemlerini ilk kurulumda tohumla
	cnt_h = int(connection.execute("SELECT COUNT(*) FROM dorm_hygiene_support_stock").fetchone()[0])
	if cnt_h == 0:
		now_s = _now_iso()
		for it_nm, def_q, min_q in [
			("Hijyenik Ped (Gece / Normal)", 40.0, 15.0),
			("Yedek İç Çamaşırı", 15.0, 5.0),
			("Koruyucu Yatak Alezi (Tek Kullanımlık)", 25.0, 10.0),
			("Sıcak Su Torbası & Koruyucu Kılıf", 8.0, 3.0),
			("Bireysel Hijyen / Bit Şampuanı Seti", 10.0, 4.0),
		]:
			connection.execute(
				"INSERT OR IGNORE INTO dorm_hygiene_support_stock(item_name, quantity, min_threshold, unit, updated_at) VALUES (?, ?, ?, 'adet', ?)",
				(it_nm, def_q, min_q, now_s),
			)


# ==============================================================================
# 3. ÖĞRENCİ & DIŞ OKUL PANSİYONER KAYIT, ODA-YATAK VE TEMASLI KÜMELENME MOTORU
# ==============================================================================

def enroll_existing_school_students_to_dorm(
	connect_db_fn: Callable[[], sqlite3.Connection],
	student_ids: list[int],
	*,
	block_name: str = "A Blok",
	floor_no: str = "1. Kat",
	room_no: str = "101",
	bed_prefix: str = "Yatak",
	boarding_status: str = "Yatılı",
	entry_health_report_status: str = "Tam (Yurtta Kalabilir)",
	school_name_default: str = "",
) -> int:
	"""Okuldaki mevcut kayıtlı öğrencilerden seçilenleri yurt odasına/yatağına yerleştirir."""
	clean_ids = [int(x) for x in (student_ids or []) if str(x).isdigit()]
	if not clean_ids:
		return 0
	now_s = _now_iso()
	added = 0
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		ph = ",".join("?" for _ in clean_ids)
		rows = conn.execute(
			f"""SELECT s.*, COALESCE(c.name, '—') AS cls_name,
			           COALESCE(hp.allergies, '') AS hp_al,
			           COALESCE(hp.chronic_conditions, '') AS hp_ch,
			           COALESCE(hp.medical_devices, '') AS hp_dev
			    FROM students s
			    LEFT JOIN classes c ON c.id = s.class_id
			    LEFT JOIN health_profiles hp ON hp.student_id = s.id
			    WHERE s.id IN ({ph})""",
			clean_ids,
		).fetchall()
		for idx, r in enumerate(rows, start=1):
			sid = int(r["id"])
			home_loc = " / ".join(filter(None, [str(r["province"] or "").strip(), str(r["district"] or "").strip()]))
			med_summary = " · ".join(filter(None, [str(r["hp_al"]).strip(), str(r["hp_ch"]).strip()]))
			bed_lbl = f"{bed_prefix}-{idx}" if len(rows) > 1 else f"{bed_prefix}-1"
			conn.execute(
				"""INSERT INTO dorm_residents(
					student_id, resident_source, origin_school_name, full_name, sex, student_no,
					class_name, birth_date, guardian_name, guardian_phone, home_location,
					boarding_status, block_name, floor_no, room_no, bed_bunk_no,
					entry_health_report_status, uses_devices, chronic_or_allergy_note,
					is_active, created_at, updated_at
				) VALUES (?, 'Kurum Öğrencisi', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
				ON CONFLICT(student_id) DO UPDATE SET
					is_active=1,
					boarding_status=excluded.boarding_status,
					block_name=excluded.block_name,
					floor_no=excluded.floor_no,
					room_no=excluded.room_no,
					bed_bunk_no=excluded.bed_bunk_no,
					entry_health_report_status=excluded.entry_health_report_status,
					updated_at=excluded.updated_at""",
				(
					sid,
					str(school_name_default or "Kurum Öğrencisi").strip(),
					str(r["full_name"]),
					str(r["sex"] or "E"),
					str(r["student_no"] or ""),
					str(r["cls_name"] or ""),
					str(r["birth_date"] or ""),
					str(r["guardian_name"] or ""),
					str(r["guardian_phone"] or ""),
					home_loc,
					boarding_status,
					str(block_name or "A Blok").strip(),
					str(floor_no or "1. Kat").strip(),
					str(room_no or "101").strip(),
					bed_lbl,
					entry_health_report_status,
					str(r["hp_dev"] or "").strip(),
					med_summary,
					now_s,
					now_s,
				),
			)
			added += 1
	return added


def enroll_external_school_student_to_dorm(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	full_name: str,
	origin_school_name: str,
	sex: str = "E",
	class_name: str = "9. Sınıf",
	student_no: str = "",
	birth_date: str = "",
	guardian_name: str = "",
	guardian_phone: str = "",
	home_location: str = "",
	block_name: str = "A Blok",
	floor_no: str = "1. Kat",
	room_no: str = "101",
	bed_bunk_no: str = "Yatak-1",
	boarding_status: str = "Yatılı",
	entry_health_report_status: str = "Tam (Yurtta Kalabilir)",
	entry_report_date: str = "",
	uses_devices: str = "",
	allergies: str = "",
	chronic_conditions: str = "",
	notes: str = "",
	create_linked_revir_profile: bool = True,
) -> int:
	"""Başka okuldan gelen yatılı öğrenci için hem yurt kaydı hem de (isteğe bağlı) revir kişi/sağlık profili açar."""
	nm = re.sub(r"\s+", " ", str(full_name or "").strip())
	orig_sch = re.sub(r"\s+", " ", str(origin_school_name or "").strip())
	if not nm:
		raise ValueError("Öğrenci adı soyadı boş bırakılamaz.")
	if not orig_sch:
		raise ValueError("Başka okuldan gelen öğrenci için 'Geldiği Okul Adı' zorunludur.")

	now_s = _now_iso()
	linked_sid: int | None = None
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		if create_linked_revir_profile:
			conn.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (EXTERNAL_DORM_CLASS_NAME,))
			ext_cid = int(
				conn.execute("SELECT id FROM classes WHERE name=?", (EXTERNAL_DORM_CLASS_NAME,)).fetchone()["id"]
			)
			disp_note = f"[Dış Okul Pansiyoner: {orig_sch} · {class_name}] {notes}".strip()
			ex_stu = conn.execute(
				"SELECT id FROM students WHERE class_id=? AND full_name=?",
				(ext_cid, nm),
			).fetchone()
			if ex_stu:
				linked_sid = int(ex_stu["id"])
				conn.execute(
					"UPDATE students SET sex=?, student_no=?, birth_date=?, guardian_name=?, guardian_phone=?, notes=? WHERE id=?",
					(sex, student_no, birth_date, guardian_name, guardian_phone, disp_note, linked_sid),
				)
			else:
				cur_s = conn.execute(
					"""INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date, guardian_name, guardian_phone, notes)
					   VALUES (?, 'Öğrenci', ?, ?, ?, ?, ?, ?, ?)""",
					(ext_cid, student_no, nm, sex, birth_date, guardian_name, guardian_phone, disp_note),
				)
				linked_sid = int(cur_s.lastrowid)
			if allergies.strip() or chronic_conditions.strip() or uses_devices.strip():
				conn.execute(
					"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, allergies, chronic_conditions, medical_devices, updated_at)
					   VALUES (?, ?, ?, ?, ?, ?, ?)
					   ON CONFLICT(student_id) DO UPDATE SET
					       allergies=excluded.allergies,
					       chronic_conditions=excluded.chronic_conditions,
					       medical_devices=excluded.medical_devices,
					       updated_at=excluded.updated_at""",
					(linked_sid, guardian_name, guardian_phone, allergies.strip(), chronic_conditions.strip(), uses_devices.strip(), now_s),
				)

		med_comb = " · ".join(filter(None, [allergies.strip(), chronic_conditions.strip()]))
		cur = conn.execute(
			"""INSERT INTO dorm_residents(
				student_id, resident_source, origin_school_name, full_name, sex, student_no,
				class_name, birth_date, guardian_name, guardian_phone, home_location,
				boarding_status, block_name, floor_no, room_no, bed_bunk_no,
				entry_health_report_status, entry_report_date, uses_devices,
				chronic_or_allergy_note, is_active, notes, created_at, updated_at
			) VALUES (?, 'Başka Okul Öğrencisi', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
			ON CONFLICT(student_id) DO UPDATE SET
				resident_source='Başka Okul Öğrencisi',
				origin_school_name=excluded.origin_school_name,
				class_name=excluded.class_name,
				block_name=excluded.block_name,
				floor_no=excluded.floor_no,
				room_no=excluded.room_no,
				bed_bunk_no=excluded.bed_bunk_no,
				boarding_status=excluded.boarding_status,
				entry_health_report_status=excluded.entry_health_report_status,
				uses_devices=excluded.uses_devices,
				chronic_or_allergy_note=excluded.chronic_or_allergy_note,
				is_active=1,
				updated_at=excluded.updated_at""",
			(
				linked_sid,
				orig_sch,
				nm,
				str(sex or "E"),
				str(student_no or "").strip(),
				str(class_name or "").strip(),
				str(birth_date or "").strip(),
				str(guardian_name or "").strip(),
				str(guardian_phone or "").strip(),
				str(home_location or "").strip(),
				boarding_status,
				str(block_name or "A Blok").strip(),
				str(floor_no or "1. Kat").strip(),
				str(room_no or "101").strip(),
				str(bed_bunk_no or "Yatak-1").strip(),
				entry_health_report_status,
				str(entry_report_date or "").strip(),
				str(uses_devices or "").strip(),
				med_comb,
				str(notes or "").strip(),
				now_s,
				now_s,
			),
		)
		if cur.lastrowid:
			return int(cur.lastrowid)
		row_id = conn.execute("SELECT id FROM dorm_residents WHERE student_id=?", (linked_sid,)).fetchone()
		return int(row_id["id"]) if row_id else 0


def update_dorm_resident_placement(
	connect_db_fn: Callable[[], sqlite3.Connection],
	resident_id: int,
	*,
	block_name: str,
	floor_no: str,
	room_no: str,
	bed_bunk_no: str,
	boarding_status: str = "Yatılı",
	entry_health_report_status: str = "Tam (Yurtta Kalabilir)",
	uses_devices: str | None = None,
	is_active: bool = True,
) -> None:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		if uses_devices is None:
			conn.execute(
				"""UPDATE dorm_residents
				   SET block_name=?, floor_no=?, room_no=?, bed_bunk_no=?,
				       boarding_status=?, entry_health_report_status=?, is_active=?, updated_at=?
				   WHERE id=?""",
				(
					str(block_name).strip(),
					str(floor_no).strip(),
					str(room_no).strip(),
					str(bed_bunk_no).strip(),
					boarding_status,
					entry_health_report_status,
					1 if is_active else 0,
					_now_iso(),
					int(resident_id),
				),
			)
		else:
			conn.execute(
				"""UPDATE dorm_residents
				   SET block_name=?, floor_no=?, room_no=?, bed_bunk_no=?,
				       boarding_status=?, entry_health_report_status=?, uses_devices=?, is_active=?, updated_at=?
				   WHERE id=?""",
				(
					str(block_name).strip(),
					str(floor_no).strip(),
					str(room_no).strip(),
					str(bed_bunk_no).strip(),
					boarding_status,
					entry_health_report_status,
					str(uses_devices).strip(),
					1 if is_active else 0,
					_now_iso(),
					int(resident_id),
				),
			)


def get_dorm_residents(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	active_only: bool = True,
	boarding_only: bool = False,
) -> list[dict[str, Any]]:
	"""Yurt öğrencilerini sağlık/diyabet/epilepsi risk bayraklarıyla zenginleştirilmiş olarak döndürür."""
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		try:
			import insulin_diabetes as idm
			idm.ensure_diabetes_tables(conn)
		except Exception:
			pass
		where_parts = []
		if active_only:
			where_parts.append("dr.is_active = 1")
		if boarding_only:
			where_parts.append("dr.boarding_status != 'Gündüzlü'")
		where_sql = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
		rows = conn.execute(
			f"""SELECT dr.*,
			           COALESCE(hp.allergies, '') AS hp_allergies,
			           COALESCE(hp.chronic_conditions, '') AS hp_chronic,
			           COALESCE(hp.seizure_emergency_protocol, '') AS hp_seizure,
			           COALESCE(hp.emergency_notes, '') AS hp_emerg,
			           COALESCE(hp.medical_devices, '') AS hp_devices,
			           COALESCE(dp.is_diabetic, 0) AS is_diabetic,
			           COALESCE(dp.emergency_protocol, '') AS dp_hypo,
			           COALESCE(dp.glucagon_location, '') AS dp_glucagon
			    FROM dorm_residents dr
			    LEFT JOIN health_profiles hp ON hp.student_id = dr.student_id
			    LEFT JOIN diabetes_profiles dp ON dp.student_id = dr.student_id
			    {where_sql}
			    ORDER BY dr.block_name ASC, dr.floor_no ASC, dr.room_no ASC, dr.bed_bunk_no ASC, dr.full_name ASC"""
		).fetchall()
	out: list[dict[str, Any]] = []
	for r in rows:
		d = dict(r)
		comb_risk = " · ".join(
			dict.fromkeys(
				x for x in [
					str(d.get("hp_allergies") or "").strip(),
					str(d.get("hp_chronic") or "").strip(),
					str(d.get("chronic_or_allergy_note") or "").strip(),
				] if x
			)
		)
		d["combined_clinical_risk"] = comb_risk
		eff_dev = str(d.get("uses_devices") or "").strip() or str(d.get("hp_devices") or "").strip()
		d["effective_devices"] = eff_dev
		out.append(d)
	out.sort(key=lambda x: ui.tr_alpha_sort_key(x.get("full_name")))
	return out


def get_room_and_neighbor_contacts(
	connect_db_fn: Callable[[], sqlite3.Connection],
	index_resident_id: int,
	*,
	neighbor_room_delta: int = 2,
) -> dict[str, Any]:
	"""Bulaşıcı hastalık şüphesinde aynı odada kalan (1. derece) ve aynı kat/komşu odalarda kalan (2. derece) öğrencileri listeler."""
	residents = get_dorm_residents(connect_db_fn, active_only=True)
	by_id = {int(r["id"]): r for r in residents}
	idx_res = by_id.get(int(index_resident_id))
	if not idx_res:
		return {"index_resident": None, "roommates": [], "floor_neighbors": []}

	b_nm = _tr_fold(idx_res["block_name"])
	f_no = _tr_fold(idx_res["floor_no"])
	r_no = str(idx_res["room_no"]).strip()
	r_int = _parse_room_int(r_no)

	roommates: list[dict[str, Any]] = []
	floor_neighbors: list[dict[str, Any]] = []
	for cand in residents:
		if int(cand["id"]) == int(index_resident_id):
			continue
		if cand.get("boarding_status") == "Gündüzlü":
			continue
		same_block = _tr_fold(cand["block_name"]) == b_nm
		same_floor = same_block and (_tr_fold(cand["floor_no"]) == f_no)
		same_room = same_block and (str(cand["room_no"]).strip().lower() == r_no.lower())
		if same_room:
			c_copy = dict(cand)
			c_copy["contact_degree"] = "🔴 1. Derece Yakın Temaslı (Aynı Oda)"
			roommates.append(c_copy)
		elif same_floor:
			c_int = _parse_room_int(cand["room_no"])
			is_adj = (r_int is not None and c_int is not None and abs(c_int - r_int) <= neighbor_room_delta)
			c_copy = dict(cand)
			c_copy["contact_degree"] = (
				"🟠 2. Derece Komşu Oda Temaslısı"
				if is_adj
				else "🟡 Aynı Kat / Ortak Banyo-WC Temaslısı"
			)
			floor_neighbors.append(c_copy)
	return {
		"index_resident": idx_res,
		"roommates": roommates,
		"floor_neighbors": floor_neighbors,
	}


def detect_recent_contagious_room_clusters(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	within_days: int = 10,
) -> list[dict[str, Any]]:
	"""Son N günde revir vizitlerinde veya izolasyonda bulaşıcı bulgusu olan yatılı öğrencileri ve oda temaslı sayılarını bulur."""
	since_iso = (date.today() - timedelta(days=max(1, within_days))).isoformat()
	residents = get_dorm_residents(connect_db_fn, active_only=True, boarding_only=True)
	by_sid = {int(r["student_id"]): r for r in residents if r.get("student_id")}
	by_rid = {int(r["id"]): r for r in residents}
	alerts: list[dict[str, Any]] = []
	seen_rids: set[int] = set()

	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		iso_rows = conn.execute(
			"SELECT * FROM dorm_isolation_stays WHERE status='İzolasyonda' OR start_date >= ? ORDER BY start_date DESC",
			(since_iso,),
		).fetchall()
		for ir in iso_rows:
			rid = int(ir["resident_id"])
			res = by_rid.get(rid)
			if not res or rid in seen_rids:
				continue
			seen_rids.add(rid)
			ct = get_room_and_neighbor_contacts(connect_db_fn, rid)
			alerts.append({
				"resident_id": rid,
				"full_name": res["full_name"],
				"room_label": f"{res['block_name']} · {res['floor_no']} · Oda {res['room_no']}",
				"finding": str(ir["disease_or_reason"]),
				"event_date": str(ir["start_date"]),
				"roommate_count": len(ct["roommates"]),
				"neighbor_count": len(ct["floor_neighbors"]),
			})
		if by_sid:
			ph = ",".join("?" for _ in by_sid)
			v_rows = conn.execute(
				f"SELECT student_id, visit_date, complaint, symptoms FROM visits WHERE visit_date >= ? AND student_id IN ({ph}) ORDER BY visit_date DESC",
				[since_iso, *by_sid.keys()],
			).fetchall()
			for vr in v_rows:
				res = by_sid.get(int(vr["student_id"]))
				if not res or int(res["id"]) in seen_rids:
					continue
				txt = _tr_fold(f"{vr['complaint'] or ''} {vr['symptoms'] or ''}")
				if any(_tr_fold(kw) in txt for kw in CONTAGIOUS_KEYWORDS):
					rid = int(res["id"])
					seen_rids.add(rid)
					ct = get_room_and_neighbor_contacts(connect_db_fn, rid)
					alerts.append({
						"resident_id": rid,
						"full_name": res["full_name"],
						"room_label": f"{res['block_name']} · {res['floor_no']} · Oda {res['room_no']}",
						"finding": f"{vr['complaint']} ({vr['symptoms'] or ''})".strip(),
						"event_date": str(vr["visit_date"]),
						"roommate_count": len(ct["roommates"]),
						"neighbor_count": len(ct["floor_neighbors"]),
					})
	return alerts


# ==============================================================================
# 4. MÜŞAHEDE YATAĞI, NÖBET DEVİR DEFTERİ, 112 SEVK & TOPLU GIDA ZEHİRLENMESİ ALARMI
# ==============================================================================

def admit_to_observation_bed(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	resident_id: int,
	bed_label: str,
	admitted_date: str,
	admitted_time: str,
	complaint: str,
	initial_intervention: str = "",
	admitted_by: str = "",
) -> int:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		cur = conn.execute(
			"""INSERT INTO dorm_observation_stays(
				resident_id, bed_label, admitted_date, admitted_time, complaint,
				initial_intervention, outcome_status, admitted_by, created_at
			) VALUES (?, ?, ?, ?, ?, ?, 'Müşahedede (Aktif Yatış)', ?, ?)""",
			(
				int(resident_id),
				str(bed_label or "Yatak 1").strip(),
				str(admitted_date)[:10],
				str(admitted_time)[:5],
				str(complaint).strip(),
				str(initial_intervention).strip(),
				str(admitted_by).strip(),
				_now_iso(),
			),
		)
		return int(cur.lastrowid)


def add_observation_vital_log(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	stay_id: int,
	recorded_at: str,
	temperature_c: float | None = None,
	systolic_mmhg: float | None = None,
	diastolic_mmhg: float | None = None,
	pulse_bpm: float | None = None,
	spo2_pct: float | None = None,
	glucose_mg_dl: float | None = None,
	care_note: str = "",
	recorded_by: str = "",
) -> int:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		cur = conn.execute(
			"""INSERT INTO dorm_observation_vitals(
				stay_id, recorded_at, temperature_c, systolic_mmhg, diastolic_mmhg,
				pulse_bpm, spo2_pct, glucose_mg_dl, care_note, recorded_by
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				int(stay_id),
				str(recorded_at).strip() or _now_iso(),
				temperature_c,
				systolic_mmhg,
				diastolic_mmhg,
				pulse_bpm,
				spo2_pct,
				glucose_mg_dl,
				str(care_note or "").strip(),
				str(recorded_by or "").strip(),
			),
		)
		return int(cur.lastrowid)


def discharge_observation_stay(
	connect_db_fn: Callable[[], sqlite3.Connection],
	stay_id: int,
	*,
	outcome_status: str,
	discharged_date: str | None = None,
	discharged_time: str | None = None,
	discharge_note: str = "",
) -> None:
	d_dt = str(discharged_date or date.today().isoformat())[:10]
	d_tm = str(discharged_time or datetime.now().strftime("%H:%M"))[:5]
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		conn.execute(
			"""UPDATE dorm_observation_stays
			   SET outcome_status=?, discharged_date=?, discharged_time=?, discharge_note=?
			   WHERE id=?""",
			(outcome_status, d_dt, d_tm, str(discharge_note or "").strip(), int(stay_id)),
		)


def build_auto_night_handover_summary(
	connect_db_fn: Callable[[], sqlite3.Connection],
	shift_date: str,
) -> str:
	"""Gece (17:00 sonrası ve ertesi sabah 08:30 öncesi) gerçekleşen müşahede, ilaç, glikoz ve 112 kayıtlarını özetler."""
	dt_s = str(shift_date or date.today().isoformat())[:10]
	try:
		next_dt = (date.fromisoformat(dt_s) + timedelta(days=1)).isoformat()
	except Exception:
		next_dt = dt_s
	lines: list[str] = []
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		obs = conn.execute(
			"""SELECT os.*, dr.full_name, dr.room_no
			   FROM dorm_observation_stays os
			   JOIN dorm_residents dr ON dr.id = os.resident_id
			   WHERE os.admitted_date IN (?, ?) OR os.outcome_status='Müşahedede (Aktif Yatış)'""",
			(dt_s, next_dt),
		).fetchall()
		if obs:
			lines.append(f"• Revir Müşahede ({len(obs)} kayıt): " + "; ".join(
				f"{r['full_name']} (Oda {r['room_no']}, {r['bed_label']} - {r['complaint']} -> {r['outcome_status']})"
				for r in obs
			))
		emg = conn.execute(
			"""SELECT et.*, dr.full_name, dr.room_no
			   FROM dorm_emergency_transfers et
			   JOIN dorm_residents dr ON dr.id = et.resident_id
			   WHERE et.transfer_date IN (?, ?)""",
			(dt_s, next_dt),
		).fetchall()
		if emg:
			lines.append(f"• 112 / Acil Hastane Sevk ({len(emg)}): " + "; ".join(
				f"{r['full_name']} ({r['transfer_time']} -> {r['hospital_name']}, Refakat: {r['escort_staff_name']})"
				for r in emg
			))
		meds = conn.execute(
			"SELECT COUNT(*) FROM dorm_med_dose_logs WHERE dose_date IN (?, ?)",
			(dt_s, next_dt),
		).fetchone()[0]
		if int(meds or 0) > 0:
			lines.append(f"• Akşam / Gece / Sabah Yurt İlaç Dozu: {int(meds)} doz imza karşılığı uygulandı.")
		glc = conn.execute(
			"""SELECT ng.*, dr.full_name
			   FROM dorm_night_glucose_logs ng
			   JOIN dorm_residents dr ON dr.id = ng.resident_id
			   WHERE ng.log_date IN (?, ?)""",
			(dt_s, next_dt),
		).fetchall()
		if glc:
			lines.append("• Gece Diyabet Glikoz İzlemi: " + "; ".join(
				f"{g['full_name']} ({g['measured_time']}: {g['glucose_mg_dl']:g} mg/dL)"
				for g in glc
			))
	if not lines:
		return f"{dt_s} gece nöbetinde acil sevk veya aktif müşahede yatışı yaşanmamıştır; rutin oda ve ilaç kontrolleri tamamlanmıştır."
	return "\n".join(lines)


def save_shift_handover(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	shift_date: str,
	shift_type: str,
	from_staff: str,
	to_staff: str,
	summary_text: str,
	critical_students_note: str = "",
	admin_notified: bool = True,
) -> int:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		cur = conn.execute(
			"""INSERT INTO dorm_shift_handovers(
				shift_date, shift_type, from_staff, to_staff, summary_text,
				critical_students_note, admin_notified, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				str(shift_date)[:10],
				str(shift_type).strip(),
				str(from_staff).strip(),
				str(to_staff).strip(),
				str(summary_text).strip(),
				str(critical_students_note).strip(),
				1 if admin_notified else 0,
				_now_iso(),
			),
		)
		return int(cur.lastrowid)


def record_emergency_transfer(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	resident_id: int,
	transfer_date: str,
	transfer_time: str,
	reason_complaint: str,
	transport_mode: str,
	hospital_name: str,
	escort_staff_name: str,
	guardian_notified_time: str = "",
	guardian_contact_person: str = "",
	return_time: str = "",
	epicrisis_summary: str = "",
	rest_days_given: int = 0,
) -> int:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		cur = conn.execute(
			"""INSERT INTO dorm_emergency_transfers(
				resident_id, transfer_date, transfer_time, reason_complaint, transport_mode,
				hospital_name, escort_staff_name, guardian_notified_time, guardian_contact_person,
				return_time, epicrisis_summary, rest_days_given, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				int(resident_id),
				str(transfer_date)[:10],
				str(transfer_time)[:5],
				str(reason_complaint).strip(),
				str(transport_mode).strip(),
				str(hospital_name).strip(),
				str(escort_staff_name).strip(),
				str(guardian_notified_time).strip(),
				str(guardian_contact_person).strip(),
				str(return_time).strip(),
				str(epicrisis_summary).strip(),
				max(0, int(rest_days_given or 0)),
				_now_iso(),
			),
		)
		tid = int(cur.lastrowid)
		# Eğer hastaneden istirahat raporu verildiyse otomatik olarak Raporlu/İstirahat senkronizasyonuna da ekle
		if int(rest_days_given or 0) > 0:
			try:
				end_d = (date.fromisoformat(str(transfer_date)[:10]) + timedelta(days=int(rest_days_given) - 1)).isoformat()
			except Exception:
				end_d = str(transfer_date)[:10]
			conn.execute(
				"""INSERT INTO dorm_rest_and_exemptions(
					resident_id, exemption_kind, start_date, end_date, reason_diagnosis, issued_by, notes, created_at
				) VALUES (?, 'Hastane İstirahat Raporlu (Ev / Memleket)', ?, ?, ?, ?, ?, ?)""",
				(
					int(resident_id),
					str(transfer_date)[:10],
					end_d,
					f"{hospital_name}: {epicrisis_summary or reason_complaint}",
					escort_staff_name,
					"112/Acil sevk sonrası otomatik oluşturuldu",
					_now_iso(),
				),
			)
		return tid


def detect_foodborne_outbreak_clusters(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	target_date: str | None = None,
	window_hours: float = 6.0,
	min_cases: int = 2,
) -> dict[str, Any]:
	"""Aynı gün / 2-6 saatlik pencere içinde birden fazla yatılı öğrencide bulantı, kusma, karın ağrısı, ishal varsa Kırmızı Alarm üretir."""
	dt_s = str(target_date or date.today().isoformat())[:10]
	residents = get_dorm_residents(connect_db_fn, active_only=True, boarding_only=True)
	by_sid = {int(r["student_id"]): r for r in residents if r.get("student_id")}
	by_rid = {int(r["id"]): r for r in residents}
	events: list[dict[str, Any]] = []

	def _time_to_min(tm: str) -> int:
		m = re.match(r"^(\d{1,2}):(\d{2})", str(tm or "12:00").strip())
		return (int(m.group(1)) * 60 + int(m.group(2))) if m else 720

	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		if by_sid:
			ph = ",".join("?" for _ in by_sid)
			v_rows = conn.execute(
				f"SELECT student_id, visit_date, visit_time, complaint, symptoms FROM visits WHERE visit_date=? AND student_id IN ({ph})",
				[dt_s, *by_sid.keys()],
			).fetchall()
			for vr in v_rows:
				txt = _tr_fold(f"{vr['complaint'] or ''} {vr['symptoms'] or ''}")
				if any(_tr_fold(k) in txt for k in GI_POISONING_KEYWORDS):
					res = by_sid[int(vr["student_id"])]
					events.append({
						"resident_id": int(res["id"]),
						"full_name": res["full_name"],
						"room_no": f"{res['block_name']} / {res['room_no']}",
						"time": str(vr["visit_time"] or "12:00")[:5],
						"minutes": _time_to_min(vr["visit_time"]),
						"detail": f"{vr['complaint']} {vr['symptoms'] or ''}".strip(),
						"source": "Revir Başvurusu",
					})
		obs_rows = conn.execute(
			"SELECT resident_id, admitted_time, complaint FROM dorm_observation_stays WHERE admitted_date=?",
			(dt_s,),
		).fetchall()
		for ob in obs_rows:
			res = by_rid.get(int(ob["resident_id"]))
			if not res:
				continue
			if any(e["resident_id"] == int(res["id"]) for e in events):
				continue
			txt = _tr_fold(ob["complaint"])
			if any(_tr_fold(k) in txt for k in GI_POISONING_KEYWORDS):
				events.append({
					"resident_id": int(res["id"]),
					"full_name": res["full_name"],
					"room_no": f"{res['block_name']} / {res['room_no']}",
					"time": str(ob["admitted_time"] or "20:00")[:5],
					"minutes": _time_to_min(ob["admitted_time"]),
					"detail": str(ob["complaint"]),
					"source": "Revir Müşahede",
				})

	events.sort(key=lambda x: x["minutes"])
	win_min = int(window_hours * 60)
	best_cluster: list[dict[str, Any]] = []
	for i in range(len(events)):
		base_m = events[i]["minutes"]
		grp = [e for e in events if base_m <= e["minutes"] <= base_m + win_min]
		uniq_rids = {g["resident_id"] for g in grp}
		if len(uniq_rids) > len({b["resident_id"] for b in best_cluster}):
			best_cluster = grp

	is_alarm = len({b["resident_id"] for b in best_cluster}) >= min_cases
	suspected_meal = ""
	if best_cluster:
		first_m = best_cluster[0]["minutes"]
		if first_m <= 13 * 60:
			suspected_meal = "Sabah Kahvaltısı"
		elif first_m <= 18 * 60 + 30:
			suspected_meal = "Öğle Yemeği"
		else:
			suspected_meal = "Akşam Yemeği / İkindi Ara Öğünü"

	return {
		"alarm": is_alarm,
		"date": dt_s,
		"case_count": len({b["resident_id"] for b in best_cluster}),
		"cases": best_cluster,
		"suspected_meal": suspected_meal,
	}


# ==============================================================================
# 5. GECE ACİL ODA RİSK KARTLARI, RAPOR DÖNÜŞ UYARILARI & MEB ÖZ-DENETİM MOTORU
# ==============================================================================

def get_due_return_and_dorm_alerts(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	ref_date: str | None = None,
) -> dict[str, Any]:
	"""Bugün yurda/derse dönüş yapması gereken raporlu, evci veya izolasyon bitişli öğrencileri ve aktif uyarı sayılarını döndürür."""
	today_s = str(ref_date or date.today().isoformat())[:10]
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		active_obs = int(
			conn.execute("SELECT COUNT(*) FROM dorm_observation_stays WHERE outcome_status='Müşahedede (Aktif Yatış)'").fetchone()[0]
		)
		active_iso = int(
			conn.execute("SELECT COUNT(*) FROM dorm_isolation_stays WHERE status='İzolasyonda'").fetchone()[0]
		)
		due_rest = [
			dict(r)
			for r in conn.execute(
				"""SELECT re.*, dr.full_name, dr.block_name, dr.room_no, dr.class_name
				   FROM dorm_rest_and_exemptions re
				   JOIN dorm_residents dr ON dr.id = re.resident_id
				   WHERE re.returned_confirmed=0 AND re.end_date <= ?
				   ORDER BY re.end_date ASC""",
				(today_s,),
			).fetchall()
		]
		due_iso = [
			dict(r)
			for r in conn.execute(
				"""SELECT isy.*, dr.full_name, dr.block_name, dr.room_no
				   FROM dorm_isolation_stays isy
				   JOIN dorm_residents dr ON dr.id = isy.resident_id
				   WHERE isy.status='İzolasyonda' AND isy.expected_end_date <= ?
				   ORDER BY isy.expected_end_date ASC""",
				(today_s,),
			).fetchall()
		]
		due_leave = [
			dict(r)
			for r in conn.execute(
				"""SELECT lr.*, dr.full_name, dr.block_name, dr.room_no
				   FROM dorm_leave_returns lr
				   JOIN dorm_residents dr ON dr.id = lr.resident_id
				   WHERE lr.return_screening_status='Dönüş Bekleniyor' AND lr.expected_return_date <= ?
				   ORDER BY lr.expected_return_date ASC""",
				(today_s,),
			).fetchall()
		]
	food_alarm = detect_foodborne_outbreak_clusters(connect_db_fn, target_date=today_s)
	badge_total = active_obs + active_iso + len(due_rest) + len(due_iso) + (1 if food_alarm["alarm"] else 0)
	return {
		"active_observation_count": active_obs,
		"active_isolation_count": active_iso,
		"due_rest_returns": due_rest,
		"due_isolation_ends": due_iso,
		"due_leave_returns": due_leave,
		"food_alarm": food_alarm,
		"badge_total": badge_total,
	}


def get_chronic_night_risk_residents(
	connect_db_fn: Callable[[], sqlite3.Connection],
) -> list[dict[str, Any]]:
	"""Tip 1 diyabet, epilepsi, ağır astım veya anafilaksi riski taşıyan yatılı öğrencileri oda bazlı Gece Acil Müdahale protokolüyle döndürür."""
	residents = get_dorm_residents(connect_db_fn, active_only=True, boarding_only=True)
	risk_cards: list[dict[str, Any]] = []
	for r in residents:
		is_diab = int(r.get("is_diabetic") or 0) == 1
		comb_f = _tr_fold(
			f"{r.get('combined_clinical_risk') or ''} {r.get('hp_seizure') or ''} {r.get('hp_emerg') or ''}"
		)
		has_epilepsy = any(k in comb_f for k in ("epilepsi", "nobet", "nöbet", "konvulziyon"))
		has_asthma = any(k in comb_f for k in ("astim", "astım", "inhaler", "ventolin", "nebul"))
		has_anaphylaxis = any(k in comb_f for k in ("anafilaksi", "epipen", "adrenalin", "fistik", "ari alerji", "alerji"))
		if not (is_diab or has_epilepsy or has_asthma or has_anaphylaxis or str(r.get("hp_emerg") or "").strip()):
			continue

		protocols: list[str] = []
		tags: list[str] = []
		if is_diab:
			tags.append("💙 Tip 1 Diyabet (Gece Hipoglisemi Riski)")
			custom_h = str(r.get("dp_hypo") or "").strip()
			protocols.append(
				custom_h
				or "Gece terleme, titreme, kabus veya bilinç bulanıklığında DERHAL parmak ucu kan şekeri ölçün (<70 mg/dL ise bilinci açıksa 15g basit şeker/meyve suyu verin; bilinci kapalıysa ağızdan bir şey vermeyip yan yatırarak 112 ve nöbetçi hemşireyi çağırın). 23:00 ve 03:00 glikoz kontrolü!"
			)
		if has_epilepsy:
			tags.append("⚡ Epilepsi / Nöbet Riski")
			sz_p = str(r.get("hp_seizure") or "").strip()
			protocols.append(
				sz_p
				or "Gece kasılma/nöbet başlarsa başının altına yumuşak yastık koyun, yan (güvenli) pozisyona çevirin, ağzına cisim sokmayın, süreyi tutun; >5 dk sürerse 112 ve revir hemşiresini arayın (Alt ranzada yatırılmalıdır)."
			)
		if has_asthma:
			tags.append("🫁 Astım / Solunum")
			protocols.append("Gece öksürük/nefes darlığında dik oturtun, kurtarıcı inhalerini uygulatın ve revire haber verin.")
		if has_anaphylaxis:
			tags.append("🥜 Alerji / Anafilaksi")
			protocols.append(f"Bilinen risk: {r.get('combined_clinical_risk')}. Nefes darlığı/dilde şişmede Adrenalin Oto-Enjektörü + 112.")
		if str(r.get("hp_emerg") or "").strip():
			protocols.append(f"Özel Acil Not: {r['hp_emerg']}")

		risk_cards.append({
			"resident_id": int(r["id"]),
			"full_name": r["full_name"],
			"block_name": r["block_name"],
			"floor_no": r["floor_no"],
			"room_no": r["room_no"],
			"bed_bunk_no": r["bed_bunk_no"],
			"guardian_phone": r.get("guardian_phone") or "—",
			"risk_tags": " · ".join(tags) or "🟠 Özel Klinik İzlem",
			"night_protocol": " | ".join(protocols),
		})
	risk_cards.sort(key=lambda x: (x["block_name"], x["floor_no"], x["room_no"], x["full_name"]))
	return risk_cards


def evaluate_meb_dorm_inspection(
	connect_db_fn: Callable[[], sqlite3.Connection],
) -> dict[str, Any]:
	"""MEB Müfettiş / İlçe Sağlık Pansiyon Denetim maddelerini canlı veritabanı verileriyle otomatik puanlar."""
	today_iso = date.today().isoformat()
	since_30d = (date.today() - timedelta(days=30)).isoformat()
	since_7d = (date.today() - timedelta(days=7)).isoformat()
	items: list[dict[str, Any]] = []

	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		total_res = int(conn.execute("SELECT COUNT(*) FROM dorm_residents WHERE is_active=1 AND boarding_status!='Gündüzlü'").fetchone()[0])
		missing_rep = int(
			conn.execute(
				"SELECT COUNT(*) FROM dorm_residents WHERE is_active=1 AND boarding_status!='Gündüzlü' AND entry_health_report_status!='Tam (Yurtta Kalabilir)'"
			).fetchone()[0]
		)
		items.append({
			"Madde No": "1",
			"Denetim Başlığı": "Yatılı Öğrenci 'Yurtta Kalmasında Sakınca Yoktur' Sağlık Raporları",
			"Durum": "✅ Tam" if (total_res > 0 and missing_rep == 0) else ("🟡 Eksik Var" if total_res > 0 else "⚪ Kayıt Yok"),
			"Detay": f"Toplam yatılı: {total_res} · Raporu eksik/yenilenecek: {missing_rep}",
			"Puan": 15 if (total_res > 0 and missing_rep == 0) else (8 if missing_rep <= 2 and total_res > 0 else 5),
		})

		# 2. Kilitli ilaç dolabı & oda ilaç denetimi
		cab_cnt = int(conn.execute("SELECT COUNT(*) FROM dorm_med_schedules WHERE is_active=1").fetchone()[0])
		rm_insp = int(conn.execute("SELECT COUNT(*) FROM dorm_room_med_inspections WHERE inspection_date >= ?", (since_30d,)).fetchone()[0])
		items.append({
			"Madde No": "2",
			"Denetim Başlığı": "Merkezi Kilitli Revir İlaç Dolabı & Aylık Oda İlaç Denetimi",
			"Durum": "✅ Uygun" if rm_insp > 0 else "🟡 Oda Denetimi Girilmemiş",
			"Detay": f"Dolapta kayıtlı yurt ilacı: {cab_cnt} · Son 30 günde oda ilaç denetimi: {rm_insp}",
			"Puan": 15 if rm_insp > 0 else 7,
		})

		# 3. SKT geçmiş revir malzemesi veya ilaç kontrolü
		exp_inv = int(
			conn.execute(
				"SELECT COUNT(*) FROM infirmary_inventory WHERE expiry_date IS NOT NULL AND trim(expiry_date)!='' AND expiry_date < ?",
				(today_iso,),
			).fetchone()[0]
		)
		exp_med = int(
			conn.execute(
				"SELECT COUNT(*) FROM student_medications WHERE status='Aktif' AND expiry_date IS NOT NULL AND trim(expiry_date)!='' AND expiry_date < ?",
				(today_iso,),
			).fetchone()[0]
		) + int(
			conn.execute(
				"SELECT COUNT(*) FROM dorm_med_schedules WHERE is_active=1 AND trim(COALESCE(expiry_date,''))!='' AND expiry_date < ?",
				(today_iso,),
			).fetchone()[0]
		)
		skt_clean = (exp_inv == 0 and exp_med == 0)
		items.append({
			"Madde No": "3",
			"Denetim Başlığı": "Son Kullanma Tarihi (SKT) Geçmiş İlaç ve Revir Malzemesi Kontrolü",
			"Durum": "✅ Temiz (SKT Geçmiş Yok)" if skt_clean else "🔴 SKT Geçmiş Kalem Var!",
			"Detay": f"SKT geçmiş sarf: {exp_inv} · SKT geçmiş aktif öğrenci ilacı: {exp_med}",
			"Puan": 15 if skt_clean else 0,
		})

		# 4. Soğuk zincir ısı takibi
		fr_cnt = int(conn.execute("SELECT COUNT(*) FROM fridge_temp_logs WHERE log_date >= ?", (since_7d,)).fetchone()[0])
		items.append({
			"Madde No": "4",
			"Denetim Başlığı": "Soğuk Zincir Buzdolabı & Revir Ortam Isı İzlemi (Son 7 Gün)",
			"Durum": "✅ Güncel" if fr_cnt > 0 else "🟡 Son 7 Günde Kayıt Yok",
			"Detay": f"Son 7 günde {fr_cnt} ısı/nem kontrol kaydı mevcut.",
			"Puan": 10 if fr_cnt > 0 else 4,
		})

		# 5. Yatılı öğrenci boy-kilo / sağlık taraması oranı
		measured_res = int(
			conn.execute(
				"""SELECT COUNT(DISTINCT dr.id)
				   FROM dorm_residents dr
				   WHERE dr.is_active=1 AND dr.boarding_status!='Gündüzlü'
				     AND dr.student_id IS NOT NULL
				     AND (EXISTS(SELECT 1 FROM measurements m WHERE m.student_id=dr.student_id)
				          OR EXISTS(SELECT 1 FROM school_screenings sc WHERE sc.student_id=dr.student_id))"""
			).fetchone()[0]
		)
		scr_ok = (total_res > 0 and measured_res == total_res)
		items.append({
			"Madde No": "5",
			"Denetim Başlığı": "Yatılı Öğrencilerin Boy-Kilo ve Dönemsel Sağlık Taramaları",
			"Durum": "✅ Tamamlandı" if scr_ok else "🟡 Devam Ediyor",
			"Detay": f"Ölçüm/taraması bulunan yatılı öğrenci: {measured_res} / {total_res}",
			"Puan": 15 if scr_ok else (10 if measured_res > 0 else 5),
		})

		# 6. Yemekhane personel portör/sağlık muayenesi & 72 saat şahit numune
		stf_total = int(conn.execute("SELECT COUNT(*) FROM dorm_staff_health_checks").fetchone()[0])
		stf_exp = int(
			conn.execute("SELECT COUNT(*) FROM dorm_staff_health_checks WHERE valid_until_date < ?", (today_iso,)).fetchone()[0]
		)
		smp_cnt = int(conn.execute("SELECT COUNT(*) FROM dorm_food_samples WHERE sample_date >= ?", (since_7d,)).fetchone()[0])
		food_ok = (stf_total > 0 and stf_exp == 0 and smp_cnt > 0)
		items.append({
			"Madde No": "6",
			"Denetim Başlığı": "Mutfak Personeli Sağlık/Hijyen Muayeneleri & 72 Saatlik Şahit Numune",
			"Durum": "✅ Tam ve Güncel" if food_ok else "🟡 Kontrol Gerekiyor",
			"Detay": f"Personel kaydı: {stf_total} (Süresi geçmiş: {stf_exp}) · Son 7 gün şahit numune: {smp_cnt}",
			"Puan": 15 if food_ok else (8 if (stf_total > 0 or smp_cnt > 0) else 3),
		})

		# 7. Su deposu klor/analiz & ortak banyo/WC dezenfeksiyonu
		san_cnt = int(conn.execute("SELECT COUNT(*) FROM dorm_facility_sanitation_logs WHERE check_date >= ?", (since_30d,)).fetchone()[0])
		san_bad = int(
			conn.execute("SELECT COUNT(*) FROM dorm_facility_sanitation_logs WHERE check_date >= ? AND is_compliant=0", (since_30d,)).fetchone()[0]
		)
		items.append({
			"Madde No": "7",
			"Denetim Başlığı": "Su Deposu Klor/Analiz ve Ortak Banyo/WC Dezenfeksiyon Kayıtları",
			"Durum": "✅ Uygun" if (san_cnt > 0 and san_bad == 0) else "🟡 Eksik / Uygunsuz Kayıt",
			"Detay": f"Son 30 gün kontrol: {san_cnt} (Uygunsuz: {san_bad})",
			"Puan": 15 if (san_cnt > 0 and san_bad == 0) else 5,
		})

	total_score = sum(int(x["Puan"]) for x in items)
	return {
		"inspection_date": today_iso,
		"total_score": min(100, total_score),
		"items": items,
	}


# ==============================================================================
# 6. KURUMSAL RESMİ BELGE & A4 WORD (.DOCX) / HTML ÇIKTI ÜRETİCİLERİ
# ==============================================================================

def _make_docx_base(school_name: str, title: str, subtitle: str = "", landscape: bool = False):
	import docx
	from docx.enum.section import WD_ORIENT
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.shared import Cm, Pt, RGBColor

	doc = docx.Document()
	sec = doc.sections[0]
	if landscape:
		sec.orientation = WD_ORIENT.LANDSCAPE
		sec.page_width = Cm(29.7)
		sec.page_height = Cm(21.0)
	else:
		sec.orientation = WD_ORIENT.PORTRAIT
		sec.page_width = Cm(21.0)
		sec.page_height = Cm(29.7)
	sec.top_margin = Cm(1.2)
	sec.bottom_margin = Cm(1.2)
	sec.left_margin = Cm(1.4)
	sec.right_margin = Cm(1.4)

	p = doc.add_paragraph()
	p.alignment = WD_ALIGN_PARAGRAPH.CENTER
	r1 = p.add_run("T.C. MİLLÎ EĞİTİM BAKANLIĞI\n")
	r1.font.size = Pt(9.5)
	r1.font.color.rgb = RGBColor(100, 116, 139)
	r2 = p.add_run(f"{(school_name or 'OKUL PANSİYONU VE REVİR BİRİMİ').upper()}\n")
	r2.bold = True
	r2.font.size = Pt(12)
	r2.font.color.rgb = RGBColor(15, 118, 110)
	r3 = p.add_run(title)
	r3.bold = True
	r3.font.size = Pt(11)
	if subtitle:
		r4 = p.add_run(f"\n{subtitle}")
		r4.font.size = Pt(9)
		r4.font.color.rgb = RGBColor(71, 85, 105)
	return doc, Pt, RGBColor, WD_ALIGN_PARAGRAPH


def _fill_docx_simple_table(doc: Any, headers: list[str], rows: list[list[str]], Pt: Any, RGBColor: Any) -> None:
	from docx.oxml import parse_xml
	from docx.oxml.ns import nsdecls

	tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
	tbl.autofit = True
	for ci, h in enumerate(headers):
		c = tbl.cell(0, ci)
		c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F766E"/>'))
		rr = c.paragraphs[0].add_run(h)
		rr.bold = True
		rr.font.size = Pt(8.5)
		rr.font.color.rgb = RGBColor(255, 255, 255)
	for ri, r_vals in enumerate(rows, start=1):
		bg = "F8FAFC" if ri % 2 == 0 else "FFFFFF"
		for ci, val in enumerate(r_vals):
			c = tbl.cell(ri, ci)
			c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
			rr = c.paragraphs[0].add_run(str(val or "—"))
			rr.font.size = Pt(8.5)


def build_night_emergency_room_cards_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	school_name: str = "",
	nurse_name: str = "",
) -> bytes:
	"""Nöbetçi belletmen öğretmenler için Oda Bazlı Gece Acil Müdahale & Kronik Risk Kartı (.docx) üretir."""
	cards = get_chronic_night_risk_residents(connect_db_fn)
	doc, Pt, RGBColor, WD_ALIGN_PARAGRAPH = _make_docx_base(
		school_name,
		"NÖBETÇİ BELLETMEN GECE ACİL MÜDAHALE & ODA RİSK KARTLARI",
		f"Düzenlenme Tarihi: {date.today().isoformat()} · Revir Birimi ({nurse_name or 'Revir Sorumlusu'})",
		landscape=True,
	)
	if not cards:
		doc.add_paragraph("Aktif yatılı öğrenciler arasında kayıtlı yüksek riskli gece acil protokolü bulunmamaktadır.")
	else:
		headers = ["Blok / Kat / Oda - Yatak", "Öğrenci Adı Soyadı", "Klinik Risk Grubu", "Gece Acil Müdahale Adımları (Belletmen Talimatı)", "Veli Tel"]
		rows = [
			[
				f"{c['block_name']} · {c['floor_no']}\nOda {c['room_no']} ({c['bed_bunk_no']})",
				c["full_name"],
				c["risk_tags"],
				c["night_protocol"],
				c["guardian_phone"],
			]
			for c in cards
		]
		_fill_docx_simple_table(doc, headers, rows, Pt, RGBColor)
	p_f = doc.add_paragraph("\nNot: Bu belge KVKK kapsamında yalnızca nöbetçi belletmen öğretmen ve yurt idaresinin gece acil müdahale güvenliği içindir; panoya asılmaz, nöbet dosyasında muhafaza edilir.")
	p_f.runs[0].font.size = Pt(8.0)
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_rest_and_exemption_list_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	ref_date: str | None = None,
	school_name: str = "",
	nurse_name: str = "",
) -> bytes:
	"""Okul idaresi ve belletmen öğretmene sunulacak aktif Revir İstirahatli / Ders ve Etüt Muafiyetli Öğrenci Listesi (.docx)."""
	dt_s = str(ref_date or date.today().isoformat())[:10]
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		rows_db = conn.execute(
			"""SELECT re.*, dr.full_name, dr.class_name, dr.origin_school_name, dr.block_name, dr.room_no
			   FROM dorm_rest_and_exemptions re
			   JOIN dorm_residents dr ON dr.id = re.resident_id
			   WHERE re.start_date <= ? AND re.end_date >= ?
			   ORDER BY dr.block_name, dr.room_no, dr.full_name""",
			(dt_s, dt_s),
		).fetchall()
	doc, Pt, RGBColor, _ = _make_docx_base(
		school_name,
		f"{dt_s} TARİHLİ REVİR İSTİRAHATLİ & DERS / ETÜT MUAFİYETLİ YURT ÖĞRENCİ LİSTESİ",
		f"Hazırlayan: {nurse_name or 'Revir Hemşiresi'}",
	)
	if not rows_db:
		doc.add_paragraph(f"{dt_s} tarihi itibarıyla aktif revir istirahati veya etüt muafiyeti bulunan yurt öğrencisi yoktur.")
	else:
		headers = ["Öğrenci Adı Soyadı", "Okul / Sınıf", "Oda No", "Muafiyet / İstirahat Türü", "Başlangıç – Bitiş", "Neden / Açıklama"]
		rows = [
			[
				r["full_name"],
				f"{r['origin_school_name'] or ''} {r['class_name']}".strip(),
				f"{r['block_name']} / {r['room_no']}",
				r["exemption_kind"],
				f"{r['start_date']} → {r['end_date']}",
				r["reason_diagnosis"],
			]
			for r in rows_db
		]
		_fill_docx_simple_table(doc, headers, rows, Pt, RGBColor)
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_emergency_transfer_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	transfer_id: int,
	*,
	school_name: str = "",
	nurse_name: str = "",
) -> bytes:
	"""Gece 112 / Acil Hastane Sevk ve Refakat Tutanağı (.docx)."""
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		row = conn.execute(
			"""SELECT et.*, dr.full_name, dr.class_name, dr.origin_school_name, dr.block_name, dr.room_no, dr.guardian_name, dr.guardian_phone
			   FROM dorm_emergency_transfers et
			   JOIN dorm_residents dr ON dr.id = et.resident_id
			   WHERE et.id=?""",
			(int(transfer_id),),
		).fetchone()
	doc, Pt, RGBColor, _ = _make_docx_base(
		school_name,
		"PANSİYON ACİL HASTANE SEVK VE REFAKAT TUTANAĞI",
		f"Tutanak No: #{transfer_id}",
	)
	if not row:
		doc.add_paragraph("Sevk kaydı bulunamadı.")
	else:
		fields = [
			["Öğrenci Adı Soyadı", row["full_name"]],
			["Okulu / Sınıfı / Oda No", f"{row['origin_school_name'] or school_name} · {row['class_name']} · {row['block_name']} Oda {row['room_no']}"],
			["Sevk Tarihi ve Saati", f"{row['transfer_date']} — Saat: {row['transfer_time']}"],
			["Sevk Nedeni / Ön Bulgular", row["reason_complaint"]],
			["Ulaşım / Sevk Şekli & Hastane", f"{row['transport_mode']} → {row['hospital_name']}"],
			["Refakat Eden Görevli / Belletmen", row["escort_staff_name"]],
			["Veli Bilgilendirme Saati & Kişi", f"{row['guardian_notified_time'] or '—'} · {row['guardian_contact_person'] or row['guardian_name'] or ''} ({row['guardian_phone'] or ''})"],
			["Hastaneden Dönüş Saati", row["return_time"] or "Henüz dönmedi / Yatış devam ediyor"],
			["Epikriz / Tanı / Verilen İstirahat", f"{row['epicrisis_summary'] or '—'} (İstirahat: {row['rest_days_given']} gün)"],
		]
		_fill_docx_simple_table(doc, ["Bilgi Alanı", "Kayıt Detayı"], fields, Pt, RGBColor)
		doc.add_paragraph(
			f"\nRefakatçi Görevli: {row['escort_staff_name']}          Nöbetçi Revir / İdare: {nurse_name or '........................'}\nİmza:                                                                    İmza:"
		)
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_shift_handover_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	handover_id: int,
	*,
	school_name: str = "",
) -> bytes:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		row = conn.execute("SELECT * FROM dorm_shift_handovers WHERE id=?", (int(handover_id),)).fetchone()
	doc, Pt, RGBColor, _ = _make_docx_base(
		school_name,
		"PANSİYON REVİR VE BELLETMEN NÖBET DEVİR TESLİM TUTANAĞI",
		f"Kayıt #{handover_id}",
	)
	if row:
		_fill_docx_simple_table(
			doc,
			["Alan", "Açıklama"],
			[
				["Nöbet Tarihi & Vardiyası", f"{row['shift_date']} · {row['shift_type']}"],
				["Nöbeti Devreden Görevli", row["from_staff"]],
				["Nöbeti Devralan Görevli", row["to_staff"]],
				["Gece Özeti & Uygulamalar", row["summary_text"]],
				["Özel Takip Gerektiren Öğrenciler", row["critical_students_note"] or "—"],
				["Yurt İdaresi Bilgilendirme", "Evet" if row["admin_notified"] else "Hayır"],
			],
			Pt,
			RGBColor,
		)
		doc.add_paragraph(f"\nDevreden: {row['from_staff']} (İmza)                    Devralan: {row['to_staff']} (İmza)")
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_tournament_kit_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	kit_id: int,
	*,
	school_name: str = "",
	nurse_name: str = "",
) -> bytes:
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		row = conn.execute("SELECT * FROM dorm_tournament_kits WHERE id=?", (int(kit_id),)).fetchone()
	doc, Pt, RGBColor, _ = _make_docx_base(
		school_name,
		"MÜSABAKA / TURNUVA ACİL SAĞLIK ÇANTASI TESLİM-TESELLÜM TUTANAĞI",
		f"Kayıt #{kit_id}",
	)
	if row:
		_fill_docx_simple_table(
			doc,
			["Bilgi", "Detay"],
			[
				["Turnuva / Etkinlik Adı", row["event_name"]],
				["Gidilen İl / İlçe / Tesis", row["destination"]],
				["Teslim Alan Sorumlu Öğretmen / Antrenör", row["coach_or_teacher"]],
				["Verilen Çanta", row["kit_name"]],
				["Çanta İçeriği", row["contents_summary"]],
				["Teslim Tarihi → Beklenen İade", f"{row['given_date']} → {row['expected_return_date'] or '—'}"],
				["İade Tarihi & Durumu", f"{row['returned_date'] or 'Henüz iade edilmedi'} ({row['status']})"],
				["Harcanan / Eksik Malzeme Notu", row["return_missing_notes"] or "Eksik yok"],
			],
			Pt,
			RGBColor,
		)
		doc.add_paragraph(f"\nTeslim Eden Revir Sorumlusu: {nurse_name or '..................'}          Teslim Alan: {row['coach_or_teacher']}")
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_meb_dorm_inspection_docx(
	connect_db_fn: Callable[[], sqlite3.Connection],
	*,
	school_name: str = "",
	nurse_name: str = "",
	principal_name: str = "",
	extra_notes: str = "",
) -> bytes:
	"""MEB Müfettiş / İlçe Sağlık Pansiyon Denetim Kontrol Listesi ve Puanlama Raporu (.docx)."""
	ev = evaluate_meb_dorm_inspection(connect_db_fn)
	doc, Pt, RGBColor, _ = _make_docx_base(
		school_name,
		"MEB PANSİYON & REVİR SAĞLIK DENETİM HAZIRLIK VE ÖZ-DEĞERLENDİRME RAPORU",
		f"Denetim Tarihi: {ev['inspection_date']} · Genel Hazırlık Puanı: {ev['total_score']} / 100",
		landscape=True,
	)
	headers = ["No", "MEB / İlçe Sağlık Denetim Maddesi", "Durum", "Canlı Sistem Tespiti", "Puan"]
	rows = [[it["Madde No"], it["Denetim Başlığı"], it["Durum"], it["Detay"], str(it["Puan"])] for it in ev["items"]]
	_fill_docx_simple_table(doc, headers, rows, Pt, RGBColor)
	if extra_notes.strip():
		doc.add_paragraph(f"\nDenetim Ek Notları: {extra_notes.strip()}")
	doc.add_paragraph(
		f"\nRevir Sorumlusu: {nurse_name or '........................'}                    Pansiyon Md. Yrd. / Okul Müdürü: {principal_name or '........................'}"
	)
	buf = io.BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_dormitory_full_excel(connect_db_fn: Callable[[], sqlite3.Connection]) -> bytes:
	"""Tüm Okul Yurdu / Pansiyon alt tablolarını çok sekmeli kurumsal Excel (.xlsx) olarak dışa aktarır."""
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)
		sheets = {
			"Yurt_Ogrencileri": pd.read_sql_query("SELECT full_name AS 'Ad Soyad', resident_source AS 'Kaynak', origin_school_name AS 'Geldiği Okul', class_name AS 'Sınıf', boarding_status AS 'Durum', block_name AS 'Blok', floor_no AS 'Kat', room_no AS 'Oda', bed_bunk_no AS 'Yatak', entry_health_report_status AS 'Giriş Raporu', uses_devices AS 'Cihaz/Gözlük' FROM dorm_residents WHERE is_active=1 ORDER BY block_name, floor_no, room_no", conn),
			"Musahede_Yatis": pd.read_sql_query("SELECT os.admitted_date AS 'Tarih', os.admitted_time AS 'Saat', dr.full_name AS 'Öğrenci', os.bed_label AS 'Yatak', os.complaint AS 'Şikayet', os.outcome_status AS 'Sonuç', os.discharged_time AS 'Çıkış' FROM dorm_observation_stays os JOIN dorm_residents dr ON dr.id=os.resident_id ORDER BY os.id DESC", conn),
			"Nobet_Devir": pd.read_sql_query("SELECT shift_date AS 'Tarih', shift_type AS 'Vardiya', from_staff AS 'Devreden', to_staff AS 'Devralan', summary_text AS 'Özet' FROM dorm_shift_handovers ORDER BY id DESC", conn),
			"Acil_112_Sevk": pd.read_sql_query("SELECT et.transfer_date AS 'Tarih', et.transfer_time AS 'Saat', dr.full_name AS 'Öğrenci', et.hospital_name AS 'Hastane', et.escort_staff_name AS 'Refakatçi', et.guardian_notified_time AS 'Veli Haber Saati', et.return_time AS 'Dönüş', et.epicrisis_summary AS 'Epikriz' FROM dorm_emergency_transfers et JOIN dorm_residents dr ON dr.id=et.resident_id ORDER BY et.id DESC", conn),
			"Izolasyon_ve_Istirahat": pd.read_sql_query("SELECT re.start_date AS 'Başlangıç', re.end_date AS 'Bitiş', dr.full_name AS 'Öğrenci', re.exemption_kind AS 'Tür', re.reason_diagnosis AS 'Tanı/Neden' FROM dorm_rest_and_exemptions re JOIN dorm_residents dr ON dr.id=re.resident_id ORDER BY re.id DESC", conn),
			"Sahit_Numune_72s": pd.read_sql_query("SELECT sample_date AS 'Tarih', meal_slot AS 'Öğün', sampled_dishes AS 'Numuneler', fridge_temp_c AS 'Dolap °C', retention_until_iso AS '72s Saklama Bitiş', taken_by AS 'Alan' FROM dorm_food_samples ORDER BY id DESC", conn),
			"Bina_Su_Hijyen": pd.read_sql_query("SELECT check_date AS 'Tarih', check_type AS 'Kontrol Türü', location_area AS 'Alan', measured_value AS 'Değer', is_compliant AS 'Uygun', next_due_date AS 'Sonraki Tarih' FROM dorm_facility_sanitation_logs ORDER BY id DESC", conn),
		}
	ev = evaluate_meb_dorm_inspection(connect_db_fn)
	sheets["MEB_Denetim_Puani"] = pd.DataFrame(ev["items"])

	buf = io.BytesIO()
	with pd.ExcelWriter(buf, engine="openpyxl") as writer:
		for sh_nm, df_s in sheets.items():
			if df_s.empty:
				df_s = pd.DataFrame([{"Bilgi": "Henüz kayıt bulunmuyor"}])
			df_s.to_excel(writer, sheet_name=sh_nm[:31], index=False)
	return buf.getvalue()


# ==============================================================================
# 7. OKUL YURDU / PANSİYON 6 SEKMELİ KULLANICI ARAYÜZÜ (UI)
# ==============================================================================

def _resident_option_map(residents: list[dict[str, Any]], mask_fn: Callable[[Any, str], str]) -> dict[str, int]:
	out: dict[str, int] = {}
	for r in residents:
		src_tag = f" · 🏫 {r['origin_school_name']}" if r.get("resident_source") == "Başka Okul Öğrencisi" and r.get("origin_school_name") else ""
		lbl = f"{mask_fn(r['full_name'], 'name')} ({r['block_name']} Oda {r['room_no']} · {r['class_name']}{src_tag}) [#{r['id']}]"
		out[lbl] = int(r["id"])
	return out


def _render_tab_rooms_and_contacts(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	school_nm = str(app_cfg.get("school_name") or "").strip()

	with ui.expander("➕ Yurt Öğrencisi Ekle (Okul Öğrencilerinden Seç veya Başka Okuldan Gelen Öğrenci Kaydı Aç)", expanded=not bool(residents)):
		col_own, col_ext = ui.columns(2)
		with col_own:
			ui.markdown("##### 🏫 1. Okuldaki Mevcut Öğrencilerden Seçerek Odaya Yerleştir")
			enrolled_sids = {int(r["student_id"]) for r in residents if r.get("student_id")}
			with connect_db_fn() as conn:
				stu_rows = conn.execute(
					"""SELECT s.id, s.full_name, s.student_no, COALESCE(c.name, '—') AS cls
					   FROM students s
					   LEFT JOIN classes c ON c.id = s.class_id
					   WHERE s.person_type = 'Öğrenci' AND COALESCE(c.name, '') != ?
					   ORDER BY c.name ASC, s.full_name ASC""",
					(EXTERNAL_DORM_CLASS_NAME,),
				).fetchall()
			avail_stu_map = {
				f"{mask_fn(r['full_name'], 'name')} ({r['cls']}" + (f" · No:{r['student_no']}" if r["student_no"] else "") + f") [#{r['id']}]": int(r["id"])
				for r in stu_rows
				if int(r["id"]) not in enrolled_sids
			}
			picked_own = ui.multiselect(
				"Yurda Yerleştirilecek Okul Öğrencileri",
				list(avail_stu_map.keys()),
				key="dorm_pick_own_students",
			)
			oc1, oc2, oc3, oc4 = ui.columns(4)
			ow_blk = oc1.text_input("Blok", value="A Blok", key="dorm_ow_blk")
			ow_flr = oc2.text_input("Kat", value="1. Kat", key="dorm_ow_flr")
			ow_rm = oc3.text_input("Oda No", value="101", key="dorm_ow_rm")
			ow_bed = oc4.text_input("Yatak/Ranza Öneki", value="Yatak", key="dorm_ow_bed")
			oc5, oc6 = ui.columns(2)
			ow_stat = oc5.selectbox("Pansiyon Durumu", BOARDING_STATUS_OPTIONS, key="dorm_ow_stat")
			ow_rep = oc6.selectbox("Giriş Sağlık Raporu", ENTRY_REPORT_STATUS_OPTIONS, key="dorm_ow_rep")
			if ui.button("📥 Seçili Okul Öğrencilerini Yurda Kaydet", type="primary", use_container_width=True, disabled=not picked_own, key="dorm_btn_enroll_own"):
				cnt = enroll_existing_school_students_to_dorm(
					connect_db_fn,
					[avail_stu_map[k] for k in picked_own],
					block_name=ow_blk,
					floor_no=ow_flr,
					room_no=ow_rm,
					bed_prefix=ow_bed,
					boarding_status=ow_stat,
					entry_health_report_status=ow_rep,
					school_name_default=school_nm,
				)
				ui.session_state.pop("dorm_pick_own_students", None)
				ui.success(f"✅ {cnt} kurum öğrencisi yurt oda haritasına yerleştirildi.")
				ui.rerun()

		with col_ext:
			ui.markdown("##### 🚌 2. Başka Okuldan Gelen Yatılı Öğrenci Kaydı Aç")
			with ui.form("dorm_form_enroll_external", clear_on_submit=True):
				ex1, ex2 = ui.columns([1.4, 1.3])
				ex_name = ex1.text_input("Öğrenci Adı Soyadı*", placeholder="Örn: Kerem Yıldız")
				ex_orig = ex2.text_input("Geldiği Okul Adı*", placeholder="Örn: Cumhuriyet Fen Lisesi")
				ex3, ex4, ex5, ex6 = ui.columns([0.8, 1.0, 1.0, 1.1])
				ex_sex = ex3.selectbox("Cinsiyet", ["E", "K"])
				ex_cls = ex4.text_input("Sınıfı", value="9-A")
				ex_sno = ex5.text_input("Okul No", placeholder="412")
				ex_bdt = ex6.text_input("Doğum Tarihi", placeholder="YYYY-AA-GG")
				ex7, ex8, ex9 = ui.columns(3)
				ex_gnm = ex7.text_input("Veli Adı Soyadı")
				ex_gph = ex8.text_input("Veli Telefonu", placeholder="05XX...")
				ex_hloc = ex9.text_input("Memleket (İl / İlçe)", placeholder="Örn: Konya / Ereğli")
				rm1, rm2, rm3, rm4 = ui.columns(4)
				ex_blk = rm1.text_input("Blok", value="A Blok")
				ex_flr = rm2.text_input("Kat", value="2. Kat")
				ex_rm = rm3.text_input("Oda No", value="201")
				ex_bed = rm4.text_input("Yatak / Ranza", value="Üst-1")
				md1, md2, md3 = ui.columns(3)
				ex_rep = md1.selectbox("Yurtta Kalabilir Raporu", ENTRY_REPORT_STATUS_OPTIONS)
				ex_med = md2.text_input("Alerji / Kronik Hastalık", placeholder="Örn: Astım, Fıstık alerjisi")
				ex_dev = md3.text_input("Gözlük / Cihaz / Ortez", placeholder="Örn: Gözlük, İşitme Cihazı")
				if ui.form_submit_button("➕ Başka Okul Öğrencisini Yurda & Revire Kaydet", type="primary", use_container_width=True):
					try:
						enroll_external_school_student_to_dorm(
							connect_db_fn,
							full_name=ex_name,
							origin_school_name=ex_orig,
							sex=ex_sex,
							class_name=ex_cls,
							student_no=ex_sno,
							birth_date=ex_bdt,
							guardian_name=ex_gnm,
							guardian_phone=ex_gph,
							home_location=ex_hloc,
							block_name=ex_blk,
							floor_no=ex_flr,
							room_no=ex_rm,
							bed_bunk_no=ex_bed,
							entry_health_report_status=ex_rep,
							uses_devices=ex_dev,
							chronic_conditions=ex_med,
						)
						if callable(ctx_api.get("clear_all_data_caches")):
							ctx_api["clear_all_data_caches"]()
						ui.success(f"✅ {ex_name} ({ex_orig}) dış okul yatılı öğrencisi olarak kaydedildi.")
						ui.rerun()
					except Exception as exc:
						ui.error(str(exc))

	if not residents:
		ui.info("Henüz kayıtlı yurt öğrencisi bulunmuyor. Yukarıdan kurum öğrencisi seçebilir veya başka okuldan gelen öğrenci kaydı açabilirsiniz.")
		return

	# Filtre ve Özet Metrikler
	blocks = sorted({r["block_name"] for r in residents})
	flt1, flt2, flt3 = ui.columns(3)
	sel_blk = flt1.selectbox("Blok Filtresi", ["Tüm Bloklar", *blocks], key="dorm_flt_blk")
	sel_src = flt2.selectbox("Öğrenci Kaynağı", ["Tümü", *RESIDENT_SOURCE_OPTIONS], key="dorm_flt_src")
	sel_rep_f = flt3.selectbox("Sağlık Raporu Filtresi", ["Tümü", *ENTRY_REPORT_STATUS_OPTIONS], key="dorm_flt_rep")

	filtered_res = [
		r for r in residents
		if (sel_blk == "Tüm Bloklar" or r["block_name"] == sel_blk)
		and (sel_src == "Tümü" or r["resident_source"] == sel_src)
		and (sel_rep_f == "Tümü" or r["entry_health_report_status"] == sel_rep_f)
	]

	# Görsel Blok / Kat / Oda Kartları
	ui.markdown(f"#### 🛏️ Oda - Yatak Yerleşim Haritası ({len(filtered_res)} Öğrenci)")
	rooms_grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
	for r in filtered_res:
		rooms_grouped.setdefault((r["block_name"], r["floor_no"], r["room_no"]), []).append(r)

	room_keys = list(rooms_grouped.keys())
	for idx_g in range(0, len(room_keys), 3):
		cols_r = ui.columns(3)
		for sub_i, rk in enumerate(room_keys[idx_g : idx_g + 3]):
			occ = rooms_grouped[rk]
			with cols_r[sub_i]:
				lines_h = []
				for o in occ:
					src_b = f" <span style='font-size:11px;color:#0369a1;'>[🏫 {html.escape(o['origin_school_name'])}]</span>" if o["resident_source"] == "Başka Okul Öğrencisi" and o["origin_school_name"] else ""
					r_badges = []
					if int(o.get("is_diabetic") or 0) == 1:
						r_badges.append("💙 Diyabet")
					if o.get("combined_clinical_risk"):
						r_badges.append(f"⚠️ {mask_fn(o['combined_clinical_risk'], 'medical')[:28]}")
					if o["entry_health_report_status"] != "Tam (Yurtta Kalabilir)":
						r_badges.append("📄 Rapor Eksik")
					b_str = (" · " + " ".join(r_badges)) if r_badges else ""
					lines_h.append(
						f"<li style='margin-bottom:4px;'><b>{html.escape(o['bed_bunk_no'])}:</b> "
						f"{html.escape(mask_fn(o['full_name'], 'name'))} <small>({html.escape(o['class_name'])})</small>{src_b}"
						f"<span style='font-size:11px;color:#b45309;'>{html.escape(b_str)}</span></li>"
					)
				ui.markdown(
					f"<div style='border:1px solid rgba(15,118,110,0.3);border-radius:10px;padding:10px 12px;margin-bottom:8px;background:rgba(15,118,110,0.04);'>"
					f"<div style='font-weight:800;color:#0f766e;margin-bottom:6px;'>🚪 {html.escape(rk[0])} · {html.escape(rk[1])} · Oda {html.escape(rk[2])} ({len(occ)} Kişi)</div>"
					f"<ul style='margin:0;padding-left:18px;font-size:13px;'>{''.join(lines_h)}</ul></div>",
					unsafe_allow_html=True,
				)

	with ui.expander("✏️ Öğrenci Oda / Yatak Değiştir, Giriş Raporu Güncelle veya Dosyasını Aç"):
		res_map = _resident_option_map(residents, mask_fn)
		pick_ed = ui.selectbox("İşlem Yapılacak Yurt Öğrencisi", list(res_map.keys()), key="dorm_ed_res_pick")
		sel_rid = res_map[pick_ed]
		cur_r = next(x for x in residents if int(x["id"]) == sel_rid)
		ec1, ec2, ec3, ec4 = ui.columns(4)
		n_blk = ec1.text_input("Blok", value=cur_r["block_name"], key=f"dorm_ed_b_{sel_rid}")
		n_flr = ec2.text_input("Kat", value=cur_r["floor_no"], key=f"dorm_ed_f_{sel_rid}")
		n_rm = ec3.text_input("Oda No", value=cur_r["room_no"], key=f"dorm_ed_r_{sel_rid}")
		n_bed = ec4.text_input("Yatak / Ranza", value=cur_r["bed_bunk_no"], key=f"dorm_ed_bd_{sel_rid}")
		ec5, ec6, ec7 = ui.columns(3)
		n_st = ec5.selectbox("Pansiyon Durumu", BOARDING_STATUS_OPTIONS, index=BOARDING_STATUS_OPTIONS.index(cur_r["boarding_status"]) if cur_r["boarding_status"] in BOARDING_STATUS_OPTIONS else 0, key=f"dorm_ed_st_{sel_rid}")
		n_rp = ec6.selectbox("Yurt Giriş Sağlık Raporu", ENTRY_REPORT_STATUS_OPTIONS, index=ENTRY_REPORT_STATUS_OPTIONS.index(cur_r["entry_health_report_status"]) if cur_r["entry_health_report_status"] in ENTRY_REPORT_STATUS_OPTIONS else 0, key=f"dorm_ed_rp_{sel_rid}")
		n_dv = ec7.text_input("Gözlük / Protez / Cihaz", value=cur_r.get("effective_devices") or "", key=f"dorm_ed_dv_{sel_rid}")
		eb1, eb2, eb3 = ui.columns(3)
		if eb1.button("💾 Oda & Rapor Bilgisini Güncelle", type="primary", use_container_width=True, key="dorm_btn_upd_res"):
			update_dorm_resident_placement(
				connect_db_fn, sel_rid, block_name=n_blk, floor_no=n_flr, room_no=n_rm, bed_bunk_no=n_bed,
				boarding_status=n_st, entry_health_report_status=n_rp, uses_devices=n_dv, is_active=True,
			)
			ui.success("✅ Oda, yatak ve sağlık raporu künyesi güncellendi.")
			ui.rerun()
		if cur_r.get("student_id") and callable(ctx_api.get("open_person_dossier_dialog")):
			if eb2.button("📂 Tam Kişi Sağlık Dosyasını Aç", use_container_width=True, key="dorm_btn_open_dos"):
				ctx_api["open_person_dossier_dialog"](int(cur_r["student_id"]))
		if eb3.button("⏸️ Yurttan Ayrıldı (Pasife Al)", type="secondary", use_container_width=True, key="dorm_btn_deact_res"):
			update_dorm_resident_placement(
				connect_db_fn, sel_rid, block_name=n_blk, floor_no=n_flr, room_no=n_rm, bed_bunk_no=n_bed,
				boarding_status=n_st, entry_health_report_status=n_rp, is_active=False,
			)
			ui.success("Öğrenci aktif yurt listesinden çıkarıldı (geçmiş kayıtlar korundu).")
			ui.rerun()

	# Bulaşıcı Hastalık Kümelenme & Oda Arkadaşı Yakın Temaslı Tarayıcısı
	ui.markdown("---")
	ui.markdown("#### 🦠 Bulaşıcı Hastalık Kümelenme & Oda Arkadaşı / Komşu Oda Yakın Temaslı Listesi")
	auto_clusters = detect_recent_contagious_room_clusters(connect_db_fn, within_days=10)
	if auto_clusters:
		for cl in auto_clusters:
			ui.warning(
				f"⚠️ **Otomatik Kümelenme Uyarısı ({cl['event_date']}):** **{mask_fn(cl['full_name'], 'name')}** ({cl['room_label']}) — "
				f"Bulgu: _{cl['finding']}_ · Aynı odada **{cl['roommate_count']}**, aynı katta/komşu odalarda **{cl['neighbor_count']}** temaslı öğrenci var!"
			)
	res_map_ct = _resident_option_map(residents, mask_fn)
	ct_pick = ui.selectbox(
		"Bulaşıcı Hastalık (Uyuz, Bit, Suçiçeği, Grip, Gastroenterit vb.) Şüphesi Olan Vaka Öğrenciyi Seçin",
		list(res_map_ct.keys()),
		key="dorm_contact_trace_pick",
	)
	ct_res = get_room_and_neighbor_contacts(connect_db_fn, res_map_ct[ct_pick])
	all_contacts = ct_res["roommates"] + ct_res["floor_neighbors"]
	if all_contacts:
		df_ct = pd.DataFrame([
			{
				"Temas Derecesi": c["contact_degree"],
				"Ad Soyad": mask_fn(c["full_name"], "name"),
				"Blok / Kat / Oda": f"{c['block_name']} · {c['floor_no']} · Oda {c['room_no']} ({c['bed_bunk_no']})",
				"Okul / Sınıf": f"{c['origin_school_name'] + ' · ' if c['resident_source']=='Başka Okul Öğrencisi' else ''}{c['class_name']}",
				"Veli Telefon": mask_fn(c.get("guardian_phone") or "—", "phone"),
			}
			for c in all_contacts
		])
		ui.dataframe(df_ct, hide_index=True, use_container_width=True)
	else:
		ui.caption("Seçilen öğrencinin aynı odasında veya katında başka yatılı öğrenci bulunmuyor.")


def _render_tab_observation_and_night_shift(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	nurse_nm = str(app_cfg.get("nurse_name") or "Revir Hemşiresi").strip()
	school_nm = str(app_cfg.get("school_name") or "").strip()
	res_map = _resident_option_map(residents, mask_fn)

	bed_cap = max(1, min(20, int(_safe_float(app_cfg.get("dorm_observation_bed_capacity"), 4) or 4)))
	with connect_db_fn() as conn:
		active_stays = [
			dict(r)
			for r in conn.execute(
				"""SELECT os.*, dr.full_name, dr.block_name, dr.room_no, dr.class_name
				   FROM dorm_observation_stays os
				   JOIN dorm_residents dr ON dr.id = os.resident_id
				   WHERE os.outcome_status = 'Müşahedede (Aktif Yatış)'
				   ORDER BY os.admitted_date DESC, os.admitted_time DESC"""
			).fetchall()
		]
		recent_stays = [
			dict(r)
			for r in conn.execute(
				"""SELECT os.*, dr.full_name, dr.room_no
				   FROM dorm_observation_stays os
				   JOIN dorm_residents dr ON dr.id = os.resident_id
				   ORDER BY os.id DESC LIMIT 25"""
			).fetchall()
		]

	ui.markdown(f"#### 🛏️ Revir Müşahede Yatakları (Dolu: {len(active_stays)} / Kapasite: {bed_cap})")
	bed_labels = [f"Yatak {i}" for i in range(1, bed_cap + 1)]
	occ_by_bed = {s["bed_label"]: s for s in active_stays}
	b_cols = ui.columns(min(4, bed_cap))
	for bi, b_lbl in enumerate(bed_labels):
		occ_s = occ_by_bed.get(b_lbl)
		with b_cols[bi % len(b_cols)]:
			if occ_s:
				ui.error(f"🔴 **{b_lbl} — DOLU**\n\n**{mask_fn(occ_s['full_name'], 'name')}** (Oda {occ_s['room_no']})\n🕒 {occ_s['admitted_time']} · {occ_s['complaint']}")
			else:
				ui.success(f"🟢 **{b_lbl} — BOŞ**\n\nMüşahedeye hazır")

	if residents:
		with ui.expander("➕ Revir Müşahede Yatağına Öğrenci Yatır", expanded=False):
			with ui.form("dorm_form_admit_obs", clear_on_submit=True):
				oa1, oa2, oa3, oa4 = ui.columns([1.8, 1.0, 1.0, 0.9])
				o_res = oa1.selectbox("Öğrenci", list(res_map.keys()))
				free_beds = [b for b in bed_labels if b not in occ_by_bed] or bed_labels
				o_bed = oa2.selectbox("Müşahede Yatağı", free_beds)
				o_dt = oa3.text_input("Tarih", value=date.today().isoformat())
				o_tm = oa4.text_input("Yatış Saati", value=datetime.now().strftime("%H:%M"))
				oa5, oa6 = ui.columns(2)
				o_cmp = oa5.text_input("Başvuru Şikayeti / Ön Bulgu*", placeholder="Örn: 38.7 °C ateş, titreme, boğaz ağrısı")
				o_int = oa6.text_input("İlk Müdahale", placeholder="Örn: Soğuk uygulama, oral sıvı, ateş takibi")
				if ui.form_submit_button("🛏️ Müşahede Yatışını Başlat", type="primary", use_container_width=True):
					if not o_cmp.strip():
						ui.warning("Şikayet alanını doldurun.")
					else:
						admit_to_observation_bed(
							connect_db_fn, resident_id=res_map[o_res], bed_label=o_bed,
							admitted_date=o_dt, admitted_time=o_tm, complaint=o_cmp,
							initial_intervention=o_int, admitted_by=nurse_nm,
						)
						ui.success("Öğrenci müşahede yatağına alındı.")
						ui.rerun()

	if active_stays:
		with ui.container(border=True):
			ui.markdown("##### 🌡️ Yatan Öğrenciye Saatlik Gece Vital Ekle veya Çıkış / Taburcu Yap")
			stay_opts = {f"{s['bed_label']} — {mask_fn(s['full_name'], 'name')} ({s['admitted_time']} · {s['complaint']}) [#{s['id']}]": int(s["id"]) for s in active_stays}
			sel_stay_lbl = ui.selectbox("Aktif Müşahede Kaydı", list(stay_opts.keys()), key="dorm_act_stay_pick")
			sel_stay_id = stay_opts[sel_stay_lbl]

			vc1, vc2, vc3, vc4, vc5, vc6 = ui.columns(6)
			v_tm = vc1.text_input("Saat", value=datetime.now().strftime("%H:%M"), key="dorm_ov_tm")
			v_tmp = vc2.number_input("Ateş (°C)", min_value=34.0, max_value=42.5, step=0.1, value=37.0, key="dorm_ov_tmp")
			v_sys = vc3.number_input("Sistolik", min_value=0, max_value=240, value=110, key="dorm_ov_sys")
			v_dia = vc4.number_input("Diyastolik", min_value=0, max_value=150, value=70, key="dorm_ov_dia")
			v_pls = vc5.number_input("Nabız/dk", min_value=0, max_value=220, value=82, key="dorm_ov_pls")
			v_nt = vc6.text_input("Gözlem Notu", placeholder="Ateş düştü, uyuyor", key="dorm_ov_nt")
			vb1, vb2, vb3 = ui.columns([1.1, 1.4, 1.1], vertical_alignment="center")
			if vb1.button("➕ Saatlik Vital Kaydet", type="primary", use_container_width=True, key="dorm_btn_add_ov"):
				add_observation_vital_log(
					connect_db_fn, stay_id=sel_stay_id, recorded_at=f"{date.today().isoformat()} {v_tm}",
					temperature_c=float(v_tmp), systolic_mmhg=float(v_sys) if v_sys > 0 else None,
					diastolic_mmhg=float(v_dia) if v_dia > 0 else None, pulse_bpm=float(v_pls) if v_pls > 0 else None,
					care_note=v_nt, recorded_by=nurse_nm,
				)
				ui.success("Saatlik vital izlem kaydedildi.")
				ui.rerun()
			out_choice = vb2.selectbox("Çıkış Kararı", OBSERVATION_OUTCOMES[1:], key="dorm_obs_out_sel")
			if vb3.button("✅ Müşahededen Çıkış Yap", type="secondary", use_container_width=True, key="dorm_btn_disch_obs"):
				discharge_observation_stay(connect_db_fn, sel_stay_id, outcome_status=out_choice, discharge_note=v_nt)
				ui.success("Müşahede çıkışı tamamlandı.")
				ui.rerun()

			with connect_db_fn() as conn:
				v_hist = conn.execute("SELECT recorded_at AS 'Saat', temperature_c AS 'Ateş (°C)', systolic_mmhg AS 'Sistolik', diastolic_mmhg AS 'Diyastolik', pulse_bpm AS 'Nabız', care_note AS 'Not' FROM dorm_observation_vitals WHERE stay_id=? ORDER BY id DESC", (sel_stay_id,)).fetchall()
			if v_hist:
				ui.dataframe(pd.DataFrame([dict(x) for x in v_hist]), hide_index=True, use_container_width=True)

	# 2. Nöbet Devir Teslim Defteri
	ui.markdown("---")
	ui.markdown("#### 📓 Nöbetçi Belletmen Öğretmen & Gece Hemşiresi Devir Teslim Defteri")
	if ui.button("✨ Bugünün Gece Kayıtlarından Otomatik Nöbet Özeti Oluştur", key="dorm_btn_auto_handover"):
		ui.session_state["dorm_ho_summary_draft"] = build_auto_night_handover_summary(connect_db_fn, date.today().isoformat())

	with ui.form("dorm_form_handover", clear_on_submit=True):
		hc1, hc2, hc3, hc4 = ui.columns(4)
		ho_dt = hc1.text_input("Nöbet Tarihi", value=date.today().isoformat())
		ho_sh = hc2.selectbox("Vardiya", ["17:00 - 08:00 Gece Nöbeti", "Hafta Sonu 24 Saat Nöbeti", "08:00 - 17:00 Gündüz Devir"])
		ho_fr = hc3.text_input("Nöbeti Devreden*", value=nurse_nm)
		ho_to = hc4.text_input("Nöbeti Devralan (Sabah Görevlisi / İdare)*", placeholder="Örn: Belletmen Ahmet Yılmaz")
		ho_sum = ui.text_area(
			"Gece Nöbet Özeti (Başvurular, Müşahede, Verilen İlaçlar, Sevkler)*",
			value=str(ui.session_state.get("dorm_ho_summary_draft") or ""),
			height=85,
		)
		ho_crit = ui.text_input("Sabah İlk Kontrol Edilecek Öğrenciler / Oda Notu", placeholder="Örn: Oda 204 Ali Kaya sabah 08:00 ateş kontrolü")
		if ui.form_submit_button("💾 Nöbet Devir Teslimini Kaydet", type="primary", use_container_width=True):
			if not (ho_fr.strip() and ho_to.strip() and ho_sum.strip()):
				ui.warning("Devreden, devralan ve özet alanlarını doldurun.")
			else:
				save_shift_handover(
					connect_db_fn, shift_date=ho_dt, shift_type=ho_sh, from_staff=ho_fr,
					to_staff=ho_to, summary_text=ho_sum, critical_students_note=ho_crit,
				)
				ui.session_state.pop("dorm_ho_summary_draft", None)
				ui.success("✅ Nöbet devir teslim kaydı oluşturuldu.")
				ui.rerun()

	with connect_db_fn() as conn:
		ho_rows = [dict(r) for r in conn.execute("SELECT * FROM dorm_shift_handovers ORDER BY id DESC LIMIT 15").fetchall()]
	if ho_rows:
		ui.dataframe(
			pd.DataFrame([
				{"#": r["id"], "Tarih": r["shift_date"], "Vardiya": r["shift_type"], "Devreden": r["from_staff"], "Devralan": r["to_staff"], "Gece Özeti": r["summary_text"]}
				for r in ho_rows
			]),
			hide_index=True,
			use_container_width=True,
		)
		ui.download_button(
			f"📝 Son Nöbet Devir Tutanağını İndir (#{ho_rows[0]['id']} .docx)",
			data=lambda: build_shift_handover_docx(connect_db_fn, int(ho_rows[0]["id"]), school_name=school_nm),
			file_name=f"Pansiyon_Nobet_Devir_{ho_rows[0]['shift_date']}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			key="dorm_dl_ho_docx",
		)

	# 3. Gece 112 / Acil Hastane Sevk ve Refakat Tutanağı
	ui.markdown("---")
	ui.markdown("#### 🚑 Gece 112 / Acil Hastane Sevk ve Refakat Tutanağı")
	if residents:
		with ui.expander("➕ Yeni 112 / Acil Hastane Sevk ve Refakat Kaydı Ekle"):
			with ui.form("dorm_form_112_transfer", clear_on_submit=True):
				tr1, tr2, tr3, tr4 = ui.columns([1.8, 0.9, 0.8, 1.1])
				t_res = tr1.selectbox("Sevk Edilen Öğrenci", list(res_map.keys()))
				t_dt = tr2.text_input("Tarih", value=date.today().isoformat())
				t_tm = tr3.text_input("Sevk Saati", value=datetime.now().strftime("%H:%M"))
				t_mod = tr4.selectbox("Ulaşım", ["112 Ambulans", "Kurum / Nöbetçi Aracı", "Veli Aracı"])
				tr5, tr6, tr7 = ui.columns(3)
				t_hosp = tr5.text_input("Sevk Edilen Hastane*", placeholder="Örn: Eğitim ve Araştırma Hastanesi Çocuk Acil")
				t_esc = tr6.text_input("Refakat Eden Görevli / Belletmen*", placeholder="Örn: Nöbetçi Öğrt. Mehmet Demir")
				t_rsn = tr7.text_input("Sevk Nedeni / Ön Bulgular*", placeholder="Örn: Sağ alt kadran şiddetli karın ağrısı, kusma")
				tr8, tr9, tr10, tr11 = ui.columns([1.0, 1.2, 1.0, 0.9])
				t_gtm = tr8.text_input("Veliye Haber Saati", placeholder="23:20")
				t_gwho = tr9.text_input("Görüşülen Veli", placeholder="Babası Hasan Bey")
				t_ret = tr10.text_input("Hastaneden Dönüş Saati", placeholder="02:40")
				t_rdays = tr11.number_input("Verilen Rapor (Gün)", min_value=0, max_value=30, value=0)
				t_epi = ui.text_input("Epikriz / Tanı / Reçete Özeti", placeholder="Örn: Akut tonsillit, 3 gün istirahat raporu verildi")
				if ui.form_submit_button("🚑 Acil Sevk & Refakat Tutanağını Kaydet", type="primary", use_container_width=True):
					if not (t_hosp.strip() and t_esc.strip() and t_rsn.strip()):
						ui.warning("Hastane, refakatçi ve sevk nedeni zorunludur.")
					else:
						record_emergency_transfer(
							connect_db_fn, resident_id=res_map[t_res], transfer_date=t_dt, transfer_time=t_tm,
							reason_complaint=t_rsn, transport_mode=t_mod, hospital_name=t_hosp,
							escort_staff_name=t_esc, guardian_notified_time=t_gtm, guardian_contact_person=t_gwho,
							return_time=t_ret, epicrisis_summary=t_epi, rest_days_given=int(t_rdays),
						)
						ui.success("✅ 112 Acil Sevk kaydı oluşturuldu (varsa istirahat raporu yoklama listesine otomatik eklendi).")
						ui.rerun()

	with connect_db_fn() as conn:
		tr_rows = [
			dict(r)
			for r in conn.execute(
				"""SELECT et.*, dr.full_name, dr.room_no
				   FROM dorm_emergency_transfers et
				   JOIN dorm_residents dr ON dr.id = et.resident_id
				   ORDER BY et.id DESC LIMIT 15"""
			).fetchall()
		]
	if tr_rows:
		ui.dataframe(
			pd.DataFrame([
				{
					"#": r["id"],
					"Tarih/Saat": f"{r['transfer_date']} {r['transfer_time']}",
					"Öğrenci": f"{mask_fn(r['full_name'], 'name')} (Oda {r['room_no']})",
					"Hastane & Ulaşım": f"{r['hospital_name']} ({r['transport_mode']})",
					"Refakatçi": r["escort_staff_name"],
					"Veli Haber": r["guardian_notified_time"] or "—",
					"Dönüş": r["return_time"] or "—",
					"Epikriz": r["epicrisis_summary"] or "—",
				}
				for r in tr_rows
			]),
			hide_index=True,
			use_container_width=True,
		)
		ui.download_button(
			f"📝 Son 112 Sevk & Refakat Tutanağını İndir (#{tr_rows[0]['id']} .docx)",
			data=lambda: build_emergency_transfer_docx(connect_db_fn, int(tr_rows[0]["id"]), school_name=school_nm, nurse_name=nurse_nm),
			file_name=f"Acil_112_Sevk_Tutanagi_{tr_rows[0]['transfer_date']}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			key="dorm_dl_112_docx",
		)


def _render_tab_meds_cabinet_and_night_cards(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	nurse_nm = str(app_cfg.get("nurse_name") or "Revir Hemşiresi").strip()
	school_nm = str(app_cfg.get("school_name") or "").strip()
	res_map = _resident_option_map(residents, mask_fn)

	ui.markdown("#### ⏰ 7/24 Akşam, Sabah Aç Karnına (07:00), Yatmadan Önce (21:00) & Hafta Sonu İlaç Çizelgesi")
	ui.caption("Veliden teslim alınan tüm öğrenci ilaçları merkezi kilitli revir dolabında muhafaza edilir ve Cumartesi-Pazar dahil imza karşılığı içirilir.")

	if residents:
		with ui.expander("➕ Merkezi Kilitli İlaç Dolabına Yeni Yurt Öğrenci İlacı & Gece/Hafta Sonu Dozu Tanımla"):
			with ui.form("dorm_form_add_med", clear_on_submit=True):
				mc1, mc2, mc3 = ui.columns([1.8, 1.4, 1.0])
				m_res = mc1.selectbox("Öğrenci", list(res_map.keys()))
				m_nm = mc2.text_input("İlaç Adı*", placeholder="Örn: Levotiroksin 50 mcg / Montelukast 5 mg")
				m_ds = mc3.text_input("Doz", value="1 Tablet")
				mc4, mc5, mc6, mc7 = ui.columns([1.6, 0.9, 0.9, 1.0])
				m_sl = mc4.selectbox("Uygulama Saati (7/24 Yurt Düzeni)", DORM_MED_SLOTS_24H)
				m_cab = mc5.text_input("Kilitli Dolap No", value="Dolap-1")
				m_shf = mc6.text_input("Raf / Kutu No", value="Raf-A")
				m_exp = mc7.text_input("İlaç SKT", placeholder="YYYY-AA-GG")
				mc8, mc9 = ui.columns([1.6, 1.2], vertical_alignment="center")
				m_gdel = mc8.text_input("Teslim Eden Veli / Belletmen", placeholder="Velisi imzalı teslim etti")
				m_wknd = mc9.checkbox("Cumartesi & Pazar Dahil 7 Gün Uygula", value=True)
				if ui.form_submit_button("🔒 Kilitli Dolaba Kaydet & Doz Planla", type="primary", use_container_width=True):
					if not m_nm.strip():
						ui.warning("İlaç adı girin.")
					else:
						with connect_db_fn() as conn:
							ensure_dormitory_tables(conn)
							conn.execute(
								"""INSERT INTO dorm_med_schedules(
									resident_id, medication_name, dose_text, slot_label, include_weekends,
									cabinet_no, shelf_no, delivered_by_guardian, expiry_date, is_active, created_at
								) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)""",
								(
									res_map[m_res], m_nm.strip(), m_ds.strip(), m_sl,
									1 if m_wknd else 0, m_cab.strip(), m_shf.strip(),
									m_gdel.strip(), m_exp.strip()[:10], _now_iso(),
								),
							)
						ui.success("✅ İlaç merkezi kilitli dolaba ve 7/24 doz çizelgesine kaydedildi.")
						ui.rerun()

	sel_med_dt = ui.text_input("İlaç Uygulama Tarihi (Hafta Sonu Dahil)", value=date.today().isoformat(), key="dorm_med_exec_dt")
	is_wknd_dt = False
	try:
		is_wknd_dt = date.fromisoformat(sel_med_dt[:10]).weekday() >= 5
	except Exception:
		pass

	with connect_db_fn() as conn:
		sched_rows = [
			dict(r)
			for r in conn.execute(
				"""SELECT ms.*, dr.full_name, dr.block_name, dr.room_no
				   FROM dorm_med_schedules ms
				   JOIN dorm_residents dr ON dr.id = ms.resident_id
				   WHERE ms.is_active=1
				   ORDER BY ms.slot_label ASC, dr.full_name ASC"""
			).fetchall()
		]
		given_map = {
			(int(r["schedule_id"]), str(r["slot_label"])): dict(r)
			for r in conn.execute("SELECT * FROM dorm_med_dose_logs WHERE dose_date=?", (sel_med_dt[:10],)).fetchall()
		}

	due_scheds = [s for s in sched_rows if (not is_wknd_dt or int(s.get("include_weekends") or 0) == 1)]
	if due_scheds:
		ui.dataframe(
			pd.DataFrame([
				{
					"Öğün / Saat": s["slot_label"],
					"Öğrenci (Oda)": f"{mask_fn(s['full_name'], 'name')} (Oda {s['room_no']})",
					"İlaç & Doz": f"{s['medication_name']} — {s['dose_text']}",
					"Kilitli Dolap": f"{s['cabinet_no']} / {s['shelf_no']}",
					"Hafta Sonu": "Evet" if s["include_weekends"] else "Hafta İçi",
					"Bugünkü Durum": (
						f"✅ İçirildi ({given_map[(int(s['id']), s['slot_label'])]['given_time']})"
						if (int(s["id"]), s["slot_label"]) in given_map
						else "⏳ Bekliyor"
					),
				}
				for s in due_scheds
			]),
			hide_index=True,
			use_container_width=True,
		)
		pending_s = [s for s in due_scheds if (int(s["id"]), s["slot_label"]) not in given_map]
		if pending_s:
			ps_map = {f"{s['slot_label']} — {mask_fn(s['full_name'], 'name')} · {s['medication_name']}": s for s in pending_s}
			pc1, pc2 = ui.columns([2.4, 1.1], vertical_alignment="center")
			pick_ps = pc1.selectbox("İmza Karşılığı İçirilecek Doz", list(ps_map.keys()), key="dorm_pick_pend_dose")
			if pc2.button("✍️ İmza Karşılığı İçirildi Kaydet", type="primary", use_container_width=True, key="dorm_btn_give_dose"):
				chosen_s = ps_map[pick_ps]
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT OR REPLACE INTO dorm_med_dose_logs(
							schedule_id, resident_id, dose_date, slot_label, given_time,
							status, administered_by, student_signed
						) VALUES (?, ?, ?, ?, ?, 'İmza Karşılığı İçirildi', ?, 1)""",
						(
							int(chosen_s["id"]), int(chosen_s["resident_id"]), sel_med_dt[:10],
							chosen_s["slot_label"], datetime.now().strftime("%H:%M"), nurse_nm,
						),
					)
				ui.success("✅ Doz imza karşılığı uygulandı olarak kaydedildi.")
				ui.rerun()

	# Oda İlaç Denetimi
	with ui.expander("🔍 Oda İlaç Denetimi Kaydı (Odalarda İzinsiz İlaç / Takviye Kontrolü)"):
		with ui.form("dorm_form_room_med_insp", clear_on_submit=True):
			ri1, ri2, ri3, ri4 = ui.columns([1.0, 1.0, 1.0, 1.2])
			i_dt = ri1.text_input("Denetim Tarihi", value=date.today().isoformat())
			i_blk = ri2.text_input("Blok", value="A Blok")
			i_rm = ri3.text_input("Denetlenen Oda No", value="101")
			i_fnd = ri4.checkbox("İzinsiz İlaç Bulundu mu?", value=False)
			ri5, ri6 = ui.columns(2)
			i_items = ri5.text_input("El Konulan İlaç / Takviye (Varsa)", placeholder="Örn: 1 kutu ağrı kesici revir dolabına alındı")
			i_act = ri6.text_input("Yapılan İşlem", value="Oda temiz / Öğrenci bilgilendirildi")
			if ui.form_submit_button("💾 Oda İlaç Denetimini Kaydet", use_container_width=True):
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT INTO dorm_room_med_inspections(
							inspection_date, block_name, room_no, unauthorized_med_found,
							confiscated_items, action_taken, inspected_by, created_at
						) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
						(i_dt[:10], i_blk.strip(), i_rm.strip(), 1 if i_fnd else 0, i_items.strip(), i_act.strip(), nurse_nm, _now_iso()),
					)
				ui.success("✅ Oda ilaç denetimi kaydedildi.")
				ui.rerun()

	# 2. Kronik Hastalıklı Öğrenciler Gece Acil Müdahale Oda Kartı & 23:00/03:00 Gece Glikoz
	ui.markdown("---")
	ui.markdown("#### 🚨 Oda Bazlı Belletmen Gece Acil Müdahale Kartları & 23:00 / 03:00 Gece Kan Şekeri")
	risk_cards = get_chronic_night_risk_residents(connect_db_fn)
	if risk_cards:
		ui.dataframe(
			pd.DataFrame([
				{
					"Blok / Oda / Yatak": f"{c['block_name']} · Oda {c['room_no']} ({c['bed_bunk_no']})",
					"Öğrenci": mask_fn(c["full_name"], "name"),
					"Risk Grubu": c["risk_tags"],
					"Gece Acil Müdahale Talimatı": mask_fn(c["night_protocol"], "medical"),
					"Veli Tel": mask_fn(c["guardian_phone"], "phone"),
				}
				for c in risk_cards
			]),
			hide_index=True,
			use_container_width=True,
		)
	else:
		ui.caption("Şu an yatılı öğrencilerde tanımlı diyabet/epilepsi/astım/anafilaksi gece kartı bulunmuyor.")

	ui.download_button(
		"🖨️ Belletmen Gece Acil Müdahale & Oda Risk Kartlarını İndir (.docx)",
		data=lambda: build_night_emergency_room_cards_docx(connect_db_fn, school_name=school_nm, nurse_name=nurse_nm),
		file_name=f"Belletmen_Gece_Acil_Oda_Kartlari_{date.today().isoformat()}.docx",
		mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		type="primary",
		key="dorm_dl_night_cards_docx",
	)

	if residents:
		with ui.expander("🩸 Diyabetli Yatılı Öğrenciler İçin 23:00 ve 03:00 Gece Kan Şekeri İzlem Çizelgesi"):
			with ui.form("dorm_form_night_glucose", clear_on_submit=True):
				gc1, gc2, gc3, gc4, gc5 = ui.columns([1.8, 1.2, 0.9, 1.0, 1.5])
				g_res = gc1.selectbox("Öğrenci", list(res_map.keys()))
				g_slot = gc2.selectbox("Hedef Gece Saati", ["23:00 Gece Ölçümü", "03:00 Gece Hipoglisemi Kontrolü", "Ek Gece Ölçümü"])
				g_tm = gc3.text_input("Ölçüm Saati", value="03:00")
				g_val = gc4.number_input("Kan Şekeri (mg/dL)", min_value=20.0, max_value=600.0, value=115.0, step=1.0)
				g_int = gc5.text_input("Müdahale (Meyve Suyu / Ek Karb. / İnsülin)", placeholder="Örn: <70 ise 15g meyve suyu verildi")
				if ui.form_submit_button("🩸 Gece Glikoz Ölçümünü Kaydet", use_container_width=True):
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_night_glucose_logs(
								resident_id, log_date, target_slot, measured_time, glucose_mg_dl,
								intervention_note, measured_by, created_at
							) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
							(res_map[g_res], date.today().isoformat(), g_slot, g_tm.strip(), float(g_val), g_int.strip(), nurse_nm, _now_iso()),
						)
					if float(g_val) < 70.0:
						ui.error(f"🚨 HİPOGLİSEMİ UYARISI ({g_val:g} mg/dL)! 15g basit karbonhidrat verip 15 dk sonra tekrar ölçün!")
					else:
						ui.success("✅ Gece kan şekeri kaydı eklendi.")
					ui.rerun()

			with connect_db_fn() as conn:
				gl_rows = conn.execute(
					"""SELECT ng.log_date AS 'Tarih', ng.target_slot AS 'Plan', ng.measured_time AS 'Saat',
					          dr.full_name AS 'Öğrenci', ng.glucose_mg_dl AS 'Glikoz (mg/dL)', ng.intervention_note AS 'Müdahale'
					   FROM dorm_night_glucose_logs ng JOIN dorm_residents dr ON dr.id=ng.resident_id
					   ORDER BY ng.id DESC LIMIT 15"""
				).fetchall()
			if gl_rows:
				ui.dataframe(pd.DataFrame([dict(r) for r in gl_rows]), hide_index=True, use_container_width=True)


def _render_tab_isolation_leave_rest_sports(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	nurse_nm = str(app_cfg.get("nurse_name") or "Revir Hemşiresi").strip()
	school_nm = str(app_cfg.get("school_name") or "").strip()
	res_map = _resident_option_map(residents, mask_fn)

	alerts = get_due_return_and_dorm_alerts(connect_db_fn)
	if alerts["due_rest_returns"] or alerts["due_isolation_ends"]:
		ui.warning(
			f"🔔 **Bugün Yurda / Derse Dönüş Kontrolü Bekleyen Öğrenciler:** "
			f"Rapor süresi dolan **{len(alerts['due_rest_returns'])}** öğrenci · "
			f"İzolasyon süresi dolan **{len(alerts['due_isolation_ends'])}** öğrenci bulunuyor!"
		)
		if alerts["due_rest_returns"]:
			dr_map = {f"{mask_fn(r['full_name'], 'name')} — Bitiş: {r['end_date']} ({r['exemption_kind']}) [#{r['id']}]": int(r["id"]) for r in alerts["due_rest_returns"]}
			dc1, dc2 = ui.columns([2.5, 1.0], vertical_alignment="center")
			pick_dr = dc1.selectbox("Dönüşü Onaylanacak Raporlu/İstirahatli Öğrenci", list(dr_map.keys()), key="dorm_due_rest_pick")
			if dc2.button("✅ Yurda/Derse Döndü Onayla", type="primary", use_container_width=True, key="dorm_btn_conf_ret"):
				with connect_db_fn() as conn:
					conn.execute("UPDATE dorm_rest_and_exemptions SET returned_confirmed=1 WHERE id=?", (dr_map[pick_dr],))
				ui.success("Öğrencinin yurda/derse dönüşü onaylandı.")
				ui.rerun()

	# 1. Yurt İzolasyon (Karantina) Odası
	ui.markdown("#### 🦠 Yurt İzolasyon (Karantina) Odası & Odaya Yemek Servisi Takibi")
	if residents:
		with ui.expander("➕ İzolasyon Odasına Öğrenci Al"):
			with ui.form("dorm_form_iso_stay", clear_on_submit=True):
				is1, is2, is3, is4 = ui.columns([1.8, 1.0, 1.0, 1.0])
				i_res = is1.selectbox("Öğrenci", list(res_map.keys()))
				i_rm = is2.text_input("İzolasyon Odası", value="İzolasyon-1")
				i_st = is3.text_input("Başlangıç", value=date.today().isoformat())
				i_en = is4.text_input("Beklenen Bitiş", value=(date.today() + timedelta(days=3)).isoformat())
				is5, is6, is7 = ui.columns([1.6, 1.2, 1.0], vertical_alignment="center")
				i_dis = is5.text_input("Bulaşıcı Hastalık / Neden*", placeholder="Örn: İnfluenza A / Suçiçeği / Akut Gastroenterit")
				i_diet = is6.text_input("Odaya Yemek Diyeti", value="Sıcak Çorba / Bol Sıvı Diyet")
				i_meal = is7.checkbox("Odaya Yemek Servisi Yapılsın", value=True)
				if ui.form_submit_button("🦠 İzolasyon Kaydını Başlat", type="primary", use_container_width=True):
					if not i_dis.strip():
						ui.warning("Hastalık/neden yazın.")
					else:
						with connect_db_fn() as conn:
							conn.execute(
								"""INSERT INTO dorm_isolation_stays(
									resident_id, isolation_room_no, disease_or_reason, start_date,
									expected_end_date, room_meal_service, diet_type, status, created_at
								) VALUES (?, ?, ?, ?, ?, ?, ?, 'İzolasyonda', ?)""",
								(res_map[i_res], i_rm.strip(), i_dis.strip(), i_st[:10], i_en[:10], 1 if i_meal else 0, i_diet.strip(), _now_iso()),
							)
						ui.success("✅ Öğrenci izolasyon odasına alındı.")
						ui.rerun()

	with connect_db_fn() as conn:
		iso_list = [
			dict(r)
			for r in conn.execute(
				"""SELECT isy.*, dr.full_name, dr.room_no AS orig_room
				   FROM dorm_isolation_stays isy JOIN dorm_residents dr ON dr.id=isy.resident_id
				   ORDER BY isy.status='İzolasyonda' DESC, isy.id DESC LIMIT 15"""
			).fetchall()
		]
	if iso_list:
		ui.dataframe(
			pd.DataFrame([
				{
					"#": r["id"],
					"Durum": r["status"],
					"Öğrenci": f"{mask_fn(r['full_name'], 'name')} (Asıl Oda: {r['orig_room']})",
					"İzolasyon Odası": r["isolation_room_no"],
					"Neden": mask_fn(r["disease_or_reason"], "medical"),
					"Tarih Aralığı": f"{r['start_date']} → {r['expected_end_date']}",
					"Odaya Yemek": f"Evet ({r['diet_type']})" if r["room_meal_service"] else "Hayır",
				}
				for r in iso_list
			]),
			hide_index=True,
			use_container_width=True,
		)
		act_isos = [r for r in iso_list if r["status"] == "İzolasyonda"]
		if act_isos:
			ic1, ic2 = ui.columns([2.5, 1.0], vertical_alignment="center")
			iso_end_pick = ic1.selectbox("İzolasyonu Sonlandırılacak Öğrenci", [f"#{r['id']} — {mask_fn(r['full_name'], 'name')} ({r['isolation_room_no']})" for r in act_isos], key="dorm_iso_end_sel")
			if ic2.button("✅ İzolasyonu Bitir & Odaya Döndür", use_container_width=True, key="dorm_btn_end_iso"):
				iid = int(iso_end_pick.split("—")[0].replace("#", "").strip())
				with connect_db_fn() as conn:
					conn.execute("UPDATE dorm_isolation_stays SET status='Tamamlandı — Odasına Döndü', actual_end_date=? WHERE id=?", (date.today().isoformat(), iid))
				ui.rerun()

	# 2. Evci Çıkışı ve Yurda Dönüş Sağlık Kontrolü
	ui.markdown("---")
	ui.markdown("#### 🧳 Hafta Sonu Evci Çıkışı & Yurda Dönüş Sağlık / Semptom Taraması")
	if residents:
		with ui.expander("➕ Evci / Memleket İzni ve Dönüş Sağlık Taraması Kaydet"):
			with ui.form("dorm_form_leave_ret", clear_on_submit=True):
				lv1, lv2, lv3, lv4 = ui.columns([1.7, 1.1, 1.0, 1.0])
				l_res = lv1.selectbox("Öğrenci", list(res_map.keys()))
				l_tp = lv2.selectbox("İzin Türü", ["Hafta Sonu Evci", "Memleket / Aile İzni", "Resmi / Bayram Tatili Dönüşü"])
				l_dep = lv3.text_input("Çıkış Tarihi", value=date.today().isoformat())
				l_ret = lv4.text_input("Dönüş Tarihi", value=(date.today() + timedelta(days=2)).isoformat())
				lv5, lv6, lv7 = ui.columns([1.4, 1.6, 1.0], vertical_alignment="center")
				l_scr = lv5.selectbox("Dönüş Sağlık Taraması", ["Temiz — Odaya Alındı", "Dönüş Bekleniyor", "Semptomlu — Revir / İzolasyon", "Raporlu — Dönmedi"])
				l_nt = lv6.text_input("Ateş / Boğaz / Bit-Uyuz / Rapor Notu", placeholder="Ateş 36.6°C, bulaşıcı bulgu yok")
				l_rep = lv7.checkbox("Dış Hekim Raporu Kontrol Edildi", value=False)
				if ui.form_submit_button("💾 Evci & Dönüş Taramasını Kaydet", use_container_width=True):
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_leave_returns(
								resident_id, leave_type, depart_date, expected_return_date,
								actual_return_date, return_screening_status, fever_or_symptom_note,
								medical_report_checked, screened_by, created_at
							) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
							(
								res_map[l_res], l_tp, l_dep[:10], l_ret[:10],
								l_ret[:10] if l_scr != "Dönüş Bekleniyor" else "",
								l_scr, l_nt.strip(), 1 if l_rep else 0, nurse_nm, _now_iso(),
							),
						)
					ui.success("✅ Evci çıkış/dönüş taraması kaydedildi.")
					ui.rerun()

	# 3. Revir İstirahat Raporu & Ders / Akşam Etüdü Yoklama Senkronizasyonu
	ui.markdown("---")
	ui.markdown("#### 📋 Revir İstirahat Raporu & Sabah Dersi / Akşam Etüdü Muafiyet Listesi")
	if residents:
		with ui.expander("➕ Öğrenciye Revir İstirahati / Etüt veya Ders Muafiyeti Gir"):
			with ui.form("dorm_form_rest_ex", clear_on_submit=True):
				re1, re2, re3, re4 = ui.columns([1.7, 1.5, 0.9, 0.9])
				r_res = re1.selectbox("Öğrenci", list(res_map.keys()))
				r_knd = re2.selectbox("Muafiyet / İstirahat Türü", EXEMPTION_TYPES)
				r_st = re3.text_input("Başlangıç", value=date.today().isoformat())
				r_en = re4.text_input("Bitiş", value=date.today().isoformat())
				r_rsn = ui.text_input("Neden / Klinik Açıklama*", placeholder="Örn: Gece ateşlenme sonrası sabah derslerinden muaf / revir istirahatli")
				if ui.form_submit_button("💾 İstirahat / Etüt Muafiyeti Kaydet", type="primary", use_container_width=True):
					if not r_rsn.strip():
						ui.warning("Neden yazın.")
					else:
						with connect_db_fn() as conn:
							conn.execute(
								"""INSERT INTO dorm_rest_and_exemptions(
									resident_id, exemption_kind, start_date, end_date, reason_diagnosis, issued_by, created_at
								) VALUES (?, ?, ?, ?, ?, ?, ?)""",
								(res_map[r_res], r_knd, r_st[:10], r_en[:10], r_rsn.strip(), nurse_nm, _now_iso()),
							)
						ui.success("✅ Muafiyet/istirahat kaydı oluşturuldu.")
						ui.rerun()

	ui.download_button(
		"📝 Okul İdaresi & Belletmen İçin Bugünkü Revir İstirahatli Öğrenci Listesi (.docx)",
		data=lambda: build_rest_and_exemption_list_docx(connect_db_fn, school_name=school_nm, nurse_name=nurse_nm),
		file_name=f"Revir_Istirahatli_Etut_Muaf_Listesi_{date.today().isoformat()}.docx",
		mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		type="primary",
		key="dorm_dl_rest_list_docx",
	)

	# 4. Spor Muafiyeti & Turnuva Acil Müdahale Çantası
	ui.markdown("---")
	ui.markdown("#### 🏅 Geçici Spor / Beden Eğitimi Muafiyeti & Turnuva Sağlık Çantası Teslimi")
	sp_col1, sp_col2 = ui.columns(2)
	with sp_col1:
		if residents:
			with ui.expander("🏃 Geçici Spor ve Antrenman Muafiyeti Ekle"):
				with ui.form("dorm_form_sport_ex", clear_on_submit=True):
					sp_res = ui.selectbox("Öğrenci", list(res_map.keys()))
					sp_rsn = ui.text_input("Yaralanma / Neden*", placeholder="Örn: Sağ ayak bileği burkulması / Alçı sonrası")
					s1, s2 = ui.columns(2)
					sp_st = s1.text_input("Başlangıç", value=date.today().isoformat())
					sp_en = s2.text_input("Muafiyet Bitiş", value=(date.today() + timedelta(days=14)).isoformat())
					if ui.form_submit_button("💾 Spor Muafiyeti Kaydet", use_container_width=True) and sp_rsn.strip():
						with connect_db_fn() as conn:
							conn.execute(
								"INSERT INTO dorm_sports_exemptions(resident_id, injury_or_reason, start_date, end_date, created_at) VALUES (?, ?, ?, ?, ?)",
								(res_map[sp_res], sp_rsn.strip(), sp_st[:10], sp_en[:10], _now_iso()),
							)
						ui.success("Spor muafiyeti kaydedildi.")
						ui.rerun()
	with sp_col2:
		with ui.expander("🧰 Turnuva / Müsabaka Sağlık Çantası Teslim-Tesellüm Tutanağı"):
			with ui.form("dorm_form_tourn_kit", clear_on_submit=True):
				tk_ev = ui.text_input("Müsabaka / Turnuva Adı*", placeholder="Örn: İl Gençler Voleybol Turnuvası")
				tk1, tk2 = ui.columns(2)
				tk_dst = tk1.text_input("Gidilen Yer*", placeholder="Merkez Spor Salonu")
				tk_tch = tk2.text_input("Teslim Alan Antrenör/Öğrt.*", placeholder="Beden Eğt. Öğrt. Can Kaya")
				tk3, tk4 = ui.columns(2)
				tk_gdt = tk3.text_input("Teslim Tarihi", value=date.today().isoformat())
				tk_rdt = tk4.text_input("Beklenen İade", value=(date.today() + timedelta(days=2)).isoformat())
				if ui.form_submit_button("🧰 Çanta Teslim Kaydı Aç", use_container_width=True) and tk_ev.strip() and tk_tch.strip():
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_tournament_kits(event_name, destination, coach_or_teacher, given_date, expected_return_date, created_at)
							   VALUES (?, ?, ?, ?, ?, ?)""",
							(tk_ev.strip(), tk_dst.strip(), tk_tch.strip(), tk_gdt[:10], tk_rdt[:10], _now_iso()),
						)
					ui.success("Turnuva çantası teslim kaydı oluşturuldu.")
					ui.rerun()

		with connect_db_fn() as conn:
			last_kit = conn.execute("SELECT * FROM dorm_tournament_kits ORDER BY id DESC LIMIT 1").fetchone()
		if last_kit:
			ui.download_button(
				f"📝 Turnuva Çantası Teslim Tutanağı (#{last_kit['id']} .docx)",
				data=lambda: build_tournament_kit_docx(connect_db_fn, int(last_kit["id"]), school_name=school_nm, nurse_name=nurse_nm),
				file_name=f"Turnuva_Saglik_Cantasi_Tutanak_{last_kit['given_date']}.docx",
				mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				key="dorm_dl_kit_docx",
			)


def _render_tab_food_safety_hygiene_dysmenorrhea_devices(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	nurse_nm = str(app_cfg.get("nurse_name") or "Revir Hemşiresi").strip()
	res_map = _resident_option_map(residents, mask_fn)

	# 1. Toplu Gıda Zehirlenmesi Erken Uyarısı + 72 Saat Şahit Numune + Portör
	outbreak = detect_foodborne_outbreak_clusters(connect_db_fn)
	if outbreak["alarm"]:
		case_names = ", ".join(f"{mask_fn(c['full_name'], 'name')} ({c['room_no']} - {c['time']})" for c in outbreak["cases"])
		ui.error(
			f"🚨 **KIRMIZI ALARM — TOPLU GIDA ZEHİRLENMESİ ERKEN UYARISI!** "
			f"Bugün kısa süre içinde **{outbreak['case_count']} yatılı öğrencide** bulantı/kusma/ishal/karın ağrısı saptandı! "
			f"Şüpheli öğün: **{outbreak['suspected_meal']}** · Etkilenenler: {case_names}. "
			"Derhal **72 saatlik yemek şahit numunelerini** koruma altına alın, su kaynağını kontrol edin ve okul idaresi / İlçe Sağlık Müdürlüğü'ne bildirin!"
		)
	else:
		ui.success("🟢 **Gıda Güvenliği Erken Uyarı Durumu:** Bugün yatılı öğrencilerde kümelenmiş gastrointestinal (bulantı/kusma/ishal) vaka saptanmadı.")

	fs1, fs2 = ui.columns(2)
	with fs1:
		ui.markdown("##### 🍱 Günlük 72 Saatlik Yemek Şahit Numunesi Çizelgesi")
		with ui.form("dorm_form_food_sample", clear_on_submit=True):
			s1, s2, s3 = ui.columns([1.1, 1.3, 0.9])
			sm_dt = s1.text_input("Tarih", value=date.today().isoformat())
			sm_sl = s2.selectbox("Öğün", ["Sabah Kahvaltısı", "Öğle Yemeği", "Akşam Yemeği", "İftar / Sahur"])
			sm_tm = s3.text_input("Saat", value=datetime.now().strftime("%H:%M"))
			s4, s5 = ui.columns([1.8, 0.9])
			sm_ds = s4.text_input("Numune Alınan Yemekler*", placeholder="Örn: Mercimek çorbası, Tavuk sote, Pirinç pilavı (250g steril kap)")
			sm_fr = s5.number_input("Numune Dolabı (°C)", value=4.0, step=0.5)
			if ui.form_submit_button("🍱 72 Saatlik Şahit Numuneyi Kaydet", use_container_width=True) and sm_ds.strip():
				try:
					ret_dt = (datetime.fromisoformat(f"{sm_dt[:10]}T{sm_tm[:5]}") + timedelta(hours=72)).strftime("%Y-%m-%d %H:%M")
				except Exception:
					ret_dt = f"{(date.today() + timedelta(days=3)).isoformat()} {sm_tm[:5]}"
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT INTO dorm_food_samples(sample_date, meal_slot, sampled_dishes, taken_time, fridge_temp_c, retention_until_iso, taken_by, created_at)
						   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
						(sm_dt[:10], sm_sl, sm_ds.strip(), sm_tm[:5], float(sm_fr), ret_dt, nurse_nm, _now_iso()),
					)
				ui.success(f"✅ Şahit numune kaydedildi (Saklama bitişi: {ret_dt}).")
				ui.rerun()

	with fs2:
		ui.markdown("##### 👨‍🍳 Yemekhane Personeli Sağlık / Portör / Hijyen Muayene Takibi")
		with ui.form("dorm_form_staff_portor", clear_on_submit=True):
			p1, p2 = ui.columns(2)
			st_nm = p1.text_input("Personel Adı Soyadı*", placeholder="Örn: Aşçı Hasan Usta")
			st_rl = p2.text_input("Görevi", value="Aşçı / Mutfak Personeli")
			p3, p4 = ui.columns(2)
			st_ex = p3.text_input("Muayene Tarihi", value=date.today().isoformat())
			st_vu = p4.text_input("Geçerlilik Bitiş", value=(date.today() + timedelta(days=365)).isoformat())
			if ui.form_submit_button("👨‍🍳 Personel Sağlık Kaydını Ekle", use_container_width=True) and st_nm.strip():
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT INTO dorm_staff_health_checks(staff_name, role_title, exam_date, valid_until_date, created_at)
						   VALUES (?, ?, ?, ?, ?)""",
						(st_nm.strip(), st_rl.strip(), st_ex[:10], st_vu[:10], _now_iso()),
					)
				ui.success("Mutfak personeli sağlık kontrolü kaydedildi.")
				ui.rerun()

	# 2. Kız Pansiyonu: Dismenore İzlemi & Hijyenik Ped / Acil Destek Stoğu
	ui.markdown("---")
	ui.markdown("#### 🌸 Dismenore (Ağrılı Regl) İzlemi & Revir Hijyenik Ped / Acil Destek Stoğu")
	dy1, dy2 = ui.columns(2)
	with dy1:
		female_res = [r for r in residents if str(r.get("sex") or "").upper() in ("K", "KADIN", "KIZ")] or residents
		f_map = _resident_option_map(female_res, mask_fn)
		if f_map:
			with ui.form("dorm_form_dysmenorrhea", clear_on_submit=True):
				d_res = ui.selectbox("Öğrenci (KVKK Gizli Kayıt)", list(f_map.keys()))
				dc1, dc2, dc3 = ui.columns(3)
				d_sev = dc1.number_input("Ağrı Şiddeti (1-10)", min_value=1, max_value=10, value=6)
				d_hot = dc2.checkbox("Sıcak Uygulama", value=True)
				d_rst = dc3.checkbox("Revir İstirahati", value=True)
				d_nt = ui.text_input("Not / Hekim Yönlendirme", placeholder="Sıcak su torbası verildi, 1 saat dinlendi")
				if ui.form_submit_button("🌸 Dismenore Destek Kaydı Ekle", use_container_width=True):
					rid_d = f_map[d_res]
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_dysmenorrhea_logs(resident_id, log_date, pain_severity, hot_pack_applied, rest_given, notes, recorded_by, created_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
							(rid_d, date.today().isoformat(), int(d_sev), 1 if d_hot else 0, 1 if d_rst else 0, d_nt.strip(), nurse_nm, _now_iso()),
						)
						rep_cnt = int(conn.execute("SELECT COUNT(*) FROM dorm_dysmenorrhea_logs WHERE resident_id=? AND pain_severity>=7", (rid_d,)).fetchone()[0])
					if rep_cnt >= 2:
						ui.warning("⚠️ Bu öğrencide tekrarlayan şiddetli dismenore saptandı; anemi / endometriozis açısından Aile Hekimi veya Kadın Doğum değerlendirmesi önerilir.")
					else:
						ui.success("Kaydedildi.")
	with dy2:
		with connect_db_fn() as conn:
			h_stocks = [dict(r) for r in conn.execute("SELECT * FROM dorm_hygiene_support_stock ORDER BY id ASC").fetchall()]
		ui.dataframe(
			pd.DataFrame([
				{
					"Destek Malzemesi": s["item_name"],
					"Mevcut": s["quantity"],
					"Kritik Eşik": s["min_threshold"],
					"Durum": "🔴 Azaldı" if float(s["quantity"]) <= float(s["min_threshold"]) else "🟢 Yeterli",
				}
				for s in h_stocks
			]),
			hide_index=True,
			use_container_width=True,
		)
		hs_c1, hs_c2, hs_c3 = ui.columns([1.6, 0.9, 1.0], vertical_alignment="center")
		hs_it = hs_c1.selectbox("Stok Kalemi", [s["item_name"] for s in h_stocks], key="dorm_hs_item")
		hs_delta = hs_c2.number_input("Değişim (+/-)", value=-1.0, step=1.0, key="dorm_hs_delta")
		if hs_c3.button("📦 Stok Güncelle", use_container_width=True, key="dorm_btn_hs_upd"):
			with connect_db_fn() as conn:
				conn.execute(
					"UPDATE dorm_hygiene_support_stock SET quantity = MAX(0, quantity + ?), updated_at=? WHERE item_name=?",
					(float(hs_delta), _now_iso(), hs_it),
				)
			ui.rerun()

	# 3. Oda Çarşaf/Banyo/Tırnak Hijyeni & Özel Eğitim Gece Gözlük/İşitme Cihazı Teslimi
	ui.markdown("---")
	ui.markdown("#### 🧼 Oda Çarşaf / Kişisel Hijyen Takvimi & Özel Eğitim Gece Cihaz Teslim-Tesellümü")
	hy1, hy2 = ui.columns(2)
	with hy1:
		with ui.form("dorm_form_room_hyg", clear_on_submit=True):
			ui.markdown("**🛏️ Oda Bazlı Çarşaf, Banyo, Tırnak & Saç Kontrolü**")
			rh1, rh2 = ui.columns(2)
			h_blk = rh1.text_input("Blok", value="A Blok")
			h_rm = rh2.text_input("Oda No", value="101")
			rh3, rh4, rh5 = ui.columns(3)
			c_lin = rh3.checkbox("Çarşaf Değişti", value=True)
			c_bth = rh4.checkbox("Banyo Tamam", value=True)
			c_nl = rh5.checkbox("Tırnak/Saç Temiz", value=True)
			if ui.form_submit_button("🧼 Oda Hijyen Kontrolünü Kaydet", use_container_width=True):
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT INTO dorm_personal_hygiene_checks(check_date, block_name, room_no, linen_changed, bath_completed, nail_hair_clean, checked_by, created_at)
						   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
						(date.today().isoformat(), h_blk.strip(), h_rm.strip(), 1 if c_lin else 0, 1 if c_bth else 0, 1 if c_nl else 0, nurse_nm, _now_iso()),
					)
				ui.success("Oda hijyen kontrolü kaydedildi.")
				ui.rerun()
	with hy2:
		if residents:
			with ui.form("dorm_form_dev_custody", clear_on_submit=True):
				ui.markdown("**👓 Gözlük, İşitme Cihazı & Ortez Gece Kutu Teslimi**")
				dv_res = ui.selectbox("Öğrenci", list(res_map.keys()))
				dv1, dv2, dv3 = ui.columns(3)
				dv_tp = dv1.selectbox("Cihaz Türü", ["Gözlük", "İşitme Cihazı / Koklear", "Ortez / Protez", "Diş Plağı / Aparat"])
				dv_ntm = dv2.text_input("Gece Teslim Saati", value="22:00")
				dv_mtm = dv3.text_input("Sabah İade Saati", placeholder="07:15")
				if ui.form_submit_button("👓 Gece Cihaz Teslimini Kaydet", use_container_width=True):
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_device_night_custody(resident_id, custody_date, device_type, received_night_time, returned_morning_time, staff_name, created_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?)""",
							(res_map[dv_res], date.today().isoformat(), dv_tp, dv_ntm.strip(), dv_mtm.strip(), nurse_nm, _now_iso()),
						)
					ui.success("Gece cihaz muhafaza kaydı eklendi.")
					ui.rerun()


def _render_tab_facility_psych_parent_inspection(ctx_api: dict[str, Any], residents: list[dict[str, Any]], app_cfg: dict[str, str]) -> None:
	connect_db_fn = ctx_api["connect_db"]
	mask_fn = ctx_api.get("mask_kvkk_value") or (lambda v, *_a, **_kw: str(v or ""))
	nurse_nm = str(app_cfg.get("nurse_name") or "Revir Hemşiresi").strip()
	school_nm = str(app_cfg.get("school_name") or "").strip()
	principal_nm = str(app_cfg.get("principal_name") or "").strip()
	res_map = _resident_option_map(residents, mask_fn)

	# 1. MEB Müfettiş / İlçe Sağlık Otomatik Puanlamalı Pansiyon Denetim Dosyası
	ui.markdown("#### 📋 MEB Müfettiş & İlçe Sağlık Pansiyon Öz-Denetim Kontrol Listesi")
	insp = evaluate_meb_dorm_inspection(connect_db_fn)
	sc_c1, sc_c2, sc_c3 = ui.columns([1.0, 1.5, 1.5], vertical_alignment="center")
	sc_c1.metric("Denetim Hazırlık Puanı", f"{insp['total_score']} / 100")
	with sc_c2:
		ui.download_button(
			"📝 Resmi MEB Pansiyon Denetim Raporu (.docx)",
			data=lambda: build_meb_dorm_inspection_docx(connect_db_fn, school_name=school_nm, nurse_name=nurse_nm, principal_name=principal_nm),
			file_name=f"MEB_Pansiyon_Denetim_Raporu_{insp['inspection_date']}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			type="primary",
			use_container_width=True,
			key="dorm_dl_meb_insp_docx",
		)
	with sc_c3:
		ui.download_button(
			"📊 Tüm Okul Yurdu Kayıtları Çalışma Kitabı (.xlsx)",
			data=lambda: build_dormitory_full_excel(connect_db_fn),
			file_name=f"Okul_Yurdu_Tam_Rapor_{insp['inspection_date']}.xlsx",
			mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			use_container_width=True,
			key="dorm_dl_full_xlsx",
		)
	ui.dataframe(pd.DataFrame(insp["items"]), hide_index=True, use_container_width=True)

	# 2. Su Deposu, Ortak Banyo/WC Dezenfeksiyon & Haşere Kontrolü
	ui.markdown("---")
	fp1, fp2 = ui.columns(2)
	with fp1:
		ui.markdown("##### 💧 Su Deposu Klor/Analiz, Banyo/WC Dezenfeksiyon & İlaçlama")
		with ui.form("dorm_form_facility_san", clear_on_submit=True):
			f_tp = ui.selectbox("Kontrol Türü", FACILITY_CHECK_TYPES)
			fc1, fc2, fc3 = ui.columns(3)
			f_loc = fc1.text_input("Alan / Depo", value="Ana Su Deposu")
			f_val = fc2.text_input("Ölçüm Değeri", placeholder="Örn: 0.35 ppm Klor")
			f_nxt = fc3.text_input("Sonraki Kontrol", value=(date.today() + timedelta(days=30)).isoformat())
			f_ok = ui.checkbox("Standartlara Uygun (0.2–0.5 ppm / Dezenfekte Edildi)", value=True)
			if ui.form_submit_button("💧 Bina Hijyen / Su Kaydını Ekle", use_container_width=True):
				with connect_db_fn() as conn:
					conn.execute(
						"""INSERT INTO dorm_facility_sanitation_logs(check_date, check_type, location_area, measured_value, is_compliant, next_due_date, performed_by, created_at)
						   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
						(date.today().isoformat(), f_tp, f_loc.strip(), f_val.strip(), 1 if f_ok else 0, f_nxt[:10], nurse_nm, _now_iso()),
					)
				ui.success("Bina/su sağlığı kaydı eklendi.")
				ui.rerun()

	# 3. KVKK Gizlilik Esaslı Uyku, Enürezis & Ev Özlemi (PDR) Takibi
	with fp2:
		ui.markdown("##### 🛡️ Gizlilik Esaslı Uyku, Enürezis & Ev Özlemi (PDR) Takibi")
		if residents:
			with ui.form("dorm_form_psych_sleep", clear_on_submit=True):
				ps_res = ui.selectbox("Öğrenci (KVKK Korumalı)", list(res_map.keys()))
				ps_cat = ui.selectbox("Kategori", SLEEP_PSYCH_CATEGORIES)
				ps_obs = ui.text_input("Gizli Klinik Gözlem*", placeholder="Örn: Gece alt ıslatma / Ev özlemi ağlama nöbeti")
				ps1, ps2 = ui.columns([1.8, 1.0], vertical_alignment="center")
				ps_pln = ps1.text_input("Koruyucu Plan", placeholder="Örn: Koruyucu alez + alt ranza + 23:00 uyandırma")
				ps_pdr = ps2.checkbox("Rehberlik (PDR) Yönlendir", value=True)
				if ui.form_submit_button("🛡️ Gizli İzlem Kaydet", use_container_width=True) and ps_obs.strip():
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_psychosocial_sleep_logs(resident_id, log_date, category, confidential_observation, protective_plan, pdr_referred, recorded_by, created_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
							(res_map[ps_res], date.today().isoformat(), ps_cat, ps_obs.strip(), ps_pln.strip(), 1 if ps_pdr else 0, nurse_nm, _now_iso()),
						)
					ui.success("KVKK korumalı psikososyal/uyku izlemi kaydedildi.")
					ui.rerun()

	# 4. Hafta Sonu Veli Ziyaret ve Sağlık Görüşmesi Notları
	ui.markdown("---")
	ui.markdown("#### 👨‍👩‍👧 Hafta Sonu Veli Ziyaret ve Yüz Yüze Sağlık Görüşmesi Notları")
	if residents:
		with ui.expander("➕ Veli Ziyaret & Sağlık Görüşmesi Kaydet"):
			with ui.form("dorm_form_parent_visit", clear_on_submit=True):
				pv1, pv2, pv3, pv4 = ui.columns([1.7, 1.3, 0.9, 0.9])
				v_res = pv1.selectbox("Öğrenci", list(res_map.keys()))
				v_vnm = pv2.text_input("Ziyarete Gelen Veli Adı*", placeholder="Örn: Annesi Ayşe Kaya")
				v_dt = pv3.text_input("Tarih", value=date.today().isoformat())
				v_tm = pv4.text_input("Saat", value="14:00")
				pv5, pv6 = ui.columns(2)
				v_top = pv5.text_input("Görüşülen Konular (Beslenme, Uyku, İlaç)*", placeholder="Örn: İlaç uyumu, uyku düzeni ve iştah durumu görüşüldü")
				v_act = pv6.text_input("Alınan Kararlar / Teslim Alınan İlaç", placeholder="Örn: Yeni reçeteli demir ilacı teslim alındı")
				if ui.form_submit_button("👨‍👩‍👧 Veli Görüşmesini Kaydet", type="primary", use_container_width=True) and v_vnm.strip() and v_top.strip():
					with connect_db_fn() as conn:
						conn.execute(
							"""INSERT INTO dorm_parent_visit_notes(resident_id, visit_date, visit_time, visitor_name, topics_discussed, agreed_actions, staff_name, created_at)
							   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
							(res_map[v_res], v_dt[:10], v_tm[:5], v_vnm.strip(), v_top.strip(), v_act.strip(), nurse_nm, _now_iso()),
						)
					ui.success("✅ Veli sağlık görüşmesi kayıt altına alındı.")
					ui.rerun()

	with connect_db_fn() as conn:
		pv_rows = conn.execute(
			"""SELECT pv.visit_date AS 'Tarih', pv.visit_time AS 'Saat', dr.full_name AS 'Öğrenci',
			          pv.visitor_name AS 'Veli', pv.topics_discussed AS 'Görüşülen Konular', pv.agreed_actions AS 'Alınan Kararlar'
			   FROM dorm_parent_visit_notes pv JOIN dorm_residents dr ON dr.id=pv.resident_id
			   ORDER BY pv.id DESC LIMIT 15"""
		).fetchall()
	if pv_rows:
		ui.dataframe(
			pd.DataFrame([
				{
					"Tarih": r["Tarih"],
					"Saat": r["Saat"],
					"Öğrenci": mask_fn(r["Öğrenci"], "name"),
					"Veli": mask_fn(r["Veli"], "name"),
					"Görüşülen Konular": mask_fn(r["Görüşülen Konular"], "medical"),
					"Alınan Kararlar": r["Alınan Kararlar"] or "—",
				}
				for r in pv_rows
			]),
			hide_index=True,
			use_container_width=True,
		)


def render_school_dormitory_page(ctx_api: dict[str, Any]) -> None:
	"""🏨 Okul Yurdu / Pansiyon Sağlık ve Revir Yönetimi ana sayfasını 6 kapsamlı sekme ile oluşturur."""
	connect_db_fn = ctx_api["connect_db"]
	with connect_db_fn() as conn:
		ensure_dormitory_tables(conn)

	app_cfg = ctx_api["get_app_settings"]()
	residents = get_dorm_residents(connect_db_fn, active_only=True)
	alerts = get_due_return_and_dorm_alerts(connect_db_fn)

	ui.markdown("## 🏨 Okul Yurdu / Pansiyon Sağlık, Gece Nöbeti & Denetim Yönetimi")
	ui.caption(
		"Oda-yatak haritası (okul öğrencisi seçimi veya başka okuldan gelen pansiyoner kaydı), temaslı kümelenme, "
		"müşahede yatağı, 7/24 ilaç & gece acil kartları, izolasyon/evci/istirahat, gıda güvenliği ve MEB denetim dosyası."
	)

	boarding_cnt = sum(1 for r in residents if r.get("boarding_status") != "Gündüzlü")
	ext_cnt = sum(1 for r in residents if r.get("resident_source") == "Başka Okul Öğrencisi")
	k1, k2, k3, k4 = ui.columns(4)
	k1.metric("Yatılı Yurt Öğrencisi", f"{boarding_cnt}", f"Dış Okul: {ext_cnt}" if ext_cnt else None)
	k2.metric("Revir Müşahede Yatan", f"{alerts['active_observation_count']}")
	k3.metric("Aktif İzolasyon Odası", f"{alerts['active_isolation_count']}")
	k4.metric("Bugün Dönüş Kontrolü", f"{len(alerts['due_rest_returns']) + len(alerts['due_isolation_ends'])}")

	if alerts["food_alarm"]["alarm"]:
		ui.error(
			f"🚨 **DİKKAT — TOPLU GIDA ZEHİRLENMESİ ALARMI:** Bugün **{alerts['food_alarm']['case_count']} yatılı öğrencide** "
			f"akut bulantı/kusma/ishal kümelenmesi var! Ayrıntılar için **🧪 Gıda, Hijyen, Ped & Cihaz** sekmesini açın."
		)

	tabs = ui.tabs([
		"🛏️ Oda-Yatak & Temaslı",
		"🌙 Müşahede, Gece Nöbeti & 112",
		"💊 7/24 İlaç, Dolap & Gece Kartı",
		"🚪 İzolasyon, Evci, İstirahat & Spor",
		"🧪 Gıda, Hijyen, Ped & Cihaz",
		"🏛️ Bina, PDR/Uyku, Veli & MEB Denetim",
	], lazy=True)
	tabs.render(0, lambda: _render_tab_rooms_and_contacts(ctx_api, residents, app_cfg))
	tabs.render(1, lambda: _render_tab_observation_and_night_shift(ctx_api, residents, app_cfg))
	tabs.render(2, lambda: _render_tab_meds_cabinet_and_night_cards(ctx_api, residents, app_cfg))
	tabs.render(3, lambda: _render_tab_isolation_leave_rest_sports(ctx_api, residents, app_cfg))
	tabs.render(4, lambda: _render_tab_food_safety_hygiene_dysmenorrhea_devices(ctx_api, residents, app_cfg))
	tabs.render(5, lambda: _render_tab_facility_psych_parent_inspection(ctx_api, residents, app_cfg))
