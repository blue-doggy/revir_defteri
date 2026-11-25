# -*- coding: utf-8 -*-
"""Okul Reviri Sarf Malzemesi, İlk Yardım Çantaları, Cihaz Kalibrasyonu ve Stok Hareketleri Modülü."""
from __future__ import annotations

from datetime import date, datetime, timedelta
import html
import io
import json
import re
from typing import Any, Callable

import pandas as pd

import document_studio as dstudio
import revir_ui as ui

SUPPLY_TABLES: tuple[str, ...] = (
    "kit_inspection_logs",
    "first_aid_kit_items",
    "first_aid_kits",
    "medical_devices",
    "supply_movements",
    "infirmary_inventory",
)

MOVEMENT_TYPES: tuple[str, ...] = (
    "Giriş / Tedarik",
    "Sarf / Kullanım",
    "Sayım Düzeltme",
    "Çantaya Aktarım",
    "Karantina",
    "SKT İmha / Zayi",
)

KIT_TYPES: list[str] = [
    "Revir Ana Dolabı",
    "Afet & Acil Durum Çantası",
    "Spor Salonu & Bahçe Çantası",
    "Fen Laboratuvarı / Atölye Seti",
    "Okul Gezisi & Servis Çantası",
    "Diyabet / Anafilaksi Acil Kiti",
    "Kat İlk Yardım Dolabı",
]

DEVICE_TYPES: list[str] = [
    "Termometre (Ateş Ölçer)",
    "Glukometre (Kan Şekeri)",
    "Tansiyon Aleti",
    "Pulse Oksimetre",
    "Nebulizatör / Solunum",
    "OED / AED Defibrilatör",
    "Baskül / Boy Ölçer",
    "Diğer Tıbbi Cihaz",
]

DEFAULT_SUPPLY_ENRICHMENTS: dict[str, dict[str, Any]] = {
    "Serum Fizyolojik": {"target_stock": 30.0, "unit_cost": 18.0, "is_mandatory_meb": 1},
    "Batikon": {"target_stock": 6.0, "unit_cost": 95.0, "is_mandatory_meb": 1},
    "Steril Gazlı Bez": {"target_stock": 40.0, "unit_cost": 22.0, "is_mandatory_meb": 1},
    "Sargı Bezi": {"target_stock": 24.0, "unit_cost": 25.0, "is_mandatory_meb": 1},
    "Elastik Bandaj": {"target_stock": 12.0, "unit_cost": 48.0, "is_mandatory_meb": 1},
    "Flaster": {"target_stock": 16.0, "unit_cost": 35.0, "is_mandatory_meb": 1},
    "Yara Bandı": {"target_stock": 10.0, "unit_cost": 55.0, "is_mandatory_meb": 1},
    "Buz Aküsü": {"target_stock": 10.0, "unit_cost": 45.0, "is_mandatory_meb": 1},
    "Muayene Eldiveni": {"target_stock": 8.0, "unit_cost": 145.0, "is_mandatory_meb": 1},
    "Ateş Ölçer Probu": {"target_stock": 6.0, "unit_cost": 80.0, "is_mandatory_meb": 0},
    "Oksijen Maskesi": {"target_stock": 6.0, "unit_cost": 70.0, "is_mandatory_meb": 1},
    "Turnike": {"target_stock": 5.0, "unit_cost": 120.0, "is_mandatory_meb": 1},
    "Yanık Kremi": {"target_stock": 6.0, "unit_cost": 115.0, "is_mandatory_meb": 1},
    "Glukometre Stribi": {"target_stock": 4.0, "unit_cost": 290.0, "is_mandatory_meb": 1},
    "Lanset": {"target_stock": 4.0, "unit_cost": 110.0, "is_mandatory_meb": 1},
}

EXTRA_MEB_SUPPLY_ITEMS: list[dict[str, Any]] = [
    {
        "item_name": "Steril Yanık Örtüsü (Jelli / Alüminyum)",
        "category": "Acil Müdahale",
        "quantity": 3.0,
        "unit": "Adet",
        "min_threshold": 2.0,
        "target_stock": 5.0,
        "unit_cost": 160.0,
        "is_mandatory_meb": 1,
        "storage_location": "Acil Çantası",
    },
    {
        "item_name": "Atel Seti (Kol / Bacak Sabitleyici)",
        "category": "İlk Yardım & Ortopedi",
        "quantity": 2.0,
        "unit": "Set",
        "min_threshold": 1.0,
        "target_stock": 3.0,
        "unit_cost": 380.0,
        "is_mandatory_meb": 1,
        "storage_location": "İlk Yardım Dolabı Raf-2",
    },
    {
        "item_name": "Suni Solunum Maskesi (CPR Pocket Mask)",
        "category": "Acil Müdahale",
        "quantity": 2.0,
        "unit": "Adet",
        "min_threshold": 1.0,
        "target_stock": 4.0,
        "unit_cost": 140.0,
        "is_mandatory_meb": 1,
        "storage_location": "Acil Çantası",
    },
    {
        "item_name": "Kesici-Delici Tıbbi Atık Kutusu (3L)",
        "category": "Koruyucu Ekipman",
        "quantity": 2.0,
        "unit": "Adet",
        "min_threshold": 1.0,
        "target_stock": 4.0,
        "unit_cost": 95.0,
        "is_mandatory_meb": 1,
        "storage_location": "Pansuman Alanı",
    },
    {
        "item_name": "El Antiseptiği / Dezenfektan Solüsyon (1L)",
        "category": "Solüsyon & Antiseptik",
        "quantity": 4.0,
        "unit": "Şişe",
        "min_threshold": 2.0,
        "target_stock": 8.0,
        "unit_cost": 130.0,
        "is_mandatory_meb": 1,
        "storage_location": "Muayene Masası Yanı",
    },
]

DEFAULT_FIRST_AID_KIT_TEMPLATES: list[dict[str, Any]] = [
    {
        "kit_name": "Revir Ana İlk Yardım Dolabı (MEB Tam Standart)",
        "kit_type": "Revir Ana Dolabı",
        "location": "Revir Odası",
        "custodian_name": "Okul Hemşiresi / Revir Sorumlusu",
        "check_interval_days": 15,
        "items": [
            ("Serum Fizyolojik (%0.9 NaCl İzotonik)", 10, "Flakon / Şişe"),
            ("Batikon (Povidon İyot Antiseptik)", 2, "Şişe"),
            ("Steril Gazlı Bez (Spanç)", 15, "Paket"),
            ("Sargı Bezi (10 cm)", 8, "Rulo"),
            ("Elastik Bandaj", 4, "Rulo"),
            ("Flaster (Tıbbi Yapışkan Bant)", 4, "Rulo"),
            ("Yara Bandı", 3, "Kutu"),
            ("Muayene Eldiveni (Pudrasız)", 2, "Kutu"),
            ("Yanık Kremi / Jeli", 2, "Tüp"),
            ("Buz Aküsü (Soğuk Kompres)", 4, "Adet"),
        ],
    },
    {
        "kit_name": "Acil Müdahale & Afet / Deprem Çantası",
        "kit_type": "Afet & Acil Durum Çantası",
        "location": "Danışma / Acil Çıkış Yanı",
        "custodian_name": "İş Sağlığı & Güvenliği Ekibi",
        "check_interval_days": 30,
        "items": [
            ("Steril Gazlı Bez (Spanç)", 10, "Paket"),
            ("Sargı Bezi (10 cm)", 6, "Rulo"),
            ("Turnike (Acil Müdahale)", 2, "Adet"),
            ("Atel Seti (Kol / Bacak Sabitleyici)", 1, "Set"),
            ("Suni Solunum Maskesi (CPR Pocket Mask)", 2, "Adet"),
            ("Steril Yanık Örtüsü (Jelli / Alüminyum)", 2, "Adet"),
            ("Serum Fizyolojik (%0.9 NaCl İzotonik)", 6, "Flakon / Şişe"),
            ("Muayene Eldiveni (Pudrasız)", 1, "Kutu"),
        ],
    },
    {
        "kit_name": "Spor Salonu & Bahçe İlk Yardım Çantası",
        "kit_type": "Spor Salonu & Bahçe Çantası",
        "location": "Spor Salonu Öğretmen Odası",
        "custodian_name": "Beden Eğitimi Zümresi",
        "check_interval_days": 30,
        "items": [
            ("Elastik Bandaj", 4, "Rulo"),
            ("Buz Aküsü (Soğuk Kompres)", 4, "Adet"),
            ("Steril Gazlı Bez (Spanç)", 6, "Paket"),
            ("Batikon (Povidon İyot Antiseptik)", 1, "Şişe"),
            ("Flaster (Tıbbi Yapışkan Bant)", 2, "Rulo"),
            ("Yara Bandı", 2, "Kutu"),
        ],
    },
    {
        "kit_name": "Fen Laboratuvarı / Atölye İlk Yardım Seti",
        "kit_type": "Fen Laboratuvarı / Atölye Seti",
        "location": "Fen Laboratuvarı Duvar Dolabı",
        "custodian_name": "Fen Bilimleri Sorumlusu",
        "check_interval_days": 30,
        "items": [
            ("Serum Fizyolojik (%0.9 NaCl İzotonik)", 6, "Flakon / Şişe"),
            ("Yanık Kremi / Jeli", 2, "Tüp"),
            ("Steril Yanık Örtüsü (Jelli / Alüminyum)", 1, "Adet"),
            ("Steril Gazlı Bez (Spanç)", 6, "Paket"),
            ("Yara Bandı", 1, "Kutu"),
            ("Muayene Eldiveni (Pudrasız)", 1, "Kutu"),
        ],
    },
    {
        "kit_name": "Okul Gezisi & Servis İlk Yardım Çantası",
        "kit_type": "Okul Gezisi & Servis Çantası",
        "location": "Revir Askılığı (Taşınabilir)",
        "custodian_name": "Nöbetçi Rehber / Gezi Sorumlusu",
        "check_interval_days": 30,
        "items": [
            ("Batikon (Povidon İyot Antiseptik)", 1, "Şişe"),
            ("Steril Gazlı Bez (Spanç)", 5, "Paket"),
            ("Sargı Bezi (10 cm)", 3, "Rulo"),
            ("Flaster (Tıbbi Yapışkan Bant)", 1, "Rulo"),
            ("Yara Bandı", 1, "Kutu"),
            ("Suni Solunum Maskesi (CPR Pocket Mask)", 1, "Adet"),
        ],
    },
    {
        "kit_name": "Diyabet, Astım & Anafilaksi Acil Müdahale Kiti",
        "kit_type": "Diyabet / Anafilaksi Acil Kiti",
        "location": "Revir Acil Rafı",
        "custodian_name": "Okul Hemşiresi",
        "check_interval_days": 15,
        "items": [
            ("Glukometre Stribi (Kan Şekeri Çubuğu)", 1, "Kutu"),
            ("Lanset (Parmak Delme İğnesi)", 1, "Kutu"),
            ("Oksijen Maskesi (Pediatrik)", 2, "Adet"),
            ("Muayene Eldiveni (Pudrasız)", 1, "Kutu"),
        ],
    },
]

DEFAULT_MEDICAL_DEVICE_TEMPLATES: list[dict[str, Any]] = [
    {"device_name": "Temassız Kızılötesi Ateş Ölçer", "device_type": "Termometre (Ateş Ölçer)", "location": "Muayene Masası", "battery_status": "İyi", "cal_days": 365},
    {"device_name": "Kan Şekeri Ölçüm Cihazı (Glukometre)", "device_type": "Glukometre (Kan Şekeri)", "location": "Diyabet / Acil Rafı", "battery_status": "İyi", "cal_days": 365},
    {"device_name": "Dijital Otomatik Tansiyon Aleti", "device_type": "Tansiyon Aleti", "location": "Muayene Masası", "battery_status": "İyi", "cal_days": 365},
    {"device_name": "Parmak Tipi Pulse Oksimetre", "device_type": "Pulse Oksimetre", "location": "Acil Çantası", "battery_status": "İyi", "cal_days": 365},
    {"device_name": "Kompresörlü Nebulizatör Cihazı", "device_type": "Nebulizatör / Solunum", "location": "Revir Müdahale Köşesi", "battery_status": "Adaptör / Şebeke", "cal_days": 180},
    {"device_name": "Otomatik Eksternal Defibrilatör (OED / AED)", "device_type": "OED / AED Defibrilatör", "location": "Revir Acil Çıkışı", "battery_status": "İyi", "cal_days": 365},
]


def _clean_text(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("none", "nan", "nat") else s


def _norm_iso_date(val: Any) -> str | None:
    s = _clean_text(val)
    if not s:
        return None
    m_tr = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", s)
    if m_tr:
        d_i, m_i, y_i = int(m_tr.group(1)), int(m_tr.group(2)), int(m_tr.group(3))
        try:
            return date(y_i, m_i, d_i).isoformat()
        except ValueError:
            return None
    m_iso = re.match(r"^(\d{4})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m_iso:
        y_i, m_i, d_i = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return date(y_i, m_i, d_i).isoformat()
        except ValueError:
            return None
    return None


def _fmt_tr_date(iso_str: Any) -> str:
    norm = _norm_iso_date(iso_str)
    if not norm:
        return "—"
    try:
        return datetime.strptime(norm, "%Y-%m-%d").strftime("%d.%m.%Y")
    except ValueError:
        return norm


def ensure_supply_schema(conn: Any) -> None:
    """Sarf malzeme, hareket defteri, ilk yardım çantaları ve cihaz tablolarını oluşturur/günceller."""
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS infirmary_inventory (
            id INTEGER PRIMARY KEY,
            item_name TEXT NOT NULL,
            category TEXT,
            quantity REAL NOT NULL DEFAULT 0,
            unit TEXT,
            min_threshold REAL NOT NULL DEFAULT 5,
            expiry_date TEXT,
            storage_location TEXT,
            notes TEXT,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS supply_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inventory_id INTEGER NOT NULL,
            item_name TEXT NOT NULL,
            movement_type TEXT NOT NULL,
            delta_qty REAL NOT NULL,
            balance_after REAL NOT NULL,
            unit TEXT,
            lot_number TEXT,
            unit_cost_snapshot REAL DEFAULT 0,
            student_id INTEGER,
            visit_id INTEGER,
            kit_id INTEGER,
            performed_by TEXT,
            reason TEXT,
            movement_date TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS first_aid_kits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kit_name TEXT NOT NULL,
            kit_type TEXT NOT NULL DEFAULT 'İlk Yardım Çantası',
            location TEXT,
            custodian_name TEXT,
            seal_code TEXT,
            check_interval_days INTEGER NOT NULL DEFAULT 30,
            last_checked_date TEXT,
            last_status TEXT DEFAULT 'Kontrol Bekliyor',
            notes TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS first_aid_kit_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kit_id INTEGER NOT NULL,
            item_name TEXT NOT NULL,
            required_qty REAL NOT NULL DEFAULT 1,
            current_qty REAL NOT NULL DEFAULT 1,
            unit TEXT DEFAULT 'Adet',
            expiry_date TEXT,
            status TEXT NOT NULL DEFAULT 'Tamam',
            notes TEXT
        );
        CREATE TABLE IF NOT EXISTS kit_inspection_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kit_id INTEGER NOT NULL,
            kit_name TEXT NOT NULL,
            check_date TEXT NOT NULL,
            checked_by TEXT,
            overall_status TEXT NOT NULL,
            missing_count INTEGER NOT NULL DEFAULT 0,
            expired_count INTEGER NOT NULL DEFAULT 0,
            replenished_count INTEGER NOT NULL DEFAULT 0,
            notes TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS medical_devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_name TEXT NOT NULL,
            device_type TEXT NOT NULL DEFAULT 'Tanı Cihazı',
            serial_no TEXT,
            location TEXT,
            battery_status TEXT DEFAULT 'İyi',
            last_calibration_date TEXT,
            next_calibration_date TEXT,
            pad_or_sensor_expiry TEXT,
            status TEXT NOT NULL DEFAULT 'Çalışıyor',
            notes TEXT,
            updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_supply_movements_inv ON supply_movements(inventory_id);
        CREATE INDEX IF NOT EXISTS idx_supply_movements_date ON supply_movements(movement_date);
        CREATE INDEX IF NOT EXISTS idx_kit_items_kit_id ON first_aid_kit_items(kit_id);
        CREATE INDEX IF NOT EXISTS idx_kit_logs_kit_id ON kit_inspection_logs(kit_id);
        """
    )
    inv_cols = {r["name"] for r in conn.execute("PRAGMA table_info(infirmary_inventory)").fetchall()}
    add_cols = [
        ("lot_number", "TEXT"),
        ("target_stock", "REAL NOT NULL DEFAULT 10"),
        ("supplier_name", "TEXT"),
        ("unit_cost", "REAL NOT NULL DEFAULT 0"),
        ("barcode", "TEXT"),
        ("is_mandatory_meb", "INTEGER NOT NULL DEFAULT 0"),
        ("quarantine_status", "TEXT NOT NULL DEFAULT 'Aktif'"),
        ("last_counted_at", "TEXT"),
    ]
    for col_nm, col_def in add_cols:
        if col_nm not in inv_cols:
            conn.execute(f"ALTER TABLE infirmary_inventory ADD COLUMN {col_nm} {col_def}")


def get_enriched_inventory(connect_db_fn: Callable[[], Any]) -> list[dict[str, Any]]:
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        rows = conn.execute(
            """SELECT id, item_name, category, quantity, unit, min_threshold,
                      expiry_date, storage_location, notes, updated_at,
                      COALESCE(lot_number, '') AS lot_number,
                      COALESCE(target_stock, MAX(min_threshold * 2, 10)) AS target_stock,
                      COALESCE(supplier_name, '') AS supplier_name,
                      COALESCE(unit_cost, 0.0) AS unit_cost,
                      COALESCE(barcode, '') AS barcode,
                      COALESCE(is_mandatory_meb, 0) AS is_mandatory_meb,
                      COALESCE(quarantine_status, 'Aktif') AS quarantine_status,
                      last_counted_at
               FROM infirmary_inventory
               ORDER BY COALESCE(category, ''), item_name, id"""
        ).fetchall()
    return [dict(r) for r in rows]


def seed_extended_meb_catalog(connect_db_fn: Callable[[], Any], base_items: list[dict[str, Any]]) -> int:
    """Varsayılan 15 + 5 ek MEB standart revir sarf malzemesini ekler; mevcutları zenginleştirir."""
    now_iso = datetime.now().isoformat(timespec="minutes")
    inserted = 0
    combined = [*base_items, *EXTRA_MEB_SUPPLY_ITEMS]
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        existing = {
            _clean_text(r["item_name"]).lower(): dict(r)
            for r in conn.execute("SELECT id, item_name, target_stock, unit_cost, is_mandatory_meb FROM infirmary_inventory").fetchall()
            if _clean_text(r["item_name"])
        }
        for it in combined:
            nm = _clean_text(it.get("item_name"))
            if not nm:
                continue
            enrich = {}
            for k_e, v_e in DEFAULT_SUPPLY_ENRICHMENTS.items():
                if k_e.lower() in nm.lower():
                    enrich = v_e
                    break
            tgt = float(it.get("target_stock") or enrich.get("target_stock") or max(10.0, float(it.get("min_threshold") or 5.0) * 2.0))
            cost = float(it.get("unit_cost") or enrich.get("unit_cost") or 0.0)
            meb_f = int(it.get("is_mandatory_meb") if it.get("is_mandatory_meb") is not None else enrich.get("is_mandatory_meb", 1))
            if nm.lower() in existing:
                row_e = existing[nm.lower()]
                if float(row_e.get("unit_cost") or 0.0) == 0.0 and cost > 0.0:
                    conn.execute(
                        "UPDATE infirmary_inventory SET target_stock=?, unit_cost=?, is_mandatory_meb=? WHERE id=?",
                        (tgt, cost, meb_f, int(row_e["id"])),
                    )
                continue
            cur = conn.execute(
                """INSERT INTO infirmary_inventory(
                    item_name, category, quantity, unit, min_threshold,
                    expiry_date, storage_location, notes, updated_at,
                    target_stock, unit_cost, is_mandatory_meb, quarantine_status
                ) VALUES (?, ?, ?, ?, ?, NULL, ?, NULL, ?, ?, ?, ?, 'Aktif')""",
                (
                    nm,
                    _clean_text(it.get("category")) or "Diğer",
                    float(it.get("quantity") or 0.0),
                    _clean_text(it.get("unit")) or "Adet",
                    float(it.get("min_threshold") or 5.0),
                    _clean_text(it.get("storage_location")) or None,
                    now_iso,
                    tgt,
                    cost,
                    meb_f,
                ),
            )
            new_id = int(cur.lastrowid)
            if float(it.get("quantity") or 0.0) > 0:
                conn.execute(
                    """INSERT INTO supply_movements(
                        inventory_id, item_name, movement_type, delta_qty, balance_after,
                        unit, unit_cost_snapshot, performed_by, reason, movement_date, created_at
                    ) VALUES (?, ?, 'Giriş / Tedarik', ?, ?, ?, ?, 'Sistem', 'Varsayılan MEB Katalog Açılışı', ?, ?)""",
                    (
                        new_id,
                        nm,
                        float(it.get("quantity") or 0.0),
                        float(it.get("quantity") or 0.0),
                        _clean_text(it.get("unit")) or "Adet",
                        cost,
                        date.today().isoformat(),
                        now_iso,
                    ),
                )
            existing[nm.lower()] = {"id": new_id}
            inserted += 1
    return inserted


def record_supply_movement(
    connect_db_fn: Callable[[], Any],
    *,
    inventory_id: int,
    movement_type: str,
    delta_qty: float,
    performed_by: str = "Revir Sorumlusu",
    reason: str = "",
    student_id: int | None = None,
    visit_id: int | None = None,
    kit_id: int | None = None,
    movement_date: str | None = None,
    new_quarantine_status: str | None = None,
    lot_number: str | None = None,
) -> dict[str, Any]:
    """Stok miktarını atomik olarak günceller ve `supply_movements` defterine işler."""
    m_date = _norm_iso_date(movement_date) or date.today().isoformat()
    now_iso = datetime.now().isoformat(timespec="minutes")
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        row = conn.execute(
            "SELECT id, item_name, quantity, unit, lot_number, unit_cost, quarantine_status FROM infirmary_inventory WHERE id=?",
            (int(inventory_id),),
        ).fetchone()
        if not row:
            raise ValueError("Malzeme kaydı bulunamadı.")
        old_qty = float(row["quantity"] or 0.0)
        new_qty = max(0.0, round(old_qty + float(delta_qty), 3))
        actual_delta = round(new_qty - old_qty, 3)

        q_stat = new_quarantine_status or (
            "İmha Edildi" if "İmha" in movement_type else (
                "Karantina" if movement_type == "Karantina" else str(row["quarantine_status"] or "Aktif")
            )
        )
        lot_use = _clean_text(lot_number) or _clean_text(row["lot_number"])
        extra_sets = "updated_at=?, quantity=?, quarantine_status=?"
        params: list[Any] = [now_iso, new_qty, q_stat]
        if lot_number is not None:
            extra_sets += ", lot_number=?"
            params.append(lot_use or None)
        if movement_type == "Sayım Düzeltme":
            extra_sets += ", last_counted_at=?"
            params.append(m_date)
        params.append(int(inventory_id))
        conn.execute(f"UPDATE infirmary_inventory SET {extra_sets} WHERE id=?", params)

        cur = conn.execute(
            """INSERT INTO supply_movements(
                inventory_id, item_name, movement_type, delta_qty, balance_after,
                unit, lot_number, unit_cost_snapshot, student_id, visit_id, kit_id,
                performed_by, reason, movement_date, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                int(inventory_id),
                str(row["item_name"]),
                movement_type,
                actual_delta,
                new_qty,
                _clean_text(row["unit"]) or "Adet",
                lot_use or None,
                float(row["unit_cost"] or 0.0),
                int(student_id) if student_id else None,
                int(visit_id) if visit_id else None,
                int(kit_id) if kit_id else None,
                _clean_text(performed_by) or "Revir Sorumlusu",
                _clean_text(reason) or movement_type,
                m_date,
                now_iso,
            ),
        )
        return {
            "movement_id": int(cur.lastrowid),
            "inventory_id": int(inventory_id),
            "item_name": str(row["item_name"]),
            "old_qty": old_qty,
            "new_qty": new_qty,
            "delta_qty": actual_delta,
            "quarantine_status": q_stat,
        }


def get_supply_movements(
    connect_db_fn: Callable[[], Any],
    *,
    limit: int = 200,
    movement_type: str | None = None,
    inventory_id: int | None = None,
    days_back: int | None = None,
) -> list[dict[str, Any]]:
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        clauses: list[str] = []
        params: list[Any] = []
        if movement_type and movement_type != "Tümü":
            clauses.append("m.movement_type = ?")
            params.append(movement_type)
        if inventory_id:
            clauses.append("m.inventory_id = ?")
            params.append(int(inventory_id))
        if days_back and int(days_back) > 0:
            cutoff = (date.today() - timedelta(days=int(days_back))).isoformat()
            clauses.append("m.movement_date >= ?")
            params.append(cutoff)
        where_sql = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        sql = f"""
            SELECT m.*, s.full_name AS student_name
            FROM supply_movements m
            LEFT JOIN students s ON s.id = m.student_id
            {where_sql}
            ORDER BY m.movement_date DESC, m.id DESC
            LIMIT ?
        """
        params.append(max(1, int(limit)))
        rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def apply_physical_inventory_count(
    connect_db_fn: Callable[[], Any],
    counted_df: pd.DataFrame,
    *,
    performed_by: str = "Revir Sorumlusu",
    count_note: str = "Fiziki Stok Sayımı",
) -> dict[str, int]:
    """Toplu fiziki sayım tablosundaki farkları Sayım Düzeltme hareketi olarak işler."""
    if counted_df is None or counted_df.empty:
        return {"adjusted": 0, "matched": 0}
    today_iso = date.today().isoformat()
    adjusted = 0
    matched = 0
    for _, row in counted_df.iterrows():
        raw_id = row.get("id")
        if raw_id is None or pd.isna(raw_id):
            continue
        iid = int(float(raw_id))
        sys_qty = float(row.get("Sistem Miktarı") or 0.0)
        phy_val = row.get("Sayılan Miktar")
        if phy_val is None or pd.isna(phy_val):
            continue
        phy_qty = max(0.0, float(phy_val))
        diff = round(phy_qty - sys_qty, 3)
        if abs(diff) >= 0.001:
            note = f"{count_note} (Sistem: {sys_qty:g} -> Sayılan: {phy_qty:g})"
            record_supply_movement(
                connect_db_fn,
                inventory_id=iid,
                movement_type="Sayım Düzeltme",
                delta_qty=diff,
                performed_by=performed_by,
                reason=note,
                movement_date=today_iso,
            )
            adjusted += 1
        else:
            with connect_db_fn() as conn:
                conn.execute("UPDATE infirmary_inventory SET last_counted_at=? WHERE id=?", (today_iso, iid))
            matched += 1
    return {"adjusted": adjusted, "matched": matched}


def get_expiry_buckets(
    items: list[dict[str, Any]],
    *,
    ref_date: date | None = None,
) -> dict[str, list[dict[str, Any]]]:
    base_d = ref_date or date.today()
    today_iso = base_d.isoformat()
    soon30_iso = (base_d + timedelta(days=30)).isoformat()
    soon90_iso = (base_d + timedelta(days=90)).isoformat()
    out: dict[str, list[dict[str, Any]]] = {
        "expired": [],
        "expiring_30": [],
        "expiring_90": [],
        "quarantined": [],
        "no_expiry": [],
    }
    for it in items:
        q_stat = _clean_text(it.get("quarantine_status")) or "Aktif"
        if q_stat != "Aktif":
            out["quarantined"].append(it)
            continue
        exp_s = _norm_iso_date(it.get("expiry_date"))
        if not exp_s:
            out["no_expiry"].append(it)
        elif exp_s < today_iso:
            out["expired"].append(it)
        elif exp_s <= soon30_iso:
            out["expiring_30"].append(it)
        elif exp_s <= soon90_iso:
            out["expiring_90"].append(it)
    return out


# ==============================================================================
# İLK YARDIM ÇANTALARI, DOLAP DENETİMİ & ACİL KİT İKMAL MOTORU
# ==============================================================================

def list_first_aid_kits(connect_db_fn: Callable[[], Any]) -> list[dict[str, Any]]:
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        kits = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM first_aid_kits ORDER BY id"
            ).fetchall()
        ]
        items = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM first_aid_kit_items ORDER BY kit_id, id"
            ).fetchall()
        ]
    by_kit: dict[int, list[dict[str, Any]]] = {}
    today_iso = date.today().isoformat()
    for it in items:
        kid = int(it["kit_id"])
        by_kit.setdefault(kid, []).append(it)
    for k in kits:
        k_items = by_kit.get(int(k["id"]), [])
        k["items"] = k_items
        missing_n = sum(1 for x in k_items if float(x.get("current_qty") or 0) < float(x.get("required_qty") or 0) or _clean_text(x.get("status")) == "Eksik")
        exp_n = sum(
            1
            for x in k_items
            if (_norm_iso_date(x.get("expiry_date")) and _norm_iso_date(x.get("expiry_date")) < today_iso)
            or _clean_text(x.get("status")) == "SKT Geçmiş"
        )
        interval_d = int(k.get("check_interval_days") or 30)
        last_chk = _norm_iso_date(k.get("last_checked_date"))
        overdue = False
        if not last_chk:
            overdue = True
        else:
            try:
                overdue = (date.today() - date.fromisoformat(last_chk)).days > interval_d
            except ValueError:
                overdue = True
        k["missing_count"] = missing_n
        k["expired_count"] = exp_n
        k["is_overdue"] = overdue
        k["needs_attention"] = bool(missing_n > 0 or exp_n > 0 or overdue)
    return kits


def seed_default_first_aid_kits(connect_db_fn: Callable[[], Any], custodian_default: str = "Okul Revir Sorumlusu") -> int:
    now_iso = datetime.now().isoformat(timespec="minutes")
    inserted = 0
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        existing_names = {
            _clean_text(r["kit_name"]).lower()
            for r in conn.execute("SELECT kit_name FROM first_aid_kits").fetchall()
        }
        for tpl in DEFAULT_FIRST_AID_KIT_TEMPLATES:
            knm = _clean_text(tpl["kit_name"])
            if knm.lower() in existing_names:
                continue
            cust = custodian_default if "Hemşiresi" in tpl["custodian_name"] and custodian_default else tpl["custodian_name"]
            cur = conn.execute(
                """INSERT INTO first_aid_kits(
                    kit_name, kit_type, location, custodian_name,
                    check_interval_days, last_checked_date, last_status, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, 'Tamam (Hazır)', ?)""",
                (
                    knm,
                    tpl["kit_type"],
                    tpl["location"],
                    cust,
                    int(tpl.get("check_interval_days") or 30),
                    date.today().isoformat(),
                    now_iso,
                ),
            )
            kid = int(cur.lastrowid)
            for item_nm, req_q, u_str in tpl["items"]:
                conn.execute(
                    """INSERT INTO first_aid_kit_items(
                        kit_id, item_name, required_qty, current_qty, unit, status
                    ) VALUES (?, ?, ?, ?, ?, 'Tamam')""",
                    (kid, item_nm, float(req_q), float(req_q), u_str),
                )
            existing_names.add(knm.lower())
            inserted += 1
    return inserted


def create_first_aid_kit(
    connect_db_fn: Callable[[], Any],
    *,
    kit_name: str,
    kit_type: str,
    location: str,
    custodian_name: str,
    seal_code: str = "",
    check_interval_days: int = 30,
    notes: str = "",
    starter_items: list[tuple[str, float, str]] | None = None,
) -> int:
    knm = _clean_text(kit_name)
    if not knm:
        raise ValueError("Çanta / dolap adı boş olamaz.")
    now_iso = datetime.now().isoformat(timespec="minutes")
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        cur = conn.execute(
            """INSERT INTO first_aid_kits(
                kit_name, kit_type, location, custodian_name, seal_code,
                check_interval_days, last_status, notes, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'Kontrol Bekliyor', ?, ?)""",
            (
                knm,
                _clean_text(kit_type) or "İlk Yardım Çantası",
                _clean_text(location) or None,
                _clean_text(custodian_name) or None,
                _clean_text(seal_code) or None,
                max(1, int(check_interval_days or 30)),
                _clean_text(notes) or None,
                now_iso,
            ),
        )
        kid = int(cur.lastrowid)
        for nm, rq, u_s in (starter_items or []):
            if _clean_text(nm):
                conn.execute(
                    """INSERT INTO first_aid_kit_items(
                        kit_id, item_name, required_qty, current_qty, unit, status
                    ) VALUES (?, ?, ?, ?, ?, 'Tamam')""",
                    (kid, _clean_text(nm), max(0.0, float(rq)), max(0.0, float(rq)), _clean_text(u_s) or "Adet"),
                )
        return kid


def save_kit_items_and_inspect(
    connect_db_fn: Callable[[], Any],
    *,
    kit_id: int,
    edited_items_df: pd.DataFrame,
    checked_by: str = "Revir Sorumlusu",
    inspection_notes: str = "",
    replenish_from_depot: bool = False,
) -> dict[str, Any]:
    """Çanta içindekiler tablosunu kaydeder, denetim tutanağı oluşturur ve istenirse ana depodan eksikleri tamamlar."""
    now_iso = datetime.now().isoformat(timespec="minutes")
    today_iso = date.today().isoformat()
    replenished_count = 0
    replenished_notes: list[str] = []

    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        k_row = conn.execute("SELECT * FROM first_aid_kits WHERE id=?", (int(kit_id),)).fetchone()
        if not k_row:
            raise ValueError("İlk yardım çantası bulunamadı.")
        kit_name = str(k_row["kit_name"])

        if edited_items_df is not None and not edited_items_df.empty:
            for _, row in edited_items_df.iterrows():
                nm = _clean_text(row.get("Malzeme"))
                if not nm:
                    continue
                req_q = max(0.0, float(row.get("Gerekli") if not pd.isna(row.get("Gerekli")) else 1.0))
                cur_q = max(0.0, float(row.get("Mevcut") if not pd.isna(row.get("Mevcut")) else 0.0))
                unit_s = _clean_text(row.get("Birim")) or "Adet"
                exp_s = _norm_iso_date(row.get("SKT (YYYY-AA-GG)"))
                st_s = _clean_text(row.get("Durum")) or ("Tamam" if cur_q >= req_q else "Eksik")
                if exp_s and exp_s < today_iso:
                    st_s = "SKT Geçmiş"
                note_s = _clean_text(row.get("Not")) or None
                raw_iid = row.get("id")
                has_iid = False
                iid_val = 0
                if raw_iid is not None and not pd.isna(raw_iid):
                    try:
                        iid_val = int(float(raw_iid))
                        has_iid = iid_val > 0
                    except (TypeError, ValueError):
                        has_iid = False
                if has_iid:
                    conn.execute(
                        """UPDATE first_aid_kit_items
                           SET item_name=?, required_qty=?, current_qty=?, unit=?, expiry_date=?, status=?, notes=?
                           WHERE id=? AND kit_id=?""",
                        (nm, req_q, cur_q, unit_s, exp_s, st_s, note_s, iid_val, int(kit_id)),
                    )
                else:
                    conn.execute(
                        """INSERT INTO first_aid_kit_items(
                            kit_id, item_name, required_qty, current_qty, unit, expiry_date, status, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (int(kit_id), nm, req_q, cur_q, unit_s, exp_s, st_s, note_s),
                    )

        kit_items = [
            dict(r)
            for r in conn.execute("SELECT * FROM first_aid_kit_items WHERE kit_id=?", (int(kit_id),)).fetchall()
        ]
        inv_rows = [
            dict(r)
            for r in conn.execute("SELECT id, item_name, quantity, unit FROM infirmary_inventory WHERE COALESCE(quarantine_status,'Aktif')='Aktif'").fetchall()
        ]

    if replenish_from_depot:
        for ki in kit_items:
            req_q = float(ki.get("required_qty") or 0.0)
            cur_q = float(ki.get("current_qty") or 0.0)
            shortfall = round(req_q - cur_q, 3)
            if shortfall <= 0:
                continue
            ki_fold = dstudio.tr_fold(ki["item_name"])
            matched_inv = None
            for inv_r in inv_rows:
                inv_fold = dstudio.tr_fold(inv_r["item_name"])
                if inv_fold == ki_fold or ki_fold in inv_fold or inv_fold in ki_fold:
                    if float(inv_r.get("quantity") or 0.0) > 0:
                        matched_inv = inv_r
                        break
            if not matched_inv:
                continue
            avail = float(matched_inv.get("quantity") or 0.0)
            xfer = min(shortfall, avail)
            if xfer <= 0:
                continue
            record_supply_movement(
                connect_db_fn,
                inventory_id=int(matched_inv["id"]),
                movement_type="Çantaya Aktarım",
                delta_qty=-xfer,
                performed_by=checked_by,
                reason=f"Çanta ikmali: {kit_name}",
                kit_id=int(kit_id),
                movement_date=today_iso,
            )
            matched_inv["quantity"] = avail - xfer
            new_cur_q = round(cur_q + xfer, 3)
            new_st = "Tamam" if new_cur_q >= req_q else "Eksik"
            with connect_db_fn() as conn:
                conn.execute(
                    "UPDATE first_aid_kit_items SET current_qty=?, status=? WHERE id=?",
                    (new_cur_q, new_st, int(ki["id"])),
                )
            ki["current_qty"] = new_cur_q
            ki["status"] = new_st
            replenished_count += 1
            replenished_notes.append(f"{ki['item_name']} (+{xfer:g})")

    missing_n = sum(1 for x in kit_items if float(x.get("current_qty") or 0) < float(x.get("required_qty") or 0))
    exp_n = sum(
        1
        for x in kit_items
        if (_norm_iso_date(x.get("expiry_date")) and _norm_iso_date(x.get("expiry_date")) < today_iso)
        or _clean_text(x.get("status")) == "SKT Geçmiş"
    )
    overall = "Tamam (Hazır)" if (missing_n == 0 and exp_n == 0) else ("SKT / Eksik Müdahale Gerekli" if exp_n > 0 else "Eksik Var")

    with connect_db_fn() as conn:
        conn.execute(
            "UPDATE first_aid_kits SET last_checked_date=?, last_status=? WHERE id=?",
            (today_iso, overall, int(kit_id)),
        )
        full_note = " · ".join(filter(None, [_clean_text(inspection_notes), ("İkmal: " + ", ".join(replenished_notes)) if replenished_notes else ""]))
        conn.execute(
            """INSERT INTO kit_inspection_logs(
                kit_id, kit_name, check_date, checked_by, overall_status,
                missing_count, expired_count, replenished_count, notes, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                int(kit_id),
                kit_name,
                today_iso,
                _clean_text(checked_by) or "Revir Sorumlusu",
                overall,
                missing_n,
                exp_n,
                replenished_count,
                full_note or None,
                now_iso,
            ),
        )
    return {
        "kit_name": kit_name,
        "overall_status": overall,
        "missing_count": missing_n,
        "expired_count": exp_n,
        "replenished_count": replenished_count,
        "replenished_items": replenished_notes,
    }


def list_kit_inspection_logs(connect_db_fn: Callable[[], Any], limit: int = 40) -> list[dict[str, Any]]:
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        rows = conn.execute(
            "SELECT * FROM kit_inspection_logs ORDER BY check_date DESC, id DESC LIMIT ?",
            (max(1, int(limit)),),
        ).fetchall()
    return [dict(r) for r in rows]


# ==============================================================================
# TIBBİ CİHAZ, KALİBRASYON, BATARYA & AED PED TAKİBİ
# ==============================================================================

def list_medical_devices(connect_db_fn: Callable[[], Any], ref_date: date | None = None) -> list[dict[str, Any]]:
    base_d = ref_date or date.today()
    today_iso = base_d.isoformat()
    soon_iso = (base_d + timedelta(days=30)).isoformat()
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        rows = [dict(r) for r in conn.execute("SELECT * FROM medical_devices ORDER BY device_type, device_name, id").fetchall()]
    for d in rows:
        warns: list[str] = []
        st_val = _clean_text(d.get("status")) or "Çalışıyor"
        bat_val = _clean_text(d.get("battery_status")) or "İyi"
        if st_val != "Çalışıyor":
            warns.append(st_val)
        if "Zayıf" in bat_val or "Değiş" in bat_val:
            warns.append(f"Pil: {bat_val}")
        ncd = _norm_iso_date(d.get("next_calibration_date"))
        if ncd:
            if ncd < today_iso:
                warns.append("Kalibrasyon Gecikmiş")
            elif ncd <= soon_iso:
                warns.append("Kalibrasyon Yakın")
        pad_exp = _norm_iso_date(d.get("pad_or_sensor_expiry"))
        if pad_exp:
            if pad_exp < today_iso:
                warns.append("Ped/Sensör SKT Geçmiş")
            elif pad_exp <= soon_iso:
                warns.append("Ped/Sensör SKT Yakın")
        d["alert_flags"] = warns
        d["needs_attention"] = bool(warns)
    return rows


def seed_default_medical_devices(connect_db_fn: Callable[[], Any]) -> int:
    today_d = date.today()
    now_iso = datetime.now().isoformat(timespec="minutes")
    added = 0
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        existing = {_clean_text(r["device_name"]).lower() for r in conn.execute("SELECT device_name FROM medical_devices").fetchall()}
        for tpl in DEFAULT_MEDICAL_DEVICE_TEMPLATES:
            if tpl["device_name"].lower() in existing:
                continue
            next_c = (today_d + timedelta(days=int(tpl.get("cal_days") or 365))).isoformat()
            pad_e = (today_d + timedelta(days=540)).isoformat() if "AED" in tpl["device_name"] else None
            conn.execute(
                """INSERT INTO medical_devices(
                    device_name, device_type, location, battery_status,
                    last_calibration_date, next_calibration_date, pad_or_sensor_expiry,
                    status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'Çalışıyor', ?)""",
                (
                    tpl["device_name"],
                    tpl["device_type"],
                    tpl["location"],
                    tpl["battery_status"],
                    today_d.isoformat(),
                    next_c,
                    pad_e,
                    now_iso,
                ),
            )
            added += 1
    return added


def upsert_medical_device(
    connect_db_fn: Callable[[], Any],
    *,
    device_name: str,
    device_type: str = "Diğer Tıbbi Cihaz",
    serial_no: str = "",
    location: str = "",
    battery_status: str = "İyi",
    last_calibration_date: str | None = None,
    next_calibration_date: str | None = None,
    pad_or_sensor_expiry: str | None = None,
    status: str = "Çalışıyor",
    notes: str = "",
    device_id: int | None = None,
) -> int:
    nm = _clean_text(device_name)
    if not nm:
        raise ValueError("Cihaz adı boş olamaz.")
    now_iso = datetime.now().isoformat(timespec="minutes")
    vals = (
        nm,
        _clean_text(device_type) or "Diğer Tıbbi Cihaz",
        _clean_text(serial_no) or None,
        _clean_text(location) or None,
        _clean_text(battery_status) or "İyi",
        _norm_iso_date(last_calibration_date),
        _norm_iso_date(next_calibration_date),
        _norm_iso_date(pad_or_sensor_expiry),
        _clean_text(status) or "Çalışıyor",
        _clean_text(notes) or None,
        now_iso,
    )
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        if device_id:
            conn.execute(
                """UPDATE medical_devices
                   SET device_name=?, device_type=?, serial_no=?, location=?, battery_status=?,
                       last_calibration_date=?, next_calibration_date=?, pad_or_sensor_expiry=?,
                       status=?, notes=?, updated_at=?
                   WHERE id=?""",
                (*vals, int(device_id)),
            )
            return int(device_id)
        cur = conn.execute(
            """INSERT INTO medical_devices(
                device_name, device_type, serial_no, location, battery_status,
                last_calibration_date, next_calibration_date, pad_or_sensor_expiry,
                status, notes, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            vals,
        )
        return int(cur.lastrowid)


# ==============================================================================
# HAZIRLIK SKORU, TÜKETİM TRENDİ & SATIN ALMA / BÜTÇE PLANLAYICI
# ==============================================================================

def get_consumption_summary(connect_db_fn: Callable[[], Any], *, days: int = 30) -> list[dict[str, Any]]:
    cutoff = (date.today() - timedelta(days=max(1, int(days)))).isoformat()
    with connect_db_fn() as conn:
        ensure_supply_schema(conn)
        inv_map = {
            int(r["id"]): dict(r)
            for r in conn.execute("SELECT id, item_name, quantity, unit, category, unit_cost FROM infirmary_inventory").fetchall()
        }
        rows = conn.execute(
            """SELECT inventory_id, item_name, unit,
                      SUM(ABS(delta_qty)) AS total_consumed,
                      COUNT(*) AS move_count
               FROM supply_movements
               WHERE delta_qty < 0
                 AND movement_type IN ('Sarf / Kullanım', 'Çantaya Aktarım', 'SKT İmha / Zayi')
                 AND movement_date >= ?
               GROUP BY inventory_id
               ORDER BY total_consumed DESC""",
            (cutoff,),
        ).fetchall()
    out: list[dict[str, Any]] = []
    for r in rows:
        iid = int(r["inventory_id"])
        inv = inv_map.get(iid) or {}
        consumed = float(r["total_consumed"] or 0.0)
        cur_qty = float(inv.get("quantity") or 0.0)
        daily_rate = consumed / float(max(1, int(days)))
        est_days = round(cur_qty / daily_rate, 1) if daily_rate > 0 else None
        out.append({
            "inventory_id": iid,
            "item_name": str(inv.get("item_name") or r["item_name"]),
            "category": _clean_text(inv.get("category")) or "Diğer",
            "unit": _clean_text(inv.get("unit") or r["unit"]) or "Adet",
            "consumed_qty": round(consumed, 2),
            "move_count": int(r["move_count"] or 0),
            "daily_rate": round(daily_rate, 3),
            "current_qty": cur_qty,
            "estimated_days_left": est_days,
        })
    return out


def build_procurement_plan(connect_db_fn: Callable[[], Any], *, include_kit_shortfalls: bool = True) -> dict[str, Any]:
    items = get_enriched_inventory(connect_db_fn)
    kits = list_first_aid_kits(connect_db_fn) if include_kit_shortfalls else []
    kit_shortfalls: dict[str, float] = {}
    for k in kits:
        for ki in k.get("items") or []:
            diff = max(0.0, float(ki.get("required_qty") or 0) - float(ki.get("current_qty") or 0))
            if diff > 0:
                fk = dstudio.tr_fold(ki["item_name"])
                kit_shortfalls[fk] = kit_shortfalls.get(fk, 0.0) + diff

    cons_map = {c["inventory_id"]: c for c in get_consumption_summary(connect_db_fn, days=30)}
    rows: list[dict[str, Any]] = []
    total_budget = 0.0
    today_iso = date.today().isoformat()

    for it in items:
        cur_q = float(it.get("quantity") or 0.0)
        min_t = float(it.get("min_threshold") or 0.0)
        tgt_q = max(float(it.get("target_stock") or 0.0), min_t * 2.0, 1.0)
        cost = float(it.get("unit_cost") or 0.0)
        exp_s = _norm_iso_date(it.get("expiry_date"))
        is_exp = bool(exp_s and exp_s < today_iso)
        is_quar = (_clean_text(it.get("quarantine_status")) or "Aktif") != "Aktif"

        extra_kit = 0.0
        fk_it = dstudio.tr_fold(it["item_name"])
        for kf, kv in kit_shortfalls.items():
            if kf == fk_it or kf in fk_it or fk_it in kf:
                extra_kit += kv

        effective_cur = 0.0 if (is_exp or is_quar) else cur_q
        need_qty = max(0.0, round((tgt_q - effective_cur) + extra_kit, 2))
        if cur_q <= min_t or is_exp or is_quar or extra_kit > 0:
            order_qty = need_qty if need_qty > 0 else max(1.0, min_t)
            line_cost = round(order_qty * cost, 2)
            total_budget += line_cost
            prio = "🔴 Acil" if (cur_q <= 0 or is_exp or is_quar) else ("🟠 Yüksek" if cur_q <= min_t else "🟡 Planlı")
            cons = cons_map.get(int(it["id"])) or {}
            rows.append({
                "id": int(it["id"]),
                "Öncelik": prio,
                "Malzeme Adı": it["item_name"],
                "Kategori": _clean_text(it.get("category")) or "Diğer",
                "Mevcut": cur_q,
                "Eşik / Hedef": f"{min_t:g} / {tgt_q:g}",
                "Çanta Eksiği": round(extra_kit, 2),
                "Önerilen Alım": order_qty,
                "Birim": _clean_text(it.get("unit")) or "Adet",
                "Birim Fiyat (TL)": round(cost, 2),
                "Tahmini Tutar (TL)": line_cost,
                "30G Tüketim": float(cons.get("consumed_qty") or 0.0),
                "Tedarikçi": _clean_text(it.get("supplier_name")) or "—",
            })

    return {"rows": rows, "total_budget_tl": round(total_budget, 2), "item_count": len(rows)}


def compute_infirmary_readiness_summary(
    connect_db_fn: Callable[[], Any],
    *,
    days_horizon: int = 30,
    ref_date: date | None = None,
) -> dict[str, Any]:
    base_d = ref_date or date.today()
    items = get_enriched_inventory(connect_db_fn)
    buckets = get_expiry_buckets(items, ref_date=base_d)
    soon_iso = (base_d + timedelta(days=max(1, int(days_horizon)))).isoformat()
    today_iso = base_d.isoformat()

    active_items = [it for it in items if (_clean_text(it.get("quarantine_status")) or "Aktif") == "Aktif"]
    critical_items = [it for it in active_items if float(it.get("quantity") or 0.0) <= float(it.get("min_threshold") or 0.0)]
    out_of_stock = [it for it in active_items if float(it.get("quantity") or 0.0) <= 0.0]

    expiring_soon: list[dict[str, Any]] = []
    for it in active_items:
        exp_s = _norm_iso_date(it.get("expiry_date"))
        if exp_s and today_iso <= exp_s <= soon_iso:
            expiring_soon.append(it)

    kits = list_first_aid_kits(connect_db_fn)
    kits_attention = [k for k in kits if k.get("needs_attention")]

    devices = list_medical_devices(connect_db_fn, ref_date=base_d)
    devices_attention = [d for d in devices if d.get("needs_attention")]

    total_val = sum(float(it.get("quantity") or 0.0) * float(it.get("unit_cost") or 0.0) for it in active_items)

    # Hazırlık skoru: 100 üzerinden eksikler ve SKT/cihaz alarmları düşülür
    penalty = (
        min(32, len(out_of_stock) * 8)
        + min(24, max(0, len(critical_items) - len(out_of_stock)) * 4)
        + min(24, len(buckets["expired"]) * 8)
        + min(16, len(kits_attention) * 6)
        + min(14, len(devices_attention) * 5)
    )
    score = max(0, min(100, 100 - penalty)) if (items or kits or devices) else 0
    if not items and not kits and not devices:
        tier = "Envanter Henüz Kurulmadı"
    elif score >= 88:
        tier = "Mükemmel Hazırlık"
    elif score >= 70:
        tier = "İyi — Küçük Eksikler Var"
    elif score >= 45:
        tier = "Dikkat — İkmal / Kontrol Gerekli"
    else:
        tier = "Kritik — Acil İkmal ve Denetim Şart"

    return {
        "total_items": len(items),
        "active_items": len(active_items),
        "total_value_tl": round(total_val, 2),
        "critical_items": critical_items,
        "out_of_stock": out_of_stock,
        "expired_items": buckets["expired"],
        "expiring_soon_items": expiring_soon,
        "expiring_90_items": buckets["expiring_90"],
        "quarantined_items": buckets["quarantined"],
        "kits": kits,
        "kits_attention": kits_attention,
        "devices": devices,
        "devices_attention": devices_attention,
        "readiness_score": int(score),
        "readiness_tier": tier,
    }


# ==============================================================================
# ÇOKLU RAPOR ÇIKTILARI: WORD (.DOCX), EXCEL (.XLSX) VE A4 PDF (.PDF)
# ==============================================================================

def build_supply_excel_bytes(connect_db_fn: Callable[[], Any]) -> bytes:
    items = get_enriched_inventory(connect_db_fn)
    proc = build_procurement_plan(connect_db_fn)
    moves = get_supply_movements(connect_db_fn, limit=500)
    kits = list_first_aid_kits(connect_db_fn)
    devices = list_medical_devices(connect_db_fn)

    df_inv = pd.DataFrame([
        {
            "ID": it["id"],
            "Malzeme Adı": it["item_name"],
            "Kategori": _clean_text(it.get("category")) or "Diğer",
            "Mevcut Miktar": float(it.get("quantity") or 0),
            "Kritik Eşik": float(it.get("min_threshold") or 0),
            "Hedef Stok": float(it.get("target_stock") or 0),
            "Birim": _clean_text(it.get("unit")) or "Adet",
            "Birim Maliyet (TL)": float(it.get("unit_cost") or 0),
            "SKT": _fmt_tr_date(it.get("expiry_date")),
            "Parti/Lot": _clean_text(it.get("lot_number")) or "—",
            "Barkod": _clean_text(it.get("barcode")) or "—",
            "Tedarikçi": _clean_text(it.get("supplier_name")) or "—",
            "Saklama Yeri": _clean_text(it.get("storage_location")) or "—",
            "Durum": _clean_text(it.get("quarantine_status")) or "Aktif",
        }
        for it in items
    ])
    df_proc = pd.DataFrame(proc["rows"]) if proc["rows"] else pd.DataFrame([{"Bilgi": "Acil satın alma gereken ürün yok"}])
    df_moves = pd.DataFrame([
        {
            "Tarih": _fmt_tr_date(m.get("movement_date")),
            "Malzeme": m.get("item_name"),
            "İşlem": m.get("movement_type"),
            "Değişim": float(m.get("delta_qty") or 0),
            "Kalan": float(m.get("balance_after") or 0),
            "Birim": m.get("unit") or "",
            "Kişi / Vaka": _clean_text(m.get("student_name")) or "—",
            "Yapan": _clean_text(m.get("performed_by")) or "—",
            "Açıklama": _clean_text(m.get("reason")) or "—",
        }
        for m in moves
    ]) if moves else pd.DataFrame([{"Bilgi": "Henüz stok hareketi yok"}])
    df_kits = pd.DataFrame([
        {
            "Çanta / Dolap Adı": k["kit_name"],
            "Tür": k.get("kit_type"),
            "Konum": _clean_text(k.get("location")) or "—",
            "Sorumlu": _clean_text(k.get("custodian_name")) or "—",
            "Son Kontrol": _fmt_tr_date(k.get("last_checked_date")),
            "Eksik Kalem": k.get("missing_count", 0),
            "SKT Geçen": k.get("expired_count", 0),
            "Durum": k.get("last_status"),
        }
        for k in kits
    ]) if kits else pd.DataFrame([{"Bilgi": "Kayıtlı ilk yardım çantası yok"}])
    df_dev = pd.DataFrame([
        {
            "Cihaz Adı": d["device_name"],
            "Tür": d.get("device_type"),
            "Seri No": _clean_text(d.get("serial_no")) or "—",
            "Konum": _clean_text(d.get("location")) or "—",
            "Pil": _clean_text(d.get("battery_status")) or "—",
            "Sonraki Kalibrasyon": _fmt_tr_date(d.get("next_calibration_date")),
            "Ped/Sensör SKT": _fmt_tr_date(d.get("pad_or_sensor_expiry")),
            "Durum": d.get("status"),
            "Uyarılar": ", ".join(d.get("alert_flags") or []) or "Uygun",
        }
        for d in devices
    ]) if devices else pd.DataFrame([{"Bilgi": "Kayıtlı tıbbi cihaz yok"}])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_inv.to_excel(writer, index=False, sheet_name="Ana Envanter")
        df_proc.to_excel(writer, index=False, sheet_name="Satin Alma Plani")
        df_moves.to_excel(writer, index=False, sheet_name="Stok Hareketleri")
        df_kits.to_excel(writer, index=False, sheet_name="Ilk Yardim Cantalari")
        df_dev.to_excel(writer, index=False, sheet_name="Tibbi Cihazlar")
        for ws in writer.book.worksheets:
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    return buf.getvalue()


def build_supply_a4_pdf_bytes(connect_db_fn: Callable[[], Any], app_cfg: dict[str, Any]) -> bytes:
    summary = compute_infirmary_readiness_summary(connect_db_fn)
    items = get_enriched_inventory(connect_db_fn)
    proc = build_procurement_plan(connect_db_fn)
    school = html.escape(_clean_text(app_cfg.get("school_name")) or "Okul Reviri")
    nurse = html.escape(_clean_text(app_cfg.get("nurse_name")) or "Okul Revir Sorumlusu")
    today_tr = date.today().strftime("%d.%m.%Y")

    rows_html: list[str] = []
    for idx, it in enumerate(items, start=1):
        rows_html.append(
            "<tr>"
            f"<td>{idx}</td>"
            f"<td>{html.escape(str(it['item_name']))}</td>"
            f"<td>{html.escape(_clean_text(it.get('category')) or '—')}</td>"
            f"<td>{float(it.get('quantity') or 0):g} {html.escape(_clean_text(it.get('unit')) or '')}</td>"
            f"<td>{float(it.get('min_threshold') or 0):g}</td>"
            f"<td>{html.escape(_fmt_tr_date(it.get('expiry_date')))}</td>"
            f"<td>{html.escape(_clean_text(it.get('quarantine_status')) or 'Aktif')}</td>"
            "</tr>"
        )

    proc_rows_html: list[str] = []
    for pr in proc["rows"][:25]:
        proc_rows_html.append(
            "<tr>"
            f"<td>{html.escape(str(pr['Öncelik']))}</td>"
            f"<td>{html.escape(str(pr['Malzeme Adı']))}</td>"
            f"<td>{float(pr['Önerilen Alım']):g} {html.escape(str(pr['Birim']))}</td>"
            f"<td>{float(pr['Tahmini Tutar (TL)']):,.2f} TL</td>"
            "</tr>"
        )

    kits_html = "".join(
        "<li><b>{}</b> ({}) — Kontrol: {} · Eksik: {} · Durum: {}</li>".format(
            html.escape(k["kit_name"]),
            html.escape(_clean_text(k.get("location")) or "—"),
            html.escape(_fmt_tr_date(k.get("last_checked_date"))),
            int(k.get("missing_count") or 0),
            html.escape(_clean_text(k.get("last_status")) or "—"),
        )
        for k in summary["kits"]
    )

    rows_body = "".join(rows_html) if rows_html else '<tr><td colspan="7">Kayıtlı malzeme bulunmamaktadır.</td></tr>'
    proc_body = "".join(proc_rows_html) if proc_rows_html else '<tr><td colspan="4">Acil ikmal gereken malzeme yok.</td></tr>'
    kits_body = kits_html if kits_html else "<li>Kayıtlı ilk yardım çantası bulunmuyor.</li>"
    body = (
        f"<h1>{school} — REVİR İLK YARDIM & SARF STOK DENETİM RAPORU</h1>"
        f"<p><b>Tarih:</b> {today_tr} · <b>Sorumlu:</b> {nurse} · "
        f"<b>Hazırlık Skoru:</b> %{summary['readiness_score']} ({html.escape(summary['readiness_tier'])}) · "
        f"<b>Envanter Değeri:</b> {summary['total_value_tl']:,.2f} TL</p>"
        "<h2>1. Revir Sarf Malzemesi Mevcut Envanteri</h2>"
        "<table><thead><tr><th>No</th><th>Malzeme</th><th>Kategori</th><th>Miktar</th><th>Eşik</th><th>SKT</th><th>Statü</th></tr></thead>"
        f"<tbody>{rows_body}</tbody></table>"
        f"<h2>2. Önerilen İkmal & Satın Alma İstem Listesi (Tahmini Toplam: {proc['total_budget_tl']:,.2f} TL)</h2>"
        "<table><thead><tr><th>Öncelik</th><th>Malzeme</th><th>Önerilen Alım</th><th>Tahmini Tutar</th></tr></thead>"
        f"<tbody>{proc_body}</tbody></table>"
        f"<h2>3. İlk Yardım Çantaları ve Acil Kitler ({len(summary['kits'])} Birim)</h2>"
        f"<ul>{kits_body}</ul>"
    )
    return dstudio.render_html_body_to_pdf_bytes(body, font_family="Arial", font_size_pt=9.5, landscape=False, margin_pt=32.0)


# ==============================================================================
# ANA ARAYÜZ: 6 SEKMELİ REVİR SARF MALZEMESİ & İLK YARDIM KOMUTA MERKEZİ
# ==============================================================================

def render_infirmary_supply_studio(ctx_api: dict[str, Any]) -> None:
    """Genişletilmiş Revir Sarf Malzemesi, İlk Yardım Çantaları, Cihaz & Stok Hareket Stüdyosu."""
    connect_db_fn = ctx_api["connect_db"]
    categories = ctx_api["INFIRMARY_INVENTORY_CATEGORIES"]
    nurse_name = ctx_api.get("nurse_name") or "Revir Sorumlusu"
    app_cfg = ctx_api["get_app_settings"]()

    with connect_db_fn() as conn:
        ensure_supply_schema(conn)

    ui.markdown("#### 🩹 Revir Sarf Malzemesi, İlk Yardım Stoku & Hazırlık Merkezi")
    summary = compute_infirmary_readiness_summary(connect_db_fn, days_horizon=30)

    # Üst Hazırlık KPI Kartları
    k1, k2, k3, k4, k5 = ui.columns(5)
    with k1:
        ui.metric("🎯 İlk Yardım Hazırlık", f"%{summary['readiness_score']}", delta=summary["readiness_tier"])
    with k2:
        ui.metric("📦 Aktif Kalem", f"{summary['active_items']}", delta=f"{summary['total_value_tl']:,.0f} TL Değer")
    with k3:
        c_cnt = len(summary["critical_items"])
        ui.metric("🚨 Kritik / Biten", f"{c_cnt} Kalem", delta=f"{len(summary['out_of_stock'])} Tükendi" if c_cnt else "Yeterli", delta_color="inverse" if c_cnt else "normal")
    with k4:
        e_cnt = len(summary["expired_items"]) + len(summary["expiring_soon_items"])
        ui.metric("⏳ SKT Uyarıları", f"{e_cnt} Kalem", delta=f"{len(summary['expired_items'])} Geçmiş" if e_cnt else "SKT Uygun", delta_color="inverse" if summary["expired_items"] else "normal")
    with k5:
        kd_cnt = len(summary["kits_attention"]) + len(summary["devices_attention"])
        ui.metric("🧰 Çanta & Cihaz", f"{len(summary['kits'])} Çanta / {len(summary['devices'])} Cihaz", delta=f"{kd_cnt} Uyarı" if kd_cnt else "Tam Hazır", delta_color="inverse" if kd_cnt else "normal")

    # Hızlı Kurulum & İndirme Çubuğu
    top_b1, top_b2, top_b3, top_b4 = ui.columns([1.4, 1.4, 1.3, 1.3])
    with top_b1:
        if ui.button("📦 MEB Standart Malzeme Kataloğu Yükle", type="secondary", use_container_width=True, key="btn_load_default_inventory"):
            added_n = ctx_api["load_default_infirmary_inventory"]()
            ui.session_state["inf_inv_editor_seed"] = ui.session_state.get("inf_inv_editor_seed", 0) + 1
            if added_n > 0:
                ui.success(f"✅ {added_n} MEB standart ilk yardım / sarf malzemesi depoya eklendi.")
            else:
                ui.info("Standart katalogdaki tüm kalemler zaten envanterde kayıtlı.")
            ui.rerun()
    with top_b2:
        if ui.button("🧰 6 Okul Çantası & 6 Cihazı Kur", type="secondary", use_container_width=True, key="btn_seed_kits_devices"):
            kn = seed_default_first_aid_kits(connect_db_fn, custodian_default=nurse_name)
            dn = seed_default_medical_devices(connect_db_fn)
            if kn > 0 or dn > 0:
                ui.success(f"✅ {kn} ilk yardım çantası ve {dn} tıbbi cihaz şablonu oluşturuldu.")
            else:
                ui.info("Varsayılan okul çantaları ve cihazlar zaten kurulu.")
            ui.rerun()
    with top_b3:
        ui.download_button(
            "🖨️ İstem & Sayım Tutanağı (.docx)",
            data=lambda: ctx_api["build_infirmary_inventory_docx"](),
            file_name=f"Revir_Sarf_Malzeme_Sayim_Istem_{date.today().isoformat()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
            key="btn_dl_inventory_docx",
        )
    with top_b4:
        ui.download_button(
            "📑 A4 Denetim Raporu (.pdf)",
            data=lambda: build_supply_a4_pdf_bytes(connect_db_fn, app_cfg),
            file_name=f"Revir_Ilk_Yardim_Denetim_A4_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            type="secondary",
            use_container_width=True,
            key="btn_dl_inventory_pdf",
        )

    # Üst Özet Uyarı Kutuları
    if summary["expired_items"] or summary["expiring_soon_items"] or summary["critical_items"]:
        ac1, ac2 = ui.columns(2)
        with ac1:
            if summary["critical_items"]:
                crit_names = ", ".join(f"{x['item_name']} ({float(x['quantity'] or 0):g} {x.get('unit') or ''})" for x in summary["critical_items"][:8])
                ui.error(f"🚨 **Kritik Stok Uyarısı ({len(summary['critical_items'])} Kalem):** {crit_names}")
            else:
                ui.success("🟢 Tüm sarf malzemeler kritik stok eşiğinin üzerinde.")
        with ac2:
            exp_msgs = []
            if summary["expired_items"]:
                exp_msgs.append("⛔ **Miadı Geçmiş:** " + ", ".join(f"{x['item_name']} ({_fmt_tr_date(x.get('expiry_date'))})" for x in summary["expired_items"][:6]))
            if summary["expiring_soon_items"]:
                exp_msgs.append("⏳ **30 Gün İçinde Dolacak:** " + ", ".join(f"{x['item_name']} ({_fmt_tr_date(x.get('expiry_date'))})" for x in summary["expiring_soon_items"][:6]))
            if exp_msgs:
                ui.warning("  \n".join(exp_msgs))
            else:
                ui.success("🟢 Miadı (SKT) yaklaşan veya geçmiş sarf malzeme yok.")
    elif summary["total_items"] > 0:
        ui.success("🟢 Tüm revir sarf malzemeleri yeterli stokta ve SKT açısından uygun.")

    # Alt Stüdyo Sekmeleri
    t_stock, t_moves, t_kits, t_exp, t_proc, t_ai = ui.tabs([
        f"📦 Ana Stok & Hızlı İşlem ({summary['total_items']})",
        "🔄 Hareketler & Fiziki Sayım",
        f"🧰 Çantalar & Cihazlar ({len(summary['kits'])}+{len(summary['devices'])})",
        f"⏳ SKT & İmha ({len(summary['expired_items']) + len(summary['expiring_soon_items']) + len(summary['quarantined_items'])})",
        "🛒 İstem, Tüketim & Çıktılar",
        "🤖 AI Stok Danışmanı",
    ])

    items_all = get_enriched_inventory(connect_db_fn)

    # 1) ANA STOK & HIZLI İŞLEM
    with t_stock:
        f1, f2, f3, f4 = ui.columns([1.8, 1.4, 1.3, 1.3])
        with f1:
            q_search = ui.text_input("🔍 Malzeme / Barkod / Lot Ara", placeholder="Örn: Gazlı bez, Batikon, Lot...", key="sup_search_q")
        with f2:
            cat_filter = ui.selectbox("🏷️ Kategori", ["Tümü", *categories], key="sup_cat_flt")
        with f3:
            st_filter = ui.selectbox("🚦 Durum", ["Tümü", "Kritik / Tükendi", "SKT Uyarısı", "Karantina / İmha", "MEB Zorunlu"], key="sup_st_flt")
        with f4:
            view_mode = ui.radio("Görünüm", ["⚡ Hızlı Kartlar", "📊 Tablo Editörü"], horizontal=True, key="sup_view_mode")

        filtered_items = []
        today_iso = date.today().isoformat()
        soon_iso = (date.today() + timedelta(days=30)).isoformat()
        for it in items_all:
            if q_search.strip():
                hay = dstudio.tr_fold(f"{it['item_name']} {it.get('barcode','')} {it.get('lot_number','')} {it.get('storage_location','')}")
                if dstudio.tr_fold(q_search) not in hay:
                    continue
            if cat_filter != "Tümü" and (_clean_text(it.get("category")) or "Diğer") != cat_filter:
                continue
            qty_f = float(it.get("quantity") or 0.0)
            thr_f = float(it.get("min_threshold") or 0.0)
            exp_s = _norm_iso_date(it.get("expiry_date"))
            q_st = _clean_text(it.get("quarantine_status")) or "Aktif"
            if st_filter == "Kritik / Tükendi" and not (qty_f <= thr_f):
                continue
            if st_filter == "SKT Uyarısı" and not (exp_s and exp_s <= soon_iso):
                continue
            if st_filter == "Karantina / İmha" and q_st == "Aktif":
                continue
            if st_filter == "MEB Zorunlu" and not int(it.get("is_mandatory_meb") or 0):
                continue
            filtered_items.append(it)

        if view_mode == "⚡ Hızlı Kartlar":
            if not filtered_items:
                ui.info("Filtreye uygun malzeme bulunamadı. Yukarıdaki '📦 MEB Standart Malzeme Kataloğu Yükle' butonunu kullanabilir veya aşağıdan yeni kalem ekleyebilirsiniz.")
            else:
                cols_cnt = 3
                rows = [filtered_items[i:i + cols_cnt] for i in range(0, len(filtered_items), cols_cnt)]
                for row_chunk in rows:
                    cols = ui.columns(cols_cnt)
                    for idx, it in enumerate(row_chunk):
                        with cols[idx]:
                            iid = int(it["id"])
                            qty_f = float(it.get("quantity") or 0.0)
                            thr_f = float(it.get("min_threshold") or 0.0)
                            tgt_f = float(it.get("target_stock") or max(10.0, thr_f * 2.0))
                            q_st = _clean_text(it.get("quarantine_status")) or "Aktif"
                            exp_s = _norm_iso_date(it.get("expiry_date"))
                            is_exp = bool(exp_s and exp_s < today_iso)
                            is_soon = bool(exp_s and today_iso <= exp_s <= soon_iso)

                            badges: list[str] = []
                            if int(it.get("is_mandatory_meb") or 0):
                                badges.append("🏛️ MEB")
                            if q_st != "Aktif":
                                badges.append(f"☣️ {q_st}")
                            elif qty_f <= 0:
                                badges.append("🔴 TÜKENDİ")
                            elif qty_f <= thr_f:
                                badges.append("🟠 Kritik")
                            else:
                                badges.append("🟢 Yeterli")
                            if is_exp:
                                badges.append("⛔ SKT Geçmiş")
                            elif is_soon:
                                badges.append("⏳ SKT ≤30g")

                            with ui.container(border=True):
                                ui.markdown(f"**{html.escape(it['item_name'])}**", unsafe_allow_html=True)
                                ui.caption(f"{' · '.join(badges)} | 📁 {it.get('category') or 'Diğer'}")
                                ui.markdown(
                                    f"<div style='font-size:18px;font-weight:800;color:#0f172a;margin:2px 0;'>"
                                    f"{qty_f:g} <span style='font-size:12px;font-weight:600;color:#64748b;'>{html.escape(it.get('unit') or 'Adet')}</span>"
                                    f" <span style='font-size:11px;font-weight:500;color:#64748b;'>(Eşik: {thr_f:g} / Hedef: {tgt_f:g})</span></div>",
                                    unsafe_allow_html=True,
                                )
                                loc_s = _clean_text(it.get("storage_location")) or "—"
                                lot_s = _clean_text(it.get("lot_number")) or "—"
                                cost_s = f"{float(it.get('unit_cost') or 0):g} TL" if float(it.get("unit_cost") or 0) > 0 else "—"
                                ui.caption(f"📍 {loc_s} · 📅 SKT: {_fmt_tr_date(exp_s)} · Lot: {lot_s} · Birim: {cost_s}")

                                b1, b2, b3 = ui.columns(3)
                                with b1:
                                    if ui.button("➖1 Kullan", key=f"sup_use1_{iid}", disabled=(qty_f <= 0 or q_st != "Aktif"), use_container_width=True):
                                        record_supply_movement(
                                            connect_db_fn,
                                            inventory_id=iid,
                                            movement_type="Sarf / Kullanım",
                                            delta_qty=-1.0,
                                            performed_by=nurse_name,
                                            reason="Karttan hızlı revir kullanımı (-1)",
                                        )
                                        ui.session_state["inf_inv_editor_seed"] = ui.session_state.get("inf_inv_editor_seed", 0) + 1
                                        ui.success(f"➖ {it['item_name']} stoktan 1 {it.get('unit') or 'Adet'} düşüldü.")
                                        ui.rerun()
                                with b2:
                                    with ui.popover("➕ Giriş", use_container_width=True):
                                        add_q = ui.number_input("Eklenecek Miktar", min_value=0.5, value=5.0, step=1.0, key=f"sup_add_q_{iid}")
                                        add_lot = ui.text_input("Parti / Lot No", value=_clean_text(it.get("lot_number")), key=f"sup_add_lot_{iid}")
                                        add_rsn = ui.text_input("Tedarik / Açıklama", value="Yeni alım / tedarik", key=f"sup_add_rsn_{iid}")
                                        if ui.button("Stoğa Ekle", type="primary", key=f"sup_add_btn_{iid}", use_container_width=True):
                                            record_supply_movement(
                                                connect_db_fn,
                                                inventory_id=iid,
                                                movement_type="Giriş / Tedarik",
                                                delta_qty=float(add_q),
                                                performed_by=nurse_name,
                                                reason=add_rsn,
                                                lot_number=add_lot or None,
                                                new_quarantine_status="Aktif",
                                            )
                                            ui.session_state["inf_inv_editor_seed"] = ui.session_state.get("inf_inv_editor_seed", 0) + 1
                                            ui.success(f"➕ {it['item_name']} +{float(add_q):g} eklendi.")
                                            ui.rerun()
                                with b3:
                                    with ui.popover("🧹 SKT/İmha", use_container_width=True):
                                        if ui.button("☣️ Karantinaya Al", key=f"sup_quar_{iid}", use_container_width=True):
                                            record_supply_movement(
                                                connect_db_fn,
                                                inventory_id=iid,
                                                movement_type="Karantina",
                                                delta_qty=0.0,
                                                performed_by=nurse_name,
                                                reason="Kullanımdan ayrıldı (Karantina)",
                                                new_quarantine_status="Karantina",
                                            )
                                            ui.rerun()
                                        if ui.button("🗑️ İmha Et (Sıfırla)", type="secondary", key=f"sup_disp_{iid}", use_container_width=True):
                                            record_supply_movement(
                                                connect_db_fn,
                                                inventory_id=iid,
                                                movement_type="SKT İmha / Zayi",
                                                delta_qty=-qty_f,
                                                performed_by=nurse_name,
                                                reason="SKT dolumu / tıbbi atık imha tutanağı",
                                                new_quarantine_status="İmha Edildi",
                                            )
                                            ui.rerun()
        else:
            if items_all:
                inv_rows = []
                for it in filtered_items or items_all:
                    st_lbl, _ = ctx_api["_inventory_row_status_label"](it)
                    inv_rows.append({
                        "id": int(it["id"]),
                        "Malzeme Adı": it["item_name"],
                        "Kategori": _clean_text(it.get("category")) or "Diğer",
                        "Miktar": float(it.get("quantity") or 0.0),
                        "Birim": _clean_text(it.get("unit")) or "Adet",
                        "Kritik Eşik": float(it.get("min_threshold") or 0.0),
                        "Hedef Stok": float(it.get("target_stock") or 10.0),
                        "Birim Maliyet (TL)": float(it.get("unit_cost") or 0.0),
                        "SKT (YYYY-AA-GG)": _clean_text(it.get("expiry_date")) or "",
                        "Parti/Lot": _clean_text(it.get("lot_number")) or "",
                        "Barkod": _clean_text(it.get("barcode")) or "",
                        "Tedarikçi": _clean_text(it.get("supplier_name")) or "",
                        "Saklama Yeri": _clean_text(it.get("storage_location")) or "",
                        "Notlar": _clean_text(it.get("notes")) or "",
                        "Durum": st_lbl,
                    })
                inv_df = pd.DataFrame(inv_rows)
                inv_seed = ui.session_state.get("inf_inv_editor_seed", 0)
                edited_inv_df = ui.data_editor(
                    inv_df,
                    key=f"inf_inv_data_editor_{inv_seed}",
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    column_config={
                        "id": None,
                        "Malzeme Adı": ui.column_config.TextColumn("Malzeme Adı", required=True),
                        "Kategori": ui.column_config.SelectboxColumn("Kategori", options=categories),
                        "Miktar": ui.column_config.NumberColumn("Miktar", min_value=0.0, step=1.0),
                        "Birim": ui.column_config.TextColumn("Birim"),
                        "Kritik Eşik": ui.column_config.NumberColumn("Kritik Eşik", min_value=0.0, step=1.0),
                        "Hedef Stok": ui.column_config.NumberColumn("Hedef Stok", min_value=0.0, step=1.0),
                        "Birim Maliyet (TL)": ui.column_config.NumberColumn("Birim (TL)", min_value=0.0, step=5.0),
                        "SKT (YYYY-AA-GG)": ui.column_config.TextColumn("SKT (YYYY-AA-GG)"),
                        "Parti/Lot": ui.column_config.TextColumn("Parti/Lot"),
                        "Barkod": ui.column_config.TextColumn("Barkod"),
                        "Tedarikçi": ui.column_config.TextColumn("Tedarikçi"),
                        "Saklama Yeri": ui.column_config.TextColumn("Saklama Yeri"),
                        "Notlar": ui.column_config.TextColumn("Notlar"),
                        "Durum": ui.column_config.TextColumn("Durum", disabled=True),
                    },
                )
                if ui.button("💾 Tablo Değişikliklerini Kaydet", type="primary", key="btn_save_inf_inv_edits", use_container_width=True):
                    saved_n = ctx_api["save_infirmary_inventory_edits"](edited_inv_df)
                    ui.session_state["inf_inv_editor_seed"] = inv_seed + 1
                    ui.success(f"✅ {saved_n} sarf malzeme kaydı güncellendi.")
                    ui.rerun()

        with ui.expander("➕ Yeni Sarf Malzeme Kalemi Ekle", expanded=False):
            with ui.form("form_add_inf_inventory_item", clear_on_submit=True):
                ic1, ic2, ic3, ic4 = ui.columns([2, 1.3, 1.0, 1.1])
                with ic1:
                    n_name = ui.text_input("Malzeme Adı*", placeholder="Örn: Steril Serum Fizyolojik", key="inv_add_name")
                with ic2:
                    n_cat = ui.selectbox("Kategori", categories, key="inv_add_cat")
                with ic3:
                    n_unit = ui.text_input("Birim", value="Adet", key="inv_add_unit")
                with ic4:
                    n_meb = ui.checkbox("MEB Zorunlu Kalem", value=True, key="inv_add_meb")

                ic5, ic6, ic7, ic8 = ui.columns([1, 1, 1, 1])
                with ic5:
                    n_qty = ui.number_input("Mevcut Miktar", min_value=0.0, step=1.0, value=10.0, key="inv_add_qty")
                with ic6:
                    n_thr = ui.number_input("Kritik Eşik", min_value=0.0, step=1.0, value=3.0, key="inv_add_thr")
                with ic7:
                    n_tgt = ui.number_input("Hedef Stok", min_value=0.0, step=1.0, value=15.0, key="inv_add_tgt")
                with ic8:
                    n_cost = ui.number_input("Birim Maliyet (TL)", min_value=0.0, step=5.0, value=0.0, key="inv_add_cost")

                ic9, ic10, ic11, ic12 = ui.columns([1.2, 1.2, 1.2, 1.2])
                with ic9:
                    n_exp_use = ui.checkbox("SKT Gir", value=False, key="inv_add_exp_chk")
                    n_exp_d = ui.date_input("Son Kullanma Tarihi", value=date.today() + timedelta(days=365), key="inv_add_exp") if n_exp_use else None
                with ic10:
                    n_lot = ui.text_input("Parti / Lot No", placeholder="LOT-2026A", key="inv_add_lot")
                with ic11:
                    n_loc = ui.text_input("Saklama Yeri", placeholder="İlk Yardım Dolabı", key="inv_add_loc")
                with ic12:
                    n_sup = ui.text_input("Tedarikçi / Barkod", placeholder="Medikal A.Ş.", key="inv_add_sup")

                n_note = ui.text_input("Not (Opsiyonel)", key="inv_add_note")
                if ui.form_submit_button("💾 Malzemeyi Envantere Ekle", type="primary", use_container_width=True):
                    if not _clean_text(n_name):
                        ui.error("Lütfen malzeme adını yazın.")
                    else:
                        new_id = ctx_api["add_infirmary_inventory_item"](
                            item_name=n_name,
                            category=n_cat,
                            quantity=float(n_qty),
                            unit=n_unit,
                            min_threshold=float(n_thr),
                            expiry_date=n_exp_d.isoformat() if n_exp_d else None,
                            storage_location=n_loc,
                            notes=n_note,
                        )
                        with connect_db_fn() as conn:
                            conn.execute(
                                "UPDATE infirmary_inventory SET lot_number=?, target_stock=?, supplier_name=?, unit_cost=?, is_mandatory_meb=? WHERE id=?",
                                (_clean_text(n_lot) or None, float(n_tgt), _clean_text(n_sup) or None, float(n_cost), int(bool(n_meb)), int(new_id)),
                            )
                        if float(n_qty) > 0:
                            record_supply_movement(
                                connect_db_fn,
                                inventory_id=int(new_id),
                                movement_type="Giriş / Tedarik",
                                delta_qty=float(n_qty),
                                performed_by=nurse_name,
                                reason="Yeni ürün açılış kaydı",
                                lot_number=n_lot or None,
                            )
                        ui.session_state["inf_inv_editor_seed"] = ui.session_state.get("inf_inv_editor_seed", 0) + 1
                        ui.success(f"'{_clean_text(n_name)}' revir envanterine eklendi.")
                        ui.rerun()

        if items_all:
            with ui.expander("🗑️ Envanterden Malzeme Sil"):
                del_opts = {f"#{int(it['id'])} — {it['item_name']} ({float(it.get('quantity') or 0):g} {it.get('unit') or ''})": int(it["id"]) for it in items_all}
                sel_del = ui.selectbox("Silinecek Malzeme", list(del_opts.keys()), key="inv_del_sel")
                if ui.button("Seçili Malzemeyi Sil", key="inv_del_btn", type="secondary"):
                    ctx_api["delete_infirmary_inventory_item"](del_opts[sel_del])
                    ui.session_state["inf_inv_editor_seed"] = ui.session_state.get("inf_inv_editor_seed", 0) + 1
                    ui.success("Malzeme envanterden silindi.")
                    ui.rerun()

    # 2) HAREKETLER & FİZİKİ SAYIM
    with t_moves:
        mc_l, mc_r = ui.columns([1.15, 1.35])
        with mc_l:
            ui.markdown("##### 📝 Manuel Stok Hareketi (Giriş / Sarf / Zayi)")
            if items_all:
                with ui.form("form_manual_supply_move", clear_on_submit=True):
                    mv_opts = {f"{it['item_name']} (Mevcut: {float(it.get('quantity') or 0):g} {it.get('unit') or ''}) [#{it['id']}]": int(it["id"]) for it in items_all}
                    sel_mv_item = ui.selectbox("Malzeme Seçin", list(mv_opts.keys()), key="mv_item_sel")
                    mv_type = ui.selectbox("İşlem Türü", list(MOVEMENT_TYPES), key="mv_type_sel")
                    mv_qty = ui.number_input("Miktar", min_value=0.5, step=1.0, value=1.0, key="mv_qty_inp")
                    mv_reason = ui.text_input("Açıklama / Gerekçe", placeholder="Örn: Spor müsabakası pansumanı, yeni fatura teslimatı...", key="mv_reason_inp")
                    if ui.form_submit_button("✅ Stok Hareketini İşle", type="primary", use_container_width=True):
                        iid_mv = mv_opts[sel_mv_item]
                        signed = float(mv_qty) if mv_type in ("Giriş / Tedarik", "Sayım Düzeltme") else (-float(mv_qty) if mv_type != "Karantina" else 0.0)
                        res_mv = record_supply_movement(
                            connect_db_fn,
                            inventory_id=iid_mv,
                            movement_type=mv_type,
                            delta_qty=signed,
                            performed_by=nurse_name,
                            reason=mv_reason or mv_type,
                        )
                        ui.success(f"✅ {res_mv['item_name']} güncellendi (Yeni bakiye: {res_mv['new_qty']:g}).")
                        ui.rerun()
            else:
                ui.info("Hareket eklemek için önce malzeme kataloğunu yükleyin veya yeni kalem ekleyin.")

        with mc_r:
            ui.markdown("##### 📊 Son Stok Hareketleri Defteri")
            mv_flt = ui.selectbox("İşlem Filtresi", ["Tümü", *MOVEMENT_TYPES], key="mv_hist_flt")
            moves = get_supply_movements(connect_db_fn, limit=120, movement_type=mv_flt)
            if moves:
                df_m = pd.DataFrame([
                    {
                        "Tarih": _fmt_tr_date(m.get("movement_date")),
                        "Malzeme": m.get("item_name"),
                        "İşlem": m.get("movement_type"),
                        "Değişim": f"{'+' if float(m.get('delta_qty') or 0) > 0 else ''}{float(m.get('delta_qty') or 0):g}",
                        "Bakiye": f"{float(m.get('balance_after') or 0):g} {m.get('unit') or ''}",
                        "Öğrenci / Personel": _clean_text(m.get("student_name")) or "—",
                        "Yapan": _clean_text(m.get("performed_by")) or "—",
                        "Açıklama": _clean_text(m.get("reason")) or "—",
                    }
                    for m in moves
                ])
                ui.dataframe(df_m, use_container_width=True, hide_index=True, height=320)
            else:
                ui.info("Henüz kayıtlı stok hareketi bulunmuyor.")

        ui.divider()
        ui.markdown("##### 🧮 Toplu Fiziki Stok Sayımı & Fark Eşitleme")
        ui.caption("Dolap ve rafları saydıktan sonra 'Sayılan Miktar' sütununu güncelleyin; sistem ile fark varsa tek tuşla 'Sayım Düzeltme' fişi oluşturulur.")
        if items_all:
            df_cnt = pd.DataFrame([
                {
                    "id": int(it["id"]),
                    "Malzeme Adı": it["item_name"],
                    "Saklama Yeri": _clean_text(it.get("storage_location")) or "—",
                    "Sistem Miktarı": float(it.get("quantity") or 0.0),
                    "Sayılan Miktar": float(it.get("quantity") or 0.0),
                    "Birim": _clean_text(it.get("unit")) or "Adet",
                }
                for it in items_all
            ])
            cnt_ed = ui.data_editor(
                df_cnt,
                key="sup_phys_count_ed",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "id": None,
                    "Malzeme Adı": ui.column_config.TextColumn("Malzeme Adı", disabled=True),
                    "Saklama Yeri": ui.column_config.TextColumn("Konum", disabled=True),
                    "Sistem Miktarı": ui.column_config.NumberColumn("Sistem", disabled=True),
                    "Sayılan Miktar": ui.column_config.NumberColumn("Sayılan", min_value=0.0, step=1.0),
                    "Birim": ui.column_config.TextColumn("Birim", disabled=True),
                },
            )
            if ui.button("🧮 Fiziki Sayım Farklarını Stoka İşle", type="primary", key="btn_apply_phys_count", use_container_width=True):
                res_cnt = apply_physical_inventory_count(connect_db_fn, cnt_ed, performed_by=nurse_name)
                ui.success(f"✅ Sayım tamamlandı: {res_cnt['adjusted']} kalemde stok farkı düzeltildi, {res_cnt['matched']} kalem birebir tuttu.")
                ui.rerun()

    # 3) ÇANTALAR, DOLAP DENETİMİ & TIBBİ CİHAZLAR
    with t_kits:
        kits = summary["kits"]
        if kits:
            kit_opts = {
                f"{'⚠️' if k['needs_attention'] else '🟢'} {k['kit_name']} ({k.get('location') or '—'}) [#{k['id']}]": k
                for k in kits
            }
            sel_k_lbl = ui.selectbox("Denetlenecek İlk Yardım Çantası / Acil Kit Seçin", list(kit_opts.keys()), key="sup_kit_sel")
            sel_kit = kit_opts[sel_k_lbl]
            kid_cur = int(sel_kit["id"])

            ui.info(
                f"📍 **Konum:** {_clean_text(sel_kit.get('location')) or '—'} · "
                f"👤 **Zimmetli:** {_clean_text(sel_kit.get('custodian_name')) or '—'} · "
                f"📅 **Son Kontrol:** {_fmt_tr_date(sel_kit.get('last_checked_date'))} · "
                f"🚦 **Durum:** {sel_kit.get('last_status')}"
            )

            k_items_df = pd.DataFrame([
                {
                    "id": int(ki["id"]),
                    "Malzeme": ki["item_name"],
                    "Gerekli": float(ki.get("required_qty") or 1.0),
                    "Mevcut": float(ki.get("current_qty") or 0.0),
                    "Birim": _clean_text(ki.get("unit")) or "Adet",
                    "SKT (YYYY-AA-GG)": _clean_text(ki.get("expiry_date")) or "",
                    "Durum": _clean_text(ki.get("status")) or "Tamam",
                    "Not": _clean_text(ki.get("notes")) or "",
                }
                for ki in (sel_kit.get("items") or [])
            ])
            edited_ki_df = ui.data_editor(
                k_items_df,
                key=f"sup_kit_ed_{kid_cur}",
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "id": None,
                    "Malzeme": ui.column_config.TextColumn("Malzeme", required=True),
                    "Gerekli": ui.column_config.NumberColumn("Gerekli", min_value=0.0, step=1.0),
                    "Mevcut": ui.column_config.NumberColumn("Mevcut", min_value=0.0, step=1.0),
                    "Birim": ui.column_config.TextColumn("Birim"),
                    "SKT (YYYY-AA-GG)": ui.column_config.TextColumn("SKT"),
                    "Durum": ui.column_config.SelectboxColumn("Durum", options=["Tamam", "Eksik", "SKT Geçmiş"]),
                    "Not": ui.column_config.TextColumn("Not"),
                },
            )
            insp_note = ui.text_input("Denetim Notu", placeholder="Örn: Aylık periyodik kontrol yapıldı, mühür yenilendi...", key=f"sup_kit_note_{kid_cur}")
            kb1, kb2 = ui.columns(2)
            with kb1:
                if ui.button("💾 Çanta Kontrolünü Kaydet", type="secondary", key=f"btn_save_kit_{kid_cur}", use_container_width=True):
                    res_k = save_kit_items_and_inspect(
                        connect_db_fn,
                        kit_id=kid_cur,
                        edited_items_df=edited_ki_df,
                        checked_by=nurse_name,
                        inspection_notes=insp_note,
                        replenish_from_depot=False,
                    )
                    ui.success(f"✅ {res_k['kit_name']} kontrolü kaydedildi ({res_k['overall_status']}).")
                    ui.rerun()
            with kb2:
                if ui.button("🔄 Ana Depodan Eksikleri Tamamla & Onayla", type="primary", key=f"btn_replenish_kit_{kid_cur}", use_container_width=True):
                    res_r = save_kit_items_and_inspect(
                        connect_db_fn,
                        kit_id=kid_cur,
                        edited_items_df=edited_ki_df,
                        checked_by=nurse_name,
                        inspection_notes=insp_note,
                        replenish_from_depot=True,
                    )
                    ui.success(
                        f"✅ {res_r['kit_name']} güncellendi · Ana depodan {res_r['replenished_count']} kalem ikmal edildi "
                        f"({', '.join(res_r['replenished_items']) or 'İkmal gerekmedi'})."
                    )
                    ui.rerun()
        else:
            ui.info("Henüz okul ilk yardım çantası veya acil kit oluşturulmamış. Üstteki '🧰 6 Okul Çantası & 6 Cihazı Kur' butonuna basın.")

        with ui.expander("➕ Yeni Özel İlk Yardım Çantası / Kat Dolabı Ekle"):
            with ui.form("form_new_first_aid_kit", clear_on_submit=True):
                nk1, nk2, nk3, nk4 = ui.columns(4)
                with nk1:
                    nk_nm = ui.text_input("Çanta / Dolap Adı*", placeholder="2. Kat İlk Yardım Dolabı", key="nk_name")
                with nk2:
                    nk_tp = ui.selectbox("Tür", KIT_TYPES, key="nk_type")
                with nk3:
                    nk_loc = ui.text_input("Konum", placeholder="2. Kat Koridoru", key="nk_loc")
                with nk4:
                    nk_cust = ui.text_input("Zimmetli / Sorumlu", value=nurse_name, key="nk_cust")
                if ui.form_submit_button("🧰 Çantayı Oluştur", type="primary", use_container_width=True):
                    if _clean_text(nk_nm):
                        create_first_aid_kit(
                            connect_db_fn,
                            kit_name=nk_nm,
                            kit_type=nk_tp,
                            location=nk_loc,
                            custodian_name=nk_cust,
                            starter_items=[("Steril Gazlı Bez (Spanç)", 5, "Paket"), ("Batikon (Povidon İyot Antiseptik)", 1, "Şişe"), ("Yara Bandı", 1, "Kutu")],
                        )
                        ui.success(f"✅ '{_clean_text(nk_nm)}' oluşturuldu.")
                        ui.rerun()

        ui.divider()
        ui.markdown("##### 🩺 Tıbbi Cihazlar, Kalibrasyon, Pil & AED Ped Takibi")
        devs = summary["devices"]
        if devs:
            df_dev_ui = pd.DataFrame([
                {
                    "ID": int(d["id"]),
                    "Cihaz Adı": d["device_name"],
                    "Tür": d.get("device_type"),
                    "Konum": _clean_text(d.get("location")) or "—",
                    "Pil": _clean_text(d.get("battery_status")) or "İyi",
                    "Sonraki Kalibrasyon": _fmt_tr_date(d.get("next_calibration_date")),
                    "Ped/Sensör SKT": _fmt_tr_date(d.get("pad_or_sensor_expiry")),
                    "Durum": d.get("status"),
                    "Uyarı": ", ".join(d.get("alert_flags") or []) or "🟢 Hazır",
                }
                for d in devs
            ])
            ui.dataframe(df_dev_ui, use_container_width=True, hide_index=True)

        with ui.expander("➕ Yeni Tıbbi Cihaz Ekle / Kalibrasyon Kaydet"):
            with ui.form("form_add_med_device", clear_on_submit=True):
                dc1, dc2, dc3, dc4 = ui.columns(4)
                with dc1:
                    d_nm = ui.text_input("Cihaz Adı*", placeholder="Örn: OED Defibrilatör", key="dev_nm")
                with dc2:
                    d_tp = ui.selectbox("Cihaz Türü", DEVICE_TYPES, key="dev_tp")
                with dc3:
                    d_loc = ui.text_input("Konum", placeholder="Revir Odası", key="dev_loc")
                with dc4:
                    d_bat = ui.selectbox("Pil / Güç Durumu", ["İyi", "Zayıf / Değişmeli", "Adaptör / Şebeke"], key="dev_bat")
                dc5, dc6, dc7 = ui.columns(3)
                with dc5:
                    d_ncal = ui.date_input("Sonraki Kalibrasyon Tarihi", value=date.today() + timedelta(days=365), key="dev_ncal")
                with dc6:
                    d_has_pad = ui.checkbox("Ped/Sensör SKT Gir (AED vb.)", value=False, key="dev_has_pad")
                    d_pad = ui.date_input("Ped/Sensör SKT", value=date.today() + timedelta(days=365), key="dev_pad") if d_has_pad else None
                with dc7:
                    d_st = ui.selectbox("Çalışma Durumu", ["Çalışıyor", "Bakımda / Arızalı", "Kalibrasyon Bekliyor"], key="dev_st")
                if ui.form_submit_button("💾 Cihazı Kaydet", type="primary", use_container_width=True):
                    if _clean_text(d_nm):
                        upsert_medical_device(
                            connect_db_fn,
                            device_name=d_nm,
                            device_type=d_tp,
                            location=d_loc,
                            battery_status=d_bat,
                            last_calibration_date=date.today().isoformat(),
                            next_calibration_date=d_ncal.isoformat(),
                            pad_or_sensor_expiry=d_pad.isoformat() if d_pad else None,
                            status=d_st,
                        )
                        ui.success(f"✅ '{_clean_text(d_nm)}' cihaz listesine kaydedildi.")
                        ui.rerun()

    # 4) SKT TAKVİMİ & İMHA YÖNETİMİ
    with t_exp:
        buckets = get_expiry_buckets(items_all)
        eb1, eb2 = ui.columns(2)
        with eb1:
            ui.markdown("##### ⛔ SKT Geçmiş (Acil Ayırma / İmha)")
            if buckets["expired"]:
                for x in buckets["expired"]:
                    c_a, c_b = ui.columns([2.2, 1.0])
                    with c_a:
                        ui.error(f"**{x['item_name']}** — {float(x.get('quantity') or 0):g} {x.get('unit') or ''} (SKT: {_fmt_tr_date(x.get('expiry_date'))})")
                    with c_b:
                        if ui.button("🗑️ İmha Et", key=f"exp_disp_{x['id']}", use_container_width=True):
                            record_supply_movement(
                                connect_db_fn,
                                inventory_id=int(x["id"]),
                                movement_type="SKT İmha / Zayi",
                                delta_qty=-float(x.get("quantity") or 0),
                                performed_by=nurse_name,
                                reason=f"SKT dolduğu için tıbbi atık imha ({_fmt_tr_date(x.get('expiry_date'))})",
                                new_quarantine_status="İmha Edildi",
                            )
                            ui.rerun()
            else:
                ui.success("🟢 SKT'si geçmiş ürün yok.")

            ui.markdown("##### 🟠 31 – 90 Gün İçinde Dolacaklar")
            if buckets["expiring_90"]:
                for x in buckets["expiring_90"]:
                    ui.info(f"📅 **{x['item_name']}** — {float(x.get('quantity') or 0):g} {x.get('unit') or ''} (SKT: {_fmt_tr_date(x.get('expiry_date'))})")
            else:
                ui.caption("31–90 gün penceresinde yaklaşan SKT yok.")

        with eb2:
            ui.markdown("##### 🔴 0 – 30 Gün İçinde Dolacaklar (Öncelikli Kullanım - FEFO)")
            if buckets["expiring_30"]:
                for x in buckets["expiring_30"]:
                    ui.warning(f"⏳ **{x['item_name']}** — {float(x.get('quantity') or 0):g} {x.get('unit') or ''} (SKT: {_fmt_tr_date(x.get('expiry_date'))})")
            else:
                ui.success("🟢 30 gün içinde dolacak ürün yok.")

            ui.markdown("##### ☣️ Karantina / İmha Edilmiş Kayıtlar")
            if buckets["quarantined"]:
                for x in buckets["quarantined"]:
                    ui.caption(f"• **{x['item_name']}** — Statü: `{x.get('quarantine_status')}` · Kalan: {float(x.get('quantity') or 0):g}")
            else:
                ui.caption("Karantina veya imha kaydı yok.")

    # 5) SATIN ALMA İSTEM PLANI, TÜKETİM TRENDİ & ÇIKTILAR
    with t_proc:
        proc = build_procurement_plan(connect_db_fn)
        cons = get_consumption_summary(connect_db_fn, days=30)

        pc1, pc2, pc3 = ui.columns(3)
        with pc1:
            ui.metric("🛒 İstem Gereken Kalem", f"{proc['item_count']} Kalem")
        with pc2:
            ui.metric("💳 Tahmini Alım Bütçesi", f"{proc['total_budget_tl']:,.2f} TL")
        with pc3:
            total_cons_ops = sum(c["move_count"] for c in cons)
            ui.metric("📉 30 Günlük Tüketim İşlemi", f"{total_cons_ops} Hareket")

        if proc["rows"]:
            ui.markdown("##### 🛒 Otomatik İkmal & Satın Alma Sepeti")
            ui.dataframe(pd.DataFrame(proc["rows"]).drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
        else:
            ui.success("🟢 Tüm ana stok ve ilk yardım çantaları hedef stok seviyesindedir; satın alma gerekmiyor.")

        if cons:
            ui.markdown("##### 📊 Son 30 Gün En Çok Kullanılan Malzemeler & Tükenme Tahmini")
            df_cons = pd.DataFrame([
                {
                    "Malzeme": c["item_name"],
                    "30G Tüketim": f"{c['consumed_qty']:g} {c['unit']}",
                    "Mevcut": f"{c['current_qty']:g} {c['unit']}",
                    "Tahmini Yetecek Gün": f"~{c['estimated_days_left']} gün" if c["estimated_days_left"] is not None else "—",
                }
                for c in cons
            ])
            ui.dataframe(df_cons, use_container_width=True, hide_index=True)

        ex1, ex2, ex3 = ui.columns(3)
        with ex1:
            ui.download_button(
                "📊 Tüm Envanter & Çantalar (.xlsx)",
                data=lambda: build_supply_excel_bytes(connect_db_fn),
                file_name=f"Revir_Sarf_Canta_Envanter_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True,
                key="btn_dl_supply_xlsx",
            )
        with ex2:
            ui.download_button(
                "📄 Resmi İstem & Sayım Tutanağı (.docx)",
                data=lambda: ctx_api["build_infirmary_inventory_docx"](),
                file_name=f"Revir_Resmi_Istem_Sayim_{date.today().isoformat()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="secondary",
                use_container_width=True,
                key="btn_dl_supply_docx2",
            )
        with ex3:
            ui.download_button(
                "📑 İmzaya Hazır A4 Rapor (.pdf)",
                data=lambda: build_supply_a4_pdf_bytes(connect_db_fn, app_cfg),
                file_name=f"Revir_Ilk_Yardim_A4_{date.today().isoformat()}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True,
                key="btn_dl_supply_pdf2",
            )

    # 6) GEMINI AI STOK DANIŞMANI
    with t_ai:
        ctx_api["render_gemini_inventory_ai_panel"](
            {
                "critical_stock": summary["critical_items"],
                "expired": summary["expired_items"],
                "expiring_soon": summary["expiring_soon_items"],
            },
            summary["total_items"],
            extra_context={
                "readiness_score": summary["readiness_score"],
                "kits_attention": [
                    {"canta": k["kit_name"], "eksik": k.get("missing_count"), "skt": k.get("expired_count")}
                    for k in summary["kits_attention"]
                ],
                "top_consumed_30d": get_consumption_summary(connect_db_fn, days=30)[:10],
                "estimated_budget_tl": build_procurement_plan(connect_db_fn)["total_budget_tl"],
            },
        )
