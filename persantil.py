from __future__ import annotations
import mimetypes

import copy
import sys
import gc
import warnings
warnings.filterwarnings('ignore', message='.*use_container_width.*')
import html
import json
import math
import re
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timedelta, time
from hashlib import sha256
import io
from io import BytesIO
from pathlib import Path
from typing import Iterator

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
try:
	import docx
	from docx.shared import Inches, Pt, RGBColor, Cm
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
	from docx.enum.section import WD_ORIENT
	from docx.oxml import parse_xml
	from docx.oxml.ns import nsdecls
	HAS_DOCX = True
except ImportError:
	docx = None
	Inches = Pt = RGBColor = Cm = None
	WD_ALIGN_PARAGRAPH = WD_TABLE_ALIGNMENT = WD_ALIGN_VERTICAL = WD_ORIENT = None
	parse_xml = nsdecls = None
	HAS_DOCX = False


APP_DIR = Path(__file__).resolve().parent
APP_VERSION = "1.0.0"
APP_ICON_PATH = APP_DIR / "app_icon.png"
APP_ICON_ICO = APP_DIR / "app_icon.ico"
DB_PATH = APP_DIR / "revir_defteri.db"
PERSON_DATA_DIR = APP_DIR / "kisi_verileri"
PERSANTIL_DATA_DIR = APP_DIR / "persantil veriler" 
TR_LOCATIONS_PATH = APP_DIR / "turkiye-il-ilce-mahalle-koy.json"
EVRAKLAR_DIR = APP_DIR / "revir_evraklari"
EGITIM_YILLARI_DIR = APP_DIR / "egitim_yillari"
EVRAKLAR_DIR.mkdir(parents=True, exist_ok=True)
EGITIM_YILLARI_DIR.mkdir(parents=True, exist_ok=True)

PAGES = [
	"📊 Genel Bakış",
	"👥 Öğrenciler",
	"🧑‍💼 Yetişkinler",
	"🏫 Okul Taraması",
	"🩺 Revir Kaydı",
	"💊 İlaç Takip Sistemi",
	"💓 Vital Takip",
	"📋 Sağlık Profili ve Bakım Planı",
	"📏 Ölçüm ve Gelişim",
	"🖨️ Raporlar",
	"📁 Evraklar",
	"⚙️ Ayarlar",
]


@contextmanager
def connect_master_db() -> Iterator[sqlite3.Connection]:
	connection = sqlite3.connect(DB_PATH, timeout=15)
	connection.row_factory = sqlite3.Row
	connection.execute("PRAGMA foreign_keys = ON")
	try:
		yield connection
		connection.commit()
	except Exception:
		connection.rollback()
		raise
	finally:
		connection.close()


def get_current_db_path() -> Path:
	try:
		ay = get_active_academic_year()
		folder = ay.get("folder_name", "")
		if folder:
			year_dir = APP_DIR / folder
			year_dir.mkdir(parents=True, exist_ok=True)
			year_db = year_dir / "revir_defteri.db"
			if year_db.exists():
				return year_db
			if DB_PATH.exists():
				import shutil
				shutil.copy2(DB_PATH, year_db)
				return year_db
			return year_db
	except Exception:
		pass
	return DB_PATH


@contextmanager
def connect_db() -> Iterator[sqlite3.Connection]:
	connection = sqlite3.connect(get_current_db_path(), timeout=15)
	connection.row_factory = sqlite3.Row
	connection.execute("PRAGMA foreign_keys = ON")
	try:
		yield connection
		connection.commit()
	except Exception:
		connection.rollback()
		raise
	finally:
		connection.close()


DEFAULT_APP_SETTINGS: dict[str, str] = {
	"font_family": "DM Sans",
	"font_size": "Normal (15px)",
	"ui_density": "Standart",
	"theme_name": "Klinik Turkuaz (Varsayılan)",
	"dark_mode_adaptation": "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)",
	"language": "Türkçe (TR)",
	"default_country": "Türkiye",
	"default_province": "Ankara",
	"default_district": "Merkez",
	"default_village": "",
	"institution_top_header": "T.C. MİLLÎ EĞİTİM BAKANLIĞI",
	"school_name": "Okul / Kurum Reviri",
	"nurse_name": "Okul Revir Sorumlusu",
	"principal_name": "Okul Müdürü",
	"institution_phone": "",
	"show_address_in_summary": "1",
	"default_table_height": "420",
	"custom_student_symptoms": "Baş ağrısı, Karın ağrısı, Mide bulantısı / kusma, Ateş / halsizlik, Boğaz ağrısı, Düşme / çarpma, Burun kanaması, Yara / sıyrık, Göz batması / sulanması, Diş ağrısı, Alerjik kaşıntı / döküntü, Baş dönmesi",
	"custom_student_actions": "Buz / soğuk kompres uygulandı, Pansuman ve yara bandı yapıldı, Ateş ölçüldü ve revirde dinlendirildi, Burun/ağız tamponu yapıldı, Veli arandı ve bilgilendirildi, Veliye teslim edildi / eve gönderildi, Sağlık kuruluşuna / 112'ye yönlendirildi, Dinlendirildi ve derse döndü",
	"custom_civilian_symptoms": "Tansiyon yüksekliği / düşüklüğü, Baş dönmesi / bayılma hissi, Baş ağrısı / migren, Göğüs sıkışması / çarpıntı, Halsizlik / yorgunluk, Bel / boyun / kas spazmı, Mide bulantısı / kramp, Şeker düşüklüğü / titreme, Düşme / burkulma / travma, Kesik / yara, Alerjik reaksiyon",
	"custom_civilian_actions": "Tansiyon ve nabız ölçüldü, Kan şekeri ölçümü yapıldı, Revirde istirahat ettirildi, Buz / soğuk kompres uygulandı, Pansuman ve yara bakımı yapıldı, Ailesine / yakınına bilgi verildi, 112 Acil Servis çağrıldı, Hastaneye / acile sevk edildi, Şikayeti hafifledi ve görevine döndü, İlaç desteği / kontrol yapıldı",
	"custom_care_plan_templates": "",
	"custom_cp_diagnoses": "",
	"custom_cp_facts": "",
	"custom_cp_goals": "",
	"custom_cp_interventions": "",
	"custom_cp_monitorings": "",
	"custom_cp_escalations": "",
	"screening_visible_fields": "height_cm,weight_kg,head_circ_cm,temperature_c,pulse_bpm,spo2_pct,wears_glasses,vision_right,vision_left,hair_screening,hearing_right,hearing_left,health_behavior,harmful_habits,psychosocial_status",
	"custom_class_templates": "{}",
	"medication_visible_fields": "med_route,dosage_pattern,dose_unit,meal_relation,total_volume,duration_days,active_days,storage_condition,guardian_consent,expiry_date,special_notes",
}

COUNTRY_OPTIONS = [
	"Türkiye",
	"KKTC (Kuzey Kıbrıs Türk Cumhuriyeti)",
	"Azerbaycan",
	"Almanya",
	"Fransa",
	"Hollanda",
	"İngiltere (Birleşik Krallık)",
	"ABD",
	"Suriye",
	"Irak",
	"Diğer",
]

LANGUAGE_OPTIONS = [
	"Türkçe (TR)",
	"English (EN)",
	"Deutsch (DE)",
	"Français (FR)",
	"العربية (AR)",
]

DARK_MODE_OPTIONS = [
	"Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)",
	"Kapalı (Dark Mode'a Adapte Olma - Sabit Aydınlık)",
	"Açık (Daima Karanlık Mod - Yüksek Kontrastlı Koyu)",
]


def is_dark_mode_active(app_cfg: dict | None = None) -> bool:
	if app_cfg is None:
		app_cfg = get_app_settings()
	theme_name = app_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
	dm_mode = app_cfg.get("dark_mode_adaptation", "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)")
	if dm_mode == "Açık (Daima Karanlık Mod - Yüksek Kontrastlı Koyu)":
		return True
	if dm_mode == "Kapalı (Dark Mode'a Adapte Olma - Sabit Aydınlık)":
		return False
	return "Dark" in theme_name or "Koyu" in theme_name


def build_app_css(font_css: str, font_url: str, base_fsize: str, pad_top: str, theme: dict[str, str], theme_name: str, dark_mode_mode: str) -> str:
	is_theme_dark = "Dark" in theme_name or "Koyu" in theme_name
	is_forced_dark = (dark_mode_mode == "Açık (Daima Karanlık Mod - Yüksek Kontrastlı Koyu)") or is_theme_dark
	is_forced_light = (dark_mode_mode == "Kapalı (Dark Mode'a Adapte Olma - Sabit Aydınlık)") and not is_theme_dark

	base_rules = f"""
	@import url('{font_url}');
	:root {{
		--ink: {theme['ink']};
		--teal: {theme['teal']};
		--teal-hover: {theme['teal_hover']};
		--mint: {theme['mint']};
		--coral: #d76c4c;
		--line: {theme['line']};
		--card-bg: {theme['card_bg']};
		--sidebar-bg: {theme['sidebar']};
		--bg-page: {theme['bg_grad']};
		--input-bg: #ffffff;
		--input-text: #0f172a;
		--text-main: {theme['ink']};
		--text-heading: {theme['ink']};
		--text-muted: #475569;
	}}
	html, body, [class*="css"] {{ font-family: {font_css}; color: var(--text-main); font-size: {base_fsize}; }}
	h1, h2, h3, h4, h5, h6 {{ font-family: {font_css}; letter-spacing: 0; color: var(--text-heading); font-weight: 700; }}
	.stApp {{ background: var(--bg-page); }}
	.block-container {{ padding-top: {pad_top} !important; padding-bottom: 2rem !important; }}
	[data-testid="stSidebar"] {{ background: var(--sidebar-bg); border-right: 1px solid var(--line); }}
	[data-testid="stSidebar"] * {{ color: var(--text-main); }}
	[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{ color: var(--text-muted) !important; }}
	[data-testid="stMetric"] {{ background: var(--card-bg); border: 1px solid var(--line); border-radius: 8px; padding: 14px 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
	[data-testid="stMetricLabel"] p {{ color: var(--text-muted) !important; font-size: 13.5px; font-weight: 600; }}
	[data-testid="stMetricValue"] div {{ color: var(--text-heading) !important; font-weight: 700; }}
	[data-testid="stVerticalBlockBorderWrapper"] > div {{ background: var(--card-bg); border: 1px solid var(--line) !important; border-radius: 8px; }}
	.stButton > button, .stDownloadButton > button {{ border-radius: 6px; border: 1px solid var(--teal); color: var(--teal); background: var(--card-bg); font-weight: 600; transition: all 0.15s ease-in-out; }}
	.stButton > button:hover, .stDownloadButton > button:hover {{ background: var(--mint); border-color: var(--teal); }}
	.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{ background: var(--teal); color: white; border: 1px solid var(--teal); }}
	.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {{ background: var(--teal-hover); }}
	[data-testid="stDataFrame"] {{ border: 1px solid var(--line); border-radius: 6px; background: var(--card-bg); }}
	div[data-baseweb="input"] input, div[data-baseweb="base-input"] input, textarea {{ border-radius: 6px !important; background-color: var(--input-bg); color: var(--input-text); }}
	div[data-baseweb="select"] > div {{ border-radius: 6px !important; background-color: var(--input-bg); color: var(--input-text); }}
	div[data-baseweb="select"] ul li svg {{ display: none !important; }}
	div[data-baseweb="select"] ul li[aria-selected="true"]::before {{ content: "📂 "; font-size: 14px; }}

	/* Streamlit Header, Menü, Deploy, GitHub Profil Badge ve Footer Gizleme */
	#MainMenu, 
	button[data-testid="stMainMenu"], 
	[data-testid="stMainMenu"],
	.stDeployButton, 
	[data-testid="stDeployButton"], 
	button[data-testid="stDeployButton"],
	div[data-testid="stToolbarActions"],
	[data-testid="stToolbar"],
	header [data-testid="stToolbar"],
	header [data-testid="stHeaderActionElements"],
	[data-testid="stStatusWidget"],
	[data-testid="manage-app-button"],
	div[data-testid="stDecoration"],
	footer, 
	[data-testid="stFooter"],
	footer:after, 
	footer:before,
	div[class*="viewerBadge"], 
	a[class*="viewerBadge"], 
	div[class*="ProfileBadge"],
	a[class*="ProfileBadge"],
	img[class*="ProfileBadge"],
	a[href*="streamlit.io/cloud"],
	a[href*="streamlit.io"],
	header img,
	header a[href*="github.com"] {{
		display: none !important;
		visibility: hidden !important;
		height: 0 !important;
		width: 0 !important;
		opacity: 0 !important;
		pointer-events: none !important;
	}}
	header[data-testid="stHeader"] {{
		background: transparent !important;
	}}

	/* Sol Alt Köşe - Menüyü Aç Düğmesi Tasarımı */
	#fab-menu-toggle,
	div[data-testid="collapsedControl"],
	[data-testid="collapsedControl"],
	[data-testid="stSidebarCollapsedControl"] {{
		position: fixed !important;
		bottom: 24px !important;
		left: 24px !important;
		top: auto !important;
		z-index: 999999 !important;
		display: flex !important;
		align-items: center !important;
		justify-content: center !important;
		gap: 8px !important;
		background: #0d9488 !important;
		color: #ffffff !important;
		padding: 10px 20px !important;
		border-radius: 50px !important;
		box-shadow: 0 4px 18px rgba(13, 148, 136, 0.45) !important;
		border: 2px solid rgba(255, 255, 255, 0.35) !important;
		cursor: pointer !important;
		font-family: inherit !important;
		font-size: 14.5px !important;
		font-weight: 700 !important;
		letter-spacing: 0.3px !important;
		transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
		visibility: visible !important;
		opacity: 1 !important;
		user-select: none !important;
		text-decoration: none !important;
	}}

	#fab-menu-toggle:hover,
	div[data-testid="collapsedControl"]:hover,
	[data-testid="collapsedControl"]:hover,
	[data-testid="stSidebarCollapsedControl"]:hover {{
		background: #0f766e !important;
		transform: translateY(-2px) scale(1.05) !important;
		box-shadow: 0 6px 24px rgba(13, 148, 136, 0.6) !important;
	}}

	#fab-menu-toggle:active,
	div[data-testid="collapsedControl"]:active,
	[data-testid="collapsedControl"]:active {{
		transform: translateY(0) scale(0.98) !important;
	}}

	div[data-testid="collapsedControl"] button,
	[data-testid="collapsedControl"] button,
	[data-testid="stSidebarCollapsedControl"] button {{
		background: transparent !important;
		border: none !important;
		color: #ffffff !important;
		padding: 0 !important;
		margin: 0 !important;
		display: flex !important;
		align-items: center !important;
		gap: 8px !important;
		font-size: 14.5px !important;
		font-weight: 700 !important;
		cursor: pointer !important;
		box-shadow: none !important;
	}}

	div[data-testid="collapsedControl"] button:hover,
	[data-testid="collapsedControl"] button:hover,
	[data-testid="stSidebarCollapsedControl"] button:hover {{
		background: transparent !important;
		border: none !important;
	}}

	div[data-testid="collapsedControl"] svg,
	[data-testid="collapsedControl"] svg,
	[data-testid="stSidebarCollapsedControl"] svg {{
		color: #ffffff !important;
		fill: #ffffff !important;
		stroke: #ffffff !important;
		width: 18px !important;
		height: 18px !important;
	}}

	div[data-testid="collapsedControl"] button::after,
	[data-testid="collapsedControl"] button::after,
	[data-testid="stSidebarCollapsedControl"] button::after {{
		content: "Menüyü Aç";
		color: #ffffff !important;
		font-weight: 700 !important;
		font-size: 14.5px !important;
		letter-spacing: 0.3px !important;
		white-space: nowrap !important;
	}}

	/* Sidebar açıkken floating menü açma butonunun gizlenmesi */
	section[data-testid="stSidebar"]:not([aria-expanded="false"]) ~ * #fab-menu-toggle,
	.stApp:has(section[data-testid="stSidebar"]:not([aria-expanded="false"]):not([style*="display: none"])) #fab-menu-toggle {{
		display: none !important;
	}}
	"""

	dark_rules = """
	:root {
		color-scheme: dark !important;
		--ink: #f8fafc !important;
		--teal: #2dd4bf !important;
		--teal-hover: #14b8a6 !important;
		--mint: #1e293b !important;
		--coral: #fb7185 !important;
		--line: #334155 !important;
		--card-bg: #1e293b !important;
		--sidebar-bg: #090d16 !important;
		--bg-page: linear-gradient(155deg, #090d16 0%, #0f172a 50%, #1e293b 100%) !important;
		--input-bg: #1e293b !important;
		--input-text: #f8fafc !important;
		--text-main: #f8fafc !important;
		--text-heading: #ffffff !important;
		--text-muted: #94a3b8 !important;
	}
	.stApp {
		background: var(--bg-page) !important;
		color: var(--text-main) !important;
	}
	html, body, [class*="css"], .stMarkdown, .stMarkdown p, p, span, label, div {
		color: var(--text-main);
	}
	h1, h2, h3, h4, h5, h6 {
		color: var(--text-heading) !important;
		font-weight: 700 !important;
	}
	.stCaption, caption, [data-testid="stCaptionContainer"] p, small {
		color: var(--text-muted) !important;
	}
	[data-testid="stSidebar"] {
		background: var(--sidebar-bg) !important;
		border-right: 1px solid var(--line) !important;
	}
	[data-testid="stSidebar"] * {
		color: var(--text-main) !important;
	}
	[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
		color: var(--text-muted) !important;
	}
	[data-testid="stVerticalBlockBorderWrapper"] > div {
		background: var(--card-bg) !important;
		border: 1px solid var(--line) !important;
	}
	[data-testid="stMetric"] {
		background: var(--card-bg) !important;
		border: 1px solid var(--line) !important;
	}
	[data-testid="stMetricLabel"] p {
		color: var(--text-muted) !important;
	}
	[data-testid="stMetricValue"] div {
		color: var(--text-heading) !important;
	}
	div[data-baseweb="input"] input, div[data-baseweb="base-input"] input, textarea {
		background-color: var(--input-bg) !important;
		color: var(--input-text) !important;
		border-color: var(--line) !important;
	}
	div[data-baseweb="select"] > div {
		background-color: var(--input-bg) !important;
		color: var(--input-text) !important;
		border-color: var(--line) !important;
	}
	div[data-baseweb="select"] * {
		color: var(--input-text) !important;
	}
	div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {
		background-color: var(--card-bg) !important;
		border: 1px solid var(--line) !important;
	}
	li[role="option"] {
		background-color: var(--card-bg) !important;
		color: var(--text-main) !important;
	}
	li[role="option"]:hover, li[role="option"][aria-selected="true"] {
		background-color: #334155 !important;
		color: #ffffff !important;
	}
	button[kind="secondary"], .stButton > button:not([kind="primary"]), .stDownloadButton > button:not([kind="primary"]) {
		background-color: var(--card-bg) !important;
		color: var(--teal) !important;
		border: 1px solid var(--teal) !important;
	}
	button[kind="secondary"]:hover, .stButton > button:not([kind="primary"]):hover {
		background-color: #334155 !important;
		border-color: var(--teal-hover) !important;
	}
	button[kind="primary"], .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
		background-color: #0f766e !important;
		color: #ffffff !important;
		border: 1px solid var(--teal) !important;
	}
	button[kind="primary"]:hover, .stButton > button[kind="primary"]:hover {
		background-color: #115e59 !important;
	}
	[data-testid="stExpander"] {
		background-color: var(--card-bg) !important;
		border: 1px solid var(--line) !important;
	}
	[data-testid="stExpander"] summary {
		color: var(--text-heading) !important;
	}
	[data-testid="stExpander"] summary * {
		color: var(--text-heading) !important;
	}
	button[data-baseweb="tab"] {
		color: var(--text-muted) !important;
	}
	button[data-baseweb="tab"][aria-selected="true"] {
		color: var(--teal) !important;
		border-bottom-color: var(--teal) !important;
		font-weight: 700 !important;
	}
	[data-testid="stDataFrame"] {
		background-color: var(--card-bg) !important;
		border: 1px solid var(--line) !important;
	}
	div[role="dialog"], [data-testid="stModal"] {
		background-color: var(--card-bg) !important;
		color: var(--text-main) !important;
		border: 1px solid var(--line) !important;
	}
	div[role="dialog"] * {
		color: var(--text-main);
	}
	div[role="dialog"] h1, div[role="dialog"] h2, div[role="dialog"] h3, div[role="dialog"] h4 {
		color: #ffffff !important;
	}
	[data-testid="stAlert"] {
		background-color: var(--card-bg) !important;
		border-color: var(--line) !important;
	}
	[data-testid="stAlert"] * {
		color: var(--text-main) !important;
	}
	[data-testid="stRadio"] label, [data-testid="stCheckbox"] label {
		color: var(--text-main) !important;
	}
	[data-testid="stRadio"] label p, [data-testid="stCheckbox"] label p {
		color: var(--text-main) !important;
	}
	table tr td {
		background-color: var(--card-bg) !important;
		color: var(--text-main) !important;
		border-color: var(--line) !important;
	}
	table tr th {
		background-color: #0f766e !important;
		color: #ffffff !important;
		border-color: var(--line) !important;
	}
	hr {
		border-color: var(--line) !important;
	}
	span[style*="background:#fef2f2"], div[style*="background:#fee2e2"] {
		background-color: #450a0a !important;
		color: #fca5a5 !important;
		border-color: #991b1b !important;
	}
	span[style*="background:#f0fdf4"], div[style*="background:#dcfce7"] {
		background-color: #052e16 !important;
		color: #86efac !important;
		border-color: #166534 !important;
	}
	span[style*="background:#eff6ff"], div[style*="background:#e0f2fe"] {
		background-color: #082f49 !important;
		color: #7dd3fc !important;
		border-color: #0369a1 !important;
	}
	span[style*="background:#fffbeb"] {
		background-color: #451a03 !important;
		color: #fcd34d !important;
		border-color: #92400e !important;
	}
	div[style*="background: rgba(0,0,0,0.03)"] {
		background: rgba(255, 255, 255, 0.06) !important;
	}
	"""

	if is_forced_dark:
		return f"<style>{base_rules}{dark_rules}</style>"
	elif is_forced_light:
		light_lock = f"""
		:root, .stApp {{
			color-scheme: light !important;
		}}
		.stApp {{
			background: {theme['bg_grad']} !important;
			color: {theme['ink']} !important;
		}}
		h1, h2, h3, h4, h5, h6 {{
			color: {theme['ink']} !important;
		}}
		p, span, label, div {{
			color: {theme['ink']};
		}}
		[data-testid="stSidebar"] {{
			background: {theme['sidebar']} !important;
		}}
		[data-testid="stSidebar"] * {{
			color: {theme['ink']} !important;
		}}
		[data-testid="stVerticalBlockBorderWrapper"] > div {{
			background: {theme['card_bg']} !important;
			border: 1px solid {theme['line']} !important;
		}}
		div[data-baseweb="input"] input, textarea {{
			background-color: #ffffff !important;
			color: #0f172a !important;
		}}
		div[data-baseweb="select"] > div {{
			background-color: #ffffff !important;
			color: #0f172a !important;
		}}
		div[data-baseweb="select"] * {{
			color: #0f172a !important;
		}}
		"""
		return f"<style>{base_rules}{light_lock}</style>"
	else:
		# Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)
		auto_rules = f"""
		@media (prefers-color-scheme: dark) {{
			{dark_rules}
		}}
		[data-theme="dark"], .stApp[data-theme="dark"], [data-base="dark"] {{
			{dark_rules}
		}}
		"""
		return f"<style>{base_rules}{auto_rules}</style>"


FONT_OPTIONS: dict[str, tuple[str, str]] = {
	"DM Sans": ("'DM Sans', sans-serif", "https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap"),
	"Manrope": ("'Manrope', sans-serif", "https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap"),
	"Inter": ("'Inter', sans-serif", "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap"),
	"Plus Jakarta Sans": ("'Plus Jakarta Sans', sans-serif", "https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap"),
	"Nunito": ("'Nunito', sans-serif", "https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&display=swap"),
	"Roboto": ("'Roboto', sans-serif", "https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap"),
	"Merriweather": ("'Merriweather', serif", "https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&display=swap"),
}

THEME_PRESETS: dict[str, dict[str, str]] = {
	"Klinik Turkuaz (Varsayılan)": {
		"ink": "#183b3a",
		"teal": "#176b68",
		"teal_hover": "#12524f",
		"mint": "#eaf3f1",
		"sidebar": "#e7f0ee",
		"bg_grad": "linear-gradient(155deg,#f4f8f7 0%,#f8faf8 48%,#f4f7f7 100%)",
		"card_bg": "#ffffff",
		"line": "#d7e2df",
	},
	"Okyanus Mavisi": {
		"ink": "#0f172a",
		"teal": "#0284c7",
		"teal_hover": "#0369a1",
		"mint": "#e0f2fe",
		"sidebar": "#f0f9ff",
		"bg_grad": "linear-gradient(155deg,#f0f9ff 0%,#f8fafc 50%,#e0f2fe 100%)",
		"card_bg": "#ffffff",
		"line": "#bae6fd",
	},
	"Zümrüt Yeşili": {
		"ink": "#14532d",
		"teal": "#15803d",
		"teal_hover": "#166534",
		"mint": "#dcfce7",
		"sidebar": "#f0fdf4",
		"bg_grad": "linear-gradient(155deg,#f0fdf4 0%,#f8fafc 50%,#dcfce7 100%)",
		"card_bg": "#ffffff",
		"line": "#bbf7d0",
	},
	"Sıcak Toprak & Mercan": {
		"ink": "#431407",
		"teal": "#c2410c",
		"teal_hover": "#9a3412",
		"mint": "#ffedd5",
		"sidebar": "#fff7ed",
		"bg_grad": "linear-gradient(155deg,#fff7ed 0%,#fafaf9 50%,#ffedd5 100%)",
		"card_bg": "#ffffff",
		"line": "#fed7aa",
	},
	"Lavanta & İndigo": {
		"ink": "#1e1b4b",
		"teal": "#4f46e5",
		"teal_hover": "#4338ca",
		"mint": "#e0e7ff",
		"sidebar": "#eef2ff",
		"bg_grad": "linear-gradient(155deg,#eef2ff 0%,#f8fafc 50%,#e0e7ff 100%)",
		"card_bg": "#ffffff",
		"line": "#c7d2fe",
	},
	"Koyu Gece (Dark)": {
		"ink": "#f1f5f9",
		"teal": "#14b8a6",
		"teal_hover": "#0d9488",
		"mint": "#1e293b",
		"sidebar": "#0f172a",
		"bg_grad": "linear-gradient(155deg,#0f172a 0%,#1e293b 100%)",
		"card_bg": "#1e293b",
		"line": "#334155",
	},
}


DEFAULT_NANDA_CARE_PLAN_TEMPLATES: dict[str, dict[str, str]] = {
	"📉 Büyüme & Gelişme Yetersizliği Riski (NANDA: 00113)": {
		"category": "Büyüme & Gelişme",
		"diagnosis": "Büyümede yetersizlik ve gelişme geriliği riski (NANDA Kodu: 00113). İlişkili Faktörler: Yetersiz kalori ve mikro besin alımı, iştahsızlık veya emilim bozukluğu. Belirtiler: Persantil eğrisinde kanal kaybı, Z-skoru < -2 SDS (3. persentil altı bodurluk/zayıflık).",
		"observed_facts": "Antropometrik ölçümlerde boy/kilo persantili yaş ve cinsiyet medyanının gerisinde (-2 SDS altı). Persantil kanalında düşüş ve kilo alma hızında duraklama gözlendi.",
		"goals": "1. 3 ay içerisinde persantil kanal kaybının durdurulması ve kanal stabilitesinin yakalanması.\n2. Yaşa ve cinsiyete uygun günlük kalori ve protein ihtiyacının karşılanması.\n3. Kilo kaybının önlenmesi ve büyüme hızının pozitif ivmeye geçmesi.",
		"interventions": "1. Revirde 15 günde bir standart koşullarda (sabah aç karnına, ayakkabısız) boy ve kilo ölçümü yapılarak persantil grafiğine işlenir.\n2. Veli ile görüşülerek 3 günlük ayrıntılı beslenme günlüğü (öğün saatleri, atıştırmalıklar) tutulması istenir.\n3. Okul saatlerinde beslenme çantası içeriği ve yemekhane tüketimi izlenir; ara öğünlerde ceviz, fındık, kuru meyve ve süt teşvik edilir.\n4. Sınıf öğretmeniyle görüşülerek ders içi odaklanma, yorgunluk ve uyuklama halleri takip edilir.\n5. Büyüme hormonu, çölyak, parazitoz veya demir eksikliği açısından aile hekimi ve çocuk endokrinolojisi randevusu koordine edilir.",
		"monitoring": "15 günde bir persantil/Z-skor takibi; ayda bir veli görüşmesi ve beslenme değerlendirmesi.",
		"escalation": "Ardışık 2 ölçümde kilo kaybı devam ederse veya boy persantili -3 SDS altına gerilerse derhal Çocuk Sağlığı ve Çocuk Endokrinolojisi Uzmanına sevk.",
		"status": "Aktif - İzlemde",
	},
	"🍔 Dengesiz Beslenme: Fazla Kilo & Obezite Riski (NANDA: 00232)": {
		"category": "Beslenme & Metabolizma",
		"diagnosis": "Dengesiz Beslenme: Beden Gereksiniminden Fazla Beslenme / Obezite Riski (NANDA Kodu: 00232). İlişkili Faktörler: Enerji yoğunluğu yüksek işlenmiş gıdalar, şekerli içecek tüketimi, sedanter okul/ev yaşamı. Belirtiler: VKİ Z-skoru > +2 SDS (97. persentil üzeri), çabuk yorulma, efor dispnesi.",
		"observed_facts": "Vücut Kitle İndeksi (VKİ) persantili 97. persentilin üzerinde (fazla kilolu/obezite sınırında). Okul içi fiziksel aktivitelere katılımda isteksizlik, kantinden paketli hazır gıda tüketim sıklığı mevcut.",
		"goals": "1. Kilo artış hızının boy uzamasıyla dengelenmesi (kilo koruma ve lineer büyüme ile VKİ normalizasyonu).\n2. Günlük fiziksel aktivite süresinin en az 60 dakikaya çıkarılması.\n3. Şekerli içecek ve rafine gıda tüketiminin sonlandırılması.",
		"interventions": "1. Ayda bir boy, kilo ve VKİ persantil takibi yapılarak değişim eğrisi çıkarılır.\n2. Beden eğitimi öğretmeni ile iş birliği yapılarak öğrencinin akran zorbalığına maruz kalmadan eğlenceli fiziksel aktivitelere (yürüyüş, ip atlama, takım oyunları) katılımı sağlanır.\n3. Okul kantininde şekerli içecek ve cips yerine ayran, süt, su ve meyve tercih etmesi konusunda sağlık danışmanlığı verilir.\n4. Aileye porsiyon kontrolü, ekran süresinin günde 1 saatin altına indirilmesi ve düzenli aile yürüyüşleri konularında rehberlik sağlanır.\n5. Rutin tansiyon ve açlık kan şekeri taraması yapılır.",
		"monitoring": "Ayda bir VKİ persantil takibi ve 3 ayda bir vital (tansiyon) kontrolü.",
		"escalation": "VKİ Z-skoru > +3 SDS olması, persantil artışının durdurulamaması veya tansiyon yüksekliği halinde Çocuk Sağlığı / Çocuk Endokrinolojisi ve Diyetisyene yönlendirme.",
		"status": "Aktif - İzlemde",
	},
	"🥗 Dengesiz Beslenme: Gereksinimden Az / Malnütrisyon Riski (NANDA: 00002)": {
		"category": "Beslenme & Metabolizma",
		"diagnosis": "Dengesiz Beslenme: Beden Gereksiniminden Az Beslenme (NANDA Kodu: 00002). İlişkili Faktörler: Yetersiz kalori/besin alımı, seçici yeme alışkanlığı, iştahsızlık, sosyoekonomik güçlükler. Belirtiler: VKİ < 3. persentil (-2 SDS altı), halsizlik, solukluk, deri turgorunda azalma.",
		"observed_facts": "Öğrencinin ağırlığı yaşıtlarına göre belirgin düşük, derslerde odaklanma güçlüğü ve halsizlik şikayeti var. Cilt ve mukozalar soluk.",
		"goals": "1. 1 ay içerisinde kilo kaybının durdurulması ve en az %5 ağırlık artışı sağlanması.\n2. Günlük enerji ve sıvı ihtiyacının eksiksiz karşılanması.\n3. Halsizlik ve odaklanma güçlüğünün giderilmesi.",
		"interventions": "1. Haftalık ağırlık ölçümü ve günlük sıvı tüketim takibi yapılır.\n2. Okul yönetimi ve rehberlik servisiyle görüşülerek ücretsiz sağlıklı okul beslenme desteği veya beslenme çantası takviyesi planlanır.\n3. Yemek öncesi su içmeyip yemek sırasında ve sonrasında tüketmesi; kalori ve besin değeri yoğun gıdaların seçilmesi anlatılır.\n4. Anemi ve vitamin eksiklikleri açısından revirde konjonktiva ve tırnak muayenesi yapılır.\n5. Veli ile ev içi öğün düzeni planlanır.",
		"monitoring": "Haftalık ağırlık kontrolü, ayda bir persantil güncellemesi.",
		"escalation": "Kilo kaybının devam etmesi, bayılma/baş dönmesi veya belirgin anemi bulgularında Çocuk Sağlığı ve Hastalıkları uzmanına sevk.",
		"status": "Aktif - İzlemde",
	},
	"🫁 Etkisiz Solunum Örüntüsü / Astım ve Bronkospazm (NANDA: 00032)": {
		"category": "Solunum & Alerji",
		"diagnosis": "Etkisiz Solunum Örüntüsü ve Gaz Değişiminde Bozulma Riski (NANDA Kodu: 00032). İlişkili Faktörler: Bronşiyal hiperreaktivite, alerjenler (polen, toz, tebeşir, akarlar), soğuk hava veya eforla tetiklenen bronkospazm. Belirtiler: Nefes darlığı, göğüste hırıltı (ronküs), kuru öksürük, SpO2 düşüşü.",
		"observed_facts": "Bilinen astım / reaktif hava yolu öyküsü. Beden eğitimi dersi veya soğuk havada öksürük nöbetleri ve ekspiratuar hırıltı şikayetiyle revire başvuru.",
		"goals": "1. Okul saatlerinde SpO2 düzeyinin %95 ve üzerinde tutulması.\n2. Solunum sıkıntısının erken evrede kontrol altına alınması.\n3. Öğrencinin tetikleyicilerden korunması ve inhaler kullanım becerisinin tam olması.",
		"interventions": "1. Revire başvurduğunda derhal rahat oturur (High-Fowler) pozisyona alınır, yakası ve sıkan giysileri gevşetilir, ortam havalandırılır.\n2. SpO2, nabız ve solunum sayısı ölçülür.\n3. Reçeteli hızlı etkili bronkodilatör (Salbutamol inhaler) varsa hazne (spacer) ile doğru teknikle uygulanması sağlanır/gözetlenir.\n4. Beden eğitimi dersleri öncesinde ısınma egzersizleri ve hekim önerisiyle koruyucu inhaler protokolü uygulanır.\n5. Sınıf ortamının tozsuz tutulması, tebeşir tozundan kaçınılması ve oda kokularının kullanılmaması sınıf öğretmenine bildirilir.",
		"monitoring": "Atak anında 5-10 dakikada bir SpO2 ve solunum sesleri takibi.",
		"escalation": "İnhaler uygulamasına rağmen 15 dakikada düzelmeme, SpO2 <%92, konuşurken duraklama, interkostal çekilme veya dudaklarda siyanoz durumunda derhal 112 Acil Ambulans ve veli bildirimi.",
		"status": "Aktif - İzlemde",
	},
	"🥜 Alerjik Reaksiyon ve Anafilaksi Riski (NANDA: 00217)": {
		"category": "Solunum & Alerji",
		"diagnosis": "Alerjik Reaksiyon ve Hayatı Tehdit Eden Anafilaksi Riski (NANDA Kodu: 00217). İlişkili Faktörler: Besin alerjenleri (kuruyemiş, süt, yumurta, gluten, deniz ürünleri), arı/böcek sokması, lateks veya ilaç teması. Belirtiler: Yaygın ürtiker, anjiyoödem (dudak, göz kapağı şişmesi), stridor, laringeal spazm, hipotansiyon.",
		"observed_facts": "Öğrencinin bilinen besin/ilaç alerjisi mevcut. Şüpheli temas sonrası kaşıntı, döküntü veya nefes darlığı riski.",
		"goals": "1. Okul ortamında alerjenle temasın %100 engellenmesi.\n2. Olası maruziyette 3 dakika içinde acil eylem planının devreye sokulması.\n3. Solunum yolu tıkanıklığı ve anafilaktik şokun önlenmesi.",
		"interventions": "1. Öğrencinin alerjileri yemekhane, kantin ve nöbetçi öğretmenlerin listesine acil not olarak eklenir; akranlarıyla yiyecek paylaşımı kesinlikle yasaklanır.\n2. Aile tarafından sağlanan Adrenalin Oto-Enjektörü (EpiPen / Penepin) revirde oda sıcaklığında, kolay erişilebilir özel acil dolabında ve son kullanma tarihi etiketli olarak saklanır.\n3. Öğretmenlere oto-enjektör uygulama tekniği (uyluk dış yan bölgesine dik açıyla 10 saniye basılı tutma) konusunda uygulamalı bilgi verilir.\n4. Reaksiyon belirtilerinde (kaşıntı, ses kısıklığı, öksürük, şişlik) çocuk asla yalnız bırakılmaz, sırtüstü yatırılıp bacakları 30 cm yükseltilir (şok pozisyonu).",
		"monitoring": "Maruziyet şüphesinde dakikalık nabız, tansiyon, solunum ve deri döküntüsü izlemi.",
		"escalation": "Dudak/boğazda şişme, ses kısıklığı, nefes darlığı veya bayılma hissi başladığı an derhal Adrenalin oto-enjektörü uygulanır ve beklemeden 112 Acil Ambulans çağrılır.",
		"status": "Aktif - İzlemde",
	},
	"🩸 Kararsız Kan Glikoz Düzeyi Riski / Tip 1 Diyabet (NANDA: 00179)": {
		"category": "Kronik Durumlar & Endokrin",
		"diagnosis": "Kararsız Kan Glikoz Düzeyi Riski (NANDA Kodu: 00179). İlişkili Faktörler: İnsülin dozajı ile karbonhidrat alımı/fiziksel aktivite arasındaki dengesizlik, öğün gecikmesi, enfeksiyonlar. Belirtiler: Hipoglisemi (glikoz <70 mg/dL, soğuk terleme, titreme, solukluk, çarpıntı, baş dönmesi, konfüzyon) veya Hiperglisemi (glikoz >250 mg/dL, ketonüri).",
		"observed_facts": "Tip 1 Diyabet tanısı, cilt altı insülin pompası veya kalem kullanımı, kan şekeri sensörü takibi.",
		"goals": "1. Okul saatleri süresince kan glikozunun 80-180 mg/dL hedef aralığında tutulması.\n2. Ağır hipoglisemi ve bilinç kaybının önlenmesi.\n3. Öğrencinin bağımsız kan şekeri kontrol ve öğün yönetim becerisinin desteklenmesi.",
		"interventions": "1. Öğrencinin öğle yemeği öncesi ve beden eğitimi öncesi/sonrası revirde veya sınıfta glikoz ölçümü yapması sağlanır/kaydedilir.\n2. Hipoglisemi durumunda 15-15 Kuralı uygulanır: Kan şekeri <70 mg/dL ise 15 gram hızlı karbonhidrat (3-4 adet kesme şeker, yarım kutu meyve suyu) verilir; 15 dakika sonra glikoz tekrar ölçülür; normale dönene kadar dinlendirilir.\n3. Öğrencinin çantasında ve revirde acil hipoglisemi kiti (meyve suyu, glukagon hipokit) hazır bulundurulur.\n4. Sınıf öğretmenlerine hipoglisemi erken uyarı işaretleri (dalgınlık, terleme, ani sinirlilik) öğretilir; öğrencinin derste su içmesine ve tuvalete çıkmasına izin verilir.",
		"monitoring": "Günlük öğün ve egzersiz öncesi glikoz takibi.",
		"escalation": "Bilinç bulanıklığı, yutkunamama veya bayılma durumunda ağızdan hiçbir şey verilmez, derhal İntramüsküler Glukagon uygulanır ve 112 Acil Ambulans çağrılır.",
		"status": "Aktif - İzlemde",
	},
	"⚡ Travma / Nöbet Güvenliği Riski (Epilepsi) (NANDA: 00039)": {
		"category": "Kronik Durumlar & Nöroloji",
		"diagnosis": "Travma ve Yaralanma Riski / Nöbet Güvenliği (NANDA Kodu: 00039). İlişkili Faktörler: Serebral elektriksel deşarj bozukluğu, uyku eksikliği, stres, parlak ışıklar veya enfeksiyon. Belirtiler: Bilinç kaybı, tonik-klonik kasılmalar, diş kenetlenmesi, ağızdan köpük gelmesi, idrar kaçırma.",
		"observed_facts": "Bilinen epilepsi / konvülsiyon öyküsü, düzenli antiepileptik ilaç kullanımı.",
		"goals": "1. Nöbet sırasında ve sonrasında fiziksel yaralanma ve kafa travmasının önlenmesi.\n2. Solunum yolu açıklığının ve oksijenasyonun korunması (aspirasyonun önlenmesi).\n3. Nöbet süresinin kayıt altına alınması.",
		"interventions": "1. Nöbet başladığı an sakin olunur, öğrenci yere yavaşça yatırılır ve başının altına yumuşak bir yastık veya mont konur.\n2. Çevredeki sivri, sert masa, sıra ve eşyalar derhal uzaklaştırılır; gözlük varsa çıkarılır, sıkan yaka açılır.\n3. ASLA ağzına kaşık, bez sokulmaz, çenesi zorla açılmaya çalışılmaz, kasılan uzuvları bastırılarak durdurulmaz.\n4. Nöbet bitiminde salgıların akciğere kaçmaması için yan yatış (koma/iyileşme pozisyonu) verilir.\n5. Nöbetin başladığı saat kronometre ile saniye saniye takip edilir.",
		"monitoring": "Nöbet süresi, nöbet tipi, nöbet sonrası bilinç toparlanma süresi kaydı.",
		"escalation": "Nöbet 5 dakikadan uzun sürerse, arka arkaya tekrarlarsa, nöbet sırasında kafa travması oluşursa veya nöbet sonrası solunum güçlüğü devam ederse derhal 112 Acil Ambulans ve veli bildirimi.",
		"status": "Aktif - İzlemde",
	},
	"🌡️ Hipertermi / Akut Yüksek Ateş Yönetimi (NANDA: 00007)": {
		"category": "Akut Bakım & Enfeksiyon",
		"diagnosis": "Hipertermi (NANDA Kodu: 00007). İlişkili Faktörler: Viral/bakteriyel üst/alt solunum yolu enfeksiyonu, dehidratasyon, aşırı giydirilme. Belirtiler: Vücut sıcaklığı > 38.0 °C, taşikardi, titreme veya sıcak/kırmızı cilt, baş ağrısı, halsizlik.",
		"observed_facts": "Vücut sıcaklığı ölçümünde 38.0 °C ve üzerinde ateş saptandı, öğrenci halsiz ve derste başını sıraya koyuyor.",
		"goals": "1. Vücut sıcaklığının güvenli sınırlara (37.0 °C - 37.5 °C) düşürülmesi.\n2. Ateşli havale (febril konvülsiyon) riskinin önlenmesi.\n3. Dehidratasyonun (sıvı kaybının) telafi edilmesi.",
		"interventions": "1. Kalın giysiler, hırka ve mont çıkarılır; oda sıcaklığı 21-22 °C seviyesinde havalandırılır.\n2. Boyun, koltuk altı ve alın bölgesine oda sıcaklığında (ılık) ıslatılmış pamuklu kompres uygulanır (ASLA çok soğuk su, buz veya ispirto/alkol sürülmez).\n3. Öğrenciye yudum yudum ılık su veya bitki çayı içirilir.\n4. Veliye telefonla ulaşılarak hekim reçeteli ateş düşürücü kullanımı için onay ve bilgi alınır.",
		"monitoring": "15-20 dakikada bir timpanik veya aksiller ateş ölçümü, nabız ve bilinç takibi.",
		"escalation": "Ateşin 38.8 °C üzerine çıkması, kontrolsüz titreme, morarma, ense sertliği veya bilinç bulanıklığı görülmesi halinde 112 Acil ve acil poliklinik sevki.",
		"status": "Aktif - İzlemde",
	},
	"🩹 Akut Ağrı ve Doku Yaralanması / Burkulma-Kesi (NANDA: 00132 & 00046)": {
		"category": "Akut Bakım & Travma",
		"diagnosis": "Akut Ağrı ve Doku Bütünlüğünde Bozulma Riski (NANDA Kodu: 00132 & 00046). İlişkili Faktörler: Okul bahçesinde veya spor salonunda düşme, çarpışma, burkulma, sıyrık veya kesikler. Belirtiler: Lokal ödem, ekimoz, kanama, hareketle artan şiddetli ağrı.",
		"observed_facts": "Teneffüste düşme sonucu eklemde burkulma veya deri bütünlüğünde sıyrık/kesi ile revire başvuru.",
		"goals": "1. Ağrı skorunun (Wong-Baker yüz skalası ile) düşürülmesi.\n2. Kanamanın durdurulması ve yara enfeksiyonunun engellenmesi.\n3. Eklem hareket ve fonksiyonunun korunması.",
		"interventions": "1. Kesik ve sıyrıklarda yara steril serum fizyolojik ile yıkanır, antiseptik uygulanıp steril gazlı bezle kapatılır.\n2. Yumuşak doku travması ve burkulmalarda R.I.C.E. Protokolü uygulanır:\n   * Rest (İstirahat): Eklem hareketsiz bırakılır.\n   * Ice (Buz): Cilt üzerine havlu sarılarak 15 dakika soğuk kompres uygulanır.\n   * Compression (Bası): Elastik bandaj ile hafif sıkılıkta sarılır.\n   * Elevation (Yükseltme): Ödemi azaltmak için kalp seviyesinin üstüne kaldırılır.\n3. Kırık/çıkık şüphesinde kesinlikle ekstremite çekilmez veya yerine oturtulmaya çalışılmaz; mevcut pozisyonda atele alınır.",
		"monitoring": "15 dakikada bir dolaşım kontrolü (parmak rengi, nabız, soğukluk).",
		"escalation": "Şekil bozukluğu, kemik krepitasyonu, durdurulamayan arteriyel kanama veya derin kesilerde 112 Acil / Ortopedi sevki.",
		"status": "Aktif - İzlemde",
	},
	"🤕 Tekrarlayan Karın/Baş Ağrısı & Okul Stresi (NANDA: 00133 & 00146)": {
		"category": "Psikososyal & Fonksiyonel",
		"diagnosis": "Kronik Ağrı ve Okul Uyumuna Bağlı Anksiyete (NANDA Kodu: 00133 & 00146). İlişkili Faktörler: Sınav kaygısı, akran ilişkileri stresi, ayrılık anksiyetesi, yetersiz uyku veya beslenme. Belirtiler: Sık tekrarlayan organik bulgusuz karın ağrısı veya gerilim baş ağrısı, ders sırasında revire gelme isteği.",
		"observed_facts": "Haftada 2 veya daha fazla kez organik patoloji olmaksızın benzer saatlerde baş/karın ağrısı şikayetiyle revir başvurusu.",
		"goals": "1. Ağrının organik ve somatizasyon ayrımının yapılması.\n2. Öğrencinin gevşeme ve baş etme teknikleri kazanması.\n3. Revirde kalış süresinin sınırlanarak sınıfa uyumun güçlendirilmesi.",
		"interventions": "1. Karın muayenesinde defans, rebaund, ateş kontrolü yapılır; organik akut batın bulguları ekarte edilir.\n2. Revirde sakin, loş bir köşede 15-20 dakika istirahat, derin nefes egzersizi ve bir bardak ılık su içirilmesi sağlanır.\n3. Ağrının okul saatleri, belirli dersler veya sınav günleri ile korelasyonu takvimlenir.\n4. Okul psikolojik danışmanı / rehberlik servisi ile vaka değerlendirme toplantısı yapılır.\n5. Veli ile görüşülerek ev içi rutinler, uyku düzeni ve göz muayenesi ihtiyacı sorgulanır.",
		"monitoring": "Haftalık revir başvuru sıklığı ve rehberlik takibi.",
		"escalation": "Ağrının gece uykudan uyandırması, kusma, kilo kaybı veya görme bozukluğu eşlik etmesi halinde Çocuk Nörolojisi / Gastroenteroloji hekimine sevk.",
		"status": "Aktif - İzlemde",
	},
	"🧴 Bulaşıcı Enfeksiyon & Paraziter Hijyen Riski (NANDA: 00004)": {
		"category": "Akut Bakım & Enfeksiyon",
		"diagnosis": "Bulaşıcı Enfeksiyon Riski (NANDA Kodu: 00004). İlişkili Faktörler: Patojen mikroorganizmalarla temas, yetersiz el hijyeni, paraziter enfeksiyonlar (pediküloz, uyuz) veya döküntülü çocukluk çağı hastalıkları. Belirtiler: Kaşıntı, saçlı deride sirke/bit, cilt döküntüleri, konjonktivit.",
		"observed_facts": "Saçlı deride yoğun kaşıntı ve bit sirkeleri tespit edildi veya ciltte bulaşıcı lezyon şüphesi.",
		"goals": "1. Okul içi ve sınıf akranlarına bulaşın sıfırlanması.\n2. Doğru hijyen ve parazit eliminasyon protokolünün uygulanması.\n3. Cilt/saç bütünlüğünün iyileştirilmesi.",
		"interventions": "1. Öğrenci revirde hassasiyetle, akranları tarafından etiketlenmeden muayene edilir ve durum gizlilik esasına göre veliye bildirilir.\n2. Aileye Sağlık Bakanlığı onaylı medikal pediküloz şampuanı/losyonu kullanım rehberi ve sık dişli metal tarak kullanımı anlatılır.\n3. Evdeki tüm yastık kılıfı, çarşaf ve giysilerin en az 60 °C sıcaklıkta yıkanması veya ağzı kapalı poşette 14 gün bekletilmesi talimatı verilir.\n4. Sınıfta genel tarama ve el yıkama farkındalık bilgilendirmesi yapılır.",
		"monitoring": "Tedavi başlangıcından sonra 7. ve 14. günlerde revirde kontrol taraması.",
		"escalation": "İkincil bakteriyel enfeksiyon (ciltte akıntı, impetigo) veya dirençli parazitoz durumunda Aile Hekimi / Cildiye Polikliniğine sevk.",
		"status": "Aktif - İzlemde",
	},
	"💤 Kronik Yorgunluk ve Uyku Düzeni Bozukluğu (NANDA: 00093)": {
		"category": "Psikososyal & Yaşam Tarzı",
		"diagnosis": "Yorgunluk ve Uyku Örüntüsünde Bozulma (NANDA Kodu: 00093). İlişkili Faktörler: Gece geç saatlere kadar ekran kullanımı, uyku apnesi / burun tıkanıklığı, düzensiz uyku saatleri, anksiyete. Belirtiler: Sabah uyanamama, ilk derslerde uyuklama, dikkat eksikliği, göz altlarında morluk.",
		"observed_facts": "Ders saatlerinde sürekli esneme, dikkat dağınıklığı ve ilk derslerde sıraya yatarak uyuma şikayeti ile revire yönlendirildi.",
		"goals": "1. Öğrencinin yaş grubuna uygun (günde en az 9-10 saat) kesintisiz gece uykusu alışkanlığı kazanması.\n2. Ders içi uyanıklık ve konsantrasyonun artırılması.\n3. Ekran maruziyetinin uyku öncesinde sınırlandırılması.",
		"interventions": "1. Öğrencinin uyku öncesi telefon/tablet/bilgisayar kullanım saatleri ve yatak saati sorgulanır.\n2. Aileye uyku hijyeni rehberi verilir: Uykudan en az 1 saat önce tüm mavi ışıklı ekranların kapatılması, yatak odasının karanlık, sessiz ve serin tutulması önerilir.\n3. Burun tıkanıklığı, geniz eti veya horlama şikayeti varsa KBB muayenesi önerilir.\n4. Gün içinde açık havada fiziksel aktivite ve güneş ışığı alması teşvik edilir.",
		"monitoring": "Haftalık ders takip formu ve veli bildirimi.",
		"escalation": "Gündüz aşırı uyku hali (narkolepsi şüphesi) veya gece solunum durması (uyku apnesi) şüphesinde Çocuk Nörolojisi ve Uyku Polikliniğine sevk.",
		"status": "Aktif - İzlemde",
	},
}


NANDA_CARE_PLAN_TEMPLATES = DEFAULT_NANDA_CARE_PLAN_TEMPLATES


def get_care_plan_templates() -> dict[str, dict[str, str]]:
	settings = get_app_settings()
	raw = settings.get("custom_care_plan_templates", "")
	if raw and raw.strip():
		try:
			parsed = json.loads(raw)
			if isinstance(parsed, dict) and parsed:
				return parsed
		except Exception:
			pass
	return copy.deepcopy(DEFAULT_NANDA_CARE_PLAN_TEMPLATES)


def save_care_plan_templates(templates: dict[str, dict[str, str]]) -> None:
	save_app_settings({
		"custom_care_plan_templates": json.dumps(templates, ensure_ascii=False, indent=2)
	})


def reset_care_plan_templates() -> None:
	save_app_settings({
		"custom_care_plan_templates": ""
	})

DEFAULT_NANDA_DIAGNOSES: dict[str, str] = {
	"📉 Büyüme & Gelişme Yetersizliği Riski (NANDA: 00113)": "Büyümede yetersizlik ve gelişme geriliği riski (NANDA Kodu: 00113). İlişkili Faktörler: Yetersiz kalori ve mikro besin alımı, iştahsızlık veya emilim bozukluğu. Belirtiler: Persantil eğrisinde kanal kaybı, Z-skoru < -2 SDS (3. persentil altı bodurluk/zayıflık).",
	"🍔 Dengesiz Beslenme: Fazla Kilo & Obezite Riski (NANDA: 00232)": "Dengesiz Beslenme: Beden Gereksiniminden Fazla Beslenme / Obezite Riski (NANDA Kodu: 00232). İlişkili Faktörler: Enerji yoğunluğu yüksek işlenmiş gıdalar, şekerli içecek tüketimi, sedanter okul/ev yaşamı. Belirtiler: VKİ Z-skoru > +2 SDS (97. persentil üzeri), çabuk yorulma, efor dispnesi.",
	"🥗 Dengesiz Beslenme: Gereksinimden Az / Malnütrisyon (NANDA: 00002)": "Dengesiz Beslenme: Beden Gereksiniminden Az Beslenme (NANDA Kodu: 00002). İlişkili Faktörler: Yetersiz kalori/besin alımı, seçici yeme alışkanlığı, iştahsızlık, sosyoekonomik güçlükler. Belirtiler: VKİ < 3. persentil (-2 SDS altı), halsizlik, solukluk, deri turgorunda azalma.",
	"🫁 Etkisiz Solunum Örüntüsü / Astım ve Bronkospazm (NANDA: 00032)": "Etkisiz Solunum Örüntüsü ve Gaz Değişiminde Bozulma Riski (NANDA Kodu: 00032). İlişkili Faktörler: Bronşiyal hiperreaktivite, alerjenler (polen, toz, tebeşir, akarlar), soğuk hava veya eforla tetiklenen bronkospazm. Belirtiler: Nefes darlığı, göğüste hırıltı (ronküs), kuru öksürük, SpO2 düşüşü.",
	"🥜 Alerjik Reaksiyon ve Anafilaksi Riski (NANDA: 00217)": "Alerjik Reaksiyon ve Hayatı Tehdit Eden Anafilaksi Riski (NANDA Kodu: 00217). İlişkili Faktörler: Besin alerjenleri (kuruyemiş, süt, yumurta, gluten, deniz ürünleri), arı/böcek sokması, lateks veya ilaç teması. Belirtiler: Yaygın ürtiker, anjiyoödem (dudak, göz kapağı şişmesi), stridor, laringeal spazm, hipotansiyon.",
	"🩸 Kararsız Kan Glikoz Düzeyi Riski / Tip 1 Diyabet (NANDA: 00179)": "Kararsız Kan Glikoz Düzeyi Riski (NANDA Kodu: 00179). İlişkili Faktörler: İnsülin dozajı ile karbonhidrat alımı/fiziksel aktivite arasındaki dengesizlik, öğün gecikmesi, enfeksiyonlar. Belirtiler: Hipoglisemi (glikoz <70 mg/dL, soğuk terleme, titreme, solukluk, çarpıntı, baş dönmesi, konfüzyon) veya Hiperglisemi (glikoz >250 mg/dL, ketonüri).",
	"⚡ Travma / Nöbet Güvenliği Riski (Epilepsi) (NANDA: 00039)": "Travma ve Yaralanma Riski / Nöbet Güvenliği (NANDA Kodu: 00039). İlişkili Faktörler: Serebral elektriksel deşarj bozukluğu, uyku eksikliği, stres, parlak ışıklar veya enfeksiyon. Belirtiler: Bilinç kaybı, tonik-klonik kasılmalar, diş kenetlenmesi, ağızdan köpük gelmesi, idrar kaçırma.",
	"🌡️ Hipertermi / Akut Yüksek Ateş (NANDA: 00007)": "Hipertermi (NANDA Kodu: 00007). İlişkili Faktörler: Viral/bakteriyel üst/alt solunum yolu enfeksiyonu, dehidratasyon, aşırı giydirilme. Belirtiler: Vücut sıcaklığı > 38.0 °C, taşikardi, titreme veya sıcak/kırmızı cilt, baş ağrısı, halsizlik.",
	"🩹 Akut Ağrı ve Doku Bütünlüğünde Bozulma Riski (NANDA: 00132 & 00046)": "Akut Ağrı ve Doku Bütünlüğünde Bozulma Riski (NANDA Kodu: 00132 & 00046). İlişkili Faktörler: Okul bahçesinde veya spor salonunda düşme, çarpışma, burkulma, sıyrık veya kesikler. Belirtiler: Lokal ödem, ekimoz, kanama, hareketle artan şiddetli ağrı.",
	"🤕 Tekrarlayan Karın/Baş Ağrısı & Okul Stresi (NANDA: 00133 & 00146)": "Kronik Ağrı ve Okul Uyumuna Bağlı Anksiyete (NANDA Kodu: 00133 & 00146). İlişkili Faktörler: Sınav kaygısı, akran ilişkileri stresi, ayrılık anksiyetesi, yetersiz uyku veya beslenme. Belirtiler: Sık tekrarlayan organik bulgusuz karın ağrısı veya gerilim baş ağrısı, ders sırasında revire gelme isteği.",
	"🧴 Bulaşıcı Enfeksiyon & Paraziter Hijyen Riski (NANDA: 00004)": "Bulaşıcı Enfeksiyon Riski (NANDA Kodu: 00004). İlişkili Faktörler: Patojen mikroorganizmalarla temas, yetersiz el hijyeni, paraziter enfeksiyonlar (pediküloz, uyuz) veya döküntülü çocukluk çağı hastalıkları. Belirtiler: Kaşıntı, saçlı deride sirke/bit, cilt döküntüleri, konjonktivit.",
	"💤 Kronik Yorgunluk ve Uyku Düzeni Bozukluğu (NANDA: 00093)": "Yorgunluk ve Uyku Örüntüsünde Bozulma (NANDA Kodu: 00093). İlişkili Faktörler: Gece geç saatlere kadar ekran kullanımı, uyku apnesi / burun tıkanıklığı, düzensiz uyku saatleri, anksiyete. Belirtiler: Sabah uyanamama, ilk derslerde uyuklama, dikkat eksikliği, göz altlarında morluk.",
	"💧 Sıvı Hacim Eksikliği Riski / Dehidratasyon (NANDA: 00028)": "Sıvı Hacim Eksikliği Riski (NANDA Kodu: 00028). İlişkili Faktörler: Akut gastroenterit, kusma, ishal, yetersiz sıvı alımı veya aşırı terleme. Belirtiler: Dudak ve dilde kuruluk, göz kürelerinde çöküklük, halsizlik, idrar miktarında azalma.",
	"🫀 Etkisiz Doku Perfüzyonu / Postüral Hipotansiyon (NANDA: 00204)": "Etkisiz Doku Perfüzyonu Riski (NANDA Kodu: 00204). İlişkili Faktörler: Uzun süre ayakta durma, açlık, dehidratasyon veya ani ayağa kalkma. Belirtiler: Göz kararması, baş dönmesi, soğuk terleme, solukluk, sistolik tansiyon düşüşü.",
	"🦷 Ağız Mukozasında Bozulma Riski / Diş Ağrısı (NANDA: 00045)": "Ağız Mukozasında Bozulma ve Akut Diş Ağrısı (NANDA Kodu: 00045). İlişkili Faktörler: Diş çürükleri, yetersiz ağız bakımı, diş apsesi veya gingivit. Belirtiler: Çiğneme güçlüğü, lokalize zonklayıcı ağrı, diş etinde hiperemi ve şişlik.",
	"😟 Durumsal Düşük Benlik Saygısı & Sosyal İzolasyon (NANDA: 00120 & 00053)": "Durumsal Düşük Benlik Saygısı ve Sosyal İzolasyon Riski (NANDA Kodu: 00120 & 00053). İlişkili Faktörler: Kronik hastalık damgası, akran zorbalığı, fiziksel aktivite kısıtlılığı veya iletişim güçlüğü. Belirtiler: Teneffüslerde yalnız kalma, içe kapanma, revire sık sığınma.",
	"👁️ Görme / İşitme Bozukluğuna Bağlı Düşme Riski (NANDA: 00155)": "Düşme ve Travma Riski (NANDA Kodu: 00155). İlişkili Faktörler: Kırılmamış gözlük, düzeltilmemiş kırma kusuru, işitme cihazı arızası veya nöromusküler koordinasyon eksikliği. Belirtiler: Merdiven inerken dengesizlik, sık takılma/düşme.",
	"🚽 Boşaltım Örüntüsünde Bozulma / Enürezis & Konstipasyon (NANDA: 00016 & 00011)": "Boşaltım Örüntüsünde Bozulma (NANDA Kodu: 00016 & 00011). İlişkili Faktörler: Okul tuvaletini kullanmaktan çekinme, yetersiz lif/su tüketimi, aşırı aktif mesane. Belirtiler: Karında şişkinlik, dışkılama erteleme, gündüz/gece idrar kaçırma.",
}

DEFAULT_NANDA_FACTS: dict[str, str] = {
	"📉 Persantil Kanal Kaybı & Bodurluk": "Antropometrik ölçümlerde boy/kilo persantili yaş ve cinsiyet medyanının gerisinde (-2 SDS altı). Persantil kanalında belirgin düşüş ve kilo alma hızında duraklama gözlendi.",
	"🍔 VKİ 97. Persentil Üzeri & Efor Dispnesi": "Vücut Kitle İndeksi (VKİ) persantili 97. persentilin üzerinde (+2.2 SDS). Merdiven çıkarken efor dispnesi ve akran oyunlarına katılımda çekinme gözlendi.",
	"🥗 Malnütrisyon / Belirgin Zayıflık": "Öğrencinin ağırlığı yaşıtlarına göre belirgin düşük (VKİ < 3. persentil, -2.4 SDS). Cilt soluk, derslerde başını sıraya koyma ve halsizlik mevcut.",
	"🫁 Ekspiratuar Wheezing & SpO2 Düşüşü": "Bilinen astım tanısı. Beden eğitimi dersi sonrası öksürük nöbeti, interkostal çekilme, dinlemekle ekspiratuar ronküs ve SpO2 %92 saptandı.",
	"🥜 Besin Teması Sonrası Ürtiker & Kaşıntı": "Öğle yemeği sonrası boyun ve gövdede eritematöz ürtikeryal döküntü, periorbital kaşıntı ve hafif ses kısıklığı başlangıcı gözlendi.",
	"🩸 Kan Şekeri 64 mg/dL & Soğuk Terleme": "Tip 1 Diyabet tanılı öğrenci. Derste dalgınlık, ellerde titreme ve soğuk terleme ile revire geldi. Parmak ucu glikoz ölçümü: 64 mg/dL.",
	"⚡ Tonik-Klonik Nöbet Gözlemi": "Sınıfta aniden yere düşme, 90 saniye süren generalize tonik-klonik kasılma, ağızda köpürme ve nöbet sonrası postiktal konfüzyon gözlendi.",
	"🌡️ Vücut Sıcaklığı 38.8 °C & Taşikardi": "Aksiller ateş ölçümünde 38.8 °C sıcaklık, nabız 128/dk, yanaklarda kızarıklık, titreme ve genel kas ağrısı saptandı.",
	"🩹 Ayak Bileğinde Ödem & Burkulma": "Okul bahçesinde koşarken ayak bileğini burkma. Dış malleol çevresinde ödem, ekimoz, basmakla şiddetli hassasiyet ve basamama durumu mevcut.",
	"🤕 Tekrarlayan Sınav Öncesi Karın Ağrısı": "Haftada 2-3 gün sabah ilk derslerde organik batın bulgusu (defans/rebound/ateş) olmaksızın şiddetli karın ağrısı şikayeti ile revire başvuru.",
	"🧴 Saçlı Deride Sirke ve Yoğun Kaşıntı": "Öğrencinin ensesinde ve kulak arkasında yoğun kaşıntı şikayeti üzerine yapılan muayenede saç tellerine yapışık çok sayıda canlı sirke tespit edildi.",
	"💤 Derste Sürekli Uyuklama & Yorgunluk": "İlk 3 derste sürekli uyuklama, esneme, göz altlarında belirgin morluk. Veli ile görüşüldüğünde gece 01:00'e kadar tablet kullandığı öğrenildi.",
	"💧 Dudaklarda Kuruluk & Koyu İdrar": "Dün başlayan 4 kez sulu ishal ve 2 kez kusma öyküsü. Cilt turgoru azalmış, dudaklar kuru, taşikardik ve halsiz görünümde.",
	"🫀 Ani Ayağa Kalkma Sonrası Presinkop": "Törende ayakta dururken ani göz kararması, baş dönmesi ve soğuk terleme. Ölçülen tansiyon: 85/55 mmHg, nabız: 96/dk.",
}

DEFAULT_NANDA_GOALS: dict[str, str] = {
	"📉 Kanal Stabilitesi & Pozitif Büyüme İvmesi": "1. 3 ay içerisinde persantil kanal kaybının durdurulması ve kanal stabilitesinin yakalanması.\n2. Yaşa ve cinsiyete uygun günlük kalori ve protein ihtiyacının karşılanması.\n3. Büyüme hızının pozitif ivmeye geçmesi.",
	"🍔 Kilo Artışının Dengelenmesi & Fiziksel Aktivite": "1. Kilo artış hızının boy uzamasıyla dengelenmesi (VKİ Z-skorunun normale yaklaşması).\n2. Günlük fiziksel aktivite süresinin en az 60 dakikaya çıkarılması.\n3. Şekerli içecek ve rafine gıda tüketiminin sıfırlanması.",
	"🥗 Ağırlık Kazanımı & Halsizliğin Giderilmesi": "1. 1 ay içerisinde kilo kaybının durdurulması ve en az %5 ağırlık artışı sağlanması.\n2. Günlük enerji ve sıvı ihtiyacının eksiksiz karşılanması.\n3. Halsizlik ve odaklanma güçlüğünün giderilmesi.",
	"🫁 SpO2 >= %95 & Hava Yolu Rahatlaması": "1. Okul saatlerinde SpO2 düzeyinin %95 ve üzerinde tutulması.\n2. Solunum sıkıntısının 15 dakika içinde kontrol altına alınması.\n3. Öğrencinin inhaler kullanım becerisinin tam olması.",
	"🥜 Alerjen Maruziyetinin %100 Önlenmesi": "1. Okul ortamında alerjenle temasın %100 engellenmesi.\n2. Olası maruziyette 3 dakika içinde acil eylem planının devreye sokulması.\n3. Anafilaktik şok gelişiminin önlenmesi.",
	"🩸 Kan Glikozunun 80-180 mg/dL Hedefinde Tutulması": "1. Okul saatleri süresince kan glikozunun 80-180 mg/dL hedef aralığında tutulması.\n2. Ağır hipoglisemi ve bilinç kaybının önlenmesi.\n3. Öğrencinin bağımsız kan şekeri kontrol ve öğün yönetim becerisinin desteklenmesi.",
	"⚡ Travma / Nöbet Güvenliği & Aspirasyonun Önlenmesi": "1. Nöbet sırasında ve sonrasında fiziksel yaralanma ve kafa travmasının önlenmesi.\n2. Solunum yolu açıklığının ve oksijenasyonun korunması (aspirasyonun önlenmesi).\n3. Nöbet süresinin saniye saniye kayıt altına alınması.",
	"🌡️ Vücut Sıcaklığının 37.0 - 37.5 °C'ye İndirilmesi": "1. Vücut sıcaklığının 30 dakika içinde güvenli sınırlara (37.0 °C - 37.5 °C) düşürülmesi.\n2. Ateşli havale (febril konvülsiyon) riskinin önlenmesi.\n3. Sıvı kaybının (dehidratasyon) telafi edilmesi.",
	"🩹 Ağrı Skorunun Düşürülmesi & Doku İyileşmesi": "1. Ağrı skorunun (Wong-Baker yüz skalası ile) <= 2 seviyesine indirilmesi.\n2. Kanamanın durdurulması ve yara enfeksiyonunun engellenmesi.\n3. Eklem hareket ve fonksiyonunun korunması.",
	"🤕 Somatizasyonun Azaltılması & Sınıfa Uyum": "1. Ağrının organik ve somatizasyon ayrımının yapılması.\n2. Öğrencinin gevşeme ve stresle baş etme teknikleri kazanması.\n3. Revirde kalış süresinin sınırlandırılarak sınıfa adaptasyonun güçlendirilmesi.",
	"🧴 Parazit Eliminasyonu & Bulaşın Sıfırlanması": "1. Okul içi ve sınıf akranlarına bulaşın sıfırlanması.\n2. Doğru hijyen ve medikal parazit eliminasyon protokolünün uygulanması.\n3. Cilt/saç bütünlüğünün iyileştirilmesi.",
	"💤 Düzenli Gece Uykusu & Gün İçi Uyanıklık": "1. Öğrencinin yaş grubuna uygun (günde en az 9-10 saat) kesintisiz gece uykusu alışkanlığı kazanması.\n2. Ders içi uyanıklık ve konsantrasyonun artırılması.\n3. Ekran maruziyetinin uyku öncesinde tamamen kesilmesi.",
	"💧 Sıvı Dengesi & Mukozal Nemliliğin Sağlanması": "1. Oral rehidratasyon ile sıvı açığının 2 saat içinde kapatılması.\n2. Vital bulguların (tansiyon, nabız) stabilize edilmesi.\n3. Halsizliğin giderilmesi.",
}

DEFAULT_NANDA_INTERVENTIONS: dict[str, str] = {
	"📉 Büyüme & Beslenme Takip Protokolü": "1. Revirde 15 günde bir standart koşullarda (sabah aç karnına, ayakkabısız) boy ve kilo ölçümü yapılarak persantil grafiğine işlenir.\n2. Veli ile görüşülerek 3 günlük ayrıntılı beslenme günlüğü tutulması istenir.\n3. Okul saatlerinde beslenme çantası içeriği ve yemekhane tüketimi izlenir; ara öğünlerde ceviz, fındık, kuru meyve ve süt teşvik edilir.\n4. Sınıf öğretmeniyle görüşülerek ders içi odaklanma ve yorgunluk halleri takip edilir.\n5. Çölyak, parazitoz veya demir eksikliği açısından çocuk sağlığı hekimine yönlendirilir.",
	"🍔 Obezite & Sağlıklı Yaşam Destek Protokolü": "1. Ayda bir boy, kilo ve VKİ persantil takibi yapılarak değişim eğrisi çıkarılır.\n2. Beden eğitimi öğretmeni ile iş birliği yapılarak öğrencinin akran zorbalığına maruz kalmadan eğlenceli fiziksel aktivitelere (yürüyüş, ip atlama, takım oyunları) katılımı sağlanır.\n3. Okul kantininde şekerli içecek ve cips yerine ayran, süt, su ve meyve tercih etmesi konusunda sağlık danışmanlığı verilir.\n4. Aileye porsiyon kontrolü ve ekran süresinin kısıtlanması konusunda rehberlik sağlanır.",
	"🥗 Malnütrisyon Beslenme Destek Protokolü": "1. Haftalık ağırlık ölçümü ve günlük sıvı tüketim takibi yapılır.\n2. Okul yönetimi ve rehberlik servisiyle görüşülerek ücretsiz sağlıklı okul beslenme desteği veya beslenme çantası takviyesi planlanır.\n3. Yemek öncesi su içmeyip yemek sırasında ve sonrasında tüketmesi; kalori ve besin değeri yoğun gıdaların seçilmesi anlatılır.\n4. Anemi ve vitamin eksiklikleri açısından revir muayenesi yapılır; veli ile ev içi öğün düzeni koordine edilir.",
	"🫁 Astım Atağı & Bronkodilatör Uygulama Protokolü": "1. Revire başvurduğunda derhal rahat oturur (High-Fowler) pozisyona alınır, yakası gevşetilir, ortam havalandırılır.\n2. SpO2, nabız ve solunum sayısı ölçülür.\n3. Reçeteli hızlı etkili bronkodilatör (Salbutamol inhaler) hazne (spacer) ile doğru teknikle uygulanması sağlanır/gözetlenir.\n4. Beden eğitimi dersleri öncesinde ısınma egzersizleri ve koruyucu inhaler protokolü uygulanır.\n5. Sınıf ortamının tozsuz tutulması sınıf öğretmenine bildirilir.",
	"🥜 Anafilaksi Acil Eylem & Oto-Enjektör Protokolü": "1. Öğrencinin alerjileri yemekhane, kantin ve nöbetçi öğretmen listesine eklenir; yiyecek paylaşımı kesinlikle yasaklanır.\n2. Aile tarafından sağlanan Adrenalin Oto-Enjektörü revir acil dolabında hazır bulundurulur.\n3. Öğretmenlere oto-enjektör uygulama tekniği (uyluk dış yan bölgesine dik açıyla 10 saniye) anlatılır.\n4. Reaksiyon belirtilerinde çocuk asla yalnız bırakılmaz, sırtüstü yatırılıp bacakları 30 cm yükseltilir (şok pozisyonu).",
	"🩸 Hipoglisemi 15-15 Yönetim Protokolü": "1. Öğrencinin öğle yemeği ve spor öncesi/sonrası kan şekeri ölçümü yapması sağlanır.\n2. Hipoglisemi durumunda 15-15 Kuralı uygulanır: Kan şekeri <70 mg/dL ise 15 gram hızlı karbonhidrat (3-4 kesme şeker, yarım kutu meyve suyu) verilir; 15 dakika sonra tekrar ölçülür.\n3. Öğrencinin çantasında ve revirde acil hipoglisemi kiti (glukagon hipokit) hazır bulundurulur.\n4. Sınıf öğretmenlerine hipoglisemi erken uyarı işaretleri öğretilir.",
	"⚡ Nöbet Güvenliği & Koma Pozisyonu Protokolü": "1. Nöbet başladığı an sakin olunur, öğrenci yere yavaşça yatırılır ve başının altına yumuşak mont konur.\n2. Çevredeki sert masa ve eşyalar uzaklaştırılır; gözlük çıkarılır, sıkan yaka açılır.\n3. ASLA ağzına kaşık/bez sokulmaz, çenesi zorla açılmaz, kasılan uzuvları bastırılmaz.\n4. Nöbet bitiminde yan yatış (koma/iyileşme pozisyonu) verilir.\n5. Nöbet başlangıç saati ve süresi saniye saniye kaydedilir.",
	"🌡️ Yüksek Ateş Ilık Kompres & Soğutma Protokolü": "1. Kalın giysiler ve mont çıkarılır; oda 21-22 °C seviyesinde havalandırılır.\n2. Boyun, koltuk altı ve alın bölgesine oda sıcaklığında (ılık) ıslatılmış pamuklu kompres uygulanır (ASLA çok soğuk su, buz veya ispirto sürülmez).\n3. Öğrenciye yudum yudum ılık su içirilir.\n4. Veliye telefonla ulaşılarak hekim reçeteli ateş düşürücü kullanımı için onay ve bilgi alınır.",
	"🩹 R.I.C.E. Yumuşak Doku & Yara Pansuman Protokolü": "1. Kesik ve sıyrıklarda yara steril serum fizyolojik ile yıkanır, antiseptik uygulanıp gazlı bezle kapatılır.\n2. Burkulma ve yumuşak doku travmalarında R.I.C.E. Protokolü uygulanır:\n   * Rest (İstirahat): Eklem hareketsiz bırakılır.\n   * Ice (Buz): Havluya sarılı buz ile 15 dakika soğuk kompres uygulanır.\n   * Compression (Bası): Elastik bandaj ile hafif sıkılıkta sarılır.\n   * Elevation (Yükseltme): Kalp seviyesinin üstüne kaldırılır.\n3. Kırık/çıkık şüphesinde ekstremite çekilmez, mevcut pozisyonda atele alınır.",
	"🤕 Somatizasyon & Anksiyete Sakinleştirme Protokolü": "1. Karın muayenesinde defans, rebaund, ateş kontrolü yapılır; organik akut batın ekarte edilir.\n2. Revirde loş ve sakin bir köşede 15-20 dakika istirahat, derin nefes egzersizi ve bir bardak ılık su sağlanır.\n3. Ağrının sınav günleri veya belirli derslerle ilişkisi takvimlenir.\n4. Okul rehberlik servisi ile vaka değerlendirmesi yapılır.\n5. Veli ile görüşülerek ev içi rutinler sorgulanır.",
	"🧴 Pediküloz & Bulaşıcı Hijyen Protokolü": "1. Öğrenci revirde hassasiyetle ve gizlilik esasına göre muayene edilir, veliye bilgi verilir.\n2. Aileye Sağlık Bakanlığı onaylı medikal pediküloz losyonu ve sık dişli metal tarak kullanımı anlatılır.\n3. Evdeki nevresim ve giysilerin 60 °C sıcaklıkta yıkanması bildirilir.\n4. Sınıfta genel el yıkama ve kişisel hijyen bilgilendirmesi yapılır.",
	"💤 Uyku Hijyeni Rehberliği Protokolü": "1. Uyku öncesi telefon/tablet kullanım saatleri sorgulanır.\n2. Aileye uyku hijyeni rehberi verilir: Uykudan 1 saat önce mavi ışıklı ekranların kapatılması, yatak odasının karanlık ve serin tutulması önerilir.\n3. Burun tıkanıklığı veya horlama varsa KBB muayenesi önerilir.\n4. Gün içinde açık hava fiziksel aktivitesi teşvik edilir.",
}

DEFAULT_NANDA_MONITORINGS: dict[str, str] = {
	"⏱️ 15 Günde Bir Persantil & Beslenme İzlemi": "15 günde bir standart boy/kilo persantil ve Z-skor takibi; ayda bir veli görüşmesi ve beslenme değerlendirmesi.",
	"⏱️ Haftalık Ağırlık & Sıvı Tüketim Kontrolü": "Haftalık ağırlık kontrolü, günlük sıvı tüketim takibi ve ayda bir persantil güncellemesi.",
	"⏱️ Aylık VKİ Persantili & Dönemsel Tansiyon İzlemi": "Ayda bir VKİ persantil takibi ve 3 ayda bir vital (tansiyon/nabız) kontrolü.",
	"⏱️ Atak Sırasında 5-10 Dakikada Bir SpO2 & Solunum İzlemi": "Atak anında 5-10 dakikada bir SpO2, solunum sayısı ve akciğer sesleri takibi.",
	"⏱️ Alerjen Temasında Dakikalık Vital & Döküntü İzlemi": "Maruziyet şüphesinde 3-5 dakikada bir nabız, tansiyon, solunum ve deri döküntüsü izlemi.",
	"⏱️ Günlük Öğün & Egzersiz Öncesi Kan Şekeri İzlemi": "Günlük öğle yemeği ve egzersiz öncesi parmak ucu glikoz takibi.",
	"⏱️ Nöbet Süresi & Postiktal Toparlanma Takibi": "Nöbet anında kronometre ile nöbet süresi, nöbet tipi ve nöbet sonrası bilinç toparlanma takibi.",
	"⏱️ 15-20 Dakikada Bir Ateş & Nabız Kontrolü": "15-20 dakikada bir timpanik veya aksiller ateş ölçümü, nabız ve bilinç takibi.",
	"⏱️ 15 Dakikada Bir Ekstremite Dolaşım Kontrolü": "15 dakikada bir periferik dolaşım kontrolü (parmak rengi, kapiller dolum, nabız, soğukluk).",
	"⏱️ Haftalık Revir Başvuru Sıklığı & Rehberlik İzlemi": "Haftalık revir başvuru sıklığı, ders saatleri korelasyonu ve rehberlik servisi takibi.",
	"⏱️ Tedavinin 7. ve 14. Günlerinde Hijyen Kontrolü": "Tedavi başlangıcından sonra 7. ve 14. günlerde revirde kontrol saç ve cilt taraması.",
	"⏱️ Haftalık Ders Odaklanma & Veli Bildirimi": "Haftalık ders takip formu, sınıfta odaklanma gözlemi ve veli bildirimi.",
}

DEFAULT_NANDA_ESCALATIONS: dict[str, str] = {
	"🚨 Ardışık Kilo Kaybında Çocuk Endokrinolojisine Sevk": "Ardışık 2 ölçümde kilo kaybı devam ederse veya boy persantili -3 SDS altına gerilerse derhal Çocuk Sağlığı ve Çocuk Endokrinolojisi Uzmanına sevk.",
	"🚨 VKİ > +3 SDS veya Hipertansiyonda Uzman Sevki": "VKİ Z-skoru > +3 SDS olması, persantil artışının durdurulamaması veya tansiyon yüksekliği halinde Çocuk Endokrinolojisi ve Diyetisyene yönlendirme.",
	"🚨 İnhalere Rağmen SpO2 <%92 Durumunda 112 Acil Ambulans": "İnhaler uygulamasına rağmen 15 dakikada düzelmeme, SpO2 <%92, konuşurken duraklama, interkostal çekilme veya dudaklarda siyanoz durumunda derhal 112 Acil Ambulans ve veli bildirimi.",
	"🚨 Anafilakside Derhal Adrenalin ve 112 Acil Çağrısı": "Dudak/boğazda şişme, ses kısıklığı, stridor, nefes darlığı veya bayılma hissi başladığı an derhal Adrenalin oto-enjektörü uygulanır ve beklemeden 112 Acil Ambulans çağrılır.",
	"🚨 Bilinç Kaybında Ağızdan Alım Yok, IM Glukagon & 112 Acil": "Bilinç bulanıklığı, yutkunamama veya bayılma durumunda ağızdan hiçbir şey verilmez, derhal İntramüsküler Glukagon uygulanır ve 112 Acil Ambulans çağrılır.",
	"🚨 Nöbet > 5 Dakika Sürerse Derhal 112 Acil Ambulans": "Nöbet 5 dakikadan uzun sürerse, arka arkaya tekrarlarsa, kafa travması eşlik ederse veya nöbet sonrası solunum güçlüğü devam ederse derhal 112 Acil Ambulans ve veli bildirimi.",
	"🚨 Ateş > 38.8 °C, Ense Sertliği veya Havalede Acil Sevk": "Ateşin 38.8 °C üzerine çıkması, kontrolsüz titreme, morarma, ense sertliği veya bilinç bulanıklığı görülmesi halinde 112 Acil ve acil poliklinik sevki.",
	"🚨 Deformite, Kırık Şüphesi veya Durdurulamayan Kanamada 112 Acil": "Şekil bozukluğu, kemik krepitasyonu, durdurulamayan arteriyel kanama veya derin kesilerde derhal 112 Acil / Ortopedi sevki.",
	"🚨 Uykudan Uyandıran Ağrı veya Kilo Kaybında Çocuk Nöroloji/Gastro": "Ağrının gece uykudan uyandırması, kusma, kilo kaybı veya görme bozukluğu eşlik etmesi halinde Çocuk Nörolojisi / Gastroenteroloji hekimine sevk.",
	"🚨 İkincil Cilt Enfeksiyonunda Cildiye Polikliniğine Sevk": "İkincil bakteriyel enfeksiyon (ciltte akıntı, impetigo) veya dirençli parazitoz durumunda Aile Hekimi / Cildiye Polikliniğine sevk.",
	"🚨 Uyku Apnesi veya Narkolepsi Şüphesinde Çocuk Nörolojiye Sevk": "Gündüz aşırı uyku hali (narkolepsi şüphesi) veya gece solunum durması (uyku apnesi) şüphesinde Çocuk Nörolojisi ve Uyku Polikliniğine sevk.",
	"🚨 Sıvı Kaybında Ağızdan Sıvı Alamama Halinde Acil Sevk": "Oral alımın tamamen durması, inatçı kusma, göz kürelerinde belirgin çökme ve letarji halinde derhal 112 Acil ve Çocuk Acil Servise sevk.",
}

SECTION_DEFAULTS: dict[str, dict[str, str]] = {
	"diagnoses": DEFAULT_NANDA_DIAGNOSES,
	"facts": DEFAULT_NANDA_FACTS,
	"goals": DEFAULT_NANDA_GOALS,
	"interventions": DEFAULT_NANDA_INTERVENTIONS,
	"monitorings": DEFAULT_NANDA_MONITORINGS,
	"escalations": DEFAULT_NANDA_ESCALATIONS,
}

SECTION_SETTING_KEYS: dict[str, str] = {
	"diagnoses": "custom_cp_diagnoses",
	"facts": "custom_cp_facts",
	"goals": "custom_cp_goals",
	"interventions": "custom_cp_interventions",
	"monitorings": "custom_cp_monitorings",
	"escalations": "custom_cp_escalations",
}


def get_care_plan_section_templates(section: str) -> dict[str, str]:
	default_dict = SECTION_DEFAULTS.get(section, {})
	setting_key = SECTION_SETTING_KEYS.get(section)
	if not setting_key:
		return copy.deepcopy(default_dict)
	settings = get_app_settings()
	raw = settings.get(setting_key, "")
	if raw and raw.strip():
		try:
			parsed = json.loads(raw)
			if isinstance(parsed, dict) and parsed:
				return parsed
		except Exception:
			pass
	return copy.deepcopy(default_dict)


def save_care_plan_section_templates(section: str, templates: dict[str, str]) -> None:
	setting_key = SECTION_SETTING_KEYS.get(section)
	if setting_key:
		save_app_settings({
			setting_key: json.dumps(templates, ensure_ascii=False, indent=2)
		})


def reset_care_plan_section_templates(section: str | None = None) -> None:
	if section:
		setting_key = SECTION_SETTING_KEYS.get(section)
		if setting_key:
			save_app_settings({setting_key: ""})
	else:
		updates = {k: "" for k in SECTION_SETTING_KEYS.values()}
		save_app_settings(updates)



def get_app_settings() -> dict[str, str]:
	settings = dict(DEFAULT_APP_SETTINGS)
	try:
		with connect_db() as connection:
			rows = connection.execute("SELECT key, value FROM app_settings").fetchall()
			for r in rows:
				if r["value"] is not None:
					settings[r["key"]] = str(r["value"])
	except Exception:
		pass
	return settings


def save_app_settings(updates: dict[str, str]) -> None:
	with connect_db() as connection:
		for k, v in updates.items():
			connection.execute(
				"INSERT INTO app_settings(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
				(k, str(v) if v is not None else ""),
			)


def import_tr_locations_to_db(force: bool = False) -> int:
	if not TR_LOCATIONS_PATH.exists():
		return 0
	with connect_db() as connection:
		existing = connection.execute("SELECT COUNT(*) FROM tr_locations").fetchone()[0]
		if existing > 50000 and not force:
			return int(existing)
		data = json.loads(TR_LOCATIONS_PATH.read_text(encoding="utf-8"))
		rows_to_insert: list[tuple[str, str, str]] = []
		for prov, districts in data.items():
			for dist, villages in districts.items():
				if not villages:
					rows_to_insert.append((prov, dist, "Merkez"))
				else:
					for vil in villages:
						rows_to_insert.append((prov, dist, vil))
		if force:
			connection.execute("DELETE FROM tr_locations")
		connection.executemany(
			"INSERT OR IGNORE INTO tr_locations(province, district, village) VALUES (?, ?, ?)",
			rows_to_insert,
		)
		return len(rows_to_insert)


@st.cache_data(show_spinner=False)
def get_tr_provinces() -> list[str]:
	with connect_db() as connection:
		rows = connection.execute("SELECT DISTINCT province FROM tr_locations ORDER BY province").fetchall()
	if rows:
		return [r["province"] for r in rows]
	if TR_LOCATIONS_PATH.exists():
		data = json.loads(TR_LOCATIONS_PATH.read_text(encoding="utf-8"))
		return sorted(data.keys())
	return ["Amasya", "Ankara", "İstanbul", "İzmir"]


@st.cache_data(show_spinner=False)
def get_tr_districts(province: str) -> list[str]:
	if not province:
		return []
	with connect_db() as connection:
		rows = connection.execute(
			"SELECT DISTINCT district FROM tr_locations WHERE province=? ORDER BY district",
			(province,),
		).fetchall()
	if rows:
		return [r["district"] for r in rows]
	if TR_LOCATIONS_PATH.exists():
		data = json.loads(TR_LOCATIONS_PATH.read_text(encoding="utf-8"))
		return sorted(data.get(province, {}).keys())
	return ["Merkez"]


@st.cache_data(show_spinner=False)
def get_tr_villages(province: str, district: str) -> list[str]:
	if not province or not district:
		return []
	with connect_db() as connection:
		rows = connection.execute(
			"SELECT village FROM tr_locations WHERE province=? AND district=? ORDER BY village",
			(province, district),
		).fetchall()
	if rows:
		return [r["village"] for r in rows]
	if TR_LOCATIONS_PATH.exists():
		data = json.loads(TR_LOCATIONS_PATH.read_text(encoding="utf-8"))
		return sorted(data.get(province, {}).get(district, []))
	return []


def initialize_db() -> None:
	with connect_db() as connection:
		connection.executescript(
			"""
			CREATE TABLE IF NOT EXISTS academic_years (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				name TEXT UNIQUE NOT NULL,
				start_date TEXT NOT NULL,
				end_date TEXT NOT NULL,
				folder_name TEXT NOT NULL,
				is_active INTEGER DEFAULT 0,
				created_at TEXT DEFAULT CURRENT_TIMESTAMP
			);
			CREATE TABLE IF NOT EXISTS app_settings (
				key TEXT PRIMARY KEY,
				value TEXT NOT NULL
			);
			CREATE TABLE IF NOT EXISTS tr_locations (
				id INTEGER PRIMARY KEY,
				province TEXT NOT NULL,
				district TEXT NOT NULL,
				village TEXT NOT NULL,
				UNIQUE(province, district, village)
			);
			CREATE INDEX IF NOT EXISTS idx_tr_loc_prov_dist ON tr_locations(province, district);
			CREATE TABLE IF NOT EXISTS classes (
				id INTEGER PRIMARY KEY,
				name TEXT NOT NULL UNIQUE
			);
			CREATE TABLE IF NOT EXISTS students (
				id INTEGER PRIMARY KEY,
				class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
				person_type TEXT NOT NULL DEFAULT 'Öğrenci',
				student_no TEXT,
				full_name TEXT NOT NULL,
				sex TEXT,
				birth_date TEXT,
				guardian_name TEXT,
				guardian_phone TEXT,
				notes TEXT,
				country TEXT,
				province TEXT,
				district TEXT,
				village TEXT,
				UNIQUE(class_id, full_name)
			);
			CREATE TABLE IF NOT EXISTS visits (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				visit_date TEXT NOT NULL,
				visit_time TEXT NOT NULL,
				complaint TEXT NOT NULL,
				symptoms TEXT,
				observation TEXT,
				action TEXT,
				guardian_contacted INTEGER NOT NULL DEFAULT 0,
				referral TEXT,
				follow_up TEXT
			);
			CREATE TABLE IF NOT EXISTS measurements (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				measured_on TEXT NOT NULL,
				height_cm REAL,
				weight_kg REAL,
				head_circ_cm REAL,
				age_months INTEGER,
				bmi REAL
			);
			CREATE TABLE IF NOT EXISTS lms_reference (
				id INTEGER PRIMARY KEY,
				indicator TEXT NOT NULL,
				sex TEXT NOT NULL,
				age_months INTEGER NOT NULL,
				l REAL NOT NULL,
				m REAL NOT NULL,
				s REAL NOT NULL,
				UNIQUE(indicator, sex, age_months)
			);
			CREATE TABLE IF NOT EXISTS wfa_reference (
				sex TEXT NOT NULL,
				age_days INTEGER NOT NULL,
				l REAL NOT NULL,
				m REAL NOT NULL,
				s REAL NOT NULL,
				source TEXT NOT NULL,
				PRIMARY KEY(sex, age_days)
			);
			CREATE TABLE IF NOT EXISTS vital_readings (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				measured_at TEXT NOT NULL,
				systolic_mmhg REAL,
				diastolic_mmhg REAL,
				pulse_bpm REAL,
				temperature_c REAL,
				spo2_pct REAL,
				glucose_mg_dl REAL,
				glucose_context TEXT,
				notes TEXT
			);
			CREATE TABLE IF NOT EXISTS vital_reference (
				id INTEGER PRIMARY KEY,
				indicator TEXT NOT NULL,
				sex TEXT NOT NULL DEFAULT '*',
				age_min_months INTEGER NOT NULL,
				age_max_months INTEGER NOT NULL,
				ref_low REAL NOT NULL,
				ref_high REAL NOT NULL,
				critical_low REAL,
				critical_high REAL,
				unit TEXT NOT NULL,
				source TEXT NOT NULL,
				UNIQUE(indicator, sex, age_min_months, age_max_months, source)
			);
			CREATE TABLE IF NOT EXISTS lab_reports (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				report_date TEXT NOT NULL,
				file_name TEXT NOT NULL,
				content_type TEXT NOT NULL,
				content BLOB NOT NULL,
				extracted_text TEXT
			);
			CREATE TABLE IF NOT EXISTS lab_results (
				id INTEGER PRIMARY KEY,
				report_id INTEGER NOT NULL REFERENCES lab_reports(id) ON DELETE CASCADE,
				test_name TEXT NOT NULL,
				value REAL NOT NULL,
				unit TEXT,
				ref_low REAL,
				ref_high REAL,
				lab_flag TEXT
			);
			CREATE TABLE IF NOT EXISTS health_profiles (
				student_id INTEGER PRIMARY KEY REFERENCES students(id) ON DELETE CASCADE,
				guardian_name TEXT,
				guardian_phone TEXT,
				allergies TEXT,
				chronic_conditions TEXT,
				medical_history TEXT,
				medications TEXT,
				blood_group TEXT,
				emergency_notes TEXT,
				updated_at TEXT NOT NULL,
				reviewed_by TEXT
			);
			CREATE TABLE IF NOT EXISTS health_forms (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				uploaded_at TEXT NOT NULL,
				file_name TEXT NOT NULL,
				content BLOB NOT NULL,
				extracted_text TEXT,
				review_status TEXT NOT NULL DEFAULT 'İnceleme bekliyor'
			);
			CREATE TABLE IF NOT EXISTS care_plans (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				created_at TEXT NOT NULL,
				reviewed_by TEXT,
				nursing_diagnosis TEXT,
				observed_facts TEXT,
				goals TEXT,
				interventions TEXT,
				monitoring TEXT,
				escalation TEXT,
				status TEXT NOT NULL DEFAULT 'Taslak - uzman incelemesi gerekli'
			);
			CREATE TABLE IF NOT EXISTS school_screenings (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				screening_date TEXT NOT NULL,
				height_cm REAL,
				weight_kg REAL,
				head_circ_cm REAL,
				temperature_c REAL,
				pulse_bpm REAL,
				spo2_pct REAL,
				wears_glasses TEXT,
				vision_right TEXT,
				vision_left TEXT,
				hair_screening TEXT,
				hearing_right TEXT,
				hearing_left TEXT,
				health_behavior TEXT,
				harmful_habits TEXT,
				psychosocial_status TEXT,
				notes TEXT,
				screened_by TEXT,
				created_at TEXT NOT NULL
			);
			CREATE TABLE IF NOT EXISTS student_vaccinations (
				id INTEGER PRIMARY KEY,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				vaccine_code TEXT NOT NULL,
				vaccine_name TEXT NOT NULL,
				dose_name TEXT NOT NULL,
				period_label TEXT NOT NULL,
				period_months INTEGER NOT NULL DEFAULT 0,
				is_done INTEGER NOT NULL DEFAULT 0,
				administered_date TEXT,
				facility TEXT,
				lot_number TEXT,
				notes TEXT,
				updated_at TEXT,
				UNIQUE(student_id, vaccine_code)
			);
			CREATE TABLE IF NOT EXISTS student_medications (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				medication_name TEXT NOT NULL,
				med_type TEXT,
				route TEXT,
				dosage_pattern TEXT,
				dose_amount REAL,
				dose_unit TEXT,
				stock_unit TEXT DEFAULT 'adet',
				meal_relation TEXT,
				scheduled_times TEXT,
				total_volume REAL,
				remaining_volume REAL,
				start_date TEXT NOT NULL,
				duration_days INTEGER DEFAULT 7,
				end_date TEXT,
				active_days TEXT,
				storage_condition TEXT,
				guardian_consent INTEGER DEFAULT 1,
				delivered_by_guardian TEXT,
				delivery_date TEXT,
				expiry_date TEXT,
				special_notes TEXT,
				status TEXT DEFAULT 'Aktif',
				created_at TEXT NOT NULL
			);
			CREATE TABLE IF NOT EXISTS medication_logs (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				medication_id INTEGER NOT NULL REFERENCES student_medications(id) ON DELETE CASCADE,
				student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
				log_date TEXT NOT NULL,
				scheduled_time TEXT NOT NULL,
				given_time TEXT,
				dose_given REAL,
				dose_unit TEXT,
				status TEXT NOT NULL DEFAULT 'Verildi',
				administered_by TEXT,
				notes TEXT,
				created_at TEXT NOT NULL,
				UNIQUE(medication_id, log_date, scheduled_time)
			);
			"""
		)
		student_columns = {row["name"] for row in connection.execute("PRAGMA table_info(students)")}
		if "person_type" not in student_columns:
			connection.execute("ALTER TABLE students ADD COLUMN person_type TEXT NOT NULL DEFAULT 'Öğrenci'")
		for loc_col in ("country", "province", "district", "village"):
			if loc_col not in student_columns:
				connection.execute(f"ALTER TABLE students ADD COLUMN {loc_col} TEXT")

		for def_k, def_v in DEFAULT_APP_SETTINGS.items():
			connection.execute("INSERT OR IGNORE INTO app_settings(key, value) VALUES (?, ?)", (def_k, def_v))

		# Ensure default academic year exists in academic_years table
		try:
			ch_ay = connection.execute("SELECT COUNT(*) FROM academic_years").fetchone()[0]
			if ch_ay == 0:
				connection.execute(
					"INSERT INTO academic_years(name, start_date, end_date, folder_name, is_active) VALUES (?, ?, ?, ?, 1)",
					("2026-2027", "2026-09-01", "2027-06-20", "egitim_yillari/2026_2027")
				)
		except Exception:
			pass

		# Ensure newly introduced fields like harmful_habits are appended if setting already existed
		row_scr = connection.execute("SELECT value FROM app_settings WHERE key='screening_visible_fields'").fetchone()
		if row_scr and row_scr["value"]:
			existing_val = str(row_scr["value"])
			if "harmful_habits" not in existing_val:
				connection.execute("UPDATE app_settings SET value=? WHERE key='screening_visible_fields'", (f"{existing_val},harmful_habits",))

		hp_columns = {row["name"] for row in connection.execute("PRAGMA table_info(health_profiles)")}
		if "medical_history" not in hp_columns:
			connection.execute("ALTER TABLE health_profiles ADD COLUMN medical_history TEXT")
		for scr_col in (
			"wears_glasses",
			"vision_right",
			"vision_left",
			"hair_screening",
			"hearing_right",
			"hearing_left",
			"health_behavior",
			"harmful_habits",
			"psychosocial_status",
			"last_screened_at",
		):
			if scr_col not in hp_columns:
				connection.execute(f"ALTER TABLE health_profiles ADD COLUMN {scr_col} TEXT")

		scr_table_cols = {row["name"] for row in connection.execute("PRAGMA table_info(school_screenings)")}
		if "harmful_habits" not in scr_table_cols:
			connection.execute("ALTER TABLE school_screenings ADD COLUMN harmful_habits TEXT")

		meas_columns = {row["name"] for row in connection.execute("PRAGMA table_info(measurements)")}
		if "head_circ_cm" not in meas_columns:
			connection.execute("ALTER TABLE measurements ADD COLUMN head_circ_cm REAL")

		vital_columns = {row["name"] for row in connection.execute("PRAGMA table_info(vital_readings)")}
		if "spo2_pct" not in vital_columns:
			connection.execute("ALTER TABLE vital_readings ADD COLUMN spo2_pct REAL")

		med_columns = {row["name"] for row in connection.execute("PRAGMA table_info(student_medications)")}
		if "stock_unit" not in med_columns:
			connection.execute("ALTER TABLE student_medications ADD COLUMN stock_unit TEXT DEFAULT 'adet'")

		# Resilient person_type normalization
		connection.execute("UPDATE students SET person_type='Öğrenci' WHERE person_type LIKE '%renci%'")

@st.cache_resource(show_spinner=False)
def ensure_db_initialized() -> bool:
	"""Tek seferlik veritabanı şema ve başlangıç kontrolü (Uygulama açılışında 1 kere çalışır, rerun'larda atlanır)."""
	initialize_db()
	return True


def _slugify_person_name(name: str) -> str:
	table = str.maketrans({
		"İ": "i", "I": "i", "ı": "i", "Ğ": "g", "ğ": "g",
		"Ü": "u", "ü": "u", "Ş": "s", "ş": "s", "Ö": "o", "ö": "o",
		"Ç": "c", "ç": "c",
	})
	cleaned = re.sub(r"[^a-zA-Z0-9]+", "_", (name or "").translate(table)).strip("_").lower()
	return cleaned or "kisi"


def get_all_academic_years() -> list[dict]:
	try:
		with connect_master_db() as connection:
			tbl_check = connection.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='academic_years'").fetchone()
			if not tbl_check:
				return []
			rows = connection.execute("SELECT * FROM academic_years ORDER BY start_date DESC").fetchall()
			return [dict(r) for r in rows]
	except Exception:
		return []


def get_active_academic_year() -> dict:
	default_year = {
		"id": 1,
		"name": "2026-2027",
		"start_date": "2026-09-01",
		"end_date": "2027-06-20",
		"folder_name": "egitim_yillari/2026_2027",
		"is_active": 1,
	}
	sel_id = st.session_state.get("active_academic_year_id")
	try:
		with connect_master_db() as connection:
			tbl_check = connection.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='academic_years'").fetchone()
			if not tbl_check:
				return default_year
			if sel_id is not None:
				row = connection.execute("SELECT * FROM academic_years WHERE id=?", (sel_id,)).fetchone()
				if row:
					return dict(row)
			row = connection.execute("SELECT * FROM academic_years WHERE is_active=1").fetchone()
			if row:
				return dict(row)
			row = connection.execute("SELECT * FROM academic_years ORDER BY start_date DESC LIMIT 1").fetchone()
			if row:
				return dict(row)
	except Exception:
		pass
	return default_year


def set_active_academic_year(year_id: int) -> None:
	with connect_master_db() as connection:
		connection.execute("UPDATE academic_years SET is_active=0")
		connection.execute("UPDATE academic_years SET is_active=1 WHERE id=?", (year_id,))
		row = connection.execute("SELECT * FROM academic_years WHERE id=?", (year_id,)).fetchone()
		if row:
			st.session_state["active_academic_year_id"] = year_id
			# Clear all session caches to guarantee zero cross-year leakage
			for k in list(st.session_state.keys()):
				if any(t in k for t in [
					"active_person_id", "active_class_filter", "students_page",
					"_cls_filter", "_person_select", "roster_table", "shared_person",
					"active_person", "sel_acad_year_switch", "sb_quick_pick_ay",
					"ay_active_year_radio", "target_page"
				]):
					st.session_state.pop(k, None)
			year_dir = APP_DIR / row["folder_name"] / "kisi_verileri"
			year_dir.mkdir(parents=True, exist_ok=True)
			year_db = APP_DIR / row["folder_name"] / "revir_defteri.db"
			if not year_db.exists() and DB_PATH.exists():
				import shutil
				shutil.copy2(DB_PATH, year_db)


def create_academic_year(
	name: str,
	start_date: str,
	end_date: str,
	student_policy: str = "fresh"  # "fresh", "promote", "clone", "exact"
) -> int:
	import shutil
	name = name.strip()
	slug = re.sub(r"[^a-zA-Z0-9]+", "_", name).strip("_")
	folder_rel = f"egitim_yillari/{slug}"
	with connect_master_db() as connection:
		cur = connection.execute(
			"INSERT INTO academic_years(name, start_date, end_date, folder_name, is_active) VALUES (?, ?, ?, ?, 0)",
			(name, start_date, end_date, folder_rel)
		)
		new_id = cur.lastrowid

	year_dir = APP_DIR / folder_rel
	year_dir.mkdir(parents=True, exist_ok=True)
	kisi_dir = year_dir / "kisi_verileri"
	kisi_dir.mkdir(parents=True, exist_ok=True)
	year_db = year_dir / "revir_defteri.db"

	# Copy template from master DB to inherit schemas, WHO tables, TR locations, app settings
	if DB_PATH.exists():
		shutil.copy2(DB_PATH, year_db)

	# Configure year_db according to the chosen student policy
	conn = sqlite3.connect(year_db)
	conn.execute("PRAGMA foreign_keys = ON")
	try:
		if student_policy == "fresh":
			# Sıfırdan Temiz Sayfa: Yeni öğrenciler için sıfır kilometre başlangıç
			for tbl in [
				"medication_logs", "student_medications", "student_vaccinations",
				"school_screenings", "care_plans", "health_forms", "health_profiles",
				"lab_results", "lab_reports", "vital_readings", "measurements",
				"visits", "students", "classes"
			]:
				try:
					conn.execute(f"DELETE FROM {tbl}")
				except Exception:
					pass
		elif student_policy == "clone":
			# Öğrenci ve sınıf kadrosunu koru, revir ve ölçüm kayıtlarını sıfırla
			for tbl in [
				"medication_logs", "school_screenings", "vital_readings",
				"measurements", "visits"
			]:
				try:
					conn.execute(f"DELETE FROM {tbl}")
				except Exception:
					pass
			src_kisi = get_current_person_data_dir()
			if src_kisi.exists():
				for jf in src_kisi.glob("kisi_*.json"):
					try:
						(kisi_dir / jf.name).write_bytes(jf.read_bytes())
					except Exception:
						pass
		elif student_policy == "promote":
			# Sınıf Terfisi & Mezuniyet: Sınıfları bir üst seviyeye aktar
			for tbl in [
				"medication_logs", "school_screenings", "vital_readings",
				"measurements", "visits"
			]:
				try:
					conn.execute(f"DELETE FROM {tbl}")
				except Exception:
					pass
			classes = conn.execute("SELECT id, name FROM classes").fetchall()
			for cid, cname in classes:
				cname_clean = cname.strip()
				m = re.match(r"^(\d+)(.*)$", cname_clean)
				if m:
					g = int(m.group(1))
					sfx = m.group(2)
					if g >= 4:
						new_c = f"Mezun {g}{sfx}"
					else:
						new_c = f"{g + 1}{sfx}"
					try:
						conn.execute("UPDATE classes SET name=? WHERE id=?", (new_c, cid))
					except Exception:
						pass
			src_kisi = get_current_person_data_dir()
			if src_kisi.exists():
				for jf in src_kisi.glob("kisi_*.json"):
					try:
						(kisi_dir / jf.name).write_bytes(jf.read_bytes())
					except Exception:
						pass
		conn.commit()
	finally:
		conn.close()

	return new_id


def update_academic_year(year_id: int, name: str, start_date: str, end_date: str) -> None:
	with connect_master_db() as connection:
		connection.execute(
			"UPDATE academic_years SET name=?, start_date=?, end_date=? WHERE id=?",
			(name.strip(), start_date, end_date, year_id)
		)


def delete_academic_year(year_id: int) -> tuple[bool, str]:
	"""Deletes an academic year, its independent database, and its files permanently."""
	import shutil
	with connect_master_db() as connection:
		all_rows = connection.execute("SELECT * FROM academic_years").fetchall()
		if len(all_rows) <= 1:
			return False, "Sistemde en az bir eğitim öğretim yılı bulunmalıdır; tek kalan eğitim yılı silinemez."

		row = connection.execute("SELECT * FROM academic_years WHERE id=?", (year_id,)).fetchone()
		if not row:
			return False, "Silinmek istenen eğitim öğretim yılı bulunamadı."

		year_name = row["name"]
		folder_name = row["folder_name"]
		was_active = bool(row["is_active"])

		# If deleting the currently active year, activate another available one first
		if was_active or st.session_state.get("active_academic_year_id") == year_id:
			other_row = next((r for r in all_rows if r["id"] != year_id), None)
			if other_row:
				connection.execute("UPDATE academic_years SET is_active=1 WHERE id=?", (other_row["id"],))
				st.session_state["active_academic_year_id"] = other_row["id"]

		connection.execute("DELETE FROM academic_years WHERE id=?", (year_id,))

	# Safely remove year folder from disk
	if folder_name and "egitim_yillari" in folder_name:
		target_dir = APP_DIR / folder_name
		if target_dir.exists():
			try:
				shutil.rmtree(target_dir, ignore_errors=True)
			except Exception:
				pass

	# Clear all person, query and filter session state caches
	for k in list(st.session_state.keys()):
		if any(t in k for t in [
			"active_person_id", "active_class_filter", "students_page",
			"_cls_filter", "_person_select", "roster_table", "shared_person",
			"active_person", "sel_acad_year_switch", "sb_quick_pick_ay",
			"ay_active_year_radio", "target_page", "sel_delete_academic_year",
			"transfer_picked_students"
		]):
			st.session_state.pop(k, None)

	return True, f"'{year_name}' eğitim öğretim yılı, veritabanı ve tüm dosyaları kalıcı olarak silindi."


def get_current_person_data_dir() -> Path:
	ay = get_active_academic_year()
	folder = ay.get("folder_name", "")
	if folder:
		pdir = APP_DIR / folder / "kisi_verileri"
		pdir.mkdir(parents=True, exist_ok=True)
		return pdir
	return PERSON_DATA_DIR


def reset_current_year_students_and_records() -> None:
	"""Reset students, classes, visits, measurements, screenings for CURRENT academic year only."""
	with connect_db() as conn:
		for tbl in [
			"medication_logs", "student_medications", "student_vaccinations",
			"school_screenings", "care_plans", "health_forms", "health_profiles",
			"lab_results", "lab_reports", "vital_readings", "measurements",
			"visits", "students", "classes"
		]:
			try:
				conn.execute(f"DELETE FROM {tbl}")
			except Exception:
				pass
	pdir = get_current_person_data_dir()
	if pdir.exists():
		for f in pdir.glob("kisi_*.json"):
			try:
				f.unlink()
			except Exception:
				pass


def transfer_students_from_year(src_year_id: int, student_ids: list[int], target_class_id: int | None = None) -> int:
	"""Transfers selected students from another academic year into the current academic year."""
	all_years = get_all_academic_years()
	src_ay = next((y for y in all_years if y["id"] == src_year_id), None)
	if not src_ay:
		return 0

	src_db_path = APP_DIR / src_ay["folder_name"] / "revir_defteri.db"
	src_kisi_dir = APP_DIR / src_ay["folder_name"] / "kisi_verileri"
	if not src_db_path.exists():
		return 0

	cur_pdir = get_current_person_data_dir()
	cur_pdir.mkdir(parents=True, exist_ok=True)

	transferred_count = 0
	src_conn = sqlite3.connect(src_db_path)
	src_conn.row_factory = sqlite3.Row
	try:
		with connect_db() as cur_conn:
			classes_cache: dict[str, int] = {}
			for r in cur_conn.execute("SELECT id, name FROM classes").fetchall():
				classes_cache[r["name"]] = r["id"]

			for sid in student_ids:
				srow = src_conn.execute(
					"SELECT s.*, c.name as class_name FROM students s JOIN classes c ON c.id=s.class_id WHERE s.id=?",
					(sid,)
				).fetchone()
				if not srow:
					continue

				c_name = srow["class_name"]
				if target_class_id:
					final_class_id = target_class_id
				else:
					if c_name not in classes_cache:
						cur = cur_conn.execute("INSERT INTO classes(name) VALUES (?)", (c_name,))
						classes_cache[c_name] = cur.lastrowid
					final_class_id = classes_cache[c_name]

				cur_conn.execute(
					"""INSERT OR REPLACE INTO students(
						id, class_id, person_type, student_no, full_name, sex,
						birth_date, guardian_name, guardian_phone, notes,
						country, province, district, village
					) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						srow["id"], final_class_id, srow["person_type"], srow["student_no"],
						srow["full_name"], srow["sex"], srow["birth_date"],
						srow["guardian_name"], srow["guardian_phone"], srow["notes"],
						srow["country"], srow["province"], srow["district"], srow["village"]
					)
				)

				hp = src_conn.execute("SELECT * FROM health_profiles WHERE student_id=?", (sid,)).fetchone()
				if hp:
					hp_dict = dict(hp)
					cols = list(hp_dict.keys())
					placeholders = ", ".join(["?"] * len(cols))
					col_names = ", ".join(cols)
					cur_conn.execute(
						f"INSERT OR REPLACE INTO health_profiles({col_names}) VALUES ({placeholders})",
						list(hp_dict.values())
					)

				if src_kisi_dir.exists():
					for f in src_kisi_dir.glob(f"kisi_{int(sid):04d}_*.json"):
						try:
							(cur_pdir / f.name).write_bytes(f.read_bytes())
						except Exception:
							pass

				transferred_count += 1
	finally:
		src_conn.close()

	return transferred_count


def format_file_size(num_bytes: int) -> str:
	if num_bytes < 1024:
		return f"{num_bytes} B"
	elif num_bytes < 1024 * 1024:
		return f"{num_bytes / 1024:.1f} KB"
	else:
		return f"{num_bytes / (1024 * 1024):.1f} MB"


def get_person_data_file_path(person_id: int, full_name: str | None = None) -> Path:
	pdir = get_current_person_data_dir()
	pdir.mkdir(parents=True, exist_ok=True)
	if not full_name:
		with connect_db() as connection:
			row = connection.execute("SELECT full_name FROM students WHERE id=?", (int(person_id),)).fetchone()
			full_name = row["full_name"] if row else f"kisi_{person_id}"
	slug = _slugify_person_name(full_name)
	return pdir / f"kisi_{int(person_id):04d}_{slug}.json"


def delete_person_data_file(person_id: int) -> None:
	pdir = get_current_person_data_dir()
	if not pdir.exists():
		return
	for old_file in pdir.glob(f"kisi_{int(person_id):04d}_*.json"):
		try:
			old_file.unlink()
		except Exception:
			pass
	legacy_file = pdir / f"kisi_{int(person_id)}.json"
	if legacy_file.exists():
		try:
			legacy_file.unlink()
		except Exception:
			pass


def save_person_data_file(person_id: int) -> Path | None:
	pdir = get_current_person_data_dir()
	pdir.mkdir(parents=True, exist_ok=True)
	with connect_db() as connection:
		stu = connection.execute(
			"""SELECT s.id, s.person_type, c.name AS class_name, s.student_no, s.full_name,
					  s.sex, s.birth_date, s.guardian_name, s.guardian_phone, s.notes,
					  s.country, s.province, s.district, s.village
			   FROM students s JOIN classes c ON c.id = s.class_id WHERE s.id = ?""",
			(int(person_id),),
		).fetchone()
		if not stu:
			delete_person_data_file(person_id)
			return None

		hp = connection.execute(
			"""SELECT * FROM health_profiles WHERE student_id = ?""",
			(int(person_id),),
		).fetchone()

		visits = [
			dict(r)
			for r in connection.execute(
				"""SELECT id, visit_date, visit_time, complaint, symptoms, observation,
						  action, guardian_contacted, referral, follow_up
				   FROM visits WHERE student_id = ?
				   ORDER BY visit_date DESC, visit_time DESC, id DESC""",
				(int(person_id),),
			).fetchall()
		]

		measurements = [
			dict(r)
			for r in connection.execute(
				"""SELECT id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi
				   FROM measurements WHERE student_id = ?
				   ORDER BY measured_on DESC, id DESC""",
				(int(person_id),),
			).fetchall()
		]

		vitals = [
			dict(r)
			for r in connection.execute(
				"""SELECT id, measured_at, systolic_mmhg, diastolic_mmhg, pulse_bpm,
						  temperature_c, spo2_pct, glucose_mg_dl, glucose_context, notes
				   FROM vital_readings WHERE student_id = ?
				   ORDER BY measured_at DESC, id DESC""",
				(int(person_id),),
			).fetchall()
		]

		screenings = [
			dict(r)
			for r in connection.execute(
				"""SELECT id, screening_date, height_cm, weight_kg, head_circ_cm,
						  temperature_c, pulse_bpm, spo2_pct, wears_glasses,
						  vision_right, vision_left, hair_screening, hearing_right,
						  hearing_left, health_behavior, harmful_habits, psychosocial_status, notes,
						  screened_by, created_at
				   FROM school_screenings WHERE student_id = ?
				   ORDER BY screening_date DESC, id DESC""",
				(int(person_id),),
			).fetchall()
		]

		care_plans = [
			dict(r)
			for r in connection.execute(
				"""SELECT id, created_at, reviewed_by, nursing_diagnosis, observed_facts,
						  goals, interventions, monitoring, escalation, status
				   FROM care_plans WHERE student_id = ?
				   ORDER BY created_at DESC, id DESC""",
				(int(person_id),),
			).fetchall()
		]

		lab_results = [
			dict(r)
			for r in connection.execute(
				"""SELECT x.id, r.report_date, r.file_name, x.test_name, x.value,
						  x.unit, x.ref_low, x.ref_high, x.lab_flag
				   FROM lab_results x JOIN lab_reports r ON r.id = x.report_id
				   WHERE r.student_id = ?
				   ORDER BY r.report_date DESC, x.id DESC""",
				(int(person_id),),
			).fetchall()
		]

		vaccinations = [
			dict(r)
			for r in connection.execute(
				"""SELECT vaccine_code, vaccine_name, dose_name, period_label,
						  period_months, is_done, administered_date, facility, lot_number, notes, updated_at
				   FROM student_vaccinations WHERE student_id = ?
				   ORDER BY period_months, id""",
				(int(person_id),),
			).fetchall()
		]

		medications = [
			dict(r)
			for r in connection.execute(
				"""SELECT * FROM student_medications WHERE student_id = ? ORDER BY id DESC""",
				(int(person_id),),
			).fetchall()
		]

	target_path = get_person_data_file_path(int(person_id), stu["full_name"])
	for old_file in pdir.glob(f"kisi_{int(person_id):04d}_*.json"):
		if old_file != target_path:
			try:
				old_file.unlink()
			except Exception:
				pass

	payload = {
		"person_id": int(stu["id"]),
		"full_name": stu["full_name"],
		"updated_at": datetime.now().isoformat(timespec="seconds"),
		"profile": dict(stu),
		"health_profile": dict(hp) if hp else {},
		"visits": visits,
		"measurements": measurements,
		"vital_readings": vitals,
		"screenings": screenings,
		"vaccinations": vaccinations,
		"lab_results": lab_results,
		"care_plans": care_plans,
		"medications": medications,
	}
	target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
	return target_path


def load_person_data_file(person_id: int) -> dict:
	with connect_db() as connection:
		row = connection.execute("SELECT full_name FROM students WHERE id=?", (int(person_id),)).fetchone()
	if not row:
		return {}
	fpath = get_person_data_file_path(int(person_id), row["full_name"])
	if not fpath.exists():
		save_person_data_file(int(person_id))
	try:
		data = json.loads(fpath.read_text(encoding="utf-8"))
		if int(data.get("person_id", -1)) != int(person_id):
			save_person_data_file(int(person_id))
			data = json.loads(fpath.read_text(encoding="utf-8"))
		return data
	except Exception:
		save_person_data_file(int(person_id))
		try:
			return json.loads(fpath.read_text(encoding="utf-8"))
		except Exception:
			return {}


def apply_person_data_file_to_db(person_id: int, data: dict) -> None:
	if not data or int(data.get("person_id", person_id)) != int(person_id):
		raise ValueError("Veri dosyasındaki kişi kimliği (person_id) seçili kişiyle uyuşmuyor.")
	prof = data.get("profile") or {}
	hp = data.get("health_profile") or {}
	now_iso = datetime.now().isoformat(timespec="minutes")
	with connect_db() as connection:
		if prof.get("class_name"):
			connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (str(prof["class_name"]).strip(),))
			cid_row = connection.execute("SELECT id FROM classes WHERE name=?", (str(prof["class_name"]).strip(),)).fetchone()
			if cid_row:
				connection.execute("UPDATE students SET class_id=? WHERE id=?", (cid_row["id"], int(person_id)))
		if prof:
			connection.execute(
				"""UPDATE students SET
					 student_no=COALESCE(?, student_no),
					 sex=?,
					 birth_date=?,
					 guardian_name=?,
					 guardian_phone=?,
					 notes=?,
					 country=COALESCE(?, country),
					 province=COALESCE(?, province),
					 district=COALESCE(?, district),
					 village=?
				   WHERE id=?""",
				(
					prof.get("student_no"),
					prof.get("sex"),
					prof.get("birth_date"),
					prof.get("guardian_name"),
					prof.get("guardian_phone"),
					prof.get("notes"),
					prof.get("country"),
					prof.get("province"),
					prof.get("district"),
					prof.get("village"),
					int(person_id),
				),
			)
		if hp is not None:
			connection.execute(
				"""INSERT INTO health_profiles(
					 student_id, guardian_name, guardian_phone, allergies, chronic_conditions,
					 medical_history, medications, blood_group, emergency_notes, updated_at, reviewed_by,
					 wears_glasses, vision_right, vision_left, hair_screening, hearing_right, hearing_left,
					 health_behavior, harmful_habits, psychosocial_status, last_screened_at
				   ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				   ON CONFLICT(student_id) DO UPDATE SET
					 guardian_name=excluded.guardian_name,
					 guardian_phone=excluded.guardian_phone,
					 allergies=excluded.allergies,
					 chronic_conditions=excluded.chronic_conditions,
					 medical_history=excluded.medical_history,
					 medications=excluded.medications,
					 blood_group=excluded.blood_group,
					 emergency_notes=excluded.emergency_notes,
					 updated_at=excluded.updated_at,
					 reviewed_by=excluded.reviewed_by,
					 wears_glasses=excluded.wears_glasses,
					 vision_right=excluded.vision_right,
					 vision_left=excluded.vision_left,
					 hair_screening=excluded.hair_screening,
					 hearing_right=excluded.hearing_right,
					 hearing_left=excluded.hearing_left,
					 health_behavior=excluded.health_behavior,
					 harmful_habits=excluded.harmful_habits,
					 psychosocial_status=excluded.psychosocial_status,
					 last_screened_at=excluded.last_screened_at""",
				(
					int(person_id),
					hp.get("guardian_name") or prof.get("guardian_name"),
					hp.get("guardian_phone") or prof.get("guardian_phone"),
					hp.get("allergies"),
					hp.get("chronic_conditions"),
					hp.get("medical_history"),
					hp.get("medications"),
					hp.get("blood_group"),
					hp.get("emergency_notes"),
					now_iso,
					hp.get("reviewed_by"),
					hp.get("wears_glasses"),
					hp.get("vision_right"),
					hp.get("vision_left"),
					hp.get("hair_screening"),
					hp.get("hearing_right"),
					hp.get("hearing_left"),
					hp.get("health_behavior"),
					hp.get("harmful_habits"),
					hp.get("psychosocial_status"),
					hp.get("last_screened_at"),
				),
			)
		if "visits" in data and isinstance(data["visits"], list):
			connection.execute("DELETE FROM visits WHERE student_id=?", (int(person_id),))
			for v in data["visits"]:
				connection.execute(
					"""INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms,
					   observation, action, guardian_contacted, referral, follow_up)
					   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						int(person_id),
						str(v.get("visit_date") or date.today().isoformat()),
						str(v.get("visit_time") or "09:00"),
						str(v.get("complaint") or "Kontrol"),
						v.get("symptoms"),
						v.get("observation"),
						v.get("action"),
						int(bool(v.get("guardian_contacted"))),
						v.get("referral"),
						v.get("follow_up"),
					),
				)
		if "measurements" in data and isinstance(data["measurements"], list):
			connection.execute("DELETE FROM measurements WHERE student_id=?", (int(person_id),))
			for m in data["measurements"]:
				h_v = numeric_value(m.get("height_cm"))
				w_v = numeric_value(m.get("weight_kg"))
				hc_v = numeric_value(m.get("head_circ_cm"))
				b_v = (w_v / ((h_v / 100.0) ** 2)) if (h_v and w_v and h_v > 0 and w_v > 0) else numeric_value(m.get("bmi"))
				connection.execute(
					"""INSERT INTO measurements(student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
					   VALUES (?, ?, ?, ?, ?, ?, ?)""",
					(
						int(person_id),
						str(m.get("measured_on") or date.today().isoformat()),
						h_v,
						w_v,
						hc_v,
						m.get("age_months"),
						b_v,
					),
				)
		if "vital_readings" in data and isinstance(data["vital_readings"], list):
			connection.execute("DELETE FROM vital_readings WHERE student_id=?", (int(person_id),))
			for vr in data["vital_readings"]:
				connection.execute(
					"""INSERT INTO vital_readings(student_id, measured_at, systolic_mmhg, diastolic_mmhg,
					   pulse_bpm, temperature_c, spo2_pct, glucose_mg_dl, glucose_context, notes)
					   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						int(person_id),
						str(vr.get("measured_at") or datetime.now().isoformat(timespec="minutes")),
						numeric_value(vr.get("systolic_mmhg")),
						numeric_value(vr.get("diastolic_mmhg")),
						numeric_value(vr.get("pulse_bpm")),
						numeric_value(vr.get("temperature_c")),
						numeric_value(vr.get("spo2_pct")),
						numeric_value(vr.get("glucose_mg_dl")),
						vr.get("glucose_context"),
						vr.get("notes"),
					),
				)
		if "screenings" in data and isinstance(data["screenings"], list):
			connection.execute("DELETE FROM school_screenings WHERE student_id=?", (int(person_id),))
			for sc in data["screenings"]:
				connection.execute(
					"""INSERT INTO school_screenings(student_id, screening_date, height_cm, weight_kg,
					   head_circ_cm, temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right,
					   vision_left, hair_screening, hearing_right, hearing_left, health_behavior,
					   harmful_habits, psychosocial_status, notes, screened_by, created_at)
					   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						int(person_id),
						str(sc.get("screening_date") or date.today().isoformat()),
						numeric_value(sc.get("height_cm")),
						numeric_value(sc.get("weight_kg")),
						numeric_value(sc.get("head_circ_cm")),
						numeric_value(sc.get("temperature_c")),
						numeric_value(sc.get("pulse_bpm")),
						numeric_value(sc.get("spo2_pct")),
						sc.get("wears_glasses"),
						sc.get("vision_right"),
						sc.get("vision_left"),
						sc.get("hair_screening"),
						sc.get("hearing_right"),
						sc.get("hearing_left"),
						sc.get("health_behavior"),
						sc.get("harmful_habits"),
						sc.get("psychosocial_status"),
						sc.get("notes"),
						sc.get("screened_by"),
						str(sc.get("created_at") or now_iso),
					),
				)
		if "vaccinations" in data and isinstance(data["vaccinations"], list):
			connection.execute("DELETE FROM student_vaccinations WHERE student_id=?", (int(person_id),))
			for v in data["vaccinations"]:
				connection.execute(
					"""INSERT INTO student_vaccinations(
						student_id, vaccine_code, vaccine_name, dose_name, period_label,
						period_months, is_done, administered_date, facility, lot_number, notes, updated_at
					) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(
						int(person_id),
						v.get("vaccine_code"),
						v.get("vaccine_name") or "",
						v.get("dose_name") or "",
						v.get("period_label") or "",
						int(v.get("period_months") or 0),
						1 if v.get("is_done") else 0,
						v.get("administered_date"),
						v.get("facility"),
						v.get("lot_number"),
						v.get("notes"),
						v.get("updated_at") or now_iso,
					),
				)
	sync_student_measurements(int(person_id))
	save_person_data_file(int(person_id))


def sync_all_person_data_files() -> int:
	cur_dir = get_current_person_data_dir()
	cur_dir.mkdir(parents=True, exist_ok=True)
	with connect_db() as connection:
		rows = connection.execute("SELECT id FROM students ORDER BY id").fetchall()
	valid_ids = {int(r["id"]) for r in rows}
	for existing_file in cur_dir.glob("kisi_*.json"):
		m = re.match(r"^kisi_0*(\d+)(?:_.*)?\.json$", existing_file.name)
		if m and int(m.group(1)) not in valid_ids:
			try:
				existing_file.unlink()
			except Exception:
				pass
	for pid in valid_ids:
		save_person_data_file(pid)
	return len(valid_ids)


def sync_nurse_and_principal_everywhere(nurse_name: str | None = None, principal_name: str | None = None) -> dict[str, Any]:
	"""
	Ayarlarda girilen Okul Revir Sorumlusu / Hemşire ve Okul Müdürü bilgilerini
	tüm sağlık profilleri (health_profiles.reviewed_by), bakım planları (care_plans.reviewed_by),
	kisi_verileri JSON dosyaları ve form alanları ile tarayarak senkronize eder.
	"""
	app_cfg = get_app_settings()
	n_name = (nurse_name if nurse_name is not None else app_cfg.get("nurse_name", "")).strip()
	p_name = (principal_name if principal_name is not None else app_cfg.get("principal_name", "")).strip()

	updates = {}
	if n_name:
		updates["nurse_name"] = n_name
	if p_name:
		updates["principal_name"] = p_name
	if updates:
		save_app_settings(updates)

	hp_updated = 0
	cp_updated = 0

	with connect_db() as connection:
		if n_name:
			cur_hp = connection.execute(
				"UPDATE health_profiles SET reviewed_by = ?, updated_at = ? WHERE reviewed_by IS NULL OR reviewed_by = '' OR reviewed_by != ?",
				(n_name, datetime.now().isoformat(timespec="minutes"), n_name)
			)
			hp_updated = cur_hp.rowcount

			cur_cp = connection.execute(
				"UPDATE care_plans SET reviewed_by = ? WHERE reviewed_by IS NULL OR reviewed_by = '' OR reviewed_by = 'Okul Revir Sorumlusu'",
				(n_name,)
			)
			cp_updated = cur_cp.rowcount

	# Re-sync all JSON files so disk files contain the updated nurse / reviewer info
	json_count = sync_all_person_data_files()

	# Sync current session state keys for any active forms
	if n_name:
		for k in list(st.session_state.keys()):
			if k.startswith("cp_rev_") or k.startswith("dossier_hp_rev_"):
				st.session_state[k] = n_name

	return {
		"nurse_name": n_name,
		"principal_name": p_name,
		"health_profiles_updated": hp_updated,
		"care_plans_updated": cp_updated,
		"json_files_synced": json_count,
	}


HAIR_SCREENING_OPTIONS: list[str] = [
	"Temiz / Parazit Yok",
	"Bit / Sirke Tespit Edildi",
	"İlaç / Şampuan Tedavisi Önerildi",
	"Tedavi Sürecinde (Kontrol)",
	"Yeniden Tarama Gerekiyor",
	"Kepek / Aşırı Pullanma",
	"Egzama / Seboreik Dermatit",
	"Cilt Hassasiyeti / Kaşıntı",
	"Saç Kıran / Alopesi Şüphesi",
	"Kişisel Hijyen Eğitimi Verildi",
]

HEARING_SCREENING_OPTIONS: list[str] = [
	"Normal / İyi Duyuyor",
	"Hafif İşitme Kaybı Şüphesi",
	"Orta / İleri İşitme Kaybı",
	"Fısıltı Testi Negatif (Duyamadı)",
	"KBB Uzmanına Sevk Edildi",
	"İşitme Cihazı Kullanıyor",
	"Biyonik Kulak (Koklear İmplant)",
	"Kulakta Akıntı / İltihap",
	"Kulak Ağrısı Şikayeti",
	"Kulak Kiri (Buşon / Tıkaç)",
	"Sık Kulak Enfeksiyonu Öyküsü",
	"İşitme Testi (Odyometri) Önerildi",
]

HEALTH_BEHAVIOR_OPTIONS: list[str] = [
	"Beslenme ve Hijyen Çok İyi",
	"Düzenli Diş Fırçalama Alışkanlığı Var",
	"Diş Fırçalama / Ağız Bakımı Geliştirilmeli",
	"Yetersiz Su Tüketimi (Günde <1L)",
	"Aşırı Şekerli / Fast-Food Tüketimi",
	"Kahvaltı Yapma Alışkanlığı Yok",
	"Yetersiz / Dengesiz Beslenme",
	"Düzenli Fiziksel Aktivite / Spor Yapıyor",
	"Hareketsiz Yaşam / Egzersiz Yetersiz",
	"El Yıkama / Kişisel Hijyen Çok İyi",
	"El Yıkama / Hijyen Eğitimi Gerekli",
	"Özel Diyet / Alerji Takibi Yapılıyor",
	"Düzenli Uyku Alışkanlığı Var",
]

HARMFUL_HABITS_OPTIONS: list[str] = [
	"Zararlı Alışkanlık Tespit Edilmedi",
	"Aşırı Ekran / Dijital Medya Bağımlılığı",
	"Gece Geç Uyumak / Düzensiz Uyku",
	"Aşırı Kafein / Enerji İçeceği Tüketimi",
	"Tırnak Yeme / Parmak Emme",
	"Tütün / Sigara Kullanımı Şüphesi / Bildirimi",
	"Elektronik Sigara / Puff / Buhar Kullanımı",
	"Alkol Kullanımı Şüphesi / Bildirimi",
	"Uyuşturucu / Madde Bağımlılığı Riski",
	"Kumar / Bahis / Sanal Oyun Harcaması Riski",
	"Düzensiz / Güvensiz Sosyal Medya Kullanımı",
	"Rehberlik / Psikososyal Destek Önerildi",
]


def parse_multiselect_values(val: Any, default_fallback: list[str] | None = None) -> list[str]:
	if not val:
		return list(default_fallback) if default_fallback else []
	if isinstance(val, (list, tuple, set)):
		items = [str(x).strip() for x in val if str(x).strip()]
		return items if items else (list(default_fallback) if default_fallback else [])
	raw_str = str(val).strip()
	if not raw_str or raw_str in ("—", "-", "None"):
		return list(default_fallback) if default_fallback else []
	parts = [p.strip() for p in raw_str.split(",") if p.strip()]
	return parts if parts else (list(default_fallback) if default_fallback else [])


def ensure_options_contain(options: list[str], current_vals: list[str]) -> list[str]:
	res = list(options)
	for v in current_vals:
		if v and v not in res:
			res.append(v)
	return res


def format_multiselect_save_str(selected_items: list[str] | None, extra_note: str | None = None) -> str | None:
	items = [str(x).strip() for x in (selected_items or []) if str(x).strip()]
	if extra_note and extra_note.strip():
		clean_extra = extra_note.strip()
		if clean_extra not in items:
			items.append(clean_extra)
	return ", ".join(items) if items else None


SCREENING_FIELDS_INFO: dict[str, dict[str, str]] = {
	"height_cm": {"name": "Boy (cm)", "cat": "Antropometri", "desc": "Öğrencinin ayakta boy ölçümü", "icon": "📏"},
	"weight_kg": {"name": "Kilo (kg)", "cat": "Antropometri", "desc": "Öğrencinin vücut ağırlığı", "icon": "⚖️"},
	"head_circ_cm": {"name": "Baş Çevresi (cm)", "cat": "Antropometri", "desc": "0-5 yaş baş çevresi takibi", "icon": "👶"},
	"temperature_c": {"name": "Ateş (°C)", "cat": "Vital Bulgular", "desc": "Vücut sıcaklığı ölçümü", "icon": "🌡️"},
	"pulse_bpm": {"name": "Nabız (/dk)", "cat": "Vital Bulgular", "desc": "Dakikadaki kalp atım hızı", "icon": "💓"},
	"spo2_pct": {"name": "Saturasyon (%)", "cat": "Vital Bulgular", "desc": "Kandaki oksijen doygunluğu (SpO2)", "icon": "🫁"},
	"wears_glasses": {"name": "Gözlük Kullanımı", "cat": "Göz & Görme", "desc": "Gözlük takıp takmadığı", "icon": "👓"},
	"vision_right": {"name": "Sağ Göz Görme/Numara", "cat": "Göz & Görme", "desc": "Sağ göz görme değerlendirmesi", "icon": "👁️"},
	"vision_left": {"name": "Sol Göz Görme/Numara", "cat": "Göz & Görme", "desc": "Sol göz görme değerlendirmesi", "icon": "👁️"},
	"hair_screening": {"name": "Saç Taraması", "cat": "Hijyen & Parazit", "desc": "Pediküloz / bit-sirke ve hijyen taraması", "icon": "💇"},
	"hearing_right": {"name": "Sağ Kulak İşitme", "cat": "İşitme Durumu", "desc": "Sağ kulak işitme değerlendirmesi", "icon": "👂"},
	"hearing_left": {"name": "Sol Kulak İşitme", "cat": "İşitme Durumu", "desc": "Sol kulak işitme değerlendirmesi", "icon": "👂"},
	"health_behavior": {"name": "Sağlık Davranışı & Alışkanlıklar", "cat": "Gelişim & Alışkanlık", "desc": "Beslenme, diş fırçalama, hijyen, uyku", "icon": "🍎"},
	"harmful_habits": {"name": "Zararlı Alışkanlıklar & Riskler", "cat": "Gelişim & Risk Faktörleri", "desc": "Dijital bağımlılık, tütün, enerji içecekleri, madde riski vb.", "icon": "⚠️"},
	"psychosocial_status": {"name": "Genel Psikososyal Durum", "cat": "Gelişim & Uyum", "desc": "Sınıf içi uyum, iletişim, dikkat", "icon": "🧠"},
}


def get_screening_visible_fields() -> list[str]:
	cfg = get_app_settings()
	val = cfg.get("screening_visible_fields", "")
	if not val.strip():
		return list(SCREENING_FIELDS_INFO.keys())
	fields = [f.strip() for f in val.split(",") if f.strip() in SCREENING_FIELDS_INFO]
	return fields if fields else list(SCREENING_FIELDS_INFO.keys())


MEDICATION_FIELDS_INFO: dict[str, dict[str, str]] = {
	"med_route": {"name": "Uygulama Yolu (Oral, Solunum, Göz, vb.)", "label": "Uygulama Yolu", "cat": "Uygulama Bilgisi", "desc": "Ağızdan, İnhalasyon, Damla, Pomad, İnsülin vb.", "icon": "🛣️"},
	"dosage_pattern": {"name": "Dozaj Şablonu (3x2, 2x1 vb.)", "label": "Dozaj Şablonu", "cat": "Dozaj & Sıklık", "desc": "Günde kaç kez ve kaçar doz alınacağını belirten şablon", "icon": "⏱️"},
	"dose_unit": {"name": "Doz Birimi (Adet, Kaşık, ml, mg, Puf vb.)", "label": "Doz Birimi", "cat": "Dozaj & Sıklık", "desc": "Ölçek, kaşık, ml, mg, tablet veya damla cinsinden birimi", "icon": "🥄"},
	"meal_relation": {"name": "Açlık / Tokluk Durumu", "label": "Açlık / Tokluk Durumu", "cat": "Uygulama Bilgisi", "desc": "Aç karnına, Tok karnına, Yemek arasında", "icon": "🍽️"},
	"total_volume": {"name": "Teslim Edilen İlaç Miktarı", "label": "Teslim Edilen İlaç Miktarı", "cat": "Stok & Miktar", "desc": "Veliden teslim alınan ilaç miktarı (örn. 100 ml, 30 adet tablet) ve doz takibi", "icon": "📦"},
	"duration_days": {"name": "Kullanım Süresi (Kaç Gün Kullanılacağı)", "label": "Kullanım Süresi", "cat": "Tedavi Takvimi", "desc": "Tedavi süresi (örn. 5 gün, 7 gün, 10 gün) veya Sürekli kullanım", "icon": "🗓️"},
	"active_days": {"name": "Hangi Günler Kullanılacağı", "label": "Hangi Günler Kullanılacağı", "cat": "Tedavi Takvimi", "desc": "Sadece okul günleri (Hafta İçi) veya haftanın her günü", "icon": "📅"},
	"storage_condition": {"name": "Saklama Koşulu (Oda Sıcaklığı / Buzdolabı)", "label": "Saklama Koşulu", "cat": "Muhafaza & Güvenlik", "desc": "Oda Sıcaklığı (15-25°C) veya Buzdolabı (2-8°C / Soğuk Zincir)", "icon": "❄️"},
	"guardian_consent": {"name": "Veli İlaç Teslim Tutanağı & İzin Belgesi", "label": "Veli İzin Belgesi", "cat": "Muhafaza & Güvenlik", "desc": "Velinin yazılı izni, ilacı teslim eden veli ve teslim tarihi", "icon": "📝"},
	"expiry_date": {"name": "Son Kullanma Tarihi (SKT)", "label": "Son Kullanma Tarihi", "cat": "Muhafaza & Güvenlik", "desc": "İlacın son kullanma tarihi ve miad kontrolü", "icon": "⏳"},
	"special_notes": {"name": "Özel Hemşirelik Notu & Yan Etki Uyarısı", "label": "Özel Hemşirelik Notu", "cat": "Klinik Not & Uyarı", "desc": "Özel kullanım talimatları, dikkat edilecek yan etkiler", "icon": "⚠️"},
}


def get_medication_visible_fields() -> list[str]:
	cfg = get_app_settings()
	val = cfg.get("medication_visible_fields", "")
	if not val.strip():
		return list(MEDICATION_FIELDS_INFO.keys())
	fields = [f.strip() for f in val.split(",") if f.strip() in MEDICATION_FIELDS_INFO]
	return fields if fields else list(MEDICATION_FIELDS_INFO.keys())


def add_student_medication(
	student_id: int,
	medication_name: str,
	med_type: str = "Şurup / Süspansiyon",
	route: str = "Ağızdan (Oral)",
	dosage_pattern: str = "2x1",
	dose_amount: float = 1.0,
	dose_unit: str = "Ölçek / Kaşık (5 ml)",
	stock_unit: str = "adet",
	meal_relation: str = "Tok Karnına",
	scheduled_times: str = "09:30, 13:30",
	total_volume: float | None = None,
	duration_days: int = 7,
	start_date: str | None = None,
	end_date: str | None = None,
	active_days: str = "Sadece Okul Günleri (Pazartesi - Cuma)",
	storage_condition: str = "Oda Sıcaklığı (15-25 °C)",
	guardian_consent: int = 1,
	delivered_by_guardian: str = "",
	delivery_date: str | None = None,
	expiry_date: str | None = None,
	special_notes: str = "",
	status: str = "Aktif",
) -> int:
	now_iso = datetime.now().isoformat(timespec="seconds")
	if not start_date:
		start_date = date.today().isoformat()
	if not delivery_date:
		delivery_date = start_date
	if not end_date and duration_days > 0:
		try:
			s_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
			end_date = (s_dt + timedelta(days=duration_days - 1)).isoformat()
		except Exception:
			end_date = None

	rem_vol = total_volume
	with connect_db() as connection:
		cursor = connection.execute(
			"""INSERT INTO student_medications(
				student_id, medication_name, med_type, route, dosage_pattern, dose_amount,
				dose_unit, stock_unit, meal_relation, scheduled_times, total_volume, remaining_volume,
				start_date, duration_days, end_date, active_days, storage_condition,
				guardian_consent, delivered_by_guardian, delivery_date, expiry_date,
				special_notes, status, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				int(student_id), medication_name.strip(), med_type, route, dosage_pattern, float(dose_amount),
				dose_unit, stock_unit.strip(), meal_relation, scheduled_times.strip(), total_volume, rem_vol,
				start_date, int(duration_days), end_date, active_days, storage_condition,
				int(guardian_consent), delivered_by_guardian.strip(), delivery_date, expiry_date,
				special_notes.strip(), status, now_iso
			),
		)
		med_id = cursor.lastrowid
	save_person_data_file(int(student_id))
	return med_id


def update_student_medication(
	medication_id: int | None = None,
	med_id: int | None = None,
	medication_name: str | None = None,
	med_type: str | None = None,
	route: str | None = None,
	dosage_pattern: str | None = None,
	dose_amount: float | None = None,
	dose_unit: str | None = None,
	stock_unit: str | None = None,
	meal_relation: str | None = None,
	scheduled_times: str | None = None,
	total_volume: float | None = None,
	remaining_volume: float | None = None,
	start_date: str | None = None,
	duration_days: int | None = None,
	end_date: str | None = None,
	active_days: str | None = None,
	storage_condition: str | None = None,
	guardian_consent: int | None = None,
	delivered_by_guardian: str | None = None,
	delivery_date: str | None = None,
	expiry_date: str | None = None,
	special_notes: str | None = None,
	status: str | None = None,
) -> None:
	actual_id = medication_id if medication_id is not None else med_id
	if actual_id is None:
		return
	stu_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT * FROM student_medications WHERE id = ?", (actual_id,)).fetchone()
		if not row:
			return
		stu_id = int(row["student_id"])

		up_name = medication_name.strip() if medication_name is not None else row["medication_name"]
		up_type = med_type if med_type is not None else row["med_type"]
		up_route = route if route is not None else row["route"]
		up_pattern = dosage_pattern if dosage_pattern is not None else row["dosage_pattern"]
		up_dose = float(dose_amount) if dose_amount is not None else row["dose_amount"]
		up_dunit = dose_unit if dose_unit is not None else row["dose_unit"]
		up_sunit = stock_unit.strip() if stock_unit is not None else (row["stock_unit"] or "adet")
		up_meal = meal_relation if meal_relation is not None else row["meal_relation"]
		up_times = scheduled_times.strip() if scheduled_times is not None else row["scheduled_times"]
		up_tot = total_volume if total_volume is not None else row["total_volume"]
		up_rem = remaining_volume if remaining_volume is not None else row["remaining_volume"]
		up_start = start_date if start_date is not None else row["start_date"]
		up_dur = int(duration_days) if duration_days is not None else row["duration_days"]
		up_end = end_date if end_date is not None else row["end_date"]
		up_act = active_days if active_days is not None else row["active_days"]
		up_store = storage_condition if storage_condition is not None else row["storage_condition"]
		up_consent = int(guardian_consent) if guardian_consent is not None else row["guardian_consent"]
		up_deliv = delivered_by_guardian.strip() if delivered_by_guardian is not None else row["delivered_by_guardian"]
		up_ddate = delivery_date if delivery_date is not None else row["delivery_date"]
		up_exp = expiry_date if expiry_date is not None else row["expiry_date"]
		up_notes = special_notes.strip() if special_notes is not None else row["special_notes"]
		up_status = status if status is not None else row["status"]

		connection.execute(
			"""UPDATE student_medications SET
				medication_name=?, med_type=?, route=?, dosage_pattern=?, dose_amount=?,
				dose_unit=?, stock_unit=?, meal_relation=?, scheduled_times=?, total_volume=?, remaining_volume=?,
				start_date=?, duration_days=?, end_date=?, active_days=?, storage_condition=?,
				guardian_consent=?, delivered_by_guardian=?, delivery_date=?, expiry_date=?,
				special_notes=?, status=?
			   WHERE id = ?""",
			(
				up_name, up_type, up_route, up_pattern, up_dose,
				up_dunit, up_sunit, up_meal, up_times, up_tot, up_rem,
				up_start, up_dur, up_end, up_act, up_store,
				up_consent, up_deliv, up_ddate, up_exp,
				up_notes, up_status, actual_id
			),
		)
	if stu_id:
		save_person_data_file(stu_id)


def delete_student_medication(med_id: int) -> None:
	stu_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT student_id FROM student_medications WHERE id = ?", (med_id,)).fetchone()
		if row:
			stu_id = int(row["student_id"])
		connection.execute("DELETE FROM student_medications WHERE id = ?", (med_id,))
		connection.execute("DELETE FROM medication_logs WHERE medication_id = ?", (med_id,))
	if stu_id:
		save_person_data_file(stu_id)


def log_medication_dose(
	medication_id: int,
	student_id: int,
	log_date: str,
	scheduled_time: str,
	given_time: str,
	dose_given: float,
	dose_unit: str,
	status: str = "Verildi",
	administered_by: str = "Revir Sorumlusu",
	notes: str = "",
) -> None:
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db() as connection:
		prev = connection.execute(
			"SELECT id, status FROM medication_logs WHERE medication_id=? AND log_date=? AND scheduled_time=?",
			(medication_id, log_date, scheduled_time),
		).fetchone()

		connection.execute(
			"""INSERT INTO medication_logs(
				medication_id, student_id, log_date, scheduled_time, given_time,
				dose_given, dose_unit, status, administered_by, notes, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(medication_id, log_date, scheduled_time) DO UPDATE SET
				given_time=excluded.given_time,
				dose_given=excluded.dose_given,
				dose_unit=excluded.dose_unit,
				status=excluded.status,
				administered_by=excluded.administered_by,
				notes=excluded.notes,
				created_at=excluded.created_at""",
			(
				medication_id, student_id, log_date, scheduled_time, given_time,
				dose_given, dose_unit, status, administered_by, notes.strip(), now_iso
			),
		)
		if status == "Verildi" and (not prev or prev["status"] != "Verildi"):
			med_row = connection.execute("SELECT remaining_volume FROM student_medications WHERE id=?", (medication_id,)).fetchone()
			if med_row and med_row["remaining_volume"] is not None:
				new_rem = max(0.0, float(med_row["remaining_volume"]) - float(dose_given))
				connection.execute("UPDATE student_medications SET remaining_volume=? WHERE id=?", (new_rem, medication_id))
	save_person_data_file(student_id)


def undo_medication_dose(log_id_or_med_id: int, log_date: str | None = None, scheduled_time: str | None = None) -> None:
	stu_id = None
	with connect_db() as connection:
		if log_date is not None and scheduled_time is not None:
			log = connection.execute(
				"SELECT * FROM medication_logs WHERE medication_id=? AND log_date=? AND scheduled_time=?",
				(log_id_or_med_id, log_date, scheduled_time),
			).fetchone()
		else:
			log = connection.execute("SELECT * FROM medication_logs WHERE id=?", (log_id_or_med_id,)).fetchone()
		if log:
			log_id = int(log["id"])
			stu_id = int(log["student_id"])
			med_id = int(log["medication_id"])
			was_given = (log["status"] == "Verildi")
			dose = float(log["dose_given"] or 0.0)
			connection.execute("DELETE FROM medication_logs WHERE id=?", (log_id,))
			if was_given and dose > 0:
				med_row = connection.execute("SELECT remaining_volume, total_volume FROM student_medications WHERE id=?", (med_id,)).fetchone()
				if med_row and med_row["remaining_volume"] is not None:
					max_v = float(med_row["total_volume"]) if med_row["total_volume"] is not None else 999999.0
					restored = min(max_v, float(med_row["remaining_volume"]) + dose)
					connection.execute("UPDATE student_medications SET remaining_volume=? WHERE id=?", (restored, med_id))
	if stu_id:
		save_person_data_file(stu_id)


def get_scheduled_doses_for_date(target_date: date) -> list[dict]:
	d_str = target_date.isoformat()
	weekday_idx = target_date.weekday()
	is_weekday = weekday_idx < 5
	weekday_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
	cur_day_name = weekday_names[weekday_idx]

	doses = []
	with connect_db() as connection:
		meds = connection.execute(
			"""SELECT m.*, s.full_name, s.student_no, c.name as class_name,
			          hp.allergies, hp.chronic_conditions
			   FROM student_medications m
			   JOIN students s ON s.id = m.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   LEFT JOIN health_profiles hp ON hp.student_id = s.id
			   WHERE m.status = 'Aktif'
			     AND m.start_date <= ?
			     AND (m.end_date IS NULL OR m.end_date = '' OR m.end_date >= ?)
			   ORDER BY c.name, s.full_name""",
			(d_str, d_str),
		).fetchall()

		for m in meds:
			act_days = m["active_days"] or "Sadece Okul Günleri (Pazartesi - Cuma)"
			applies = False
			if "Her Gün" in act_days:
				applies = True
			elif "Sadece Okul Günleri" in act_days or "Hafta İçi" in act_days:
				applies = is_weekday
			elif cur_day_name in act_days:
				applies = True

			if not applies:
				continue

			raw_times = m["scheduled_times"] or "12:00"
			times = [t.strip() for t in raw_times.split(",") if t.strip()]
			if not times:
				times = ["12:00"]

			for t in times:
				log = connection.execute(
					"SELECT * FROM medication_logs WHERE medication_id=? AND log_date=? AND scheduled_time=?",
					(m["id"], d_str, t),
				).fetchone()

				doses.append({
					"medication_id": int(m["id"]),
					"student_id": int(m["student_id"]),
					"full_name": m["full_name"],
					"student_no": m["student_no"] or "",
					"class_name": m["class_name"] or "—",
					"medication_name": m["medication_name"],
					"med_type": m["med_type"] or "İlaç",
					"route": m["route"] or "Ağızdan (Oral)",
					"dosage_pattern": m["dosage_pattern"] or "1x1",
					"dose_amount": float(m["dose_amount"] or 1.0),
					"dose_unit": m["dose_unit"] or "Ölçek",
					"stock_unit": m["stock_unit"] if ("stock_unit" in m.keys() and m["stock_unit"]) else (m["dose_unit"] or "adet"),
					"meal_relation": m["meal_relation"] or "Fark Etmez",
					"scheduled_time": t,
					"remaining_volume": m["remaining_volume"],
					"total_volume": m["total_volume"],
					"storage_condition": m["storage_condition"] or "Oda Sıcaklığı (15-25 °C)",
					"special_notes": m["special_notes"] or "",
					"allergies": m["allergies"] or "",
					"chronic_conditions": m["chronic_conditions"] or "",
					"expiry_date": m["expiry_date"],
					"is_logged": bool(log),
					"log_id": int(log["id"]) if log else None,
					"log_status": log["status"] if log else "Bekliyor",
					"given_time": log["given_time"] if log else None,
					"administered_by": log["administered_by"] if log else None,
					"log_notes": log["notes"] if log else None,
				})

	doses.sort(key=lambda x: (x["scheduled_time"], x["class_name"], x["full_name"]))
	return doses


def get_all_student_medications(status: str | None = None) -> list[dict]:
	query = """SELECT m.*, s.full_name, s.student_no, c.name as class_name,
	                  hp.allergies, hp.chronic_conditions
	           FROM student_medications m
	           JOIN students s ON s.id = m.student_id
	           LEFT JOIN classes c ON c.id = s.class_id
	           LEFT JOIN health_profiles hp ON hp.student_id = s.id"""
	params = ()
	if status:
		query += " WHERE m.status = ?"
		params = (status,)
	query += " ORDER BY m.id DESC"
	with connect_db() as connection:
		return [dict(r) for r in connection.execute(query, params).fetchall()]


def save_school_screening(
	student_id: int,
	screening_date: str,
	height_cm: float | None = None,
	weight_kg: float | None = None,
	head_circ_cm: float | None = None,
	temperature_c: float | None = None,
	pulse_bpm: float | None = None,
	spo2_pct: float | None = None,
	wears_glasses: str | None = None,
	vision_right: str | None = None,
	vision_left: str | None = None,
	hair_screening: str | list[str] | None = None,
	hearing_right: str | list[str] | None = None,
	hearing_left: str | list[str] | None = None,
	health_behavior: str | list[str] | None = None,
	harmful_habits: str | list[str] | None = None,
	psychosocial_status: str | None = None,
	notes: str | None = None,
	screened_by: str | None = None,
) -> int:
	now_iso = datetime.now().isoformat(timespec="seconds")
	scr_date = str(screening_date or date.today().isoformat())[:10]
	app_cfg = get_app_settings()
	scr_by = (screened_by or app_cfg.get("nurse_name", "")).strip() or None

	clean_hair = format_multiselect_save_str(hair_screening) if isinstance(hair_screening, (list, tuple, set)) else (hair_screening.strip() if hair_screening else None)
	clean_hr = format_multiselect_save_str(hearing_right) if isinstance(hearing_right, (list, tuple, set)) else (hearing_right.strip() if hearing_right else None)
	clean_hl = format_multiselect_save_str(hearing_left) if isinstance(hearing_left, (list, tuple, set)) else (hearing_left.strip() if hearing_left else None)
	clean_beh = format_multiselect_save_str(health_behavior) if isinstance(health_behavior, (list, tuple, set)) else (health_behavior.strip() if health_behavior else None)
	clean_harm = format_multiselect_save_str(harmful_habits) if isinstance(harmful_habits, (list, tuple, set)) else (harmful_habits.strip() if harmful_habits else None)

	with connect_db() as connection:
		# 1. Insert into school_screenings
		cursor = connection.execute(
			"""INSERT INTO school_screenings(
				student_id, screening_date, height_cm, weight_kg, head_circ_cm,
				temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right,
				vision_left, hair_screening, hearing_right, hearing_left,
				health_behavior, harmful_habits, psychosocial_status, notes, screened_by, created_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				int(student_id), scr_date, height_cm, weight_kg, head_circ_cm,
				temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right,
				vision_left, clean_hair, clean_hr, clean_hl,
				clean_beh, clean_harm, psychosocial_status, notes, scr_by, now_iso
			),
		)
		screening_id = cursor.lastrowid

		# 2. Update health_profiles with latest screening status
		connection.execute(
			"""INSERT INTO health_profiles(
				student_id, updated_at, wears_glasses, vision_right, vision_left,
				hair_screening, hearing_right, hearing_left, health_behavior,
				harmful_habits, psychosocial_status, last_screened_at, reviewed_by
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(student_id) DO UPDATE SET
				updated_at=excluded.updated_at,
				wears_glasses=COALESCE(excluded.wears_glasses, health_profiles.wears_glasses),
				vision_right=COALESCE(excluded.vision_right, health_profiles.vision_right),
				vision_left=COALESCE(excluded.vision_left, health_profiles.vision_left),
				hair_screening=COALESCE(excluded.hair_screening, health_profiles.hair_screening),
				hearing_right=COALESCE(excluded.hearing_right, health_profiles.hearing_right),
				hearing_left=COALESCE(excluded.hearing_left, health_profiles.hearing_left),
				health_behavior=COALESCE(excluded.health_behavior, health_profiles.health_behavior),
				harmful_habits=COALESCE(excluded.harmful_habits, health_profiles.harmful_habits),
				psychosocial_status=COALESCE(excluded.psychosocial_status, health_profiles.psychosocial_status),
				last_screened_at=excluded.last_screened_at,
				reviewed_by=COALESCE(health_profiles.reviewed_by, excluded.reviewed_by)""",
			(
				int(student_id), now_iso, wears_glasses, vision_right, vision_left,
				clean_hair, clean_hr, clean_hl, clean_beh, clean_harm,
				psychosocial_status, scr_date, scr_by
			),
		)

		# 3. Synchronize measurements if height, weight or head circumference provided
		stu_row = connection.execute("SELECT sex, birth_date FROM students WHERE id = ?", (int(student_id),)).fetchone()
		if height_cm is not None or weight_kg is not None or head_circ_cm is not None:
			h_clean = float(height_cm) if (height_cm is not None and height_cm > 0) else None
			w_clean = float(weight_kg) if (weight_kg is not None and weight_kg > 0) else None
			hc_clean = float(head_circ_cm) if (head_circ_cm is not None and head_circ_cm > 0) else None
			bmi_val = (w_clean / ((h_clean / 100.0) ** 2)) if (h_clean and w_clean) else None
			m_age = age_in_months(stu_row["birth_date"], date.fromisoformat(scr_date)) if (stu_row and stu_row["birth_date"]) else None

			connection.execute(
				"""INSERT INTO measurements(student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
				   VALUES (?, ?, ?, ?, ?, ?, ?)""",
				(int(student_id), scr_date, h_clean, w_clean, hc_clean, m_age, bmi_val),
			)

		# 4. Synchronize vitals if temperature, pulse or spo2 provided
		if temperature_c is not None or pulse_bpm is not None or spo2_pct is not None:
			connection.execute(
				"""INSERT INTO vital_readings(student_id, measured_at, temperature_c, pulse_bpm, spo2_pct, notes)
				   VALUES (?, ?, ?, ?, ?, ?)""",
				(int(student_id), f"{scr_date}T09:30", temperature_c, pulse_bpm, spo2_pct, "Okul Taraması"),
			)

	save_person_data_file(int(student_id))
	return int(screening_id)


def get_student_latest_screening(student_id: int) -> dict[str, Any]:
	with connect_db() as connection:
		row = connection.execute(
			"""SELECT * FROM school_screenings WHERE student_id = ? ORDER BY screening_date DESC, id DESC LIMIT 1""",
			(int(student_id),),
		).fetchone()
		hp = connection.execute(
			"""SELECT * FROM health_profiles WHERE student_id = ?""",
			(int(student_id),),
		).fetchone()
	res = dict(row) if row else {}
	if hp:
		for k in ("wears_glasses", "vision_right", "vision_left", "hair_screening", "hearing_right", "hearing_left", "health_behavior", "harmful_habits", "psychosocial_status", "last_screened_at"):
			if k not in res or res[k] is None:
				res[k] = hp[k]
	return res


def get_class_screening_stats(class_name: str) -> dict[str, Any]:
	with connect_db() as connection:
		students = connection.execute(
			"""SELECT s.id, h.last_screened_at,
					  (SELECT COUNT(*) FROM school_screenings sc WHERE sc.student_id = s.id) as scr_cnt
			   FROM students s
			   JOIN classes c ON c.id = s.class_id
			   LEFT JOIN health_profiles h ON h.student_id = s.id
			   WHERE c.name = ?""",
			(class_name,),
		).fetchall()
	total = len(students)
	screened = sum(1 for s in students if (s["last_screened_at"] or (s["scr_cnt"] and s["scr_cnt"] > 0)))
	pct = round((screened / total * 100), 1) if total > 0 else 0
	return {
		"total": total,
		"screened": screened,
		"pending": total - screened,
		"percent": pct,
		"is_complete": total > 0 and screened == total,
	}


# ---------------------------------------------------------------------------
# T.C. SAĞLIK BAKANLIĞI ÇOCUKLUK DÖNEMİ AŞI TAKVİMİ & BAĞIŞIKLAMA SİSTEMİ
# ---------------------------------------------------------------------------

MOH_VACCINE_SCHEDULE: list[dict[str, Any]] = [
	# 0. Ay / Doğumda
	{"code": "hep_b_1", "vaccine": "Hepatit B", "dose": "1. Doz", "period_months": 0, "period_label": "Doğumda (0. Ay)", "type": "routine", "color": "#2563eb", "desc": "Hepatit B virüsüne karşı doğumda hastanede uygulanır."},
	# 1. Ay
	{"code": "hep_b_2", "vaccine": "Hepatit B", "dose": "2. Doz", "period_months": 1, "period_label": "1. Ayın Sonu", "type": "routine", "color": "#2563eb", "desc": "Hepatit B aşısının 2. dozu (Aile Sağlığı Merkezi)."},
	# 2. Ay
	{"code": "bcg", "vaccine": "BCG (Verem)", "dose": "1. Doz", "period_months": 2, "period_label": "2. Ayın Sonu", "type": "routine", "color": "#16a34a", "desc": "Tüberküloz (verem) hastalığına karşı sol omuzdan uygulanır."},
	{"code": "dabt_ipa_hib_1", "vaccine": "DaBT-İPA-Hib (5'li Karma)", "dose": "1. Doz", "period_months": 2, "period_label": "2. Ayın Sonu", "type": "routine", "color": "#7c3aed", "desc": "Difteri, Boğmaca, Tetanoz, Çocuk Felci, Menenjit (Hib)."},
	{"code": "kpa_1", "vaccine": "KPA (Konjuge Pnömokok)", "dose": "1. Doz", "period_months": 2, "period_label": "2. Ayın Sonu", "type": "routine", "color": "#ea580c", "desc": "Zatürre ve menenjite neden olan pnömokok bakterilerine karşı."},
	# 4. Ay
	{"code": "dabt_ipa_hib_2", "vaccine": "DaBT-İPA-Hib (5'li Karma)", "dose": "2. Doz", "period_months": 4, "period_label": "4. Ayın Sonu", "type": "routine", "color": "#7c3aed", "desc": "5'li karma aşının ikinci dozu."},
	{"code": "kpa_2", "vaccine": "KPA (Konjuge Pnömokok)", "dose": "2. Doz", "period_months": 4, "period_label": "4. Ayın Sonu", "type": "routine", "color": "#ea580c", "desc": "Pnömokok aşısının ikinci dozu."},
	# 6. Ay
	{"code": "dabt_ipa_hib_3", "vaccine": "DaBT-İPA-Hib (5'li Karma)", "dose": "3. Doz", "period_months": 6, "period_label": "6. Ayın Sonu", "type": "routine", "color": "#7c3aed", "desc": "5'li karma aşının üçüncü dozu."},
	{"code": "kpa_3", "vaccine": "KPA (Konjuge Pnömokok)", "dose": "3. Doz", "period_months": 6, "period_label": "6. Ayın Sonu", "type": "routine", "color": "#ea580c", "desc": "Pnömokok aşısının üçüncü dozu."},
	{"code": "hep_b_3", "vaccine": "Hepatit B", "dose": "3. Doz", "period_months": 6, "period_label": "6. Ayın Sonu", "type": "routine", "color": "#2563eb", "desc": "Hepatit B aşısının üçüncü dozu."},
	{"code": "opa_1", "vaccine": "OPA (Oral Polio)", "dose": "1. Doz", "period_months": 6, "period_label": "6. Ayın Sonu", "type": "routine", "color": "#0891b2", "desc": "Ağızdan damla şeklinde canlı çocuk felci aşısı."},
	# 12. Ay (1 Yaş)
	{"code": "kkk_1", "vaccine": "KKK (Kızamık-Kızamıkçık-Kabakulak)", "dose": "1. Doz", "period_months": 12, "period_label": "12. Ay (1 Yaş)", "type": "routine", "color": "#dc2626", "desc": "Kızamık, kızamıkçık ve kabakulak virüslerine karşı."},
	{"code": "kpa_r", "vaccine": "KPA (Konjuge Pnömokok)", "dose": "Pekiştirme (R)", "period_months": 12, "period_label": "12. Ay (1 Yaş)", "type": "routine", "color": "#ea580c", "desc": "Pnömokok pekiştirme (rapel) dozu."},
	{"code": "su_cicegi", "vaccine": "Suçiçeği (Varisella)", "dose": "1. Doz", "period_months": 12, "period_label": "12. Ay (1 Yaş)", "type": "routine", "color": "#d97706", "desc": "Suçiçeği virüsüne karşı koruma."},
	# 18. Ay
	{"code": "dabt_ipa_hib_r", "vaccine": "DaBT-İPA-Hib (5'li Karma)", "dose": "Pekiştirme (R)", "period_months": 18, "period_label": "18. Ayın Sonu", "type": "routine", "color": "#7c3aed", "desc": "5'li karma aşının rapel dozu."},
	{"code": "opa_2", "vaccine": "OPA (Oral Polio)", "dose": "2. Doz", "period_months": 18, "period_label": "18. Ayın Sonu", "type": "routine", "color": "#0891b2", "desc": "Oral polio aşısının ikinci dozu."},
	{"code": "hep_a_1", "vaccine": "Hepatit A", "dose": "1. Doz", "period_months": 18, "period_label": "18. Ayın Sonu", "type": "routine", "color": "#059669", "desc": "Hepatit A (sarılık) virüsüne karşı ilk doz."},
	# 24. Ay (2 Yaş)
	{"code": "hep_a_2", "vaccine": "Hepatit A", "dose": "2. Doz", "period_months": 24, "period_label": "24. Ayın Sonu", "type": "routine", "color": "#059669", "desc": "Hepatit A aşısının ikinci dozu."},
	# İlkokul 1. Sınıf (48. Ay / 4-6 Yaş)
	{"code": "kkk_2", "vaccine": "KKK (Kızamık-Kızamıkçık-Kabakulak)", "dose": "2. Doz (Pekiştirme)", "period_months": 48, "period_label": "İlkokul 1 (48. Ay)", "type": "routine", "color": "#dc2626", "desc": "KKK pekiştirme dozu (48. ayda / ilkokul 1. sınıf çağında uygulanır)."},
	{"code": "dabt_ipa_r", "vaccine": "DaBT-İPA (4'lü Karma)", "dose": "Pekiştirme (R)", "period_months": 48, "period_label": "İlkokul 1 (48. Ay)", "type": "routine", "color": "#7c3aed", "desc": "4'lü karma pekiştirme dozu (48. ayda ASM'de uygulanır)."},
	# 8. Sınıf (13 Yaş / 156. Ay)
	{"code": "td_r", "vaccine": "Td (Erişkin Difteri-Tetanoz)", "dose": "Pekiştirme (R)", "period_months": 156, "period_label": "İlköğretim 8 (13 Yaş)", "type": "routine", "color": "#475569", "desc": "Erişkin tipi difteri ve tetanoz pekiştirme dozu."},
	# Özel / Ek Aşılar (Opsiyonel)
	{"code": "rota_1", "vaccine": "Rotavirüs", "dose": "1. Doz", "period_months": 2, "period_label": "2-3. Ay", "type": "optional", "color": "#db2777", "desc": "Bebeklik çağı ağır ishallerine karşı oral aşı."},
	{"code": "rota_2", "vaccine": "Rotavirüs", "dose": "2. Doz", "period_months": 4, "period_label": "4-5. Ay", "type": "optional", "color": "#db2777", "desc": "Rotavirüs aşısının ikinci dozu."},
	{"code": "meningokok_1", "vaccine": "Meningokok (MenACWY/B)", "dose": "1. Doz", "period_months": 3, "period_label": "3. Ay ve sonrası", "type": "optional", "color": "#9333ea", "desc": "Meningokokal menenjit enfeksiyonlarına karşı."},
	{"code": "meningokok_2", "vaccine": "Meningokok (MenACWY/B)", "dose": "2. Doz / Pekiştirme", "period_months": 12, "period_label": "12. Ay ve sonrası", "type": "optional", "color": "#9333ea", "desc": "Meningokok aşısının rapel dozu."},
	{"code": "hpv", "vaccine": "HPV (İnsan Papilloma)", "dose": "Tam Doz", "period_months": 108, "period_label": "9-14 Yaş", "type": "optional", "color": "#4f46e5", "desc": "HPV ilişkili kanser ve lezyonlara karşı koruyucu aşı."},
	{"code": "influenza", "vaccine": "İnfluenza (Grip)", "dose": "Mevsimsel Doz", "period_months": 6, "period_label": "Her Yıl (Sonbahar)", "type": "optional", "color": "#0d9488", "desc": "Mevsimsel grip aşısı."},
]


def get_student_vaccinations(student_id: int) -> dict[str, dict[str, Any]]:
	with connect_db() as connection:
		rows = connection.execute(
			"""SELECT * FROM student_vaccinations WHERE student_id = ?""",
			(int(student_id),),
		).fetchall()
	res = {}
	for r in rows:
		res[r["vaccine_code"]] = dict(r)
	return res


def save_student_vaccination(
	student_id: int,
	vaccine_code: str,
	is_done: bool,
	administered_date: str | None = None,
	facility: str | None = None,
	lot_number: str | None = None,
	notes: str | None = None,
) -> None:
	vac_meta = next((v for v in MOH_VACCINE_SCHEDULE if v["code"] == vaccine_code), None)
	if not vac_meta:
		return
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db() as connection:
		connection.execute(
			"""INSERT INTO student_vaccinations(
				student_id, vaccine_code, vaccine_name, dose_name, period_label,
				period_months, is_done, administered_date, facility, lot_number, notes, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
			ON CONFLICT(student_id, vaccine_code) DO UPDATE SET
				is_done=excluded.is_done,
				administered_date=excluded.administered_date,
				facility=excluded.facility,
				lot_number=excluded.lot_number,
				notes=excluded.notes,
				updated_at=excluded.updated_at""",
			(
				int(student_id),
				vaccine_code,
				vac_meta["vaccine"],
				vac_meta["dose"],
				vac_meta["period_label"],
				vac_meta["period_months"],
				1 if is_done else 0,
				administered_date,
				facility,
				lot_number,
				notes,
				now_iso,
			),
		)
	save_person_data_file(int(student_id))


def save_bulk_student_vaccinations(student_id: int, records: list[dict[str, Any]]) -> None:
	now_iso = datetime.now().isoformat(timespec="seconds")
	with connect_db() as connection:
		for rec in records:
			v_code = rec.get("vaccine_code")
			vac_meta = next((v for v in MOH_VACCINE_SCHEDULE if v["code"] == v_code), None)
			if not vac_meta:
				continue
			connection.execute(
				"""INSERT INTO student_vaccinations(
					student_id, vaccine_code, vaccine_name, dose_name, period_label,
					period_months, is_done, administered_date, facility, lot_number, notes, updated_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				ON CONFLICT(student_id, vaccine_code) DO UPDATE SET
					is_done=excluded.is_done,
					administered_date=excluded.administered_date,
					facility=excluded.facility,
					lot_number=excluded.lot_number,
					notes=excluded.notes,
					updated_at=excluded.updated_at""",
				(
					int(student_id),
					v_code,
					vac_meta["vaccine"],
					vac_meta["dose"],
					vac_meta["period_label"],
					vac_meta["period_months"],
					1 if rec.get("is_done") else 0,
					rec.get("administered_date"),
					rec.get("facility"),
					rec.get("lot_number"),
					rec.get("notes"),
					now_iso,
				),
			)
	save_person_data_file(int(student_id))


def quick_fill_routine_vaccinations_for_student(student_id: int, birth_date: str | None = None, up_to_age: bool = True) -> int:
	cur_months = age_in_months(birth_date, date.today()) if birth_date else None
	b_date = None
	if birth_date:
		try:
			b_date = date.fromisoformat(str(birth_date)[:10])
		except Exception:
			b_date = None

	now_iso = datetime.now().isoformat(timespec="seconds")
	count = 0
	with connect_db() as connection:
		for v in MOH_VACCINE_SCHEDULE:
			if v["type"] != "routine":
				continue
			if up_to_age and cur_months is not None and v["period_months"] > cur_months:
				continue

			est_date = None
			if b_date:
				try:
					from datetime import timedelta
					est_date = (b_date + timedelta(days=int(v["period_months"] * 30.4375))).isoformat()
					if est_date > date.today().isoformat():
						est_date = date.today().isoformat()
				except Exception:
					est_date = date.today().isoformat()
			else:
				est_date = date.today().isoformat()

			connection.execute(
				"""INSERT INTO student_vaccinations(
					student_id, vaccine_code, vaccine_name, dose_name, period_label,
					period_months, is_done, administered_date, facility, lot_number, notes, updated_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				ON CONFLICT(student_id, vaccine_code) DO UPDATE SET
					is_done=1,
					administered_date=COALESCE(student_vaccinations.administered_date, excluded.administered_date),
					facility=COALESCE(student_vaccinations.facility, excluded.facility),
					updated_at=excluded.updated_at""",
				(
					int(student_id),
					v["code"],
					v["vaccine"],
					v["dose"],
					v["period_label"],
					v["period_months"],
					1,
					est_date,
					"Aile Sağlığı Merkezi (ASM)",
					None,
					"Otomatik takvim onayı",
					now_iso,
				),
			)
			count += 1
	save_person_data_file(int(student_id))
	return count


def clear_student_vaccinations(student_id: int) -> None:
	with connect_db() as connection:
		connection.execute("DELETE FROM student_vaccinations WHERE student_id = ?", (int(student_id),))
	save_person_data_file(int(student_id))


def get_student_vaccination_summary(student_id: int, birth_date: str | None = None) -> dict[str, Any]:
	vac_records = get_student_vaccinations(student_id)
	cur_months = age_in_months(birth_date, date.today()) if birth_date else None

	routine_items = [v for v in MOH_VACCINE_SCHEDULE if v["type"] == "routine"]
	total_routine = len(routine_items)

	if cur_months is not None:
		applicable_routine = [v for v in routine_items if v["period_months"] <= cur_months]
	else:
		applicable_routine = routine_items

	done_routine = sum(1 for v in routine_items if vac_records.get(v["code"], {}).get("is_done") == 1)

	overdue_list = []
	if cur_months is not None:
		for v in applicable_routine:
			rec = vac_records.get(v["code"])
			if not rec or not rec.get("is_done"):
				overdue_list.append(f"{v['vaccine']} ({v['dose']})")

	total_app = len(applicable_routine) if applicable_routine else total_routine
	app_done = sum(1 for v in applicable_routine if vac_records.get(v["code"], {}).get("is_done") == 1)
	pct = round((app_done / total_app * 100), 1) if total_app > 0 else 0

	return {
		"total_routine": total_routine,
		"applicable_total": total_app,
		"completed_count": done_routine,
		"applicable_done_count": app_done,
		"percent": pct,
		"overdue_list": overdue_list,
		"is_up_to_date": (len(overdue_list) == 0 and app_done > 0),
		"records": vac_records,
	}


def build_vaccination_matrix_html(student_id: int, birth_date: str | None = None) -> str:
	vac_summary = get_student_vaccination_summary(student_id, birth_date)
	vac_records = vac_summary["records"]
	cur_months = age_in_months(birth_date, date.today()) if birth_date else None

	periods = [
		(0, "Doğumda"),
		(1, "1. Ay"),
		(2, "2. Ay"),
		(4, "4. Ay"),
		(6, "6. Ay"),
		(12, "12. Ay (1y)"),
		(18, "18. Ay"),
		(24, "24. Ay (2y)"),
		(48, "İlkokul 1 (4y)"),
		(156, "8. Sınıf (13y)"),
	]

	antigen_matrix = [
		("Hepatit B", {0: ("hep_b_1", "I. Doz"), 1: ("hep_b_2", "II. Doz"), 6: ("hep_b_3", "III. Doz")}),
		("BCG (Verem)", {2: ("bcg", "I. Doz")}),
		("DaBT-İPA-Hib (5'li Karma)", {2: ("dabt_ipa_hib_1", "I. Doz"), 4: ("dabt_ipa_hib_2", "II. Doz"), 6: ("dabt_ipa_hib_3", "III. Doz"), 18: ("dabt_ipa_hib_r", "R (IV. Doz)")}),
		("KPA (Pnömokok)", {2: ("kpa_1", "I. Doz"), 4: ("kpa_2", "II. Doz"), 6: ("kpa_3", "III. Doz"), 12: ("kpa_r", "R (Pekiştirme)")}),
		("KKK (Kızamıkçık)", {12: ("kkk_1", "I. Doz"), 48: ("kkk_2", "R (II. Doz)")}),
		("Suçiçeği", {12: ("su_cicegi", "I. Doz")}),
		("OPA (Çocuk Felci)", {6: ("opa_1", "I. Doz"), 18: ("opa_2", "II. Doz")}),
		("Hepatit A", {18: ("hep_a_1", "I. Doz"), 24: ("hep_a_2", "II. Doz")}),
		("DaBT-İPA (4'lü Karma)", {48: ("dabt_ipa_r", "R (Pekiştirme)")}),
		("Td (Difteri-Tetanoz)", {156: ("td_r", "R (Pekiştirme)")}),
	]

	th_html = "".join([f"<th style='text-align:center;padding:7px 5px;background:#0f766e;color:white;font-size:11px;border:1px solid #14b8a6;'>{p[1]}</th>" for p in periods])
	header_row = f"<tr><th style='text-align:left;padding:7px 8px;background:#0f766e;color:white;font-size:12px;border:1px solid #14b8a6;min-width:140px;'>Aşı / Antijen</th>{th_html}</tr>"

	body_rows = []
	for ant_name, dose_map in antigen_matrix:
		td_cells = []
		for p_months, p_lbl in periods:
			if p_months in dose_map:
				vcode, dose_lbl = dose_map[p_months]
				rec = vac_records.get(vcode, {})
				is_done = bool(rec.get("is_done"))
				adm_dt = rec.get("administered_date") or ""

				if is_done:
					dt_txt = f"<br><span style='font-size:9px;color:#166534;font-weight:normal;'>{adm_dt}</span>" if adm_dt else ""
					cell_content = f"<div style='background:#dcfce7;color:#15803d;border:1px solid #86efac;border-radius:4px;padding:3px 2px;text-align:center;font-weight:600;font-size:10px;line-height:1.2;'>✔ {dose_lbl}{dt_txt}</div>"
				elif cur_months is not None and cur_months >= p_months:
					cell_content = f"<div style='background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;border-radius:4px;padding:3px 2px;text-align:center;font-weight:600;font-size:10px;line-height:1.2;'>✖ {dose_lbl}<br><span style='font-size:9px;color:#991b1b;font-weight:normal;'>Eksik Doz</span></div>"
				else:
					cell_content = f"<div style='background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1;border-radius:4px;padding:3px 2px;text-align:center;font-size:10px;line-height:1.2;'>⏳ {dose_lbl}<br><span style='font-size:8px;font-weight:normal;'>Gelecek</span></div>"
			else:
				cell_content = "<div style='text-align:center;color:#cbd5e1;font-size:11px;'>—</div>"
			td_cells.append(f"<td style='padding:5px 3px;border:1px solid #e2e8f0;vertical-align:middle;'>{cell_content}</td>")

		body_rows.append(f"<tr><td style='padding:6px 8px;border:1px solid #e2e8f0;font-weight:600;font-size:12px;background:#f8fafc;'>{ant_name}</td>{''.join(td_cells)}</tr>")

	table_html = f"""
	<div style='overflow-x:auto;margin:12px 0 16px 0;border:1px solid #cbd5e1;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.05);'>
		<table style='width:100%;border-collapse:collapse;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;'>
			<thead>{header_row}</thead>
			<tbody>{''.join(body_rows)}</tbody>
		</table>
	</div>
	<div style='display:flex;gap:14px;font-size:12px;color:#475569;margin-bottom:12px;flex-wrap:wrap;'>
		<div><span style='display:inline-block;width:12px;height:12px;background:#dcfce7;border:1px solid #86efac;border-radius:3px;vertical-align:middle;margin-right:4px;'></span> <b>Yeşil (✔):</b> Yapıldı / Tamamlandı</div>
		<div><span style='display:inline-block;width:12px;height:12px;background:#fee2e2;border:1px solid #fca5a5;border-radius:3px;vertical-align:middle;margin-right:4px;'></span> <b>Kırmızı (✖):</b> Yaşı Gelmiş / Gecikmiş Eksik Doz</div>
		<div><span style='display:inline-block;width:12px;height:12px;background:#f1f5f9;border:1px solid #cbd5e1;border-radius:3px;vertical-align:middle;margin-right:4px;'></span> <b>Gri (⏳):</b> Yaşı Henüz Gelmemiş Gelecek Doz</div>
	</div>
	"""
	return table_html


def set_cell_background(cell, fill_hex: str) -> None:
	tcPr = cell._element.get_or_add_tcPr()
	shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
	tcPr.append(shd)


def set_cell_margins(cell, top: int = 80, bottom: int = 80, left: int = 100, right: int = 100) -> None:
	tcPr = cell._element.get_or_add_tcPr()
	tcMar = parse_xml(f'''
		<w:tcMar {nsdecls("w")}>
			<w:top w:w="{top}" w:type="dxa"/>
			<w:bottom w:w="{bottom}" w:type="dxa"/>
			<w:left w:w="{left}" w:type="dxa"/>
			<w:right w:w="{right}" w:type="dxa"/>
		</w:tcMar>
	''')
	tcPr.append(tcMar)


def set_table_borders(table, color: str = "CBD5E1", sz: str = "4", val: str = "single") -> None:
	tblPr = table._element.xpath('w:tblPr')
	if tblPr:
		borders = parse_xml(f'''
			<w:tblBorders {nsdecls("w")}>
				<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
				<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
				<w:left w:val="none"/>
				<w:right w:val="none"/>
				<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
				<w:insideV w:val="none"/>
			</w:tblBorders>
		''')
		tblPr[0].append(borders)


def add_styled_df_table(doc: docx.Document, df: pd.DataFrame | None, title: str | None = None, empty_msg: str = "Kayıt bulunmuyor.") -> None:
	if title:
		p_t = doc.add_paragraph()
		p_t.paragraph_format.space_before = Pt(8)
		p_t.paragraph_format.space_after = Pt(2)
		r_t = p_t.add_run(title)
		r_t.font.name = "Calibri"
		r_t.font.size = Pt(10.5)
		r_t.bold = True
		r_t.font.color.rgb = RGBColor(15, 118, 110)

	if df is None or df.empty:
		p_e = doc.add_paragraph()
		p_e.paragraph_format.space_after = Pt(4)
		r_e = p_e.add_run(empty_msg)
		r_e.font.name = "Calibri"
		r_e.font.size = Pt(8.5)
		r_e.font.italic = True
		r_e.font.color.rgb = RGBColor(148, 163, 184)
		return

	cols = [str(c).replace("_", " ") for c in df.columns]
	table = doc.add_table(rows=1 + len(df), cols=len(cols))
	table.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(table, color="CBD5E1", sz="4", val="single")

	# Header
	hdr = table.rows[0]
	hdr_trPr = hdr._element.get_or_add_trPr()
	hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

	for c_i, c_name in enumerate(cols):
		c = hdr.cells[c_i]
		set_cell_background(c, "0F766E")
		set_cell_margins(c, top=60, bottom=60, left=70, right=70)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(c_name)
		r.font.name = "Calibri"
		r.font.size = Pt(8)
		r.bold = True
		r.font.color.rgb = RGBColor(255, 255, 255)

	# Rows
	for r_i, (_, row_data) in enumerate(df.iterrows()):
		r = table.rows[1 + r_i]
		bg_col = "F8FAFC" if (r_i % 2 == 1) else "FFFFFF"
		for c_i, val in enumerate(row_data):
			cell = r.cells[c_i]
			if bg_col != "FFFFFF":
				set_cell_background(cell, bg_col)
			set_cell_margins(cell, top=45, bottom=45, left=60, right=60)
			cp = cell.paragraphs[0]
			val_str = "—" if (pd.isna(val) or str(val).strip() == "") else str(val).strip()
			if len(val_str) <= 12 and not any(ch.isalpha() for ch in val_str[:3]):
				cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
			else:
				cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
			r_run = cp.add_run(val_str)
			r_run.font.name = "Calibri"
			r_run.font.size = Pt(7.5)
			r_run.font.color.rgb = RGBColor(30, 41, 59)


def build_vaccination_card_docx(student_id: int) -> bytes:
	app_cfg = get_app_settings()
	inst_top = app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI"
	inst_school = app_cfg.get("school_name") or ""
	inst_nurse = app_cfg.get("nurse_name") or "Okul Revir Sorumlusu"
	inst_princ = app_cfg.get("principal_name") or "Okul Müdürü / İdare"

	student = get_person_by_id(student_id)
	if not student:
		doc = docx.Document()
		doc.add_paragraph("Öğrenci bulunamadı.")
		buf = BytesIO()
		doc.save(buf)
		return buf.getvalue()

	vac_summary = get_student_vaccination_summary(student_id, student["birth_date"])
	vac_records = vac_summary["records"]

	s_age_m = age_in_months(student["birth_date"], date.today()) if student["birth_date"] else None
	s_age_str = f"{s_age_m // 12} yaş {s_age_m % 12} ay" if s_age_m is not None else "—"
	s_sex = "Kız" if student["sex"] == "K" else ("Erkek" if student["sex"] == "E" else "—")
	safe_name = student["full_name"]
	safe_class = student["class_name"] or "—"
	safe_sno = str(student["student_no"] or "—")
	safe_birth = str(student["birth_date"] or "—")

	doc = docx.Document()
	section = doc.sections[0]
	section.orientation = WD_ORIENT.PORTRAIT
	section.page_width = Cm(21.0)
	section.page_height = Cm(29.7)
	section.top_margin = Cm(1.5)
	section.bottom_margin = Cm(1.5)
	section.left_margin = Cm(1.5)
	section.right_margin = Cm(1.5)

	# Header Box
	p_top = doc.add_paragraph()
	p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
	r_top = p_top.add_run(f"T.C. SAĞLIK BAKANLIĞI & {inst_top}")
	r_top.font.name = "Calibri"
	r_top.font.size = Pt(10)
	r_top.font.color.rgb = RGBColor(100, 116, 139)

	if inst_school:
		p_sch = doc.add_paragraph()
		p_sch.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p_sch.paragraph_format.space_before = Pt(0)
		p_sch.paragraph_format.space_after = Pt(2)
		r_sch = p_sch.add_run(inst_school)
		r_sch.font.name = "Calibri"
		r_sch.font.size = Pt(14)
		r_sch.bold = True
		r_sch.font.color.rgb = RGBColor(15, 118, 110)

	p_title = doc.add_paragraph()
	p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_title.paragraph_format.space_before = Pt(2)
	p_title.paragraph_format.space_after = Pt(8)
	r_title = p_title.add_run("💉 ÖĞRENCİ AŞI VE BAĞIŞIKLAMA İZLEM KARTI")
	r_title.font.name = "Calibri"
	r_title.font.size = Pt(13)
	r_title.bold = True
	r_title.font.color.rgb = RGBColor(30, 41, 59)

	# Student Info Box (table)
	info_tbl = doc.add_table(rows=2, cols=3)
	info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(info_tbl, color="CBD5E1", sz="4", val="single")
	info_data = [
		[("Öğrenci Adı Soyadı", safe_name), ("Okul No", safe_sno), ("Sınıf", safe_class)],
		[("Doğum Tarihi / Yaş", f"{safe_birth} ({s_age_str})"), ("Cinsiyet", s_sex), ("Düzenleme Tarihi", date.today().strftime("%d.%m.%Y"))],
	]
	for r_i, r_data in enumerate(info_data):
		for c_i, (lbl, val) in enumerate(r_data):
			cell = info_tbl.cell(r_i, c_i)
			set_cell_background(cell, "F8FAFC")
			set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
			cp = cell.paragraphs[0]
			r1 = cp.add_run(f"{lbl}: ")
			r1.font.name = "Calibri"
			r1.font.size = Pt(9.5)
			r1.font.color.rgb = RGBColor(100, 116, 139)
			r2 = cp.add_run(val)
			r2.font.name = "Calibri"
			r2.font.size = Pt(9.5)
			r2.bold = True
			r2.font.color.rgb = RGBColor(30, 41, 59)

	doc.add_paragraph().paragraph_format.space_after = Pt(4)

	# Summary Bar table
	sum_tbl = doc.add_table(rows=1, cols=3)
	sum_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(sum_tbl, color="CCFBF1", sz="6", val="single")
	sum_data = [
		("Aşı Durumu", f"%{vac_summary['percent']} Tamamlandı"),
		("Yapılan Doz", f"{vac_summary['applicable_done_count']} / {vac_summary['applicable_total']} Doz"),
		("Eksik / Gecikmiş", f"{len(vac_summary['overdue_list'])} Doz"),
	]
	for c_i, (lbl, val) in enumerate(sum_data):
		cell = sum_tbl.cell(0, c_i)
		set_cell_background(cell, "F0FDFA")
		set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
		cp = cell.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r1 = cp.add_run(f"{lbl}: ")
		r1.font.name = "Calibri"
		r1.font.size = Pt(9.5)
		r1.bold = True
		r2 = cp.add_run(val)
		r2.font.name = "Calibri"
		r2.font.size = Pt(9.5)

	doc.add_paragraph().paragraph_format.space_after = Pt(6)

	# Main Vaccination Schedule Table
	v_cols = ["Aşı Adı", "Doz", "Takvim Dönemi", "Uygulama Durumu", "Tarih", "Sağlık Kuruluşu / Notlar"]
	table = doc.add_table(rows=1 + len(MOH_VACCINE_SCHEDULE), cols=len(v_cols))
	table.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(table, color="CBD5E1", sz="4", val="single")

	# Header row
	hdr_row = table.rows[0]
	hdr_trPr = hdr_row._element.get_or_add_trPr()
	hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

	for idx, col_name in enumerate(v_cols):
		c = hdr_row.cells[idx]
		set_cell_background(c, "0F766E")
		set_cell_margins(c, top=80, bottom=80, left=100, right=100)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(col_name)
		r.font.name = "Calibri"
		r.font.size = Pt(9)
		r.bold = True
		r.font.color.rgb = RGBColor(255, 255, 255)

	for v_idx, v in enumerate(MOH_VACCINE_SCHEDULE):
		rec = vac_records.get(v["code"])
		is_done = bool(rec and rec.get("is_done"))
		dt = rec.get("administered_date") if (rec and rec.get("administered_date")) else ""
		fac = rec.get("facility") if (rec and rec.get("facility")) else ""
		notes = rec.get("notes") if (rec and rec.get("notes")) else ""

		row_idx = v_idx + 1
		row = table.rows[row_idx]
		bg_color = "F8FAFC" if (v_idx % 2 == 1) else "FFFFFF"

		if is_done:
			st_text = "✔ YAPILDI"
			st_color = RGBColor(21, 128, 61)
		elif s_age_m is not None and s_age_m >= v["period_months"]:
			st_text = "✖ EKSİK / GECİKMİŞ"
			st_color = RGBColor(185, 28, 28)
		else:
			st_text = "⏳ Zamanı Gelmedi"
			st_color = RGBColor(100, 116, 139)

		fac_notes = " · ".join(filter(None, [fac, notes])) or "—"

		vals = [
			(v["vaccine"], True, RGBColor(30, 41, 59)),
			(v["dose"], False, RGBColor(30, 41, 59)),
			(v["period_label"], False, RGBColor(30, 41, 59)),
			(st_text, True, st_color),
			(dt or "—", False, RGBColor(30, 41, 59)),
			(fac_notes, False, RGBColor(30, 41, 59)),
		]

		for c_idx, (c_val, is_bold, color) in enumerate(vals):
			c = row.cells[c_idx]
			if bg_color != "FFFFFF":
				set_cell_background(c, bg_color)
			set_cell_margins(c, top=60, bottom=60, left=80, right=80)
			cp = c.paragraphs[0]
			if c_idx in (1, 2, 3, 4):
				cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
			else:
				cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
			r = cp.add_run(c_val)
			r.font.name = "Calibri"
			r.font.size = Pt(8.5)
			r.bold = is_bold
			r.font.color.rgb = color

	# Signature section table
	doc.add_paragraph().paragraph_format.space_after = Pt(12)
	sig_table = doc.add_table(rows=1, cols=3)
	sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
	sig_data = [
		(f"{inst_nurse}\nOkul Revir Sorumlusu\n\n(İmza / Kaşe)"),
		("Aile Hekimi / Sağlık Kuruluşu\nBağışıklama Birimi\n\n(İmza / Kaşe)"),
		(f"{inst_princ}\nOkul Müdürü\n\n(Mühür / İmza)"),
	]
	for idx, text in enumerate(sig_data):
		c = sig_table.cell(0, idx)
		set_cell_margins(c, top=100, bottom=100, left=150, right=150)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(text)
		r.font.name = "Calibri"
		r.font.size = Pt(9.5)

	# Footer disclaimer
	p_foot = doc.add_paragraph()
	p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_foot.paragraph_format.space_before = Pt(10)
	r_foot = p_foot.add_run("Bu belge T.C. Sağlık Bakanlığı Çocukluk Çağı Aşı Takvimi ve okul revir kayıtları doğrultusunda düzenlenmiştir.")
	r_foot.font.name = "Calibri"
	r_foot.font.size = Pt(8)
	r_foot.font.color.rgb = RGBColor(148, 163, 184)

	buf = BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_vaccination_card_html(student_id: int) -> str:
	app_cfg = get_app_settings()
	inst_top = html.escape(app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	inst_school = html.escape(app_cfg.get("school_name") or "")
	inst_nurse = html.escape(app_cfg.get("nurse_name") or "Okul Revir Sorumlusu")
	inst_princ = html.escape(app_cfg.get("principal_name") or "Okul Müdürü / İdare")

	student = get_person_by_id(student_id)
	if not student:
		return "Öğrenci bulunamadı."

	vac_summary = get_student_vaccination_summary(student_id, student["birth_date"])
	vac_records = vac_summary["records"]

	s_age_m = age_in_months(student["birth_date"], date.today()) if student["birth_date"] else None
	s_age_str = f"{s_age_m // 12} yaş {s_age_m % 12} ay" if s_age_m is not None else "—"
	s_sex = "Kız" if student["sex"] == "K" else ("Erkek" if student["sex"] == "E" else "—")
	safe_name = html.escape(student["full_name"])
	safe_class = html.escape(student["class_name"] or "—")
	safe_sno = html.escape(str(student["student_no"] or "—"))
	safe_birth = html.escape(str(student["birth_date"] or "—"))

	rows_html = []
	for v in MOH_VACCINE_SCHEDULE:
		rec = vac_records.get(v["code"])
		is_done = bool(rec and rec.get("is_done"))
		dt = rec.get("administered_date") if (rec and rec.get("administered_date")) else ""
		fac = rec.get("facility") if (rec and rec.get("facility")) else ""
		notes = rec.get("notes") if (rec and rec.get("notes")) else ""

		if is_done:
			st_badge = "<span style='color:#15803d;font-weight:bold;'>✔ YAPILDI</span>"
		elif s_age_m is not None and s_age_m >= v["period_months"]:
			st_badge = "<span style='color:#b91c1c;font-weight:bold;'>✖ EKSİK / GECİKMİŞ</span>"
		else:
			st_badge = "<span style='color:#64748b;'>⏳ Zamanı Gelmedi</span>"

		fac_notes = " · ".join(filter(None, [fac, notes])) or "—"

		row = f"""
		<tr>
			<td><b>{html.escape(v['vaccine'])}</b></td>
			<td>{html.escape(v['dose'])}</td>
			<td>{html.escape(v['period_label'])}</td>
			<td style='text-align:center;'>{st_badge}</td>
			<td>{html.escape(dt or '—')}</td>
			<td>{html.escape(fac_notes)}</td>
		</tr>
		"""
		rows_html.append(row)

	today_str = date.today().strftime("%d.%m.%Y")

	html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{safe_name} - Öğrenci Aşı ve Bağışıklama Kartı</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; padding: 25px; line-height: 1.4; font-size: 13px; }}
.header-box {{ border-bottom: 2px solid #0f766e; padding-bottom: 12px; margin-bottom: 16px; text-align: center; }}
.header-box h1 {{ color: #0f766e; margin: 0; font-size: 20px; text-transform: uppercase; }}
.header-box h3 {{ color: #475569; margin: 3px 0 0; font-size: 13px; font-weight: normal; }}
.student-card {{ background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 6px; padding: 12px 16px; margin-bottom: 16px; font-size: 13px; }}
.summary-bar {{ display: flex; justify-content: space-between; background: #f0fdfa; border: 1px solid #ccfbf1; padding: 10px 14px; border-radius: 6px; margin-bottom: 16px; font-size: 13px; }}
table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 12px; }}
th, td {{ border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; }}
th {{ background-color: #f1f5f9; color: #0f766e; font-weight: bold; white-space: nowrap; }}
tr:nth-child(even) {{ background-color: #f8fafc; }}
.signature-section {{ display: flex; justify-content: space-between; margin-top: 35px; page-break-inside: avoid; }}
.sig-box {{ border-top: 1px dashed #64748b; width: 30%; text-align: center; padding-top: 8px; font-size: 12px; }}
@media print {{
  body {{ padding: 0; font-size: 11px; }}
  .header-box h1 {{ font-size: 16px; }}
  table {{ font-size: 10px; }}
  th, td {{ padding: 4px 6px; }}
}}
</style>
</head>
<body>
<div class='header-box'>
	<h3>T.C. SAĞLIK BAKANLIĞI & {inst_top}</h3>
	<h1>{inst_school}</h1>
	<h3>💉 ÖĞRENCİ AŞI VE BAĞIŞIKLAMA İZLEM KARTI</h3>
</div>

<div class='student-card'>
	<b>Öğrenci Adı Soyadı:</b> {safe_name} &nbsp;·&nbsp; <b>Okul No:</b> {safe_sno} &nbsp;·&nbsp; <b>Sınıf:</b> {safe_class}<br>
	<b>Doğum Tarihi:</b> {safe_birth} ({s_age_str}) &nbsp;·&nbsp; <b>Cinsiyet:</b> {s_sex} &nbsp;·&nbsp; <b>Düzenleme Tarihi:</b> {today_str}
</div>

<div class='summary-bar'>
	<div><b>Aşı Durumu:</b> %{vac_summary['percent']} Tamamlandı</div>
	<div><b>Yapılan Doz:</b> {vac_summary['applicable_done_count']} / {vac_summary['applicable_total']} Doz</div>
	<div><b>Eksik / Gecikmiş:</b> {len(vac_summary['overdue_list'])} Doz</div>
</div>

<table>
	<thead>
		<tr>
			<th>Aşı / Antijen</th>
			<th>Doz</th>
			<th>Takvim Dönemi</th>
			<th>Durum</th>
			<th>Uygulama Tarihi</th>
			<th>Uygulayan Kurum / Not</th>
		</tr>
	</thead>
	<tbody>
		{''.join(rows_html)}
	</tbody>
</table>

<div class='signature-section'>
	<div class='sig-box'>
		<b>{inst_nurse}</b><br>
		Okul Revir Sorumlusu<br><br>
		(İmza / Kaşe)
	</div>
	<div class='sig-box'>
		<b>Öğrenci Velisi</b><br>
		Bilgilendirme Onayı<br><br>
		(Tarih / İmza)
	</div>
	<div class='sig-box'>
		<b>{inst_princ}</b><br>
		Okul Müdürü<br><br>
		(Mühür / İmza)
	</div>
</div>

<p style='margin-top: 25px; font-size: 10px; color: #94a3b8; text-align: center;'>
	Bu belge T.C. Sağlık Bakanlığı Çocukluk Dönemi Aşı Takvimi esas alınarak okul revir sistemi tarafından düzenlenmiştir.
</p>
</body>
</html>"""
	return html_content


def _normalize_name_for_match(name: str) -> str:
	if not name:
		return ""
	s = str(name).strip().upper()
	trans = str.maketrans("İIĞÜŞÖÇ", "IIGUSOC")
	return re.sub(r"[^A-Z0-9]", "", s.translate(trans))


def sync_master_documents_and_files() -> str:
	"""
	Resmi PDF (Karaman Ozel Egitim Anaokulu Sinif/Yas Listesi) ve Excel listelerini
	veritabani (students, classes, health_profiles) ve kisi_verileri JSON dosyalariyla
	tamamen esitler ve senkronize eder.
	"""
	import pymupdf
	pdf_path = PERSANTIL_DATA_DIR / "ogrenci_dogum_tarihleri_ve_yaslari.pdf"
	if not pdf_path.exists():
		pdf_path = None
	xl_path = PERSANTIL_DATA_DIR / "ogrenci_adres_ve_telefon_listesi.xlsx"
	if not xl_path.exists():
		xl_path = None

	pdf_lookup = {}
	if pdf_path and pdf_path.exists():
		try:
			doc = pymupdf.open(pdf_path)
			for page in doc:
				text = page.get_text()
				m_cls = re.search(r"(Özel Eğitim Ana Sınıfı\s*/\s*[A-ZĞÜŞİÖÇ]\s*Şubesi)", text)
				cls_name = re.sub(r"\s+", " ", m_cls.group(1).replace("  ", " ")).strip() if m_cls else "Özel Eğitim Ana Sınıfı"
				lines = [l.strip() for l in text.splitlines() if l.strip()]
				for i, l in enumerate(lines):
					if re.match(r"^\d{2}/\d{2}/\d{4}$", l) and not l.startswith("19/10/2026"):
						b_date_str = l
						try:
							dt_iso = datetime.strptime(b_date_str, "%d/%m/%Y").strftime("%Y-%m-%d")
						except Exception:
							dt_iso = None
						sex_val = lines[i+2] if i+2 < len(lines) else None
						sex_code = "K" if ("Kız" in str(sex_val) or "Kz" in str(sex_val)) else "E"
						soyad = lines[i-2] if i-2 >= 0 else ""
						ad = lines[i-3] if i-3 >= 0 else ""
						stu_no = re.sub(r"[^\d]", "", lines[i-4] if i-4 >= 0 else "")
						fn = f"{ad} {soyad}".strip()
						pdf_lookup[_normalize_name_for_match(fn)] = {
							"class_name": cls_name,
							"student_no": stu_no or None,
							"full_name": fn,
							"birth_date": dt_iso,
							"sex": sex_code,
						}
		except Exception as e:
			print(f"PDF sync error: {e}")

	xl_lookup = {}
	if xl_path and xl_path.exists():
		try:
			df_xl = pd.read_excel(xl_path)
			df_xl.columns = ["full_name", "koy", "mahalle", "telefon"]
			for _, r in df_xl.iterrows():
				fn = str(r["full_name"]).strip()
				tel = str(r["telefon"]).strip().replace(".0", "").replace(" ", "").replace("-", "")
				if tel.isdigit() and len(tel) == 10 and not tel.startswith("0"):
					tel = "0" + tel
				elif not tel.isdigit() or len(tel) < 7:
					tel = None
				koy = str(r["koy"]).strip() if pd.notna(r["koy"]) and str(r["koy"]).strip().lower() != "nan" else None
				mah = str(r["mahalle"]).strip() if pd.notna(r["mahalle"]) and str(r["mahalle"]).strip().lower() != "nan" else None
				xl_lookup[_normalize_name_for_match(fn)] = {
					"full_name": fn,
					"telefon": tel,
					"koy": koy,
					"mahalle": mah,
				}
		except Exception as e:
			print(f"Excel sync error: {e}")

	with connect_db() as connection:
		for c in set(r["class_name"] for r in pdf_lookup.values()):
			connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (c,))

		class_map = {r["name"]: r["id"] for r in connection.execute("SELECT id, name FROM classes").fetchall()}
		db_students = connection.execute("SELECT * FROM students").fetchall()
		updated_cnt = 0

		for st in db_students:
			k = _normalize_name_for_match(st["full_name"])
			pdf_r = pdf_lookup.get(k)
			xl_r = xl_lookup.get(k)

			target_sno = (pdf_r["student_no"] if pdf_r and pdf_r["student_no"] else st["student_no"]) or None
			target_bdate = (pdf_r["birth_date"] if pdf_r and pdf_r["birth_date"] else st["birth_date"]) or None
			target_sex = (pdf_r["sex"] if pdf_r and pdf_r["sex"] else st["sex"]) or None

			target_cid = st["class_id"]
			if pdf_r and pdf_r["class_name"] in class_map:
				target_cid = class_map[pdf_r["class_name"]]

			target_phone = st["guardian_phone"]
			if xl_r and xl_r["telefon"]:
				target_phone = xl_r["telefon"]

			target_vil = st["village"]
			if xl_r and xl_r["koy"]:
				target_vil = xl_r["koy"].title()
			elif xl_r and xl_r["mahalle"] and (not target_vil or target_vil == "Merkez"):
				target_vil = xl_r["mahalle"].title()

			target_notes = st["notes"]
			if xl_r and xl_r["mahalle"]:
				target_notes = xl_r["mahalle"]

			connection.execute(
				"""UPDATE students SET
				   class_id=?, student_no=?, sex=?, birth_date=?, guardian_phone=?, village=?, notes=?
				   WHERE id=?""",
				(target_cid, target_sno, target_sex, target_bdate, target_phone, target_vil, target_notes, st["id"])
			)
			connection.execute(
				"""INSERT INTO health_profiles(student_id, guardian_phone, updated_at) VALUES (?, ?, ?)
				   ON CONFLICT(student_id) DO UPDATE SET guardian_phone=COALESCE(excluded.guardian_phone, health_profiles.guardian_phone)""",
				(st["id"], target_phone, datetime.now().isoformat(timespec="seconds"))
			)
			updated_cnt += 1

	sync_student_measurements()
	with connect_db() as connection:
		for st in connection.execute("SELECT id FROM students").fetchall():
			save_person_data_file(st["id"])

	return f"Senkronizasyon Başarılı: {updated_cnt} öğrenci kaydı PDF ({len(pdf_lookup)} kişi) ve Excel ({len(xl_lookup)} kişi) listeleriyle, veritabanı ve kişisel JSON dosyalarıyla eşitlendi."


def sync_student_measurements(student_id: int | None = None) -> None:
	with connect_db() as connection:
		if student_id is not None:
			rows = connection.execute(
				"""SELECT m.id, m.measured_on, m.height_cm, m.weight_kg, s.birth_date
				   FROM measurements m JOIN students s ON s.id = m.student_id
				   WHERE m.student_id = ?""",
				(student_id,),
			).fetchall()
		else:
			rows = connection.execute(
				"""SELECT m.id, m.measured_on, m.height_cm, m.weight_kg, s.birth_date
				   FROM measurements m JOIN students s ON s.id = m.student_id"""
			).fetchall()
		updates = []
		for r in rows:
			try:
				m_date = date.fromisoformat(str(r["measured_on"])[:10])
			except Exception:
				m_date = date.today()
			months = age_in_months(r["birth_date"], m_date)
			h_val = float(r["height_cm"]) if r["height_cm"] is not None else None
			w_val = float(r["weight_kg"]) if r["weight_kg"] is not None else None
			bmi_val = (w_val / ((h_val / 100.0) ** 2)) if (h_val and w_val and h_val > 0 and w_val > 0) else None
			updates.append((months, bmi_val, r["id"]))
		if updates:
			connection.executemany(
				"UPDATE measurements SET age_months = ?, bmi = ? WHERE id = ?",
				updates,
			)
	if student_id is not None:
		save_person_data_file(student_id)


def delete_student(student_id: int) -> None:
	with connect_db() as connection:
		connection.execute("DELETE FROM visits WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM measurements WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM vital_readings WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM health_profiles WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM health_forms WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM care_plans WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM lab_reports WHERE student_id = ?", (student_id,))
		connection.execute("DELETE FROM students WHERE id = ?", (student_id,))
	delete_person_data_file(student_id)


def delete_visit(visit_id: int) -> None:
	person_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT student_id FROM visits WHERE id = ?", (visit_id,)).fetchone()
		if row:
			person_id = int(row["student_id"])
		connection.execute("DELETE FROM visits WHERE id = ?", (visit_id,))
	if person_id is not None:
		save_person_data_file(person_id)


def update_visit(visit_id: int, complaint: str, symptoms: str, observation: str, action: str,
				 guardian_contacted: bool, referral: str, follow_up: str) -> None:
	person_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT student_id FROM visits WHERE id = ?", (visit_id,)).fetchone()
		if row:
			person_id = int(row["student_id"])
		connection.execute(
			"""UPDATE visits SET complaint=?, symptoms=?, observation=?, action=?,
			   guardian_contacted=?, referral=?, follow_up=? WHERE id=?""",
			(complaint, symptoms, observation, action, int(guardian_contacted), referral, follow_up, visit_id),
		)
	if person_id is not None:
		save_person_data_file(person_id)


def delete_measurement(measurement_id: int) -> None:
	person_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT student_id FROM measurements WHERE id = ?", (measurement_id,)).fetchone()
		if row:
			person_id = int(row["student_id"])
		connection.execute("DELETE FROM measurements WHERE id = ?", (measurement_id,))
	if person_id is not None:
		save_person_data_file(person_id)


def delete_vital(vital_id: int) -> None:
	person_id = None
	with connect_db() as connection:
		row = connection.execute("SELECT student_id FROM vital_readings WHERE id = ?", (vital_id,)).fetchone()
		if row:
			person_id = int(row["student_id"])
		connection.execute("DELETE FROM vital_readings WHERE id = ?", (vital_id,))
	if person_id is not None:
		save_person_data_file(person_id)


def clean_text(value: object) -> str:
	if value is None or pd.isna(value):
		return ""
	return " ".join(str(value).strip().split())


def normalize_header(value: object) -> str:
	text = clean_text(value).lower()
	table = str.maketrans({"ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c"})
	return text.translate(table).replace("_", " ")


def parse_birth_date(value: object) -> str | None:
	if value is None or pd.isna(value) or not clean_text(value):
		return None
	if isinstance(value, (datetime, date, pd.Timestamp)):
		return value.strftime("%Y-%m-%d")
	text = clean_text(value)
	for pattern in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
		try:
			return datetime.strptime(text, pattern).strftime("%Y-%m-%d")
		except ValueError:
			pass
	return None


def normalize_sex(value: object) -> str | None:
	text = normalize_header(value)
	if text in {"k", "kiz", "kadin", "female", "f"}:
		return "K"
	if text in {"e", "erkek", "male", "m"}:
		return "E"
	return None


def _infer_sex_from_turkish_name(full_name: str) -> str | None:
	tokens = [normalize_header(t) for t in full_name.split() if t.strip()]
	if not tokens:
		return None
	first_tokens = tokens[:-1] if len(tokens) > 1 else tokens
	female_names = {
		"azra", "busra", "liya", "ipek", "hatice", "kubra", "esila", "nurcan", "eslim", "songul",
		"zeynep", "ikra", "naz", "nare", "nisa", "seycan", "zubeyde", "nur", "saliha", "ecrin",
		"fadime", "rahime", "ayse", "sacideh", "fatma", "emine", "elif", "zehra", "merve", "rabia",
	}
	male_names = {
		"tugrul", "adar", "adem", "hamza", "ali", "asaf", "alparslan", "canpolat", "ahmet", "yigit",
		"ayaz", "goktug", "hasan", "huseyin", "kayra", "ekrem", "emin", "ensar", "metehan", "mirac",
		"safa", "omer", "tuna", "eymen", "efe", "yusuf", "gokbilge", "mustafa", "ramazan", "aybars",
		"yagiz", "han", "baran", "abdullah", "mehmet", "muharrem", "recep", "veysel", "yavuz", "selim",
		"alper", "zahit", "azat", "ilker", "eren", "ural", "kerim", "sami", "mikail", "emirhan",
	}
	for tok in first_tokens:
		if tok in female_names:
			return "K"
		if tok in male_names:
			return "E"
	return None


def _tr_title(text: str) -> str:
	if not text:
		return ""
	lower_map = str.maketrans({"İ": "i", "I": "ı", "Ç": "ç", "Ğ": "ğ", "Ö": "ö", "Ş": "ş", "Ü": "ü"})
	upper_map = {"i": "İ", "ı": "I", "ç": "Ç", "ğ": "Ğ", "ö": "Ö", "ş": "Ş", "ü": "Ü"}
	words = []
	for w in text.strip().split():
		lw = w.translate(lower_map).lower()
		if lw:
			first = upper_map.get(lw[0], lw[0].upper())
			words.append(first + lw[1:])
	return " ".join(words)


def _norm_student_name_key(text: str | None) -> str:
	table = str.maketrans({
		"İ": "i", "I": "i", "ı": "i",
		"Ğ": "g", "ğ": "g",
		"Ü": "u", "ü": "u",
		"Ş": "s", "ş": "s",
		"Ö": "o", "ö": "o",
		"Ç": "c", "ç": "c",
	})
	return " ".join((text or "").translate(table).lower().split())


def parse_eokul_roster_pdf(content: bytes) -> pd.DataFrame:
	text = extract_pdf_text(content)
	current_class = "Sınıf Belirtilmedi"
	rows: list[dict[str, str]] = []
	class_re = re.compile(r"([^\n\r]*?Şubesi)\s+Doğum\s+Tarihleri", re.IGNORECASE)
	row_re = re.compile(
		r"^\s*\d+\s+(\d+)\s+([A-ZÇĞİÖŞÜa-zçğıöşü\s]+?)\s+(Erkek|Kız|Kiz|E|K)\s+(\d{2}[./-]\d{2}[./-]\d{4})",
		re.IGNORECASE,
	)
	for raw_line in text.splitlines():
		line = " ".join(raw_line.strip().split())
		if not line:
			continue
		c_match = class_re.search(line)
		if c_match:
			cls_txt = c_match.group(1).strip()
			if cls_txt:
				current_class = cls_txt
			continue
		r_match = row_re.match(line)
		if r_match:
			sno, full_name, sex_txt, bdate_txt = r_match.groups()
			rows.append({
				"class": current_class,
				"number": sno.strip(),
				"name": " ".join(full_name.strip().split()),
				"sex": "K" if sex_txt.upper().startswith("K") else "E",
				"birth": bdate_txt.strip(),
			})
	if not rows:
		raise ValueError("PDF içerisinde e-Okul sınıf/öğrenci doğum tarihi listesi formatında satır bulunamadı.")
	return pd.DataFrame(rows)


def import_roster(uploaded_file, fallback_class: str = "Sınıf Belirtilmedi") -> tuple[int, int, list[str]]:
	file_name = getattr(uploaded_file, "name", str(uploaded_file)).lower()
	if file_name.endswith(".pdf"):
		content = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else Path(str(uploaded_file)).read_bytes()
		frame = parse_eokul_roster_pdf(content)
	elif file_name.endswith(".csv"):
		frame = pd.read_csv(uploaded_file, dtype=object)
	else:
		workbook = pd.ExcelFile(uploaded_file)
		frame = pd.read_excel(workbook, sheet_name=workbook.sheet_names[0], dtype=object)

	aliases = {
		"sinif": "class", "sube": "class", "class": "class", "class name": "class", "sinif sube": "class",
		"ogrenci adi": "name", "ad soyad": "name", "ogrenci": "name", "student name": "name", "name": "name", "adi soyadi": "name",
		"okul no": "number", "ogrenci no": "number", "numara": "number", "student no": "number",
		"cinsiyet": "sex", "gender": "sex", "sex": "sex",
		"dogum tarihi": "birth", "dogum tarihi (gg.aa.yyyy)": "birth", "dogum tarihi gg.aa.yyyy": "birth", "birth date": "birth",
		"veli adi": "guardian", "veli adi soyadi": "guardian", "guardian": "guardian", "veli": "guardian",
		"veli telefonu": "phone", "telefon": "phone", "guardian phone": "phone", "tel": "phone", "iletisim": "phone",
		"ogrenci koy adi": "village", "koy adi": "village", "koy": "village", "yerlesim": "village",
		"ogrenci mahalle adi": "neighborhood", "mahalle adi": "neighborhood", "mahalle": "neighborhood", "adres": "neighborhood", "acik adres": "neighborhood",
		"il": "province", "sehir": "province",
		"ilce": "district",
		"ulke": "country",
	}
	mapped = {column: aliases.get(normalize_header(column)) for column in frame.columns}
	frame = frame.rename(columns={column: field for column, field in mapped.items() if field})
	if "name" not in frame.columns:
		raise ValueError("Yüklenen dosyada en az 'Öğrenci Adı' (veya 'Ad Soyad') sütunu bulunmalıdır.")

	app_cfg = get_app_settings()
	def_country = app_cfg.get("default_country") or "Türkiye"
	def_province = app_cfg.get("default_province") or "Karaman"
	def_district = app_cfg.get("default_district") or "Merkez"

	imported = 0
	skipped = 0
	issues: list[str] = []
	now_iso = datetime.now().isoformat(timespec="minutes")

	with connect_db() as connection:
		for row_number, row in frame.iterrows():
			name = clean_text(row.get("name"))
			if not name:
				skipped += 1
				continue

			raw_class = clean_text(row.get("class")) if "class" in frame.columns else ""
			raw_village = clean_text(row.get("village")) if "village" in frame.columns else ""
			raw_neighborhood = clean_text(row.get("neighborhood")) if "neighborhood" in frame.columns else ""
			raw_province = clean_text(row.get("province")) if "province" in frame.columns else ""
			raw_district = clean_text(row.get("district")) if "district" in frame.columns else ""
			raw_country = clean_text(row.get("country")) if "country" in frame.columns else ""

			# Find any existing student(s) matching by normalized Turkish name
			norm_target = _norm_student_name_key(name)
			all_db_students = connection.execute(
				"SELECT s.*, c.name AS class_name FROM students s JOIN classes c ON c.id=s.class_id WHERE c.name != 'Sivil / Diğer' ORDER BY s.id"
			).fetchall()
			matching_students = [s for s in all_db_students if _norm_student_name_key(s["full_name"]) == norm_target]
			existing_stu = matching_students[0] if matching_students else None

			# If duplicate rows exist for the same student, merge them into the primary record
			if len(matching_students) > 1:
				primary_id = existing_stu["id"]
				for dup in matching_students[1:]:
					dup_id = dup["id"]
					connection.execute("UPDATE measurements SET student_id=? WHERE student_id=?", (primary_id, dup_id))
					connection.execute("UPDATE visits SET student_id=? WHERE student_id=?", (primary_id, dup_id))
					connection.execute(
						"""UPDATE students SET
						   student_no=COALESCE(student_no, ?),
						   sex=COALESCE(sex, ?),
						   birth_date=COALESCE(birth_date, ?),
						   guardian_name=COALESCE(guardian_name, ?),
						   guardian_phone=COALESCE(guardian_phone, ?),
						   notes=COALESCE(notes, ?),
						   country=COALESCE(country, ?),
						   province=COALESCE(province, ?),
						   district=COALESCE(district, ?),
						   village=COALESCE(village, ?)
						   WHERE id=?""",
						(dup["student_no"], dup["sex"], dup["birth_date"], dup["guardian_name"],
						 dup["guardian_phone"], dup["notes"], dup["country"], dup["province"],
						 dup["district"], dup["village"], primary_id),
					)
					connection.execute("DELETE FROM health_profiles WHERE student_id=?", (dup_id,))
					connection.execute("DELETE FROM students WHERE id=?", (dup_id,))
				existing_stu = connection.execute(
					"SELECT s.*, c.name AS class_name FROM students s JOIN classes c ON c.id=s.class_id WHERE s.id=?",
					(primary_id,),
				).fetchone()

			if raw_class:
				class_name = raw_class
			elif existing_stu and existing_stu["class_name"]:
				class_name = existing_stu["class_name"]
			else:
				class_name = fallback_class

			connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (class_name,))
			class_id = connection.execute("SELECT id FROM classes WHERE name = ?", (class_name,)).fetchone()["id"]

			birth_date = parse_birth_date(row.get("birth"))
			if clean_text(row.get("birth")) and not birth_date:
				issues.append(f"Satır {row_number + 2}: doğum tarihi okunamadı.")

			sex_val = normalize_sex(row.get("sex")) or (existing_stu["sex"] if existing_stu else None) or _infer_sex_from_turkish_name(name)

			raw_phone = clean_text(row.get("phone")) if "phone" in frame.columns else ""
			if raw_phone in {"…", "...", "-", "—", "Yok", "YOK"}:
				phone_val = None
			elif raw_phone:
				digits = re.sub(r"\D", "", raw_phone)
				if len(digits) == 10 and digits.startswith("5"):
					phone_val = "0" + digits
				else:
					phone_val = raw_phone
			else:
				phone_val = None

			guardian_val = clean_text(row.get("guardian")) if "guardian" in frame.columns else None
			guardian_val = guardian_val or None
			country_val = raw_country or (existing_stu["country"] if existing_stu else None) or def_country
			province_val = raw_province or (existing_stu["province"] if existing_stu else None) or def_province
			district_val = raw_district or ("Merkez" if raw_village.upper() == "MERKEZ" else ((existing_stu["district"] if existing_stu else None) or def_district))

			# Determine village/neighborhood display and address note (never used as class!)
			if raw_village and raw_village.upper() not in {"MERKEZ", "İLÇE MERKEZİ"}:
				village_val = f"{_tr_title(raw_village)} Köyü" if "köy" not in raw_village.lower() else _tr_title(raw_village)
			elif raw_neighborhood:
				m_match = re.match(r"^([A-ZÇĞİÖŞÜa-zçğıöşü\s]+?(?:MAH\.?|MAHALLESİ))\b", raw_neighborhood, flags=re.IGNORECASE)
				village_val = _tr_title(m_match.group(1).strip()) if m_match else "Merkez"
			else:
				village_val = "Merkez" if raw_village.upper() == "MERKEZ" else (_tr_title(raw_village) if raw_village else None)

			notes_val = raw_neighborhood if raw_neighborhood else None
			num_val = clean_text(row.get("number")) if "number" in frame.columns else ""
			num_val = num_val or None

			if existing_stu:
				stu_id = existing_stu["id"]
				connection.execute(
					"""UPDATE students SET
					   class_id = ?,
					   person_type = 'Öğrenci',
					   full_name = ?,
					   student_no = COALESCE(?, student_no),
					   sex = COALESCE(?, sex),
					   birth_date = COALESCE(?, birth_date),
					   guardian_name = COALESCE(?, guardian_name),
					   guardian_phone = COALESCE(?, guardian_phone),
					   notes = COALESCE(?, notes),
					   country = COALESCE(?, country),
					   province = COALESCE(?, province),
					   district = COALESCE(?, district),
					   village = COALESCE(?, village)
					   WHERE id = ?""",
					(class_id, name, num_val, sex_val, birth_date, guardian_val, phone_val,
					 notes_val, country_val, province_val, district_val, village_val, stu_id),
				)
			else:
				cur = connection.execute(
					"""INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date,
					   guardian_name, guardian_phone, notes, country, province, district, village)
					   VALUES (?, 'Öğrenci', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(class_id, num_val, name, sex_val, birth_date, guardian_val, phone_val, notes_val,
					 country_val, province_val, district_val, village_val),
				)
				stu_id = cur.lastrowid

			if stu_id and (guardian_val or phone_val):
				connection.execute(
					"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, updated_at)
					   VALUES (?, ?, ?, ?)
					   ON CONFLICT(student_id) DO UPDATE SET
						 guardian_name=COALESCE(excluded.guardian_name, health_profiles.guardian_name),
						 guardian_phone=COALESCE(excluded.guardian_phone, health_profiles.guardian_phone),
						 updated_at=excluded.updated_at""",
					(stu_id, guardian_val, phone_val, now_iso),
				)
			imported += 1
	sync_all_person_data_files()
	return imported, skipped, issues


BUILTIN_CLASS_TEMPLATES: dict[str, list[str]] = {
	"Anaokulu (3-5 Yaş & Anasınıfları)": [
		"3 Yaş Grubu",
		"4 Yaş / Papatyalar",
		"5 Yaş / Yıldızlar",
		"Anasınıfı A",
		"Anasınıfı B",
		"Anasınıfı C",
	],
	"İlkokul (1-4. Sınıflar - 3 Şube)": [
		"1-A", "1-B", "1-C",
		"2-A", "2-B", "2-C",
		"3-A", "3-B", "3-C",
		"4-A", "4-B", "4-C",
	],
	"Ortaokul (5-8. Sınıflar - 3 Şube)": [
		"5-A", "5-B", "5-C",
		"6-A", "6-B", "6-C",
		"7-A", "7-B", "7-C",
		"8-A", "8-B", "8-C",
	],
	"Lise (9-12. Sınıflar - 4 Şube)": [
		"9-A", "9-B", "9-C", "9-D",
		"10-A", "10-B", "10-C", "10-D",
		"11-A", "11-B", "11-C", "11-D",
		"12-A", "12-B", "12-C", "12-D",
	],
	"Özel Eğitim (1, 2 ve 3. Kademe)": [
		"1. Kademe A", "1. Kademe B",
		"2. Kademe A", "2. Kademe B",
		"3. Kademe A", "3. Kademe B",
	],
}


def get_custom_class_templates() -> dict[str, list[str]]:
	settings = get_app_settings()
	raw = settings.get("custom_class_templates", "")
	if raw and raw.strip():
		try:
			parsed = json.loads(raw)
			if isinstance(parsed, dict):
				return parsed
		except Exception:
			pass
	return {}


def save_custom_class_template(name: str, classes: list[str]) -> bool:
	clean_name = name.strip()
	if not clean_name:
		return False
	current = get_custom_class_templates()
	seen = set()
	cleaned = []
	for c in classes:
		c_str = str(c).strip()
		if c_str and c_str not in seen and c_str not in {"Sivil / Diğer", "Sınıf Belirtilmedi"}:
			seen.add(c_str)
			cleaned.append(c_str)
	if not cleaned:
		return False
	current[clean_name] = cleaned
	save_app_settings({"custom_class_templates": json.dumps(current, ensure_ascii=False, indent=2)})
	return True


def delete_custom_class_template(name: str) -> bool:
	current = get_custom_class_templates()
	if name in current:
		del current[name]
		save_app_settings({"custom_class_templates": json.dumps(current, ensure_ascii=False, indent=2)})
		return True
	return False


def get_all_class_templates() -> dict[str, list[str]]:
	templates = dict(BUILTIN_CLASS_TEMPLATES)
	custom = get_custom_class_templates()
	for k, v in custom.items():
		templates[f"⭐ {k} (Özel Şablon)"] = v
	return templates


def apply_class_template(classes_to_add: list[str], clear_empty: bool = False) -> tuple[int, int]:
	added = 0
	removed = 0
	with connect_db() as connection:
		if clear_empty:
			empty_rows = connection.execute(
				"""SELECT c.id, c.name FROM classes c
				   LEFT JOIN students s ON s.class_id = c.id
				   WHERE c.name NOT IN ('Sivil / Diğer', 'Sınıf Belirtilmedi')
				   GROUP BY c.id HAVING COUNT(s.id) = 0"""
			).fetchall()
			for r in empty_rows:
				connection.execute("DELETE FROM classes WHERE id=?", (r["id"],))
				removed += 1
		for c_name in classes_to_add:
			c_clean = str(c_name).strip()
			if c_clean:
				cursor = connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (c_clean,))
				if cursor.rowcount > 0:
					added += 1
	return added, removed


def get_classes() -> list[str]:
	with connect_db() as connection:
		return [row["name"] for row in connection.execute("SELECT name FROM classes ORDER BY name")]


def get_students(class_name: str | None = None) -> pd.DataFrame:
	query = """SELECT s.id, s.person_type AS kayit_turu, c.name AS sinif, s.student_no AS okul_no, s.full_name AS ad_soyad,
					  s.sex AS cinsiyet, s.birth_date AS dogum_tarihi,
					  s.guardian_name AS veli, s.guardian_phone AS veli_telefon, s.notes AS notlar
			   FROM students s JOIN classes c ON c.id=s.class_id"""
	params: tuple = ()
	if class_name and class_name != "Tüm sınıflar":
		query += " WHERE c.name = ?"
		params = (class_name,)
	query += " ORDER BY c.name, s.full_name"
	with connect_db() as connection:
		return pd.read_sql_query(query, connection, params=params)


def age_in_months(birth_date: str | None, on_date: date) -> int | None:
	if not birth_date or not isinstance(birth_date, str) or pd.isna(birth_date):
		return None
	try:
		born = date.fromisoformat(birth_date.strip()[:10])
	except Exception:
		return None
	months = (on_date.year - born.year) * 12 + on_date.month - born.month
	if on_date.day < born.day:
		months -= 1
	return max(0, months)


def age_in_days(birth_date: str | None, on_date: date) -> int | None:
	if not birth_date or not isinstance(birth_date, str) or pd.isna(birth_date):
		return None
	try:
		return max(0, (on_date - date.fromisoformat(birth_date.strip()[:10])).days)
	except Exception:
		return None


def z_score(value: float, l_value: float, m_value: float, s_value: float) -> float:
	if value <= 0 or m_value <= 0 or s_value <= 0:
		raise ValueError("Ölçüm ve LMS değerleri sıfırdan büyük olmalı.")
	if abs(l_value) < 1e-9:
		return math.log(value / m_value) / s_value
	return ((value / m_value) ** l_value - 1) / (l_value * s_value)


def percentile(z_value: float) -> float:
	return 50 * (1 + math.erf(z_value / math.sqrt(2)))


def numeric_value(value: object) -> float | None:
	if value is None or pd.isna(value) or not clean_text(value):
		return None
	if isinstance(value, str):
		value = value.strip().replace(",", ".")
	try:
		return float(value)
	except (TypeError, ValueError):
		return None


def assess_range(value: float, low: float | None, high: float | None) -> str:
	if low is None or high is None or high <= low:
		return "Referans aralığı yok"
	if value < low:
		return "Referansın altında"
	if value > high:
		return "Referansın üstünde"
	margin = (high - low) * 0.1
	if value <= low + margin:
		return "Alt sınıra yakın"
	if value >= high - margin:
		return "Üst sınıra yakın"
	return "Rapor aralığında"


def assess_lab_history(frame: pd.DataFrame) -> pd.DataFrame:
	result = frame.copy()
	if result.empty:
		result["Değerlendirme"] = pd.Series(dtype="object")
		return result
	statuses = []
	for _, row in result.iterrows():
		status = assess_range(float(row["Sonuç"]), numeric_value(row.get("Referans_Alt")),
						  numeric_value(row.get("Referans_Üst")))
		flag = clean_text(row.get("Rapor_İşareti"))
		if flag:
			status = f"{status}; rapor işareti: {flag}"
		statuses.append(status)
	result["Değerlendirme"] = statuses
	return result


def lab_result_columns() -> list[str]:
	return ["Test", "Sonuç", "Birim", "Referans alt", "Referans üst", "Laboratuvar işareti"]


def parse_reference_range(value: object) -> tuple[float | None, float | None]:
	text = clean_text(value).replace(",", ".")
	match = re.search(r"(-?\d+(?:\.\d+)?)\s*[-–]\s*(-?\d+(?:\.\d+)?)", text)
	if not match:
		return None, None
	return float(match.group(1)), float(match.group(2))


@st.cache_resource
def get_ocr_reader():
	import truststore
	truststore.inject_into_ssl()
	import easyocr
	return easyocr.Reader(["tr", "en"], gpu=False, verbose=False)


def extract_pdf_text(content: bytes) -> str:
	import pymupdf
	import numpy as np
	from PIL import Image

	with pymupdf.open(stream=content, filetype="pdf") as document:
		pages = []
		for page in document:
			page_text = page.get_text("text")
			if len(re.sub(r"\s", "", page_text)) < 30:
				pixmap = page.get_pixmap(dpi=240, alpha=False)
				image = np.asarray(Image.open(BytesIO(pixmap.tobytes("png"))).convert("RGB"))
				ocr_lines = get_ocr_reader().readtext(image, detail=0, paragraph=False)
				page_text = "\n".join(ocr_lines)
			pages.append(page_text)
			widgets = page.widgets()
			if widgets:
				pages.extend(f"{widget.field_name}: {widget.field_value}" for widget in widgets if widget.field_value)
	return "\n".join(pages)


def extract_health_form_fields(text: str) -> dict[str, str]:
	fields = {}
	for line in text.splitlines():
		parts = re.split(r"\s*[:：]\s*", line.strip(), maxsplit=1)
		if len(parts) < 2:
			parts = re.split(r"\s+[–-]\s+", line.strip(), maxsplit=1)
		if len(parts) < 2:
			continue
		label = normalize_header(parts[0]).replace("\ufffd", "i").replace("\u00b7", "i")
		value = " ".join(parts[1].strip().split())
		if not value:
			continue
		if label.startswith(("veli telefon", "ebeveyn telefon", "acil iletisim telefon", "telefon", "gsm")):
			field = "guardian_phone"
		elif label.startswith(("veli adi", "veli ismi", "ebeveyn adi", "ebeveyn ismi", "yasal temsilci", "acil durumda aranacak kisi")):
			field = "guardian_name"
		elif "alerji" in label:
			field = "allergies"
		elif any(term in label for term in ("kronik hastalik", "kronik rahatsizlik", "kronik durum")):
			field = "chronic_conditions"
		elif any(term in label for term in ("saglik gecmisi", "saglik oykusu", "gecirilen hastalik", "gecirilen ameliyat", "bilinen hastalik", "ozgecmis", "saglik sorunu", "tani", "yasadigi durum")):
			field = "medical_history"
		elif any(term in label for term in ("duzenli ilac", "kullandigi ilac", "ilaclar", "ilac kullanimi")):
			field = "medications"
		elif label.startswith(("kan grubu", "blood type")):
			field = "blood_group"
		elif any(term in label for term in ("acil durum notu", "ozel durum", "acil mudahale", "saglik notu")):
			field = "emergency_notes"
		else:
			continue
		fields.setdefault(field, value)
	return fields


def parse_lab_upload(uploaded_file) -> tuple[pd.DataFrame, str]:
	content = uploaded_file.getvalue()
	suffix = Path(uploaded_file.name).suffix.lower()
	columns = lab_result_columns()
	if suffix == ".pdf":
		text = extract_pdf_text(content)
		pattern = re.compile(
			r"^(?P<test>[\wİıĞğÜüŞşÖöÇç ()/.,+-]{2,60}?)\s+"
			r"(?P<value>-?\d+(?:[.,]\d+)?)\s*"
			r"(?P<unit>[^\s\d]+(?:/[^\s\d]+)?)?\s*"
			r"(?:(?P<low>-?\d+(?:[.,]\d+)?)\s*[-–]\s*(?P<high>-?\d+(?:[.,]\d+)?))?\s*"
			r"(?P<flag>HIGH|LOW|H|L|YÜKSEK|DÜŞÜK|\*)?\s*$", re.IGNORECASE
		)
		rows = []
		for line in text.splitlines():
			match = pattern.match(line.strip())
			if match:
				rows.append({"Test": match["test"].strip(), "Sonuç": numeric_value(match["value"]),
							 "Birim": match["unit"] or "", "Referans alt": numeric_value(match["low"]),
							 "Referans üst": numeric_value(match["high"]), "Laboratuvar işareti": match["flag"] or ""})
		return pd.DataFrame(rows, columns=columns), text
	if suffix == ".csv":
		frame = pd.read_csv(BytesIO(content))
	elif suffix in {".xlsx", ".xlsm"}:
		frame = pd.read_excel(BytesIO(content))
	else:
		raise ValueError("Yalnızca PDF, Excel veya CSV laboratuvar raporu yüklenebilir.")
	aliases = {
		"test": "Test", "test name": "Test", "test adi": "Test", "analit": "Test", "tahlil": "Test", "parametre": "Test",
		"result": "Sonuç", "sonuc": "Sonuç", "value": "Sonuç", "deger": "Sonuç",
		"unit": "Birim", "birim": "Birim",
		"reference low": "Referans alt", "ref low": "Referans alt", "referans alt": "Referans alt", "alt sinir": "Referans alt",
		"reference high": "Referans üst", "ref high": "Referans üst", "referans ust": "Referans üst", "ust sinir": "Referans üst",
		"reference range": "Referans", "referans araligi": "Referans",
		"flag": "Laboratuvar işareti", "durum": "Laboratuvar işareti", "isaret": "Laboratuvar işareti",
	}
	frame = frame.rename(columns={column: aliases.get(normalize_header(column)) for column in frame.columns})
	if "Test" not in frame.columns or "Sonuç" not in frame.columns:
		raise ValueError("Dosyada test adı ve sonuç sütunu bulunamadı. PDF/Excel/CSV dosyasını yükleyip satırları elle girebilirsiniz.")
	rows = []
	for _, row in frame.iterrows():
		low = numeric_value(row.get("Referans alt"))
		high = numeric_value(row.get("Referans üst"))
		if (low is None or high is None) and clean_text(row.get("Referans")):
			low, high = parse_reference_range(row.get("Referans"))
		rows.append({"Test": clean_text(row.get("Test")), "Sonuç": numeric_value(row.get("Sonuç")),
					 "Birim": clean_text(row.get("Birim")), "Referans alt": low, "Referans üst": high,
					 "Laboratuvar işareti": clean_text(row.get("Laboratuvar işareti"))})
	return pd.DataFrame(rows, columns=columns), ""


def parse_vital_reference_pdf(uploaded_file) -> pd.DataFrame:
	text = extract_pdf_text(uploaded_file.getvalue())
	rows = []
	age_pattern = re.compile(r"(?P<age_min>\d+)\s*[-–]\s*(?P<age_max>\d+)\s*(?P<age_unit>months?|mos?|mo|ay|years?|yrs?|yr|yil|yıl|y)\b", re.IGNORECASE)
	number_range = re.compile(r"(-?\d+(?:[.,]\d+)?)\s*[-–]\s*(-?\d+(?:[.,]\d+)?)")
	unit_pattern = re.compile(r"mm\s?hg|beats\s?/\s?min|bpm|atım\s?/\s?dk|mg\s?/\s?dl|mmol\s?/\s?l|°\s?[cf]|celsius|fahrenheit|%|yüzde", re.IGNORECASE)
	for line in text.splitlines():
		normalized = normalize_header(line)
		if "systolic" in normalized or "buyuk tansiyon" in normalized:
			indicator = "systolic_bp"
		elif "diastolic" in normalized or "kucuk tansiyon" in normalized:
			indicator = "diastolic_bp"
		elif "heart rate" in normalized or "pulse" in normalized or "nabiz" in normalized:
			indicator = "pulse"
		elif "temperature" in normalized or "body temperature" in normalized or "ates" in normalized:
			indicator = "temperature"
		elif "spo2" in normalized or "saturasyon" in normalized or "oxygen saturation" in normalized or "oksijen" in normalized:
			indicator = "spo2"
		elif "blood glucose" in normalized or "glucose" in normalized or "kan sekeri" in normalized:
			indicator = "glucose"
		else:
			continue
		age_match = age_pattern.search(line)
		if not age_match:
			continue
		ranges = list(number_range.finditer(line))
		value_range = next((match for match in ranges if match.start() >= age_match.end()), None)
		if value_range is None:
			value_range = next((match for match in ranges if match.end() <= age_match.start()), None)
		unit_match = unit_pattern.search(line)
		if value_range is None or unit_match is None:
			continue
		age_min, age_max = int(age_match["age_min"]), int(age_match["age_max"])
		age_unit = normalize_header(age_match["age_unit"])
		if age_unit in {"year", "years", "yr", "yrs", "yil", "y"}:
			age_min, age_max = age_min * 12, age_max * 12
		sex = "*"
		if re.search(r"\b(female|girl|kiz)\b", normalized):
			sex = "K"
		elif re.search(r"\b(male|boy|erkek)\b", normalized):
			sex = "E"
		rows.append({
			"Gösterge": indicator, "Cinsiyet": sex, "Yaş alt ay": age_min, "Yaş üst ay": age_max,
			"Referans alt": numeric_value(value_range.group(1)), "Referans üst": numeric_value(value_range.group(2)),
			"Birim": unit_match.group(0), "Kaynak": Path(uploaded_file.name).name,
		})
	if not rows:
		raise ValueError("PDF'den yaş, gösterge, birim ve aralık içeren okunabilir satır çıkarılamadı. PDF'yi kontrol edin veya şablona elle aktarın.")
	return pd.DataFrame(rows)


def parse_vital_reference_upload(uploaded_file) -> tuple[int, list[str]]:
	suffix = Path(uploaded_file.name).suffix.lower()
	if suffix == ".pdf":
		frame = parse_vital_reference_pdf(uploaded_file)
	elif suffix == ".csv":
		frame = pd.read_csv(uploaded_file)
	else:
		frame = pd.read_excel(uploaded_file)
	aliases = {
		"gosterge": "indicator", "indicator": "indicator", "olcum": "indicator",
		"cinsiyet": "sex", "sex": "sex", "gender": "sex",
		"yas alt ay": "age_min", "yas min ay": "age_min", "age min months": "age_min", "age_min_months": "age_min",
		"yas ust ay": "age_max", "yas max ay": "age_max", "age max months": "age_max", "age_max_months": "age_max",
		"referans alt": "low", "ref low": "low", "low": "low",
		"referans ust": "high", "ref high": "high", "high": "high",
		"kritik alt": "critical_low", "critical low": "critical_low",
		"kritik ust": "critical_high", "critical high": "critical_high",
		"birim": "unit", "unit": "unit", "kaynak": "source", "source": "source",
	}
	frame = frame.rename(columns={column: aliases.get(normalize_header(column)) for column in frame.columns})
	required = {"indicator", "age_min", "age_max", "low", "high", "unit", "source"}
	if not required.issubset(frame.columns):
		raise ValueError("Gerekli sütunlar: Gösterge, Yaş alt ay, Yaş üst ay, Referans alt, Referans üst, Birim, Kaynak.")
	indicator_map = {"systolic bp": "systolic_bp", "systolic blood pressure": "systolic_bp", "buyuk tansiyon": "systolic_bp",
					 "diastolic bp": "diastolic_bp", "diastolic blood pressure": "diastolic_bp", "kucuk tansiyon": "diastolic_bp",
					 "pulse": "pulse", "heart rate": "pulse", "nabiz": "pulse",
					 "temperature": "temperature", "body temperature": "temperature", "ates": "temperature",
					 "spo2": "spo2", "saturasyon": "spo2", "oksijen saturasyonu": "spo2", "oxygen saturation": "spo2",
					 "glucose": "glucose", "blood glucose": "glucose", "kan sekeri": "glucose"}
	issues = []
	count = 0
	with connect_db() as connection:
		for row_number, row in frame.iterrows():
			indicator = indicator_map.get(normalize_header(row["indicator"]))
			sex = normalize_sex(row.get("sex")) or "*"
			try:
				age_min, age_max = int(row["age_min"]), int(row["age_max"])
				low, high = float(row["low"]), float(row["high"])
				unit, source = clean_text(row["unit"]), clean_text(row["source"])
				critical_low = numeric_value(row.get("critical_low"))
				critical_high = numeric_value(row.get("critical_high"))
				if not indicator or age_min < 0 or age_max < age_min or high <= low or not unit or not source:
					raise ValueError
				normalized_unit = normalize_header(unit).replace(" ", "").replace("°", "")
				if indicator in {"systolic_bp", "diastolic_bp"}:
					if normalized_unit != "mmhg":
						raise ValueError
					unit = "mmHg"
				elif indicator == "pulse":
					if normalized_unit not in {"bpm", "beats/min", "atim/dk"}:
						raise ValueError
					unit = "atım/dk"
				elif indicator == "temperature":
					if normalized_unit in {"f", "fahrenheit", "derecef"}:
						low, high = (low - 32) * 5 / 9, (high - 32) * 5 / 9
						critical_low = (critical_low - 32) * 5 / 9 if critical_low is not None else None
						critical_high = (critical_high - 32) * 5 / 9 if critical_high is not None else None
					elif normalized_unit not in {"c", "celsius", "derecec"}:
						raise ValueError
					unit = "°C"
				elif indicator == "spo2":
					if normalized_unit not in {"%", "yuzde", "pct", "spo2"}:
						raise ValueError
					unit = "%"
				elif indicator == "glucose":
					if normalized_unit in {"mmol/l", "mmolperliter"}:
						low, high = low * 18.0182, high * 18.0182
						critical_low = critical_low * 18.0182 if critical_low is not None else None
						critical_high = critical_high * 18.0182 if critical_high is not None else None
					elif normalized_unit != "mg/dl":
						raise ValueError
					unit = "mg/dL"
			except (TypeError, ValueError, OverflowError):
				issues.append(f"Satır {row_number + 2}: gösterge, yaş, referans, birim veya kaynak geçersiz.")
				continue
			connection.execute(
				"""INSERT INTO vital_reference(indicator, sex, age_min_months, age_max_months, ref_low, ref_high,
				   critical_low, critical_high, unit, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				   ON CONFLICT(indicator, sex, age_min_months, age_max_months, source) DO UPDATE SET
				   ref_low=excluded.ref_low, ref_high=excluded.ref_high, critical_low=excluded.critical_low,
				   critical_high=excluded.critical_high, unit=excluded.unit""",
				(indicator, sex, age_min, age_max, low, high, critical_low, critical_high, unit, source),
			)
			count += 1
	return count, issues


def vital_assessment(indicator: str, value: float, student: sqlite3.Row, on_date: date) -> dict | None:
	months = age_in_months(student["birth_date"], on_date)
	if months is not None:
		sex = student["sex"] or "*"
		with connect_db() as connection:
			reference = connection.execute(
				"""SELECT * FROM vital_reference WHERE indicator=? AND age_min_months<=? AND age_max_months>=?
				   AND sex IN (?, '*') ORDER BY (sex=?) DESC, (age_max_months-age_min_months), id DESC LIMIT 1""",
				(indicator, months, months, sex, sex),
			).fetchone()
		if reference is not None:
			status = assess_range(value, reference["ref_low"], reference["ref_high"])
			if reference["critical_low"] is not None and value < reference["critical_low"]:
				status = "Kaynakta kritik düşük"
			elif reference["critical_high"] is not None and value > reference["critical_high"]:
				status = "Kaynakta kritik yüksek"
			return {"status": status, "unit": reference["unit"], "source": reference["source"]}
	if indicator == "spo2":
		if value < 90:
			return {"status": "Kaynakta kritik düşük (<%90 Hipoksi)", "unit": "%", "source": "Standart Klinik Tarama (95-100%)"}
		elif value < 95:
			return {"status": "Referansın altında (<%95)", "unit": "%", "source": "Standart Klinik Tarama (95-100%)"}
		return {"status": "Rapor aralığında (Normal %95-100)", "unit": "%", "source": "Standart Klinik Tarama (95-100%)"}
	return None


def make_vital_reference_template() -> bytes:
	workbook = Workbook()
	sheet = workbook.active
	sheet.title = "Vital Referansları"
	sheet.append(["Gösterge", "Cinsiyet", "Yaş alt ay", "Yaş üst ay", "Referans alt", "Referans üst", "Kritik alt", "Kritik üst", "Birim", "Kaynak"])
	info = workbook.create_sheet("Bilgi")
	info.append(["Alan", "Açıklama"])
	info.append(["Gösterge", "systolic_bp, diastolic_bp, pulse, temperature, spo2 veya glucose"])
	info.append(["Cinsiyet", "K, E veya * (kaynak aralığı cinsiyete göre ayırmıyorsa)"])
	info.append(["Yaş", "Ay cinsinden alt ve üst sınır; sınıfa uygun resmi referans kullanın."])
	info.append(["Kaynak", "Kurumsal veya klinik kaynağın adı ve sürümü. Bu şablonda klinik eşik örneği yoktur."])
	for cell in sheet[1]:
		cell.font = Font(bold=True, color="FFFFFF")
		cell.fill = PatternFill("solid", fgColor="176B68")
	for tab in workbook.worksheets:
		for cell in tab[1]:
			cell.font = Font(bold=True, color="FFFFFF")
			cell.fill = PatternFill("solid", fgColor="176B68")
		for column_cells in tab.columns:
			tab.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max(len(str(column_cells[0].value or "")) + 3, 16), 55)
	output = BytesIO()
	workbook.save(output)
	return output.getvalue()


# WHO Child Growth Standards Head Circumference-for-Age (0-60 months) LMS reference
_WHO_HCFA_KEY_LMS = {
	"E": [
		(0, 34.4618, 0.03686), (1, 37.2759, 0.03133), (2, 39.1285, 0.02997), (3, 40.5135, 0.02918),
		(4, 41.6316, 0.02868), (5, 42.5576, 0.02832), (6, 43.3306, 0.02805), (9, 45.0169, 0.02756),
		(12, 46.0661, 0.02736), (15, 46.7982, 0.02729), (18, 47.3627, 0.02730), (24, 48.2189, 0.02741),
		(30, 48.8704, 0.02759), (36, 49.3894, 0.02779), (42, 49.8118, 0.02800), (48, 50.1636, 0.02820),
		(54, 50.4646, 0.02840), (60, 50.7289, 0.02858),
	],
	"K": [
		(0, 33.8787, 0.03496), (1, 36.5463, 0.03116), (2, 38.2521, 0.03032), (3, 39.5328, 0.02985),
		(4, 40.5817, 0.02956), (5, 41.4590, 0.02935), (6, 42.1995, 0.02918), (9, 43.8318, 0.02883),
		(12, 44.8956, 0.02864), (15, 45.6683, 0.02854), (18, 46.2772, 0.02849), (24, 47.2116, 0.02844),
		(30, 47.9261, 0.02841), (36, 48.4946, 0.02838), (42, 48.9587, 0.02835), (48, 49.3489, 0.02831),
		(54, 49.6869, 0.02828), (60, 49.9868, 0.02824),
	],
}


def get_builtin_hcfa_lms_rows() -> list[tuple[str, str, int, float, float, float]]:
	rows = []
	for sex, pts in _WHO_HCFA_KEY_LMS.items():
		for idx in range(len(pts) - 1):
			m0, med0, s0 = pts[idx]
			m1, med1, s1 = pts[idx + 1]
			for mo in range(m0, m1 + 1):
				t = (mo - m0) / (m1 - m0) if m1 > m0 else 0.0
				med = med0 + t * (med1 - med0)
				s = s0 + t * (s1 - s0)
				rows.append(("head_circ", sex, mo, 1.0, round(med, 4), round(s, 5)))
	return rows


_LMS_CACHE: dict[tuple[str, str, int], dict] = {}


def _ensure_lms_cache() -> None:
	if _LMS_CACHE:
		return
	with connect_db() as connection:
		rows = connection.execute("SELECT indicator, sex, age_months, l, m, s FROM lms_reference").fetchall()
		if len(rows) < 100:
			try:
				import_builtin_who_references()
				rows = connection.execute("SELECT indicator, sex, age_months, l, m, s FROM lms_reference").fetchall()
			except Exception:
				pass
	for r in rows:
		_LMS_CACHE[(r["indicator"], r["sex"], int(r["age_months"]))] = {
			"l": float(r["l"]),
			"m": float(r["m"]),
			"s": float(r["s"]),
			"age_months": int(r["age_months"]),
		}
	for ind, sex, mo, l_v, m_v, s_v in get_builtin_hcfa_lms_rows():
		_LMS_CACHE.setdefault((ind, sex, mo), {"l": l_v, "m": m_v, "s": s_v, "age_months": mo})


def get_lms_row(indicator: str, sex: str, months: int) -> dict | None:
	if not sex or months is None:
		return None
	_ensure_lms_cache()
	m_int = int(months)
	if indicator == "head_circ" and m_int > 60:
		m_int = 60
	for candidate in (m_int, m_int - 1, m_int + 1):
		hit = _LMS_CACHE.get((indicator, sex, int(candidate)))
		if hit is not None:
			return hit
	return None


def evaluate_measurement(value: float | None, indicator: str, sex: str | None, months: int | None) -> dict | None:
	if value is None or pd.isna(value) or value <= 0 or not sex or months is None or pd.isna(months):
		return None
	reference = get_lms_row(indicator, sex, int(months))
	if reference is None:
		return None
	z_value = z_score(float(value), reference["l"], reference["m"], reference["s"])
	pct_val = percentile(z_value)
	if indicator == "bmi":
		if z_value < -3:
			status = "Ağır zayıflık aralığı"
		elif z_value < -2:
			status = "Zayıflık aralığı"
		elif z_value > 3:
			status = "Ağır obezite aralığı"
		elif z_value > 2:
			status = "Obezite aralığı"
		elif z_value > 1:
			status = "Fazla kiloluluk aralığı"
		else:
			status = "Referans aralığında (Normal)"
	elif indicator == "height":
		if z_value < -3:
			status = "Ağır boy kısalığı (bodurluk)"
		elif z_value < -2:
			status = "Boy kısalığı (bodurluk riski)"
		elif z_value > 3:
			status = "Belirgin uzun boy"
		elif z_value > 2:
			status = "Uzun boy aralığı"
		else:
			status = "Referans aralığında (Normal boy)"
	elif indicator == "head_circ":
		if z_value < -3:
			status = "Ağır düşük baş çevresi (Mikrosefali riski)"
		elif z_value < -2:
			status = "Düşük baş çevresi sınırı"
		elif z_value > 3:
			status = "Ağır yüksek baş çevresi (Makrosefali riski)"
		elif z_value > 2:
			status = "Yüksek baş çevresi sınırı"
		else:
			status = "Referans aralığında (Normal baş çevresi)"
		if int(months) > 60:
			status += " (60 ay WHO tavan referansı)"
	else:  # weight
		if z_value < -3:
			status = "Ağır düşük kilo aralığı"
		elif z_value < -2:
			status = "Düşük kilo aralığı"
		elif z_value > 2:
			status = "Yüksek kilo aralığı"
		else:
			status = "Referans aralığında (Normal kilo)"
	return {"z": round(z_value, 2), "percentile": round(pct_val, 1), "status": status, "median": round(reference["m"], 1)}


def get_lms_value_at_z(l: float, m: float, s: float, z: float) -> float:
	if l == 0:
		return m * math.exp(s * z)
	return m * ((1.0 + l * s * z) ** (1.0 / l))


def get_percentile_band(pct: float) -> str:
	if pct < 3.0:
		return "< 3p (Çok Düşük / Kritik Sınır)"
	elif pct < 10.0:
		return "3 - 10p (Düşük Sınır)"
	elif pct < 25.0:
		return "10 - 25p (Alt Normal Bant)"
	elif pct <= 75.0:
		return "25 - 75p (Orta Medyan / İdeal)"
	elif pct <= 90.0:
		return "75 - 90p (Üst Normal Bant)"
	elif pct <= 97.0:
		return "90 - 97p (Yüksek Sınır)"
	else:
		return "> 97p (Çok Yüksek / Kritik Sınır)"



def calculate_ideal_weight(
	height_cm: float | None,
	sex: str | None = None,
	age_months: int | None = None,
	current_weight_kg: float | None = None,
) -> dict | None:
	"""
	Klinik ve Antropometrik İdeal Kilo ve Sağlıklı Kilo Aralığı Hesaplama Motoru (WHO / ÇEDD Standardı).
	- Çocuk ve ergenlerde (0-228 ay / 0-19 yaş):
	  Boya Göre İdeal Kilo (Ideal Body Weight for Height - McLaren / WHO Standardı):
	  Yaş ve cinsiyete karşılık gelen 50. persentil (medyan) VKİ değeri (M_bmi) ile çocuğun boyu baz alınır:
	  İdeal Kilo = M_bmi * (height_cm / 100)^2
	  Sağlıklı Kilo Aralığı: -2 SDS ile +1 SDS arasındaki VKİ değerleri
	- Yetişkinlerde (> 228 ay veya doğum tarihi girilmemişse):
	  WHO Yetişkin Normal VKİ Aralığı (18.5 – 24.9 kg/m²), Standart İdeal VKİ = 22.0 kg/m²
	  İdeal Kilo = 22.0 * (height_cm / 100)^2
	- Durum & Fark Değerlendirmesi:
	  diff_kg = current_weight_kg - ideal_kg
	"""
	if height_cm is not None and float(height_cm) > 0:
		h = float(height_cm)
		h_m = h / 100.0
		h_sq = h_m ** 2
		is_ped = (sex in ("K", "E") and age_months is not None and 0 <= int(age_months) <= 228)
		if is_ped:
			b_lms = get_lms_row("bmi", sex, int(age_months))
			if b_lms:
				med_bmi = float(b_lms["m"])
				ideal_kg = med_bmi * h_sq
				low_bmi = get_lms_value_at_z(float(b_lms["l"]), float(b_lms["m"]), float(b_lms["s"]), -2.0)
				high_bmi = get_lms_value_at_z(float(b_lms["l"]), float(b_lms["m"]), float(b_lms["s"]), 1.0)
				min_kg = low_bmi * h_sq
				max_kg = high_bmi * h_sq
				ref_bmi = med_bmi
				is_pediatric = True
			else:
				ref_bmi = 22.0
				ideal_kg = 22.0 * h_sq
				min_kg = 18.5 * h_sq
				max_kg = 24.9 * h_sq
				is_pediatric = False
		else:
			ref_bmi = 22.0
			ideal_kg = 22.0 * h_sq
			min_kg = 18.5 * h_sq
			max_kg = 24.9 * h_sq
			is_pediatric = False
	elif current_weight_kg is not None and float(current_weight_kg) > 0 and sex in ("K", "E") and age_months is not None and 0 <= int(age_months) <= 228:
		w_lms = get_lms_row("weight", sex, int(age_months))
		if w_lms:
			ideal_kg = float(w_lms["m"])
			min_kg = get_lms_value_at_z(float(w_lms["l"]), float(w_lms["m"]), float(w_lms["s"]), -2.0)
			max_kg = get_lms_value_at_z(float(w_lms["l"]), float(w_lms["m"]), float(w_lms["s"]), 2.0)
			ref_bmi = None
			is_pediatric = True
		else:
			return None
	else:
		return None

	diff_kg = None
	status_text = ""
	if current_weight_kg is not None and float(current_weight_kg) > 0:
		w_cur = float(current_weight_kg)
		diff_kg = round(w_cur - ideal_kg, 1)
		if abs(diff_kg) <= 0.5:
			status_text = f"🟢 İdeal Kiloda ({min_kg:.1f} – {max_kg:.1f} kg)"
		elif diff_kg > 0.5:
			if w_cur > max_kg:
				diff_over = round(w_cur - max_kg, 1)
				status_text = f"🟡 +{diff_over:.1f} kg En Çoğun Üzerinde"
			else:
				status_text = f"🟢 İdeal Aralıkta ({min_kg:.1f} – {max_kg:.1f} kg)"
		else:
			if w_cur < min_kg:
				diff_under = round(min_kg - w_cur, 1)
				status_text = f"🟠 {diff_under:.1f} kg En Azın Altında"
			else:
				status_text = f"🟢 İdeal Aralıkta ({min_kg:.1f} – {max_kg:.1f} kg)"
	else:
		status_text = f"İdeal Aralık: En Az {min_kg:.1f} kg – En Çok {max_kg:.1f} kg"

	return {
		"ideal_kg": round(ideal_kg, 1),
		"min_kg": round(min_kg, 1),
		"max_kg": round(max_kg, 1),
		"band_str": f"{min_kg:.1f} – {max_kg:.1f} kg",
		"range_str": f"İdeal En Az: {min_kg:.1f} kg — İdeal En Çok: {max_kg:.1f} kg",
		"diff_kg": diff_kg,
		"status_text": status_text,
		"reference_bmi": round(ref_bmi, 1) if ref_bmi else None,
		"is_pediatric": is_pediatric,
	}


def get_clinical_growth_summary_table(
	height_cm: float | None,
	weight_kg: float | None,
	bmi: float | None,
	sex: str | None,
	months: int | None,
	head_circ_cm: float | None = None,
	always_include_head_circ: bool = False,
) -> pd.DataFrame:
	rows = []

	# 1. Height
	h_lms = get_lms_row("height", sex, months) if (sex and months is not None) else None
	if height_cm is not None:
		if h_lms:
			z = z_score(height_cm, h_lms["l"], h_lms["m"], h_lms["s"])
			pct = percentile(z)
			med = round(float(h_lms["m"]), 1)
			low_norm = get_lms_value_at_z(h_lms["l"], h_lms["m"], h_lms["s"], -2)
			high_norm = get_lms_value_at_z(h_lms["l"], h_lms["m"], h_lms["s"], 2)
			ref_band = f"{low_norm:.1f} – {high_norm:.1f} cm"
			z_val = round(float(z), 2)
			pct_val = round(float(pct), 1)
			band_str = get_percentile_band(pct)
			if z < -3:
				status = "🔴 Ağır Boy Kısalığı (Bodurluk)"
			elif z < -2:
				status = "🟠 Boy Kısalığı Riski (Bodurluk Sınırı)"
			elif z > 2:
				status = "🟢 Uzun Boy"
			else:
				status = "🟢 Yaşa Uygun Normal Boy"
		else:
			med, ref_band, z_val, pct_val, band_str, status = None, "—", None, None, "—", "Doğum tarihi / cinsiyet gerekli"
		rows.append({
			"Antropometrik Parametre": "📏 Boy (cm)",
			"Ölçülen Değer": round(float(height_cm), 1),
			"50. Persentil (Medyan)": med,
			"Normal Referans Bandı (-2 ile +2 SDS)": ref_band,
			"Z-Skoru (SDS)": z_val,
			"Persentil (%)": pct_val,
			"Persentil Dilimi": band_str,
			"Klinik Değerlendirme": status,
		})

	# 2. Weight
	w_lms = get_lms_row("weight", sex, months) if (sex and months is not None) else None
	if weight_kg is not None:
		if w_lms:
			z = z_score(weight_kg, w_lms["l"], w_lms["m"], w_lms["s"])
			pct = percentile(z)
			med = round(float(w_lms["m"]), 1)
			low_norm = get_lms_value_at_z(w_lms["l"], w_lms["m"], w_lms["s"], -2)
			high_norm = get_lms_value_at_z(w_lms["l"], w_lms["m"], w_lms["s"], 2)
			ref_band = f"{low_norm:.1f} – {high_norm:.1f} kg"
			z_val = round(float(z), 2)
			pct_val = round(float(pct), 1)
			band_str = get_percentile_band(pct)
			if z < -3:
				status = "🔴 Ağır Düşük Kilo"
			elif z < -2:
				status = "🟠 Düşük Kilo"
			elif z > 2:
				status = "🟠 Yüksek Kilo"
			else:
				status = "🟢 Yaşa Uygun Normal Kilo"
		else:
			med, ref_band, z_val, pct_val, band_str, status = None, "—", None, None, "—", "Doğum tarihi / cinsiyet gerekli"
		rows.append({
			"Antropometrik Parametre": "⚖️ Kilo (kg)",
			"Ölçülen Değer": round(float(weight_kg), 1),
			"50. Persentil (Medyan)": med,
			"Normal Referans Bandı (-2 ile +2 SDS)": ref_band,
			"Z-Skoru (SDS)": z_val,
			"Persentil (%)": pct_val,
			"Persentil Dilimi": band_str,
			"Klinik Değerlendirme": status,
		})

	# 3. BMI
	b_lms = get_lms_row("bmi", sex, months) if (sex and months is not None) else None
	if bmi is not None:
		if b_lms:
			z = z_score(bmi, b_lms["l"], b_lms["m"], b_lms["s"])
			pct = percentile(z)
			med = round(float(b_lms["m"]), 1)
			low_norm = get_lms_value_at_z(b_lms["l"], b_lms["m"], b_lms["s"], -2)
			high_norm = get_lms_value_at_z(b_lms["l"], b_lms["m"], b_lms["s"], 2)
			ref_band = f"{low_norm:.1f} – {high_norm:.1f} kg/m²"
			z_val = round(float(z), 2)
			pct_val = round(float(pct), 1)
			band_str = get_percentile_band(pct)
			if z < -3:
				status = "🔴 Şiddetli Zayıflık"
			elif z < -2:
				status = "🟠 Zayıflık"
			elif z > 3:
				status = "🔴 Ağır Obezite"
			elif z > 2:
				status = "🔴 Obezite"
			elif z > 1:
				status = "🟡 Fazla Kiloluluk"
			else:
				status = "🟢 Sağlıklı / Normal Ağırlık"
		else:
			med, ref_band, z_val, pct_val, band_str, status = None, "—", None, None, "—", "Doğum tarihi / cinsiyet gerekli"
		rows.append({
			"Antropometrik Parametre": "🧬 VKİ (BMI, kg/m²)",
			"Ölçülen Değer": round(float(bmi), 1),
			"50. Persentil (Medyan)": med,
			"Normal Referans Bandı (-2 ile +2 SDS)": ref_band,
			"Z-Skoru (SDS)": z_val,
			"Persentil (%)": pct_val,
			"Persentil Dilimi": band_str,
			"Klinik Değerlendirme": status,
		})

	# 3b. İdeal Kilo Aralığı (Boy ve Yaşa Göre İdeal En Az – İdeal En Çok Kilo)
	ideal_eval = calculate_ideal_weight(height_cm, sex, months, current_weight_kg=weight_kg)
	if ideal_eval is not None:
		diff_str = f" ({ideal_eval['diff_kg']:+.1f} kg)" if ideal_eval["diff_kg"] is not None else ""
		rows.append({
			"Antropometrik Parametre": "🎯 İdeal Kilo Aralığı (kg)",
			"Ölçülen Değer": round(float(weight_kg), 1) if weight_kg is not None else None,
			"50. Persentil (Medyan)": ideal_eval["ideal_kg"],
			"Normal Referans Bandı (-2 ile +2 SDS)": f"En Az {ideal_eval['min_kg']:.1f} kg – En Çok {ideal_eval['max_kg']:.1f} kg",
			"Z-Skoru (SDS)": None,
			"Persentil (%)": None,
			"Persentil Dilimi": ideal_eval["status_text"],
			"Klinik Değerlendirme": f"🎯 İdeal En Az: {ideal_eval['min_kg']:.1f} kg — İdeal En Çok: {ideal_eval['max_kg']:.1f} kg · {ideal_eval['status_text']}",
		})

	# 4. Head Circumference (Baş Çevresi) - Yalnızca 0-60 ay (0-5 yaş) arasında değerlendirilebilir!
	can_eval_hc_age = (months is not None and 0 <= int(months) <= 60)
	hc_lms = get_lms_row("head_circ", sex, months) if (sex and can_eval_hc_age) else None
	if can_eval_hc_age and head_circ_cm is not None and not pd.isna(head_circ_cm) and float(head_circ_cm) > 0:
		hc_f = float(head_circ_cm)
		if hc_lms:
			z = z_score(hc_f, hc_lms["l"], hc_lms["m"], hc_lms["s"])
			pct = percentile(z)
			med = round(float(hc_lms["m"]), 1)
			low_norm = get_lms_value_at_z(hc_lms["l"], hc_lms["m"], hc_lms["s"], -2)
			high_norm = get_lms_value_at_z(hc_lms["l"], hc_lms["m"], hc_lms["s"], 2)
			ref_band = f"{low_norm:.1f} – {high_norm:.1f} cm"
			z_val = round(float(z), 2)
			pct_val = round(float(pct), 1)
			band_str = get_percentile_band(pct)
			if z < -3:
				status = "🔴 Ağır Düşük Baş Çevresi (Mikrosefali Sınırı)"
			elif z < -2:
				status = "🟠 Düşük Baş Çevresi"
			elif z > 3:
				status = "🔴 Ağır Yüksek Baş Çevresi (Makrosefali Sınırı)"
			elif z > 2:
				status = "🟠 Yüksek Baş Çevresi"
			else:
				status = "🟢 Yaşa Uygun Normal Baş Çevresi"
		else:
			med, ref_band, z_val, pct_val, band_str = None, "—", None, None, "—"
			status = f"Kaydedildi ({hc_f:.1f} cm)" if (sex and months is not None) else "Doğum tarihi / cinsiyet gerekli"
		rows.append({
			"Antropometrik Parametre": "🧠 Baş Çevresi (cm)",
			"Ölçülen Değer": round(hc_f, 1),
			"50. Persentil (Medyan)": med,
			"Normal Referans Bandı (-2 ile +2 SDS)": ref_band,
			"Z-Skoru (SDS)": z_val,
			"Persentil (%)": pct_val,
			"Persentil Dilimi": band_str,
			"Klinik Değerlendirme": status,
		})
	elif always_include_head_circ and can_eval_hc_age:
		if hc_lms:
			med = round(float(hc_lms["m"]), 1)
			low_norm = get_lms_value_at_z(hc_lms["l"], hc_lms["m"], hc_lms["s"], -2)
			high_norm = get_lms_value_at_z(hc_lms["l"], hc_lms["m"], hc_lms["s"], 2)
			ref_band = f"{low_norm:.1f} – {high_norm:.1f} cm"
		else:
			med, ref_band = None, "—"
		rows.append({
			"Antropometrik Parametre": "🧠 Baş Çevresi (cm)",
			"Ölçülen Değer": None,
			"50. Persentil (Medyan)": med,
			"Normal Referans Bandı (-2 ile +2 SDS)": ref_band,
			"Z-Skoru (SDS)": None,
			"Persentil (%)": None,
			"Persentil Dilimi": "—",
			"Klinik Değerlendirme": "— (Girilmedi; çift tıklayarak ekleyebilirsiniz)",
		})

	return pd.DataFrame(rows)


def get_clinical_growth_column_config() -> dict:
	return {
		"Ölçülen Değer": st.column_config.NumberColumn(
			"✏️ Ölçülen Değer (Düzenlenebilir)",
			help="Boy, kilo veya baş çevresi değerine çift tıklayarak doğrudan tablo üzerinden değiştirebilirsiniz.",
			min_value=1.0,
			max_value=300.0,
			step=0.1,
			format="%.1f",
		),
		"50. Persentil (Medyan)": st.column_config.NumberColumn("50. Persentil (Medyan)", format="%.1f"),
		"Z-Skoru (SDS)": st.column_config.NumberColumn("Z-Skoru (SDS)", format="%+.2f SDS"),
		"Persentil (%)": st.column_config.NumberColumn("Persentil (%)", format="%%%.1f"),
	}


def generate_who_percentile_curves(indicator: str, sex: str, min_month: int = 0, max_month: int = 72) -> pd.DataFrame:
	_ensure_lms_cache()
	percentiles = [
		(3, -1.88079, "3p (Kritik Alt)"),
		(15, -1.03643, "15p (Alt Normal)"),
		(50, 0.0, "50p (Medyan / Ortalama)"),
		(85, 1.03643, "85p (Üst Normal)"),
		(97, 1.88079, "97p (Kritik Üst)"),
	]
	rows = []
	for m in range(min_month, max_month + 1):
		lms = get_lms_row(indicator, sex, m)
		if not lms:
			continue
		l, m_val, s = lms["l"], lms["m"], lms["s"]
		for p_val, z_val, p_lbl in percentiles:
			val = get_lms_value_at_z(l, m_val, s, z_val)
			rows.append({
				"Yas_Ay": m,
				"Persantil": p_val,
				"Egri_Etiketi": p_lbl,
				"Deger": round(val, 2),
			})
	return pd.DataFrame(rows)


def get_who_reference_cutoffs_table(sex: str, months: int, height_cm: float | None, weight_kg: float | None, bmi: float | None, head_circ_cm: float | None = None) -> pd.DataFrame:
	rows = []
	for ind, name, unit, cur_val in [
		("height", "📏 Boy", "cm", height_cm),
		("weight", "⚖️ Kilo", "kg", weight_kg),
		("bmi", "🧬 VKİ (BMI)", "kg/m²", bmi),
		("head_circ", "🧠 Baş Çevresi", "cm", head_circ_cm),
	]:
		lms = get_lms_row(ind, sex, months)
		if not lms:
			continue
		vals = {z: round(get_lms_value_at_z(lms["l"], lms["m"], lms["s"], z), 1) for z in [-3, -2, -1, 0, 1, 2, 3]}
		rows.append({
			"Parametre": f"{name} ({unit})",
			"-3 SDS (0.1p)": vals[-3],
			"-2 SDS (3p Alt Sınır)": vals[-2],
			"-1 SDS (15p)": vals[-1],
			"0 SDS (50p Medyan)": vals[0],
			"+1 SDS (85p)": vals[1],
			"+2 SDS (97p Üst Sınır)": vals[2],
			"+3 SDS (99.9p)": vals[3],
			"Öğrencinin Ölçümü": round(float(cur_val), 1) if (cur_val is not None and not pd.isna(cur_val)) else None,
		})
	return pd.DataFrame(rows)


def get_who_cutoffs_column_config() -> dict:
	return {
		"-3 SDS (0.1p)": st.column_config.NumberColumn("-3 SDS (0.1p)", format="%.1f"),
		"-2 SDS (3p Alt Sınır)": st.column_config.NumberColumn("-2 SDS (3p Alt Sınır)", format="%.1f"),
		"-1 SDS (15p)": st.column_config.NumberColumn("-1 SDS (15p)", format="%.1f"),
		"0 SDS (50p Medyan)": st.column_config.NumberColumn("0 SDS (50p Medyan)", format="%.1f"),
		"+1 SDS (85p)": st.column_config.NumberColumn("+1 SDS (85p)", format="%.1f"),
		"+2 SDS (97p Üst Sınır)": st.column_config.NumberColumn("+2 SDS (97p Üst Sınır)", format="%.1f"),
		"+3 SDS (99.9p)": st.column_config.NumberColumn("+3 SDS (99.9p)", format="%.1f"),
		"Öğrencinin Ölçümü": st.column_config.NumberColumn("Öğrencinin Ölçümü", format="%.1f"),
	}


def enrich_measurements_with_percentiles(records: pd.DataFrame, sex: str | None, birth_date: str | None = None) -> pd.DataFrame:
	if records.empty:
		return records.copy()
	df = records.copy()
	sort_cols = ["Tarih", "id"] if "id" in df.columns else ["Tarih"]
	df = df.sort_values(by=sort_cols, ascending=True, kind="mergesort").reset_index(drop=True)

	out_rows = []
	prev_h, prev_w = None, None
	for _, row in df.iterrows():
		raw_date = str(row.get("Tarih", ""))[:10]
		months = int(row["Yas_ay"]) if (row.get("Yas_ay") is not None and not pd.isna(row.get("Yas_ay"))) else None
		if months is None and birth_date and raw_date:
			try:
				months = age_in_months(birth_date, date.fromisoformat(raw_date))
			except Exception:
				months = None

		h_val = round(float(row["Boy_cm"]), 1) if (row.get("Boy_cm") is not None and not pd.isna(row.get("Boy_cm"))) else None
		w_val = round(float(row["Kilo_kg"]), 1) if (row.get("Kilo_kg") is not None and not pd.isna(row.get("Kilo_kg"))) else None
		hc_raw = row.get("Bas_Cevresi_cm") if "Bas_Cevresi_cm" in row else row.get("head_circ_cm")
		hc_val = round(float(hc_raw), 1) if (hc_raw is not None and not pd.isna(hc_raw) and float(hc_raw) > 0) else None
		b_val = round(float(row["BMI"]), 1) if (row.get("BMI") is not None and not pd.isna(row.get("BMI"))) else None
		if b_val is None and h_val and w_val and h_val > 0 and w_val > 0:
			b_val = round(w_val / ((h_val / 100.0) ** 2), 1)

		diff_boy = round(h_val - prev_h, 1) if (prev_h is not None and h_val is not None) else None
		diff_kilo = round(w_val - prev_w, 1) if (prev_w is not None and w_val is not None) else None
		prev_h, prev_w = h_val, w_val

		h_eval = evaluate_measurement(h_val, "height", sex, months) if (sex and months is not None and h_val is not None) else None
		w_eval = evaluate_measurement(w_val, "weight", sex, months) if (sex and months is not None and w_val is not None) else None
		b_eval = evaluate_measurement(b_val, "bmi", sex, months) if (sex and months is not None and b_val is not None) else None
		hc_eval = evaluate_measurement(hc_val, "head_circ", sex, months) if (sex and months is not None and hc_val is not None) else None
		ideal_eval = calculate_ideal_weight(h_val, sex, months, current_weight_kg=w_val)

		if sex and months is not None:
			remarks = []
			if b_eval and "Normal" not in b_eval["status"]:
				remarks.append(b_eval["status"].replace(" aralığı", ""))
			if h_eval and "Normal" not in h_eval["status"]:
				remarks.append(h_eval["status"].replace(" aralığı", ""))
			if hc_eval and "Normal" not in hc_eval["status"]:
				remarks.append(hc_eval["status"].replace(" sınırı", ""))
			if w_eval and "Normal" not in w_eval["status"] and not remarks:
				remarks.append(w_eval["status"].replace(" aralığı", ""))
			status_str = ", ".join(remarks) if remarks else "🟢 Sağlıklı / Normal gelişim"
		else:
			status_str = "⚪ Bilgi eksik (Doğum tarihi / Cinsiyet)"

		row_dict = dict(row)
		row_dict.update({
			"Yas_ay": months,
			"Boy_cm": h_val,
			"Boy_Z": h_eval["z"] if h_eval else None,
			"Boy_Persantil": h_eval["percentile"] if h_eval else None,
			"Boy_Dilimi": get_percentile_band(h_eval["percentile"]) if h_eval else "—",
			"Boy_Farki": diff_boy,
			"Kilo_kg": w_val,
			"Kilo_Z": w_eval["z"] if w_eval else None,
			"Kilo_Persantil": w_eval["percentile"] if w_eval else None,
			"Kilo_Dilimi": get_percentile_band(w_eval["percentile"]) if w_eval else "—",
			"Kilo_Farki": diff_kilo,
			"Ideal_Kilo_kg": ideal_eval["ideal_kg"] if ideal_eval else None,
			"Ideal_Kilo_Farki": ideal_eval["diff_kg"] if ideal_eval else None,
			"Ideal_En_Az_kg": ideal_eval["min_kg"] if ideal_eval else None,
			"Ideal_En_Cok_kg": ideal_eval["max_kg"] if ideal_eval else None,
			"Ideal_Kilo_Araligi": f"{ideal_eval['min_kg']:.1f} – {ideal_eval['max_kg']:.1f} kg" if ideal_eval else "—",
			"Ideal_Kilo_Durumu": ideal_eval["status_text"] if ideal_eval else "—",
			"Bas_Cevresi_cm": hc_val,
			"Bas_Cevresi_Z": hc_eval["z"] if hc_eval else None,
			"Bas_Cevresi_Persantil": hc_eval["percentile"] if hc_eval else None,
			"BMI": b_val,
			"BMI_Z": b_eval["z"] if b_eval else None,
			"BMI_Persantil": b_eval["percentile"] if b_eval else None,
			"BMI_Dilimi": get_percentile_band(b_eval["percentile"]) if b_eval else "—",
			"Gelisim_Degerlendirmesi": status_str,
		})
		out_rows.append(row_dict)

	return pd.DataFrame(out_rows)


def get_measurement_history_column_config() -> dict:
	return {
		"id": None,
		"Tarih": st.column_config.TextColumn("Tarih"),
		"Yaş (Ay)": st.column_config.NumberColumn("Yaş (Ay)", format="%d ay"),
		"Boy (cm)": st.column_config.NumberColumn("Boy (cm)", format="%.1f cm"),
		"Boy Z-Skor": st.column_config.NumberColumn("Boy Z-Skor", format="%+.2f SDS"),
		"Boy %": st.column_config.NumberColumn("Boy %", format="%%%.1f"),
		"Boy Dilimi": st.column_config.TextColumn("Boy Dilimi"),
		"Boy Değişimi": st.column_config.NumberColumn("Boy Değişimi", format="%+.1f cm"),
		"Kilo (kg)": st.column_config.NumberColumn("Kilo (kg)", format="%.1f kg"),
		"Kilo Z-Skor": st.column_config.NumberColumn("Kilo Z-Skor", format="%+.2f SDS"),
		"Kilo %": st.column_config.NumberColumn("Kilo %", format="%%%.1f"),
		"Kilo Dilimi": st.column_config.TextColumn("Kilo Dilimi"),
		"Kilo Değişimi": st.column_config.NumberColumn("Kilo Değişimi", format="%+.1f kg"),
		"İdeal Kilo Aralığı": st.column_config.TextColumn("🎯 İdeal Kilo Aralığı (En Az – En Çok)"),
		"İdeal En Az (kg)": st.column_config.NumberColumn("İdeal En Az (kg)", format="%.1f kg"),
		"İdeal En Çok (kg)": st.column_config.NumberColumn("İdeal En Çok (kg)", format="%.1f kg"),
		"İdeal Kilo (kg)": st.column_config.NumberColumn("Medyan Kilo (kg)", format="%.1f kg"),
		"İdeal Kilo Farkı": st.column_config.NumberColumn("Fark (kg)", format="%+.1f kg"),
		"Baş Çevresi (cm)": st.column_config.NumberColumn("Baş Çevresi (cm)", format="%.1f cm"),
		"Baş Çevresi Z-Skor": st.column_config.NumberColumn("Baş Çevresi Z-Skor", format="%+.2f SDS"),
		"Baş Çevresi %": st.column_config.NumberColumn("Baş Çevresi %", format="%%%.1f"),
		"VKİ (BMI)": st.column_config.NumberColumn("VKİ (BMI)", format="%.1f kg/m²"),
		"BMI": st.column_config.NumberColumn("VKİ (BMI)", format="%.1f kg/m²"),
		"VKİ Z-Skor": st.column_config.NumberColumn("VKİ Z-Skor", format="%+.2f SDS"),
		"BMI Z-Skor": st.column_config.NumberColumn("BMI Z-Skor", format="%+.2f SDS"),
		"VKİ %": st.column_config.NumberColumn("VKİ %", format="%%%.1f"),
		"BMI %": st.column_config.NumberColumn("BMI %", format="%%%.1f"),
		"Gelişim Değerlendirmesi": st.column_config.TextColumn("Gelişim Değerlendirmesi"),
	}


def format_measurements_for_html(df: pd.DataFrame) -> pd.DataFrame:
	out = df.copy()
	fmt_map = {
		"Boy (cm)": lambda v: f"{v:.1f} cm",
		"Boy Z-Skor": lambda v: f"{v:+.2f} SDS",
		"Boy %": lambda v: f"%{v:.1f}",
		"Boy Değişimi": lambda v: f"{v:+.1f} cm",
		"Kilo (kg)": lambda v: f"{v:.1f} kg",
		"Kilo Z-Skor": lambda v: f"{v:+.2f} SDS",
		"Kilo %": lambda v: f"%{v:.1f}",
		"Kilo Değişimi": lambda v: f"{v:+.1f} kg",
		"Baş Çevresi (cm)": lambda v: f"{v:.1f} cm",
		"Baş Çevresi Z-Skor": lambda v: f"{v:+.2f} SDS",
		"Baş Çevresi %": lambda v: f"%{v:.1f}",
		"VKİ (BMI)": lambda v: f"{v:.1f}",
		"BMI": lambda v: f"{v:.1f}",
		"VKİ Z-Skor": lambda v: f"{v:+.2f} SDS",
		"BMI Z-Skor": lambda v: f"{v:+.2f} SDS",
		"VKİ %": lambda v: f"%{v:.1f}",
		"BMI %": lambda v: f"%{v:.1f}",
	}
	for col, fn in fmt_map.items():
		if col in out.columns:
			out[col] = out[col].apply(lambda x: fn(float(x)) if pd.notna(x) else "—")
	return out


def parse_lms_file(uploaded_file) -> tuple[int, list[str]]:
	if uploaded_file.name.lower().endswith(".csv"):
		frame = pd.read_csv(uploaded_file)
	else:
		frame = pd.read_excel(uploaded_file)
	aliases = {
		"gosterge": "indicator", "indicator": "indicator", "olcum": "indicator", "indeks": "indicator",
		"cinsiyet": "sex", "sex": "sex", "gender": "sex",
		"yas ay": "age_months", "yas_ay": "age_months", "age months": "age_months", "age_months": "age_months",
		"l": "l", "m": "m", "s": "s",
	}
	frame = frame.rename(columns={column: aliases.get(normalize_header(column), aliases.get(clean_text(column).lower())) for column in frame.columns})
	frame = frame.loc[:, ~frame.columns.isna()]
	required = {"indicator", "sex", "age_months", "l", "m", "s"}
	if not required.issubset(frame.columns):
		raise ValueError("Referans dosyasında Gösterge, Cinsiyet, Yaş_Ay, L, M ve S sütunları olmalı.")
	indicator_map = {
		"bmi": "bmi", "bmi for age": "bmi", "bmi yas": "bmi", "bmi-for-age": "bmi",
		"boy": "height", "height": "height", "height for age": "height", "boy yas": "height", "height-for-age": "height",
		"kilo": "weight", "weight": "weight", "weight for age": "weight", "kilo yas": "weight", "weight-for-age": "weight",
		"bas cevresi": "head_circ", "head circumference": "head_circ", "head_circ": "head_circ", "hcfa": "head_circ",
	}
	issues: list[str] = []
	inserted = 0
	with connect_db() as connection:
		for index, row in frame.iterrows():
			indicator = indicator_map.get(normalize_header(row["indicator"]))
			sex = normalize_sex(row["sex"])
			try:
				months = int(row["age_months"])
				l_value, m_value, s_value = float(row["l"]), float(row["m"]), float(row["s"])
				if not indicator or not sex or m_value <= 0 or s_value <= 0:
					raise ValueError
			except (ValueError, TypeError, OverflowError):
				issues.append(f"Satır {index + 2}: gösterge, cinsiyet, yaş veya LMS değerleri geçersiz.")
				continue
			connection.execute(
				"""INSERT INTO lms_reference(indicator, sex, age_months, l, m, s) VALUES (?, ?, ?, ?, ?, ?)
				   ON CONFLICT(indicator, sex, age_months) DO UPDATE SET l=excluded.l, m=excluded.m, s=excluded.s""",
				(indicator, sex, months, l_value, m_value, s_value),
			)
			inserted += 1
	_LMS_CACHE.clear()
	return inserted, issues


def import_who_wfa_reference_files() -> int:
	sources = {
		"E": "wfa-boys-zscore-expanded-tables.xlsx",
		"K": "wfa-girls-zscore-expanded-tables.xlsx",
	}
	folder = APP_DIR / "persantil veriler"
	if not all((folder / filename).is_file() for filename in sources.values()):
		return 0
	rows = []
	for sex, filename in sources.items():
		frame = pd.read_excel(folder / filename)
		columns = {normalize_header(column): column for column in frame.columns}
		age_column = next((columns[name] for name in ("day", "days", "age days") if name in columns), None)
		if age_column is None or not all(name in columns for name in ("l", "m", "s")):
			raise ValueError(f"WHO WFA tablosunun sütunları tanınmadı: {filename}")
		selected = frame[[age_column, columns["l"], columns["m"], columns["s"]]].copy()
		selected.columns = ["age_days", "l", "m", "s"]
		for column in selected.columns:
			selected[column] = pd.to_numeric(selected[column], errors="coerce")
		selected = selected.dropna()
		if (len(selected) != 1857 or selected["age_days"].duplicated().any()
				or selected["age_days"].min() != 0 or selected["age_days"].max() != 1856
				or (selected["m"] <= 0).any() or (selected["s"] <= 0).any()):
			raise ValueError(f"WHO WFA yaş veya LMS satırları eksik/geçersiz: {filename}")
		rows.extend((sex, int(row.age_days), float(row.l), float(row.m), float(row.s), filename)
					for row in selected.itertuples(index=False))
	with connect_db() as connection:
		connection.executemany(
			"""INSERT INTO wfa_reference(sex, age_days, l, m, s, source) VALUES (?, ?, ?, ?, ?, ?)
			   ON CONFLICT(sex, age_days) DO UPDATE SET l=excluded.l, m=excluded.m, s=excluded.s, source=excluded.source""",
			rows,
		)
	return len(rows)


def import_builtin_who_references() -> tuple[int, int]:
	folder = APP_DIR / "persantil veriler"
	inserted_lms = 0
	sources = [
		("tab_wfa_boys_p_0_5.xlsx", "weight", "E"),
		("tab_wfa_girls_p_0_5.xlsx", "weight", "K"),
		("wfa_boys_5_10.xlsx", "weight", "E"),
		("wfa_girls_5_10.xlsx", "weight", "K"),
		("tab_lhfa_boys_0_2.xlsx", "height", "E"),
		("tab_lhfa_girls_0_2.xlsx", "height", "K"),
		("tab_lhfa_boys_2_5.xlsx", "height", "E"),
		("tab_lhfa_girls_2_5.xlsx", "height", "K"),
		("hfa_boys_5_19.xlsx", "height", "E"),
		("hfa_girls_5_19.xlsx", "height", "K"),
		("tab_bmi_boys_0_2.xlsx", "bmi", "E"),
		("tab_bmi_girls_0_2.xlsx", "bmi", "K"),
		("tab_bmi_boys_2_5.xlsx", "bmi", "E"),
		("tab_bmi_girls_2_5.xlsx", "bmi", "K"),
		("bmi_boys_5_19.xlsx", "bmi", "E"),
		("bmi_girls_5_19.xlsx", "bmi", "K"),
	]
	with connect_db() as connection:
		for filename, indicator, sex in sources:
			file_path = folder / filename
			if not file_path.is_file():
				continue
			try:
				df = pd.read_excel(file_path)
				cols = {str(c).strip().lower(): c for c in df.columns}
				m_col = cols.get("month")
				l_col = cols.get("l")
				m_val_col = cols.get("m")
				s_col = cols.get("s")
				if not (m_col and l_col and m_val_col and s_col):
					continue
				rows = []
				for _, r in df.iterrows():
					try:
						mo = int(r[m_col])
						l_v = float(r[l_col])
						m_v = float(r[m_val_col])
						s_v = float(r[s_col])
						rows.append((indicator, sex, mo, l_v, m_v, s_v))
					except Exception:
						continue
				connection.executemany(
					"""INSERT INTO lms_reference(indicator, sex, age_months, l, m, s) VALUES (?, ?, ?, ?, ?, ?)
					   ON CONFLICT(indicator, sex, age_months) DO UPDATE SET l=excluded.l, m=excluded.m, s=excluded.s""",
					rows,
				)
				inserted_lms += len(rows)
			except Exception:
				continue
		hc_rows = get_builtin_hcfa_lms_rows()
		connection.executemany(
			"""INSERT INTO lms_reference(indicator, sex, age_months, l, m, s) VALUES (?, ?, ?, ?, ?, ?)
			   ON CONFLICT(indicator, sex, age_months) DO UPDATE SET l=excluded.l, m=excluded.m, s=excluded.s""",
			hc_rows,
		)
		inserted_lms += len(hc_rows)
	_LMS_CACHE.clear()
	inserted_wfa = 0
	try:
		inserted_wfa = import_who_wfa_reference_files()
	except Exception:
		pass
	return inserted_lms, inserted_wfa


def make_roster_template() -> bytes:
	workbook = Workbook()
	sheet = workbook.active
	sheet.title = "Öğrenciler"
	sheet.append([
		"Sınıf", "Okul No", "Öğrenci Adı", "Cinsiyet", "Doğum Tarihi",
		"Öğrenci Köy Adı", "Öğrenci Mahalle Adı", "Veli Adı", "TELEFON",
	])
	sheet.append([
		"5-A", "1001", "Örnek Öğrenci", "K", "15.03.2015",
		"MERKEZ", "GEVHER HATUN MAHALLESİ 1986 SOKAK NO:1", "Veli Adı", "05xx xxx xx xx",
	])
	instructions = workbook.create_sheet("Bilgi")
	instructions.append(["Sütun", "Açıklama"])
	instructions.append(["Öğrenci Adı", "Zorunlu. Ad soyad birlikte yazılmalı."])
	instructions.append(["Sınıf", "İsteğe bağlı. Yoksa 'Öğrenci Köy Adı' (servis grubu) veya varsayılan grup atanır."])
	instructions.append(["Öğrenci Köy Adı", "İsteğe bağlı. Öğrencinin geldiği köy veya merkez bilgisi."])
	instructions.append(["Öğrenci Mahalle Adı", "İsteğe bağlı. Açık adres veya mahalle/sokak bilgisi."])
	instructions.append(["TELEFON / Veli Telefonu", "İsteğe bağlı. 10 haneli (5xx...) numaraların başına otomatik 0 eklenir."])
	instructions.append(["Cinsiyet", "K veya E. Boş bırakılırsa isimden otomatik tahmin edilir."])
	instructions.append(["Doğum Tarihi", "GG.AA.YYYY biçimi önerilir. Persantil hesabı için gereklidir."])
	instructions.append(["Okul No, Veli Adı", "İsteğe bağlı."])
	for tab in workbook.worksheets:
		tab.freeze_panes = "A2"
		tab.auto_filter.ref = tab.dimensions
		for cell in tab[1]:
			cell.font = Font(bold=True, color="FFFFFF")
			cell.fill = PatternFill("solid", fgColor="176B68")
		for column_cells in tab.columns:
			width = min(max(max(len(str(cell.value or "")) for cell in column_cells) + 3, 15), 42)
			tab.column_dimensions[get_column_letter(column_cells[0].column)].width = width
	output = BytesIO()
	workbook.save(output)
	return output.getvalue()


def workbook_download(students: pd.DataFrame, visits: pd.DataFrame, measurements: pd.DataFrame,
					  vitals: pd.DataFrame | None = None, lab_results: pd.DataFrame | None = None,
					  health_profiles: pd.DataFrame | None = None, care_plans: pd.DataFrame | None = None) -> bytes:
	output = BytesIO()
	with pd.ExcelWriter(output, engine="openpyxl") as writer:
		students.drop(columns=["id"], errors="ignore").to_excel(writer, sheet_name="Öğrenciler", index=False)
		visits.drop(columns=["id", "student_id"], errors="ignore").to_excel(writer, sheet_name="Revir Kayıtları", index=False)
		measurements.drop(columns=["id", "student_id"], errors="ignore").to_excel(writer, sheet_name="Ölçümler", index=False)
		(vitals if vitals is not None else pd.DataFrame()).drop(columns=["id", "student_id"], errors="ignore").to_excel(writer, sheet_name="Vital Takip", index=False)
		(lab_results if lab_results is not None else pd.DataFrame()).drop(columns=["id", "student_id", "report_id"], errors="ignore").to_excel(writer, sheet_name="Tahlil Sonuçları", index=False)
		(health_profiles if health_profiles is not None else pd.DataFrame()).drop(columns=["student_id"], errors="ignore").to_excel(writer, sheet_name="Sağlık Profilleri", index=False)
		(care_plans if care_plans is not None else pd.DataFrame()).drop(columns=["id", "student_id"], errors="ignore").to_excel(writer, sheet_name="Bakım Planları", index=False)
		for sheet in writer.book.worksheets:
			sheet.freeze_panes = "A2"
			sheet.auto_filter.ref = sheet.dimensions
			for cell in sheet[1]:
				cell.font = Font(bold=True, color="FFFFFF")
				cell.fill = PatternFill("solid", fgColor="176B68")
				cell.alignment = Alignment(wrap_text=True)
			for column_cells in sheet.columns:
				width = min(max(max(len(str(cell.value or "")) for cell in column_cells) + 2, 12), 36)
				sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = width
	return output.getvalue()


def student_options(class_name: str | None = None, person_type: str | None = None) -> list[sqlite3.Row]:
	query = """SELECT s.id, s.person_type, s.full_name, c.name AS class_name, s.sex, s.birth_date,
					  s.student_no, s.guardian_name, s.guardian_phone, s.notes
			   FROM students s JOIN classes c ON c.id=s.class_id WHERE 1=1"""
	params: list = []
	if class_name and class_name != "Tüm sınıflar":
		query += " AND c.name=?"
		params.append(class_name)
	if person_type:
		norm_pt = clean_text(person_type).lower()
		if "sivil" in norm_pt or "diger" in norm_pt:
			query += " AND (c.name = 'Sivil / Diğer' OR s.person_type LIKE '%Sivil%' OR s.person_type LIKE '%diğer%')"
		else:
			query += " AND (c.name != 'Sivil / Diğer' AND s.person_type NOT LIKE '%Sivil%' AND s.person_type NOT LIKE '%diğer%')"
	query += " ORDER BY c.name, s.full_name"
	with connect_db() as connection:
		return connection.execute(query, tuple(params)).fetchall()


def clear_all_data_caches() -> None:
	"""Veri değişikliklerinde önbellekleri temizler ve belleği boşaltır."""
	try:
		get_students_enriched.clear()
	except Exception:
		pass
	gc.collect()


@st.cache_data(ttl=60, show_spinner=False)
def get_students_enriched(class_name: str | None = None, person_type: str | None = None) -> pd.DataFrame:
	query = """
		SELECT s.id,
			   s.person_type AS kayit_turu,
			   c.name AS sinif,
			   s.student_no AS okul_no,
			   s.full_name AS ad_soyad,
			   s.sex AS cinsiyet,
			   s.birth_date AS dogum_tarihi,
			   s.notes AS notlar,
			   s.country AS ulke,
			   s.province AS il,
			   s.district AS ilce,
			   s.village AS koy,
			   COALESCE(h.guardian_name, s.guardian_name) AS veli,
			   COALESCE(h.guardian_phone, s.guardian_phone) AS veli_telefon,
			   h.blood_group AS kan_grubu,
			   h.allergies AS alerjiler,
			   h.chronic_conditions AS kronik,
			   h.medical_history AS saglik_gecmisi,
			   h.medications AS ilaclar,
			   (SELECT COUNT(*) FROM visits v WHERE v.student_id = s.id) AS basvuru_sayisi,
			   COALESCE(
				   m.measured_on,
				   (SELECT m3.measured_on FROM measurements m3 WHERE m3.student_id = s.id ORDER BY m3.measured_on DESC, m3.id DESC LIMIT 1)
			   ) AS son_olcum_tarihi,
			   COALESCE(
				   m.height_cm,
				   (SELECT m3.height_cm FROM measurements m3 WHERE m3.student_id = s.id AND m3.height_cm IS NOT NULL ORDER BY m3.measured_on DESC, m3.id DESC LIMIT 1)
			   ) AS boy_cm,
			   COALESCE(
				   m.weight_kg,
				   (SELECT m3.weight_kg FROM measurements m3 WHERE m3.student_id = s.id AND m3.weight_kg IS NOT NULL ORDER BY m3.measured_on DESC, m3.id DESC LIMIT 1)
			   ) AS kilo_kg,
			   COALESCE(
				   m.head_circ_cm,
				   (SELECT m3.head_circ_cm FROM measurements m3 WHERE m3.student_id = s.id AND m3.head_circ_cm IS NOT NULL ORDER BY m3.measured_on DESC, m3.id DESC LIMIT 1)
			   ) AS bas_cevresi_cm,
			   m.age_months AS olcum_yas_ay,
			   m.bmi AS vki
		FROM students s
		JOIN classes c ON c.id = s.class_id
		LEFT JOIN health_profiles h ON h.student_id = s.id
		LEFT JOIN measurements m ON m.id = (
			SELECT m2.id FROM measurements m2
			WHERE m2.student_id = s.id
			ORDER BY m2.measured_on DESC, m2.id DESC LIMIT 1
		)
		WHERE 1=1
	"""
	params: list = []
	if class_name and class_name != "Tüm sınıflar":
		query += " AND c.name = ?"
		params.append(class_name)
	if person_type:
		norm_pt = clean_text(person_type).lower()
		if "sivil" in norm_pt or "diger" in norm_pt:
			query += " AND (c.name = 'Sivil / Diğer' OR s.person_type LIKE '%Sivil%' OR s.person_type LIKE '%diğer%')"
		else:
			query += " AND (c.name != 'Sivil / Diğer' AND s.person_type NOT LIKE '%Sivil%' AND s.person_type NOT LIKE '%diğer%')"
	query += " ORDER BY c.name, s.full_name, s.id"
	with connect_db() as connection:
		df = pd.read_sql_query(query, connection, params=tuple(params))
	if df.empty:
		return df

	today = date.today()
	out_rows = []
	for _, r in df.iterrows():
		sex = r["cinsiyet"]
		b_date = r["dogum_tarihi"]
		cur_age_mo = age_in_months(b_date, today) if b_date else None
		m_age_mo = int(r["olcum_yas_ay"]) if pd.notna(r["olcum_yas_ay"]) else None
		if m_age_mo is None and b_date and pd.notna(r["son_olcum_tarihi"]):
			try:
				m_age_mo = age_in_months(b_date, date.fromisoformat(str(r["son_olcum_tarihi"])[:10]))
			except Exception:
				m_age_mo = cur_age_mo

		h_val = round(float(r["boy_cm"]), 1) if pd.notna(r["boy_cm"]) else None
		w_val = round(float(r["kilo_kg"]), 1) if pd.notna(r["kilo_kg"]) else None
		hc_val = round(float(r["bas_cevresi_cm"]), 1) if (pd.notna(r["bas_cevresi_cm"]) and float(r["bas_cevresi_cm"]) > 0) else None
		b_val = round(float(r["vki"]), 1) if pd.notna(r["vki"]) else None
		if b_val is None and h_val and w_val and h_val > 0 and w_val > 0:
			b_val = round(w_val / ((h_val / 100.0) ** 2), 1)

		h_eval = evaluate_measurement(h_val, "height", sex, m_age_mo) if (sex and m_age_mo is not None and h_val) else None
		w_eval = evaluate_measurement(w_val, "weight", sex, m_age_mo) if (sex and m_age_mo is not None and w_val) else None
		b_eval = evaluate_measurement(b_val, "bmi", sex, m_age_mo) if (sex and m_age_mo is not None and b_val) else None
		hc_eval = evaluate_measurement(hc_val, "head_circ", sex, m_age_mo) if (sex and m_age_mo is not None and hc_val) else None

		if h_val is None and w_val is None and hc_val is None:
			growth_status = "—"
		elif sex and m_age_mo is not None:
			remarks = []
			if b_eval and "Normal" not in b_eval["status"]:
				remarks.append(b_eval["status"].replace(" aralığı", ""))
			if h_eval and "Normal" not in h_eval["status"]:
				remarks.append(h_eval["status"].replace(" aralığı", ""))
			if hc_eval and "Normal" not in hc_eval["status"]:
				remarks.append(hc_eval["status"].replace(" sınırı", ""))
			growth_status = ", ".join(remarks) if remarks else "🟢 Normal"
		else:
			growth_status = "⚪ Doğum/Cinsiyet Eksik"

		alert_parts = [clean_text(r["alerjiler"]), clean_text(r["kronik"]), clean_text(r["saglik_gecmisi"]), clean_text(r["ilaclar"])]
		alert_summary = " · ".join([p for p in alert_parts if p]) or "—"

		loc_parts = [clean_text(r.get("il")), clean_text(r.get("ilce")), clean_text(r.get("koy"))]
		loc_summary = " / ".join([p for p in loc_parts if p]) or "—"

		out_rows.append({
			"id": int(r["id"]),
			"Kayıt Türü": clean_text(r["kayit_turu"]) or "Öğrenci",
			"Sınıf": clean_text(r["sinif"]) or "—",
			"Görev / Statü": clean_text(r["notlar"]) or "—",
			"Okul No": clean_text(r["okul_no"]) or "—",
			"Ad Soyad": clean_text(r["ad_soyad"]),
			"Cinsiyet": clean_text(r["cinsiyet"]) or "—",
			"Doğum Tarihi": clean_text(r["dogum_tarihi"]) or "—",
			"Yaş (Ay)": cur_age_mo,
			"Boy (cm)": h_val,
			"Boy Z-Skor": h_eval["z"] if h_eval else None,
			"Boy %": h_eval["percentile"] if h_eval else None,
			"Kilo (kg)": w_val,
			"Kilo Z-Skor": w_eval["z"] if w_eval else None,
			"Kilo %": w_eval["percentile"] if w_eval else None,
			"Baş Çevresi (cm)": hc_val,
			"VKİ (BMI)": b_val,
			"VKİ Z-Skor": b_eval["z"] if b_eval else None,
			"VKİ %": b_eval["percentile"] if b_eval else None,
			"Gelişim Durumu": growth_status,
			"Kan Grubu": clean_text(r["kan_grubu"]) or "—",
			"Alerji / Kronik": alert_summary,
			"İl / İlçe / Köy": loc_summary,
			"Adres / Mahalle": clean_text(r["notlar"]) or "—",
			"Başvuru": int(r["basvuru_sayisi"] or 0),
			"Veli": clean_text(r["veli"]) or "—",
			"Telefon": clean_text(r["veli_telefon"]) or "—",
		})
	return pd.DataFrame(out_rows)


def get_students_column_config() -> dict:
	return {
		"id": None,
		"📂 Dosya": st.column_config.TextColumn(
			"📂 Dosya",
			help="Öğrencinin Kişi Dosyasını (Detay & İşlem) açmak için tıklayın",
			width="small",
		),
		"Yaş (Ay)": st.column_config.NumberColumn("Yaş (Ay)", format="%d ay"),
		"Boy (cm)": st.column_config.NumberColumn("Boy (cm)", format="%.1f cm"),
		"Boy Z-Skor": st.column_config.NumberColumn("Boy Z-Skor", format="%+.2f SDS"),
		"Boy %": st.column_config.NumberColumn("Boy %", format="%%%.1f"),
		"Kilo (kg)": st.column_config.NumberColumn("Kilo (kg)", format="%.1f kg"),
		"Kilo Z-Skor": st.column_config.NumberColumn("Kilo Z-Skor", format="%+.2f SDS"),
		"Kilo %": st.column_config.NumberColumn("Kilo %", format="%%%.1f"),
		"Baş Çevresi (cm)": st.column_config.NumberColumn("Baş Çevresi (cm)", format="%.1f cm"),
		"VKİ (BMI)": st.column_config.NumberColumn("VKİ (BMI)", format="%.1f"),
		"VKİ Z-Skor": st.column_config.NumberColumn("VKİ Z-Skor", format="%+.2f SDS"),
		"VKİ %": st.column_config.NumberColumn("VKİ %", format="%%%.1f"),
		"Başvuru": st.column_config.NumberColumn("Başvuru", format="%d"),
	}


def get_person_by_id(person_id: int | None) -> sqlite3.Row | None:
	if person_id is None:
		return None
	with connect_db() as connection:
		return connection.execute(
			"""SELECT s.id, s.person_type, s.full_name, c.name AS class_name, s.sex, s.birth_date,
					  s.student_no, s.guardian_name, s.guardian_phone, s.notes,
					  s.country, s.province, s.district, s.village
			   FROM students s JOIN classes c ON c.id=s.class_id WHERE s.id=?""",
			(int(person_id),),
		).fetchone()


def set_active_person(person_id: int | None) -> None:
	st.session_state["active_person_id"] = int(person_id) if person_id is not None else None


def select_shared_person(key_prefix: str, person_type: str | None = None, label: str = "Öğrenci / Kişi Seçin") -> sqlite3.Row | None:
	all_classes = get_classes()
	if person_type == "Öğrenci":
		classes = [c for c in all_classes if c != "Sivil / Diğer"]
	elif person_type == "Sivil":
		classes = ["Sivil / Diğer"] if "Sivil / Diğer" in all_classes else []
	else:
		classes = all_classes

	active_id = st.session_state.get("active_person_id")
	active_person = get_person_by_id(active_id)

	if person_type != "Sivil":
		class_opts = ["Tüm sınıflar", *classes]
		saved_class = st.session_state.get("active_class_filter", "Tüm sınıflar")
		cls_widget_key = f"{key_prefix}_cls_filter"

		# If an active person exists and their class doesn't match saved filter, reset to "Tüm sınıflar" so they are visible
		if active_person and saved_class != "Tüm sınıflar" and active_person["class_name"] != saved_class:
			if person_type is None or (person_type == "Öğrenci" and active_person["class_name"] != "Sivil / Diğer"):
				saved_class = "Tüm sınıflar"
				st.session_state["active_class_filter"] = "Tüm sınıflar"
				if cls_widget_key in st.session_state:
					st.session_state[cls_widget_key] = "Tüm sınıflar"

		cls_idx = class_opts.index(saved_class) if saved_class in class_opts else 0

		col_f, col_s = st.columns([1, 2.2])
		with col_f:
			chosen_class = st.selectbox("Sınıf Filtresi", class_opts, index=cls_idx, key=cls_widget_key)
			st.session_state["active_class_filter"] = chosen_class
	else:
		chosen_class = "Sivil / Diğer"
		col_s = st.container()

	all_people = student_options(None, person_type=person_type)
	if chosen_class == "Tüm sınıflar" or not chosen_class:
		people = all_people
	else:
		class_people = student_options(chosen_class, person_type=person_type)
		other_people = [p for p in all_people if p["id"] not in {cp["id"] for cp in class_people}]
		# Put class people first, then allow searching other people so no student is ever hidden
		people = class_people + other_people

	if not people:
		st.info("Kayıtlı kişi bulunamadı. 'Öğrenciler' veya 'Yeni Kayıt' bölümünden kayıt ekleyebilirsiniz.")
		return None

	def _person_label(row: sqlite3.Row) -> str:
		c_name = row["class_name"] or "Sınıf Belirtilmedi"
		prefix = f"[{c_name}] " if chosen_class == "Tüm sınıflar" or (chosen_class != c_name) else ""
		parts = [f"{prefix}{row['full_name']}"]
		if row["student_no"]:
			parts.append(f"No: {row['student_no']}")
		elif row["notes"] and (c_name == "Sivil / Diğer" or "Sivil" in str(row["person_type"] or "")):
			parts.append(str(row["notes"]))
		parts.append(f"#{row['id']}")
		return " · ".join(parts)

	lookup = {_person_label(r): r for r in people}
	labels = list(lookup.keys())
	id_to_idx = {r["id"]: i for i, r in enumerate(people)}
	default_idx = id_to_idx.get(active_id, None)

	# Keep widget state synchronized with global active_person_id when switching pages
	widget_key = f"{key_prefix}_person_select"
	last_synced_key = f"{key_prefix}_last_synced_id"
	if st.session_state.get(last_synced_key) != active_id:
		st.session_state[last_synced_key] = active_id
		st.session_state[widget_key] = labels[default_idx] if default_idx is not None else None

	with col_s:
		selected_label = st.selectbox(label, labels, index=default_idx, placeholder="İsim, sınıf veya okul numarası ile arayın...", key=widget_key)

	if not selected_label or selected_label not in lookup:
		return None

	chosen_person = lookup[selected_label]
	if chosen_person["id"] != active_id:
		set_active_person(chosen_person["id"])
		st.session_state[last_synced_key] = chosen_person["id"]
	return chosen_person


def record_linked_measurement(student_id: int, measured_on: date, height_cm: float | None, weight_kg: float | None,
							  birth_date: str | None = None, sex: str | None = None,
							  head_circ_cm: float | None = None) -> str:
	months = age_in_months(birth_date, measured_on) if birth_date else None
	h_val = float(height_cm) if (height_cm is not None and height_cm > 0) else None
	w_val = float(weight_kg) if (weight_kg is not None and weight_kg > 0) else None
	hc_val = float(head_circ_cm) if (head_circ_cm is not None and head_circ_cm > 0) else None
	bmi = (w_val / ((h_val / 100.0) ** 2)) if (h_val and w_val) else None
	with connect_db() as connection:
		existing = connection.execute(
			"SELECT id, height_cm, weight_kg, head_circ_cm, age_months, bmi FROM measurements WHERE student_id=? AND measured_on=? ORDER BY id DESC LIMIT 1",
			(student_id, measured_on.isoformat()),
		).fetchone()
		if existing:
			final_h = h_val if h_val is not None else existing["height_cm"]
			final_w = w_val if w_val is not None else existing["weight_kg"]
			final_hc = hc_val if hc_val is not None else existing["head_circ_cm"]
			final_bmi = (final_w / ((final_h / 100.0) ** 2)) if (final_h and final_w) else existing["bmi"]
			connection.execute(
				"""UPDATE measurements
				   SET height_cm=?, weight_kg=?, head_circ_cm=?, age_months=?, bmi=?
				   WHERE id=?""",
				(final_h, final_w, final_hc, months or existing["age_months"], final_bmi, existing["id"]),
			)
		else:
			connection.execute(
				"""INSERT INTO measurements(student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
				   VALUES (?, ?, ?, ?, ?, ?, ?)""",
				(student_id, measured_on.isoformat(), h_val, w_val, hc_val, months, float(bmi) if bmi else None),
			)
	save_person_data_file(student_id)
	parts = []
	if h_val is not None:
		h_str = f"Boy: {h_val:.1f} cm"
		if sex and months is not None:
			h_ev = evaluate_measurement(h_val, "height", sex, months)
			if h_ev:
				h_str += f" ({h_ev['z']:+.2f} SDS, %{h_ev['percentile']:.1f}p)"
		parts.append(h_str)
	if w_val is not None:
		w_str = f"Kilo: {w_val:.1f} kg"
		if sex and months is not None:
			w_ev = evaluate_measurement(w_val, "weight", sex, months)
			if w_ev:
				w_str += f" ({w_ev['z']:+.2f} SDS, %{w_ev['percentile']:.1f}p)"
		parts.append(w_str)
	if bmi is not None:
		b_str = f"VKİ: {bmi:.1f} kg/m²"
		if sex and months is not None:
			b_ev = evaluate_measurement(bmi, "bmi", sex, months)
			if b_ev:
				b_str += f" ({b_ev['z']:+.2f} SDS, %{b_ev['percentile']:.1f}p - {b_ev['status']})"
		parts.append(b_str)
	if hc_val is not None:
		hc_str = f"Baş Çevresi: {hc_val:.1f} cm"
		if sex and months is not None:
			hc_ev = evaluate_measurement(hc_val, "head_circ", sex, months)
			if hc_ev:
				hc_str += f" ({hc_ev['z']:+.2f} SDS, %{hc_ev['percentile']:.1f}p)"
		parts.append(hc_str)
	return " | ".join(parts)


def record_linked_vital(student_id: int, measured_at: str, systolic: float | None, diastolic: float | None,
						pulse: float | None, temperature: float | None, glucose: float | None,
						glucose_context: str | None = None, notes: str | None = None,
						spo2: float | None = None) -> str:
	with connect_db() as connection:
		connection.execute(
			"""INSERT INTO vital_readings(student_id, measured_at, systolic_mmhg, diastolic_mmhg,
			   pulse_bpm, temperature_c, spo2_pct, glucose_mg_dl, glucose_context, notes)
			   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(student_id, measured_at, systolic, diastolic, pulse, temperature, spo2, glucose, glucose_context, notes),
		)
	save_person_data_file(student_id)
	v_parts = []
	if temperature is not None:
		v_parts.append(f"Ateş: {temperature:.1f} °C")
	if spo2 is not None:
		v_parts.append(f"SpO2 (Saturasyon): %{spo2:.0f}")
	if systolic is not None and diastolic is not None:
		v_parts.append(f"Tansiyon: {systolic:.0f}/{diastolic:.0f} mmHg")
	if pulse is not None:
		v_parts.append(f"Nabız: {pulse:.0f} atım/dk")
	if glucose is not None:
		ctx = f" ({glucose_context})" if glucose_context else ""
		v_parts.append(f"Kan Şekeri: {glucose:.0f} mg/dL{ctx}")
	return " | ".join(v_parts)


def get_person_latest_measurements(person_id: int | None) -> dict:
	"""
	Tek ve Merkezi Antropometrik Veri Motoru (Single Source of Truth).
	Herhangi bir kişi (öğrenci/sivil) için en güncel boy, kilo, VKİ ve baş çevresi verilerini getirir.

	Önemli Prensipler:
	1. Tam İzolasyon: Veriler kesinlikle student_id=? ile sorgulanır. Farklı kişilerin verileri ASLA birbirine karışmaz.
	2. Boş Değer Önleme / Tarihsel Telafi: Son ölçümde baş çevresi (veya boy/kilo) girilmemiş olsa bile,
	   öğrencinin geçmişte kaydedilmiş geçerli bir baş çevresi ölçümü varsa bunu çeker ve tarihini belirtir.
	   Böylece "baş ölçümü sonucu boş değil" verileri sistemin her yerinde eksiksiz görüntülenir.
	3. Entegre WHO LMS Z-skor ve persantil hesaplaması: İlgili parametre için ölçüldüğü tarihteki yaşa göre doğru SDS/% üretir.
	"""
	empty_res = {
		"has_data": False,
		"person_id": person_id,
		"total_records": 0,
		"latest_measured_on": None,
		"height_cm": None, "height_date": None, "height_eval": None,
		"weight_kg": None, "weight_date": None, "weight_eval": None,
		"bmi": None, "bmi_date": None, "bmi_eval": None,
		"head_circ_cm": None, "head_circ_date": None, "head_circ_eval": None,
		"is_head_circ_from_earlier": False,
		"summary_items": [],
		"summary_html": "Henüz boy, kilo veya baş çevresi ölçümü girilmedi.",
	}
	if person_id is None:
		return empty_res

	with connect_db() as connection:
		person = connection.execute(
			"SELECT s.id, s.full_name, s.sex, s.birth_date, c.name AS class_name FROM students s JOIN classes c ON c.id=s.class_id WHERE s.id=?",
			(int(person_id),)
		).fetchone()
		if not person:
			return empty_res

		rows = connection.execute(
			"""SELECT id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi
			   FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC""",
			(int(person_id),)
		).fetchall()

	if not rows:
		return empty_res

	sex = person["sex"]
	b_date = person["birth_date"]
	latest_row = rows[-1]
	latest_measured_on = str(latest_row["measured_on"])[:10]

	h_val, h_date, h_mo = None, None, None
	w_val, w_date, w_mo = None, None, None
	hc_val, hc_date, hc_mo = None, None, None
	b_val, b_date_str, b_mo = None, None, None

	for r in reversed(rows):
		d_str = str(r["measured_on"])[:10]
		mo = r["age_months"]
		if mo is None and b_date:
			try:
				mo = age_in_months(b_date, date.fromisoformat(d_str))
			except Exception:
				mo = None

		if h_val is None and r["height_cm"] is not None and float(r["height_cm"]) > 0:
			h_val = float(r["height_cm"])
			h_date = d_str
			h_mo = mo

		if w_val is None and r["weight_kg"] is not None and float(r["weight_kg"]) > 0:
			w_val = float(r["weight_kg"])
			w_date = d_str
			w_mo = mo

		if hc_val is None and r["head_circ_cm"] is not None and float(r["head_circ_cm"]) > 0:
			hc_val = float(r["head_circ_cm"])
			hc_date = d_str
			hc_mo = mo

		if b_val is None and r["bmi"] is not None and float(r["bmi"]) > 0:
			b_val = float(r["bmi"])
			b_date_str = d_str
			b_mo = mo

	# If height and weight exist, recalculate latest BMI
	if h_val and w_val:
		b_val = round(w_val / ((h_val / 100.0) ** 2), 2)
		b_date_str = max(h_date or "", w_date or "")
		b_mo = max(h_mo or 0, w_mo or 0) or None

	h_ev = evaluate_measurement(h_val, "height", sex, h_mo) if (sex and h_mo is not None and h_val) else None
	w_ev = evaluate_measurement(w_val, "weight", sex, w_mo) if (sex and w_mo is not None and w_val) else None
	b_ev = evaluate_measurement(b_val, "bmi", sex, b_mo) if (sex and b_mo is not None and b_val) else None
	hc_ev = evaluate_measurement(hc_val, "head_circ", sex, hc_mo) if (sex and hc_mo is not None and hc_val) else None

	ideal_info = calculate_ideal_weight(h_val, sex, b_mo or h_mo, current_weight_kg=w_val)
	is_hc_earlier = (hc_date != latest_measured_on) if (hc_val and hc_date) else False

	m_items = []
	if h_val is not None:
		m_items.append(f"📏 <b>Boy:</b> {h_val:.1f} cm" + (f" ({h_ev['z']:+.2f} SDS · <b>%{h_ev['percentile']:.1f}p</b>)" if h_ev else ""))
	if w_val is not None:
		m_items.append(f"⚖️ <b>Kilo:</b> {w_val:.1f} kg" + (f" ({w_ev['z']:+.2f} SDS · <b>%{w_ev['percentile']:.1f}p</b>)" if w_ev else ""))
	if b_val is not None:
		m_items.append(f"🧬 <b>VKİ:</b> {b_val:.1f}" + (f" ({b_ev['z']:+.2f} SDS · <b>%{b_ev['percentile']:.1f}p</b> — {html.escape(b_ev['status'])})" if b_ev else ""))
	if ideal_info is not None:
		status_badge = f" · <span style='font-size:12.5px;font-weight:600;'>{ideal_info['status_text']}</span>" if ideal_info.get("status_text") else ""
		m_items.append(f"🎯 <b>İdeal Kilo Aralığı:</b> {ideal_info['min_kg']:.1f} – {ideal_info['max_kg']:.1f} kg <small style='opacity:0.8;'>(İdeal En Az: <b>{ideal_info['min_kg']:.1f} kg</b> · İdeal En Çok: <b>{ideal_info['max_kg']:.1f} kg</b>)</small>{status_badge}")
	if hc_val is not None:
		hc_extra = f" <small style='opacity:0.8;'>({hc_date})</small>" if is_hc_earlier else ""
		m_items.append(f"🧠 <b>Baş Çevresi:</b> {hc_val:.1f} cm{hc_extra}" + (f" ({hc_ev['z']:+.2f} SDS · <b>%{hc_ev['percentile']:.1f}p</b>)" if hc_ev else ""))

	return {
		"has_data": bool(m_items),
		"person_id": person_id,
		"total_records": len(rows),
		"latest_measured_on": latest_measured_on,
		"height_cm": h_val, "height_date": h_date, "height_eval": h_ev,
		"weight_kg": w_val, "weight_date": w_date, "weight_eval": w_ev,
		"bmi": b_val, "bmi_date": b_date_str, "bmi_eval": b_ev,
		"ideal_weight_kg": ideal_info["ideal_kg"] if ideal_info else None,
		"ideal_weight_min_kg": ideal_info["min_kg"] if ideal_info else None,
		"ideal_weight_max_kg": ideal_info["max_kg"] if ideal_info else None,
		"ideal_weight_range_str": f"{ideal_info['min_kg']:.1f} – {ideal_info['max_kg']:.1f} kg" if ideal_info else None,
		"ideal_weight_info": ideal_info,
		"head_circ_cm": hc_val, "head_circ_date": hc_date, "head_circ_eval": hc_ev,
		"is_head_circ_from_earlier": is_hc_earlier,
		"summary_items": m_items,
		"summary_html": (" &nbsp;|&nbsp; ".join(m_items) + f" <span style='font-size:12px;opacity:0.75;'>({latest_measured_on})</span>") if m_items else "Henüz boy, kilo veya baş çevresi ölçümü girilmedi.",
	}


def render_shared_person_summary(person: sqlite3.Row, section_key: str = "") -> None:
	person_id = person["id"]
	app_cfg = get_app_settings()
	theme_name = app_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
	theme = THEME_PRESETS.get(theme_name, THEME_PRESETS["Klinik Turkuaz (Varsayılan)"])
	is_dark = is_dark_mode_active(app_cfg)
	muted_col = "#94a3b8" if is_dark else "#475569"

	with connect_db() as connection:
		profile = connection.execute("SELECT * FROM health_profiles WHERE student_id=?", (person_id,)).fetchone()
		latest_vital = connection.execute(
			"SELECT * FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC LIMIT 1",
			(person_id,),
		).fetchone()
		visit_stat = connection.execute(
			"SELECT COUNT(*) AS cnt, MAX(visit_date) AS last_date FROM visits WHERE student_id=?",
			(person_id,),
		).fetchone()

	today_months = age_in_months(person["birth_date"], date.today()) if person["birth_date"] else None
	age_str = f"{today_months // 12} yaş {today_months % 12} ay ({today_months} ay)" if today_months is not None else "Doğum tarihi girilmedi"
	sex_str = {"K": "Kız / Kadın ♀", "E": "Erkek ♂"}.get(person["sex"], "Cinsiyet belirtilmedi")
	guardian_name = (profile["guardian_name"] if profile and profile["guardian_name"] else person["guardian_name"]) or "—"
	guardian_phone = (profile["guardian_phone"] if profile and profile["guardian_phone"] else person["guardian_phone"]) or "—"
	blood_group = (profile["blood_group"] if profile and profile["blood_group"] else None) or "Belirtilmedi"
	loc_parts = [person["province"] if "province" in person.keys() else None, person["district"] if "district" in person.keys() else None, person["village"] if "village" in person.keys() else None]
	loc_str = " / ".join([str(p) for p in loc_parts if p])
	addr_note = clean_text(person["notes"]) if "notes" in person.keys() else ""
	show_addr = app_cfg.get("show_address_in_summary", "1") == "1"

	# Anthropometric & Percentile Summary string from single source of truth
	meas_info = get_person_latest_measurements(person_id)
	if meas_info["has_data"]:
		meas_summary_html = meas_info["summary_html"]
	else:
		meas_summary_html = f"📏 <span style='color:{muted_col};'>Henüz boy, kilo veya baş çevresi ölçümü girilmedi.</span>"

	# Build Vital Summary string
	if latest_vital:
		v_items = []
		if latest_vital["temperature_c"] is not None:
			v_items.append(f"<b>Ateş:</b> {latest_vital['temperature_c']:.1f} °C")
		if "spo2_pct" in latest_vital.keys() and latest_vital["spo2_pct"] is not None:
			v_items.append(f"<b>SpO2:</b> %{latest_vital['spo2_pct']:.0f}")
		if latest_vital["systolic_mmhg"] is not None and latest_vital["diastolic_mmhg"] is not None:
			v_items.append(f"<b>Tansiyon:</b> {latest_vital['systolic_mmhg']:.0f}/{latest_vital['diastolic_mmhg']:.0f} mmHg")
		if latest_vital["pulse_bpm"] is not None:
			v_items.append(f"<b>Nabız:</b> {latest_vital['pulse_bpm']:.0f}/dk")
		if latest_vital["glucose_mg_dl"] is not None:
			v_items.append(f"<b>Şeker:</b> {latest_vital['glucose_mg_dl']:.0f} mg/dL")
		vital_summary_html = "💓 " + " &nbsp;|&nbsp; ".join(v_items) + f" <span style='color:{muted_col};font-size:12px;'>({latest_vital['measured_at']})</span>"
	else:
		vital_summary_html = f"💓 <span style='color:{muted_col};'>Henüz vital ölçüm kaydı yok.</span>"

	visit_cnt = visit_stat["cnt"] if visit_stat else 0
	last_v_date = visit_stat["last_date"] if (visit_stat and visit_stat["last_date"]) else "Yok"

	has_alert = profile and any(profile[k] for k in ["allergies", "chronic_conditions", "medical_history", "medications", "emergency_notes"])
	alert_banner_html = ""
	if has_alert:
		a_items = []
		if profile["allergies"]:
			a_items.append(f"🔴 <b>Alerji:</b> {html.escape(profile['allergies'])}")
		if profile["chronic_conditions"]:
			a_items.append(f"🟠 <b>Kronik Hastalıklar:</b> {html.escape(profile['chronic_conditions'])}")
		if profile["medical_history"]:
			a_items.append(f"📋 <b>Sağlık Geçmişi:</b> {html.escape(profile['medical_history'])}")
		if profile["medications"]:
			a_items.append(f"💊 <b>Düzenli İlaç:</b> {html.escape(profile['medications'])}")
		if profile["emergency_notes"]:
			a_items.append(f"⚠️ <b>Acil Not:</b> {html.escape(profile['emergency_notes'])}")
		alert_bg = "#450a0a" if is_dark else "#fef2f2"
		alert_col = "#fecaca" if is_dark else "#7f1d1d"
		alert_banner_html = (
			f"<div style='background:{alert_bg};border:1px solid #f87171;border-left:5px solid #dc2626;padding:8px 12px;border-radius:6px;margin-top:8px;font-size:13.5px;color:{alert_col};'>"
			+ " &nbsp;·&nbsp; ".join(a_items)
			+ "</div>"
		)

	with st.container(border=True):
		c_top1, c_top2 = st.columns([1.6, 1.4])
		sno_badge = f" · No: {person['student_no']}" if person['student_no'] else ""
		c_top1.markdown(f"**👤 {html.escape(person['full_name'])}** · {html.escape(person['class_name'])}{sno_badge} · {sex_str} · {age_str}")
		loc_txt = f" · 📍 {html.escape(loc_str)}" if loc_str else ""
		c_top2.caption(f"🩸 **Kan:** {html.escape(blood_group)} · 📞 **Veli:** {html.escape(guardian_name)} ({html.escape(guardian_phone)}){loc_txt} · 🩺 **Revir:** {visit_cnt} ({last_v_date})")

		st.markdown(meas_summary_html, unsafe_allow_html=True)
		st.markdown(vital_summary_html, unsafe_allow_html=True)

		if has_alert:
			a_items = []
			if profile["allergies"]:
				a_items.append(f"🔴 **Alerji:** {profile['allergies']}")
			if profile["chronic_conditions"]:
				a_items.append(f"🟠 **Kronik:** {profile['chronic_conditions']}")
			if profile["medical_history"]:
				a_items.append(f"📋 **Geçmiş:** {profile['medical_history']}")
			if profile["medications"]:
				a_items.append(f"💊 **İlaç:** {profile['medications']}")
			if profile["emergency_notes"]:
				a_items.append(f"⚠️ **Acil:** {profile['emergency_notes']}")
			st.warning("  ·  ".join(a_items))



BLOOD_GROUP_OPTIONS = ["Belirtilmedi", "A Rh+", "A Rh-", "B Rh+", "B Rh-", "AB Rh+", "AB Rh-", "0 Rh+", "0 Rh-"]


def normalize_blood_group(value: str | None) -> str:
	if not value:
		return "Belirtilmedi"
	v = clean_text(value).upper().replace("0", "0").replace("O", "0")
	if v in BLOOD_GROUP_OPTIONS:
		return v
	for opt in BLOOD_GROUP_OPTIONS[1:]:
		base = opt.replace(" ", "")
		if v.replace(" ", "") == base:
			return opt
	return "Belirtilmedi"


def render_person_data_file_editor(person_id: int, section_key: str = "overview") -> None:
	person_data = load_person_data_file(person_id)
	if not person_data:
		st.info("Bu kişiye ait veri dosyası bulunamadı.")
		return
	fpath = get_person_data_file_path(person_id, person_data.get("full_name"))
	st.caption(f"📂 **Kişiye Özel Veri Dosyası:** `{fpath.name}` (`kisi_verileri/{fpath.name}`) · Son Güncelleme: `{person_data.get('updated_at', '—')}`")

	d_c1, d_c2, d_c3 = st.columns([1.4, 1.4, 1.2])
	with d_c1:
		st.download_button(
			f"📥 {person_data.get('full_name')} Veri Dosyasını İndir (.json)",
			data=fpath.read_bytes() if fpath.exists() else json.dumps(person_data, ensure_ascii=False, indent=2).encode("utf-8"),
			file_name=fpath.name,
			mime="application/json",
			key=f"dl_pfile_{section_key}_{person_id}",
			use_container_width=True,
		)
	with d_c2:
		if st.button("🔄 Dosyayı Veritabanıyla Yeniden Senkronize Et", key=f"sync_pfile_{section_key}_{person_id}", use_container_width=True):
			save_person_data_file(person_id)
			st.success(f"{person_data.get('full_name')} veri dosyası güncellendi.")
			st.rerun()
	with d_c3:
		if st.button("🧹 Kişinin Ölçüm/Revir Kayıtlarını Sıfırla", key=f"clear_pfile_recs_{section_key}_{person_id}", use_container_width=True):
			with connect_db() as connection:
				connection.execute("DELETE FROM visits WHERE student_id=?", (int(person_id),))
				connection.execute("DELETE FROM measurements WHERE student_id=?", (int(person_id),))
				connection.execute("DELETE FROM vital_readings WHERE student_id=?", (int(person_id),))
			save_person_data_file(int(person_id))
			st.success(f"{person_data.get('full_name')} adlı kişinin revir, ölçüm ve vital kayıtları temizlendi.")
			st.rerun()

	ed_tab_v, ed_tab_m, ed_tab_vr, ed_tab_json = st.tabs([
		f"🩺 Revir Kayıtları ({len(person_data.get('visits', []))})",
		f"📏 Boy-Kilo Ölçümleri ({len(person_data.get('measurements', []))})",
		f"💓 Vital Bulgular ({len(person_data.get('vital_readings', []))})",
		"📝 Ham Veri Dosyası (JSON Düzenle)",
	])

	with ed_tab_v:
		v_list = person_data.get("visits", [])
		v_df = pd.DataFrame(v_list) if v_list else pd.DataFrame(columns=["visit_date", "visit_time", "complaint", "symptoms", "observation", "action", "guardian_contacted", "referral", "follow_up"])
		v_df = v_df.drop(columns=["id"], errors="ignore")
		edited_v = st.data_editor(
			v_df,
			num_rows="dynamic",
			width="stretch",
			hide_index=True,
			key=f"pfile_v_ed_{section_key}_{person_id}",
			column_config={
				"visit_date": st.column_config.TextColumn("Tarih (YYYY-AA-GG)"),
				"visit_time": st.column_config.TextColumn("Saat (SS:DD)"),
				"complaint": st.column_config.TextColumn("Başvuru Nedeni"),
				"symptoms": st.column_config.TextColumn("Belirtiler / Şikayet"),
				"observation": st.column_config.TextColumn("Gözlem ve Ölçüm"),
				"action": st.column_config.TextColumn("Yapılan İşlem"),
				"guardian_contacted": st.column_config.CheckboxColumn("Veli Arandı"),
				"referral": st.column_config.TextColumn("Yönlendirme"),
				"follow_up": st.column_config.TextColumn("İzlem"),
			},
		)
		if st.button("💾 Revir Kayıtlarını Kişi Dosyasına Kaydet", type="primary", key=f"save_pfile_v_{section_key}_{person_id}"):
			person_data["visits"] = edited_v.dropna(how="all").to_dict(orient="records")
			apply_person_data_file_to_db(person_id, person_data)
			st.success("Kişinin revir kayıtları kendi veri dosyasına kaydedildi.")
			st.rerun()

	with ed_tab_m:
		m_list = person_data.get("measurements", [])
		m_df = pd.DataFrame(m_list) if m_list else pd.DataFrame(columns=["measured_on", "height_cm", "weight_kg", "head_circ_cm"])
		for col in ("measured_on", "height_cm", "weight_kg", "head_circ_cm"):
			if col not in m_df.columns:
				m_df[col] = None
		m_df = m_df[["measured_on", "height_cm", "weight_kg", "head_circ_cm"]]
		edited_m = st.data_editor(
			m_df,
			num_rows="dynamic",
			width="stretch",
			hide_index=True,
			key=f"pfile_m_ed_{section_key}_{person_id}",
			column_config={
				"measured_on": st.column_config.TextColumn("Ölçüm Tarihi (YYYY-AA-GG)"),
				"height_cm": st.column_config.NumberColumn("Boy (cm)", format="%.1f"),
				"weight_kg": st.column_config.NumberColumn("Kilo (kg)", format="%.1f"),
				"head_circ_cm": st.column_config.NumberColumn("Baş Çevresi (cm)", format="%.1f"),
			},
		)
		if st.button("💾 Boy-Kilo Ölçümlerini Kişi Dosyasına Kaydet", type="primary", key=f"save_pfile_m_{section_key}_{person_id}"):
			person_data["measurements"] = edited_m.dropna(how="all").to_dict(orient="records")
			apply_person_data_file_to_db(person_id, person_data)
			st.success("Kişinin boy-kilo ölçümleri kendi veri dosyasına kaydedildi.")
			st.rerun()

	with ed_tab_vr:
		vr_list = person_data.get("vital_readings", [])
		vr_df = pd.DataFrame(vr_list) if vr_list else pd.DataFrame(columns=["measured_at", "temperature_c", "spo2_pct", "systolic_mmhg", "diastolic_mmhg", "pulse_bpm", "glucose_mg_dl", "notes"])
		for col in ("measured_at", "temperature_c", "spo2_pct", "systolic_mmhg", "diastolic_mmhg", "pulse_bpm", "glucose_mg_dl", "notes"):
			if col not in vr_df.columns:
				vr_df[col] = None
		vr_df = vr_df[["measured_at", "temperature_c", "spo2_pct", "systolic_mmhg", "diastolic_mmhg", "pulse_bpm", "glucose_mg_dl", "notes"]]
		edited_vr = st.data_editor(
			vr_df,
			num_rows="dynamic",
			width="stretch",
			hide_index=True,
			key=f"pfile_vr_ed_{section_key}_{person_id}",
			column_config={
				"measured_at": st.column_config.TextColumn("Tarih/Saat (YYYY-AA-GGTHH:MM)"),
				"temperature_c": st.column_config.NumberColumn("Ateş (°C)", format="%.1f"),
				"spo2_pct": st.column_config.NumberColumn("SpO2 (%)", format="%.0f"),
				"systolic_mmhg": st.column_config.NumberColumn("Büyük Tansiyon", format="%.0f"),
				"diastolic_mmhg": st.column_config.NumberColumn("Küçük Tansiyon", format="%.0f"),
				"pulse_bpm": st.column_config.NumberColumn("Nabız (/dk)", format="%.0f"),
				"glucose_mg_dl": st.column_config.NumberColumn("Kan Şekeri", format="%.0f"),
				"notes": st.column_config.TextColumn("Notlar"),
			},
		)
		if st.button("💾 Vital Kayıtları Kişi Dosyasına Kaydet", type="primary", key=f"save_pfile_vr_{section_key}_{person_id}"):
			person_data["vital_readings"] = edited_vr.dropna(how="all").to_dict(orient="records")
			apply_person_data_file_to_db(person_id, person_data)
			st.success("Kişinin vital kayıtları kendi veri dosyasına kaydedildi.")
			st.rerun()

	with ed_tab_json:
		raw_json_str = st.text_area(
			f"{fpath.name} İçeriği (JSON)",
			value=json.dumps(person_data, ensure_ascii=False, indent=2),
			height=280,
			key=f"pfile_raw_json_{section_key}_{person_id}",
		)
		if st.button("💾 JSON Dosyasını Kaydet ve Uygula", key=f"save_pfile_json_{section_key}_{person_id}"):
			try:
				parsed_json = json.loads(raw_json_str)
				apply_person_data_file_to_db(person_id, parsed_json)
				st.success("Kişi veri dosyası güncellendi ve uygulandı.")
				st.rerun()
			except Exception as e:
				st.error(f"JSON biçimi geçersiz: {e}")


# ---------------------------------------------------------------------------
# BÜTÜNLEŞİK KİŞİ DOSYASI (EHR 360° PATIENT CHART)
# ---------------------------------------------------------------------------

def _render_dossier_tab_overview(person: sqlite3.Row, person_id: int, hp: sqlite3.Row | None, visit_stats: sqlite3.Row | None, theme: dict, muted_col: str) -> None:
	meas_info = get_person_latest_measurements(person_id)
	with connect_db() as connection:
		latest_vital = connection.execute(
			"SELECT * FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC LIMIT 1",
			(person_id,),
		).fetchone()
		recent_visits = connection.execute(
			"SELECT visit_date, visit_time, complaint, symptoms, observation, action, referral, guardian_contacted FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC LIMIT 4",
			(person_id,),
		).fetchall()
		recent_labs = connection.execute(
			"SELECT r.report_date, x.test_name, x.value, x.unit, x.ref_low, x.ref_high, x.lab_flag FROM lab_results x JOIN lab_reports r ON r.id=x.report_id WHERE r.student_id=? ORDER BY r.report_date DESC, x.id DESC LIMIT 6",
			(person_id,),
		).fetchall()

	# Metric Cards with Ideal Weight & Conditional Head Circumference
	h_v = meas_info["height_cm"]
	h_ev = meas_info["height_eval"]
	w_v = meas_info["weight_kg"]
	w_ev = meas_info["weight_eval"]
	b_v = meas_info["bmi"]
	b_ev = meas_info["bmi_eval"]
	ideal_info = meas_info.get("ideal_weight_info")
	hc_v = meas_info["head_circ_cm"]
	hc_ev = meas_info["head_circ_eval"]
	today_mo = age_in_months(person["birth_date"], date.today()) if person["birth_date"] else None
	can_eval_hc_dossier = (today_mo is not None and 0 <= today_mo <= 60 and hc_v is not None)

	if can_eval_hc_dossier:
		mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
		with mc1:
			st.metric("📏 Boy", f"{h_v:.1f} cm" if h_v else "—", delta=f"{h_ev['z']:+.2f} SDS (%{h_ev['percentile']:.1f}p)" if h_ev else None)
			st.caption(f"Tarama: **{h_ev['status']}**" if h_ev else "Kayıt yok")
		with mc2:
			st.metric("⚖️ Kilo", f"{w_v:.1f} kg" if w_v else "—", delta=f"{w_ev['z']:+.2f} SDS (%{w_ev['percentile']:.1f}p)" if w_ev else None)
			st.caption(f"Tarama: **{w_ev['status']}**" if w_ev else "Kayıt yok")
		with mc3:
			st.metric("🧬 VKİ (BMI)", f"{b_v:.1f} kg/m²" if b_v else "—", delta=f"{b_ev['z']:+.2f} SDS (%{b_ev['percentile']:.1f}p)" if b_ev else None)
			st.caption(f"Tarama: **{b_ev['status']}**" if b_ev else "Kayıt yok")
		with mc4:
			rng_val = f"{ideal_info['min_kg']:.1f} – {ideal_info['max_kg']:.1f} kg" if ideal_info else "—"
			st.metric("🎯 İdeal Kilo Aralığı", rng_val, delta=ideal_info['status_text'] if ideal_info else None)
			if ideal_info:
				st.caption(f"İdeal En Az: **{ideal_info['min_kg']:.1f} kg**<br>İdeal En Çok: **{ideal_info['max_kg']:.1f} kg**", unsafe_allow_html=True)
			else:
				st.caption("Boy girilmedi")
		with mc5:
			st.metric("🧠 Baş Çevresi", f"{hc_v:.1f} cm" if hc_v else "—", delta=f"{hc_ev['z']:+.2f} SDS (%{hc_ev['percentile']:.1f}p)" if hc_ev else None)
			dt_s = f" ({meas_info['head_circ_date']})" if meas_info["is_head_circ_from_earlier"] else ""
			st.caption(f"Tarama: **{hc_ev['status'] if hc_ev else 'Kaydedildi'}**{dt_s}")
		with mc6:
			if latest_vital:
				v_items = []
				if latest_vital["temperature_c"] is not None: v_items.append(f"{latest_vital['temperature_c']:.1f}°C")
				if latest_vital["systolic_mmhg"] is not None and latest_vital["diastolic_mmhg"] is not None: v_items.append(f"{latest_vital['systolic_mmhg']:.0f}/{latest_vital['diastolic_mmhg']:.0f}")
				if latest_vital["pulse_bpm"] is not None: v_items.append(f"{latest_vital['pulse_bpm']:.0f}/dk")
				st.metric("💓 Son Vital", " · ".join(v_items) if v_items else "Kaydedildi")
				st.caption(f"Tarih: **{latest_vital['measured_at']}**")
			else:
				st.metric("💓 Son Vital", "—")
				st.caption("Henüz vital kaydı yok")
	else:
		mc1, mc2, mc3, mc4, mc5 = st.columns(5)
		with mc1:
			st.metric("📏 Boy", f"{h_v:.1f} cm" if h_v else "—", delta=f"{h_ev['z']:+.2f} SDS (%{h_ev['percentile']:.1f}p)" if h_ev else None)
			st.caption(f"Tarama: **{h_ev['status']}**" if h_ev else "Kayıt yok")
		with mc2:
			st.metric("⚖️ Kilo", f"{w_v:.1f} kg" if w_v else "—", delta=f"{w_ev['z']:+.2f} SDS (%{w_ev['percentile']:.1f}p)" if w_ev else None)
			st.caption(f"Tarama: **{w_ev['status']}**" if w_ev else "Kayıt yok")
		with mc3:
			st.metric("🧬 VKİ (BMI)", f"{b_v:.1f} kg/m²" if b_v else "—", delta=f"{b_ev['z']:+.2f} SDS (%{b_ev['percentile']:.1f}p)" if b_ev else None)
			st.caption(f"Tarama: **{b_ev['status']}**" if b_ev else "Kayıt yok")
		with mc4:
			rng_val = f"{ideal_info['min_kg']:.1f} – {ideal_info['max_kg']:.1f} kg" if ideal_info else "—"
			st.metric("🎯 İdeal Kilo Aralığı", rng_val, delta=ideal_info['status_text'] if ideal_info else None)
			if ideal_info:
				st.caption(f"İdeal En Az: **{ideal_info['min_kg']:.1f} kg**<br>İdeal En Çok: **{ideal_info['max_kg']:.1f} kg**", unsafe_allow_html=True)
			else:
				st.caption("Boy girilmedi")
		with mc5:
			if latest_vital:
				v_items = []
				if latest_vital["temperature_c"] is not None: v_items.append(f"{latest_vital['temperature_c']:.1f}°C")
				if latest_vital["systolic_mmhg"] is not None and latest_vital["diastolic_mmhg"] is not None: v_items.append(f"{latest_vital['systolic_mmhg']:.0f}/{latest_vital['diastolic_mmhg']:.0f}")
				if latest_vital["pulse_bpm"] is not None: v_items.append(f"{latest_vital['pulse_bpm']:.0f}/dk")
				st.metric("💓 Son Vital", " · ".join(v_items) if v_items else "Kaydedildi")
				st.caption(f"Tarih: **{latest_vital['measured_at']}**")
			else:
				st.metric("💓 Son Vital", "—")
				st.caption("Henüz vital kaydı yok")

	st.markdown("---")
	ov_col_l, ov_col_r = st.columns([1.5, 1])

	with ov_col_l:
		st.markdown("##### 🩺 Son Revir Ziyaretleri ve Müdahaleler")
		if recent_visits:
			for v in recent_visits:
				with st.container(border=True):
					v_col1, v_col2 = st.columns([1.2, 1.8])
					v_col1.markdown(f"**📅 {v['visit_date']} · {v['visit_time']}**")
					v_badges = [f"**{v['complaint']}**"]
					if v["referral"]:
						v_badges.append(f"<span style='background:#fee2e2;color:#991b1b;padding:1px 6px;border-radius:4px;font-size:11px;'>Sevk: {html.escape(v['referral'])}</span>")
					if v["guardian_contacted"]:
						v_badges.append("<span style='background:#fef3c7;color:#92400e;padding:1px 6px;border-radius:4px;font-size:11px;'>Veli Arandı</span>")
					v_col2.markdown(" ".join(v_badges), unsafe_allow_html=True)

					if v["symptoms"]:
						st.markdown(f"🔍 **Belirtiler:** {html.escape(v['symptoms'])}")
					if v["action"]:
						st.markdown(f"🛠️ **Yapılan İşlem:** {html.escape(v['action'])}")
					if v["observation"]:
						st.caption(f"📋 **Bulgular:** {html.escape(v['observation'])}")
		else:
			st.info("Kayıtlı revir ziyareti bulunmuyor. '🩺 Revir Kayıtları' sekmesinden yeni giriş yapabilirsiniz.")

		st.markdown("##### 🧪 Son Tahlil & Laboratuvar Bulguları")
		if recent_labs:
			lab_rows = [{"Tarih": lr["report_date"], "Test": lr["test_name"], "Sonuç": f"{lr['value']} {lr['unit'] or ''}", "Referans": f"{lr['ref_low'] or '—'} - {lr['ref_high'] or '—'}", "Uyarı": lr["lab_flag"] or "Normal"} for lr in recent_labs]
			st.dataframe(pd.DataFrame(lab_rows), width="stretch", hide_index=True)
		else:
			st.caption("Kayıtlı tahlil veya laboratuvar sonucu bulunmuyor.")

	with ov_col_r:
		st.markdown("##### 📑 Tıbbi Profil & Kritik Notlar")
		med_hist = (hp["medical_history"] if hp and hp["medical_history"] else "").strip()
		emer_notes = (hp["emergency_notes"] if hp and hp["emergency_notes"] else "").strip()
		with st.container(border=True):
			st.markdown(f"🩸 **Kan Grubu:** `{(hp['blood_group'] if hp and hp['blood_group'] else None) or 'Belirtilmedi'}`")
			st.markdown(f"🔴 **Alerjiler:** `{(hp['allergies'] if hp and hp['allergies'] else None) or 'Yok / Bildirilmedi'}`")
			st.markdown(f"🟠 **Kronik Durumlar:** `{(hp['chronic_conditions'] if hp and hp['chronic_conditions'] else None) or 'Yok / Bildirilmedi'}`")
			st.markdown(f"💊 **Sürekli İlaçlar:** `{(hp['medications'] if hp and hp['medications'] else None) or 'Yok / Bildirilmedi'}`")
			if med_hist:
				st.markdown(f"📜 **Geçmiş:** {html.escape(med_hist)}")
			if emer_notes:
				st.error(f"🚨 **Acil Notu:** {html.escape(emer_notes)}")

		st.markdown("##### 📁 Veri Dosyası & Senkronizasyon")
		with st.container(border=True):
			p_fpath = get_person_data_file_path(person_id, person["full_name"])
			st.caption(f"Kişisel JSON dosyası: `{p_fpath.name}`")
			dl_c1, dl_c2 = st.columns(2)
			with dl_c1:
				st.download_button(
					"📥 JSON İndir",
					data=p_fpath.read_bytes() if p_fpath.exists() else json.dumps(load_person_data_file(person_id), ensure_ascii=False, indent=2).encode("utf-8"),
					file_name=p_fpath.name,
					mime="application/json",
					key=f"dossier_dl_json_{person_id}",
					use_container_width=True,
				)
			with dl_c2:
				if st.button("🔄 Senkronize Et", key=f"dossier_sync_json_{person_id}", use_container_width=True):
					save_person_data_file(person_id)
					st.success("Senkronize edildi.")
					st.rerun()

	st.markdown("---")
	st.markdown("### 🩺 Okul Sağlık Taraması ve Bireysel Gelişim Bulguları")
	st.caption("Öğrencinin tarama bulgularını inceleyebilir, her kutucuğun sağ üst köşesindeki **✏️ Düzenle** butonuna tıklayarak doğrudan düzeltebilirsiniz.")

	top_scr_btn_c1, top_scr_btn_c2 = st.columns([2.5, 1.5])
	with top_scr_btn_c1:
		scr_stat_txt = f"🗓️ Son Okul Taraması: **{hp['last_screened_at']}**" if (hp and hp["last_screened_at"]) else "⏳ Bu öğrenciye henüz okul taraması yapılmadı."
		st.info(scr_stat_txt)
	with top_scr_btn_c2:
		if st.button("🩺 Tam Taramayı Aç / Düzenle ↗️", type="primary", use_container_width=True, key=f"dossier_btn_full_scr_{person_id}"):
			open_screening_dialog(person_id)

	scr_c1, scr_c2 = st.columns(2)
	active_box_key = f"active_edit_box_{person_id}"

	# Card 1: Göz ve Görme
	with scr_c1:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 👓 Göz ve Görme Durumu")
			is_editing_vis = (st.session_state.get(active_box_key) == "vision")
			btn_label = "✖ Kapat" if is_editing_vis else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_vis_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_vis else "vision"
				st.rerun()

			if is_editing_vis:
				with st.form(f"form_inline_vis_{person_id}"):
					gl_opts = ["Hayır (Kullanmıyor)", "Evet (Sürekli)", "Bazen / Sadece Okurken", "Gözlük Önerildi / Muayene Bekleniyor", "Göz Tembelliği Kapaması Yapılıyor"]
					cur_gl = (hp["wears_glasses"] if hp and hp["wears_glasses"] else "Hayır (Kullanmıyor)")
					i_gl = gl_opts.index(cur_gl) if cur_gl in gl_opts else 0
					new_gl = st.selectbox("Gözlük Kullanımı", gl_opts, index=i_gl)
					new_vr = st.text_input("Sağ Göz Durumu / Numara", value=(hp["vision_right"] if hp and hp["vision_right"] else "Tam / Normal (10/10)"))
					new_vl = st.text_input("Sol Göz Durumu / Numara", value=(hp["vision_left"] if hp and hp["vision_left"] else "Tam / Normal (10/10)"))
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET wears_glasses=?, vision_right=?, vision_left=?, updated_at=? WHERE student_id=?",
								(new_gl, new_vr.strip() or None, new_vl.strip() or None, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Göz ve görme bilgisi güncellendi.")
						st.rerun()
			else:
				gl_val = hp["wears_glasses"] if hp and hp["wears_glasses"] else "Kayıt Yok"
				vr_val = hp["vision_right"] if hp and hp["vision_right"] else "—"
				vl_val = hp["vision_left"] if hp and hp["vision_left"] else "—"
				st.markdown(f"👓 **Gözlük Takıyor mu:** `{gl_val}`")
				st.markdown(f"👁️ **Sağ Göz:** `{vr_val}` &nbsp;·&nbsp; 👁️ **Sol Göz:** `{vl_val}`")

	# Card 2: Sağlık Davranışı & Alışkanlıkları (Çoklu Seçim)
	with scr_c2:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 🍎 Sağlık Davranışı & Alışkanlıkları")
			is_editing_beh = (st.session_state.get(active_box_key) == "behavior")
			btn_label = "✖ Kapat" if is_editing_beh else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_beh_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_beh else "behavior"
				st.rerun()

			if is_editing_beh:
				with st.form(f"form_inline_beh_{person_id}"):
					cur_beh_list = parse_multiselect_values(hp["health_behavior"] if hp and hp["health_behavior"] else None, ["Beslenme ve Hijyen Çok İyi"])
					opts_beh = ensure_options_contain(HEALTH_BEHAVIOR_OPTIONS, cur_beh_list)
					sel_beh = st.multiselect("Sağlık Davranışı & Alışkanlıklar (Çoklu Seçilebilir)", opts_beh, default=cur_beh_list, key=f"il_beh_sel_{person_id}")
					cust_beh = st.text_input("Açıklama / Serbest Not (Opsiyonel)", value="", placeholder="Diyet, su tüketimi, hijyen notu...", key=f"il_beh_note_{person_id}")
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_beh = format_multiselect_save_str(sel_beh, cust_beh)
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET health_behavior=?, updated_at=? WHERE student_id=?",
								(res_beh, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Sağlık davranışı bilgisi güncellendi.")
						st.rerun()
			else:
				beh_val = hp["health_behavior"] if hp and hp["health_behavior"] else "Kayıt Yok / Değerlendirilmedi"
				beh_items = parse_multiselect_values(beh_val)
				if beh_items and beh_items != ["Kayıt Yok / Değerlendirilmedi"]:
					chips_html = "".join([f"<span style='background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;padding:2px 7px;border-radius:4px;font-size:12px;margin-right:5px;display:inline-block;margin-top:3px;'>🍎 {html.escape(c)}</span>" for c in beh_items])
					st.markdown(f"🍎 **Beslenme, Diş ve Hijyen Alışkanlıkları:**<br>{chips_html}", unsafe_allow_html=True)
				else:
					st.markdown(f"🍎 **Beslenme, Diş ve Hijyen:** `{beh_val}`")

	# Card 3: Saç Taraması & Hijyen (Çoklu Seçim)
	with scr_c1:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 💇 Saç Taraması & Hijyen (Pediküloz)")
			is_editing_hair = (st.session_state.get(active_box_key) == "hair")
			btn_label = "✖ Kapat" if is_editing_hair else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_hair_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_hair else "hair"
				st.rerun()

			if is_editing_hair:
				with st.form(f"form_inline_hair_{person_id}"):
					cur_hair_list = parse_multiselect_values(hp["hair_screening"] if hp and hp["hair_screening"] else None, ["Temiz / Parazit Yok"])
					opts_hair = ensure_options_contain(HAIR_SCREENING_OPTIONS, cur_hair_list)
					sel_hair = st.multiselect("Saç Taraması & Hijyen Durumu (Çoklu Seçilebilir)", opts_hair, default=cur_hair_list, key=f"il_hair_sel_{person_id}")
					cust_hair = st.text_input("Açıklama / Serbest Not (Opsiyonel)", value="", placeholder="İlaç, sirke durumu, kontrol notu...", key=f"il_hair_note_{person_id}")
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_hair = format_multiselect_save_str(sel_hair, cust_hair)
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET hair_screening=?, updated_at=? WHERE student_id=?",
								(res_hair, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Saç taraması bilgisi güncellendi.")
						st.rerun()
			else:
				hair_val = hp["hair_screening"] if hp and hp["hair_screening"] else "Tarama Yapılmadı"
				hair_items = parse_multiselect_values(hair_val)
				if any("Bit" in x or "Sirke" in x or "Tedavi" in x or "İlaç" in x for x in hair_items):
					chips_html = "".join([f"<span style='background:#fef3c7;color:#92400e;border:1px solid #fde68a;padding:2px 7px;border-radius:4px;font-size:12px;margin-right:5px;display:inline-block;margin-top:3px;'>💇 {html.escape(c)}</span>" for c in hair_items])
					st.markdown(f"💇 **Saç / Pediküloz Bulguları:**<br>{chips_html}", unsafe_allow_html=True)
				elif hair_items and hair_items != ["Tarama Yapılmadı"]:
					chips_html = "".join([f"<span style='background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;padding:2px 7px;border-radius:4px;font-size:12px;margin-right:5px;display:inline-block;margin-top:3px;'>💇 {html.escape(c)}</span>" for c in hair_items])
					st.markdown(f"💇 **Saç / Pediküloz Durumu:**<br>{chips_html}", unsafe_allow_html=True)
				else:
					st.markdown(f"💇 **Saç / Pediküloz Durumu:** `{hair_val}`")

	# Card 4: Genel Psikososyal Durumu
	with scr_c2:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 🧠 Genel Psikososyal Durumu")
			is_editing_psy = (st.session_state.get(active_box_key) == "psychosocial")
			btn_label = "✖ Kapat" if is_editing_psy else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_psy_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_psy else "psychosocial"
				st.rerun()

			if is_editing_psy:
				with st.form(f"form_inline_psy_{person_id}"):
					psy_opts = [
						"Sınıf İçi Uyum ve İletişimi Çok İyi", "Çekingen / İçe Kapanık",
						"Aşırı Hareketli / Dikkat Dağınıklığı", "Okul Kaygısı / Anneye Bağlılık",
						"Öfke Kontrolü / Akran Çatışması", "Göz Teması ve İletişim Desteklenmeli", "Diğer / Serbest Not"
					]
					cur_psy = (hp["psychosocial_status"] if hp and hp["psychosocial_status"] else "Sınıf İçi Uyum ve İletişimi Çok İyi")
					i_psy = psy_opts.index(cur_psy) if cur_psy in psy_opts else 0
					sel_psy = st.selectbox("Psikososyal Uyum", psy_opts, index=i_psy)
					cust_psy = st.text_input("Açıklama / Özel Not", value="" if sel_psy != "Diğer / Serbest Not" else cur_psy)
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_psy = cust_psy.strip() if sel_psy == "Diğer / Serbest Not" and cust_psy.strip() else sel_psy
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET psychosocial_status=?, updated_at=? WHERE student_id=?",
								(res_psy, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Psikososyal durum bilgisi güncellendi.")
						st.rerun()
			else:
				psy_val = hp["psychosocial_status"] if hp and hp["psychosocial_status"] else "Kayıt Yok / Değerlendirilmedi"
				st.markdown(f"🧠 **Sınıf İçi Uyum ve İletişim:**<br>`{psy_val}`", unsafe_allow_html=True)

	# Card 5: İşitme Durumu (Kulak - Çoklu Seçim)
	with scr_c1:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 👂 İşitme Durumu (Kulak)")
			is_editing_hear = (st.session_state.get(active_box_key) == "hearing")
			btn_label = "✖ Kapat" if is_editing_hear else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_hear_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_hear else "hearing"
				st.rerun()

			if is_editing_hear:
				with st.form(f"form_inline_hear_{person_id}"):
					cur_hr_list = parse_multiselect_values(hp["hearing_right"] if hp and hp["hearing_right"] else None, ["Normal / İyi Duyuyor"])
					cur_hl_list = parse_multiselect_values(hp["hearing_left"] if hp and hp["hearing_left"] else None, ["Normal / İyi Duyuyor"])
					opts_hr = ensure_options_contain(HEARING_SCREENING_OPTIONS, cur_hr_list)
					opts_hl = ensure_options_contain(HEARING_SCREENING_OPTIONS, cur_hl_list)
					sel_hr = st.multiselect("👂 Sağ Kulak İşitme (Çoklu)", opts_hr, default=cur_hr_list, key=f"il_hr_sel_{person_id}")
					sel_hl = st.multiselect("👂 Sol Kulak İşitme (Çoklu)", opts_hl, default=cur_hl_list, key=f"il_hl_sel_{person_id}")
					cust_hear = st.text_input("Açıklama / KBB Sevk Notu (Opsiyonel)", value="", placeholder="Odyometri sonucu, hekim muayene notu...", key=f"il_hear_note_{person_id}")
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_hr = format_multiselect_save_str(sel_hr, cust_hear if not sel_hl else None)
						res_hl = format_multiselect_save_str(sel_hl)
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET hearing_right=?, hearing_left=?, updated_at=? WHERE student_id=?",
								(res_hr, res_hl, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("İşitme bilgisi güncellendi.")
						st.rerun()
			else:
				hr_val = hp["hearing_right"] if hp and hp["hearing_right"] else "—"
				hl_val = hp["hearing_left"] if hp and hp["hearing_left"] else "—"
				st.markdown(f"👂 **Sağ Kulak:** `{hr_val}`<br>👂 **Sol Kulak:** `{hl_val}`", unsafe_allow_html=True)

	# Card 6: Zararlı Alışkanlıklar & Risk Faktörleri (Çoklu Seçim)
	with scr_c2:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### ⚠️ Zararlı Alışkanlıklar & Riskler")
			is_editing_harm = (st.session_state.get(active_box_key) == "harmful_habits")
			btn_label = "✖ Kapat" if is_editing_harm else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_harm_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_harm else "harmful_habits"
				st.rerun()

			if is_editing_harm:
				with st.form(f"form_inline_harm_{person_id}"):
					cur_harm_list = parse_multiselect_values(hp["harmful_habits"] if hp and "harmful_habits" in hp.keys() and hp["harmful_habits"] else None, ["Zararlı Alışkanlık Tespit Edilmedi"])
					opts_harm = ensure_options_contain(HARMFUL_HABITS_OPTIONS, cur_harm_list)
					sel_harm = st.multiselect("⚠️ Zararlı Alışkanlıklar & Risk Faktörleri (Çoklu)", opts_harm, default=cur_harm_list, key=f"il_harm_sel_{person_id}")
					cust_harm = st.text_input("Açıklama / Serbest Not (Opsiyonel)", value="", placeholder="Dijital ekran süresi, tütün, kafein vb...", key=f"il_harm_note_{person_id}")
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_harm = format_multiselect_save_str(sel_harm, cust_harm)
						with connect_db() as connection:
							connection.execute(
								"UPDATE health_profiles SET harmful_habits=?, updated_at=? WHERE student_id=?",
								(res_harm, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Zararlı alışkanlıklar bilgisi güncellendi.")
						st.rerun()
			else:
				harm_val = hp["harmful_habits"] if (hp and "harmful_habits" in hp.keys() and hp["harmful_habits"]) else "Zararlı Alışkanlık Tespit Edilmedi"
				harm_items = parse_multiselect_values(harm_val)
				if any(x not in ("Zararlı Alışkanlık Tespit Edilmedi", "Kayıt Yok", "Yok") for x in harm_items):
					chips_html = "".join([f"<span style='background:#fee2e2;color:#991b1b;border:1px solid #fecaca;padding:2px 7px;border-radius:4px;font-size:12px;margin-right:5px;display:inline-block;margin-top:3px;'>⚠️ {html.escape(c)}</span>" for c in harm_items])
					st.markdown(f"⚠️ **Tespit Edilen Riskler:**<br>{chips_html}", unsafe_allow_html=True)
				else:
					st.markdown(f"⚠️ **Zararlı Alışkanlık Durumu:** `{harm_val}`")

	# Card 6: Klinik Tıbbi Profil
	with scr_c2:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 📋 Klinik Sağlık Özeti")
			is_editing_clin = (st.session_state.get(active_box_key) == "clinical")
			btn_label = "✖ Kapat" if is_editing_clin else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_clin_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_clin else "clinical"
				st.rerun()

			if is_editing_clin:
				with st.form(f"form_inline_clin_{person_id}"):
					cur_bg = normalize_blood_group(hp["blood_group"] if hp else None)
					bg_idx = BLOOD_GROUP_OPTIONS.index(cur_bg) if cur_bg in BLOOD_GROUP_OPTIONS else 0
					new_bg = st.selectbox("Kan Grubu", BLOOD_GROUP_OPTIONS, index=bg_idx)
					new_al = st.text_input("Alerjiler", value=(hp["allergies"] if hp and hp["allergies"] else ""))
					new_ch = st.text_input("Kronik Durumlar", value=(hp["chronic_conditions"] if hp and hp["chronic_conditions"] else ""))
					new_med = st.text_input("Sürekli İlaçlar", value=(hp["medications"] if hp and hp["medications"] else ""))
					new_em = st.text_input("Acil Durum Notu", value=(hp["emergency_notes"] if hp and hp["emergency_notes"] else ""))
					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						res_bg = None if new_bg == "Belirtilmedi" else new_bg
						with connect_db() as connection:
							connection.execute(
								"""UPDATE health_profiles SET blood_group=?, allergies=?, chronic_conditions=?,
								   medications=?, emergency_notes=?, updated_at=? WHERE student_id=?""",
								(res_bg, new_al.strip() or None, new_ch.strip() or None, new_med.strip() or None,
								 new_em.strip() or None, datetime.now().isoformat(timespec="minutes"), person_id)
							)
						save_person_data_file(person_id)
						st.session_state[active_box_key] = None
						st.success("Klinik profil güncellendi.")
						st.rerun()
			else:
				st.markdown(f"🩸 **Kan Grubu:** `{(hp['blood_group'] if hp and hp['blood_group'] else None) or 'Belirtilmedi'}`")
				st.markdown(f"🔴 **Alerjiler:** `{(hp['allergies'] if hp and hp['allergies'] else None) or 'Yok'}`")
				st.markdown(f"🟠 **Kronik Durumlar:** `{(hp['chronic_conditions'] if hp and hp['chronic_conditions'] else None) or 'Yok'}`")

	# Card 7: Aşı Takvimi ve Bağışıklama Özeti
	with scr_c1:
		with st.container(border=True):
			hdr1, hdr2 = st.columns([3, 1.2])
			hdr1.markdown("##### 💉 Aşı Takvimi & Bağışıklama")
			is_editing_vac = (st.session_state.get(active_box_key) == "vaccination")
			btn_label = "✖ Kapat" if is_editing_vac else "✏️ Düzenle"
			if hdr2.button(btn_label, key=f"btn_edit_vac_{person_id}", use_container_width=True):
				st.session_state[active_box_key] = None if is_editing_vac else "vaccination"
				st.rerun()

			vac_sum = get_student_vaccination_summary(person_id, person["birth_date"])
			if is_editing_vac:
				with st.form(f"form_inline_vac_{person_id}"):
					st.caption("Temel okul çağı ve rutin aşıları hızlıca işaretleyin:")
					key_vacs = [
						("kkk_1", "KKK 1. Doz (12. Ay)"),
						("kkk_2", "KKK 2. Doz (İlkokul 1)"),
						("dabt_ipa_hib_r", "5'li Karma Pekiştirme (18. Ay)"),
						("dabt_ipa_r", "4'lü Karma Pekiştirme (İlkokul 1)"),
						("bcg", "BCG Verem Aşısı"),
						("hep_b_3", "Hepatit B Serisi (3 Doz)"),
						("su_cicegi", "Suçiçeği Aşısı"),
						("hep_a_2", "Hepatit A Serisi (2 Doz)"),
						("td_r", "Td Erişkin Difteri-Tetanoz (8. Sınıf)"),
					]
					vac_updates = {}
					for vcode, vtitle in key_vacs:
						cur_d = bool(vac_sum["records"].get(vcode, {}).get("is_done"))
						vac_updates[vcode] = st.checkbox(vtitle, value=cur_d, key=f"il_vac_{person_id}_{vcode}")

					if st.form_submit_button("💾 Kaydet", type="primary", use_container_width=True):
						for vcode, is_checked in vac_updates.items():
							save_student_vaccination(person_id, vcode, is_checked)
						st.session_state[active_box_key] = None
						st.success("Aşı bilgileri güncellendi.")
						st.rerun()
			else:
				if vac_sum["applicable_done_count"] > 0 and len(vac_sum["overdue_list"]) == 0:
					st.markdown(f"🟢 **Durum:** `Takvime Tam Uygun (%{vac_sum['percent']})`")
				elif len(vac_sum["overdue_list"]) > 0:
					st.markdown(f"⚠️ **Durum:** `%{vac_sum['percent']} ({len(vac_sum['overdue_list'])} Doz Eksik)`")
					st.caption(f"Eksik: **{', '.join(vac_sum['overdue_list'][:3])}**" + ("..." if len(vac_sum['overdue_list']) > 3 else ""))
				else:
					st.markdown("⚪ **Durum:** `Aşı kaydı girilmedi`")

				st.markdown(f"💉 **Yapılan:** `{vac_sum['applicable_done_count']} / {vac_sum['applicable_total']}` &nbsp;·&nbsp; **Toplam:** `{vac_sum['completed_count']} Doz`")

	# Card 8: Sağlık Bakanlığı Takvim Yönlendirmesi
	with scr_c2:
		with st.container(border=True):
			st.markdown("##### 🏛️ T.C. Sağlık Bakanlığı Takvimi")
			st.caption("Öğrencinin tüm bebeklik ve okul çağı aşılarını resmi Sağlık Bakanlığı matris tablosu üstünde incelemek ve işaretlemek için:")
			st.info("💡 Yukarıdaki sekmelerden **'💉 Aşı Takvimi'** sekmesine geçerek tüm aşıları renkli matriste görebilir, tek tıkla tamamlayabilir veya Aşı Kartı çıktısı alabilirsiniz.")

def _render_dossier_tab_visits(person: sqlite3.Row, person_id: int) -> None:
	st.markdown("#### ➕ Bu Kişiye Yeni Revir Ziyareti Kaydet")
	with st.expander("📝 Yeni Revir Kayıt Formunu Aç", expanded=True):
		with st.form(f"dossier_form_new_visit_{person_id}"):
			vc1, vc2, vc3 = st.columns([1.2, 1, 1.8])
			v_date = vc1.date_input("Ziyaret Tarihi", value=date.today(), key=f"dossier_v_date_{person_id}")
			v_time = vc2.time_input("Ziyaret Saati", value=datetime.now().time(), key=f"dossier_v_time_{person_id}")
			v_complaint = vc3.selectbox(
				"Başvuru Nedeni",
				["Hastalık / rahatsızlık", "Düşme / çarpma", "Yaralanma / kesik", "Karın ağrısı", "Baş ağrısı", "Ateş", "Alerjik reaksiyon", "İlaç uygulaması", "Rutin kontrol", "Diğer"],
				key=f"dossier_v_comp_{person_id}"
			)

			v_symptoms = st.text_area("Belirtiler / Şikayet", placeholder="Örn: 2 saattir devam eden baş ağrısı, hafif halsizlik...", height=70, key=f"dossier_v_symp_{person_id}")
			v_obs = st.text_area("Gözlem ve Muayene Bulguları", placeholder="Örn: Boğaz hiperemik, ciltte döküntü yok...", height=70, key=f"dossier_v_obs_{person_id}")
			v_action = st.text_area("Yapılan Müdahale / Tedavi", placeholder="Örn: İstirahat verildi, ılık su içirildi, veliye bilgi verildi...", height=70, key=f"dossier_v_act_{person_id}")

			v_chk1, v_chk2, v_chk3 = st.columns([1.2, 1.4, 1.4])
			v_guardian_cont = v_chk1.checkbox("Veli Bilgilendirildi", value=False, key=f"dossier_v_gcont_{person_id}")
			v_referral = v_chk2.text_input("Sevk Durumu", placeholder="Örn: Aile Hekimi / Devlet Hastanesi", key=f"dossier_v_ref_{person_id}")
			v_follow_up = v_chk3.text_input("İzlem Planı", placeholder="Örn: 1 saat sonra ateş kontrolü", key=f"dossier_v_fup_{person_id}")

			st.markdown("###### 💓 Bu vizitte alınan ölçümler (İsteğe bağlı - otomatik bağlantılanır):")
			m_c1, m_c2, m_c3, m_c4, m_c5 = st.columns(5)
			v_temp = m_c1.number_input("Ateş (°C)", min_value=30.0, max_value=45.0, value=None, step=0.1, format="%.1f", key=f"dossier_v_temp_{person_id}")
			v_sys = m_c2.number_input("Sistolik", min_value=40.0, max_value=250.0, value=None, step=1.0, key=f"dossier_v_sys_{person_id}")
			v_dia = m_c3.number_input("Diyastolik", min_value=20.0, max_value=150.0, value=None, step=1.0, key=f"dossier_v_dia_{person_id}")
			v_pulse = m_c4.number_input("Nabız (/dk)", min_value=30.0, max_value=220.0, value=None, step=1.0, key=f"dossier_v_pls_{person_id}")
			v_spo2 = m_c5.number_input("SpO2 (%)", min_value=50.0, max_value=100.0, value=None, step=1.0, key=f"dossier_v_spo2_{person_id}")

			hw_c1, hw_c2, hw_c3 = st.columns(3)
			v_height = hw_c1.number_input("Boy (cm)", min_value=40.0, max_value=230.0, value=None, step=0.1, format="%.1f", key=f"dossier_v_h_{person_id}")
			v_weight = hw_c2.number_input("Kilo (kg)", min_value=2.0, max_value=200.0, value=None, step=0.1, format="%.1f", key=f"dossier_v_w_{person_id}")
			v_head_circ = hw_c3.number_input("Baş Çevresi (cm)", min_value=20.0, max_value=100.0, value=None, step=0.1, format="%.1f", key=f"dossier_v_hc_{person_id}")

			submit_v = st.form_submit_button("💾 Revir Ziyaretini Bu Kişiye Kaydet", type="primary", use_container_width=True)

		if submit_v:
			auto_notes = []
			if any(x is not None for x in [v_temp, v_sys, v_pulse, v_spo2]):
				v_summary = record_linked_vital(
					person_id,
					f"{v_date.isoformat()} {v_time.strftime('%H:%M')}",
					v_sys, v_dia, v_pulse, v_temp, None, notes="Revir viziti esnasında alındı", spo2=v_spo2,
				)
				if v_summary:
					auto_notes.append(f"[Vital: {v_summary}]")
			has_hw = (v_height is not None and v_weight is not None and v_height > 0 and v_weight > 0) or (v_head_circ is not None and v_head_circ > 0)
			if has_hw:
				hw_summary = record_linked_measurement(
					person_id, v_date, v_height, v_weight,
					person["birth_date"], person["sex"], head_circ_cm=v_head_circ
				)
				if hw_summary:
					auto_notes.append(f"[Antropometri: {hw_summary}]")

			final_observation = " · ".join(filter(None, [v_obs.strip(), *auto_notes]))
			with connect_db() as connection:
				connection.execute(
					"""INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms, observation,
					   action, guardian_contacted, referral, follow_up) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
					(person_id, v_date.isoformat(), v_time.strftime("%H:%M"), v_complaint, v_symptoms.strip(),
					 final_observation, v_action.strip(), int(v_guardian_cont), v_referral.strip() or None, v_follow_up.strip() or None),
				)
			save_person_data_file(person_id)
			st.success(f"{person['full_name']} için yeni revir kaydı başarıyla eklendi.")
			st.rerun()

	st.markdown("#### 📋 Geçmiş Revir Kayıtları")
	with connect_db() as connection:
		visits_df = pd.read_sql_query(
			"""SELECT id, visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru,
					  symptoms AS Belirtiler, observation AS Gözlem_ve_Bulgular, action AS Yapılan_İşlem,
					  referral AS Sevk, guardian_contacted AS Veli_Arandı, follow_up AS İzlem
			   FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC""",
			connection, params=(person_id,)
		)
	if visits_df.empty:
		st.info("Bu kişi için henüz kayıtlı revir ziyareti bulunmuyor.")
	else:
		st.dataframe(visits_df.drop(columns=["id", "follow_up"], errors="ignore"), width="stretch", hide_index=True)
		del_c1, del_c2 = st.columns([3, 1])
		with del_c2:
			visit_opts = {f"{r['Tarih']} {r['Saat']} - {r['Başvuru']} (#{r['id']})": int(r["id"]) for _, r in visits_df.iterrows()}
			sel_del_v = st.selectbox("Silinecek vizit", list(visit_opts.keys()), key=f"dossier_del_v_sel_{person_id}")
			if st.button("🗑️ Seçili Viziti Sil", type="secondary", key=f"dossier_btn_del_v_{person_id}", use_container_width=True):
				delete_visit(visit_opts[sel_del_v])
				st.success("Revir kaydı silindi.")
				st.rerun()


def render_clinical_growth_curve(enriched_records: pd.DataFrame, person: dict | sqlite3.Row,
								indicator_name: str, y_col: str, title: str, unit: str, default_max_mo: int = 72) -> None:
	has_sex = person["sex"] in {"K", "E"}
	sex_label = "Kız Çocuk" if person["sex"] == "K" else ("Erkek Çocuk" if person["sex"] == "E" else "Cinsiyet Yok")
	if not has_sex:
		st.warning("WHO klinik büyüme eğrilerini çizmek için öğrenci profilinde cinsiyet bilgisi gereklidir.")
		return

	stu_months = enriched_records["Yas_ay"].dropna().astype(int).tolist() if "Yas_ay" in enriched_records.columns else []
	max_stu_mo = max(stu_months) if stu_months else 60
	graph_max_mo = min(default_max_mo, max(72, max_stu_mo + 6))

	curves_df = generate_who_percentile_curves(indicator_name, person["sex"], min_month=0, max_month=graph_max_mo)
	if curves_df.empty:
		st.info(f"'{title}' için WHO referans verisi bulunamadı.")
		return

	color_scale = alt.Scale(
		domain=["3p (Kritik Alt)", "15p (Alt Normal)", "50p (Medyan / Ortalama)", "85p (Üst Normal)", "97p (Kritik Üst)"],
		range=["#ef4444", "#f59e0b", "#10b981", "#f59e0b", "#ef4444"],
	)

	curve_chart = alt.Chart(curves_df).mark_line(interpolate="monotone", strokeWidth=1.8).encode(
		x=alt.X("Yas_Ay:Q", title="Yaş (Ay)", scale=alt.Scale(zero=False, domain=[0, graph_max_mo])),
		y=alt.Y("Deger:Q", title=f"{title} ({unit})", scale=alt.Scale(zero=False)),
		color=alt.Color("Egri_Etiketi:N", title="WHO Persantil Kanalları", scale=color_scale),
		tooltip=[
			alt.Tooltip("Yas_Ay:Q", title="Yaş (Ay)"),
			alt.Tooltip("Egri_Etiketi:N", title="Persantil Eğrisi"),
			alt.Tooltip("Deger:Q", title=f"Referans Değer ({unit})", format=".1f"),
		],
	)

	col_map = {
		"Boy_cm": ("Boy_Persantil", "Boy_Z"),
		"Kilo_kg": ("Kilo_Persantil", "Kilo_Z"),
		"BMI": ("BMI_Persantil", "BMI_Z"),
		"Bas_Cevresi_cm": ("Bas_Cevresi_Persantil", "Bas_Cevresi_Z"),
	}
	pct_col, z_col = col_map.get(y_col, (None, None))
	needed_cols = ["Tarih", "Yas_ay", y_col]
	if pct_col and pct_col in enriched_records.columns:
		needed_cols.append(pct_col)
	if z_col and z_col in enriched_records.columns:
		needed_cols.append(z_col)

	stu_df = enriched_records[needed_cols].dropna(subset=[y_col, "Yas_ay"]).copy()
	stu_df["Yas_Ay"] = stu_df["Yas_ay"].astype(int)
	stu_df["Deger"] = stu_df[y_col].astype(float)
	stu_df["Ogrenci"] = person["full_name"]

	if not stu_df.empty:
		stu_line = alt.Chart(stu_df).mark_line(color="#1e1b4b", strokeWidth=3.2, strokeDash=[4, 4]).encode(
			x=alt.X("Yas_Ay:Q"),
			y=alt.Y("Deger:Q"),
		)
		tooltips = [
			alt.Tooltip("Ogrenci:N", title="Öğrenci"),
			alt.Tooltip("Tarih:N", title="Ölçüm Tarihi"),
			alt.Tooltip("Yas_Ay:Q", title="Ölçüm Yaşı (Ay)"),
			alt.Tooltip("Deger:Q", title=f"Ölçülen Değer ({unit})", format=".1f"),
		]
		if pct_col and pct_col in stu_df.columns:
			tooltips.append(alt.Tooltip(f"{pct_col}:Q", title="Persantil (%)", format=".1f"))
		if z_col and z_col in stu_df.columns:
			tooltips.append(alt.Tooltip(f"{z_col}:Q", title="Z-Skoru (SDS)", format="+.2f"))

		stu_points = alt.Chart(stu_df).mark_circle(size=110, color="#b91c1c", stroke="#ffffff", strokeWidth=2).encode(
			x=alt.X("Yas_Ay:Q"),
			y=alt.Y("Deger:Q"),
			tooltip=tooltips,
		)
		final_chart = alt.layer(curve_chart, stu_line, stu_points).properties(height=380)
	else:
		final_chart = curve_chart.properties(height=380)

	st.altair_chart(final_chart, use_container_width=True)
	st.caption(f"📌 **Grafik Rehberi ({sex_label}):** Yeşil kalın çizgi **50. persentil (ortalama / medyan)**, sarı çizgiler **15p ve 85p (normal referans sınırları)**, kırmızı çizgiler ise **3p ve 97p (klinik alt/üst kritik sınırlar)** değerleridir. Kırmızı noktalı çizgi öğrencinin kendi büyüme izidir.")


def _render_dossier_tab_measurements(person: sqlite3.Row, person_id: int) -> None:
	st.markdown("#### ➕ Yeni Boy, Kilo veya Baş Çevresi Ölçümü Ekle")
	with st.expander("📏 Yeni Ölçüm Giriş Formunu Aç", expanded=True):
		mc1, mc2, mc3, mc4 = st.columns(4)
		m_date = mc1.date_input("Ölçüm Tarihi", value=date.today(), key=f"dossier_m_date_{person_id}")
		m_height = mc2.number_input("Boy (cm)", min_value=0.0, max_value=250.0, value=None, step=0.1, format="%.1f", key=f"dossier_m_h_{person_id}")
		m_weight = mc3.number_input("Kilo (kg)", min_value=0.0, max_value=300.0, value=None, step=0.1, format="%.1f", key=f"dossier_m_w_{person_id}")
		m_head_circ = mc4.number_input("Baş Çevresi (cm, opsiyonel)", min_value=0.0, max_value=100.0, value=None, step=0.1, format="%.1f", key=f"dossier_m_hc_{person_id}")

		# Instant live calculation preview
		has_hw_in = m_height is not None and m_weight is not None and m_height > 0 and m_weight > 0
		has_hc_in = m_head_circ is not None and m_head_circ > 0
		if has_hw_in or has_hc_in:
			prev_mo = age_in_months(person["birth_date"], m_date) if person["birth_date"] else None
			p_items = []
			if prev_mo is not None:
				p_items.append(f"**Ölçüm Yaşı:** {prev_mo // 12} yaş {prev_mo % 12} ay ({prev_mo} ay)")
			if has_hw_in:
				p_bmi = m_weight / ((m_height / 100.0) ** 2)
				p_items.append(f"**VKİ:** {p_bmi:.1f} kg/m²")
				p_ideal = calculate_ideal_weight(m_height, person["sex"], prev_mo, current_weight_kg=m_weight)
				if p_ideal:
					p_items.append(f"**🎯 İdeal Kilo Aralığı:** {p_ideal['min_kg']:.1f} – {p_ideal['max_kg']:.1f} kg (İdeal En Az: {p_ideal['min_kg']:.1f} kg · İdeal En Çok: {p_ideal['max_kg']:.1f} kg · {p_ideal['status_text']})")
				if person["sex"] and prev_mo is not None:
					ph_ev = evaluate_measurement(m_height, "height", person["sex"], prev_mo)
					pw_ev = evaluate_measurement(m_weight, "weight", person["sex"], prev_mo)
					pb_ev = evaluate_measurement(p_bmi, "bmi", person["sex"], prev_mo)
					if ph_ev: p_items.append(f"**Boy:** {ph_ev['z']:+.2f} SDS (%{ph_ev['percentile']:.1f}p)")
					if pw_ev: p_items.append(f"**Kilo:** {pw_ev['z']:+.2f} SDS (%{pw_ev['percentile']:.1f}p)")
					if pb_ev: p_items.append(f"**VKİ:** {pb_ev['z']:+.2f} SDS (%{pb_ev['percentile']:.1f}p)")
			if has_hc_in:
				phc_ev = evaluate_measurement(m_head_circ, "head_circ", person["sex"], prev_mo) if (person["sex"] and prev_mo is not None) else None
				if phc_ev:
					p_items.append(f"**Baş Çevresi:** {m_head_circ:.1f} cm · {phc_ev['z']:+.2f} SDS (%{phc_ev['percentile']:.1f}p)")
				else:
					p_items.append(f"**Baş Çevresi:** {m_head_circ:.1f} cm")
			st.info("⚡ **Canlı Önizleme:** " + "  |  ".join(p_items))

		if st.button("💾 Ölçümü Bu Kişiye Kaydet", type="primary", key=f"dossier_btn_save_m_{person_id}"):
			if (m_height is not None and m_height > 0) != (m_weight is not None and m_weight > 0):
				st.error("Boy ve kilo ölçümünü birlikte girin.")
			elif not has_hw_in and not has_hc_in:
				st.error("Kaydetmek için boy-kilo veya baş çevresi girin.")
			else:
				summary_txt = record_linked_measurement(
					person_id, m_date,
					m_height if has_hw_in else None,
					m_weight if has_hw_in else None,
					person["birth_date"], person["sex"],
					head_circ_cm=m_head_circ if has_hc_in else None,
				)
				st.success(f"Ölçüm kaydedildi: {summary_txt}")
				st.rerun()

	with connect_db() as connection:
		records = pd.read_sql_query(
			"""SELECT id, measured_on AS Tarih, height_cm AS Boy_cm, weight_kg AS Kilo_kg,
					  head_circ_cm AS Bas_Cevresi_cm, age_months AS Yas_ay, bmi AS BMI
			   FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC""",
			connection, params=(person_id,)
		)
	if records.empty:
		st.info("Henüz ölçüm kaydı yok. Yukarıdaki formdan ilk ölçümü ekleyebilirsiniz.")
		return

	records_by_id = {int(r["id"]): r for _, r in records.iterrows()}
	record_ids_desc = list(reversed(list(records_by_id.keys())))
	active_meas_key = f"dossier_active_meas_{person_id}"
	if st.session_state.get(active_meas_key) not in records_by_id:
		st.session_state[active_meas_key] = record_ids_desc[0]

	active_meas_id = st.session_state[active_meas_key]
	latest = records_by_id[active_meas_id]
	months = int(latest["Yas_ay"]) if not pd.isna(latest["Yas_ay"]) else (
		age_in_months(person["birth_date"], date.fromisoformat(str(latest["Tarih"])[:10])) if person["birth_date"] else None
	)

	person_meas = get_person_latest_measurements(person_id)
	latest_h = float(latest["Boy_cm"]) if pd.notna(latest["Boy_cm"]) else None
	latest_w = float(latest["Kilo_kg"]) if pd.notna(latest["Kilo_kg"]) else None
	latest_b = float(latest["BMI"]) if pd.notna(latest["BMI"]) else None
	latest_hc = float(latest["Bas_Cevresi_cm"]) if (pd.notna(latest.get("Bas_Cevresi_cm")) and float(latest["Bas_Cevresi_cm"]) > 0) else None
	display_hc = latest_hc if latest_hc is not None else person_meas["head_circ_cm"]

	st.markdown("#### 🩺 Antropometrik Z-Skoru (SDS) ve Persantil Değerlendirme Tablosu")
	st.caption("💡 Tablodaki **'✏️ Ölçülen Değer'** hücresine çift tıklayarak Boy, Kilo veya Baş Çevresi değerini doğrudan güncelleyebilirsiniz.")

	can_eval_age = (months is not None and 0 <= months <= 228)
	can_eval_hc = (months is not None and 0 <= months <= 60)
	clinical_summary_df = get_clinical_growth_summary_table(
		latest_h, latest_w, latest_b, person["sex"], months,
		head_circ_cm=display_hc if can_eval_hc else None,
		always_include_head_circ=can_eval_hc
	)
	edited_clinical_df = st.data_editor(
		clinical_summary_df,
		width="stretch",
		hide_index=True,
		column_config=get_clinical_growth_column_config(),
		disabled=["Antropometrik Parametre", "50. Persentil (Medyan)", "Normal Referans Bandı (-2 ile +2 SDS)", "Z-Skoru (SDS)", "Persentil (%)", "Persentil Dilimi", "Klinik Değerlendirme"],
		key=f"dossier_clinical_editor_{person_id}_{active_meas_id}",
	)

	if not edited_clinical_df.empty:
		param_map = {str(r["Antropometrik Parametre"]): numeric_value(r["Ölçülen Değer"]) for _, r in edited_clinical_df.iterrows()}
		new_h_val = param_map.get("📏 Boy (cm)", latest_h)
		new_w_val = param_map.get("⚖️ Kilo (kg)", latest_w)
		new_hc_val = param_map.get("🧠 Baş Çevresi (cm)", display_hc)
		orig_h_val = round(float(latest_h), 1) if latest_h is not None else None
		orig_w_val = round(float(latest_w), 1) if latest_w is not None else None
		orig_hc_val = round(float(display_hc), 1) if display_hc is not None else None

		h_changed = (new_h_val is not None) != (orig_h_val is not None) or (new_h_val is not None and orig_h_val is not None and abs(new_h_val - orig_h_val) > 1e-4)
		w_changed = (new_w_val is not None) != (orig_w_val is not None) or (new_w_val is not None and orig_w_val is not None and abs(new_w_val - orig_w_val) > 1e-4)
		hc_changed = (new_hc_val is not None) != (orig_hc_val is not None) or (new_hc_val is not None and orig_hc_val is not None and abs(new_hc_val - orig_hc_val) > 1e-4)

		if h_changed or w_changed or hc_changed:
			if (new_h_val is not None and new_h_val <= 0) or (new_w_val is not None and new_w_val <= 0):
				st.error("Boy ve kilo sıfırdan büyük olmalıdır.")
			else:
				new_bmi = (new_w_val / ((new_h_val / 100.0) ** 2)) if (new_h_val and new_w_val and new_h_val > 0 and new_w_val > 0) else None
				orig_date = date.fromisoformat(str(latest["Tarih"])[:10])
				new_months = age_in_months(person["birth_date"], orig_date) if person["birth_date"] else None
				clean_hc = float(new_hc_val) if (new_hc_val is not None and new_hc_val > 0) else None
				with connect_db() as connection:
					connection.execute(
						"UPDATE measurements SET height_cm=?, weight_kg=?, head_circ_cm=?, age_months=?, bmi=? WHERE id=? AND student_id=?",
						(float(new_h_val) if new_h_val else None, float(new_w_val) if new_w_val else None, clean_hc, new_months, float(new_bmi) if new_bmi else None, active_meas_id, person_id),
					)
				save_person_data_file(person_id)
				st.success("Tablo üzerinden ölçüm güncellendi.")
				st.rerun()

	# WHO Clinical Growth Curves & Percentile tracking
	st.markdown("#### 📈 WHO Klinik Büyüme Eğrileri ve Takip Grafiği")
	enriched_records = enrich_measurements_with_percentiles(records, person["sex"], person["birth_date"])

	if not can_eval_age:
		if not person["birth_date"]:
			st.warning("⚠️ Bu öğrencinin doğum tarihi girilmediği için yaşa göre persantil ve Z-skoru büyüme eğrileri değerlendirilememektedir. '✏️ Kimlik & İletişim Düzenle' sekmesinden doğum tarihini kaydedebilirsiniz.")
		else:
			st.info("ℹ️ Kişi 19 yaşından büyük (yetişkin) olduğu için çocukluk çağı WHO büyüme eğrileri ve persantil kanalları uygulanmamaktadır.")
	else:
		if can_eval_hc:
			chart_sub_tabs = st.tabs(["📏 Boy", "⚖️ Kilo", "🧬 VKİ", "🧠 Baş Çevresi", "📊 Kronolojik Takip"])
			with chart_sub_tabs[0]:
				render_clinical_growth_curve(enriched_records, person, "height", "Boy_cm", "Boy", "cm", default_max_mo=228)
			with chart_sub_tabs[1]:
				render_clinical_growth_curve(enriched_records, person, "weight", "Kilo_kg", "Kilo", "kg", default_max_mo=120)
			with chart_sub_tabs[2]:
				render_clinical_growth_curve(enriched_records, person, "bmi", "BMI", "VKİ (BMI)", "kg/m²", default_max_mo=228)
			with chart_sub_tabs[3]:
				render_clinical_growth_curve(enriched_records, person, "head_circ", "Bas_Cevresi_cm", "Baş Çevresi", "cm", default_max_mo=60)
			tab_chrono = chart_sub_tabs[4]
		else:
			chart_sub_tabs = st.tabs(["📏 Boy", "⚖️ Kilo", "🧬 VKİ", "📊 Kronolojik Takip"])
			with chart_sub_tabs[0]:
				render_clinical_growth_curve(enriched_records, person, "height", "Boy_cm", "Boy", "cm", default_max_mo=228)
			with chart_sub_tabs[1]:
				render_clinical_growth_curve(enriched_records, person, "weight", "Kilo_kg", "Kilo", "kg", default_max_mo=120)
			with chart_sub_tabs[2]:
				render_clinical_growth_curve(enriched_records, person, "bmi", "BMI", "VKİ (BMI)", "kg/m²", default_max_mo=228)
			tab_chrono = chart_sub_tabs[3]

		with tab_chrono:
			st.caption("Kişinin zaman içindeki persantil dilim değişimi:")
			c_rec = enriched_records.assign(Tarih=pd.to_datetime(enriched_records["Tarih"]))
			av_vars = [c for c in ["Boy_Persantil", "Kilo_Persantil", "BMI_Persantil", "Bas_Cevresi_Persantil"] if c in c_rec.columns and c_rec[c].notna().any()]
			if av_vars:
				p_long = c_rec.melt(id_vars=["Tarih"], value_vars=av_vars, var_name="Gösterge", value_name="Persantil").dropna(subset=["Persantil"])
				p_long["Gösterge"] = p_long["Gösterge"].map({
					"Boy_Persantil": "📏 Boy (%p)", "Kilo_Persantil": "⚖️ Kilo (%p)", "BMI_Persantil": "🧬 VKİ (%p)", "Bas_Cevresi_Persantil": "🧠 Baş Çevresi (%p)"
				})
				p_chart = alt.Chart(p_long).mark_line(point=True).encode(
					x=alt.X("Tarih:T", title="Tarih"),
					y=alt.Y("Persantil:Q", title="Persantil (%)", scale=alt.Scale(domain=[0, 100])),
					color=alt.Color("Gösterge:N", title="Parametre"),
					tooltip=["Tarih:T", "Gösterge:N", alt.Tooltip("Persantil:Q", format=".1f")]
				).properties(height=320)
				st.altair_chart(p_chart, use_container_width=True)

	# Chronological table & deletion
	st.markdown("#### 📅 Kronolojik Ölçüm Geçmişi")
	disp_cols = ["id", "Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Kilo_kg", "Kilo_Z", "Kilo_Persantil", "BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"]
	if can_eval_hc or (enriched_records["Bas_Cevresi_cm"].notna() & (enriched_records["Bas_Cevresi_cm"] > 0)).any():
		disp_cols[9:9] = ["Bas_Cevresi_cm", "Bas_Cevresi_Z", "Bas_Cevresi_Persantil"]
	if not can_eval_age:
		disp_cols = [c for c in disp_cols if "_Z" not in c and "Persantil" not in c and c != "Gelisim_Degerlendirmesi"]
	disp_hist = enriched_records[[c for c in disp_cols if c in enriched_records.columns]].copy()
	st.dataframe(disp_hist.drop(columns=["id"], errors="ignore"), width="stretch", hide_index=True)
	del_m1, del_m2 = st.columns([3, 1])
	with del_m2:
		m_opts = {f"{r['Tarih']} - Boy: {r['Boy_cm']} cm, Kilo: {r['Kilo_kg']} kg (#{r['id']})": int(r["id"]) for _, r in records.iterrows()}
		sel_del_m = st.selectbox("Silinecek ölçüm", list(m_opts.keys()), key=f"dossier_del_m_sel_{person_id}")
		if st.button("🗑️ Seçili Ölçümü Sil", type="secondary", key=f"dossier_btn_del_m_{person_id}", use_container_width=True):
			delete_measurement(m_opts[sel_del_m])
			st.success("Ölçüm kaydı silindi.")
			st.rerun()


def _render_dossier_tab_vitals(person: sqlite3.Row, person_id: int) -> None:
	st.markdown("#### ➕ Yeni Vital Bulgusu Kaydet")
	with st.expander("💓 Yeni Vital Giriş Formunu Aç", expanded=True):
		with st.form(f"dossier_form_new_vital_{person_id}"):
			vt_c1, vt_c2 = st.columns(2)
			vr_date = vt_c1.date_input("Ölçüm Tarihi", value=date.today(), key=f"dossier_vr_date_{person_id}")
			vr_time = vt_c2.time_input("Ölçüm Saati", value=datetime.now().time(), key=f"dossier_vr_time_{person_id}")

			v1, v2, v3, v4 = st.columns(4)
			v_temp = v1.number_input("Ateş (°C)", min_value=32.0, max_value=43.0, value=None, step=0.1, format="%.1f", key=f"dossier_vr_temp_{person_id}")
			v_spo2 = v2.number_input("SpO2 Saturasyon (%)", min_value=50.0, max_value=100.0, value=None, step=1.0, key=f"dossier_vr_spo2_{person_id}")
			v_sys = v3.number_input("Sistolik (Büyük Tansiyon)", min_value=40.0, max_value=250.0, value=None, step=1.0, key=f"dossier_vr_sys_{person_id}")
			v_dia = v4.number_input("Diyastolik (Küçük Tansiyon)", min_value=20.0, max_value=150.0, value=None, step=1.0, key=f"dossier_vr_dia_{person_id}")

			v5, v6, v7 = st.columns(3)
			v_pulse = v5.number_input("Nabız (atım/dk)", min_value=30.0, max_value=220.0, value=None, step=1.0, key=f"dossier_vr_pls_{person_id}")
			v_gluc = v6.number_input("Kan Şekeri (mg/dL)", min_value=20.0, max_value=600.0, value=None, step=1.0, key=f"dossier_vr_glu_{person_id}")
			v_gctx = v7.selectbox("Şeker Durumu", ["Belirtilmedi", "Açlık", "Tokluk", "Rastgele"], key=f"dossier_vr_gctx_{person_id}")

			v_notes = st.text_input("Ek Not / Açıklama", placeholder="Örn: Dinlenmiş vaziyette ölçüldü", key=f"dossier_vr_not_{person_id}")
			save_vr = st.form_submit_button("💾 Vital Bulguyu Bu Kişiye Kaydet", type="primary", use_container_width=True)

		if save_vr:
			if all(x is None for x in [v_temp, v_spo2, v_sys, v_dia, v_pulse, v_gluc]):
				st.error("Kaydetmek için en az bir vital değer girin.")
			else:
				meas_at = f"{vr_date.isoformat()} {vr_time.strftime('%H:%M')}"
				summary_v = record_linked_vital(
					person_id, meas_at, v_sys, v_dia, v_pulse, v_temp, v_gluc,
					glucose_context=None if v_gctx == "Belirtilmedi" else v_gctx,
					notes=v_notes.strip() or None,
					spo2=v_spo2
				)
				st.success(f"Vital bulgu kaydedildi: {summary_v}")
				st.rerun()

	st.markdown("#### 📋 Geçmiş Vital Ölçümleri")
	with connect_db() as connection:
		vitals_df = pd.read_sql_query(
			"""SELECT id, measured_at AS Tarih_Saat, temperature_c AS Ateş_C, spo2_pct AS SpO2,
					  systolic_mmhg AS Sistolik, diastolic_mmhg AS Diyastolik, pulse_bpm AS Nabız,
					  glucose_mg_dl AS Şeker, glucose_context AS Şeker_Durumu, notes AS Notlar
			   FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC""",
			connection, params=(person_id,)
		)
	if vitals_df.empty:
		st.info("Bu kişi için kayıtlı vital bulgu bulunmuyor.")
	else:
		st.dataframe(vitals_df.drop(columns=["id"]), width="stretch", hide_index=True)
		del_v1, del_v2 = st.columns([3, 1])
		with del_v2:
			vr_opts = {f"{r['Tarih_Saat']} (#{r['id']})": int(r["id"]) for _, r in vitals_df.iterrows()}
			sel_del_vr = st.selectbox("Silinecek vital", list(vr_opts.keys()), key=f"dossier_del_vr_sel_{person_id}")
			if st.button("🗑️ Seçili Vital Kaydını Sil", type="secondary", key=f"dossier_btn_del_vr_{person_id}", use_container_width=True):
				delete_vital(vr_opts[sel_del_vr])
				st.success("Vital kaydı silindi.")
				st.rerun()


def _render_dossier_tab_labs(person: sqlite3.Row, person_id: int) -> None:
	st.markdown("#### ➕ Yeni Tahlil Belgesi (PDF / Görsel) Yükle")
	with st.expander("🧪 Tahlil / Laboratuvar Belgesi Yükleme", expanded=True):
		uploaded_lab = st.file_uploader(
			f"📄 {person['full_name']} adına laboratuvar tahlil belgesi yükleyin",
			type=["pdf", "png", "jpg", "jpeg"],
			key=f"dossier_lab_uploader_{person_id}"
		)
		if uploaded_lab:
			fbytes = uploaded_lab.read()
			rep_date = st.date_input("Tahlil / Rapor Tarihi", value=date.today(), key=f"dossier_lab_date_{person_id}")
			st.caption("Belgeden çıkarılan laboratuvar test değerlerini kontrol edin:")
			try:
				if uploaded_lab.name.lower().endswith(".pdf"):
					import pymupdf
					doc = pymupdf.open(stream=fbytes, filetype="pdf")
					text_content = "\n".join([page.get_text() for page in doc])
				else:
					text_content = ""
				parsed_tests = extract_lab_results_from_text(text_content) if text_content else []
			except Exception:
				parsed_tests = []

			if not parsed_tests:
				parsed_tests = [
					{"test_name": "Hemoglobin (HGB)", "value": None, "unit": "g/dL", "ref_low": 11.5, "ref_high": 15.5},
					{"test_name": "Demir (Fe)", "value": None, "unit": "µg/dL", "ref_low": 50.0, "ref_high": 120.0},
					{"test_name": "Glukoz (Açlık)", "value": None, "unit": "mg/dL", "ref_low": 70.0, "ref_high": 100.0},
				]

			lab_df = pd.DataFrame(parsed_tests)
			edited_lab_df = st.data_editor(lab_df, num_rows="dynamic", width="stretch", key=f"dossier_lab_edit_{person_id}")
			if st.button("💾 Tahlil Sonuçlarını Bu Kişinin Dosyasına Kaydet", type="primary", key=f"dossier_save_lab_{person_id}"):
				with connect_db() as connection:
					cur = connection.execute(
						"INSERT INTO lab_reports(student_id, report_date, file_name, content, content_type) VALUES (?, ?, ?, ?, ?)",
						(person_id, rep_date.isoformat(), uploaded_lab.name, fbytes, uploaded_lab.type)
					)
					rep_id = cur.lastrowid
					for _, lr in edited_lab_df.iterrows():
						t_name = str(lr.get("test_name") or "").strip()
						t_val = numeric_value(lr.get("value"))
						if t_name and t_val is not None:
							connection.execute(
								"INSERT INTO lab_results(report_id, test_name, value, unit, ref_low, ref_high, lab_flag) VALUES (?, ?, ?, ?, ?, ?, ?)",
								(rep_id, t_name, float(t_val), lr.get("unit"), numeric_value(lr.get("ref_low")), numeric_value(lr.get("ref_high")), lr.get("lab_flag"))
							)
				save_person_data_file(person_id)
				st.success("Tahlil sonuçları kaydedildi.")
				st.rerun()

	st.markdown("#### 📋 Kayıtlı Tahlil Sonuçları & Orijinal Belgeler")
	with connect_db() as connection:
		lab_history = pd.read_sql_query(
			"""SELECT r.report_date AS Tarih, r.file_name AS Dosya, x.test_name AS Test,
					  x.value AS Sonuç, x.unit AS Birim, x.ref_low AS Referans_Alt, x.ref_high AS Referans_Üst,
					  x.lab_flag AS Durum FROM lab_results x JOIN lab_reports r ON r.id=x.report_id
			   WHERE r.student_id=? ORDER BY r.report_date DESC, x.id""",
			connection, params=(person_id,)
		)
		lab_files = connection.execute("SELECT id, report_date, file_name, content, content_type FROM lab_reports WHERE student_id=? ORDER BY report_date DESC", (person_id,)).fetchall()

	if not lab_history.empty:
		st.dataframe(lab_history, width="stretch", hide_index=True)
	else:
		st.info("Kayıtlı tahlil sonucu bulunmuyor.")

	if lab_files:
		st.markdown("###### 📥 Orijinal Tahlil Belgelerini İndir:")
		for lf in lab_files:
			st.download_button(f"📥 {lf['report_date']} · {lf['file_name']}", bytes(lf["content"]), file_name=lf["file_name"], mime=lf["content_type"], key=f"dossier_dl_lab_{lf['id']}")


def _render_dossier_tab_vaccinations(person: sqlite3.Row, person_id: int, hp: sqlite3.Row | None) -> None:
	vac_summary = get_student_vaccination_summary(person_id, person["birth_date"])
	vac_records = vac_summary["records"]
	cur_months = age_in_months(person["birth_date"], date.today()) if person["birth_date"] else None
	age_desc = f"{cur_months // 12} yaş {cur_months % 12} ay ({cur_months} ay)" if cur_months is not None else "Doğum tarihi belirtilmedi"

	st.markdown("#### 💉 T.C. Sağlık Bakanlığı Çocukluk Dönemi Aşı Takvimi ve İzlem Formu")
	st.caption("Öğrencinin resmi aşı takvimini inceleyebilir, yapılan aşıları takvim üstünde işaretleyebilir veya resmi Aşı Kartı çıktısı alabilirsiniz.")

	# Top Metric Cards
	m_col1, m_col2, m_col3, m_col4 = st.columns(4)
	with m_col1:
		if vac_summary["is_up_to_date"]:
			st.metric("Aşı Durumu", f"%{vac_summary['percent']}", delta="🟢 Takvime Tam Uygun")
		elif vac_summary["overdue_list"]:
			st.metric("Aşı Durumu", f"%{vac_summary['percent']}", delta=f"⚠️ {len(vac_summary['overdue_list'])} Eksik Doz", delta_color="inverse")
		else:
			st.metric("Aşı Durumu", f"%{vac_summary['percent']}", delta="Kayıt Bekleniyor")
	with m_col2:
		st.metric("Yapılan Rutin Doz", f"{vac_summary['applicable_done_count']} / {vac_summary['applicable_total']}")
	with m_col3:
		st.metric("Gecikmiş / Eksik Doz", f"{len(vac_summary['overdue_list'])} Doz")
	with m_col4:
		st.metric("Öğrenci Yaşı", age_desc.split(" (")[0])

	if vac_summary["overdue_list"]:
		st.warning(f"⚠️ **Yaşına Göre Eksik / Gecikmiş Aşılar ({len(vac_summary['overdue_list'])}):** {', '.join(vac_summary['overdue_list'])} — Lütfen veli veya aile hekimliği ile iletişime geçiniz.")

	# Action Toolbar
	t_c1, t_c2, t_c3, t_c4 = st.columns([2.2, 1.8, 1.3, 2.2])
	with t_c1:
		if st.button("✅ Yaşına Uygun Aşıları Tamamla", key=f"vac_btn_fill_age_{person_id}", type="primary", use_container_width=True):
			cnt = quick_fill_routine_vaccinations_for_student(person_id, person["birth_date"], up_to_age=True)
			st.success(f"{cnt} adet yaşa uygun aşı yapıldı olarak işaretlendi.")
			st.rerun()
	with t_c2:
		if st.button("📋 Tüm Takvimi İşaretle", key=f"vac_btn_fill_all_{person_id}", use_container_width=True):
			cnt = quick_fill_routine_vaccinations_for_student(person_id, person["birth_date"], up_to_age=False)
			st.success("Tüm rutin aşılar yapıldı olarak işaretlendi.")
			st.rerun()
	with t_c3:
		if st.button("🔄 Sıfırla", key=f"vac_btn_clear_{person_id}", use_container_width=True):
			clear_student_vaccinations(person_id)
			st.warning("Tüm aşı kayıtları temizlendi.")
			st.rerun()
	with t_c4:
		vac_docx = build_vaccination_card_docx(person_id)
		st.download_button(
			"📝 Aşı Kartını İndir (.docx)",
			data=vac_docx,
			file_name=f"asi_karti_{person['full_name']}_{date.today().isoformat()}.docx",
			mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
			type="primary",
			key=f"dl_vac_card_{person_id}",
			use_container_width=True,
		)

	st.markdown("---")
	st.markdown("##### 🏛️ T.C. Sağlık Bakanlığı Resmi Aşı Takvimi Matrisi")
	st.caption("Aşağıdaki takvimde antijenlerin aylara göre doz dağılımı ve öğrencinin mevcut aşı durumu gösterilmektedir:")
	st.markdown(build_vaccination_matrix_html(person_id, person["birth_date"]), unsafe_allow_html=True)

	st.markdown("---")
	st.markdown("##### ✍️ Aşı İşaretleme ve Detaylı Düzenleme Formu")
	st.caption("Aşağıdaki dönem sekmelerinden öğrencinin uygulanan aşılarını işaretleyebilir, tarih ve sağlık kuruluşu notu ekleyebilirsiniz:")

	v_tab_infant, v_tab_toddler, v_tab_school, v_tab_extra = st.tabs([
		"👶 Bebeklik Dönemi (0 - 6. Ay)",
		"🧒 Küçük Çocukluk Dönemi (12 - 24. Ay)",
		"🎒 Okul Çağı Aşıları (İlkokul 1 & 8. Sınıf)",
		"⭐ Özel & İsteğe Bağlı Aşılar (Opsiyonel)",
	])

	infant_vaccines = [v for v in MOH_VACCINE_SCHEDULE if v["type"] == "routine" and v["period_months"] <= 6]
	toddler_vaccines = [v for v in MOH_VACCINE_SCHEDULE if v["type"] == "routine" and 12 <= v["period_months"] <= 24]
	school_vaccines = [v for v in MOH_VACCINE_SCHEDULE if v["type"] == "routine" and v["period_months"] >= 48]
	extra_vaccines = [v for v in MOH_VACCINE_SCHEDULE if v["type"] == "optional"]

	with st.form(f"form_vaccinations_checklist_{person_id}"):
		vac_form_entries = {}

		def _render_vac_group(group_list: list[dict]):
			for v in group_list:
				vcode = v["code"]
				rec = vac_records.get(vcode, {})
				is_done = bool(rec.get("is_done"))
				def_date_str = rec.get("administered_date")
				try:
					def_date = date.fromisoformat(def_date_str) if def_date_str else date.today()
				except Exception:
					def_date = date.today()
				cur_fac = rec.get("facility") or "Aile Sağlığı Merkezi (ASM)"
				cur_notes = rec.get("notes") or ""

				with st.container(border=True):
					c1, c2, c3, c4 = st.columns([2.5, 1.2, 1.8, 1.5])
					with c1:
						chk = st.checkbox(
							f"💉 **{v['vaccine']}** — {v['dose']} *({v['period_label']})*",
							value=is_done,
							key=f"vchk_{person_id}_{vcode}",
							help=v["desc"],
						)
					with c2:
						adm_d = st.date_input("Uygulama Tarihi", value=def_date, key=f"vdt_{person_id}_{vcode}")
					with c3:
						fac_opts = ["Aile Sağlığı Merkezi (ASM)", "Okul Reviri", "Devlet Hastanesi", "Özel Hastane / Klinik", "Diğer"]
						f_idx = fac_opts.index(cur_fac) if cur_fac in fac_opts else 0
						fac_val = st.selectbox("Kurum", fac_opts, index=f_idx, key=f"vfac_{person_id}_{vcode}")
					with c4:
						note_val = st.text_input("Lot / Not", value=cur_notes, placeholder="Örn: Lot no veya not", key=f"vnot_{person_id}_{vcode}")

					vac_form_entries[vcode] = {
						"vaccine_code": vcode,
						"is_done": chk,
						"administered_date": adm_d.isoformat() if chk else None,
						"facility": fac_val if chk else None,
						"lot_number": None,
						"notes": note_val.strip() if (chk and note_val.strip()) else None,
					}

		with v_tab_infant:
			_render_vac_group(infant_vaccines)
		with v_tab_toddler:
			_render_vac_group(toddler_vaccines)
		with v_tab_school:
			_render_vac_group(school_vaccines)
		with v_tab_extra:
			_render_vac_group(extra_vaccines)

		st.write("")
		sub_c1, sub_c2 = st.columns([3, 1])
		with sub_c1:
			submitted = st.form_submit_button("💾 Aşı Kayıtlarını Kaydet ve Senkronize Et", type="primary", use_container_width=True)

		if submitted:
			records_to_save = list(vac_form_entries.values())
			save_bulk_student_vaccinations(person_id, records_to_save)
			st.success("✅ Aşı kayıtları başarıyla kaydedildi ve tüm profil ile eşitlendi.")
			st.rerun()


def compile_person_clinical_facts(person_id: int) -> str:
	with connect_db() as connection:
		profile = connection.execute("SELECT * FROM health_profiles WHERE student_id=?", (person_id,)).fetchone()
		latest_vital = connection.execute(
			"SELECT measured_at, systolic_mmhg, diastolic_mmhg, pulse_bpm, temperature_c, spo2_pct, glucose_mg_dl FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC LIMIT 1",
			(person_id,),
		).fetchone()
		latest_lab = connection.execute(
			"SELECT r.report_date, x.test_name, x.value, x.unit, x.ref_low, x.ref_high FROM lab_results x JOIN lab_reports r ON r.id=x.report_id WHERE r.student_id=? ORDER BY r.report_date DESC, x.id DESC LIMIT 8",
			(person_id,),
		).fetchall()
		recent_visits = connection.execute(
			"SELECT visit_date, complaint, symptoms, observation, action FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC LIMIT 3",
			(person_id,),
		).fetchall()

	facts = []
	if profile:
		profile_labels = {
			"allergies": "Bilinen alerjiler",
			"chronic_conditions": "Kronik hastalıklar",
			"medical_history": "Sağlık geçmişi",
			"medications": "Düzenli ilaçlar",
			"blood_group": "Kan grubu",
			"emergency_notes": "Acil durum notları",
		}
		for field, lbl in profile_labels.items():
			if field in profile.keys() and profile[field]:
				facts.append(f"{lbl}: {profile[field]}")

	latest_ant = get_person_latest_measurements(person_id)
	if latest_ant and latest_ant["has_data"]:
		m_parts = []
		if latest_ant["height_cm"] is not None:
			h_ev = latest_ant["height_eval"]
			m_parts.append(f"Boy {latest_ant['height_cm']:.1f} cm" + (f" ({h_ev['z']:+.2f} SDS, %{h_ev['percentile']:.1f}p - {h_ev['status']})" if h_ev else ""))
		if latest_ant["weight_kg"] is not None:
			w_ev = latest_ant["weight_eval"]
			m_parts.append(f"Kilo {latest_ant['weight_kg']:.1f} kg" + (f" ({w_ev['z']:+.2f} SDS, %{w_ev['percentile']:.1f}p - {w_ev['status']})" if w_ev else ""))
		if latest_ant["bmi"] is not None:
			b_ev = latest_ant["bmi_eval"]
			m_parts.append(f"VKİ {latest_ant['bmi']:.1f} kg/m²" + (f" ({b_ev['z']:+.2f} SDS, %{b_ev['percentile']:.1f}p - {b_ev['status']})" if b_ev else ""))
		if latest_ant["head_circ_cm"] is not None:
			hc_ev = latest_ant["head_circ_eval"]
			hc_dt = f" ({latest_ant['head_circ_date']})" if latest_ant["is_head_circ_from_earlier"] else ""
			m_parts.append(f"Baş çevresi {latest_ant['head_circ_cm']:.1f} cm{hc_dt}" + (f" ({hc_ev['z']:+.2f} SDS, %{hc_ev['percentile']:.1f}p - {hc_ev['status']})" if hc_ev else ""))
		if m_parts:
			facts.append(f"Antropometrik ölçüm ({latest_ant['latest_measured_on']}): " + ", ".join(m_parts))

	if latest_vital:
		v_meas = []
		if latest_vital["systolic_mmhg"] is not None:
			v_meas.append(f"Tansiyon {latest_vital['systolic_mmhg']:g}/{latest_vital['diastolic_mmhg']:g} mmHg")
		for fld, lbl, unt in [("pulse_bpm", "Nabız", "atım/dk"), ("temperature_c", "Ateş", "°C"), ("spo2_pct", "SpO2", "%"), ("glucose_mg_dl", "Şeker", "mg/dL")]:
			if fld in latest_vital.keys() and latest_vital[fld] is not None:
				v_meas.append(f"{lbl} {latest_vital[fld]:g} {unt}")
		facts.append(f"Vital bulgular ({latest_vital['measured_at']}): " + ", ".join(v_meas))

	for result in latest_lab:
		facts.append(f"Tahlil ({result['report_date']}): {result['test_name']} {result['value']:g} {result['unit'] or ''}")

	for rv in recent_visits:
		facts.append(f"Revir ziyareti ({rv['visit_date']} - {rv['complaint']}): {rv['symptoms'] or ''} | {rv['observation'] or ''} | İşlem: {rv['action'] or ''}")

	return "\n".join(facts) if facts else "Henüz doğrulanmış sağlık profili, ölçüm veya revir kaydı bulunmuyor."


def render_nanda_care_plan_editor(person_id: int, prefix: str = "cp") -> None:
	app_cfg = get_app_settings()
	default_nurse = app_cfg.get("nurse_name", "Okul Revir Sorumlusu")
	compiled_facts = compile_person_clinical_facts(person_id)

	# Session state keys for this person
	k_diag = f"{prefix}_diag_{person_id}"
	k_facts = f"{prefix}_facts_{person_id}"
	k_goals = f"{prefix}_goals_{person_id}"
	k_interv = f"{prefix}_interv_{person_id}"
	k_mon = f"{prefix}_mon_{person_id}"
	k_esc = f"{prefix}_esc_{person_id}"
	k_rev = f"{prefix}_rev_{person_id}"
	k_stat = f"{prefix}_stat_{person_id}"

	# Initialize defaults in session state if not set
	if k_diag not in st.session_state:
		st.session_state[k_diag] = ""
	if k_facts not in st.session_state:
		st.session_state[k_facts] = compiled_facts
	if k_goals not in st.session_state:
		st.session_state[k_goals] = ""
	if k_interv not in st.session_state:
		st.session_state[k_interv] = ""
	if k_mon not in st.session_state:
		st.session_state[k_mon] = ""
	if k_esc not in st.session_state:
		st.session_state[k_esc] = ""
	if k_rev not in st.session_state or not st.session_state[k_rev]:
		st.session_state[k_rev] = default_nurse
	if k_stat not in st.session_state:
		st.session_state[k_stat] = "Aktif - İzlemde"

	care_plan_templates = get_care_plan_templates()
	template_keys = list(care_plan_templates.keys())

	with st.expander("➕ Yeni Hemşirelik Bakım Planı Oluştur", expanded=True):
		# Master Template Quick Loader
		with st.container():
			st.markdown("###### ⚡ (İsteğe Bağlı) Tüm Planı Tek Seferde Doldur")
			m_col1, m_col2 = st.columns([3.2, 1.3])
			with m_col1:
				chosen_master_tpl = st.selectbox(
					"Tüm Alanları Dolduracak Hazır NANDA-I Protokolü",
					["-- Seçin (Tüm Formu Doldurmak İçin) --"] + template_keys,
					key=f"{prefix}_master_tpl_sel_{person_id}",
				)
			with m_col2:
				st.write("")
				st.write("")
				if st.button("🚀 Tüm Forma Uygula", key=f"{prefix}_apply_all_btn_{person_id}", use_container_width=True):
					if chosen_master_tpl in care_plan_templates:
						t_data = care_plan_templates[chosen_master_tpl]
						st.session_state[k_diag] = t_data["diagnosis"]
						st.session_state[k_facts] = (compiled_facts + "\n\n[Klinik Protokol Olgusu]: " + t_data["observed_facts"]).strip()
						st.session_state[k_goals] = t_data["goals"]
						st.session_state[k_interv] = t_data["interventions"]
						st.session_state[k_mon] = t_data["monitoring"]
						st.session_state[k_esc] = t_data["escalation"]
						st.session_state[k_stat] = t_data.get("status", "Aktif - İzlemde")
						st.success(f"'{chosen_master_tpl}' şablonu tüm alanlara aktarıldı!")
						st.rerun()

		st.markdown("---")

		# 1. HEMŞİRELİK TANISI
		st.markdown("##### 1. 🩺 Hemşirelik Tanısı (NANDA-I)")
		st.text_area(
			"Hemşirelik Tanısı",
			key=k_diag,
			placeholder="Örn: Büyümede yetersizlik ve gelişme geriliği riski (NANDA: 00113)...",
			height=85,
			label_visibility="collapsed",
		)
		# Section below field 1: Diagnosis templates
		diag_snippets = get_care_plan_section_templates("diagnoses")
		diag_c1, diag_c2 = st.columns([3.2, 1.3])
		with diag_c1:
			sel_diag_tpl = st.selectbox(
				"💡 Hazır NANDA-I Tanı Şablonu Seçin:",
				["-- Hazır Tanı Seçin --"] + list(diag_snippets.keys()),
				key=f"{prefix}_sub_diag_sel_{person_id}",
			)
		with diag_c2:
			st.write("")
			st.write("")
			if st.button("➕ Tanıyı Alana Yaz", key=f"{prefix}_sub_diag_btn_{person_id}", use_container_width=True):
				if sel_diag_tpl in diag_snippets:
					st.session_state[k_diag] = diag_snippets[sel_diag_tpl]
					st.success("Tanı alanı güncellendi.")
					st.rerun()

		st.write("")

		# 2. GÖZLENEN OLGULAR & BULGULAR
		st.markdown("##### 2. 🔍 Gözlenen Olgular, Ölçüm ve Tahlil Bulguları")
		st.text_area(
			"Gözlenen Olgular",
			key=k_facts,
			placeholder="Öğrencinin boy-kilo persantili, vital bulguları, alerji ve revir başvuru notları...",
			height=110,
			label_visibility="collapsed",
		)
		# Section below field 2: Facts templates & Auto-compile
		facts_snippets = get_care_plan_section_templates("facts")
		facts_c1, facts_c2, facts_c3 = st.columns([2.2, 1.5, 1.3])
		with facts_c1:
			sel_facts_tpl = st.selectbox(
				"💡 Hazır Klinik Olgu / Belirti Şablonu Seçin:",
				["-- Örnek Durum Olgusu Seçin --"] + list(facts_snippets.keys()),
				key=f"{prefix}_sub_facts_sel_{person_id}",
			)
		with facts_c2:
			st.write("")
			st.write("")
			if st.button("➕ Olgu Notunu Ekle", key=f"{prefix}_sub_facts_btn_{person_id}", use_container_width=True):
				if sel_facts_tpl in facts_snippets:
					add_txt = "[Klinik Durum Olgusu]: " + facts_snippets[sel_facts_tpl]
					cur_f = st.session_state.get(k_facts, "").strip()
					st.session_state[k_facts] = (cur_f + "\n\n" + add_txt).strip() if cur_f else add_txt
					st.success("Olgu notu eklendi.")
					st.rerun()
		with facts_c3:
			st.write("")
			st.write("")
			if st.button("🔄 Sistem Verilerini Çek", key=f"{prefix}_pull_facts_btn_{person_id}", use_container_width=True):
				st.session_state[k_facts] = compile_person_clinical_facts(person_id)
				st.success("Sistemdeki güncel veriler çekildi.")
				st.rerun()

		st.write("")

		# 3. BAKIM HEDEFLERİ
		st.markdown("##### 3. 🎯 Bakım Hedefleri (Kısa ve Uzun Vadeli Beklenen Sonuçlar)")
		st.text_area(
			"Bakım Hedefleri",
			key=k_goals,
			placeholder="Örn: 3 ay içinde boy-kilo kanal stabilitesi sağlanması, SpO2 > %95 tutulması...",
			height=90,
			label_visibility="collapsed",
		)
		# Section below field 3: Goals templates
		goals_snippets = get_care_plan_section_templates("goals")
		goals_c1, goals_c2 = st.columns([3.2, 1.3])
		with goals_c1:
			sel_goals_tpl = st.selectbox(
				"💡 Hazır Bakım Hedefleri (NOC) Şablonu Seçin:",
				["-- Hazır Bakım Hedefi Seçin --"] + list(goals_snippets.keys()),
				key=f"{prefix}_sub_goals_sel_{person_id}",
			)
		with goals_c2:
			st.write("")
			st.write("")
			if st.button("➕ Hedefleri Alana Ekle", key=f"{prefix}_sub_goals_btn_{person_id}", use_container_width=True):
				if sel_goals_tpl in goals_snippets:
					g_txt = goals_snippets[sel_goals_tpl]
					cur_g = st.session_state.get(k_goals, "").strip()
					st.session_state[k_goals] = (cur_g + "\n" + g_txt).strip() if cur_g else g_txt
					st.success("Bakım hedefleri eklendi.")
					st.rerun()

		st.write("")

		# 4. HEMŞİRELİK GİRİŞİMLERİ VE MÜDAHALELER
		st.markdown("##### 4. 🛠️ Planlanan Hemşirelik Girişimleri ve Müdahaleler")
		st.text_area(
			"Hemşirelik Girişimleri",
			key=k_interv,
			placeholder="Örn: 15 günde bir persantil ölçümü, beslenme günlüğü, inhaler takip protokolü...",
			height=130,
			label_visibility="collapsed",
		)
		# Section below field 4: Interventions templates
		interv_snippets = get_care_plan_section_templates("interventions")
		interv_c1, interv_c2 = st.columns([3.2, 1.3])
		with interv_c1:
			sel_interv_tpl = st.selectbox(
				"💡 Hazır Hemşirelik Girişim Protokolü (NIC) Seçin:",
				["-- Hazır Girişim Protokolü Seçin --"] + list(interv_snippets.keys()),
				key=f"{prefix}_sub_interv_sel_{person_id}",
			)
		with interv_c2:
			st.write("")
			st.write("")
			if st.button("➕ Girişimleri Alana Ekle", key=f"{prefix}_sub_interv_btn_{person_id}", use_container_width=True):
				if sel_interv_tpl in interv_snippets:
					i_txt = interv_snippets[sel_interv_tpl]
					cur_i = st.session_state.get(k_interv, "").strip()
					st.session_state[k_interv] = (cur_i + "\n\n" + i_txt).strip() if cur_i else i_txt
					st.success("Girişim protokolü eklendi.")
					st.rerun()

		st.write("")

		# 5. İZLEM VE DEĞERLENDİRME ZAMANI
		st.markdown("##### 5. ⏱️ İzlem ve Değerlendirme Zamanı / Sıklığı")
		st.text_input(
			"İzlem Sıklığı",
			key=k_mon,
			placeholder="Örn: 15 günde bir persantil izlemi, haftalık kilo kontrolü...",
			label_visibility="collapsed",
		)
		# Section below field 5: Monitoring frequency templates
		mon_snippets = get_care_plan_section_templates("monitorings")
		mon_c1, mon_c2 = st.columns([3.2, 1.3])
		with mon_c1:
			sel_mon_tpl = st.selectbox(
				"💡 Hazır İzlem & Değerlendirme Sıklığı Seçin:",
				["-- Hazır İzlem Sıklığı Seçin --"] + list(mon_snippets.keys()),
				key=f"{prefix}_sub_mon_sel_{person_id}",
			)
		with mon_c2:
			st.write("")
			st.write("")
			if st.button("➕ Sıklığı Seç", key=f"{prefix}_sub_mon_btn_{person_id}", use_container_width=True):
				if sel_mon_tpl in mon_snippets:
					st.session_state[k_mon] = mon_snippets[sel_mon_tpl]
					st.success("İzlem sıklığı güncellendi.")
					st.rerun()

		st.write("")

		# 6. YÖNLENDİRME / SEVK / ACİL EYLEM KRİTERİ
		st.markdown("##### 6. 🚨 Yönlendirme / Sevk / Acil Eylem Kriteri")
		st.text_input(
			"Sevk Kriteri",
			key=k_esc,
			placeholder="Örn: SpO2 <%92 durumunda derhal 112 Acil, kilo kaybı devamında Çocuk Endokrinolojisine sevk...",
			label_visibility="collapsed",
		)
		# Section below field 6: Escalation criteria templates
		esc_snippets = get_care_plan_section_templates("escalations")
		esc_c1, esc_c2 = st.columns([3.2, 1.3])
		with esc_c1:
			sel_esc_tpl = st.selectbox(
				"💡 Hazır Sevk / Acil Eylem Kriteri Seçin:",
				["-- Hazır Sevk Kriteri Seçin --"] + list(esc_snippets.keys()),
				key=f"{prefix}_sub_esc_sel_{person_id}",
			)
		with esc_c2:
			st.write("")
			st.write("")
			if st.button("➕ Kriteri Seç", key=f"{prefix}_sub_esc_btn_{person_id}", use_container_width=True):
				if sel_esc_tpl in esc_snippets:
					st.session_state[k_esc] = esc_snippets[sel_esc_tpl]
					st.success("Sevk kriteri güncellendi.")
					st.rerun()

		st.write("")
		st.markdown("---")

		# REVIEWER & STATUS
		status_options = [
			"Aktif - İzlemde",
			"Taslak - Uzman İncelemesi Gerekli",
			"Uzman Değerlendirmesi Tamamlandı",
			"Tamamlandı - Hedefe Ulaşıldı",
			"İptal / İhtiyaç Kalmadı",
		]
		r_c1, r_c2 = st.columns(2)
		with r_c1:
			st.text_input("Sorumlu Sağlık Personeli / Hemşire", key=k_rev, placeholder=default_nurse or "Ad Soyad, Unvan")
		with r_c2:
			cur_s = st.session_state.get(k_stat, "Aktif - İzlemde")
			s_idx = status_options.index(cur_s) if cur_s in status_options else 0
			st.selectbox("Bakım Planı Durumu", status_options, index=s_idx, key=k_stat)

		st.write("")
		if st.button("💾 Hemşirelik Bakım Planını Kaydet", type="primary", use_container_width=True, key=f"{prefix}_save_final_btn_{person_id}"):
			val_diag = st.session_state.get(k_diag, "").strip()
			val_facts = st.session_state.get(k_facts, "").strip()
			val_goals = st.session_state.get(k_goals, "").strip()
			val_interv = st.session_state.get(k_interv, "").strip()
			val_mon = st.session_state.get(k_mon, "").strip()
			val_esc = st.session_state.get(k_esc, "").strip()
			val_rev = st.session_state.get(k_rev, "").strip()
			val_stat = st.session_state.get(k_stat, "Aktif - İzlemde")

			if not val_diag:
				st.error("Hemşirelik tanısı boş bırakılamaz. Lütfen tanıyı yazın veya hazır şablonlardan seçin.")
			else:
				with connect_db() as connection:
					connection.execute(
						"""INSERT INTO care_plans(student_id, created_at, reviewed_by, nursing_diagnosis, observed_facts,
						   goals, interventions, monitoring, escalation, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(person_id, datetime.now().isoformat(timespec="minutes"), val_rev or None,
						 val_diag, val_facts or None, val_goals or None,
						 val_interv or None, val_mon or None, val_esc or None, val_stat),
					)
				save_person_data_file(person_id)
				for k in [k_diag, k_facts, k_goals, k_interv, k_mon, k_esc]:
					st.session_state.pop(k, None)
				st.session_state[k_rev] = default_nurse
				st.success("Hemşirelik bakım planı başarıyla kaydedildi.")
				st.rerun()

	# Display existing care plans
	with connect_db() as connection:
		plans = connection.execute(
			"""SELECT id, created_at, reviewed_by, nursing_diagnosis, observed_facts, goals,
					  interventions, monitoring, escalation, status
			   FROM care_plans WHERE student_id = ? ORDER BY created_at DESC, id DESC""",
			(person_id,),
		).fetchall()

	st.markdown("##### 📑 Kayıtlı Bakım Planları")
	if not plans:
		st.caption("Kişiye ait henüz kayıtlı bir bakım planı bulunmuyor.")
		return

	status_badge_colors = {
		"Aktif - İzlemde": ("#16a34a", "#f0fdf4"),
		"Taslak - Uzman İncelemesi Gerekli": ("#d97706", "#fffbeb"),
		"Uzman Değerlendirmesi Tamamlandı": ("#0284c7", "#f0f9ff"),
		"Tamamlandı - Hedefe Ulaşıldı": ("#0d9488", "#f0fdfa"),
		"İptal / İhtiyaç Kalmadı": ("#64748b", "#f8fafc"),
	}

	for p in plans:
		p_id = p["id"]
		p_stat = p["status"] or "Aktif - İzlemde"
		fg_col, bg_col = status_badge_colors.get(p_stat, ("#475569", "#f1f5f9"))
		first_line_diag = (p['nursing_diagnosis'] or '').split('\n')[0]
		card_title = f"{p['created_at'][:10]} · {first_line_diag[:60]}..." if len(first_line_diag) > 60 else f"{p['created_at'][:10]} · {first_line_diag or 'Tanı Girilmedi'}"

		with st.expander(f"📋 {card_title} ({p_stat})", expanded=False):
			h_c1, h_c2, h_c3 = st.columns([2.5, 1.5, 1])
			with h_c1:
				st.markdown(
					f"<span style='background-color:{bg_col}; color:{fg_col}; font-weight:600; padding:3px 10px; border-radius:12px; border:1px solid {fg_col}; font-size:12px;'>● {p_stat}</span> &nbsp; "
					f"<small style='color:#64748b;'>Tarih: {p['created_at']} · Sorumlu: <b>{p['reviewed_by'] or 'Belirtilmedi'}</b></small>",
					unsafe_allow_html=True
				)
			with h_c3:
				if st.button("🗑️ Planı Sil", key=f"del_cp_{prefix}_{p_id}", type="secondary", use_container_width=True):
					with connect_db() as connection:
						connection.execute("DELETE FROM care_plans WHERE id=? AND student_id=?", (p_id, person_id))
					save_person_data_file(person_id)
					st.warning("Bakım planı silindi.")
					st.rerun()

			st.info(f"🩺 **Hemşirelik Tanısı (NANDA-I):** {p['nursing_diagnosis'] or '—'}")

			d_col1, d_col2 = st.columns(2)
			with d_col1:
				with st.container(border=True):
					st.markdown("**🔍 Gözlenen Olgular & Bulgular:**")
					st.write(p['observed_facts'] or '—')
				with st.container(border=True):
					st.markdown("**🎯 Bakım Hedefleri:**")
					st.write(p['goals'] or '—')
			with d_col2:
				with st.container(border=True):
					st.markdown("**🛠️ Hemşirelik Girişimleri:**")
					st.write(p['interventions'] or '—')
				with st.container(border=True):
					st.markdown(f"⏱️ **İzlem Sıklığı:** {p['monitoring'] or '—'}")
					st.markdown(f"🚨 **Yönlendirme / Sevk Kriteri:** {p['escalation'] or '—'}")


def _render_dossier_tab_care_plans(person: sqlite3.Row, person_id: int, hp: sqlite3.Row | None) -> None:
	st.markdown("#### 🩺 Sağlık Profilini Düzenle")
	cur_bg = normalize_blood_group(hp["blood_group"] if hp else None)
	bg_idx = BLOOD_GROUP_OPTIONS.index(cur_bg) if cur_bg in BLOOD_GROUP_OPTIONS else 0

	with st.form(f"dossier_hp_form_{person_id}"):
		p_c1, p_c2 = st.columns(2)
		f_bg = p_c1.selectbox("Kan Grubu", BLOOD_GROUP_OPTIONS, index=bg_idx, key=f"dossier_hp_bg_{person_id}")
		app_cfg = get_app_settings()
		cur_rev = (hp["reviewed_by"] if hp and hp["reviewed_by"] else "") or app_cfg.get("nurse_name", "Okul Revir Sorumlusu")
		f_rev = p_c2.text_input("Bilgileri Doğrulayan Görevli", value=cur_rev, key=f"dossier_hp_rev_{person_id}")

		f_al = st.text_area("🔴 Alerjiler", value=hp["allergies"] if hp and hp["allergies"] else "", placeholder="İlaç, gıda, polen, böcek alerjileri...", height=65, key=f"dossier_hp_al_{person_id}")
		f_ch = st.text_area("🟠 Kronik Durumlar", value=hp["chronic_conditions"] if hp and hp["chronic_conditions"] else "", placeholder="Astım, epilepsi, diyabet, kalp vb.", height=65, key=f"dossier_hp_ch_{person_id}")
		f_med = st.text_area("💊 Sürekli İlaçlar", value=hp["medications"] if hp and hp["medications"] else "", placeholder="Düzenli kullanılan ilaçlar ve dozları...", height=65, key=f"dossier_hp_med_{person_id}")
		f_mh = st.text_area("Tıbbi Geçmiş", value=hp["medical_history"] if hp and hp["medical_history"] else "", placeholder="Geçirilen ameliyatlar, hastane yatışları...", height=65, key=f"dossier_hp_mh_{person_id}")
		f_em = st.text_area("🚨 Acil Durum Notları", value=hp["emergency_notes"] if hp and hp["emergency_notes"] else "", placeholder="Acil bir durumda dikkat edilecek ilk yardım talimatları...", height=65, key=f"dossier_hp_em_{person_id}")

		st.markdown("###### 🩺 Okul Sağlık Taraması Bulguları")
		scr_f_c1, scr_f_c2 = st.columns(2)
		gl_opts = ["Hayır (Kullanmıyor)", "Evet (Sürekli)", "Bazen / Sadece Okurken", "Gözlük Önerildi / Muayene Bekleniyor", "Göz Tembelliği Kapaması Yapılıyor"]
		cur_gl = hp["wears_glasses"] if hp and hp["wears_glasses"] else "Hayır (Kullanmıyor)"
		f_gl = scr_f_c1.selectbox("Gözlük Kullanımı", gl_opts, index=gl_opts.index(cur_gl) if cur_gl in gl_opts else 0, key=f"dossier_hp_gl_{person_id}")
		f_hair = scr_f_c2.text_input("Saç Taraması (Pediküloz)", value=hp["hair_screening"] if hp and hp["hair_screening"] else "Temiz / Parazit Yok", key=f"dossier_hp_hair_{person_id}")

		eye_f_c1, eye_f_c2 = st.columns(2)
		f_vr = eye_f_c1.text_input("Sağ Göz Durumu", value=hp["vision_right"] if hp and hp["vision_right"] else "Tam / Normal (10/10)", key=f"dossier_hp_vr_{person_id}")
		f_vl = eye_f_c2.text_input("Sol Göz Durumu", value=hp["vision_left"] if hp and hp["vision_left"] else "Tam / Normal (10/10)", key=f"dossier_hp_vl_{person_id}")

		ear_f_c1, ear_f_c2 = st.columns(2)
		f_hr = ear_f_c1.text_input("Sağ Kulak İşitme", value=hp["hearing_right"] if hp and hp["hearing_right"] else "Normal / İyi Duyuyor", key=f"dossier_hp_hr_{person_id}")
		f_hl = ear_f_c2.text_input("Sol Kulak İşitme", value=hp["hearing_left"] if hp and hp["hearing_left"] else "Normal / İyi Duyuyor", key=f"dossier_hp_hl_{person_id}")

		dev_f_c1, dev_f_c2 = st.columns(2)
		f_beh = dev_f_c1.text_input("Sağlık Davranışı (Beslenme, Diş, Hijyen)", value=hp["health_behavior"] if hp and hp["health_behavior"] else "Beslenme ve Hijyen Çok İyi", key=f"dossier_hp_beh_{person_id}")
		f_psy = dev_f_c2.text_input("Genel Psikososyal Durum (Uyum, İletişim)", value=hp["psychosocial_status"] if hp and hp["psychosocial_status"] else "Sınıf İçi Uyum ve İletişimi Çok İyi", key=f"dossier_hp_psy_{person_id}")
		f_harm = st.text_input("⚠️ Zararlı Alışkanlıklar & Risk Faktörleri", value=hp["harmful_habits"] if hp and "harmful_habits" in hp.keys() and hp["harmful_habits"] else "Zararlı Alışkanlık Tespit Edilmedi", key=f"dossier_hp_harm_{person_id}")

		save_hp = st.form_submit_button("💾 Sağlık Profilini Güncelle", type="primary", use_container_width=True)

	if save_hp:
		res_bg = None if f_bg == "Belirtilmedi" else f_bg
		with connect_db() as connection:
			connection.execute(
				"""INSERT INTO health_profiles(student_id, blood_group, allergies, chronic_conditions, medications,
				   medical_history, emergency_notes, reviewed_by, updated_at, wears_glasses, vision_right,
				   vision_left, hair_screening, hearing_right, hearing_left, health_behavior, harmful_habits, psychosocial_status)
				   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				   ON CONFLICT(student_id) DO UPDATE SET blood_group=excluded.blood_group, allergies=excluded.allergies,
				   chronic_conditions=excluded.chronic_conditions, medications=excluded.medications,
				   medical_history=excluded.medical_history, emergency_notes=excluded.emergency_notes,
				   reviewed_by=excluded.reviewed_by, updated_at=excluded.updated_at,
				   wears_glasses=excluded.wears_glasses, vision_right=excluded.vision_right,
				   vision_left=excluded.vision_left, hair_screening=excluded.hair_screening,
				   hearing_right=excluded.hearing_right, hearing_left=excluded.hearing_left,
				   health_behavior=excluded.health_behavior, harmful_habits=excluded.harmful_habits, psychosocial_status=excluded.psychosocial_status""",
				(
					person_id, res_bg, f_al.strip() or None, f_ch.strip() or None, f_med.strip() or None,
					f_mh.strip() or None, f_em.strip() or None, f_rev.strip() or None,
					datetime.now().isoformat(timespec="minutes"), f_gl, f_vr.strip() or None,
					f_vl.strip() or None, f_hair.strip() or None, f_hr.strip() or None,
					f_hl.strip() or None, f_beh.strip() or None, f_harm.strip() or None, f_psy.strip() or None,
				)
			)
		save_person_data_file(person_id)
		st.success("Sağlık profili ve tarama bilgileri başarıyla güncellendi.")
		st.rerun()

	st.markdown("---")
	st.markdown("#### 📋 Hemşirelik Bakım Planı Yönetimi")
	render_nanda_care_plan_editor(person_id, prefix="dossier_cp")


def _render_dossier_tab_edit_demographics(person: sqlite3.Row, person_id: int, hp: sqlite3.Row | None) -> None:
	st.markdown("#### ✏️ Kimlik, Sınıf ve İletişim Bilgilerini Düzenle")
	app_cfg = get_app_settings()
	classes = [c for c in get_classes() if c != "Sivil / Diğer"]
	if "Sınıf Belirtilmedi" not in classes:
		classes = ["Sınıf Belirtilmedi", *classes]

	with st.form(f"dossier_edit_demo_{person_id}"):
		ed_c1, ed_c2, ed_c3 = st.columns([1.5, 1, 1])
		f_name = ed_c1.text_input("Ad Soyad", value=person["full_name"], key=f"dossier_ed_name_{person_id}")
		f_ptype = ed_c2.selectbox("Kayıt Türü", ["Öğrenci", "Sivil / Diğer"], index=0 if person["person_type"] == "Öğrenci" else 1, key=f"dossier_ed_ptype_{person_id}")
		if f_ptype == "Öğrenci":
			f_sno = ed_c3.text_input("Okul Numarası", value=str(person["student_no"] or ""), key=f"dossier_ed_sno_{person_id}")
			f_role = ""
		else:
			f_sno = ""
			f_role = ed_c3.text_input("Görev / Statü", value=person["notes"] or "", key=f"dossier_ed_role_{person_id}")

		ed_c4, ed_c5, ed_c6 = st.columns(3)
		cur_sex = {"K": "Kız", "E": "Erkek"}.get(person["sex"], "Belirtilmedi")
		f_sex = ed_c4.selectbox("Cinsiyet", ["Belirtilmedi", "Kız", "Erkek"], index=["Belirtilmedi", "Kız", "Erkek"].index(cur_sex), key=f"dossier_ed_sex_{person_id}")
		b_val = date.fromisoformat(person["birth_date"]) if person["birth_date"] else None
		f_birth = ed_c5.date_input("Doğum Tarihi", value=b_val, min_value=date(1940, 1, 1), max_value=date.today(), key=f"dossier_ed_birth_{person_id}")

		if f_ptype == "Öğrenci":
			cur_cls = person["class_name"]
			cls_opts = list(classes)
			if cur_cls and cur_cls not in cls_opts:
				cls_opts.insert(0, cur_cls)
			cls_idx = cls_opts.index(cur_cls) if cur_cls in cls_opts else 0
			f_class = ed_c6.selectbox("Sınıfı", cls_opts, index=cls_idx, key=f"dossier_ed_cls_{person_id}")
		else:
			f_class = "Sivil / Diğer"
			ed_c6.text_input("Sınıfı", value="Sivil / Diğer", disabled=True, key=f"dossier_ed_cls_dis_{person_id}")

		ed_c7, ed_c8 = st.columns(2)
		cur_g = (hp["guardian_name"] if hp and hp["guardian_name"] else person["guardian_name"]) or ""
		cur_ph = (hp["guardian_phone"] if hp and hp["guardian_phone"] else person["guardian_phone"]) or ""
		f_guard = ed_c7.text_input("Veli / Yakını Adı", value=cur_g, key=f"dossier_ed_guard_{person_id}")
		f_phone = ed_c8.text_input("İletişim Telefonu", value=cur_ph, key=f"dossier_ed_phone_{person_id}")

		f_addr = st.text_input("Açık Adres / Mahalle / Not", value=person["notes"] or "", key=f"dossier_ed_addr_{person_id}")

		st.markdown("###### 📍 Yerleşim Yeri (Ülke / İl / İlçe / Köy)")
		l_c1, l_c2, l_c3, l_c4 = st.columns(4)
		cur_cntry = person["country"] or app_cfg.get("default_country") or "Türkiye"
		c_idx = COUNTRY_OPTIONS.index(cur_cntry) if cur_cntry in COUNTRY_OPTIONS else 0
		f_country = l_c1.selectbox("Ülke", COUNTRY_OPTIONS, index=c_idx, key=f"dossier_ed_cntry_{person_id}")

		if f_country == "Türkiye":
			provinces = get_tr_provinces()
			cur_prov = person["province"] or app_cfg.get("default_province") or "Karaman"
			p_idx = provinces.index(cur_prov) if cur_prov in provinces else 0
			f_province = l_c2.selectbox("İl", provinces, index=p_idx, key=f"dossier_ed_prov_{person_id}")
			districts = get_tr_districts(f_province)
			cur_dist = person["district"] or app_cfg.get("default_district") or "Merkez"
			d_idx = districts.index(cur_dist) if cur_dist in districts else 0
			f_district = l_c3.selectbox("İlçe", districts, index=d_idx, key=f"dossier_ed_dist_{person_id}")
			villages = get_tr_villages(f_province, f_district)
			vil_opts = ["İlçe Merkezi / Belirtilmedi", *villages]
			cur_vil = person["village"] or ""
			v_idx = vil_opts.index(cur_vil) if cur_vil in vil_opts else 0
			f_vil_raw = l_c4.selectbox("Köy / Mahalle", vil_opts, index=v_idx, key=f"dossier_ed_vil_{person_id}")
			f_village = "" if f_vil_raw == "İlçe Merkezi / Belirtilmedi" else f_vil_raw
		else:
			f_province = l_c2.text_input("İl / Eyalet", value=person["province"] or "", key=f"dossier_ed_prov_txt_{person_id}")
			f_district = l_c3.text_input("İlçe / Şehir", value=person["district"] or "", key=f"dossier_ed_dist_txt_{person_id}")
			f_village = l_c4.text_input("Köy / Mahalle", value=person["village"] or "", key=f"dossier_ed_vil_txt_{person_id}")

		save_demo = st.form_submit_button("💾 Kimlik & İletişim Bilgilerini Kaydet", type="primary", use_container_width=True)

	if save_demo:
		if not f_name.strip():
			st.error("Ad soyad boş bırakılamaz.")
		else:
			resolved_class = "Sivil / Diğer" if f_ptype != "Öğrenci" else f_class.strip()
			clean_sex = {"Kız": "K", "Erkek": "E"}.get(f_sex)
			clean_birth = f_birth.isoformat() if f_birth else None
			clean_notes = f_addr.strip() if f_ptype == "Öğrenci" else (f_role.strip() or f_addr.strip())
			with connect_db() as connection:
				connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (resolved_class,))
				cid = connection.execute("SELECT id FROM classes WHERE name=?", (resolved_class,)).fetchone()["id"]
				connection.execute(
					"""UPDATE students SET class_id=?, person_type=?, student_no=?, full_name=?, sex=?, birth_date=?,
					   guardian_name=?, guardian_phone=?, notes=?, country=?, province=?, district=?, village=? WHERE id=?""",
					(cid, f_ptype, f_sno.strip() or None, f_name.strip(), clean_sex, clean_birth,
					 f_guard.strip() or None, f_phone.strip() or None, clean_notes or None,
					 f_country or None, (f_province or "").strip() or None,
					 (f_district or "").strip() or None, (f_village or "").strip() or None, person_id)
				)
				connection.execute(
					"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, updated_at) VALUES (?, ?, ?, ?)
					   ON CONFLICT(student_id) DO UPDATE SET guardian_name=excluded.guardian_name, guardian_phone=excluded.guardian_phone, updated_at=excluded.updated_at""",
					(person_id, f_guard.strip() or None, f_phone.strip() or None, datetime.now().isoformat(timespec="minutes"))
				)
			sync_student_measurements(person_id)
			save_person_data_file(person_id)
			st.success("Kimlik ve iletişim bilgileri başarıyla güncellendi.")
			st.rerun()

	st.markdown("---")
	del_chk_col, del_btn_col = st.columns([2, 1])
	with del_chk_col:
		confirm_del = st.checkbox("Bu kişiyi ve tüm ilişkili kayıtlarını silmeyi onayla", key=f"dossier_del_conf_{person_id}")
	with del_btn_col:
		if st.button("🗑️ Kişiyi Kalıcı Olarak Sil", type="secondary", disabled=not confirm_del, key=f"dossier_del_btn_{person_id}", use_container_width=True):
			delete_student(person_id)
			set_active_person(None)
			st.session_state["target_page"] = "👥 Öğrenciler"
			st.success("Kişi ve bağlı tüm kayıtlar silindi.")
			st.rerun()


def _render_dossier_tab_printable(person: sqlite3.Row, person_id: int) -> None:
	st.markdown("#### 🖨️ Yazdırılabilir Resmi Sağlık Dosyası")
	app_cfg = get_app_settings()
	inst_top = html.escape(app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	inst_school = html.escape(app_cfg.get("school_name") or "")
	inst_nurse = html.escape(app_cfg.get("nurse_name") or "Okul Revir Sorumlusu")
	inst_princ = html.escape(app_cfg.get("principal_name") or "Okul Müdürü / İdare")

	single_body = build_student_report_section_html(person, inst_top, inst_school, inst_nurse, inst_princ, is_multi=False)
	css_header = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Öğrenci Sağlık Dosyası</title>
<style>
@media print {
	@page { size: A4; margin: 12mm 15mm; }
	.no-print { display: none !important; }
	body { font-size: 11pt; }
}
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; margin: 20px auto; max-width: 900px; color: #1e293b; line-height: 1.4; }
.report-page { page-break-after: always; padding: 10px 0; }
table { width: 100%; border-collapse: collapse; margin: 10px 0 16px 0; font-size: 11px; }
th, td { border: 1px solid #cbd5e1; padding: 5px 8px; text-align: left; }
th { background: #f1f5f9; font-weight: 600; color: #0f172a; }
h2 { color: #0f172a; border-bottom: 2px solid #0f766e; padding-bottom: 6px; margin-top: 20px; font-size: 16px; }
h3 { color: #0f766e; margin-top: 14px; margin-bottom: 6px; font-size: 13px; }
.header-box { text-align: center; margin-bottom: 15px; border-bottom: 2px double #94a3b8; padding-bottom: 10px; }
.header-box h1 { margin: 0; font-size: 15px; text-transform: uppercase; color: #0f172a; }
.header-box h2 { border: none; margin: 2px 0; font-size: 13px; color: #0f766e; }
.sign-box { display: flex; justify-content: space-between; margin-top: 35px; page-break-inside: avoid; }
.sign-item { text-align: center; width: 220px; border-top: 1px solid #64748b; padding-top: 5px; font-size: 11px; }
button.print-btn { background: #0f766e; color: white; border: none; padding: 10px 20px; font-size: 14px; font-weight: 600; border-radius: 6px; cursor: pointer; }
button.print-btn:hover { background: #115e59; }
</style>
</head>
<body>
<div class="no-print" style="margin-bottom: 20px; text-align: right;">
	<button class="print-btn" onclick="window.print()">🖨️ Sayfayı Yazdır / PDF Olarak Kaydet</button>
</div>
"""
	full_print_html = css_header + single_body + "</body>\n</html>"

	st.components.v1.html(full_print_html, height=850, scrolling=True)
	print_name = re.sub(r"[^A-Za-z0-9_-]+", "_", person["full_name"]).strip("_") or "kisi"
	rep_docx = build_student_report_docx(person)
	st.download_button(
		"📝 Word Rapor Dosyasını İndir (.docx)",
		data=rep_docx,
		file_name=f"{print_name}_saglik_dosyasi.docx",
		mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		type="primary",
		key=f"dossier_dl_print_docx_{person_id}",
		use_container_width=True
	)


def render_person_dossier_content(person_id: int, is_dialog: bool = False) -> None:
	app_cfg = get_app_settings()
	theme_name = app_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
	theme = THEME_PRESETS.get(theme_name, THEME_PRESETS["Klinik Turkuaz (Varsayılan)"])
	is_dark = is_dark_mode_active(app_cfg)
	muted_col = "#94a3b8" if is_dark else "#475569"

	person = get_person_by_id(person_id)
	if not person:
		st.info("Kişi bulunamadı.")
		return

	with connect_db() as connection:
		hp = connection.execute("SELECT * FROM health_profiles WHERE student_id=?", (person_id,)).fetchone()
		visit_stats = connection.execute("SELECT COUNT(*) cnt, MAX(visit_date) last_date FROM visits WHERE student_id=?", (person_id,)).fetchone()
		meas_cnt = connection.execute("SELECT COUNT(*) cnt FROM measurements WHERE student_id=?", (person_id,)).fetchone()["cnt"]
		vital_cnt = connection.execute("SELECT COUNT(*) cnt FROM vital_readings WHERE student_id=?", (person_id,)).fetchone()["cnt"]

	today = date.today()
	today_months = age_in_months(person["birth_date"], today) if person["birth_date"] else None
	age_str = f"{today_months // 12} yaş {today_months % 12} ay ({today_months} ay)" if today_months is not None else "Doğum tarihi yok"
	sex_str = {"K": "Kız ♀", "E": "Erkek ♂"}.get(person["sex"], "Belirtilmedi")
	blood = (hp["blood_group"] if hp and hp["blood_group"] else None) or "Belirtilmedi"
	guardian = (hp["guardian_name"] if hp and hp["guardian_name"] else person["guardian_name"]) or "—"
	phone = (hp["guardian_phone"] if hp and hp["guardian_phone"] else person["guardian_phone"]) or "—"
	loc_str = " / ".join(filter(None, [person["province"], person["district"], person["village"]]))
	addr = person["notes"] or "—"

	allergies = (hp["allergies"] if hp and hp["allergies"] else "").strip()
	chronic = (hp["chronic_conditions"] if hp and hp["chronic_conditions"] else "").strip()
	meds = (hp["medications"] if hp and hp["medications"] else "").strip()
	emergency = (hp["emergency_notes"] if hp and hp["emergency_notes"] else "").strip()

	badge_htmls = []
	if allergies:
		badge_htmls.append(f"<span style='background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;padding:3px 9px;border-radius:12px;font-size:12px;font-weight:600;'>🔴 Alerji: {html.escape(allergies)}</span>")
	if chronic:
		badge_htmls.append(f"<span style='background:#fffbeb;color:#b45309;border:1px solid #fde68a;padding:3px 9px;border-radius:12px;font-size:12px;font-weight:600;'>🟠 Kronik: {html.escape(chronic)}</span>")
	if meds:
		badge_htmls.append(f"<span style='background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0;padding:3px 9px;border-radius:12px;font-size:12px;font-weight:600;'>💊 İlaç: {html.escape(meds)}</span>")
	if emergency:
		badge_htmls.append(f"<span style='background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;padding:3px 9px;border-radius:12px;font-size:12px;font-weight:600;'>🚨 Acil: {html.escape(emergency)}</span>")

	with st.container(border=True):
		hero_l, hero_r = st.columns([2.2, 1.0])
		with hero_l:
			sno_badge = f" <span style='background:#e0f2fe;color:#0369a1;padding:2px 8px;border-radius:6px;font-size:13px;font-weight:600;'>No: {html.escape(str(person['student_no']))}</span>" if person["student_no"] else ""
			st.markdown(f"<h3 style='margin:0;padding:0;color:{theme['ink']};'>👤 {html.escape(person['full_name'])} <small style='color:{theme['teal']};font-size:16px;font-weight:600;'>· {html.escape(person['class_name'])}</small>{sno_badge}</h3>", unsafe_allow_html=True)
			st.markdown(f"🎂 **Doğum:** {person['birth_date'] or '—'} ({age_str}) &nbsp;·&nbsp; {sex_str} &nbsp;·&nbsp; 🩸 **Kan Grubu:** `{blood}`")
			st.caption(f"📞 **Veli / Yakını:** {guardian} ({phone}) &nbsp;·&nbsp; 📍 **Yerleşim:** {loc_str or addr}")
		with hero_r:
			st.markdown(f"🏥 **Toplam Revir Başvurusu:** `{visit_stats['cnt']}`")
			st.markdown(f"📏 **Ölçüm:** `{meas_cnt}` &nbsp;·&nbsp; 💓 **Vital:** `{vital_cnt}`")
			st.caption(f"🗓️ **Son Başvuru Tarihi:** {visit_stats['last_date'] or 'Yok'}")

		if badge_htmls:
			st.markdown("<div style='margin-top:6px;display:flex;gap:8px;flex-wrap:wrap;'>" + "".join(badge_htmls) + "</div>", unsafe_allow_html=True)

	d_tabs = st.tabs([
		"📋 360° Genel Özet",
		f"🩺 Revir Kayıtları ({visit_stats['cnt']})",
		f"📈 Büyüme & Persantil ({meas_cnt})",
		f"💓 Vital Bulgular ({vital_cnt})",
		"🧪 Tahlil & Laboratuvar",
		"💉 Aşı Takvimi (Sağlık Bakanlığı)",
		"📑 Sağlık Profili & Bakım Planı",
		"✏️ Kimlik & İletişim Düzenle",
		"🖨️ Yazdırılabilir Sağlık Dosyası",
	])

	with d_tabs[0]:
		_render_dossier_tab_overview(person, person_id, hp, visit_stats, theme, muted_col)
	with d_tabs[1]:
		_render_dossier_tab_visits(person, person_id)
	with d_tabs[2]:
		_render_dossier_tab_measurements(person, person_id)
	with d_tabs[3]:
		_render_dossier_tab_vitals(person, person_id)
	with d_tabs[4]:
		_render_dossier_tab_labs(person, person_id)
	with d_tabs[5]:
		_render_dossier_tab_vaccinations(person, person_id, hp)
	with d_tabs[6]:
		_render_dossier_tab_care_plans(person, person_id, hp)
	with d_tabs[7]:
		_render_dossier_tab_edit_demographics(person, person_id, hp)
	with d_tabs[8]:
		_render_dossier_tab_printable(person, person_id)

	if is_dialog:
		st.write("")
		if st.button("✖ Pencereyi Kapat", key=f"dossier_dlg_close_btn_{person_id}", type="secondary", use_container_width=True):
			st.session_state["active_dossier_modal_pid"] = None
			st.session_state["roster_table_seed"] = st.session_state.get("roster_table_seed", 0) + 1
			st.rerun()


@st.dialog("📂 Kişi Dosyası (Detay & İşlem)", width="large")
def open_person_dossier_dialog(person_id: int) -> None:
	render_person_dossier_content(person_id, is_dialog=True)


def render_person_dossier() -> None:
	app_cfg = get_app_settings()
	theme_name = app_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
	theme = THEME_PRESETS.get(theme_name, THEME_PRESETS["Klinik Turkuaz (Varsayılan)"])
	is_dark = is_dark_mode_active(app_cfg)
	muted_col = "#94a3b8" if is_dark else "#475569"

	active_id = st.session_state.get("active_person_id")
	person = get_person_by_id(active_id) if active_id else None

	top_c1, top_c2 = st.columns([1.2, 2.8])
	with top_c1:
		is_civ = person and (person["class_name"] == "Sivil / Diğer" or "Sivil" in str(person["person_type"] or ""))
		back_label = "🧑‍💼 Yetişkinler Listesine Dön" if is_civ else "👥 Öğrenci Listesine Dön"
		back_target = "🧑‍💼 Yetişkinler" if is_civ else "👥 Öğrenciler"
		if st.button(f"⬅️ {back_label}", key="dossier_btn_back", use_container_width=True):
			st.session_state["target_page"] = back_target
			st.rerun()
	with top_c2:
		chosen = select_shared_person("dossier_top", label="📂 Başka Bir Kişinin Dosyasını Aç")
		if chosen and chosen["id"] != active_id:
			set_active_person(chosen["id"])
			st.rerun()

	active_id = st.session_state.get("active_person_id")
	if not active_id:
		st.info("💡 Görüntülemek için yukarıdaki listeden veya 'Öğrenciler'/'Yetişkinler' sayfasından bir kişi seçin.")
		return

	render_person_dossier_content(active_id, is_dialog=False)


@st.dialog("✏️ Revir Başvuru Kaydını Düzenle", width="large")
def open_edit_visit_dialog(visit_id: int) -> None:
	with connect_db() as connection:
		v_row = connection.execute(
			"""SELECT v.*, s.full_name, s.student_no, s.person_type, c.name as class_name
			   FROM visits v
			   JOIN students s ON s.id = v.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   WHERE v.id = ?""",
			(visit_id,),
		).fetchone()
	if not v_row:
		st.error("Revir başvuru kaydı bulunamadı.")
		return

	stu_name = v_row["full_name"]
	stu_class = v_row["class_name"] or "Belirtilmedi"
	stu_no = f" · No: {v_row['student_no']}" if v_row["student_no"] else ""
	st.markdown(f"### 👤 {html.escape(stu_name)} <small style='color:#0f766e;'>({html.escape(stu_class)}{stu_no})</small>", unsafe_allow_html=True)
	st.caption("Revir başvuru kaydının tarih, saat, şikayet, bulgu ve yapılan işlem bilgilerini güncelleyebilirsiniz.")

	try:
		cur_v_date = datetime.strptime(v_row["visit_date"], "%Y-%m-%d").date()
	except Exception:
		cur_v_date = date.today()
	try:
		cur_v_time = datetime.strptime(v_row["visit_time"], "%H:%M").time()
	except Exception:
		cur_v_time = datetime.now().time()

	with st.form(f"form_modal_edit_visit_{visit_id}"):
		dt1, dt2, dt3 = st.columns(3)
		with dt1:
			new_v_date = st.date_input("🗓️ Başvuru Tarihi", value=cur_v_date)
		with dt2:
			new_v_time = st.time_input("⏰ Saat", value=cur_v_time)
		with dt3:
			comp_opts = ["Hastalık / rahatsızlık", "Yaralanma", "İlaç / bakım", "Kontrol", "Diğer"]
			c_idx = comp_opts.index(v_row["complaint"]) if v_row["complaint"] in comp_opts else 0
			new_complaint = st.selectbox("📌 Başvuru Nedeni", comp_opts, index=c_idx)

		new_syms = st.text_area("🩺 Şikayet / Belirtiler", value=v_row["symptoms"] or "", height=80)
		new_obs = st.text_area("🔍 Gözlem ve Ölçümler", value=v_row["observation"] or "", height=80)
		new_act = st.text_area("🛠️ Yapılan İşlem / Müdahale", value=v_row["action"] or "", height=80)

		c_ext1, c_ext2, c_ext3 = st.columns([1, 1.5, 1.5])
		with c_ext1:
			new_cont = st.checkbox("📞 Veli Arandı", value=bool(v_row["guardian_contacted"]))
		with c_ext2:
			new_ref = st.text_input("🚨 Yönlendirme / Sevk", value=v_row["referral"] or "")
		with c_ext3:
			new_foll = st.text_input("⏱️ İzlem / Takip", value=v_row["follow_up"] or "")

		btn_save, btn_del = st.columns([2, 1])
		with btn_save:
			if st.form_submit_button("💾 Değişiklikleri Kaydet", type="primary", use_container_width=True):
				with connect_db() as connection:
					connection.execute(
						"""UPDATE visits SET visit_date=?, visit_time=?, complaint=?, symptoms=?, observation=?,
						                      action=?, guardian_contacted=?, referral=?, follow_up=? WHERE id=?""",
						(new_v_date.isoformat(), new_v_time.strftime("%H:%M"), new_complaint,
						 new_syms.strip(), new_obs.strip(), new_act.strip(), int(new_cont),
						 new_ref.strip(), new_foll.strip(), visit_id),
					)
				save_person_data_file(int(v_row["student_id"]))
				st.success("Revir başvuru kaydı güncellendi.")
				st.rerun()

	with st.expander("🗑️ Bu Başvuru Kaydını Sil"):
		st.warning("Bu revir kaydını silmek istediğinize emin misiniz?")
		conf_del = st.checkbox("Evet, bu kaydı kalıcı olarak sil", key=f"modal_conf_del_visit_{visit_id}")
		if st.button("🗑️ Kaydı Sil", type="secondary", disabled=not conf_del, key=f"modal_btn_del_visit_{visit_id}"):
			delete_visit(visit_id)
			st.success("Revir kaydı silindi.")
			st.rerun()


def get_recent_operations(limit: int = 50) -> list[dict]:
	ops = []
	with connect_db() as connection:
		# 1. Visits
		visits = connection.execute(
			"""SELECT v.id, v.student_id, v.visit_date, v.visit_time, v.complaint, v.symptoms, v.observation, v.action,
			          s.full_name, s.student_no, COALESCE(c.name, '—') as class_name
			   FROM visits v
			   JOIN students s ON s.id = v.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   ORDER BY v.visit_date DESC, v.visit_time DESC, v.id DESC LIMIT ?""",
			(limit,),
		).fetchall()
		for v in visits:
			v_date = v["visit_date"] or ""
			v_time = v["visit_time"] or ""
			sort_key = f"{v_date} {v_time}"
			date_fmt = f"{v_date[8:10]}.{v_date[5:7]}.{v_date[:4]} {v_time}".strip() if len(v_date) >= 10 else v_date

			info = []
			sym = (v["symptoms"] or "").strip()
			obs = (v["observation"] or "").strip()
			act = (v["action"] or "").strip()
			comp = (v["complaint"] or "").strip()

			# Varsa şikayeti, yoksa sadece doldurulan değerler
			if sym:
				info.append(sym)
				if act:
					info.append(act)
			else:
				if obs:
					clean_obs = obs.replace("[Antropometri:", "").replace("[", "").replace("]", "").strip()
					info.append(clean_obs)
				if act:
					info.append(act)
				if not obs and not act and comp:
					info.append(comp)

			ops.append({
				"type": "visit",
				"key": f"v_{v['id']}",
				"id": int(v["id"]),
				"student_id": int(v["student_id"]),
				"sort_key": sort_key,
				"date_str": date_fmt,
				"class_name": v["class_name"],
				"full_name": v["full_name"],
				"student_no": str(v["student_no"]).strip() if v["student_no"] is not None and str(v["student_no"]).strip() and str(v["student_no"]).strip() != "None" else "",
				"details": info,
			})

		# 2. School Screenings
		screenings = connection.execute(
			"""SELECT sc.id, sc.student_id, sc.screening_date, sc.created_at, sc.height_cm, sc.weight_kg,
			          sc.head_circ_cm, sc.temperature_c, sc.pulse_bpm, sc.spo2_pct, sc.wears_glasses,
			          sc.vision_right, sc.vision_left, sc.hair_screening, sc.hearing_right, sc.hearing_left,
			          sc.health_behavior, sc.harmful_habits, sc.psychosocial_status, sc.notes,
			          s.full_name, s.student_no, COALESCE(c.name, '—') as class_name
			   FROM school_screenings sc
			   JOIN students s ON s.id = sc.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   ORDER BY sc.screening_date DESC, sc.created_at DESC, sc.id DESC LIMIT ?""",
			(limit,),
		).fetchall()
		for sc in screenings:
			s_date = sc["screening_date"] or ""
			c_at = sc["created_at"] or ""
			time_part = c_at[11:16] if len(c_at) >= 16 else ""
			sort_key = f"{s_date} {time_part}".strip()
			date_fmt = f"{s_date[8:10]}.{s_date[5:7]}.{s_date[:4]} {time_part}".strip() if len(s_date) >= 10 else s_date

			info = []
			if sc["height_cm"] is not None:
				info.append(f"{sc['height_cm']} cm")
			if sc["weight_kg"] is not None:
				info.append(f"{sc['weight_kg']} kg")
			if sc["head_circ_cm"] is not None:
				info.append(f"{sc['head_circ_cm']} cm")
			if sc["temperature_c"] is not None:
				info.append(f"{sc['temperature_c']} °C")
			if sc["pulse_bpm"] is not None:
				info.append(f"{int(sc['pulse_bpm'])} /dk")
			if sc["spo2_pct"] is not None:
				info.append(f"%{int(sc['spo2_pct'])}")
			if sc["wears_glasses"] and "Evet" in sc["wears_glasses"]:
				info.append(sc["wears_glasses"])
			if sc["hair_screening"]:
				info.append(sc["hair_screening"].strip())
			if sc["hearing_right"] and sc["hearing_right"] == sc["hearing_left"]:
				info.append(sc["hearing_right"].strip())
			else:
				if sc["hearing_right"]:
					info.append(sc["hearing_right"].strip())
				if sc["hearing_left"]:
					info.append(sc["hearing_left"].strip())
			if sc["health_behavior"]:
				info.append(sc["health_behavior"].strip())
			if sc["harmful_habits"]:
				info.append(sc["harmful_habits"].strip())
			if sc["psychosocial_status"]:
				info.append(sc["psychosocial_status"].strip())
			if sc["notes"]:
				info.append(sc["notes"].strip())

			ops.append({
				"type": "screening",
				"key": f"sc_{sc['id']}",
				"id": int(sc["id"]),
				"student_id": int(sc["student_id"]),
				"sort_key": sort_key,
				"date_str": date_fmt,
				"class_name": sc["class_name"],
				"full_name": sc["full_name"],
				"student_no": str(sc["student_no"]).strip() if sc["student_no"] is not None and str(sc["student_no"]).strip() and str(sc["student_no"]).strip() != "None" else "",
				"details": info,
			})

	ops.sort(key=lambda x: x["sort_key"], reverse=True)
	return ops[:limit]


def render_overview() -> None:
	active_ay = get_active_academic_year()
	s_date = active_ay["start_date"]
	e_date = active_ay["end_date"]
	today = date.today().isoformat()
	is_archive = e_date < today

	with connect_db() as connection:
		counts = connection.execute(
			"""SELECT (SELECT COUNT(*) FROM students) students,
					  (SELECT COUNT(*) FROM classes) classes,
					  (SELECT COUNT(*) FROM visits WHERE visit_date=?) today_visits,
					  (SELECT COUNT(*) FROM visits WHERE visit_date >= ? AND visit_date <= ?) total_visits,
					  (SELECT COUNT(*) FROM measurements WHERE measured_on >= ? AND measured_on <= ?) measurements,
					  (SELECT COUNT(*) FROM health_profiles WHERE (allergies IS NOT NULL AND allergies != '') OR (chronic_conditions IS NOT NULL AND chronic_conditions != '')) risk_profiles""",
			(today, s_date, e_date, s_date, e_date),
		).fetchone()

	st.markdown(f"## 📊 Revir Genel Durumu ve Göstergeler ({active_ay['name']})")
	metric_cols = st.columns(6)
	metric_cols[0].metric("Kayıtlı Kişi", counts["students"])
	metric_cols[1].metric("Sınıf / Grup", counts["classes"])
	metric_cols[2].metric("Bugünkü Başvuru", counts["today_visits"])
	metric_cols[3].metric(f"Revir ({active_ay['name']})", counts["total_visits"])
	metric_cols[4].metric(f"Ölçüm ({active_ay['name']})", counts["measurements"])
	metric_cols[5].metric("Alerji / Kronik", counts["risk_profiles"])

	top_op1, top_op2 = st.columns([4, 1.2])
	with top_op1:
		st.markdown("### 🕒 Son Yapılan İşlemler")
	with top_op2:
		op_limit = st.selectbox("Kayıt Limiti", [15, 30, 50, 100], index=1, key="sel_recent_ops_limit", label_visibility="collapsed")
	recent_ops = get_recent_operations(limit=op_limit)
	if not recent_ops:
		st.info("Henüz revir başvurusu veya sağlık taraması kaydı bulunmuyor.")
	else:
		for op in recent_ops:
			r_col, btn_col = st.columns([12.5, 1.1])
			with r_col:
				badges = []
				if op["date_str"]:
					badges.append(f"<span style='color:#64748b; font-size:12.5px; font-weight:600;'>{html.escape(op['date_str'])}</span>")
				if op["class_name"] and op["class_name"] != "—":
					badges.append(f"<span style='background-color:#e0f2fe; color:#0369a1; padding:2px 8px; border-radius:6px; font-weight:700; font-size:13.5px;'>{html.escape(op['class_name'])}</span>")
				badges.append(f"<span style='font-weight:700; color:#0f172a; font-size:15px;'>{html.escape(op['full_name'])}</span>")
				if op["student_no"]:
					badges.append(f"<span style='background-color:#f1f5f9; color:#475569; padding:2px 7px; border-radius:6px; font-weight:600; font-size:13px;'>{html.escape(op['student_no'])}</span>")
				for d in op["details"]:
					clean_d = html.escape(d)
					badges.append(f"<span style='color:#94a3b8;'>·</span><span style='color:#334155; font-size:14px; font-weight:500;'>{clean_d}</span>")

				content_html = " ".join(badges)
				st.markdown(
					f"<div style='background-color:#ffffff; border:1px solid #e2e8f0; border-radius:8px; padding:10px 14px; margin-bottom:6px; display:flex; flex-wrap:wrap; align-items:center; gap:8px; box-shadow:0 1px 2px rgba(0,0,0,0.03);'>{content_html}</div>",
					unsafe_allow_html=True,
				)
			with btn_col:
				if st.button("✏️", key=f"btn_edit_op_{op['key']}", help="Bu işlemi düzenle", use_container_width=True):
					if op["type"] == "visit":
						open_edit_visit_dialog(op["id"])
					elif op["type"] == "screening":
						open_screening_dialog(op["student_id"], op["class_name"])


@st.dialog("🏫 Okul Sınıfları Yönetimi (Ekle / Çıkar / Öğrenci Ata)", width="large")
def open_class_manager_dialog() -> None:
	hdr_l, hdr_r = st.columns([8.5, 1.5])
	with hdr_l:
		st.caption("Okuldaki mevcut sınıfları ekleyip çıkarabilir, hazır sınıf kademelerini tek tıkla oluşturabilir veya öğrencilerin sınıflarını topluca atayabilirsiniz. (İl, ilçe veya köy bilgileri sınıf olarak kullanılmaz.)")
	with hdr_r:
		if st.button("✖ Kapat", key="close_class_mgr_modal", use_container_width=True):
			st.rerun()

	tab_list, tab_assign = st.tabs(["🏫 Mevcut Sınıflar & Ekle / Sil", "👥 Öğrencileri Sınıflara Ata"])

	with tab_list:
		with connect_db() as connection:
			cls_df = pd.read_sql_query(
				"""SELECT c.id, c.name AS Sinif_Adi, COUNT(s.id) AS Ogrenci_Sayisi
				   FROM classes c LEFT JOIN students s ON s.class_id = c.id
				   WHERE c.name != 'Sivil / Diğer'
				   GROUP BY c.id, c.name ORDER BY c.name""",
				connection,
			)

		add_c1, add_c2 = st.columns([2.2, 1.8])
		with add_c1:
			st.markdown("##### ➕ Yeni Sınıf Ekle")
			in_c1, in_c2 = st.columns([2.5, 1.2])
			with in_c1:
				new_cls_input = st.text_input("Sınıf Adı", placeholder="Örn: 1-A, 5-B, 9-A, Anasınıfı (Virgülle birden fazla yazabilirsiniz)", key="modal_new_cls_in", label_visibility="collapsed")
			with in_c2:
				if st.button("➕ Sınıf Ekle", type="primary", key="modal_btn_add_cls", use_container_width=True):
					parts = [p.strip() for p in new_cls_input.split(",") if p.strip()]
					if parts:
						with connect_db() as connection:
							for p in parts:
								connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (p,))
						st.success(f"{len(parts)} sınıf eklendi.")
						st.rerun()

		with add_c2:
			st.markdown("##### ⚡ Hazır Şablon & Kademe Ekle")
			all_dlg_tpls = get_all_class_templates()
			dlg_tpl_opts = ["Seçiniz..."] + list(all_dlg_tpls.keys())
			preset_choice = st.selectbox(
				"Okul Türü / Şablon",
				dlg_tpl_opts,
				key="modal_cls_preset",
				label_visibility="collapsed",
			)
			if st.button("📦 Şablon Sınıfları Ekle", key="modal_btn_preset_cls", use_container_width=True):
				to_add = all_dlg_tpls.get(preset_choice, [])
				if to_add:
					added_cnt, _ = apply_class_template(to_add, clear_empty=False)
					st.success(f"'{preset_choice}' şablonundan {added_cnt} yeni sınıf eklendi.")
					st.rerun()
				else:
					st.warning("Lütfen eklemek istediğiniz şablonu seçin.")

		st.markdown("##### 📋 Kayıtlı Sınıflar Listesi")
		st.dataframe(
			cls_df.drop(columns=["id"], errors="ignore").rename(columns={"Sinif_Adi": "Sınıf", "Ogrenci_Sayisi": "Kayıtlı Öğrenci Sayısı"}),
			width="stretch",
			hide_index=True,
			height=230,
		)

		removable_classes = [r["Sinif_Adi"] for _, r in cls_df.iterrows() if r["Sinif_Adi"] != "Sınıf Belirtilmedi"]
		if removable_classes:
			st.markdown("##### 🗑️ Sınıf Çıkar / Sil veya Adını Değiştir")
			del_c1, del_c2, del_c3 = st.columns([1.6, 1.6, 1.2])
			with del_c1:
				target_cls = st.selectbox("İşlem Yapılacak Sınıf", removable_classes, key="modal_del_cls_sel")
			with del_c2:
				rename_val = st.text_input("Yeni Sınıf Adı (Değiştirmek için)", value=target_cls, key="modal_ren_cls_in")
			with del_c3:
				st.write("")
				if st.button("✏️ Adını Güncelle", key="modal_btn_ren_cls", use_container_width=True):
					if rename_val.strip() and rename_val.strip() != target_cls:
						with connect_db() as connection:
							old_r = connection.execute("SELECT id FROM classes WHERE name=?", (target_cls,)).fetchone()
							new_r = connection.execute("SELECT id FROM classes WHERE name=?", (rename_val.strip(),)).fetchone()
							if old_r:
								if new_r:
									connection.execute("UPDATE OR IGNORE students SET class_id=? WHERE class_id=?", (new_r["id"], old_r["id"]))
									connection.execute("DELETE FROM classes WHERE id=?", (old_r["id"],))
								else:
									connection.execute("UPDATE classes SET name=? WHERE id=?", (rename_val.strip(), old_r["id"]))
						st.success(f"'{target_cls}' -> '{rename_val.strip()}' olarak güncellendi.")
						st.rerun()

			if st.button(f"🗑️ '{target_cls}' Sınıfını Çıkar (İçindeki öğrenciler 'Sınıf Belirtilmedi' olarak korunur)", type="secondary", key="modal_btn_del_cls", use_container_width=True):
				with connect_db() as connection:
					connection.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sınıf Belirtilmedi')")
					unassigned_id = connection.execute("SELECT id FROM classes WHERE name='Sınıf Belirtilmedi'").fetchone()["id"]
					old_r = connection.execute("SELECT id FROM classes WHERE name=?", (target_cls,)).fetchone()
					if old_r:
						connection.execute("UPDATE OR IGNORE students SET class_id=? WHERE class_id=?", (unassigned_id, old_r["id"]))
						connection.execute("DELETE FROM classes WHERE id=?", (old_r["id"],))
				if st.session_state.get("active_class_filter") == target_cls:
					st.session_state["active_class_filter"] = "Tüm sınıflar"
				st.success(f"'{target_cls}' sınıfı kaldırıldı.")
				st.rerun()

	with tab_assign:
		st.caption("Aşağıdaki listeden bir veya birden fazla öğrenci seçip okuldaki sınıflarına tek tıkla atayabilirsiniz.")
		all_stu = student_options(person_type="Öğrenci")
		valid_classes = [c for c in get_classes() if c != "Sivil / Diğer"]
		if not valid_classes:
			valid_classes = ["Sınıf Belirtilmedi"]
		if all_stu:
			stu_label_map = {f"{s['full_name']} (Mevcut Sınıfı: {s['class_name']})": s["id"] for s in all_stu}
			picked_students = st.multiselect("Sınıfı Değiştirilecek Öğrencileri Seçin", list(stu_label_map.keys()), key="modal_bulk_stu_pick")
			dest_class = st.selectbox("Atanacak Yeni Sınıf", valid_classes, key="modal_bulk_dest_cls")
			if st.button("💾 Seçili Öğrencileri Bu Sınıfa Ata", type="primary", disabled=not picked_students, key="modal_btn_bulk_assign", use_container_width=True):
				with connect_db() as connection:
					connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (dest_class,))
					dest_cid = connection.execute("SELECT id FROM classes WHERE name=?", (dest_class,)).fetchone()["id"]
					for lbl in picked_students:
						sid = stu_label_map[lbl]
						connection.execute("UPDATE OR IGNORE students SET class_id=? WHERE id=?", (dest_cid, sid))
				st.success(f"{len(picked_students)} öğrenci '{dest_class}' sınıfına atandı.")
				st.rerun()


@st.dialog("➕ Yeni Profil ve İlk Ölçüm / Sağlık Kaydı", width="large")
def open_new_profile_dialog(default_person_type: str = "Öğrenci") -> None:
	top_l, top_r = st.columns([8.5, 1.5])
	with top_l:
		st.caption("Kişi bilgilerini, okuldaki sınıfını, yerleşim yerini (il / ilçe / köy-mahalle), varsa ilk antropometrik ölçümlerini ve sağlık geçmişini girip kaydedin. İptal etmek için sağ üstteki ✖ butonuna tıklayabilirsiniz.")
	with top_r:
		if st.button("✖ İptal", key=f"close_new_prof_modal_{default_person_type}", use_container_width=True):
			st.rerun()

	app_cfg = get_app_settings()
	def_country = app_cfg.get("default_country") or "Türkiye"
	def_prov = app_cfg.get("default_province") or "Karaman"
	def_dist = app_cfg.get("default_district") or "Merkez"
	def_vil = app_cfg.get("default_village") or ""

	existing_classes = [c for c in get_classes() if c != "Sivil / Diğer"]
	if "Sınıf Belirtilmedi" not in existing_classes:
		existing_classes = ["Sınıf Belirtilmedi", *existing_classes]

	new_nonce = st.session_state.setdefault("new_profile_nonce", 0)
	type_opts = ["Öğrenci", "Sivil / diğer kişi"]
	default_type_idx = 0 if default_person_type == "Öğrenci" else 1
	person_type = st.selectbox("Kayıt türü", type_opts, index=default_type_idx, key=f"modal_person_type_{default_person_type}_{new_nonce}")
	first, second, third = st.columns(3)
	with first:
		if person_type == "Öğrenci":
			cls_choices = [*existing_classes, "➕ Yeni Sınıf Yaz..."]
			picked_cls = st.selectbox("Sınıfı *", cls_choices, index=0, key=f"new_class_sel_{new_nonce}")
			if picked_cls == "➕ Yeni Sınıf Yaz...":
				class_name = st.text_input("Yeni Sınıf Adı *", placeholder="Örn: 5-A veya 1. Kademe", key=f"new_class_custom_{new_nonce}")
			else:
				class_name = picked_cls
			role_note = st.text_input("Açık Adres / Mahalle / Sokak (İsteğe bağlı)", placeholder="Örn. Gevher Hatun Mah. 1782 Sok. No:7", key=f"new_stu_addr_{new_nonce}")
		else:
			class_name = ""
			role_note = st.selectbox(
				"Görev / Statü",
				["Öğretmen", "İdari Personel", "Destek / Güvenlik Personeli", "Veli", "Ziyaretçi / Diğer"],
				key=f"new_role_{new_nonce}",
			)
		full_name = st.text_input("Ad soyad *", key=f"new_name_{new_nonce}")
		if person_type == "Öğrenci":
			student_no = st.text_input("Okul numarası", key=f"new_number_{new_nonce}")
		else:
			student_no = ""
			st.text_input("Kayıt grubu", value="Sivil / Diğer", disabled=True, key=f"new_group_{new_nonce}")
		sex = st.selectbox("Cinsiyet", ["Belirtilmedi", "Kız / Kadın", "Erkek"], key=f"new_sex_{new_nonce}")
	with second:
		birth = st.date_input("Doğum tarihi", value=None, min_value=date(1940, 1, 1), max_value=date.today(), key=f"new_birth_{new_nonce}")
		guardian_lbl = "Veli adı" if person_type == "Öğrenci" else "Acil durumda aranacak yakını"
		phone_lbl = "Veli telefonu" if person_type == "Öğrenci" else "İletişim telefonu"
		guardian = st.text_input(guardian_lbl, key=f"new_guardian_{new_nonce}")
		phone = st.text_input(phone_lbl, key=f"new_phone_{new_nonce}")
		blood_grp = st.selectbox("Kan grubu", BLOOD_GROUP_OPTIONS, key=f"new_blood_{new_nonce}")
	with third:
		init_height = st.number_input("İlk Boy Ölçümü (cm, isteğe bağlı)", min_value=0.0, max_value=250.0, value=None, step=0.1, format="%.1f", key=f"new_h_{new_nonce}")
		init_weight = st.number_input("İlk Kilo Ölçümü (kg, isteğe bağlı)", min_value=0.0, max_value=300.0, value=None, step=0.1, format="%.1f", key=f"new_w_{new_nonce}")
		if person_type == "Öğrenci":
			init_head_circ = st.number_input("İlk Baş Çevresi (cm, isteğe bağlı)", min_value=0.0, max_value=100.0, value=None, step=0.1, format="%.1f", key=f"new_hc_{new_nonce}")
		else:
			init_head_circ = None
		init_allergy = st.text_input("Bilinen Alerji (İsteğe bağlı)", key=f"new_allergy_{new_nonce}")
		init_chronic = st.text_input("Kronik Hastalıklar (İsteğe bağlı)", key=f"new_chronic_{new_nonce}")
		init_med_hist = st.text_input("Sağlık Geçmişi / Diğer Yaşadığı Durumlar (İsteğe bağlı)", key=f"new_medhist_{new_nonce}")

	st.markdown("##### 📍 Yerleşim Yeri (Ülke / İl / İlçe / Köy-Mahalle)")
	st.caption("Varsayılan konum Ayarlar'dan otomatik gelir; kişi farklı bir il, ilçe veya köyden geliyorsa aşağıdan değiştirebilirsiniz.")
	loc_c1, loc_c2, loc_c3, loc_c4 = st.columns(4)
	with loc_c1:
		c_idx = COUNTRY_OPTIONS.index(def_country) if def_country in COUNTRY_OPTIONS else 0
		sel_country = st.selectbox("Ülke", COUNTRY_OPTIONS, index=c_idx, key=f"new_country_{new_nonce}")
	if sel_country == "Türkiye":
		provinces = get_tr_provinces()
		p_idx = provinces.index(def_prov) if def_prov in provinces else 0
		with loc_c2:
			sel_province = st.selectbox("İl", provinces, index=p_idx, key=f"new_prov_{new_nonce}")
		districts = get_tr_districts(sel_province)
		d_idx = districts.index(def_dist) if def_dist in districts else 0
		with loc_c3:
			sel_district = st.selectbox("İlçe", districts, index=d_idx, key=f"new_dist_{new_nonce}")
		villages = get_tr_villages(sel_province, sel_district)
		vil_opts = ["İlçe Merkezi / Belirtilmedi", *villages]
		v_idx = vil_opts.index(def_vil) if def_vil in vil_opts else 0
		with loc_c4:
			sel_village_raw = st.selectbox("Köy / Mahalle / Belde", vil_opts, index=v_idx, key=f"new_vil_{new_nonce}")
			sel_village = "" if sel_village_raw == "İlçe Merkezi / Belirtilmedi" else sel_village_raw
	else:
		with loc_c2:
			sel_province = st.text_input("İl / Eyalet", value="", key=f"new_prov_txt_{new_nonce}")
		with loc_c3:
			sel_district = st.text_input("İlçe / Şehir", value="", key=f"new_dist_txt_{new_nonce}")
		with loc_c4:
			sel_village = st.text_input("Köy / Mahalle", value="", key=f"new_vil_txt_{new_nonce}")

	btn_save_col, btn_cancel_col = st.columns([3, 1])
	with btn_save_col:
		save_clicked = st.button("💾 Profili ve Bağlı Verileri Kaydet", type="primary", key=f"save_new_profile_{new_nonce}", use_container_width=True)
	with btn_cancel_col:
		if st.button("✖ Vazgeç", type="secondary", key=f"cancel_bottom_new_profile_{new_nonce}", use_container_width=True):
			st.rerun()

	if save_clicked:
		resolved_class = "Sivil / Diğer" if person_type != "Öğrenci" else class_name.strip()
		if not resolved_class or not full_name.strip():
			st.error("Ad soyad ve öğrenciler için sınıf zorunludur.")
		else:
			try:
				resolved_sex = {"Kız / Kadın": "K", "Kız": "K", "Erkek": "E"}.get(sex)
				resolved_birth = birth.isoformat() if birth else None
				resolved_note = (role_note or "").strip() or None
				with connect_db() as connection:
					connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (resolved_class,))
					class_id = connection.execute("SELECT id FROM classes WHERE name=?", (resolved_class,)).fetchone()["id"]
					cur = connection.execute(
						"""INSERT INTO students(class_id, person_type, student_no, full_name, sex, birth_date,
						   guardian_name, guardian_phone, notes, country, province, district, village)
						   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(class_id, person_type, student_no.strip() or None, full_name.strip(),
						 resolved_sex, resolved_birth,
						 guardian.strip() or None, phone.strip() or None, resolved_note,
						 sel_country or None, (sel_province or "").strip() or None,
						 (sel_district or "").strip() or None, (sel_village or "").strip() or None),
					)
					new_id = cur.lastrowid
					if blood_grp != "Belirtilmedi" or init_allergy.strip() or init_chronic.strip() or init_med_hist.strip() or guardian.strip() or phone.strip():
						connection.execute(
							"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, allergies,
							   chronic_conditions, medical_history, blood_group, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
							(new_id, guardian.strip() or None, phone.strip() or None,
							 init_allergy.strip() or None, init_chronic.strip() or None, init_med_hist.strip() or None,
							 None if blood_grp == "Belirtilmedi" else blood_grp,
							 datetime.now().isoformat(timespec="minutes")),
						)
				has_init_hw = init_height is not None and init_weight is not None and init_height > 0 and init_weight > 0
				has_init_hc = init_head_circ is not None and init_head_circ > 0
				if has_init_hw or has_init_hc:
					record_linked_measurement(
						new_id,
						date.today(),
						float(init_height) if has_init_hw else None,
						float(init_weight) if has_init_hw else None,
						resolved_birth,
						resolved_sex,
						head_circ_cm=float(init_head_circ) if has_init_hc else None,
					)
				save_person_data_file(new_id)
				set_active_person(new_id)
			except sqlite3.IntegrityError:
				st.error("Bu sınıfta aynı isimle kayıtlı bir öğrenci zaten var.")
			else:
				st.session_state["new_profile_nonce"] = new_nonce + 1
				st.session_state["profile_saved_toast"] = f"{full_name.strip()} başarıyla kaydedildi ve aktif kişi olarak seçildi."
				st.rerun()


@st.dialog("✏️ Profili Düzenle / Sil", width="large")
def open_edit_profile_dialog(edit_id: int, key_suffix: str = "modal") -> None:
	person = get_person_by_id(edit_id)
	if person is None:
		st.session_state["edit_profile_open_id"] = None
		st.session_state["edit_profile_loaded_id"] = None
		st.rerun()
		return

	hdr_l, hdr_r = st.columns([8.5, 1.5])
	with hdr_l:
		st.markdown(f"### ✏️ {person['full_name']}")
		st.caption("Kişi bilgilerini, okuldaki sınıfını, açık adresini ve yerleşim yerini güncelleyebilir, kaydı silebilir veya sağ üstteki ✖ butonuna tıklayarak iptal edebilirsiniz.")
	with hdr_r:
		if st.button("✖ İptal", key=f"close_profile_editor_{key_suffix}", use_container_width=True):
			st.session_state["edit_profile_open_id"] = None
			st.session_state["edit_profile_loaded_id"] = None
			st.rerun()

	app_cfg = get_app_settings()
	existing_classes = [c for c in get_classes() if c != "Sivil / Diğer"]
	if "Sınıf Belirtilmedi" not in existing_classes:
		existing_classes = ["Sınıf Belirtilmedi", *existing_classes]

	with connect_db() as connection:
		edit_hp = connection.execute("SELECT blood_group FROM health_profiles WHERE student_id=?", (edit_id,)).fetchone()
	if st.session_state.get("edit_profile_loaded_id") != edit_id:
		st.session_state["edit_profile_loaded_id"] = edit_id
		st.session_state["edit_person_type"] = person["person_type"]
		st.session_state["edit_class_name"] = person["class_name"] if person["person_type"] == "Öğrenci" else ""
		st.session_state["edit_full_name"] = person["full_name"]
		st.session_state["edit_student_no"] = person["student_no"] or ""
		st.session_state["edit_notes"] = person["notes"] or ""
		st.session_state["edit_sex"] = {"K": "Kız / Kadın", "E": "Erkek"}.get(person["sex"], "Belirtilmedi")
		st.session_state["edit_birth"] = date.fromisoformat(person["birth_date"]) if person["birth_date"] else None
		st.session_state["edit_guardian"] = person["guardian_name"] or ""
		st.session_state["edit_phone"] = person["guardian_phone"] or ""
		st.session_state["edit_blood_group"] = normalize_blood_group(edit_hp["blood_group"] if edit_hp else None)
		st.session_state["edit_country"] = (person["country"] if "country" in person.keys() and person["country"] else None) or app_cfg.get("default_country") or "Türkiye"
		st.session_state["edit_province"] = (person["province"] if "province" in person.keys() and person["province"] else None) or app_cfg.get("default_province") or "Karaman"
		st.session_state["edit_district"] = (person["district"] if "district" in person.keys() and person["district"] else None) or app_cfg.get("default_district") or "Merkez"
		st.session_state["edit_village"] = (person["village"] if "village" in person.keys() and person["village"] else None) or ""

	edit_type = st.selectbox("Kayıt türü", ["Öğrenci", "Sivil / diğer kişi"], key="edit_person_type")
	edit_first, edit_second = st.columns(2)
	with edit_first:
		if edit_type == "Öğrenci":
			cur_cls = st.session_state.get("edit_class_name", "Sınıf Belirtilmedi")
			cls_opts = list(existing_classes)
			if cur_cls and cur_cls not in cls_opts:
				cls_opts.insert(0, cur_cls)
			cls_opts.append("➕ Yeni Sınıf Yaz...")
			e_cls_idx = cls_opts.index(cur_cls) if cur_cls in cls_opts else 0
			picked_edit_cls = st.selectbox("Sınıfı", cls_opts, index=e_cls_idx, key=f"edit_cls_sel_{edit_id}")
			if picked_edit_cls == "➕ Yeni Sınıf Yaz...":
				edit_class = st.text_input("Yeni Sınıf Adı", value="", placeholder="Örn: 5-A veya 1. Kademe", key=f"edit_cls_custom_{edit_id}")
			else:
				edit_class = picked_edit_cls
			edit_notes = st.text_input("Açık Adres / Mahalle / Not", key="edit_notes")
		else:
			edit_class = ""
			st.text_input("Kayıt grubu", value="Sivil / Diğer", disabled=True, key="edit_group_display")
			edit_notes = st.text_input("Görev / Statü", key="edit_notes")
		edit_name = st.text_input("Ad soyad", key="edit_full_name")
		if edit_type == "Öğrenci":
			edit_number = st.text_input("Okul numarası", key="edit_student_no")
		else:
			edit_number = ""
		edit_sex = st.selectbox("Cinsiyet", ["Belirtilmedi", "Kız / Kadın", "Erkek"], key="edit_sex")
	with edit_second:
		edit_birth = st.date_input("Doğum tarihi", value=None, min_value=date(1940, 1, 1), max_value=date.today(), key="edit_birth")
		edit_guardian = st.text_input("Veli / Yakını adı", key="edit_guardian")
		edit_phone = st.text_input("İletişim telefonu", key="edit_phone")
		edit_blood = st.selectbox("Kan grubu", BLOOD_GROUP_OPTIONS, key="edit_blood_group")

	st.markdown("##### 📍 Yerleşim Yeri (Ülke / İl / İlçe / Köy-Mahalle)")
	el_c1, el_c2, el_c3, el_c4 = st.columns(4)
	cur_c = st.session_state.get("edit_country", "Türkiye")
	with el_c1:
		ec_idx = COUNTRY_OPTIONS.index(cur_c) if cur_c in COUNTRY_OPTIONS else 0
		edit_country = st.selectbox("Ülke", COUNTRY_OPTIONS, index=ec_idx, key=f"edit_country_sel_{edit_id}")
	if edit_country == "Türkiye":
		provinces = get_tr_provinces()
		cur_p = st.session_state.get("edit_province", "Karaman")
		ep_idx = provinces.index(cur_p) if cur_p in provinces else 0
		with el_c2:
			edit_province = st.selectbox("İl", provinces, index=ep_idx, key=f"edit_prov_sel_{edit_id}")
		districts = get_tr_districts(edit_province)
		cur_d = st.session_state.get("edit_district", "Merkez")
		ed_idx = districts.index(cur_d) if cur_d in districts else 0
		with el_c3:
			edit_district = st.selectbox("İlçe", districts, index=ed_idx, key=f"edit_dist_sel_{edit_id}")
		villages = get_tr_villages(edit_province, edit_district)
		evil_opts = ["İlçe Merkezi / Belirtilmedi", *villages]
		cur_v = st.session_state.get("edit_village", "")
		ev_idx = evil_opts.index(cur_v) if cur_v in evil_opts else 0
		with el_c4:
			edit_village_raw = st.selectbox("Köy / Mahalle / Belde", evil_opts, index=ev_idx, key=f"edit_vil_sel_{edit_id}")
			edit_village = "" if edit_village_raw == "İlçe Merkezi / Belirtilmedi" else edit_village_raw
	else:
		with el_c2:
			edit_province = st.text_input("İl / Eyalet", value=st.session_state.get("edit_province", ""), key=f"edit_prov_txt_{edit_id}")
		with el_c3:
			edit_district = st.text_input("İlçe / Şehir", value=st.session_state.get("edit_district", ""), key=f"edit_dist_txt_{edit_id}")
		with el_c4:
			edit_village = st.text_input("Köy / Mahalle", value=st.session_state.get("edit_village", ""), key=f"edit_vil_txt_{edit_id}")

	save_col, cancel_col = st.columns([3, 1])
	with save_col:
		if st.button("💾 Değişiklikleri Kaydet", type="primary", key=f"save_edit_profile_{edit_id}_{key_suffix}", use_container_width=True):
			resolved_class = "Sivil / Diğer" if edit_type != "Öğrenci" else edit_class.strip()
			if not resolved_class or not edit_name.strip():
				st.error("Ad soyad ve öğrenciler için sınıf zorunludur.")
			else:
				try:
					resolved_bg = None if edit_blood == "Belirtilmedi" else edit_blood
					with connect_db() as connection:
						connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (resolved_class,))
						class_id = connection.execute("SELECT id FROM classes WHERE name=?", (resolved_class,)).fetchone()["id"]
						connection.execute(
							"""UPDATE students SET class_id=?, person_type=?, student_no=?, full_name=?, sex=?, birth_date=?,
							   guardian_name=?, guardian_phone=?, notes=?, country=?, province=?, district=?, village=? WHERE id=?""",
							(class_id, edit_type, edit_number.strip() or None, edit_name.strip(),
							 {"Kız / Kadın": "K", "E": "E", "Erkek": "E"}.get(edit_sex), edit_birth.isoformat() if edit_birth else None,
							 edit_guardian.strip() or None, edit_phone.strip() or None, edit_notes.strip() or None,
							 edit_country or None, (edit_province or "").strip() or None,
							 (edit_district or "").strip() or None, (edit_village or "").strip() or None, edit_id),
						)
						connection.execute(
							"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, blood_group, updated_at)
							   VALUES (?, ?, ?, ?, ?)
							   ON CONFLICT(student_id) DO UPDATE SET
							   guardian_name=excluded.guardian_name, guardian_phone=excluded.guardian_phone,
							   blood_group=excluded.blood_group, updated_at=excluded.updated_at""",
							(edit_id, edit_guardian.strip() or None, edit_phone.strip() or None,
							 resolved_bg, datetime.now().isoformat(timespec="minutes")),
						)
					sync_student_measurements(edit_id)
				except sqlite3.IntegrityError:
					st.error("Bu sınıfta aynı isimle kayıtlı başka bir kişi var.")
				else:
					st.session_state["edit_profile_open_id"] = None
					st.session_state["edit_profile_loaded_id"] = None
					st.session_state["profile_saved_toast"] = "Profil güncellendi ve bağlı persantil/ölçüm verileri senkronize edildi."
					st.rerun()
	with cancel_col:
		if st.button("✖ Vazgeç", type="secondary", key=f"cancel_bottom_edit_profile_{edit_id}_{key_suffix}", use_container_width=True):
			st.session_state["edit_profile_open_id"] = None
			st.session_state["edit_profile_loaded_id"] = None
			st.rerun()

	st.divider()
	del_chk_col, del_btn_col = st.columns([2, 1])
	with del_chk_col:
		confirm_del = st.checkbox("Bu kişiyi ve tüm ilişkili kayıtlarını silmeyi onayla", key=f"del_conf_{edit_id}_{key_suffix}")
	with del_btn_col:
		if st.button("🗑️ Kişiyi Kalıcı Olarak Sil", type="secondary", disabled=not confirm_del, key=f"del_btn_{edit_id}_{key_suffix}", use_container_width=True):
			delete_student(edit_id)
			if st.session_state.get("active_person_id") == edit_id:
				set_active_person(None)
			st.session_state["edit_profile_open_id"] = None
			st.session_state["edit_profile_loaded_id"] = None
			st.session_state["profile_saved_toast"] = "Kişi profili ve tüm ilişkili kayıtlar silindi."
			st.rerun()


def render_students() -> None:
	top_title_col, top_sync_btn_col, top_cls_btn_col, top_btn_col = st.columns([2.5, 1.5, 1.4, 1.2])
	with top_title_col:
		st.markdown("## Öğrenci profilleri ve ortak sağlık özeti")
	with top_sync_btn_col:
		if st.button("🔄 Verileri Senkronize Et", type="secondary", key="btn_sync_all_master", use_container_width=True, help="PDF ve Excel listelerindeki şube, doğum tarihi, okul no, telefon ve yerleşim verilerini tüm dosyalarla eşitler"):
			with st.spinner("Resmi listeler ve kişisel veri dosyaları senkronize ediliyor..."):
				summary_msg = sync_master_documents_and_files()
			st.success(summary_msg)
			st.rerun()
	with top_cls_btn_col:
		if st.button("🏫 Sınıfları Yönet (Ekle/Çıkar)", type="secondary", key="btn_open_class_mgr_modal", use_container_width=True):
			open_class_manager_dialog()
	with top_btn_col:
		if st.button("➕ Yeni Kayıt", type="primary", key="btn_open_new_student_modal", use_container_width=True):
			open_new_profile_dialog("Öğrenci")
	if st.session_state.get("profile_saved_toast"):
		st.success(st.session_state.pop("profile_saved_toast"))

	classes = [c for c in get_classes() if c != "Sivil / Diğer"]
	saved_cls = st.session_state.get("active_class_filter", "Tüm sınıflar")
	cls_opts = ["Tüm sınıflar", *classes]
	cls_idx = cls_opts.index(saved_cls) if saved_cls in cls_opts else 0

	filter_col, search_col, pick_col = st.columns([0.9, 1.1, 1.8])
	with filter_col:
		selected_class = st.selectbox("Sınıf filtresi", cls_opts, index=cls_idx, key="students_page_class_filter")
		st.session_state["active_class_filter"] = selected_class
	with search_col:
		search_q = st.text_input("🔍 Hızlı Arama (Ad, Mahalle, Köy, Tel)", placeholder="Örn: Gevher Hatun, 0532, Recep...", key="students_page_search_q")

	students = get_students_enriched(selected_class, person_type="Öğrenci").reset_index(drop=True)
	if search_q.strip() and not students.empty:
		q_low = search_q.strip().lower()
		mask = (
			students["Ad Soyad"].astype(str).str.lower().str.contains(q_low, na=False)
			| students["Adres / Mahalle"].astype(str).str.lower().str.contains(q_low, na=False)
			| students["İl / İlçe / Köy"].astype(str).str.lower().str.contains(q_low, na=False)
			| students["Telefon"].astype(str).str.lower().str.contains(q_low, na=False)
			| students["Okul No"].astype(str).str.lower().str.contains(q_low, na=False)
		)
		students = students[mask].reset_index(drop=True)

	display_students = students.drop(columns=["Kayıt Türü", "Görev / Statü"], errors="ignore") if not students.empty else students
	if not display_students.empty:
		display_students = display_students.copy()
		if "📂 Dosya" not in display_students.columns:
			display_students.insert(0, "📂 Dosya", "📂 Aç")
	st.caption(f"Toplam {len(students)} öğrenci listeleniyor · Sol kutudan öğrenciyi seçebilir, **📂 Aç** sembolüne tıklayarak Kişi Dosyasını açabilirsiniz.")

	app_cfg = get_app_settings()
	tbl_h = int(app_cfg.get("default_table_height", "420") or 420)

	roster_seed = st.session_state.get("roster_table_seed", 0)
	table_event = st.dataframe(
		display_students,
		width="stretch",
		height=tbl_h,
		hide_index=True,
		on_select="rerun",
		selection_mode=["single-row", "single-cell"],
		key=f"profile_roster_table_{roster_seed}",
		column_config=get_students_column_config(),
	)
	selected_cells = getattr(table_event.selection, "cells", []) or []
	selected_rows = getattr(table_event.selection, "rows", []) or []

	dosya_clicked_row = None
	for c in selected_cells:
		c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
		c_col = c[1] if isinstance(c, (tuple, list)) else (c.get("column") if isinstance(c, dict) else None)
		if c_col in ("📂 Dosya", "Dosya", "📂") and c_row is not None and 0 <= int(c_row) < len(students):
			dosya_clicked_row = int(c_row)
			break

	if dosya_clicked_row is not None:
		row_person_id = int(students.iloc[dosya_clicked_row]["id"])
		set_active_person(row_person_id)
		cell_token = (roster_seed, dosya_clicked_row)
		if st.session_state.get("last_opened_cell_token") != cell_token:
			st.session_state["last_opened_cell_token"] = cell_token
			open_person_dossier_dialog(row_person_id)
		else:
			st.session_state["last_opened_cell_token"] = None
			st.session_state["roster_table_seed"] = roster_seed + 1
			st.rerun()
	elif selected_rows and 0 <= selected_rows[0] < len(students):
		row_person_id = int(students.iloc[selected_rows[0]]["id"])
		if row_person_id != st.session_state.get("active_person_id"):
			set_active_person(row_person_id)
			st.rerun()
	elif selected_cells:
		c = selected_cells[0]
		c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
		if c_row is not None and 0 <= int(c_row) < len(students):
			row_person_id = int(students.iloc[int(c_row)]["id"])
			if row_person_id != st.session_state.get("active_person_id"):
				set_active_person(row_person_id)
				st.rerun()

	active_id = st.session_state.get("active_person_id")
	person_map = {
		f"📂 {r['Sınıf']} · {r['Ad Soyad']}" + (f" (No: {r['Okul No']})" if r["Okul No"] != "—" else ""): int(r["id"])
		for _, r in students.iterrows()
	} if not students.empty else {}
	id_to_label = {v: k for k, v in person_map.items()}
	with pick_col:
		if person_map:
			labels_list = list(person_map.keys())
			cur_lbl = id_to_label.get(active_id)
			def_idx = labels_list.index(cur_lbl) if cur_lbl in labels_list else None
			picked_lbl = st.selectbox(
				"Aktif Öğrenci (Tüm Pencerelerde Ortak Seçili)",
				labels_list,
				index=def_idx,
				placeholder="Tablodan tıklayın veya buradan öğrenci arayın...",
				key="students_page_active_picker",
			)
			if picked_lbl and person_map[picked_lbl] != active_id:
				set_active_person(person_map[picked_lbl])
				active_id = person_map[picked_lbl]

	selected_person_id = st.session_state.get("active_person_id")
	active_person_obj = get_person_by_id(selected_person_id)
	if active_person_obj is not None and active_person_obj["class_name"] != "Sivil / Diğer":
		render_shared_person_summary(active_person_obj, "students_page")
		act_c1, act_c2, act_c3 = st.columns([2.2, 1.2, 1.4])
		with act_c1:
			if st.button(f"📂 {active_person_obj['full_name']} Dosyasını Aç (Açılır Pencere) ↗️", type="primary", key="open_selected_profile_dossier", use_container_width=True):
				open_person_dossier_dialog(active_person_obj["id"])
		with act_c2:
			if st.button("✏️ Hızlı Düzenle", type="secondary", key="edit_selected_profile", use_container_width=True):
				st.session_state["edit_profile_loaded_id"] = None
				open_edit_profile_dialog(active_person_obj["id"], "stu")
		with act_c3:
			p_fpath = get_person_data_file_path(active_person_obj["id"], active_person_obj["full_name"])
			with st.expander("📁 JSON Dosyası"):
				render_person_data_file_editor(active_person_obj["id"], section_key="stu_page")


def render_civilians() -> None:
	top_title_col, top_btn_col = st.columns([4, 1.2])
	with top_title_col:
		st.markdown("## 🧑‍💼 Yetişkinler (Personel, Öğretmen ve Ziyaretçi Kayıtları)")
	with top_btn_col:
		if st.button("➕ Yeni Kayıt", type="primary", key="btn_open_new_civilian_modal", use_container_width=True):
			open_new_profile_dialog("Sivil / diğer kişi")

	if st.session_state.get("profile_saved_toast"):
		st.success(st.session_state.pop("profile_saved_toast"))

	civilians = get_students_enriched("Sivil / Diğer", person_type="Sivil").reset_index(drop=True)
	display_civilians = civilians.drop(columns=["Kayıt Türü", "Sınıf", "Okul No", "Baş Çevresi (cm)", "Adres / Mahalle"], errors="ignore") if not civilians.empty else civilians

	role_col, civ_search_col, pick_col = st.columns([0.9, 1.1, 1.8])
	with role_col:
		roles = sorted({str(r) for r in civilians["Görev / Statü"].dropna().unique() if str(r) != "—"}) if not civilians.empty and "Görev / Statü" in civilians.columns else []
		selected_role = st.selectbox("Görev / Statü filtresi", ["Tüm görevler", *roles], key="civilians_role_filter")
	with civ_search_col:
		civ_search_q = st.text_input("🔍 Hızlı Arama (Ad, Görev, Tel)", placeholder="İsim veya telefon ara...", key="civilians_page_search_q")
	if selected_role != "Tüm görevler" and not civilians.empty:
		civilians = civilians[civilians["Görev / Statü"] == selected_role].reset_index(drop=True)
	if civ_search_q.strip() and not civilians.empty:
		cq = civ_search_q.strip().lower()
		cmask = (
			civilians["Ad Soyad"].astype(str).str.lower().str.contains(cq, na=False)
			| civilians["Görev / Statü"].astype(str).str.lower().str.contains(cq, na=False)
			| civilians["Telefon"].astype(str).str.lower().str.contains(cq, na=False)
		)
		civilians = civilians[cmask].reset_index(drop=True)
	display_civilians = civilians.drop(columns=["Kayıt Türü", "Sınıf", "Okul No", "Baş Çevresi (cm)", "Adres / Mahalle"], errors="ignore") if not civilians.empty else civilians
	if not display_civilians.empty:
		display_civilians = display_civilians.copy()
		if "📂 Dosya" not in display_civilians.columns:
			display_civilians.insert(0, "📂 Dosya", "📂 Aç")

	st.caption(f"Toplam {len(civilians)} yetişkin / personel listeleniyor · Sol kutudan kişiyi seçebilir, **📂 Aç** sembolüne tıklayarak Kişi Dosyasını açabilirsiniz.")

	if civilians.empty:
		st.info("Henüz kayıtlı yetişkin kişi veya personel bulunmuyor. Sağ üstteki '➕ Yeni Kayıt' butonuna tıklayarak ekleyebilirsiniz.")
	else:
		civ_roster_seed = st.session_state.get("civ_roster_table_seed", 0)
		civ_table_event = st.dataframe(
			display_civilians,
			width="stretch",
			hide_index=True,
			on_select="rerun",
			selection_mode=["single-row", "single-cell"],
			key=f"civilian_roster_table_{civ_roster_seed}",
			column_config=get_students_column_config(),
		)
		selected_civ_cells = getattr(civ_table_event.selection, "cells", []) or []
		selected_civ_rows = getattr(civ_table_event.selection, "rows", []) or []

		dosya_clicked_civ_row = None
		for c in selected_civ_cells:
			c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
			c_col = c[1] if isinstance(c, (tuple, list)) else (c.get("column") if isinstance(c, dict) else None)
			if c_col in ("📂 Dosya", "Dosya", "📂") and c_row is not None and 0 <= int(c_row) < len(civilians):
				dosya_clicked_civ_row = int(c_row)
				break

		if dosya_clicked_civ_row is not None:
			row_civ_id = int(civilians.iloc[dosya_clicked_civ_row]["id"])
			set_active_person(row_civ_id)
			cell_token = (civ_roster_seed, dosya_clicked_civ_row)
			if st.session_state.get("last_opened_civ_cell_token") != cell_token:
				st.session_state["last_opened_civ_cell_token"] = cell_token
				open_person_dossier_dialog(row_civ_id)
			else:
				st.session_state["last_opened_civ_cell_token"] = None
				st.session_state["civ_roster_table_seed"] = civ_roster_seed + 1
				st.rerun()
		elif selected_civ_rows and 0 <= selected_civ_rows[0] < len(civilians):
			row_civ_id = int(civilians.iloc[selected_civ_rows[0]]["id"])
			if row_civ_id != st.session_state.get("active_person_id"):
				set_active_person(row_civ_id)
				st.rerun()
		elif selected_civ_cells:
			c = selected_civ_cells[0]
			c_row = c[0] if isinstance(c, (tuple, list)) else (c.get("row") if isinstance(c, dict) else None)
			if c_row is not None and 0 <= int(c_row) < len(civilians):
				row_civ_id = int(civilians.iloc[int(c_row)]["id"])
				if row_civ_id != st.session_state.get("active_person_id"):
					set_active_person(row_civ_id)
					st.rerun()

	active_id = st.session_state.get("active_person_id")
	civ_map = {
		f"📂 {r['Ad Soyad']}" + (f" · {r['Görev / Statü']}" if r.get("Görev / Statü") and r["Görev / Statü"] != "—" else ""): int(r["id"])
		for _, r in civilians.iterrows()
	} if not civilians.empty else {}
	id_to_civ_label = {v: k for k, v in civ_map.items()}
	with pick_col:
		if civ_map:
			civ_labels = list(civ_map.keys())
			cur_civ_lbl = id_to_civ_label.get(active_id)
			def_civ_idx = civ_labels.index(cur_civ_lbl) if cur_civ_lbl in civ_labels else None
			picked_civ_lbl = st.selectbox(
				"Aktif Sivil / Personel (Tüm Pencerelerde Ortak Seçili)",
				civ_labels,
				index=def_civ_idx,
				placeholder="Tablodan tıklayın veya buradan kişi arayın...",
				key="civilians_page_active_picker",
			)
			if picked_civ_lbl and civ_map[picked_civ_lbl] != active_id:
				set_active_person(civ_map[picked_civ_lbl])
				active_id = civ_map[picked_civ_lbl]

	selected_person_id = st.session_state.get("active_person_id")
	active_person_obj = get_person_by_id(selected_person_id)
	if active_person_obj is not None and (active_person_obj["class_name"] == "Sivil / Diğer" or "Sivil" in (active_person_obj["person_type"] or "")):
		render_shared_person_summary(active_person_obj, "civilians_page")
		act_c1, act_c2, act_c3 = st.columns([2.2, 1.2, 1.4])
		with act_c1:
			if st.button(f"📂 {active_person_obj['full_name']} Dosyasını Aç (Açılır Pencere) ↗️", type="primary", key="open_selected_civ_dossier", use_container_width=True):
				open_person_dossier_dialog(active_person_obj["id"])
		with act_c2:
			if st.button("✏️ Hızlı Düzenle", type="secondary", key="edit_selected_civ_profile", use_container_width=True):
				st.session_state["edit_profile_loaded_id"] = None
				open_edit_profile_dialog(active_person_obj["id"], "civ")
		with act_c3:
			c_fpath = get_person_data_file_path(active_person_obj["id"], active_person_obj["full_name"])
			with st.expander("📁 JSON Dosyası"):
				render_person_data_file_editor(active_person_obj["id"], section_key="civ_page")

	st.divider()
	st.markdown("### 🩺 Sivil / Personel Revir Başvurusu ve Sağlık İzlemi")
	render_civilian_visit_tab(show_new_button=False)


def render_health_profile_and_plan() -> None:
	st.markdown("## Sağlık profili ve bakım planı")
	person = select_shared_person("hp", label="Sağlık Profili ve Bakım Planı Görüntülenecek Kişi")
	if not person:
		return
	render_shared_person_summary(person, "hp")
	person_id = person["id"]
	with connect_db() as connection:
		profile = connection.execute("SELECT * FROM health_profiles WHERE student_id=?", (person_id,)).fetchone()
		form_files = connection.execute("SELECT id, uploaded_at, file_name, content FROM health_forms WHERE student_id=? ORDER BY uploaded_at DESC", (person_id,)).fetchall()
		plans = connection.execute("SELECT * FROM care_plans WHERE student_id=? ORDER BY created_at DESC", (person_id,)).fetchall()
	profile_fields = ["guardian_name", "guardian_phone", "allergies", "chronic_conditions", "medical_history", "medications", "blood_group", "emergency_notes"]
	profile_fallbacks = {
		"guardian_name": person["guardian_name"] or "",
		"guardian_phone": person["guardian_phone"] or "",
	}
	def profile_default(field: str) -> str:
		if field == "blood_group":
			return normalize_blood_group(profile["blood_group"] if profile else None)
		if profile and field in profile.keys() and profile[field]:
			return profile[field]
		return profile_fallbacks.get(field, "")

	profile_labels = {
		"guardian_name": "Veli / acil durumda aranacak kişi",
		"guardian_phone": "İletişim telefonu",
		"allergies": "Bilinen alerjiler",
		"chronic_conditions": "Kronik hastalıklar",
		"medical_history": "Sağlık geçmişi (kronik veya değil geçirdiği hastalıklar, ameliyatlar ve yaşadığı diğer durumlar)",
		"medications": "Bildirilen düzenli ilaçlar",
		"blood_group": "Kan grubu",
		"emergency_notes": "Acil durum ve özel notlar",
	}
	profile_placeholders = {
		"chronic_conditions": "Örn: Astım, Tip 1 Diyabet, Epilepsi, Çölyak vb. sürekli takip gerektiren kronik hastalıklar",
		"medical_history": "Örn: Geçirdiği ameliyatlar, kırık/travmalar, ateşli havale, hastanede yatış öyküsü, dönemsel veya kronik olmayan diğer sağlık durumları",
	}
	profile_tab, plan_tab = st.tabs(["Sağlık formu ve profil", "Bakım planı taslağı"])
	with profile_tab:
		st.caption(f"{person['person_type']} profili · {person['full_name']}")
		uploaded = st.file_uploader("Sağlık tanılama formu PDF", type=["pdf"], key=f"health_form_upload_{person_id}")
		extracted_text = ""
		parsed = {}
		if uploaded is not None:
			content = uploaded.getvalue()
			if len(content) > 20 * 1024 * 1024:
				st.error("En fazla 20 MB PDF yükleyin.")
				uploaded = None
			else:
				signature = sha256(content).hexdigest()
				signature_key = f"health_form_signature_{person_id}"
				text_key = f"health_form_text_{person_id}"
				if st.session_state.get(signature_key) != signature:
					try:
						extracted_text = extract_pdf_text(content)
						parsed = extract_health_form_fields(extracted_text)
						st.session_state[text_key] = extracted_text
						st.session_state[signature_key] = signature
						for field in profile_fields:
							state_key = f"health_profile_{person_id}_{field}"
							existing = profile_default(field)
							val = parsed.get(field, existing)
							st.session_state[state_key] = normalize_blood_group(val) if field == "blood_group" else val
					except Exception as error:
						st.error(f"PDF okunamadı: {error}")
				else:
					extracted_text = st.session_state.get(text_key, "")
					parsed = extract_health_form_fields(extracted_text)
				if not extracted_text.strip():
					st.warning("PDF'de seçilebilir metin bulunamadı. Taranmış PDF otomatik okunamaz; aşağıdaki alanları elle doldurun.")
				elif not parsed:
					st.info("Form alanları otomatik eşleşmedi. PDF metnini kontrol edip bilgileri aşağıya elle aktarın.")
				else:
					st.caption(f"PDF'den {len(parsed)} alan taslak olarak çıkarıldı; kaydetmeden önce formun kendisiyle doğrulayın.")
					with st.expander("PDF'den çıkarılan metin"):
						st.text(extracted_text[:12000])
		else:
			for field in profile_fields:
				state_key = f"health_profile_{person_id}_{field}"
				if state_key not in st.session_state:
					st.session_state[state_key] = profile_default(field)
				elif field == "blood_group" and st.session_state[state_key] not in BLOOD_GROUP_OPTIONS:
					st.session_state[state_key] = normalize_blood_group(st.session_state[state_key])
		with st.form(f"health_profile_form_{person_id}"):
			id_c1, id_c2, id_c3 = st.columns(3)
			cur_sex_opt = {"K": "Kız", "E": "Erkek"}.get(person["sex"], "Belirtilmedi")
			hp_sex = id_c1.selectbox("Cinsiyet (Ortak)", ["Belirtilmedi", "Kız", "Erkek"], index=["Belirtilmedi", "Kız", "Erkek"].index(cur_sex_opt))
			cur_b_val = date.fromisoformat(person["birth_date"]) if person["birth_date"] else None
			hp_birth = id_c2.date_input("Doğum Tarihi (Ortak)", value=cur_b_val, min_value=date(1960, 1, 1), max_value=date.today())
			values = {}
			values["blood_group"] = id_c3.selectbox(profile_labels["blood_group"], BLOOD_GROUP_OPTIONS, key=f"health_profile_{person_id}_blood_group")
			for field in profile_fields:
				if field == "blood_group":
					continue
				if field in {"chronic_conditions", "medical_history", "medications", "emergency_notes"}:
					values[field] = st.text_area(
						profile_labels[field],
						placeholder=profile_placeholders.get(field, ""),
						key=f"health_profile_{person_id}_{field}",
					)
				else:
					values[field] = st.text_input(profile_labels[field], key=f"health_profile_{person_id}_{field}")
			app_cfg = get_app_settings()
			def_nurse_val = app_cfg.get("nurse_name", "Okul Revir Sorumlusu")
			reviewed_by = st.text_input("Bilgileri doğrulayan görevli", value=def_nurse_val)
			confirmed = st.checkbox("PDF'den çıkarılan bilgileri kaynak formla karşılaştırıp doğruladım.", value=False)
			save_profile = st.form_submit_button("Sağlık profilini kaydet", type="primary")
		if save_profile:
			if uploaded is not None and not confirmed:
				st.error("PDF bilgilerini kaydetmeden önce kaynak formla karşılaştırıp onaylayın.")
			else:
				resolved_fields = []
				for field in profile_fields:
					val = values[field].strip() if values.get(field) else ""
					if field == "blood_group" and val == "Belirtilmedi":
						val = ""
					resolved_fields.append(val or None)
				with connect_db() as connection:
					connection.execute(
						"""INSERT INTO health_profiles(student_id, guardian_name, guardian_phone, allergies, chronic_conditions,
						   medical_history, medications, blood_group, emergency_notes, updated_at, reviewed_by)
						   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
						   ON CONFLICT(student_id) DO UPDATE SET guardian_name=excluded.guardian_name,
						   guardian_phone=excluded.guardian_phone, allergies=excluded.allergies,
						   chronic_conditions=excluded.chronic_conditions, medical_history=excluded.medical_history,
						   medications=excluded.medications, blood_group=excluded.blood_group,
						   emergency_notes=excluded.emergency_notes, updated_at=excluded.updated_at,
						   reviewed_by=excluded.reviewed_by""",
						(person_id, *resolved_fields,
						 datetime.now().isoformat(timespec="minutes"), reviewed_by.strip() or None),
					)
					connection.execute(
						"UPDATE students SET guardian_name=?, guardian_phone=?, sex=?, birth_date=? WHERE id=?",
						(values["guardian_name"].strip() or None, values["guardian_phone"].strip() or None,
						 {"Kız": "K", "Erkek": "E"}.get(hp_sex), hp_birth.isoformat() if hp_birth else None, person_id),
					)
					if uploaded is not None:
						connection.execute(
							"INSERT INTO health_forms(student_id, uploaded_at, file_name, content, extracted_text, review_status) VALUES (?, ?, ?, ?, ?, ?)",
							(person_id, datetime.now().isoformat(timespec="minutes"), Path(uploaded.name).name,
							 uploaded.getvalue(), extracted_text or None, "Alanlar kullanıcı tarafından doğrulandı"),
						)
				sync_student_measurements(person_id)
				st.success("Doğrulanmış sağlık profili kaydedildi ve tüm sekmelerle senkronize edildi.")
				st.rerun()
		if form_files:
			st.markdown("### Saklanan kaynak formlar")
			for form_file in form_files:
				st.download_button(f"{form_file['uploaded_at']} · {form_file['file_name']}", bytes(form_file["content"]),
							   file_name=form_file["file_name"], mime="application/pdf", key=f"health_form_download_{form_file['id']}")
		st.caption("Metinli veya taranmış PDF'den Türkçe/İngilizce OCR ile alan taslağı çıkarılır; bilgileri kaynak formla doğrulayın. Bu bilgisayarda saklanan sağlık verileri hassastır.")
	with plan_tab:
		st.info("Otomatik taslak; sağlık profili, boy-kilo-persantil ölçümleri, vital bulgular, tahliller ve revir başvurularındaki olguları bir araya getirir. Hemşirelik tanısı, hedef ve girişimler yetkili sağlık personeli tarafından değerlendirilmeli ve onaylanmalıdır.")
		render_nanda_care_plan_editor(person_id, prefix="main_hp")


def sync_template_to_text(current_text: str, old_picked: list[str], new_picked: list[str]) -> str:
	parts = [p.strip() for p in current_text.split(",") if p.strip()]
	for r in old_picked:
		if r not in new_picked and r in parts:
			parts.remove(r)
	for a in new_picked:
		if a not in old_picked and a not in parts:
			parts.append(a)
	return ", ".join(parts)


def render_student_visit_tab() -> None:
	student = select_shared_person("visit_stu", person_type="Öğrenci", label="Öğrenci adı")
	if not student:
		st.caption("Seçilen öğrenciye ait alerji uyarıları, boy-kilo persantil durumu, vital bulgular ve sağlık geçmişi bu alanda görüntülenir.")
		return

	render_shared_person_summary(student, section_key="visit_stu")
	if not student["birth_date"] or not student["sex"]:
		warn_c1, warn_c2 = st.columns([3.2, 1])
		with warn_c1:
			st.info("ℹ️ Bu öğrencinin doğum tarihi veya cinsiyeti henüz girilmemiş. Boy-kilo persantil hesabı için sağdaki butondan hızlıca ekleyebilirsiniz.")
		with warn_c2:
			if st.button("✏️ Profili Düzenle", key=f"visit_quick_edit_{student['id']}", use_container_width=True):
				st.session_state["edit_profile_loaded_id"] = None
				open_edit_profile_dialog(student["id"], "v_stu")

	app_cfg = get_app_settings()
	raw_s_sym = app_cfg.get("custom_student_symptoms", DEFAULT_APP_SETTINGS["custom_student_symptoms"])
	raw_s_act = app_cfg.get("custom_student_actions", DEFAULT_APP_SETTINGS["custom_student_actions"])
	quick_symptom_options = [s.strip() for s in raw_s_sym.split(",") if s.strip()]
	quick_action_options = [a.strip() for a in raw_s_act.split(",") if a.strip()]

	stu_id = student["id"]
	s_txt_key = f"vis_sym_{stu_id}"
	s_tpl_key = f"vis_tpl_sym_{stu_id}"
	s_prev_key = f"vis_prev_sym_{stu_id}"

	a_txt_key = f"vis_act_{stu_id}"
	a_tpl_key = f"vis_tpl_act_{stu_id}"
	a_prev_key = f"vis_prev_act_{stu_id}"

	if s_txt_key not in st.session_state:
		st.session_state[s_txt_key] = ""
	if s_prev_key not in st.session_state:
		st.session_state[s_prev_key] = []

	if a_txt_key not in st.session_state:
		st.session_state[a_txt_key] = ""
	if a_prev_key not in st.session_state:
		st.session_state[a_prev_key] = []

	def on_stu_sym_change():
		new_p = st.session_state.get(s_tpl_key, [])
		old_p = st.session_state.get(s_prev_key, [])
		cur_val = st.session_state.get(s_txt_key, "")
		st.session_state[s_txt_key] = sync_template_to_text(cur_val, old_p, new_p)
		st.session_state[s_prev_key] = list(new_p)

	def on_stu_act_change():
		new_p = st.session_state.get(a_tpl_key, [])
		old_p = st.session_state.get(a_prev_key, [])
		cur_val = st.session_state.get(a_txt_key, "")
		st.session_state[a_txt_key] = sync_template_to_text(cur_val, old_p, new_p)
		st.session_state[a_prev_key] = list(new_p)

	col_form, col_history = st.columns([1.15, 0.85], gap="large")
	with col_form:
		with st.container(border=True):
			st.markdown(f"### 🩺 {student['full_name']} <small>· {student['class_name']}</small>", unsafe_allow_html=True)
			dt_col1, dt_col2, dt_col3 = st.columns(3)
			visit_date = dt_col1.date_input("Başvuru tarihi", value=date.today(), key=f"v_date_{stu_id}")
			visit_time = dt_col2.time_input("Saat", value=datetime.now().replace(second=0, microsecond=0).time(), key=f"v_time_{stu_id}")
			complaint = dt_col3.selectbox("Başvuru nedeni", ["Hastalık / rahatsızlık", "Yaralanma", "İlaç / bakım", "Kontrol", "Diğer"], key=f"v_comp_{stu_id}")

			st.markdown("###### 🌡️📏 Başvuru Sırasında Yapılan Ölçümler")
			st.caption("Buraya girilen ateş, saturasyon (SpO2), tansiyon, nabız, şeker, boy-kilo ve baş çevresi değerleri ilgili takip pencerelerine ve persantil hesaplamalarına otomatik kaydedilir.")
			v_c1, v_c2, v_c3, v_c4 = st.columns(4)
			v_temp = v_c1.number_input("Ateş (°C)", min_value=30.0, max_value=45.0, value=None, step=0.1, format="%.1f", key=f"v_temp_{stu_id}")
			v_spo2 = v_c2.number_input("Saturasyon - SpO2 (%)", min_value=30.0, max_value=100.0, value=None, step=1.0, format="%.0f", key=f"v_spo2_{stu_id}")
			v_pulse = v_c3.number_input("Nabız (/dk)", min_value=20.0, max_value=240.0, value=None, step=1.0, format="%.0f", key=f"v_pulse_{stu_id}")
			v_glu = v_c4.number_input("Kan Şekeri (mg/dL)", min_value=1.0, max_value=1500.0, value=None, step=1.0, format="%.0f", key=f"v_glu_{stu_id}")
			m_c1, m_c2, m_c3, m_c4, m_c5 = st.columns(5)
			v_sys = m_c1.number_input("Büyük Tansiyon", min_value=40.0, max_value=260.0, value=None, step=1.0, format="%.0f", key=f"v_sys_{stu_id}")
			v_dia = m_c2.number_input("Küçük Tansiyon", min_value=20.0, max_value=180.0, value=None, step=1.0, format="%.0f", key=f"v_dia_{stu_id}")
			v_height = m_c3.number_input("Boy (cm)", min_value=30.0, max_value=250.0, value=None, step=0.1, format="%.1f", key=f"v_height_{stu_id}")
			v_weight = m_c4.number_input("Kilo (kg)", min_value=1.0, max_value=300.0, value=None, step=0.1, format="%.1f", key=f"v_weight_{stu_id}")
			v_head_circ = m_c5.number_input("Baş Çevresi (cm)", min_value=20.0, max_value=100.0, value=None, step=0.1, format="%.1f", key=f"v_head_circ_{stu_id}")

			st.markdown("---")
			st.markdown("##### 🤒 Şikayet / Belirtiler *")
			st.caption("Öğrencinin şikayetini ve belirtilerini el ile yazabilir veya kutucuğun altındaki şablondan seçip ekleyebilirsiniz:")
			symptoms = st.text_area(
				"Şikayet / Belirtiler",
				key=s_txt_key,
				placeholder="Öğrencinin ifadesini ve şikayetini el ile yazabilir veya hemen aşağıdaki şablondan seçebilirsiniz...",
				label_visibility="collapsed"
			)
			st.multiselect(
				"📋 Sık Şikayetler Şablonu (Seçtikleriniz yukarıdaki şikayet kutusuna eklenir)",
				quick_symptom_options,
				key=s_tpl_key,
				on_change=on_stu_sym_change,
				placeholder="Sık karşılaşılan şikayetleri buradan seçebilirsiniz..."
			)

			st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
			observation = st.text_area(
				"👁️ Gözlem ve Ek Klinik Notlar",
				placeholder="Gözlenen klinik durum (Yukarıdaki ölçümler otomatik eklenecektir)",
				key=f"vis_obs_{stu_id}"
			)

			st.markdown("##### 🩹 Yapılan İşlem / Müdahale *")
			st.caption("Uygulanan işlem ve müdahaleyi el ile yazabilir veya kutucuğun altındaki şablondan seçip ekleyebilirsiniz:")
			action = st.text_area(
				"Yapılan İşlemler / Verilen Destek",
				key=a_txt_key,
				placeholder="Uygulanan tedavi, pansuman, ilk yardım veya verilen desteği el ile yazabilir veya hemen aşağıdaki şablondan seçebilirsiniz...",
				label_visibility="collapsed"
			)
			st.multiselect(
				"📋 Sık Yapılan İşlemler Şablonu (Seçtikleriniz yukarıdaki işlem kutusuna eklenir)",
				quick_action_options,
				key=a_tpl_key,
				on_change=on_stu_act_change,
				placeholder="Sık uygulanan işlemleri buradan seçebilirsiniz..."
			)

			st.markdown("---")
			g_c1, g_c2 = st.columns([1, 2])
			with g_c1:
				guardian_contacted = st.checkbox("📞 Veliye bilgi verildi", key=f"vis_gc_{stu_id}")
			with g_c2:
				referral = st.text_input("➡️ Yönlendirme", placeholder="Örn: Veliye teslim edildi / sağlık kuruluşuna sevk", key=f"vis_ref_{stu_id}")
			follow_up = st.text_area("📝 İzlem / Ek Not", key=f"vis_fu_{stu_id}", placeholder="İzlem veya takip planı notu...")

			save = st.button("💾 Revir Kaydını ve Ölçümleri Kaydet", type="primary", use_container_width=True, key=f"btn_save_visit_{stu_id}")

		if save:
			if (v_sys is None) != (v_dia is None):
				st.error("Tansiyon kaydı için büyük ve küçük tansiyonu birlikte girin.")
			elif (v_height is None) != (v_weight is None):
				st.error("Boy-kilo persantil kaydı için boy ve kiloyu birlikte girin.")
			else:
				auto_notes = []
				if any(x is not None for x in (v_temp, v_spo2, v_pulse, v_sys, v_dia, v_glu)):
					measured_at = datetime.combine(visit_date, visit_time).isoformat(timespec="minutes")
					vital_summary = record_linked_vital(
						student["id"], measured_at, v_sys, v_dia, v_pulse, v_temp, v_glu,
						glucose_context="Revir Başvurusu", notes=f"Revir başvurusu ({complaint})", spo2=v_spo2
					)
					if vital_summary:
						auto_notes.append(f"[Vital: {vital_summary}]")
				if (v_height is not None and v_weight is not None and v_height > 0 and v_weight > 0) or (v_head_circ is not None and v_head_circ > 0):
					meas_summary = record_linked_measurement(
						student["id"], visit_date, v_height, v_weight, student["birth_date"], student["sex"], head_circ_cm=v_head_circ
					)
					if meas_summary:
						auto_notes.append(f"[Antropometri: {meas_summary}]")
				final_obs = " · ".join(filter(None, [observation.strip(), *auto_notes]))
				with connect_db() as connection:
					connection.execute(
						"""INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms, observation,
						   action, guardian_contacted, referral, follow_up) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(student["id"], visit_date.isoformat(), visit_time.strftime("%H:%M"), complaint, symptoms.strip(),
						 final_obs, action.strip(), int(guardian_contacted), referral.strip(), follow_up.strip()),
					)
				save_person_data_file(student["id"])

				# Clear input fields for next entry
				for k_clear in [
					s_txt_key, s_tpl_key, s_prev_key, a_txt_key, a_tpl_key, a_prev_key,
					f"vis_obs_{stu_id}", f"vis_ref_{stu_id}", f"vis_fu_{stu_id}",
					f"v_temp_{stu_id}", f"v_spo2_{stu_id}", f"v_pulse_{stu_id}", f"v_glu_{stu_id}",
					f"v_sys_{stu_id}", f"v_dia_{stu_id}", f"v_height_{stu_id}", f"v_weight_{stu_id}", f"v_head_circ_{stu_id}"
				]:
					if k_clear in st.session_state:
						del st.session_state[k_clear]

				st.success("Revir kaydı ve bağlantılı ölçümler başarıyla kaydedildi.")
				st.rerun()

	with col_history:
		st.markdown("### Revir başvuru geçmişi")
		with connect_db() as connection:
			history = pd.read_sql_query(
				"""SELECT id, visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru,
						  symptoms AS Belirtiler, observation AS Gözlem_ve_Ölçümler, action AS İşlem, referral AS Yönlendirme,
						  guardian_contacted AS Veli_Arandı, follow_up AS İzlem
				   FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC LIMIT 15""",
				connection, params=(student["id"],),
			)
		if history.empty:
			st.caption("Bu öğrenci için henüz revir kaydı bulunmuyor.")
		else:
			display_history = history.drop(columns=["id", "follow_up"], errors="ignore")
			st.dataframe(display_history, width="stretch", hide_index=True)

			with st.expander("🛠️ Seçili Kaydı Düzenle veya Sil"):
				visit_options_map = {
					f"{r['Tarih']} {r['Saat']} · {r['Başvuru']}": int(r["id"])
					for _, r in history.iterrows()
				}
				chosen_v_label = st.selectbox("Ziyaret Kaydı Seçin", list(visit_options_map), key=f"sel_v_hist_{student['id']}")
				chosen_v_id = visit_options_map[chosen_v_label]
				selected_v_row = history[history["id"] == chosen_v_id].iloc[0]

				complaint_options = ["Hastalık / rahatsızlık", "Yaralanma", "İlaç / bakım", "Kontrol", "Diğer"]
				comp_idx = complaint_options.index(selected_v_row["Başvuru"]) if selected_v_row["Başvuru"] in complaint_options else 0
				edit_v_complaint = st.selectbox("Başvuru nedeni", complaint_options, index=comp_idx, key=f"ev_comp_{chosen_v_id}")
				edit_v_syms = st.text_area("Belirtiler", value=selected_v_row["Belirtiler"] or "", key=f"ev_sym_{chosen_v_id}")
				edit_v_obs = st.text_area("Gözlem ve Ölçümler", value=selected_v_row["Gözlem_ve_Ölçümler"] or "", key=f"ev_obs_{chosen_v_id}")
				edit_v_act = st.text_area("İşlem", value=selected_v_row["İşlem"] or "", key=f"ev_act_{chosen_v_id}")
				edit_v_cont = st.checkbox("Veli arandı", value=bool(selected_v_row["Veli_Arandı"]), key=f"ev_cont_{chosen_v_id}")
				edit_v_ref = st.text_input("Yönlendirme", value=selected_v_row["Yönlendirme"] or "", key=f"ev_ref_{chosen_v_id}")
				edit_v_foll = st.text_area("İzlem / Not", value=selected_v_row["İzlem"] or "", key=f"ev_foll_{chosen_v_id}")

				btn_c1, btn_c2 = st.columns(2)
				with btn_c1:
					if st.button("Kaydı Güncelle", type="primary", key=f"btn_upd_v_{chosen_v_id}"):
						update_visit(chosen_v_id, edit_v_complaint, edit_v_syms.strip(), edit_v_obs.strip(),
									 edit_v_act.strip(), edit_v_cont, edit_v_ref.strip(), edit_v_foll.strip())
						st.success("Ziyaret kaydı güncellendi.")
						st.rerun()
				with btn_c2:
					conf_del_v = st.checkbox("Bu kaydı silmeyi onayla", key=f"conf_del_v_{chosen_v_id}")
					if st.button("Kaydı Sil", type="secondary", disabled=not conf_del_v, key=f"btn_del_v_{chosen_v_id}"):
						delete_visit(chosen_v_id)
						st.success("Ziyaret kaydı silindi.")
						st.rerun()


def render_civilian_visit_tab(show_new_button: bool = True) -> None:
	if show_new_button:
		hdr_c, btn_c = st.columns([4, 1.2])
		with hdr_c:
			st.caption("Kayıtlı sivil / personel seçin veya sağdaki butondan yeni sivil profili açın.")
		with btn_c:
			if st.button("➕ Yeni Sivil Kayıt", type="primary", key="btn_open_civ_modal_from_visit", use_container_width=True):
				open_new_profile_dialog("Sivil / diğer kişi")

	civ_person = select_shared_person("visit_civ", person_type="Sivil", label="Sivil / Personel Seçin")
	if not civ_person:
		st.caption("Revir kaydı oluşturmak için listeden sivil / personel seçin veya sağ üstteki '➕ Yeni Kayıt' butonundan yeni kayıt açın.")
		return

	if show_new_button:
		render_shared_person_summary(civ_person, section_key="visit_civ")

	app_cfg = get_app_settings()
	raw_c_sym = app_cfg.get("custom_civilian_symptoms", DEFAULT_APP_SETTINGS["custom_civilian_symptoms"])
	raw_c_act = app_cfg.get("custom_civilian_actions", DEFAULT_APP_SETTINGS["custom_civilian_actions"])
	civ_symptom_options = [s.strip() for s in raw_c_sym.split(",") if s.strip()]
	civ_action_options = [a.strip() for a in raw_c_act.split(",") if a.strip()]

	civ_id = civ_person["id"]
	cs_txt_key = f"civ_sym_{civ_id}"
	cs_tpl_key = f"civ_tpl_sym_{civ_id}"
	cs_prev_key = f"civ_prev_sym_{civ_id}"

	ca_txt_key = f"civ_act_{civ_id}"
	ca_tpl_key = f"civ_tpl_act_{civ_id}"
	ca_prev_key = f"civ_prev_act_{civ_id}"

	if cs_txt_key not in st.session_state:
		st.session_state[cs_txt_key] = ""
	if cs_prev_key not in st.session_state:
		st.session_state[cs_prev_key] = []

	if ca_txt_key not in st.session_state:
		st.session_state[ca_txt_key] = ""
	if ca_prev_key not in st.session_state:
		st.session_state[ca_prev_key] = []

	def on_civ_sym_change():
		new_p = st.session_state.get(cs_tpl_key, [])
		old_p = st.session_state.get(cs_prev_key, [])
		cur_val = st.session_state.get(cs_txt_key, "")
		st.session_state[cs_txt_key] = sync_template_to_text(cur_val, old_p, new_p)
		st.session_state[cs_prev_key] = list(new_p)

	def on_civ_act_change():
		new_p = st.session_state.get(ca_tpl_key, [])
		old_p = st.session_state.get(ca_prev_key, [])
		cur_val = st.session_state.get(ca_txt_key, "")
		st.session_state[ca_txt_key] = sync_template_to_text(cur_val, old_p, new_p)
		st.session_state[ca_prev_key] = list(new_p)

	col_civ_form, col_civ_hist = st.columns([1.15, 0.85], gap="large")
	with col_civ_form:
		with st.container(border=True):
			st.markdown(f"### 🩺 {civ_person['full_name']} <small>· {civ_person['notes'] or 'Sivil / Diğer'}</small>", unsafe_allow_html=True)
			cdt1, cdt2, cdt3 = st.columns(3)
			civ_visit_date = cdt1.date_input("Başvuru tarihi", value=date.today(), key=f"cv_date_{civ_id}")
			civ_visit_time = cdt2.time_input("Saat", value=datetime.now().replace(second=0, microsecond=0).time(), key=f"cv_time_{civ_id}")
			civ_complaint = cdt3.selectbox("Başvuru nedeni", ["Hastalık / rahatsızlık", "Yaralanma / iş kazası", "Tansiyon / şeker kontrolü", "İlaç / pansuman", "Diğer"], key=f"cv_comp_{civ_id}")

			st.markdown("###### 🌡️📏 Başvuru Sırasında Yapılan Ölçümler")
			st.caption("Buraya girilen ateş, saturasyon (SpO2), tansiyon, nabız, şeker ve boy-kilo değerleri ilgili takip pencerelerine otomatik kaydedilir.")
			cv_c1, cv_c2, cv_c3, cv_c4 = st.columns(4)
			cv_temp = cv_c1.number_input("Ateş (°C)", min_value=30.0, max_value=45.0, value=None, step=0.1, format="%.1f", key=f"cv_temp_{civ_id}")
			cv_spo2 = cv_c2.number_input("Saturasyon - SpO2 (%)", min_value=30.0, max_value=100.0, value=None, step=1.0, format="%.0f", key=f"cv_spo2_{civ_id}")
			cv_pulse = cv_c3.number_input("Nabız (/dk)", min_value=20.0, max_value=240.0, value=None, step=1.0, format="%.0f", key=f"cv_pulse_{civ_id}")
			cv_glu = cv_c4.number_input("Kan Şekeri (mg/dL)", min_value=1.0, max_value=1500.0, value=None, step=1.0, format="%.0f", key=f"cv_glu_{civ_id}")
			cm_c1, cm_c2, cm_c3, cm_c4 = st.columns(4)
			cv_sys = cm_c1.number_input("Büyük Tansiyon", min_value=40.0, max_value=260.0, value=None, step=1.0, format="%.0f", key=f"cv_sys_{civ_id}")
			cv_dia = cm_c2.number_input("Küçük Tansiyon", min_value=20.0, max_value=180.0, value=None, step=1.0, format="%.0f", key=f"cv_dia_{civ_id}")
			cv_height = cm_c3.number_input("Boy (cm)", min_value=30.0, max_value=250.0, value=None, step=0.1, format="%.1f", key=f"cv_height_{civ_id}")
			cv_weight = cm_c4.number_input("Kilo (kg)", min_value=1.0, max_value=300.0, value=None, step=0.1, format="%.1f", key=f"cv_weight_{civ_id}")

			st.markdown("---")
			st.markdown("##### 🤒 Şikayet / Beyan *")
			st.caption("Kişinin şikayetini ve belirtilerini el ile yazabilir veya kutucuğun altındaki şablondan seçip ekleyebilirsiniz:")
			civ_symptoms = st.text_area(
				"Şikayet / Beyan",
				key=cs_txt_key,
				placeholder="Kişinin şikayetini ve belirtilerini el ile yazabilir veya hemen aşağıdaki şablondan seçebilirsiniz...",
				label_visibility="collapsed"
			)
			st.multiselect(
				"📋 Sık Şikayetler Şablonu (Seçtikleriniz yukarıdaki şikayet kutusuna eklenir)",
				civ_symptom_options,
				key=cs_tpl_key,
				on_change=on_civ_sym_change,
				placeholder="Sık karşılaşılan şikayetleri buradan seçebilirsiniz..."
			)

			st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
			civ_observation = st.text_area(
				"👁️ Gözlem ve Ek Klinik Notlar",
				placeholder="Gözlenen durum (Yukarıdaki ölçümler otomatik eklenecektir)",
				key=f"civ_obs_{civ_id}"
			)

			st.markdown("##### 🩹 Yapılan İşlem / Müdahale *")
			st.caption("Uygulanan işlem ve müdahaleyi el ile yazabilir veya kutucuğun altındaki şablondan seçip ekleyebilirsiniz:")
			civ_action = st.text_area(
				"Yapılan İşlemler / Verilen Destek",
				key=ca_txt_key,
				placeholder="Uygulanan tedavi, pansuman, ilk yardım veya verilen desteği el ile yazabilir veya hemen aşağıdaki şablondan seçebilirsiniz...",
				label_visibility="collapsed"
			)
			st.multiselect(
				"📋 Sık Yapılan İşlemler Şablonu (Seçtikleriniz yukarıdaki işlem kutusuna eklenir)",
				civ_action_options,
				key=ca_tpl_key,
				on_change=on_civ_act_change,
				placeholder="Sık uygulanan işlemleri buradan seçebilirsiniz..."
			)

			st.markdown("---")
			cg_c1, cg_c2 = st.columns([1, 2])
			with cg_c1:
				civ_guardian_contacted = st.checkbox("📞 Yakınına / ailesine bilgi verildi", key=f"civ_gc_{civ_id}")
			with cg_c2:
				civ_referral = st.text_input("➡️ Yönlendirme", placeholder="Örn: 112 Acil sevk / evine istirahate yollandı", key=f"civ_ref_{civ_id}")
			civ_follow_up = st.text_area("📝 İzlem / Ek Not", key=f"civ_fu_{civ_id}", placeholder="İzlem veya takip planı notu...")

			civ_save = st.button("💾 Sivil Revir Kaydını ve Ölçümleri Kaydet", type="primary", use_container_width=True, key=f"btn_save_civ_visit_{civ_id}")

		if civ_save:
			if (cv_sys is None) != (cv_dia is None):
				st.error("Tansiyon kaydı için büyük ve küçük tansiyonu birlikte girin.")
			elif (cv_height is None) != (cv_weight is None):
				st.error("Boy ve kilo ölçümünü birlikte girin.")
			else:
				c_auto_notes = []
				if any(x is not None for x in (cv_temp, cv_spo2, cv_pulse, cv_sys, cv_dia, cv_glu)):
					c_measured_at = datetime.combine(civ_visit_date, civ_visit_time).isoformat(timespec="minutes")
					c_vital_sum = record_linked_vital(
						civ_person["id"], c_measured_at, cv_sys, cv_dia, cv_pulse, cv_temp, cv_glu,
						glucose_context="Revir Başvurusu", notes=f"Sivil başvuru ({civ_complaint})", spo2=cv_spo2
					)
					if c_vital_sum:
						c_auto_notes.append(f"[Vital: {c_vital_sum}]")
				if cv_height is not None and cv_weight is not None and cv_height > 0 and cv_weight > 0:
					c_meas_sum = record_linked_measurement(
						civ_person["id"], civ_visit_date, cv_height, cv_weight, civ_person["birth_date"], civ_person["sex"]
					)
					if c_meas_sum:
						c_auto_notes.append(f"[Antropometri: {c_meas_sum}]")
				c_final_obs = " · ".join(filter(None, [civ_observation.strip(), *c_auto_notes]))
				with connect_db() as connection:
					connection.execute(
						"""INSERT INTO visits(student_id, visit_date, visit_time, complaint, symptoms, observation,
						   action, guardian_contacted, referral, follow_up) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
						(civ_person["id"], civ_visit_date.isoformat(), civ_visit_time.strftime("%H:%M"), civ_complaint, civ_symptoms.strip(),
						 c_final_obs, civ_action.strip(), int(civ_guardian_contacted), civ_referral.strip(), civ_follow_up.strip()),
					)
				save_person_data_file(civ_person["id"])

				# Clear input fields for next entry
				for k_clear in [
					cs_txt_key, cs_tpl_key, cs_prev_key, ca_txt_key, ca_tpl_key, ca_prev_key,
					f"civ_obs_{civ_id}", f"civ_ref_{civ_id}", f"civ_fu_{civ_id}",
					f"cv_temp_{civ_id}", f"cv_spo2_{civ_id}", f"cv_pulse_{civ_id}", f"cv_glu_{civ_id}",
					f"cv_sys_{civ_id}", f"cv_dia_{civ_id}", f"cv_height_{civ_id}", f"cv_weight_{civ_id}"
				]:
					if k_clear in st.session_state:
						del st.session_state[k_clear]

				st.success("Sivil revir kaydı ve bağlantılı ölçümler başarıyla kaydedildi.")
				st.rerun()

	with col_civ_hist:
		st.markdown("### Revir başvuru geçmişi")
		with connect_db() as connection:
			civ_history = pd.read_sql_query(
				"""SELECT id, visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru,
						  symptoms AS Belirtiler, observation AS Gözlem_ve_Ölçümler, action AS İşlem, referral AS Yönlendirme,
						  guardian_contacted AS Yakını_Arandı, follow_up AS İzlem
				   FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC LIMIT 15""",
				connection, params=(civ_person["id"],),
			)
		if civ_history.empty:
			st.caption("Bu kişi için henüz revir kaydı bulunmuyor.")
		else:
			display_civ_history = civ_history.drop(columns=["id", "follow_up"], errors="ignore")
			st.dataframe(display_civ_history, width="stretch", hide_index=True)

			with st.expander("🛠️ Seçili Kaydı Düzenle veya Sil"):
				civ_options_map = {
					f"{r['Tarih']} {r['Saat']} · {r['Başvuru']}": int(r["id"])
					for _, r in civ_history.iterrows()
				}
				chosen_cv_label = st.selectbox("Ziyaret Kaydı Seçin", list(civ_options_map), key=f"sel_cv_hist_{civ_person['id']}")
				chosen_cv_id = civ_options_map[chosen_cv_label]
				selected_cv_row = civ_history[civ_history["id"] == chosen_cv_id].iloc[0]

				c_complaint_opts = ["Hastalık / rahatsızlık", "Yaralanma / iş kazası", "Tansiyon / şeker kontrolü", "İlaç / pansuman", "Diğer"]
				cc_idx = c_complaint_opts.index(selected_cv_row["Başvuru"]) if selected_cv_row["Başvuru"] in c_complaint_opts else 0
				e_cv_comp = st.selectbox("Başvuru nedeni", c_complaint_opts, index=cc_idx, key=f"ecv_comp_{chosen_cv_id}")
				e_cv_sym = st.text_area("Belirtiler", value=selected_cv_row["Belirtiler"] or "", key=f"ecv_sym_{chosen_cv_id}")
				e_cv_obs = st.text_area("Gözlem ve Ölçümler", value=selected_cv_row["Gözlem_ve_Ölçümler"] or "", key=f"ecv_obs_{chosen_cv_id}")
				e_cv_act = st.text_area("İşlem", value=selected_cv_row["İşlem"] or "", key=f"ecv_act_{chosen_cv_id}")
				e_cv_cont = st.checkbox("Yakınına bilgi verildi", value=bool(selected_cv_row["Yakını_Arandı"]), key=f"ecv_cont_{chosen_cv_id}")
				e_cv_ref = st.text_input("Yönlendirme", value=selected_cv_row["Yönlendirme"] or "", key=f"ecv_ref_{chosen_cv_id}")
				e_cv_foll = st.text_area("İzlem / Not", value=selected_cv_row["İzlem"] or "", key=f"ecv_foll_{chosen_cv_id}")

				c_b1, c_b2 = st.columns(2)
				with c_b1:
					if st.button("Kaydı Güncelle", type="primary", key=f"btn_upd_cv_{chosen_cv_id}"):
						update_visit(chosen_cv_id, e_cv_comp, e_cv_sym.strip(), e_cv_obs.strip(),
									 e_cv_act.strip(), e_cv_cont, e_cv_ref.strip(), e_cv_foll.strip())
						st.success("Ziyaret kaydı güncellendi.")
						st.rerun()
				with c_b2:
					conf_del_cv = st.checkbox("Bu kaydı silmeyi onayla", key=f"conf_del_cv_{chosen_cv_id}")
					if st.button("Kaydı Sil", type="secondary", disabled=not conf_del_cv, key=f"btn_del_cv_{chosen_cv_id}"):
						delete_visit(chosen_cv_id)
						st.success("Ziyaret kaydı silindi.")
						st.rerun()


def render_visit() -> None:
	st.markdown("## Öğrenci revir kaydı")
	st.caption("Sivil / personel / ziyaretçi kayıtları ve başvuruları sol menüdeki '🧑‍💼 Yetişkinler' bölümüne taşınmıştır.")
	render_student_visit_tab()




def create_daily_medication_docx(target_date: date, doses: list[dict]) -> io.BytesIO:
	import docx
	from docx.shared import Inches, Pt, RGBColor
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.enum.table import WD_TABLE_ALIGNMENT
	from docx.oxml import parse_xml
	from docx.oxml.ns import nsdecls

	doc = docx.Document()
	for s in doc.sections:
		s.top_margin = Inches(0.6)
		s.bottom_margin = Inches(0.6)
		s.left_margin = Inches(0.6)
		s.right_margin = Inches(0.6)

	app_cfg = get_app_settings()
	inst_top = app_cfg.get("institution_top_header", "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	school_name = app_cfg.get("school_name", "Okul Reviri")
	nurse_name = app_cfg.get("nurse_name", "Okul Revir Sorumlusu")

	p_top = doc.add_paragraph()
	p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
	r_top = p_top.add_run(f"{inst_top}\n{school_name.upper()}\nREVİR BİRİMİ GÜNLÜK ÖĞRENCİ İLAÇ TAKİP VE UYGULAMA ÇİZELGESİ\n")
	r_top.bold = True
	r_top.font.size = Pt(11)
	r_top.font.name = "Arial"

	turkish_day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
	turkish_month_names = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
	date_str = f"{target_date.day} {turkish_month_names[target_date.month]} {target_date.year}, {turkish_day_names[target_date.weekday()]}"

	p_meta = doc.add_paragraph()
	r_d = p_meta.add_run(f"Takip Tarihi: {date_str}       |       Revir Sorumlusu: {nurse_name}")
	r_d.bold = True
	r_d.font.size = Pt(9.5)
	r_d.font.name = "Arial"

	headers = ["Saat", "Öğrenci Adı - No", "Sınıf", "İlaç Adı & Formu", "Doz", "Uygulama Yolu", "Aç/Tok", "Durum", "Uygulayan / İmza"]
	col_widths = [Inches(0.6), Inches(1.5), Inches(0.7), Inches(1.5), Inches(0.8), Inches(0.9), Inches(0.7), Inches(0.9), Inches(0.9)]

	table = doc.add_table(rows=1, cols=len(headers))
	table.alignment = WD_TABLE_ALIGNMENT.CENTER
	table.autofit = False

	hdr_row = table.rows[0]
	for idx, h_text in enumerate(headers):
		cell = hdr_row.cells[idx]
		cell.width = col_widths[idx]
		cell.text = h_text
		p = cell.paragraphs[0]
		p.alignment = WD_ALIGN_PARAGRAPH.CENTER
		for run in p.runs:
			run.bold = True
			run.font.size = Pt(8.5)
			run.font.color.rgb = RGBColor(255, 255, 255)
		shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="006666"/>')
		cell._tc.get_or_add_tcPr().append(shading_elm)

	sorted_doses = sorted(doses, key=lambda x: str(x.get("scheduled_time", "00:00")))

	if not sorted_doses:
		r = table.add_row()
		r.cells[0].merge(r.cells[len(headers) - 1])
		r.cells[0].text = "Bu tarih için planlanmış ilaç uygulaması bulunmamaktadır."
		r.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
	else:
		for d in sorted_doses:
			row = table.add_row()
			sno = f" ({d['student_no']})" if d.get('student_no') else ""
			status_text = d.get('log_status', 'Bekliyor')
			if d.get('is_logged') and d.get('given_time'):
				status_text += f" ({d['given_time']})"

			vals = [
				str(d.get("scheduled_time", "-")),
				f"{d.get('full_name', '-')}{sno}",
				str(d.get("class_name", "-")),
				str(d.get("medication_name", "-")),
				f"{d.get('dose_amount', '-')} {d.get('dose_unit', '')}",
				str(d.get("route", "-")),
				str(d.get("meal_relation", "-")),
				status_text,
				str(d.get("administered_by", ""))
			]
			for idx, v in enumerate(vals):
				cell = row.cells[idx]
				cell.width = col_widths[idx]
				cell.text = v
				p = cell.paragraphs[0]
				p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx in [0, 2, 4, 6, 7] else WD_ALIGN_PARAGRAPH.LEFT
				for run in p.runs:
					run.font.size = Pt(8)

	buf = io.BytesIO()
	doc.save(buf)
	buf.seek(0)
	return buf


DEFAULT_DOSE_TIMES = {
	1: [time(12, 0)],
	2: [time(9, 30), time(13, 30)],
	3: [time(9, 0), time(12, 30), time(15, 30)],
	4: [time(9, 0), time(11, 30), time(14, 0), time(16, 30)],
	5: [time(8, 30), time(10, 30), time(12, 30), time(14, 30), time(16, 30)],
	6: [time(8, 30), time(10, 0), time(11, 30), time(13, 0), time(14, 30), time(16, 0)],
}

@st.dialog("➕ Yeni İlaç ve Tedavi Planı Ekle / Düzenle", width="large")
def open_add_medication_dialog(preselected_student_id: int | None = None, edit_medication_id: int | None = None) -> None:
	vis_fields = get_medication_visible_fields()

	# Retrieve existing medications to allow switching or editing
	with connect_db() as conn:
		all_meds = conn.execute(
			"""SELECT m.id, m.student_id, m.medication_name, m.dosage_pattern, m.status, s.full_name, c.name as class_name
			   FROM student_medications m
			   JOIN students s ON s.id = m.student_id
			   LEFT JOIN classes c ON c.id = s.class_id
			   ORDER BY m.status ASC, s.full_name ASC"""
		).fetchall()

	target_edit_id = edit_medication_id
	if target_edit_id is None and "med_dialog_target_id" in st.session_state:
		target_edit_id = st.session_state.get("med_dialog_target_id")

	# Mode selection
	is_edit = False
	active_edit_id = None
	if all_meds:
		mode_c1, mode_c2 = st.columns([1.6, 2.4])
		with mode_c1:
			action_mode = st.radio(
				"İşlem Türü",
				["➕ Yeni İlaç Ekle", "✏️ Kayıtlı İlacı Düzenle"],
				index=1 if (target_edit_id is not None) else 0,
				horizontal=True,
				key="med_dlg_mode_radio"
			)
		if action_mode == "✏️ Kayıtlı İlacı Düzenle":
			is_edit = True
			med_options = []
			med_id_map = {}
			init_idx = 0
			for idx, row in enumerate(all_meds):
				st_tag = "" if row["status"] == "Aktif" else f" [{row['status']}]"
				lbl = f"{row['full_name']} ({row['class_name'] or 'Sınıfsız'}) — {row['medication_name']} ({row['dosage_pattern'] or '1x1'}){st_tag}"
				med_options.append(lbl)
				med_id_map[lbl] = row["id"]
				if target_edit_id and row["id"] == target_edit_id:
					init_idx = idx

			with mode_c2:
				sel_lbl = st.selectbox(
					"📋 Düzenlenecek Kayıtlı İlaç",
					med_options,
					index=init_idx,
					key="med_dlg_pick_med"
				)
				active_edit_id = med_id_map.get(sel_lbl)

	med = None
	if is_edit and active_edit_id:
		with connect_db() as conn:
			med = conn.execute(
				"""SELECT m.*, s.full_name, s.student_no, c.name as class_name, hp.allergies, hp.chronic_conditions, hp.emergency_notes
				   FROM student_medications m
				   JOIN students s ON s.id = m.student_id
				   LEFT JOIN classes c ON c.id = s.class_id
				   LEFT JOIN health_profiles hp ON hp.student_id = s.id
				   WHERE m.id = ?""",
				(active_edit_id,)
			).fetchone()

	# Key prefix for widgets to prevent state leaks
	k = f"dlg_m_{active_edit_id}_" if is_edit else "dlg_new_"

	selected_stu_id = None
	med_status = "Aktif"

	if is_edit and med:
		selected_stu_id = int(med["student_id"])
		st.markdown("##### 👤 Öğrenci Bilgisi & İlaç Durumu")
		col_stu_info, col_stu_st = st.columns([2.5, 1.2])
		with col_stu_info:
			st.markdown(f"🎓 **{med['full_name']}** (No: {med['student_no'] or '-'}) — Sınıf: **{med['class_name'] or 'Sınıfsız'}**")
		with col_stu_st:
			status_opts = ["Aktif", "Tamamlandı", "İptal Edildi"]
			cur_st_idx = status_opts.index(med["status"]) if med["status"] in status_opts else 0
			med_status = st.selectbox("📌 Durum", status_opts, index=cur_st_idx, key=f"{k}status")

		crit_alerts = []
		if med["allergies"] and str(med["allergies"]).strip():
			crit_alerts.append(f"🚨 **Alerjiler:** {str(med['allergies']).strip()}")
		if med["chronic_conditions"] and str(med["chronic_conditions"]).strip():
			crit_alerts.append(f"🏥 **Kronik Durum:** {str(med['chronic_conditions']).strip()}")
		if med.get("emergency_notes") and str(med["emergency_notes"]).strip():
			crit_alerts.append(f"ℹ️ **Acil Notu:** {str(med['emergency_notes']).strip()}")
		if crit_alerts:
			st.warning(" | ".join(crit_alerts))
	else:
		classes = get_classes()
		all_students_df = get_students()
		if all_students_df.empty:
			st.warning("Sistemde kayıtlı öğrenci bulunamadı. Lütfen önce öğrenci ekleyin.")
			return

		st.markdown("##### 👤 Öğrenci Seçimi")
		col_c, col_s = st.columns([1, 2])

		with col_c:
			selected_class = st.selectbox("🏫 Sınıf Filtresi", ["Tüm Sınıflar"] + classes, index=0, key=f"{k}sel_class")

		if selected_class and selected_class != "Tüm Sınıflar":
			filtered_students = all_students_df[all_students_df["sinif"] == selected_class]
		else:
			filtered_students = all_students_df

		if filtered_students.empty:
			st.info("Bu sınıfta öğrenci bulunmamaktadır.")
			return

		stu_options = []
		stu_id_map = {}
		init_stu_idx = 0
		for i, (_, s_row) in enumerate(filtered_students.iterrows()):
			s_no = f" (No: {s_row['okul_no']})" if s_row.get("okul_no") else ""
			label = f"{s_row['ad_soyad']}{s_no} — {s_row['sinif']}"
			stu_options.append(label)
			stu_id_map[label] = int(s_row["id"])
			if preselected_student_id and int(s_row["id"]) == preselected_student_id:
				init_stu_idx = i

		with col_s:
			selected_stu_label = st.selectbox("🎓 Öğrenci Adı *", stu_options, index=init_stu_idx, key=f"{k}sel_stu")
			selected_stu_id = stu_id_map.get(selected_stu_label)

		if selected_stu_id:
			with connect_db() as conn:
				hp_row = conn.execute("SELECT allergies, chronic_conditions, emergency_notes FROM health_profiles WHERE student_id=?", (selected_stu_id,)).fetchone()
			if hp_row:
				crit_alerts = []
				if hp_row["allergies"] and hp_row["allergies"].strip():
					crit_alerts.append(f"🚨 **Alerjiler:** {hp_row['allergies'].strip()}")
				if hp_row["chronic_conditions"] and hp_row["chronic_conditions"].strip():
					crit_alerts.append(f"🏥 **Kronik Durum:** {hp_row['chronic_conditions'].strip()}")
				if hp_row["emergency_notes"] and hp_row["emergency_notes"].strip():
					crit_alerts.append(f"ℹ️ **Acil Notu:** {hp_row['emergency_notes'].strip()}")
				if crit_alerts:
					st.warning(" | ".join(crit_alerts))

	st.markdown("---")
	st.markdown("##### 💊 İlaç ve Tedavi Bilgileri")

	r1_c1, r1_c2 = st.columns([2, 1.2])
	with r1_c1:
		init_med_name = med["medication_name"] if is_edit and med else ""
		med_name = st.text_input("İlaç Adı ve Formu *", value=init_med_name, placeholder="Örn: Augmentin-BID 400 mg Süspansiyon, Ventolin İnhaler, Calpol 120mg", key=f"{k}name")

	med_route = "Ağızdan (Oral)"
	with r1_c2:
		if "med_route" in vis_fields:
			route_opts = ["Ağızdan (Oral)", "Solunum (İnhalasyon / Hava)", "Göz Damlası", "Kulak Damlası", "Burun Spreyi", "Cilt / Dermal (Merhem/Krem)", "Enjeksiyon (Deri Altı/İnsülin)", "Rektal", "Diğer"]
			r_idx = 0
			if is_edit and med and med.get("route"):
				if med["route"] in route_opts:
					r_idx = route_opts.index(med["route"])
				else:
					route_opts.insert(0, med["route"])
					r_idx = 0
			med_route = st.selectbox("💉 Uygulama Yolu", route_opts, index=r_idx, key=f"{k}route")

	r2_c1, r2_c2, r2_c3 = st.columns([1.2, 1, 1.2])
	dosage_pattern = "2x1"
	with r2_c1:
		if "dosage_pattern" in vis_fields:
			pat_opts = ["1x1 (Günde 1 kez)", "2x1 (Günde 2 kez)", "3x1 (Günde 3 kez)", "3x2 (Günde 3 kez 2'şer doz)", "4x1 (Günde 4 kez)", "Gerektiğinde (S.O.S / PRN)", "Diğer"]
			p_idx = 1
			custom_pat_init = ""
			if is_edit and med and med.get("dosage_pattern"):
				dp = str(med["dosage_pattern"])
				if "Gerektiğinde" in dp or "PRN" in dp or "S.O.S" in dp:
					p_idx = 5
				elif "1x1" in dp or "1 kez" in dp:
					p_idx = 0
				elif "2x1" in dp or "2 kez" in dp:
					p_idx = 1
				elif "3x1" in dp or "3 kez" in dp:
					p_idx = 2
				elif "3x2" in dp:
					p_idx = 3
				elif "4x1" in dp or "4 kez" in dp:
					p_idx = 4
				else:
					p_idx = 6
					custom_pat_init = dp

			dosage_pattern = st.selectbox("🔁 Doz Şablonu", pat_opts, index=p_idx, key=f"{k}pattern")
			if "Diğer" in dosage_pattern:
				custom_pat = st.text_input("Özel Doz Şablonu", value=custom_pat_init, placeholder="Örn: 2 haftada bir", key=f"{k}pat_custom")
				if custom_pat.strip():
					dosage_pattern = custom_pat.strip()

	with r2_c2:
		cur_amt = float(med["dose_amount"] or 1.0) if is_edit and med else 1.0
		dose_amt = st.number_input("Tek Seferlik Doz *", min_value=0.1, max_value=500.0, value=cur_amt, step=0.5, key=f"{k}dose_amt")

	dose_unit = "ölçek"
	with r2_c3:
		if "dose_unit" in vis_fields:
			unit_opts = ["ölçek", "adet (tablet/kapsül)", "kaşık (5 ml)", "ml", "mg", "damla", "puf / fıs", "şase", "tüp"]
			u_idx = 0
			if is_edit and med and med.get("dose_unit"):
				if med["dose_unit"] in unit_opts:
					u_idx = unit_opts.index(med["dose_unit"])
				else:
					unit_opts.insert(0, med["dose_unit"])
					u_idx = 0
			dose_unit = st.selectbox("📏 Doz Birimi", unit_opts, index=u_idx, key=f"{k}unit")

	is_prn = "Gerektiğinde" in dosage_pattern or "PRN" in dosage_pattern or "S.O.S" in dosage_pattern

	# Calculate number of daily doses from dosage pattern
	if is_prn:
		dose_count = 1
		scheduled_times = "Gerektiğinde"
		r3_c1, r3_c2 = st.columns([2.5, 1.2])
		with r3_c1:
			st.info("ℹ️ **Doz Şablonu: Gerektiğinde (PRN / Lüzum Halinde)**\n\nSabit bir uygulama saati belirlenmez. İlaç sadece semptom geliştiğinde (ateş, ağrı, alerji, astım atağı vb.) revirde uygulanacaktır.")
		meal_rel = "Fark etmez"
		with r3_c2:
			if "meal_relation" in vis_fields:
				meal_opts = ["Fark etmez", "Tok karnına", "Aç karnına", "Yemekle birlikte", "Yemekten 30 dk önce"]
				m_idx = meal_opts.index(med["meal_relation"]) if is_edit and med and med.get("meal_relation") in meal_opts else 0
				meal_rel = st.selectbox("🍽️ Açlık / Tokluk", meal_opts, index=m_idx, key=f"{k}meal_prn")
	else:
		if "1x1" in dosage_pattern or "1 kez" in dosage_pattern:
			dose_count = 1
		elif "2x1" in dosage_pattern or "2 kez" in dosage_pattern:
			dose_count = 2
		elif "3x1" in dosage_pattern or "3x2" in dosage_pattern or "3 kez" in dosage_pattern:
			dose_count = 3
		elif "4x1" in dosage_pattern or "4 kez" in dosage_pattern:
			dose_count = 4
		else:
			if dosage_pattern and dosage_pattern[0].isdigit():
				try:
					dose_count = max(1, min(6, int(dosage_pattern.split("x")[0].strip())))
				except Exception:
					dose_count = 2
			else:
				dose_count = 2

		# Parse existing scheduled times if editing
		existing_parsed = []
		if is_edit and med and med.get("scheduled_times") and med["scheduled_times"] != "Gerektiğinde":
			for tp in med["scheduled_times"].split(","):
				tp_clean = tp.strip()
				if tp_clean:
					try:
						h, m = tp_clean.split(":")
						existing_parsed.append(time(int(h), int(m)))
					except Exception:
						pass

		r3_c1, r3_c2 = st.columns([2.5, 1.2])
		with r3_c1:
			st.markdown("##### ⏰ Revirde / Okulda Verilme Saatleri *")
			st.caption(f"Doz şablonuna göre ({dosage_pattern}) günde **{dose_count} doz** saati belirlenmiştir. Saat & dakika değiştirmek için kutucuklara tıklayabilirsiniz:")
			def_times = DEFAULT_DOSE_TIMES.get(dose_count, [time(9, 30)] * dose_count)
			time_cols = st.columns(dose_count)
			chosen_times = []

			for idx in range(dose_count):
				if idx < len(existing_parsed):
					def_t = existing_parsed[idx]
				elif idx < len(def_times):
					def_t = def_times[idx]
				else:
					def_t = time(9 + idx * 2, 0)

				with time_cols[idx]:
					t_val = st.time_input(
						f"{idx + 1}. Doz Saati",
						value=def_t,
						step=300,
						key=f"{k}time_{dosage_pattern}_{dose_count}_{idx}"
					)
					if t_val:
						chosen_times.append(t_val)

			formatted_times = [t.strftime("%H:%M") for t in sorted(chosen_times)]
			scheduled_times = ", ".join(formatted_times)
			st.caption(f"📌 Kaydedilecek Doz Saatleri: **{scheduled_times}** ({len(formatted_times)} Doz)")

		meal_rel = "Tok karnına"
		with r3_c2:
			if "meal_relation" in vis_fields:
				st.markdown("<div style='height: 28px'></div>", unsafe_allow_html=True)
				meal_opts = ["Tok karnına", "Aç karnına", "Yemekle birlikte", "Yemekten 30 dk önce", "Fark etmez"]
				m_idx = meal_opts.index(med["meal_relation"]) if is_edit and med and med.get("meal_relation") in meal_opts else 0
				meal_rel = st.selectbox("🍽️ Açlık / Tokluk", meal_opts, index=m_idx, key=f"{k}meal")

	st.markdown("---")
	st.markdown("##### 📦 Takvim, Stok & Güvenlik Koşulları")

	init_vol = 100.0
	if is_edit and med:
		if med.get("remaining_volume") is not None:
			init_vol = float(med["remaining_volume"])
		elif med.get("total_volume") is not None:
			init_vol = float(med["total_volume"])

	total_vol = None
	stock_unit = "adet"
	duration_d = 0
	start_d = date.today()
	calc_end = None

	if is_prn:
		r4_c1, r4_c2, r4_c3 = st.columns([1.6, 1.0, 2.0])
		with r4_c1:
			if "total_volume" in vis_fields:
				total_vol = st.number_input(
					"📦 Teslim Edilen İlaç Miktarı",
					min_value=0.0,
					max_value=5000.0,
					value=init_vol,
					step=1.0,
					help="Veliden teslim alınan toplam miktar (örn: 100 ml veya 20 adet tablet). Her verildiğinde otomatik düşülür.",
					key=f"{k}tot_vol_prn"
				)
		with r4_c2:
			if "total_volume" in vis_fields:
				stk_opts = ["adet", "ml", "ölçek", "damla", "puf", "şase", "tüp"]
				stk_idx = 0
				if is_edit and med:
					c_stk = med.get("stock_unit") or med.get("dose_unit") or "adet"
					if c_stk in stk_opts:
						stk_idx = stk_opts.index(c_stk)
					else:
						stk_opts.insert(0, c_stk)
						stk_idx = 0
				stock_unit = st.selectbox(
					"Birim",
					stk_opts,
					index=stk_idx,
					key=f"{k}stock_unit_prn",
					help="Tablet/kapsüller için 'adet', şurup/süspansiyonlar için 'ml' seçin."
				)
		with r4_c3:
			st.markdown("<div style='height: 28px'></div>", unsafe_allow_html=True)
			st.caption("ℹ️ **Tedavi Süresi:** Süresiz / İhtiyaç duyuldukça (Tarih ve gün sınırı olmadan aktif kalır)")
	else:
		r4_c1, r4_c2, r4_c3, r4_c4 = st.columns([1.4, 0.9, 1.0, 1.2])
		with r4_c1:
			if "total_volume" in vis_fields:
				total_vol = st.number_input(
					"📦 Teslim Edilen İlaç Miktarı",
					min_value=0.0,
					max_value=5000.0,
					value=init_vol,
					step=1.0,
					help="Veliden teslim alınan toplam miktar (örn: 100 ml veya 20 adet tablet). Her verildiğinde otomatik düşülür.",
					key=f"{k}tot_vol"
				)
		with r4_c2:
			if "total_volume" in vis_fields:
				stk_opts = ["adet", "ml", "ölçek", "damla", "puf", "şase", "tüp"]
				stk_idx = 0
				if is_edit and med:
					c_stk = med.get("stock_unit") or med.get("dose_unit") or "adet"
					if c_stk in stk_opts:
						stk_idx = stk_opts.index(c_stk)
					else:
						stk_opts.insert(0, c_stk)
						stk_idx = 0
				stock_unit = st.selectbox(
					"Birim",
					stk_opts,
					index=stk_idx,
					key=f"{k}stock_unit",
					help="Tablet/kapsüller için 'adet', şurup/süspansiyonlar için 'ml' seçin."
				)

		cur_dur = 7
		if is_edit and med and med.get("duration_days"):
			cur_dur = max(1, int(med["duration_days"]))

		duration_d = cur_dur
		with r4_c3:
			if "duration_days" in vis_fields:
				duration_d = int(st.number_input("📅 Süre (Gün)", min_value=1, max_value=365, value=cur_dur, step=1, key=f"{k}dur_days"))

		cur_start_d = date.today()
		if is_edit and med and med.get("start_date"):
			try:
				cur_start_d = date.fromisoformat(med["start_date"][:10])
			except Exception:
				pass

		with r4_c4:
			start_d = st.date_input("🗓️ Başlangıç Tarihi", value=cur_start_d, key=f"{k}start_d")
			calc_end = start_d + timedelta(days=duration_d - 1)
			st.caption(f"🏁 Tahmini Bitiş: **{calc_end.strftime('%d.%m.%Y')}**")

	r5_c1, r5_c2 = st.columns(2)
	active_days = "Sadece Okul Günleri (Pazartesi - Cuma)"
	with r5_c1:
		if "active_days" in vis_fields:
			act_opts = ["Sadece Okul Günleri (Pazartesi - Cuma)", "Her Gün (Hafta Sonu Dahil)", "Pazartesi, Çarşamba, Cuma", "Salı, Perşembe", "Hafta Sonu"]
			act_idx = 0
			if is_edit and med and med.get("active_days"):
				if med["active_days"] in act_opts:
					act_idx = act_opts.index(med["active_days"])
				else:
					act_opts.insert(0, med["active_days"])
					act_idx = 0
			active_days = st.selectbox("🗓️ Uygulanacak Günler", act_opts, index=act_idx, key=f"{k}act_days")

	storage_cond = "Oda Sıcaklığı (15-25 °C)"
	with r5_c2:
		if "storage_condition" in vis_fields:
			stor_opts = ["Oda Sıcaklığı (15-25 °C)", "Buzdolabı (+2 / +8 °C)", "Kilitli Revir Dolabı", "Işıktan Korunan Dolap"]
			stor_idx = 0
			if is_edit and med and med.get("storage_condition"):
				if med["storage_condition"] in stor_opts:
					stor_idx = stor_opts.index(med["storage_condition"])
				else:
					stor_opts.insert(0, med["storage_condition"])
					stor_idx = 0
			storage_cond = st.selectbox("❄️ Saklama Koşulu", stor_opts, index=stor_idx, key=f"{k}storage")

	r6_c1, r6_c2 = st.columns([1.5, 1.5])
	guardian_consent = 1
	with r6_c1:
		if "guardian_consent" in vis_fields:
			c_val = True
			if is_edit and med:
				c_val = bool(med.get("guardian_consent", 1))
			guardian_consent = 1 if st.checkbox("📝 Veli İlaç Teslim & Uygulama İzin Formu Alındı", value=c_val, key=f"{k}consent") else 0

	exp_d = None
	with r6_c2:
		if "expiry_date" in vis_fields:
			exp_init = date.today() + timedelta(days=180)
			if is_edit and med and med.get("expiry_date"):
				try:
					exp_init = date.fromisoformat(med["expiry_date"][:10])
				except Exception:
					pass
			exp_val = st.date_input("⏳ Son Kullanma Tarihi (SKT)", value=exp_init, key=f"{k}exp")
			exp_d = exp_val.isoformat()

	spec_notes = ""
	if "special_notes" in vis_fields:
		n_init = med["special_notes"] if (is_edit and med and med.get("special_notes")) else ""
		spec_notes = st.text_area("⚠️ Özel Hemşirelik Notu / Uyarılar", value=n_init, placeholder="Örn: Bol su ile içirilecek, astım semptomu başlarsa verilecek, ilacı aldıktan sonra 10 dk dinlendirilecek...", key=f"{k}notes")

	st.markdown("---")

	if is_edit and active_edit_id:
		col_btn_save, col_btn_del = st.columns([2.5, 1])
		with col_btn_save:
			if st.button("💾 Değişiklikleri Güncelle", type="primary", use_container_width=True, key=f"{k}save_btn"):
				if not med_name.strip():
					st.error("Lütfen ilaç adını girin.")
					return
				if not is_prn and not scheduled_times.strip():
					st.error("Lütfen en az bir uygulama saati girin (Örn: 09:30).")
					return

				update_student_medication(
					medication_id=active_edit_id,
					medication_name=med_name.strip(),
					med_type=dose_unit,
					route=med_route,
					dosage_pattern=dosage_pattern,
					dose_amount=float(dose_amt),
					dose_unit=dose_unit,
					stock_unit=stock_unit,
					meal_relation=meal_rel,
					scheduled_times=scheduled_times.strip() if not is_prn else "Gerektiğinde",
					total_volume=float(total_vol) if total_vol is not None else None,
					remaining_volume=float(total_vol) if total_vol is not None else None,
					duration_days=int(duration_d) if not is_prn else 0,
					start_date=start_d.isoformat(),
					end_date=calc_end.isoformat() if (calc_end and not is_prn) else None,
					active_days=active_days,
					storage_condition=storage_cond,
					guardian_consent=guardian_consent,
					expiry_date=exp_d,
					special_notes=spec_notes.strip(),
					status=med_status
				)
				st.session_state["med_dialog_target_id"] = None
				st.success(f"'{med_name.strip()}' ilacı ve tedavi planı başarıyla güncellendi!")
				st.rerun()

		with col_btn_del:
			if st.button("🗑️ Tedavi Planını Sil", type="secondary", use_container_width=True, key=f"{k}del_btn"):
				delete_student_medication(active_edit_id)
				st.session_state["med_dialog_target_id"] = None
				st.warning("İlaç ve tedavi planı silindi.")
				st.rerun()
	else:
		if st.button("💾 İlacı ve Tedavi Planını Kaydet", type="primary", use_container_width=True, key=f"{k}save_btn"):
			if not selected_stu_id:
				st.error("Lütfen bir öğrenci seçin.")
				return
			if not med_name.strip():
				st.error("Lütfen ilaç adını girin.")
				return
			if not is_prn and not scheduled_times.strip():
				st.error("Lütfen en az bir uygulama saati girin (Örn: 09:30).")
				return

			add_student_medication(
				student_id=selected_stu_id,
				medication_name=med_name.strip(),
				med_type=dose_unit,
				route=med_route,
				dosage_pattern=dosage_pattern,
				dose_amount=float(dose_amt),
				dose_unit=dose_unit,
				stock_unit=stock_unit,
				meal_relation=meal_rel,
				scheduled_times=scheduled_times.strip() if not is_prn else "Gerektiğinde",
				total_volume=float(total_vol) if total_vol is not None else None,
				duration_days=int(duration_d) if not is_prn else 0,
				start_date=start_d.isoformat(),
				end_date=calc_end.isoformat() if (calc_end and not is_prn) else None,
				active_days=active_days,
				storage_condition=storage_cond,
				guardian_consent=guardian_consent,
				expiry_date=exp_d,
				special_notes=spec_notes.strip(),
				status="Aktif"
			)
			st.session_state["med_add_time_slots"] = [time(9, 30)]
			st.session_state["med_dialog_target_id"] = None
			st.success(f"'{med_name.strip()}' ilacı başarıyla kaydedildi ve takip takvimine eklendi!")
			st.rerun()


def open_edit_medication_dialog(medication_id: int) -> None:
	st.session_state["med_dialog_target_id"] = medication_id
	open_add_medication_dialog(edit_medication_id=medication_id)


def render_medication_tracker() -> None:
	app_cfg = get_app_settings()
	nurse_name = app_cfg.get("nurse_name", "Okul Revir Sorumlusu")

	if "med_tracker_date" not in st.session_state:
		st.session_state["med_tracker_date"] = date.today()

	top_col1, top_col2 = st.columns([7, 3])
	with top_col1:
		st.markdown("## 💊 İlaç Takip Sistemi & Günlük İlaç Çizelgesi")
		st.caption("Öğrencilerin okulda alması gereken reçeteli ilaçları saatlik çizelge üzerinden takip edin, doz verildiğinde tek tıkla kaydedin ve ilaç stoklarını denetleyin.")
	with top_col2:
		st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
		if st.button("➕ Yeni İlaç ve Tedavi Planı Ekle", type="primary", use_container_width=True, key="btn_open_add_med_top"):
			st.session_state["med_dialog_target_id"] = None
			open_add_medication_dialog()

	nav_c1, nav_c2, nav_c3, nav_c4, nav_c5 = st.columns([1, 1.8, 1, 1, 2.5])
	with nav_c1:
		if st.button("◀️ Önceki Gün", use_container_width=True, key="btn_med_prev_day"):
			st.session_state["med_tracker_date"] = st.session_state["med_tracker_date"] - timedelta(days=1)
			st.rerun()
	with nav_c2:
		cur_date_val = st.date_input(
			"📅 Takip Tarihi",
			value=st.session_state["med_tracker_date"],
			key="med_date_picker",
			label_visibility="collapsed"
		)
		if cur_date_val != st.session_state["med_tracker_date"]:
			st.session_state["med_tracker_date"] = cur_date_val
			st.rerun()
	with nav_c3:
		if st.button("🔵 Bugün", use_container_width=True, key="btn_med_today"):
			st.session_state["med_tracker_date"] = date.today()
			st.rerun()
	with nav_c4:
		if st.button("Sonraki Gün ▶️", use_container_width=True, key="btn_med_next_day"):
			st.session_state["med_tracker_date"] = st.session_state["med_tracker_date"] + timedelta(days=1)
			st.rerun()

	selected_date = st.session_state["med_tracker_date"]
	turkish_day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
	turkish_month_names = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
	day_label = f"{selected_date.day} {turkish_month_names[selected_date.month]} {selected_date.year}, {turkish_day_names[selected_date.weekday()]}"

	with nav_c5:
		is_today_str = " (Bugün)" if selected_date == date.today() else ""
		st.markdown(f"<div style='font-size: 1.05rem; font-weight: 700; color: var(--teal); padding-top: 6px;'>🗓️ {day_label}{is_today_str}</div>", unsafe_allow_html=True)

	st.markdown("---")

	scheduled_doses = get_scheduled_doses_for_date(selected_date)

	total_doses_count = len(scheduled_doses)
	given_doses_count = sum(1 for d in scheduled_doses if d["is_logged"] and d["log_status"] == "Verildi")
	missed_doses_count = sum(1 for d in scheduled_doses if d["is_logged"] and d["log_status"] != "Verildi")
	pending_doses_count = total_doses_count - (given_doses_count + missed_doses_count)

	all_active_meds = get_all_student_medications(status="Aktif")
	low_stock_meds = [m for m in all_active_meds if m["remaining_volume"] is not None and m["remaining_volume"] <= 3 * (m["dose_amount"] or 1.0)]

	m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
	with m_col1:
		st.metric("📋 Toplam Planlanan Doz", f"{total_doses_count} Doz")
	with m_col2:
		st.metric("✅ Verilen / Uygulanan", f"{given_doses_count} Doz", delta=f"%{round((given_doses_count/total_doses_count)*100):.0f}" if total_doses_count > 0 else "0%")
	with m_col3:
		st.metric("⏳ Bekleyen Dozlar", f"{pending_doses_count} Doz", delta=None if pending_doses_count == 0 else f"{pending_doses_count} bekliyor", delta_color="inverse")
	with m_col4:
		st.metric("⚠️ Verilemeyen / Yok", f"{missed_doses_count} Doz", delta=None if missed_doses_count == 0 else "İncelenmeli", delta_color="inverse")
	with m_col5:
		st.metric("📦 Azalan İlaç Stokları", f"{len(low_stock_meds)} İlaç", delta="Kritik Seviye" if len(low_stock_meds) > 0 else "Yeterli", delta_color="inverse" if len(low_stock_meds) > 0 else "normal")

	st.markdown("<div style='height: 10px'></div>", unsafe_allow_html=True)

	tab_timetable, tab_cabinet, tab_history = st.tabs([
		f"🕒 Günlük Saatlik İlaç Çizelgesi ({total_doses_count})",
		f"📦 İlaç Dolabı & Aktif Tedavi Planları ({len(all_active_meds)})",
		"📜 Revir İlaç Defteri & Geçmiş Kayıtlar"
	])

	# TAB 1: GÜNLÜK SAATLİK ÇİZELGE
	with tab_timetable:
		tb_top_c1, tb_top_c2, tb_top_c3 = st.columns([3, 1.5, 1.5])
		with tb_top_c1:
			st.markdown(f"#### 🕒 {day_label} — Saatlik İlaç Uygulama Çizelgesi")
		with tb_top_c2:
			view_mode = st.radio(
				"Görünüm:",
				["Zaman Çizelgesi", "Tablo Görünümü"],
				horizontal=True,
				label_visibility="collapsed",
				key="med_tb_view_mode"
			)
		with tb_top_c3:
			try:
				docx_data = create_daily_medication_docx(selected_date, scheduled_doses)
				st.download_button(
					"📄 Çizelgeyi Word İndir",
					data=docx_data,
					file_name=f"Ilac_Takip_Cizelgesi_{selected_date.isoformat()}.docx",
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					use_container_width=True,
					key="btn_down_med_docx"
				)
			except Exception as e:
				st.caption(f"Word indirme hazırlanamadı: {e}")

		if not scheduled_doses:
			st.info(f"📅 **{day_label}** tarihinde verilecek planlanmış öğrenci ilacı bulunmamaktadır. Yukarıdaki '➕ Yeni İlaç ve Tedavi Planı Ekle' butonundan öğrenciye ilaç tanımlayabilirsiniz.")
		else:
			# Group doses by hour
			doses_by_hour = {}
			for d in scheduled_doses:
				t_str = str(d.get("scheduled_time", "12:00")).strip()
				doses_by_hour.setdefault(t_str, []).append(d)

			sorted_hours = sorted(doses_by_hour.keys())

			if view_mode == "Zaman Çizelgesi":
				for h_time in sorted_hours:
					h_doses = doses_by_hour[h_time]
					h_given = sum(1 for d in h_doses if d["is_logged"] and d["log_status"] == "Verildi")
					h_total = len(h_doses)
					badge_color = "#10b981" if h_given == h_total else "#f59e0b"

					time_display = f"🕒 Saat {h_time}" if h_time != "Gerektiğinde" else "🕒 İhtiyaç Halinde / Gerektiğinde (PRN)"
					time_info = f"Toplam: {h_total} Doz | Verilen: {h_given} / {h_total}" if h_time != "Gerektiğinde" else f"Toplam: {h_total} Öğrenci | Verilen: {h_given} / {h_total} (Saat Sınırı Yoktur)"

					st.markdown(
						f"""<div style='background: rgba(0,0,0,0.03); border-left: 4px solid {badge_color}; border-radius: 4px; padding: 6px 12px; margin-top: 14px; margin-bottom: 8px;'>
							<span style='font-size: 1.15rem; font-weight: 700;'>{time_display}</span>
							<span style='margin-left: 14px; font-size: 0.9rem; color: #64748b;'>{time_info}</span>
						</div>""",
						unsafe_allow_html=True
					)

					for dose_idx, dose in enumerate(h_doses):
						card_key = f"dose_{dose['medication_id']}_{dose['scheduled_time']}_{selected_date.isoformat()}"
						with st.container(border=True):
							c_stu, c_med, c_stock, c_act = st.columns([2.2, 2.8, 1.8, 2.2])

							with c_stu:
								sno_txt = f" (No: {dose['student_no']})" if dose.get("student_no") else ""
								st.markdown(f"👤 **{dose['full_name']}**{sno_txt}")
								st.caption(f"🏫 Sınıf: **{dose['class_name']}**")
								if dose.get("allergies") and dose["allergies"].strip():
									st.markdown(f"<span style='color: #ef4444; font-size: 0.8rem; font-weight: 600;'>🚨 Alerji: {dose['allergies']}</span>", unsafe_allow_html=True)
								if dose.get("chronic_conditions") and dose["chronic_conditions"].strip():
									st.caption(f"🏥 Kronik: {dose['chronic_conditions']}")

							with c_med:
								st.markdown(f"💊 **{dose['medication_name']}**")
								route_txt = dose.get("route") or "Oral"
								st.markdown(f"💉 Yol: **{route_txt}** | 📏 Doz: **{dose['dose_amount']} {dose['dose_unit']}** ({dose['dosage_pattern']})")
								st.caption(f"🍽️ {dose.get('meal_relation', 'Tok karnına')} | ❄️ {dose.get('storage_condition', 'Oda Sıcaklığı')}")
								if dose.get("special_notes") and dose["special_notes"].strip():
									st.markdown(f"<span style='font-size: 0.8rem; color: #0284c7;'>📝 Not: {dose['special_notes']}</span>", unsafe_allow_html=True)

							with c_stock:
								rem_v = dose.get("remaining_volume")
								tot_v = dose.get("total_volume")
								st_unit = dose.get("stock_unit") or dose.get("dose_unit") or "adet"
								if rem_v is not None:
									is_low = rem_v <= (3 * dose["dose_amount"])
									stock_color = "#ef4444" if is_low else "#10b981"
									tot_str = f" / {tot_v}" if tot_v is not None else ""
									st.markdown(f"📦 Stok: <b style='color:{stock_color}'>{rem_v}{tot_str} {st_unit}</b>", unsafe_allow_html=True)
									if is_low:
										st.markdown("<span style='color: #ef4444; font-size: 0.75rem; font-weight: 700;'>⚠️ Stok Azaldı!</span>", unsafe_allow_html=True)
								else:
									st.caption("📦 Stok takibi yok")

							with c_act:
								if not dose["is_logged"]:
									now_time_str = datetime.now().strftime("%H:%M")
									if st.button("✅ Verildi Olarak İşle", key=f"btn_give_{card_key}", type="primary", use_container_width=True):
										log_medication_dose(
											medication_id=dose["medication_id"],
											student_id=dose["student_id"],
											log_date=selected_date.isoformat(),
											scheduled_time=dose["scheduled_time"],
											given_time=now_time_str,
											dose_given=dose["dose_amount"],
											dose_unit=dose["dose_unit"],
											status="Verildi",
											administered_by=nurse_name,
											notes=""
										)
										st.success(f"{dose['full_name']} için ilaç verildi olarak kaydedildi.")
										st.rerun()

									with st.popover("⚠️ Diğer Durumlar", use_container_width=True):
										alt_status = st.selectbox(
											"Uygulanamama Nedeni",
											["Öğrenci Gelmedi / Yoktu", "İlaç Kalmadı / Bitti", "Öğrenci Reddetti / İçmedi", "Veli Talebiyle Verilmedi", "Diğer"],
											key=f"alt_stat_{card_key}"
										)
										alt_note = st.text_input("Açıklama / Not", placeholder="Örn: Revire çağrıldı raporluymuş", key=f"alt_note_{card_key}")
										if st.button("Kaydet", key=f"btn_save_alt_{card_key}", type="secondary", use_container_width=True):
											log_medication_dose(
												medication_id=dose["medication_id"],
												student_id=dose["student_id"],
												log_date=selected_date.isoformat(),
												scheduled_time=dose["scheduled_time"],
												given_time=now_time_str,
												dose_given=dose["dose_amount"],
												dose_unit=dose["dose_unit"],
												status=alt_status,
												administered_by=nurse_name,
												notes=alt_note.strip()
											)
											st.warning(f"Durum kaydedildi: {alt_status}")
											st.rerun()
								else:
									if dose["log_status"] == "Verildi":
										st.markdown(
											f"""<div style='background: #dcfce7; color: #166534; padding: 6px 10px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; text-align: center;'>
												✅ Verildi ({dose.get('given_time', '-')})<br>
												<span style='font-size: 0.75rem; font-weight: normal;'>{dose.get('administered_by', '')}</span>
											</div>""",
											unsafe_allow_html=True
										)
									else:
										st.markdown(
											f"""<div style='background: #fee2e2; color: #991b1b; padding: 6px 10px; border-radius: 6px; font-size: 0.85rem; font-weight: 600; text-align: center;'>
												⚠️ {dose['log_status']}<br>
												<span style='font-size: 0.75rem; font-weight: normal;'>{dose.get('log_notes') or ''}</span>
											</div>""",
											unsafe_allow_html=True
										)

									if st.button("↩️ Geri Al / Düzelt", key=f"btn_undo_{card_key}", use_container_width=True):
										undo_medication_dose(dose["medication_id"], selected_date.isoformat(), dose["scheduled_time"])
										st.info("Kayıt geri alındı.")
										st.rerun()

			else:
				# Table View
				rows_data = []
				for d in scheduled_doses:
					sno = f" (No: {d['student_no']})" if d.get('student_no') else ""
					st_txt = d['log_status']
					if d['is_logged'] and d.get('given_time'):
						st_txt += f" ({d['given_time']})"
					rows_data.append({
						"Saat": d["scheduled_time"],
						"Öğrenci": f"{d['full_name']}{sno}",
						"Sınıf": d["class_name"],
						"İlaç Adı": d["medication_name"],
						"Doz": f"{d['dose_amount']} {d['dose_unit']}",
						"Yol": d["route"],
						"Açlık/Tokluk": d["meal_relation"],
						"Saklama": d["storage_condition"],
						"Durum": st_txt,
						"Uygulayan": d.get("administered_by") or "-",
					})
				df_table = pd.DataFrame(rows_data)
				st.dataframe(df_table, use_container_width=True, hide_index=True)

	# TAB 2: İLAÇ DOLABI & AKTİF TEDAVİ PLANLARI
	with tab_cabinet:
		st.markdown("#### 📦 Okul Reviri İlaç Dolabı & Kayıtlı Tedavi Planları")
		st.caption("Revire velisi tarafından teslim edilen ilaçları, aktif tedavi kürlerini, saklama yerlerini ve kalan stokları buradan yönetebilirsiniz.")

		c_f1, c_f2, c_f3 = st.columns([2, 1.5, 1.5])
		with c_f1:
			cab_search = st.text_input("🔍 İlaç veya Öğrenci Ara", placeholder="Örn: Augmentin, Ali, 2-A...", key="med_cab_search")
		with c_f2:
			cab_status_filter = st.selectbox("Durum", ["Aktif", "Tamamlandı", "İptal Edildi", "Tümü"], index=0, key="med_cab_stat_fil")
		with c_f3:
			cab_storage_filter = st.selectbox("Saklama Yeri", ["Tümü", "Buzdolabı (+2 / +8 °C)", "Oda Sıcaklığı (15-25 °C)", "Kilitli Revir Dolabı"], index=0, key="med_cab_store_fil")

		status_param = None if cab_status_filter == "Tümü" else cab_status_filter
		med_list = get_all_student_medications(status=status_param)

		today = date.today()

		# Top Level Expiry Warning Banners
		expired_items = []
		exp_warn_items = []
		for m in med_list:
			if m.get("status") == "Aktif" and m.get("expiry_date"):
				try:
					ed = date.fromisoformat(str(m["expiry_date"])[:10])
					days = (ed - today).days
					if days < 0:
						expired_items.append((m, ed.strftime("%d.%m.%Y")))
					elif days <= 7:
						exp_warn_items.append((m, days, ed.strftime("%d.%m.%Y")))
				except Exception:
					pass

		if expired_items:
			exp_names = ", ".join([f"**{x[0]['full_name']}** ({x[0]['medication_name']} - SKT: {x[1]})" for x in expired_items])
			st.error(f"🚨 **DİKKAT: Son Kullanma Tarihi Geçmiş İlaç Var!** ({len(expired_items)} adet)\n\n{exp_names}\n\n*Bu ilaçlar kesinlikle uygulanmamalı, derhal velisine teslim edilmeli veya uygun şekilde imha edilmelidir.*")

		if exp_warn_items:
			warn_lines = [f"• **{x[0]['full_name']}** ({x[0]['class_name'] or 'Sınıf Yok'}) — *{x[0]['medication_name']}*: Son **{x[1]} gün** kaldı (SKT: {x[2]})" for x in exp_warn_items]
			st.warning(f"⚠️ **SKT Yaklaşan İlaç Uyarısı (Son 1 Hafta / ≤ 7 Gün - {len(exp_warn_items)} adet):**\n\n" + "\n\n".join(warn_lines) + "\n\n*Lütfen yeni ilaç temini için öğrenci velileri ile iletişime geçiniz.*")

		if cab_search.strip():
			q_clean = cab_search.strip().lower()
			med_list = [
				m for m in med_list
				if q_clean in str(m.get("full_name", "")).lower()
				or q_clean in str(m.get("medication_name", "")).lower()
				or q_clean in str(m.get("class_name", "")).lower()
			]

		if cab_storage_filter != "Tümü":
			med_list = [m for m in med_list if cab_storage_filter in str(m.get("storage_condition", ""))]

		if not med_list:
			st.info("Kriterlere uygun kayıtlı ilaç bulunamadı.")
		else:
			for m_item in med_list:
				# Calculate expiration dates and remaining days
				exp_str = m_item.get("expiry_date")
				exp_days_left = None
				exp_formatted = None
				if exp_str and str(exp_str).strip():
					try:
						exp_d = date.fromisoformat(str(exp_str)[:10])
						exp_days_left = (exp_d - today).days
						exp_formatted = exp_d.strftime("%d.%m.%Y")
					except Exception:
						exp_formatted = str(exp_str)[:10]

				with st.container(border=True):
					mc_1, mc_2, mc_3, mc_4 = st.columns([2.5, 2.5, 2.2, 1.8])
					with mc_1:
						sno_str = f" (No: {m_item['student_no']})" if m_item.get("student_no") else ""
						st.markdown(f"👤 **{m_item['full_name']}**{sno_str}")
						st.caption(f"🏫 Sınıf: **{m_item['class_name'] or 'Sınıfsız'}**")
						if m_item.get("allergies"):
							st.markdown(f"<span style='color: #ef4444; font-size: 0.8rem;'>🚨 Alerji: {m_item['allergies']}</span>", unsafe_allow_html=True)
						# SKT Warning Badge on the student side
						if exp_days_left is not None and exp_days_left < 0:
							st.markdown(f"<div style='background:#fee2e2;color:#991b1b;border:1px solid #f87171;padding:2px 8px;border-radius:4px;font-size:11.5px;font-weight:700;margin-top:4px;'>🚨 İLAÇ SKT DOLDU ({exp_formatted})</div>", unsafe_allow_html=True)
						elif exp_days_left is not None and exp_days_left <= 7:
							st.markdown(f"<div style='background:#fffbeb;color:#b45309;border:1px solid #fde68a;padding:2px 8px;border-radius:4px;font-size:11.5px;font-weight:700;margin-top:4px;'>⚠️ SKT'ye {exp_days_left} Gün Kaldı ({exp_formatted})</div>", unsafe_allow_html=True)

					with mc_2:
						st.markdown(f"💊 **{m_item['medication_name']}**")
						st.caption(f"💉 {m_item.get('route', 'Oral')} | 🔁 {m_item.get('dosage_pattern', '1x1')} | 📏 {m_item.get('dose_amount', 1)} {m_item.get('dose_unit', 'adet')}")
						st.caption(f"⏰ Saatler: **{m_item.get('scheduled_times', '-')}**")
						# Explicit SKT Information Display
						if exp_formatted:
							if exp_days_left is not None and exp_days_left < 0:
								st.markdown(f"⏳ **SKT:** <span style='color:#dc2626;font-weight:700;'>{exp_formatted} (Süresi Doldu!)</span>", unsafe_allow_html=True)
							elif exp_days_left is not None and exp_days_left <= 7:
								st.markdown(f"⏳ **SKT:** <span style='color:#d97706;font-weight:700;'>{exp_formatted} (Son {exp_days_left} gün kaldı!)</span>", unsafe_allow_html=True)
							else:
								st.caption(f"⏳ **SKT:** {exp_formatted} ({exp_days_left} gün kaldı)")
						else:
							st.caption("⏳ **SKT:** Belirtilmedi")

					with mc_3:
						rem_v = m_item.get("remaining_volume")
						tot_v = m_item.get("total_volume")
						st_unit = m_item.get("stock_unit") or m_item.get("dose_unit") or "adet"
						if rem_v is not None:
							is_low = rem_v <= (3 * (m_item.get("dose_amount") or 1.0))
							col_s = "#ef4444" if is_low else "#10b981"
							tot_str = f" / {tot_v}" if tot_v is not None else ""
							st.markdown(f"📦 Kalan: <b style='color:{col_s}'>{rem_v}{tot_str} {st_unit}</b>", unsafe_allow_html=True)
						st.caption(f"❄️ {m_item.get('storage_condition', 'Oda Sıcaklığı')}")
						is_item_prn = "Gerektiğinde" in str(m_item.get("dosage_pattern", "")) or str(m_item.get("scheduled_times", "")) == "Gerektiğinde"
						if is_item_prn:
							st.caption(f"📅 Tedavi: {m_item.get('start_date', '')} — Süresiz (Gerektiğinde)")
						else:
							st.caption(f"📅 Tedavi: {m_item.get('start_date', '')} — {m_item.get('end_date') or 'Süresiz'}")

					with mc_4:
						m_id = m_item["id"]
						if st.button("✏️ Düzenle", key=f"btn_edit_med_{m_id}", use_container_width=True):
							open_edit_medication_dialog(m_id)
						if m_item.get("status") == "Aktif":
							if st.button("✔️ Tamamla", key=f"btn_done_med_{m_id}", use_container_width=True):
								update_student_medication(medication_id=m_id, status="Tamamlandı")
								st.success("Tedavi planı tamamlandı olarak işaretlendi.")
								st.rerun()

	# TAB 3: REVİR İLAÇ DEFTERİ (UYGULAMA GEÇMİŞİ)
	with tab_history:
		st.markdown("#### 📜 Revir İlaç Uygulama Defteri & Denetim Kayıtları")
		st.caption("Revirde öğrencilere verilen tüm ilaçların tarih, saat, uygulayıcı hemşire ve doz bazındaki resmi geçmiş kayıtları.")

		h_c1, h_c2, h_c3 = st.columns([1.5, 1.5, 2])
		with h_c1:
			hist_start = st.date_input("Başlangıç Tarihi", value=date.today() - timedelta(days=14), key="hist_d_start")
		with h_c2:
			hist_end = st.date_input("Bitiş Tarihi", value=date.today(), key="hist_d_end")
		with h_c3:
			hist_search = st.text_input("🔍 Öğrenci / İlaç Ara", placeholder="İsim veya ilaç yazın...", key="hist_search_txt")

		with connect_db() as conn:
			query = """
				SELECT l.log_date, l.scheduled_time, l.given_time, s.full_name, s.student_no, c.name as class_name,
				       m.medication_name, l.dose_given, l.dose_unit, l.status, l.administered_by, l.notes
				FROM medication_logs l
				JOIN students s ON s.id = l.student_id
				LEFT JOIN classes c ON c.id = s.class_id
				JOIN student_medications m ON m.id = l.medication_id
				WHERE l.log_date >= ? AND l.log_date <= ?
			"""
			params = [hist_start.isoformat(), hist_end.isoformat()]
			if hist_search.strip():
				query += " AND (s.full_name LIKE ? OR m.medication_name LIKE ?)"
				s_param = f"%{hist_search.strip()}%"
				params.extend([s_param, s_param])
			query += " ORDER BY l.log_date DESC, l.given_time DESC"

			hist_df = pd.read_sql_query(query, conn, params=params)

		if hist_df.empty:
			st.info("Seçilen tarih aralığında ilaç uygulama kaydı bulunamadı.")
		else:
			hist_df = hist_df.rename(columns={
				"log_date": "Tarih",
				"scheduled_time": "Planlanan Saat",
				"given_time": "Verilme Saati",
				"full_name": "Öğrenci Adı",
				"student_no": "Okul No",
				"class_name": "Sınıf",
				"medication_name": "İlaç Adı",
				"dose_given": "Verilen Doz",
				"dose_unit": "Birim",
				"status": "Durum",
				"administered_by": "Uygulayan",
				"notes": "Notlar"
			})
			st.dataframe(hist_df, use_container_width=True, hide_index=True)

			# Excel Export
			excel_buf = io.BytesIO()
			with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
				hist_df.to_excel(writer, index=False, sheet_name="Ilac_Uygulama_Kayitlari")
			excel_buf.seek(0)

			st.download_button(
				"📥 İlaç Uygulama Defterini Excel Olarak İndir",
				data=excel_buf,
				file_name=f"Revir_Ilac_Defteri_{hist_start.isoformat()}_{hist_end.isoformat()}.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				key="btn_down_hist_excel"
			)


def render_vital_tracking() -> None:
	st.markdown("## Vital ve laboratuvar takibi")
	student = select_shared_person("vital", label="Öğrenci veya kişi seçin")
	if not student:
		st.caption("Vital bulguları, boy-kilo persantil özetini ve laboratuvar sonuçlarını görmek için kişi seçin.")
		return
	render_shared_person_summary(student, section_key="vital")

	vital_tab, lab_tab = st.tabs(["💓 Vital Ölçümleri & Antropometri", "🧪 Tahlil Sonuçları"])
	with vital_tab:
		st.caption("Kayıtlar kişinin sağlık geçmişine eklenir. Aynı anda boy, kilo veya baş çevresi girerseniz Persantil/Z-Skor tablosuna da otomatik işlenir.")
		with st.form("vital_form", clear_on_submit=True):
			when, context = st.columns(2)
			reading_date = when.date_input("Ölçüm tarihi", value=date.today(), key="vital_date")
			reading_time = context.time_input("Saat", value=datetime.now().replace(second=0, microsecond=0).time(), key="vital_time")
			bp_cols = st.columns(3)
			systolic = bp_cols[0].number_input("Büyük tansiyon (mmHg)", min_value=40.0, max_value=260.0, value=None, step=1.0)
			diastolic = bp_cols[1].number_input("Küçük tansiyon (mmHg)", min_value=20.0, max_value=180.0, value=None, step=1.0)
			spo2 = bp_cols[2].number_input("Saturasyon - SpO2 (%)", min_value=30.0, max_value=100.0, value=None, step=1.0, format="%.0f")
			other_cols = st.columns(3)
			pulse = other_cols[0].number_input("Nabız (atım/dk)", min_value=20.0, max_value=240.0, value=None, step=1.0)
			temperature = other_cols[1].number_input("Ateş (°C)", min_value=30.0, max_value=45.0, value=None, step=0.1)
			glucose = other_cols[2].number_input("Kan şekeri (mg/dL)", min_value=1.0, max_value=1500.0, value=None, step=1.0)
			hw_cols = st.columns(4)
			glucose_context = hw_cols[0].selectbox("Kan şekeri ölçüm koşulu", ["Belirtilmedi", "Açlık", "Yemek sonrası", "Rastgele", "Hekim istemiyle"])
			v_height = hw_cols[1].number_input("Boy (cm) [Persantile işlenir]", min_value=30.0, max_value=250.0, value=None, step=0.1, format="%.1f")
			v_weight = hw_cols[2].number_input("Kilo (kg) [Persantile işlenir]", min_value=1.0, max_value=300.0, value=None, step=0.1, format="%.1f")
			v_head_circ = hw_cols[3].number_input("Baş Çevresi (cm) [Persantile işlenir]", min_value=20.0, max_value=100.0, value=None, step=0.1, format="%.1f")
			notes = st.text_area("Belirti / ölçüm koşulu / not", key="vital_notes")
			save = st.form_submit_button("Ölçümleri kaydet", type="primary")
		if save:
			values = {"systolic": systolic, "diastolic": diastolic, "pulse": pulse,
					  "temperature": temperature, "spo2": spo2, "glucose": glucose}
			has_vital = any(value is not None for value in values.values())
			has_hw = (v_height is not None and v_weight is not None and v_height > 0 and v_weight > 0) or (v_head_circ is not None and v_head_circ > 0)
			if not has_vital and not has_hw:
				st.error("Kaydetmeden önce en az bir vital ölçüm veya boy-kilo / baş çevresi girin.")
			elif (systolic is None) != (diastolic is None):
				st.error("Tansiyon kaydı için büyük ve küçük tansiyonu birlikte girin.")
			elif (v_height is None) != (v_weight is None):
				st.error("Boy-kilo kaydı için hem boy hem kiloyu birlikte girin.")
			else:
				extra_note = notes.strip()
				if has_hw:
					hw_summary = record_linked_measurement(student["id"], reading_date, v_height, v_weight, student["birth_date"], student["sex"], head_circ_cm=v_head_circ)
					if hw_summary:
						extra_note = " · ".join(filter(None, [extra_note, f"[{hw_summary}]"]))
				if has_vital:
					measured_at = datetime.combine(reading_date, reading_time).isoformat(timespec="minutes")
					record_linked_vital(
						student["id"], measured_at, systolic, diastolic, pulse, temperature, glucose,
						None if glucose_context == "Belirtilmedi" else glucose_context, extra_note or None, spo2=spo2
					)
				st.success("Ölçümler öğrenci geçmişine ve bağlantılı modüllere kaydedildi.")
				st.rerun()
		with connect_db() as connection:
			vitals = pd.read_sql_query(
				"""SELECT id, measured_at AS Tarih_Saat, systolic_mmhg AS Buyuk_Tansiyon,
				   diastolic_mmhg AS Kucuk_Tansiyon, pulse_bpm AS Nabiz, temperature_c AS Ates_C,
				   spo2_pct AS Saturasyon_SpO2, glucose_mg_dl AS Kan_Sekeri, glucose_context AS Olcum_Kosulu, notes AS Notlar
				   FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC""",
				connection, params=(student["id"],),
			)
		if not vitals.empty:
			latest = vitals.iloc[0]
			st.markdown("### Son kayıt")
			latest_cols = st.columns(5)
			for column, label, key, unit in [
				(latest_cols[0], "Tansiyon", "Buyuk_Tansiyon", "mmHg"),
				(latest_cols[1], "Nabız", "Nabiz", "atım/dk"),
				(latest_cols[2], "Ateş", "Ates_C", "°C"),
				(latest_cols[3], "Saturasyon (SpO2)", "Saturasyon_SpO2", "%"),
				(latest_cols[4], "Kan şekeri", "Kan_Sekeri", "mg/dL"),
			]:
				value = latest[key]
				if pd.notna(value):
					display_value = f"{latest['Buyuk_Tansiyon']:.0f}/{latest['Kucuk_Tansiyon']:.0f}" if key == "Buyuk_Tansiyon" else f"{value:g} {unit}"
					column.metric(label, display_value)
			assessments = []
			measured_on = datetime.fromisoformat(str(latest["Tarih_Saat"])).date()
			for indicator, label, value in [
				("systolic_bp", "Büyük tansiyon", latest["Buyuk_Tansiyon"]),
				("diastolic_bp", "Küçük tansiyon", latest["Kucuk_Tansiyon"]),
				("pulse", "Nabız", latest["Nabiz"]),
				("temperature", "Ateş", latest["Ates_C"]),
				("spo2", "Saturasyon (SpO2)", latest["Saturasyon_SpO2"]),
				("glucose", "Kan şekeri", latest["Kan_Sekeri"]),
			]:
				if pd.isna(value):
					continue
				assessment = vital_assessment(indicator, float(value), student, measured_on)
				assessments.append({"Ölçüm": label, "Değer": float(value),
									"Durum": assessment["status"] if assessment else "Yaşa uygun kaynak aralığı yok",
									"Birim": assessment["unit"] if assessment else "",
									"Kaynak": assessment["source"] if assessment else ""})
			if assessments:
				st.dataframe(pd.DataFrame(assessments), width="stretch", hide_index=True)
				st.caption("‘Sınıra yakın’ yalnızca referans aralığının %10'luk kenar bandını ifade eden bir tarama işaretidir; klinik eşik değildir.")
				statuses = [row["Durum"] for row in assessments]
				if any(status.startswith("Kaynakta kritik") for status in statuses):
					st.error("Ölçüm, yüklenen kaynakta kritik işaretli aralığa ulaşıyor. Okul sağlık sorumlusuna ve veliye hemen bildirin; ciddi veya ani belirtide 112'yi arayın.")
				elif any(status in {"Referansın altında", "Referansın üstünde"} or status.startswith("Referansın altında") for status in statuses):
					st.warning("Bir veya daha fazla ölçüm referans aralığının dışında. Uygun yöntemle tekrar ölçün ve sağlık personeliyle paylaşın.")
				elif any("sınıra yakın" in status for status in statuses):
					st.info("Bir ölçüm referans sınırına yakın. Değişimi takip edin.")
			st.markdown("### Ölçüm geçmişi")
			st.dataframe(
				vitals.drop(columns=["id"]),
				width="stretch",
				hide_index=True,
				column_config={
					"Buyuk_Tansiyon": st.column_config.NumberColumn("Büyük Tansiyon (mmHg)", format="%.0f"),
					"Kucuk_Tansiyon": st.column_config.NumberColumn("Küçük Tansiyon (mmHg)", format="%.0f"),
					"Nabiz": st.column_config.NumberColumn("Nabız (/dk)", format="%.0f"),
					"Ates_C": st.column_config.NumberColumn("Ateş (°C)", format="%.1f"),
					"Saturasyon_SpO2": st.column_config.NumberColumn("Saturasyon - SpO2 (%)", format="%%%.0f"),
					"Kan_Sekeri": st.column_config.NumberColumn("Kan Şekeri (mg/dL)", format="%.0f"),
				},
			)

			with st.expander("🗑️ Hatalı Vital Kaydını Sil"):
				vital_choices = {
					f"{r['Tarih_Saat']} (Tansiyon: {r['Buyuk_Tansiyon'] or '—'}/{r['Kucuk_Tansiyon'] or '—'}, Nabız: {r['Nabiz'] or '—'}, Ateş: {r['Ates_C'] or '—'}, SpO2: {r['Saturasyon_SpO2'] or '—'}%)": int(r["id"])
					for _, r in vitals.iterrows()
				}
				chosen_vit_label = st.selectbox("Silinecek Vital Kaydı", list(vital_choices), key=f"sel_vit_del_{student['id']}")
				conf_vit_del = st.checkbox("Bu vital ölçümünü silmeyi onayla", key=f"conf_vit_del_{student['id']}")
				if st.button("Vital Kaydını Sil", type="secondary", disabled=not conf_vit_del, key=f"btn_vit_del_{student['id']}"):
					delete_vital(vital_choices[chosen_vit_label])
					st.success("Vital kaydı silindi.")
					st.rerun()
		else:
			st.caption("Bu öğrenci için henüz vital kaydı yok.")
		with st.expander("Yaşa uygun referans tablosunu yönet"):
			st.write("Evrensel tek bir çocuk vital aralığı yoktur. Yalnızca doğrulanmış klinik/kurumsal kaynağın yaş, gerekiyorsa cinsiyet ve birim eşleşen tablosunu yükleyin.")
			st.download_button("Referans şablonunu indir", make_vital_reference_template(), "vital_referans_sablonu.xlsx",
							 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
			reference_file = st.file_uploader("Kaynaklı referans tablosu", type=["xlsx", "csv", "pdf"], key="vital_reference_upload")
			if reference_file is not None and reference_file.name.lower().endswith(".pdf"):
				try:
					preview = parse_vital_reference_pdf(reference_file)
					st.markdown("#### PDF'den çıkarılan aralıklar")
					st.dataframe(preview, width="stretch", hide_index=True)
					st.caption("PDF metin satırları otomatik ayrıştırıldı. Yaş, ölçüm, birim ve kaynak dosyayla eşleştiğini doğrulamadan kaydetmeyin.")
				except Exception as error:
					st.warning(f"PDF'den otomatik aralık çıkarılamadı: {error}")
			if reference_file is not None and st.button("Referans tablosunu doğrula ve kaydet", key="save_vital_refs"):
				try:
					count, issues = parse_vital_reference_upload(reference_file)
					st.success(f"{count} yaşa uygun referans satırı kaydedildi.")
					for issue in issues[:8]:
						st.warning(issue)
				except Exception as error:
					st.error(f"Referans tablosu okunamadı: {error}")
	with lab_tab:
		st.write("PDF, Excel veya CSV raporu yükleyin. Bulunan sonuçları ve rapordaki referans sınırlarını kontrol edip onayladıktan sonra öğrenci dosyasına kaydedin.")
		uploaded = st.file_uploader("Laboratuvar raporu", type=["pdf", "xlsx", "csv"], key="lab_report_upload")
		if uploaded is not None:
			content = uploaded.getvalue()
			if len(content) > 20 * 1024 * 1024:
				st.error("Gizlilik ve performans için en fazla 20 MB dosya yükleyin.")
			else:
				signature = sha256(content).hexdigest()
				session_key = f"lab_draft_{student['id']}"
				signature_key = f"lab_signature_{student['id']}"
				if st.session_state.get(signature_key) != signature:
					try:
						draft, extracted = parse_lab_upload(uploaded)
						st.session_state[session_key] = draft if not draft.empty else pd.DataFrame(columns=lab_result_columns())
						st.session_state[f"lab_text_{student['id']}"] = extracted
						st.session_state[signature_key] = signature
					except Exception as error:
						st.error(f"Dosya okunamadı: {error}")
						st.session_state[signature_key] = None
				if st.session_state.get(signature_key) == signature:
					extracted = st.session_state.get(f"lab_text_{student['id']}", "")
					if uploaded.name.lower().endswith(".pdf") and not extracted.strip():
						st.warning("PDF'den metin veya OCR sonucu çıkarılamadı. Sonuç satırlarını aşağıya elle girin.")
					elif uploaded.name.lower().endswith(".pdf"):
						st.caption("PDF metni çıkarıldı. Otomatik bulunan satırları kaydetmeden önce raporla karşılaştırın.")
						with st.expander("PDF'den çıkarılan metni gör"):
							st.text(extracted[:12000])
					draft = st.session_state[session_key]
					edited = st.data_editor(
						draft, num_rows="dynamic", width="stretch", key=f"lab_editor_{student['id']}",
						column_config={
							"Test": st.column_config.TextColumn("Test adı", required=True),
							"Sonuç": st.column_config.NumberColumn("Sonuç", format="%.4f"),
							"Birim": st.column_config.TextColumn("Birim"),
							"Referans alt": st.column_config.NumberColumn("Referans alt", format="%.4f"),
							"Referans üst": st.column_config.NumberColumn("Referans üst", format="%.4f"),
							"Laboratuvar işareti": st.column_config.TextColumn("Rapor işareti (H/L vb.)"),
						},
					)
				report_date = st.date_input("Tahlil tarihi", value=date.today(), key=f"lab_date_{student['id']}")
				if st.button("Sonuçları öğrenci dosyasına kaydet", type="primary", key=f"save_lab_{student['id']}"):
					rows = []
					invalid = False
					for _, row in edited.iterrows():
						name, value = clean_text(row.get("Test")), numeric_value(row.get("Sonuç"))
						if not name and value is None:
							continue
						low, high = numeric_value(row.get("Referans alt")), numeric_value(row.get("Referans üst"))
						if not name or value is None or ((low is None) != (high is None)) or (low is not None and high <= low):
							invalid = True
							break
						rows.append((name, value, clean_text(row.get("Birim")) or None, low, high,
									 clean_text(row.get("Laboratuvar işareti")) or None))
					if invalid or not rows:
						st.error("Her sonuç için test adı ve sayısal değer girin; referans alt/üst sınırları birlikte ve tutarlı olmalı.")
					else:
						content_type = "application/pdf" if uploaded.name.lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if uploaded.name.lower().endswith(".xlsx") else "text/csv"
						with connect_db() as connection:
							cursor = connection.execute(
								"INSERT INTO lab_reports(student_id, report_date, file_name, content_type, content, extracted_text) VALUES (?, ?, ?, ?, ?, ?)",
								(student["id"], report_date.isoformat(), Path(uploaded.name).name, content_type, content,
								 st.session_state.get(f"lab_text_{student['id']}") or None),
							)
							connection.executemany(
								"INSERT INTO lab_results(report_id, test_name, value, unit, ref_low, ref_high, lab_flag) VALUES (?, ?, ?, ?, ?, ?, ?)",
								[(cursor.lastrowid, *row) for row in rows],
							)
						save_person_data_file(student["id"])
						st.success(f"{len(rows)} sonuç ve kaynak dosya öğrenci dosyasına kaydedildi.")
						st.rerun()
		with connect_db() as connection:
			reports = connection.execute(
				"SELECT id, report_date, file_name, content, content_type FROM lab_reports WHERE student_id=? ORDER BY report_date DESC, id DESC",
				(student["id"],),
			).fetchall()
			lab_history = pd.read_sql_query(
				"""SELECT r.report_date AS Tarih, r.file_name AS Dosya, x.test_name AS Test,
				   x.value AS Sonuç, x.unit AS Birim, x.ref_low AS Referans_Alt, x.ref_high AS Referans_Üst,
				   x.lab_flag AS Rapor_İşareti FROM lab_results x JOIN lab_reports r ON r.id=x.report_id
				   WHERE r.student_id=? ORDER BY r.report_date DESC, x.id""", connection, params=(student["id"],),
			)
		if not lab_history.empty:
			lab_history = assess_lab_history(lab_history)
			st.markdown("### Kayıtlı sonuçlar")
			st.dataframe(lab_history, width="stretch", hide_index=True)
			st.caption("Değerlendirme rapordaki referans sınırıyla sayısal karşılaştırmadır; tanı değildir. Rapor işaretleri ayrıca kontrol edilmelidir.")
			statuses = lab_history["Değerlendirme"].tolist()
			if any(status.startswith("Referansın") for status in statuses):
				st.warning("Sonuçlardan en az biri raporun referans aralığı dışında. Aile ve sağlık uzmanıyla paylaşın.")
			elif any("sınıra yakın" in status for status in statuses):
				st.info("Bir sonuç rapor aralığının sınırına yakın. Laboratuvarın referans notlarını esas alın.")
		if reports:
			st.markdown("### Kaynak dosyalar")
			for report in reports:
				st.download_button(f"{report['report_date']} · {report['file_name']}", bytes(report["content"]),
							   file_name=report["file_name"], mime=report["content_type"], key=f"download_lab_{report['id']}")


def render_measurements() -> None:
	st.markdown("## Boy, kilo ve gelişim")
	student = select_shared_person("meas", label="Öğrenci veya kişi seçin")
	if not student:
		st.caption("Ölçüm kaydı, Z-skoru (SDS) ve persantil analizi için kişi seçin.")
		return

	render_shared_person_summary(student, section_key="meas")

	columns = st.columns(4)
	measured_on = columns[0].date_input("Ölçüm tarihi", value=date.today(), key=f"new_measurement_date_{student['id']}")
	height = columns[1].number_input("Boy (cm)", min_value=0.0, max_value=250.0, value=None, step=0.1, format="%.1f", key=f"new_measurement_height_{student['id']}")
	weight = columns[2].number_input("Kilo (kg)", min_value=0.0, max_value=300.0, value=None, step=0.1, format="%.1f", key=f"new_measurement_weight_{student['id']}")
	head_circ = columns[3].number_input("Baş Çevresi (cm, opsiyonel)", min_value=0.0, max_value=100.0, value=None, step=0.1, format="%.1f", key=f"new_measurement_hc_{student['id']}")

	# Instant live calculation preview before saving
	if (height is not None and weight is not None and height > 0 and weight > 0) or (head_circ is not None and head_circ > 0):
		preview_months = age_in_months(student["birth_date"], measured_on) if student["birth_date"] else None
		p_items = []
		if preview_months is not None:
			p_items.append(f"**Ölçüm Yaşı:** {preview_months // 12} yaş {preview_months % 12} ay ({preview_months} ay)")
		if height is not None and weight is not None and height > 0 and weight > 0:
			preview_bmi = weight / ((height / 100.0) ** 2)
			p_items.append(f"**VKİ (BMI):** {preview_bmi:.1f} kg/m²")
			prev_ideal = calculate_ideal_weight(height, student["sex"], preview_months, current_weight_kg=weight)
			if prev_ideal:
				p_items.append(f"**🎯 İdeal Kilo Aralığı:** {prev_ideal['min_kg']:.1f} – {prev_ideal['max_kg']:.1f} kg (İdeal En Az: {prev_ideal['min_kg']:.1f} kg · İdeal En Çok: {prev_ideal['max_kg']:.1f} kg · {prev_ideal['status_text']})")
			if student["sex"] and preview_months is not None:
				ph_ev = evaluate_measurement(height, "height", student["sex"], preview_months)
				pw_ev = evaluate_measurement(weight, "weight", student["sex"], preview_months)
				pb_ev = evaluate_measurement(preview_bmi, "bmi", student["sex"], preview_months)
				if ph_ev:
					p_items.append(f"**Boy:** {ph_ev['z']:+.2f} SDS (%{ph_ev['percentile']:.1f}p · {ph_ev['status']})")
				if pw_ev:
					p_items.append(f"**Kilo:** {pw_ev['z']:+.2f} SDS (%{pw_ev['percentile']:.1f}p · {pw_ev['status']})")
				if pb_ev:
					p_items.append(f"**VKİ:** {pb_ev['z']:+.2f} SDS (%{pb_ev['percentile']:.1f}p · {pb_ev['status']})")
		if head_circ is not None and head_circ > 0:
			phc_ev = evaluate_measurement(head_circ, "head_circ", student["sex"], preview_months) if (student["sex"] and preview_months is not None) else None
			if phc_ev:
				p_items.append(f"**Baş Çevresi:** {head_circ:.1f} cm · {phc_ev['z']:+.2f} SDS (%{phc_ev['percentile']:.1f}p · {phc_ev['status']})")
			else:
				p_items.append(f"**Baş Çevresi:** {head_circ:.1f} cm")
		st.info("⚡ **Anlık Persantil & Z-Skor Önizlemesi:** " + "  |  ".join(p_items))

	save = st.button("Ölçümü kaydet", type="primary", key=f"save_measurement_{student['id']}")
	if save:
		has_hw_in = height is not None and weight is not None and height > 0 and weight > 0
		has_hc_in = head_circ is not None and head_circ > 0
		if (height is not None and height > 0) != (weight is not None and weight > 0):
			st.error("Boy ve kilo ölçümünü birlikte girin.")
		elif not has_hw_in and not has_hc_in:
			st.error("Kaydetmek için boy-kilo veya baş çevresi ölçümü girin.")
		else:
			summary_txt = record_linked_measurement(
				student["id"], measured_on,
				height if has_hw_in else None,
				weight if has_hw_in else None,
				student["birth_date"], student["sex"],
				head_circ_cm=head_circ if has_hc_in else None,
			)
			st.success(f"Ölçüm öğrenci profiline eklendi: {summary_txt}")
			st.rerun()

	with connect_db() as connection:
		records = pd.read_sql_query(
			"""SELECT id, measured_on AS Tarih, height_cm AS Boy_cm, weight_kg AS Kilo_kg,
					  head_circ_cm AS Bas_Cevresi_cm, age_months AS Yas_ay, bmi AS BMI
			   FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC""",
			connection, params=(student["id"],),
		)
	if records.empty:
		st.info("Henüz ölçüm yok. Yeni bir ölçüm girerek persantil gelişimini takip edebilirsiniz.")
		return

	records_by_id = {int(row["id"]): row for _, row in records.iterrows()}
	record_ids_desc = list(reversed(list(records_by_id.keys())))
	active_meas_key = f"active_meas_id_{student['id']}"
	if st.session_state.get(active_meas_key) not in records_by_id:
		st.session_state[active_meas_key] = record_ids_desc[0]

	active_meas_id = st.session_state[active_meas_key]
	latest = records_by_id[active_meas_id]
	months = int(latest["Yas_ay"]) if not pd.isna(latest["Yas_ay"]) else (
		age_in_months(student["birth_date"], date.fromisoformat(str(latest["Tarih"])[:10])) if student["birth_date"] else None
	)
	age_display = f"{months // 12} yaş {months % 12} ay ({months} ay)" if months is not None else "Doğum tarihi eksik"
	sex_display = "Kız Öğrenci ♀" if student["sex"] == "K" else ("Erkek Öğrenci ♂" if student["sex"] == "E" else "Cinsiyet belirtilmedi")

	latest_h = float(latest["Boy_cm"]) if pd.notna(latest["Boy_cm"]) else None
	latest_w = float(latest["Kilo_kg"]) if pd.notna(latest["Kilo_kg"]) else None
	latest_b = float(latest["BMI"]) if pd.notna(latest["BMI"]) else None
	latest_hc = float(latest["Bas_Cevresi_cm"]) if (pd.notna(latest.get("Bas_Cevresi_cm")) and float(latest["Bas_Cevresi_cm"]) > 0) else None

	person_meas = get_person_latest_measurements(student["id"])
	display_hc = latest_hc if latest_hc is not None else person_meas["head_circ_cm"]
	h_eval = evaluate_measurement(latest_h, "height", student["sex"], months) if (student["sex"] and months is not None and latest_h) else None
	w_eval = evaluate_measurement(latest_w, "weight", student["sex"], months) if (student["sex"] and months is not None and latest_w) else None
	b_eval = evaluate_measurement(latest_b, "bmi", student["sex"], months) if (student["sex"] and months is not None and latest_b) else None
	hc_eval = evaluate_measurement(display_hc, "head_circ", student["sex"], months) if (student["sex"] and months is not None and display_hc) else None

	st.markdown(f"#### 👤 {student['full_name']} <small style='color:#64748b;'>· {student['class_name']} · {sex_display} · {age_display}</small>", unsafe_allow_html=True)

	can_eval_age_m = (months is not None and 0 <= months <= 228)
	can_eval_hc_m = (months is not None and 0 <= months <= 60)

	ideal_meas = calculate_ideal_weight(latest_h, student["sex"], months, current_weight_kg=latest_w)

	if can_eval_hc_m:
		card_c1, card_c2, card_c3, card_c4, card_c5, card_c6 = st.columns(6)
		with card_c1:
			st.metric("📏 Boy", f"{latest_h:.1f} cm" if latest_h else "—", delta=f"{h_eval['z']:+.2f} SDS (%{h_eval['percentile']:.1f}p)" if h_eval else None)
			st.caption(f"Tarama: **{h_eval['status']}**" if h_eval else "Referans yok")
		with card_c2:
			st.metric("⚖️ Kilo", f"{latest_w:.1f} kg" if latest_w else "—", delta=f"{w_eval['z']:+.2f} SDS (%{w_eval['percentile']:.1f}p)" if w_eval else None)
			st.caption(f"Tarama: **{w_eval['status']}**" if w_eval else "Referans yok")
		with card_c3:
			st.metric("🧬 VKİ (BMI)", f"{latest_b:.1f} kg/m²" if latest_b else "—", delta=f"{b_eval['z']:+.2f} SDS (%{b_eval['percentile']:.1f}p)" if b_eval else None)
			st.caption(f"Tarama: **{b_eval['status']}**" if b_eval else "Referans yok")
		with card_c4:
			rng_val = f"{ideal_meas['min_kg']:.1f} – {ideal_meas['max_kg']:.1f} kg" if ideal_meas else "—"
			st.metric("🎯 İdeal Kilo Aralığı", rng_val, delta=ideal_meas['status_text'] if ideal_meas else None)
			if ideal_meas:
				st.caption(f"İdeal En Az: **{ideal_meas['min_kg']:.1f} kg** · İdeal En Çok: **{ideal_meas['max_kg']:.1f} kg**")
			else:
				st.caption("Hesaplanamadı")
		with card_c5:
			st.metric("🧠 Baş Çevresi", f"{display_hc:.1f} cm" if display_hc else "—", delta=f"{hc_eval['z']:+.2f} SDS (%{hc_eval['percentile']:.1f}p)" if hc_eval else None)
			if display_hc:
				if latest_hc is not None: st.caption(f"Tarama: **{hc_eval['status']}**" if hc_eval else "Kaydedildi")
				else: st.caption(f"Önceki Kayıttan ({person_meas['head_circ_date']}) · **{hc_eval['status'] if hc_eval else 'Kaydedildi'}**")
			else: st.caption("Girilmedi")
		with card_c6:
			st.metric("📅 Ölçüm Tarihi", str(latest["Tarih"]))
			st.caption(f"Kayıtlı toplam ölçüm: **{len(records)} adet**")
	else:
		card_c1, card_c2, card_c3, card_c4, card_c5 = st.columns(5)
		with card_c1:
			st.metric("📏 Boy", f"{latest_h:.1f} cm" if latest_h else "—", delta=f"{h_eval['z']:+.2f} SDS (%{h_eval['percentile']:.1f}p)" if (h_eval and can_eval_age_m) else None)
			st.caption(f"Tarama: **{h_eval['status']}**" if (h_eval and can_eval_age_m) else "Kaydedildi")
		with card_c2:
			st.metric("⚖️ Kilo", f"{latest_w:.1f} kg" if latest_w else "—", delta=f"{w_eval['z']:+.2f} SDS (%{w_eval['percentile']:.1f}p)" if (w_eval and can_eval_age_m) else None)
			st.caption(f"Tarama: **{w_eval['status']}**" if (w_eval and can_eval_age_m) else "Kaydedildi")
		with card_c3:
			st.metric("🧬 VKİ (BMI)", f"{latest_b:.1f} kg/m²" if latest_b else "—", delta=f"{b_eval['z']:+.2f} SDS (%{b_eval['percentile']:.1f}p)" if (b_eval and can_eval_age_m) else None)
			st.caption(f"Tarama: **{b_eval['status']}**" if (b_eval and can_eval_age_m) else "Hesaplandı")
		with card_c4:
			rng_val = f"{ideal_meas['min_kg']:.1f} – {ideal_meas['max_kg']:.1f} kg" if ideal_meas else "—"
			st.metric("🎯 İdeal Kilo Aralığı", rng_val, delta=ideal_meas['status_text'] if ideal_meas else None)
			if ideal_meas:
				st.caption(f"İdeal En Az: **{ideal_meas['min_kg']:.1f} kg** · İdeal En Çok: **{ideal_meas['max_kg']:.1f} kg**")
			else:
				st.caption("Hesaplanamadı")
		with card_c5:
			st.metric("📅 Ölçüm Tarihi", str(latest["Tarih"]))
			st.caption(f"Kayıtlı toplam ölçüm: **{len(records)} adet**")

	tab_report, tab_cutoffs, tab_history, tab_charts = st.tabs([
		"🏥 Klinik Büyüme & Z-Skor Raporu",
		"📋 WHO Persentil & SDS Referans Cetveli",
		"📅 Kronolojik Ölçüm Geçmişi & Değişim",
		"📈 Büyüme Eğrileri ve Takip Grafiği"
	])

	with tab_report:
		st.markdown("##### 🩺 Antropometrik Z-Skoru (SDS) ve Persantil Değerlendirme Tablosu")
		st.caption("💡 Tablodaki **'✏️ Ölçülen Değer (Düzenlenebilir)'** sütununda **Boy (cm)**, **Kilo (kg)** veya **Baş Çevresi (cm)** hücresine çift tıklayarak değeri doğrudan değiştirebilirsiniz; VKİ, Z-Skoru (SDS) ve Persantil otomatik güncellenir.")

		def _fmt_meas_opt(rid: int) -> str:
			rr = records_by_id[rid]
			h_s = f"{rr['Boy_cm']:.1f} cm" if pd.notna(rr["Boy_cm"]) else "—"
			w_s = f"{rr['Kilo_kg']:.1f} kg" if pd.notna(rr["Kilo_kg"]) else "—"
			hc_s = f" · Baş Çevresi: {rr['Bas_Cevresi_cm']:.1f} cm" if (pd.notna(rr.get("Bas_Cevresi_cm")) and float(rr["Bas_Cevresi_cm"]) > 0) else ""
			return f"{rr['Tarih']} · Boy: {h_s} · Kilo: {w_s}{hc_s}" + (" (Son Ölçüm)" if rid == record_ids_desc[0] else "")

		ctrl_c1, ctrl_c2, ctrl_c3 = st.columns([2.2, 1.2, 0.8])
		with ctrl_c1:
			if len(record_ids_desc) > 1:
				chosen_rec_id = st.selectbox(
					"Tabloda görüntülenen / düzenlenen ölçüm kaydı",
					record_ids_desc,
					index=record_ids_desc.index(active_meas_id),
					format_func=_fmt_meas_opt,
					key=f"sel_active_meas_{student['id']}",
				)
				if chosen_rec_id != active_meas_id:
					st.session_state[active_meas_key] = chosen_rec_id
					st.rerun()
			else:
				st.caption(f"📌 Kayıtlı Ölçüm: **{_fmt_meas_opt(active_meas_id)}**")
		with ctrl_c2:
			orig_date = date.fromisoformat(str(latest["Tarih"])[:10])
			edited_date = st.date_input("Ölçüm Tarihi", value=orig_date, key=f"inline_meas_date_{active_meas_id}")
			if edited_date != orig_date:
				new_months = age_in_months(student["birth_date"], edited_date) if student["birth_date"] else None
				with connect_db() as connection:
					connection.execute(
						"UPDATE measurements SET measured_on=?, age_months=? WHERE id=? AND student_id=?",
						(edited_date.isoformat(), new_months, active_meas_id, student["id"]),
					)
				save_person_data_file(student["id"])
				st.success("Ölçüm tarihi ve yaş persantil hesabı güncellendi.")
				st.rerun()
		with ctrl_c3:
			st.write("")
			if st.button("🗑️ Bu Ölçümü Sil", type="secondary", use_container_width=True, key=f"inline_del_meas_{active_meas_id}"):
				delete_measurement(active_meas_id)
				st.session_state.pop(active_meas_key, None)
				st.success("Ölçüm kaydı silindi.")
				st.rerun()

		clinical_summary_df = get_clinical_growth_summary_table(
			latest_h, latest_w, latest_b, student["sex"], months, head_circ_cm=display_hc, always_include_head_circ=True
		)
		edited_clinical_df = st.data_editor(
			clinical_summary_df,
			width="stretch",
			hide_index=True,
			column_config=get_clinical_growth_column_config(),
			disabled=[
				"Antropometrik Parametre",
				"50. Persentil (Medyan)",
				"Normal Referans Bandı (-2 ile +2 SDS)",
				"Z-Skoru (SDS)",
				"Persentil (%)",
				"Persentil Dilimi",
				"Klinik Değerlendirme",
			],
			key=f"clinical_growth_editor_{student['id']}_{active_meas_id}",
		)

		if not edited_clinical_df.empty:
			param_map = {str(r["Antropometrik Parametre"]): numeric_value(r["Ölçülen Değer"]) for _, r in edited_clinical_df.iterrows()}
			new_h_val = param_map.get("📏 Boy (cm)", latest_h)
			new_w_val = param_map.get("⚖️ Kilo (kg)", latest_w)
			new_hc_val = param_map.get("🧠 Baş Çevresi (cm)", display_hc)
			orig_h_val = round(float(latest_h), 1) if latest_h is not None else None
			orig_w_val = round(float(latest_w), 1) if latest_w is not None else None
			orig_hc_val = round(float(display_hc), 1) if display_hc is not None else None

			h_changed = (new_h_val is not None) != (orig_h_val is not None) or (new_h_val is not None and orig_h_val is not None and abs(new_h_val - orig_h_val) > 1e-4)
			w_changed = (new_w_val is not None) != (orig_w_val is not None) or (new_w_val is not None and orig_w_val is not None and abs(new_w_val - orig_w_val) > 1e-4)
			hc_changed = (new_hc_val is not None) != (orig_hc_val is not None) or (new_hc_val is not None and orig_hc_val is not None and abs(new_hc_val - orig_hc_val) > 1e-4)

			if h_changed or w_changed or hc_changed:
				if (new_h_val is not None and new_h_val <= 0) or (new_w_val is not None and new_w_val <= 0):
					st.error("Boy ve kilo sıfırdan büyük olmalıdır.")
				else:
					new_bmi = (new_w_val / ((new_h_val / 100.0) ** 2)) if (new_h_val and new_w_val and new_h_val > 0 and new_w_val > 0) else None
					new_months = age_in_months(student["birth_date"], orig_date) if student["birth_date"] else None
					clean_hc = float(new_hc_val) if (new_hc_val is not None and new_hc_val > 0) else None
					with connect_db() as connection:
						connection.execute(
							"UPDATE measurements SET height_cm=?, weight_kg=?, head_circ_cm=?, age_months=?, bmi=? WHERE id=? AND student_id=?",
							(float(new_h_val) if new_h_val else None, float(new_w_val) if new_w_val else None, clean_hc, new_months, float(new_bmi) if new_bmi else None, active_meas_id, student["id"]),
						)
					save_person_data_file(student["id"])
					st.success("Tablo üzerinden ölçüm güncellendi.")
					st.rerun()

		if student["sex"] and months is not None:
			issues = []
			if h_eval and ("kısalığı" in h_eval["status"].lower() or "bodurluk" in h_eval["status"].lower()):
				issues.append(f"📏 **Boy Gelişimi:** Öğrencinin boyu yaş grubuna göre -2 SDS (3. persentil) altında kalmaktadır ({h_eval['status']}). Büyüme hızının takip edilmesi, aile hekimi ve çocuk endokrinolojisi uzmanı görüşü alınması önerilir.")
			if b_eval and ("obezite" in b_eval["status"].lower() or "fazla kilo" in b_eval["status"].lower()):
				issues.append(f"🧬 **Ağırlık & VKİ:** Öğrencinin Vücut Kitle İndeksi fazla kiloluluk/obezite diliminde yer almaktadır ({b_eval['status']}). Sağlıklı beslenme ve fiziksel aktivite takibi tavsiye edilir.")
			elif b_eval and "zayıflık" in b_eval["status"].lower():
				issues.append(f"🧬 **Ağırlık & VKİ:** Öğrencinin Vücut Kitle İndeksi zayıflık sınırındadır ({b_eval['status']}). Beslenme desteği ve takibi gerekebilir.")
			if hc_eval and "Normal" not in hc_eval["status"]:
				issues.append(f"🧠 **Baş Çevresi:** Öğrencinin baş çevresi yaş referans sınırları dışındadır ({hc_eval['status']}).")

			if issues:
				st.warning("\n\n".join(issues))
			else:
				st.success("✅ **Klinik Değerlendirme:** Öğrencinin antropometrik ölçümleri Dünya Sağlık Örgütü (WHO) yaş ve cinsiyet referans aralığında (-2 ile +2 SDS / 3. ile 97. persentil arasında) sağlıklı ilerlemektedir.")
			st.caption("ℹ️ **Z-Skoru (SDS):** Ölçümün toplum ortalamasından (medyandan) kaç standart sapma uzakta olduğunu gösterir. **Persentil (%p):** Aynı yaş ve cinsiyetteki 100 çocuk arasındaki yüzdelik sırayı ifade eder (50p tam medyan/ortalamadır; 3p-97p arası normal kabul edilir).")
		else:
			w_c1, w_c2 = st.columns([3.2, 1])
			with w_c1:
				st.warning("Persantil ve Z-skoru değerlendirmesi için öğrenci profilinde doğum tarihi ve cinsiyet bulunmalıdır.")
			with w_c2:
				if st.button("✏️ Doğum Tarihi / Cinsiyet Ekle", type="primary", key=f"meas_quick_edit_{student['id']}", use_container_width=True):
					st.session_state["edit_profile_loaded_id"] = None
					open_edit_profile_dialog(student["id"], "meas_q")

	with tab_cutoffs:
		if student["sex"] and months is not None:
			st.markdown(f"##### 📋 {months // 12} Yaş {months % 12} Aylık {'Kız' if student['sex'] == 'K' else 'Erkek'} Çocuklar İçin Resmi WHO Referans Sınır Cetveli")
			st.caption("Bu tablo, Dünya Sağlık Örgütü (WHO) standartlarına göre çocuğun tam yaşı ve cinsiyetindeki resmi istatistiksel sınırları ve öğrencinin bu sınırlara göre yerini gösterir:")
			cutoffs_df = get_who_reference_cutoffs_table(
				student["sex"], months, latest_h, latest_w, latest_b, head_circ_cm=display_hc
			)
			st.dataframe(
				cutoffs_df,
				width="stretch",
				hide_index=True,
				column_config=get_who_cutoffs_column_config(),
			)
			st.markdown("""
			- **-3 SDS (0.1p):** Belirgin/ağır düşük sınır
			- **-2 SDS (3p):** Normalin alt klinik eşiği (Bodurluk / Düşük kilo alt sınırı)
			- **0 SDS (50p):** Toplum medyanı / ideal ortalama
			- **+1 SDS (85p):** Hafif kilolu / risk başlangıcı
			- **+2 SDS (97p):** Normalin üst klinik eşiği (Fazla kiloluluk / Obezite üst sınırı)
			- **+3 SDS (99.9p):** İleri derece yüksek sınır
			""")
		else:
			st.warning("Referans cetveli için öğrenci profilinde doğum tarihi ve cinsiyet bulunmalıdır.")

	enriched_records = enrich_measurements_with_percentiles(records, student["sex"], student["birth_date"])

	with tab_history:
		st.markdown("##### 📅 Kronolojik Ölçüm Geçmişi ve Gelişim Hızı Çizelgesi")
		display_df = enriched_records[[
			"id", "Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Boy_Dilimi", "Boy_Farki",
			"Kilo_kg", "Kilo_Z", "Kilo_Persantil", "Kilo_Dilimi", "Kilo_Farki",
			"Ideal_Kilo_Araligi", "Ideal_En_Az_kg", "Ideal_En_Cok_kg", "Ideal_Kilo_kg", "Ideal_Kilo_Farki",
			"Bas_Cevresi_cm", "Bas_Cevresi_Z", "Bas_Cevresi_Persantil",
			"BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"
		]].rename(columns={
			"Yas_ay": "Yaş (Ay)",
			"Boy_cm": "Boy (cm)",
			"Boy_Z": "Boy Z-Skor",
			"Boy_Persantil": "Boy %",
			"Boy_Dilimi": "Boy Dilimi",
			"Boy_Farki": "Boy Değişimi",
			"Kilo_kg": "Kilo (kg)",
			"Kilo_Z": "Kilo Z-Skor",
			"Kilo_Persantil": "Kilo %",
			"Kilo_Dilimi": "Kilo Dilimi",
			"Kilo_Farki": "Kilo Değişimi",
			"Ideal_Kilo_Araligi": "İdeal Kilo Aralığı",
			"Ideal_En_Az_kg": "İdeal En Az (kg)",
			"Ideal_En_Cok_kg": "İdeal En Çok (kg)",
			"Ideal_Kilo_kg": "İdeal Kilo (kg)", "Ideal_Kilo_Araligi": "İdeal Kilo Aralığı",
			"Ideal_Kilo_Farki": "İdeal Kilo Farkı",
			"Bas_Cevresi_cm": "Baş Çevresi (cm)",
			"Bas_Cevresi_Z": "Baş Çevresi Z-Skor",
			"Bas_Cevresi_Persantil": "Baş Çevresi %",
			"BMI": "VKİ (BMI)",
			"BMI_Z": "VKİ Z-Skor",
			"BMI_Persantil": "VKİ %",
			"Gelisim_Degerlendirmesi": "Gelişim Değerlendirmesi"
		})
		hist_col_cfg = get_measurement_history_column_config()
		hist_col_cfg["id"] = None
		edited_hist_df = st.data_editor(
			display_df,
			width="stretch",
			hide_index=True,
			num_rows="dynamic",
			column_config=hist_col_cfg,
			disabled=[
				"Yaş (Ay)", "Boy Z-Skor", "Boy %", "Boy Dilimi", "Boy Değişimi",
				"Kilo Z-Skor", "Kilo %", "Kilo Dilimi", "Kilo Değişimi",
				"İdeal Kilo Aralığı", "İdeal En Az (kg)", "İdeal En Çok (kg)", "İdeal Kilo (kg)", "İdeal Kilo Farkı",
				"Baş Çevresi Z-Skor", "Baş Çevresi %",
				"VKİ (BMI)", "VKİ Z-Skor", "VKİ %", "Gelişim Değerlendirmesi"
			],
			key=f"history_editor_{student['id']}",
		)
		st.caption("💡 Bu tabloda da **Tarih**, **Boy (cm)**, **Kilo (kg)** ve **Baş Çevresi (cm)** hücrelerini doğrudan düzenleyebilir veya satır seçip silebilirsiniz. Tüm Z-skoru ve persantil sütunları sayısal olarak birbirine bağlıdır.")

		# Detect inline edits or row deletions in tab_history
		orig_ids = set(display_df["id"].dropna().astype(int).tolist())
		curr_ids = set(edited_hist_df["id"].dropna().astype(int).tolist()) if "id" in edited_hist_df.columns else set()
		deleted_ids = orig_ids - curr_ids
		changed_rows = []
		if "id" in edited_hist_df.columns:
			for _, erow in edited_hist_df.iterrows():
				rid = erow.get("id")
				if pd.isna(rid) or int(rid) not in records_by_id:
					continue
				rid_int = int(rid)
				orig_r = records_by_id[rid_int]
				nh = numeric_value(erow.get("Boy (cm)"))
				nw = numeric_value(erow.get("Kilo (kg)"))
				nhc = numeric_value(erow.get("Baş Çevresi (cm)"))
				nt = str(erow.get("Tarih") or "").strip()[:10]
				orig_h = float(orig_r["Boy_cm"]) if pd.notna(orig_r["Boy_cm"]) else None
				orig_w = float(orig_r["Kilo_kg"]) if pd.notna(orig_r["Kilo_kg"]) else None
				orig_hc = float(orig_r["Bas_Cevresi_cm"]) if (pd.notna(orig_r.get("Bas_Cevresi_cm")) and float(orig_r["Bas_Cevresi_cm"]) > 0) else None
				if nt:
					h_diff = (nh is not None) != (orig_h is not None) or (nh is not None and orig_h is not None and abs(nh - orig_h) > 1e-4)
					w_diff = (nw is not None) != (orig_w is not None) or (nw is not None and orig_w is not None and abs(nw - orig_w) > 1e-4)
					hc_diff = (nhc is not None) != (orig_hc is not None) or (nhc is not None and orig_hc is not None and abs(nhc - orig_hc) > 1e-4)
					if h_diff or w_diff or hc_diff or nt != str(orig_r["Tarih"])[:10]:
						changed_rows.append((rid_int, nt, nh, nw, nhc))
		if deleted_ids or changed_rows:
			with connect_db() as connection:
				for did in deleted_ids:
					connection.execute("DELETE FROM measurements WHERE id=? AND student_id=?", (int(did), student["id"]))
				for rid_int, nt, nh, nw, nhc in changed_rows:
					try:
						dt_obj = date.fromisoformat(nt)
					except Exception:
						dt_obj = date.today()
					m_ay = age_in_months(student["birth_date"], dt_obj) if student["birth_date"] else None
					nbmi = (nw / ((nh / 100.0) ** 2)) if (nh and nw and nh > 0 and nw > 0) else None
					connection.execute(
						"UPDATE measurements SET measured_on=?, height_cm=?, weight_kg=?, head_circ_cm=?, age_months=?, bmi=? WHERE id=? AND student_id=?",
						(dt_obj.isoformat(), float(nh) if nh else None, float(nw) if nw else None, float(nhc) if (nhc and nhc > 0) else None, m_ay, nbmi, rid_int, student["id"]),
					)
			save_person_data_file(student["id"])
			st.success("Tablo üzerindeki ölçüm değişiklikleri kaydedildi.")
			st.rerun()

	with tab_charts:
		st.markdown("##### 📈 Klinik WHO Büyüme Eğrileri ve Persantil Gelişim Takip Grafiği")
		st.caption("🩺 **Klinik Büyüme Eğrileri Mantığı:** Büyüme eğrisi, çocuğun boy, kilo veya VKİ ölçümünü doğrudan yaşa göre Dünya Sağlık Örgütü (WHO) persantil kanalları (**3p, 15p, 50p, 85p, 97p**) üzerinde gösterir. Böylece çocuğun kendi genetik persantil kanalında kalıp kalmadığı net olarak değerlendirilir.")

		if not can_eval_age_m:
			if not stu_birth:
				st.warning("⚠️ Seçili öğrencinin doğum tarihi girilmediği için yaşa göre büyüme eğrisi ve Z-skorları değerlendirilememektedir. Öğrenci düzenleme formundan doğum tarihini kaydedebilirsiniz.")
			else:
				st.info("ℹ️ Kişi 19 yaşından büyük (yetişkin) olduğu için çocukluk çağı WHO büyüme eğrileri ve persantil kanalları uygulanamamaktadır.")
		else:
			has_sex = student["sex"] in {"K", "E"}
			sex_label = "Kız Çocuk" if student["sex"] == "K" else ("Erkek Çocuk" if student["sex"] == "E" else "Cinsiyet Yok")

			def _render_clinical_growth_curve(indicator_name: str, y_col: str, title: str, unit: str, default_max_mo: int = 72):
				render_clinical_growth_curve(enriched_records, student, indicator_name, y_col, title, unit, default_max_mo)

			if can_eval_hc_m:
				chart_tabs = st.tabs([
					"📏 Yaşa Göre Boy Eğrisi (WHO LHFA/HFA)",
					"⚖️ Yaşa Göre Kilo Eğrisi (WHO WFA)",
					"🧬 Yaşa Göre VKİ Eğrisi (WHO BMIFA)",
					"🧠 Yaşa Göre Baş Çevresi Eğrisi (WHO HCFA)",
					"📊 Kronolojik Persantil (%p) Seyri",
				])
				with chart_tabs[0]:
					st.markdown(f"###### 📏 Yaşa Göre Boy Persantil Eğrisi (0 - 228 Ay) · {sex_label}")
					_render_clinical_growth_curve("height", "Boy_cm", "Boy", "cm", default_max_mo=228)
				with chart_tabs[1]:
					st.markdown(f"###### ⚖️ Yaşa Göre Kilo Persantil Eğrisi (0 - 120 Ay) · {sex_label}")
					_render_clinical_growth_curve("weight", "Kilo_kg", "Kilo", "kg", default_max_mo=120)
				with chart_tabs[2]:
					st.markdown(f"###### 🧬 Yaşa Göre VKİ (Beden Kitle İndeksi) Persantil Eğrisi · {sex_label}")
					_render_clinical_growth_curve("bmi", "BMI", "VKİ (BMI)", "kg/m²", default_max_mo=228)
				with chart_tabs[3]:
					st.markdown(f"###### 🧠 Yaşa Göre Baş Çevresi Persantil Eğrisi (0 - 60 Ay) · {sex_label}")
					_render_clinical_growth_curve("head_circ", "Bas_Cevresi_cm", "Baş Çevresi", "cm", default_max_mo=60)
				tab_chrono_main = chart_tabs[4]
			else:
				chart_tabs = st.tabs([
					"📏 Yaşa Göre Boy Eğrisi (WHO LHFA/HFA)",
					"⚖️ Yaşa Göre Kilo Eğrisi (WHO WFA)",
					"🧬 Yaşa Göre VKİ Eğrisi (WHO BMIFA)",
					"📊 Kronolojik Persantil (%p) Seyri",
				])
				with chart_tabs[0]:
					st.markdown(f"###### 📏 Yaşa Göre Boy Persantil Eğrisi (0 - 228 Ay) · {sex_label}")
					_render_clinical_growth_curve("height", "Boy_cm", "Boy", "cm", default_max_mo=228)
				with chart_tabs[1]:
					st.markdown(f"###### ⚖️ Yaşa Göre Kilo Persantil Eğrisi (0 - 120 Ay) · {sex_label}")
					_render_clinical_growth_curve("weight", "Kilo_kg", "Kilo", "kg", default_max_mo=120)
				with chart_tabs[2]:
					st.markdown(f"###### 🧬 Yaşa Göre VKİ (Beden Kitle İndeksi) Persantil Eğrisi · {sex_label}")
					_render_clinical_growth_curve("bmi", "BMI", "VKİ (BMI)", "kg/m²", default_max_mo=228)
			tab_chrono_main = chart_tabs[3]

			with tab_chrono_main:
				st.markdown("###### 📊 Kronolojik Persantil (%p) Takip Çizelgesi")
				st.caption("Çocuğun zaman içindeki boy, kilo, VKİ ve baş çevresi persantil dilimlerindeki değişimi ve kanal stabilitesini izleyin:")
				chart_records = enriched_records.assign(Tarih=pd.to_datetime(enriched_records["Tarih"]))
				avail_vars = [c for c in ["Boy_Persantil", "Kilo_Persantil", "BMI_Persantil", "Bas_Cevresi_Persantil"] if c in chart_records.columns and chart_records[c].notna().any()]
				if avail_vars:
					perc_long = chart_records.melt(
						id_vars=["Tarih"],
						value_vars=avail_vars,
						var_name="Gösterge",
						value_name="Persantil",
					).dropna(subset=["Persantil"])
					perc_long["Gösterge"] = perc_long["Gösterge"].map({
						"Boy_Persantil": "📏 Boy Persantil (%p)",
						"Kilo_Persantil": "⚖️ Kilo Persantil (%p)",
						"BMI_Persantil": "🧬 VKİ Persantil (%p)",
						"Bas_Cevresi_Persantil": "🧠 Baş Çevresi Persantil (%p)",
					})

					# Reference rule lines at 3, 15, 50, 85, 97 percentiles
					ref_rules = pd.DataFrame([
						{"Persantil": 3, "Band": "3p Kritik Alt Sınır"},
						{"Persantil": 15, "Band": "15p Alt Normal Sınır"},
						{"Persantil": 50, "Band": "50p İdeal Medyan"},
						{"Persantil": 85, "Band": "85p Üst Normal Sınır"},
						{"Persantil": 97, "Band": "97p Kritik Üst Sınır"},
					])
					rule_layer = alt.Chart(ref_rules).mark_rule(strokeDash=[3, 3], color="#cbd5e1").encode(
						y="Persantil:Q"
					)

					p_chart = alt.Chart(perc_long).mark_line(point=True, strokeWidth=2.5).encode(
						x=alt.X("Tarih:T", title="Ölçüm Tarihi"),
						y=alt.Y("Persantil:Q", title="Persantil (%p)", scale=alt.Scale(domain=[0, 100])),
						color=alt.Color("Gösterge:N", title="Gösterge"),
						tooltip=[alt.Tooltip("Tarih:T", title="Tarih"), alt.Tooltip("Gösterge:N"), alt.Tooltip("Persantil:Q", format=".1f")],
					)
					st.altair_chart(alt.layer(rule_layer, p_chart).properties(height=320), use_container_width=True)
				else:
					st.info("Kronolojik persantil grafiği için boy, kilo ve doğum tarihi içeren en az bir ölçüm kaydı bulunmalıdır.")


# ---------------------------------------------------------------------------
# OKUL SAĞLIK TARAMASI & MODAL PENCERE & SINIF RAPORLARI
# ---------------------------------------------------------------------------

@st.dialog("🩺 Öğrenci Sağlık Taraması", width="large")
def open_screening_dialog(student_id: int, class_name: str | None = None) -> None:
	student = get_person_by_id(student_id)
	if not student:
		st.error("Öğrenci bulunamadı.")
		return

	app_cfg = get_app_settings()
	visible_fields = get_screening_visible_fields()
	latest = get_student_latest_screening(student_id)
	meas_info = get_person_latest_measurements(student_id)

	# Fetch latest vitals
	with connect_db() as c_conn:
		v_row = c_conn.execute(
			"SELECT * FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC LIMIT 1",
			(student_id,),
		).fetchone()

	today = date.today()
	today_months = age_in_months(student["birth_date"], today) if student["birth_date"] else None
	age_str = f"{today_months // 12} yaş {today_months % 12} ay" if today_months is not None else "Doğum tarihi yok"
	sex_str = {"K": "Kız ♀", "E": "Erkek ♂"}.get(student["sex"], "Belirtilmedi")
	stu_class = student["class_name"] or class_name or "Belirtilmedi"
	sno_txt = f" · No: {student['student_no']}" if student["student_no"] else ""

	st.markdown(f"### 👤 {html.escape(student['full_name'])} <small style='color:#0f766e;'>({html.escape(stu_class)}{sno_txt})</small>", unsafe_allow_html=True)
	st.caption(f"🎂 {student['birth_date'] or '—'} ({age_str}) &nbsp;·&nbsp; {sex_str} &nbsp;·&nbsp; 🩸 Kan Grubu: `{student['blood_group'] if 'blood_group' in student.keys() and student['blood_group'] else 'Belirtilmedi'}`")

	if latest.get("screening_date"):
		st.info(f"🗓️ Son Tarama: **{latest.get('screening_date')}** (Yetkili: {latest.get('screened_by') or 'Revir Sorumlusu'})")
	else:
		st.warning("⏳ Bu öğrenciye bu eğitim-öğretim döneminde henüz tarama yapılmadı.")

	with st.form(f"form_screening_modal_{student_id}"):
		# 1. Date & Nurse
		c_top1, c_top2 = st.columns([1.2, 1.8])
		with c_top1:
			scr_date = st.date_input("🗓️ Tarama Tarihi", value=date.today())
		with c_top2:
			def_nurse = latest.get("screened_by") or app_cfg.get("nurse_name", "")
			scr_nurse = st.text_input("👩‍⚕️ Taramayı Yapan Hemşire / Yetkili", value=def_nurse)

		# 2. Antropometri
		has_anthropometry = any(f in visible_fields for f in ("height_cm", "weight_kg", "head_circ_cm"))
		new_h = None
		new_w = None
		new_hc = None
		if has_anthropometry:
			st.markdown("##### 📏 Fiziksel Ölçümler (Antropometri)")
			anthro_cols = []
			if "height_cm" in visible_fields: anthro_cols.append("h")
			if "weight_kg" in visible_fields: anthro_cols.append("w")
			if "head_circ_cm" in visible_fields: anthro_cols.append("hc")
			cols_obj = st.columns(len(anthro_cols))
			col_idx = 0
			if "height_cm" in visible_fields:
				with cols_obj[col_idx]:
					cur_h_val = float(latest.get("height_cm") or (meas_info["height_cm"] or 0.0))
					new_h = st.number_input("📏 Boy (cm)", min_value=0.0, max_value=250.0, value=cur_h_val, step=0.5, format="%.1f")
				col_idx += 1
			if "weight_kg" in visible_fields:
				with cols_obj[col_idx]:
					cur_w_val = float(latest.get("weight_kg") or (meas_info["weight_kg"] or 0.0))
					new_w = st.number_input("⚖️ Kilo (kg)", min_value=0.0, max_value=250.0, value=cur_w_val, step=0.1, format="%.1f")
				col_idx += 1
			if "head_circ_cm" in visible_fields:
				with cols_obj[col_idx]:
					cur_hc_val = float(latest.get("head_circ_cm") or (meas_info["head_circ_cm"] or 0.0))
					hc_help = "DSÖ standartlarına göre 0-5 yaş arası çocuklarda değerlendirilir." if (today_months is not None and today_months <= 60) else "Öğrenci 5 yaş üstü; rutin taramada opsiyoneldir."
					new_hc = st.number_input("👶 Baş Çevresi (cm)", min_value=0.0, max_value=80.0, value=cur_hc_val, step=0.5, format="%.1f", help=hc_help)
				col_idx += 1

		# 3. Vital Bulgular
		has_vitals = any(f in visible_fields for f in ("temperature_c", "pulse_bpm", "spo2_pct"))
		new_temp = None
		new_pulse = None
		new_spo2 = None
		if has_vitals:
			st.markdown("##### 💓 Vital Bulgular")
			vital_cols = []
			if "temperature_c" in visible_fields: vital_cols.append("temp")
			if "pulse_bpm" in visible_fields: vital_cols.append("pulse")
			if "spo2_pct" in visible_fields: vital_cols.append("spo2")
			vcols_obj = st.columns(len(vital_cols))
			vcol_idx = 0
			if "temperature_c" in visible_fields:
				with vcols_obj[vcol_idx]:
					cur_t_val = float(latest.get("temperature_c") or (v_row["temperature_c"] if v_row and v_row["temperature_c"] else 36.5))
					new_temp = st.number_input("🌡️ Ateş (°C)", min_value=0.0, max_value=45.0, value=cur_t_val, step=0.1, format="%.1f")
				vcol_idx += 1
			if "pulse_bpm" in visible_fields:
				with vcols_obj[vcol_idx]:
					cur_p_val = int(latest.get("pulse_bpm") or (v_row["pulse_bpm"] if v_row and v_row["pulse_bpm"] else 80))
					new_pulse = st.number_input("💓 Nabız (/dk)", min_value=0, max_value=220, value=cur_p_val, step=1)
				vcol_idx += 1
			if "spo2_pct" in visible_fields:
				with vcols_obj[vcol_idx]:
					cur_s_val = int(latest.get("spo2_pct") or (v_row["spo2_pct"] if v_row and v_row["spo2_pct"] else 98))
					new_spo2 = st.number_input("🫁 Oksijen Saturasyonu (%)", min_value=0, max_value=100, value=cur_s_val, step=1)
				vcol_idx += 1

		# 4. Göz & Görme
		has_vision = any(f in visible_fields for f in ("wears_glasses", "vision_right", "vision_left"))
		new_wears = None
		new_vr = None
		new_vl = None
		if has_vision:
			st.markdown("##### 👓 Göz ve Görme Durumu")
			vis_cols = []
			if "wears_glasses" in visible_fields: vis_cols.append("wears")
			if "vision_right" in visible_fields: vis_cols.append("vr")
			if "vision_left" in visible_fields: vis_cols.append("vl")
			vis_obj = st.columns(len(vis_cols))
			vis_idx = 0
			if "wears_glasses" in visible_fields:
				with vis_obj[vis_idx]:
					gl_opts = ["Hayır (Kullanmıyor)", "Evet (Sürekli)", "Bazen / Sadece Okurken", "Gözlük Önerildi / Muayene Bekleniyor", "Göz Tembelliği Kapaması Yapılıyor"]
					cur_gl = latest.get("wears_glasses") or "Hayır (Kullanmıyor)"
					i_gl = gl_opts.index(cur_gl) if cur_gl in gl_opts else 0
					new_wears = st.selectbox("👓 Gözlük Takıyor mu?", gl_opts, index=i_gl)
				vis_idx += 1
			if "vision_right" in visible_fields:
				with vis_obj[vis_idx]:
					cur_vr = latest.get("vision_right") or "Tam / Normal (10/10)"
					new_vr = st.text_input("👁️ Sağ Göz (Numara / Durum)", value=cur_vr)
				vis_idx += 1
			if "vision_left" in visible_fields:
				with vis_obj[vis_idx]:
					cur_vl = latest.get("vision_left") or "Tam / Normal (10/10)"
					new_vl = st.text_input("👁️ Sol Göz (Numara / Durum)", value=cur_vl)
				vis_idx += 1

		# 5. Saç & Kulak (Çoklu Seçim)
		has_hair_or_hearing = any(f in visible_fields for f in ("hair_screening", "hearing_right", "hearing_left"))
		new_hair = None
		new_hr = None
		new_hl = None
		if has_hair_or_hearing:
			st.markdown("##### 💇 Saç Taraması & 👂 İşitme Değerlendirmesi")
			hh_cols = []
			if "hair_screening" in visible_fields: hh_cols.append("hair")
			if "hearing_right" in visible_fields: hh_cols.append("hr")
			if "hearing_left" in visible_fields: hh_cols.append("hl")
			hh_obj = st.columns(len(hh_cols))
			hh_idx = 0
			if "hair_screening" in visible_fields:
				with hh_obj[hh_idx]:
					cur_hair_list = parse_multiselect_values(latest.get("hair_screening"))
					opts_hair = ensure_options_contain(HAIR_SCREENING_OPTIONS, cur_hair_list)
					sel_hair = st.multiselect("💇 Saç Taraması (Pediküloz - Çoklu)", opts_hair, default=cur_hair_list, placeholder="Bulguları seçin...")
					cust_hair = st.text_input("Saç Özel Notu (Opsiyonel)", value="", key=f"scr_cust_hair_{student_id}")
					new_hair = format_multiselect_save_str(sel_hair, cust_hair)
				hh_idx += 1
			if "hearing_right" in visible_fields:
				with hh_obj[hh_idx]:
					cur_hr_list = parse_multiselect_values(latest.get("hearing_right"))
					opts_hr = ensure_options_contain(HEARING_SCREENING_OPTIONS, cur_hr_list)
					sel_hr = st.multiselect("👂 Sağ Kulak İşitme (Çoklu)", opts_hr, default=cur_hr_list, placeholder="Sağ kulak bulguları...")
					new_hr = format_multiselect_save_str(sel_hr)
				hh_idx += 1
			if "hearing_left" in visible_fields:
				with hh_obj[hh_idx]:
					cur_hl_list = parse_multiselect_values(latest.get("hearing_left"))
					opts_hl = ensure_options_contain(HEARING_SCREENING_OPTIONS, cur_hl_list)
					sel_hl = st.multiselect("👂 Sol Kulak İşitme (Çoklu)", opts_hl, default=cur_hl_list, placeholder="Sol kulak bulguları...")
					new_hl = format_multiselect_save_str(sel_hl)
				hh_idx += 1

		# 6. Sağlık Davranışı, Zararlı Alışkanlıklar & Psikososyal
		has_beh_or_harm_or_psy = any(f in visible_fields for f in ("health_behavior", "harmful_habits", "psychosocial_status"))
		new_beh = None
		new_harm = None
		new_psy = None
		if has_beh_or_harm_or_psy:
			st.markdown("##### 🍎 Sağlık Davranışı, ⚠️ Zararlı Alışkanlıklar & 🧠 Psikososyal Durum")
			if "health_behavior" in visible_fields or "harmful_habits" in visible_fields:
				bh_cols = []
				if "health_behavior" in visible_fields: bh_cols.append("beh")
				if "harmful_habits" in visible_fields: bh_cols.append("harm")
				bh_obj = st.columns(len(bh_cols))
				bh_idx = 0
				if "health_behavior" in visible_fields:
					with bh_obj[bh_idx]:
						cur_beh_list = parse_multiselect_values(latest.get("health_behavior"))
						opts_beh = ensure_options_contain(HEALTH_BEHAVIOR_OPTIONS, cur_beh_list)
						sel_beh = st.multiselect("🍎 Sağlık Davranışı & Alışkanlıklar (Çoklu)", opts_beh, default=cur_beh_list, placeholder="Beslenme, diş, hijyen bulguları...")
						cust_beh = st.text_input("Sağlık Davranışı Ek Not (Opsiyonel)", value="", key=f"scr_cust_beh_{student_id}")
						new_beh = format_multiselect_save_str(sel_beh, cust_beh)
					bh_idx += 1
				if "harmful_habits" in visible_fields:
					with bh_obj[bh_idx]:
						cur_harm_list = parse_multiselect_values(latest.get("harmful_habits"))
						opts_harm = ensure_options_contain(HARMFUL_HABITS_OPTIONS, cur_harm_list)
						sel_harm = st.multiselect("⚠️ Zararlı Alışkanlıklar & Risk Faktörleri (Çoklu)", opts_harm, default=cur_harm_list, placeholder="Risk faktörlerini seçin...")
						cust_harm = st.text_input("Zararlı Alışkanlık Ek Not (Opsiyonel)", value="", key=f"scr_cust_harm_{student_id}")
						new_harm = format_multiselect_save_str(sel_harm, cust_harm)
					bh_idx += 1

			if "psychosocial_status" in visible_fields:
				psy_opts = [
					"Sınıf İçi Uyum ve İletişimi Çok İyi", "Çekingen / İçe Kapanık",
					"Aşırı Hareketli / Dikkat Dağınıklığı", "Okul Kaygısı / Anneye Bağlılık",
					"Öfke Kontrolü / Akran Çatışması", "Göz Teması ve İletişim Desteklenmeli", "Diğer / Serbest Not"
				]
				cur_psy_opt = latest.get("psychosocial_status") or "Sınıf İçi Uyum ve İletişimi Çok İyi"
				i_psy = psy_opts.index(cur_psy_opt) if cur_psy_opt in psy_opts else 0
				c_p1, c_p2 = st.columns([1.5, 1.5])
				with c_p1:
					sel_psy = st.selectbox("🧠 Genel Psikososyal Durumu", psy_opts, index=i_psy)
				with c_p2:
					cust_psy = st.text_input("Psikososyal Durum Özel Notu", value="" if sel_psy != "Diğer / Serbest Not" else cur_psy_opt)
				new_psy = cust_psy.strip() if sel_psy == "Diğer / Serbest Not" and cust_psy.strip() else sel_psy

		# 7. Notlar & Sevk
		st.markdown("##### 📝 Tarama Notu, Veli Bildirimi & Sevk Kararı")
		scr_notes = st.text_area("Özel Bulgular, Veli Bildirimi veya Uzman Hekim Sevk Notu", value=latest.get("notes") or "", height=70)

		# Submit button
		submitted = st.form_submit_button("💾 Sağlık Taramasını Kaydet ve Profile Eşitle", type="primary", use_container_width=True)

		if submitted:
			save_school_screening(
				student_id=student_id,
				screening_date=scr_date.isoformat(),
				height_cm=float(new_h) if (new_h is not None and new_h > 0) else None,
				weight_kg=float(new_w) if (new_w is not None and new_w > 0) else None,
				head_circ_cm=float(new_hc) if (new_hc is not None and new_hc > 0) else None,
				temperature_c=float(new_temp) if (new_temp is not None and new_temp > 0) else None,
				pulse_bpm=float(new_pulse) if (new_pulse is not None and new_pulse > 0) else None,
				spo2_pct=float(new_spo2) if (new_spo2 is not None and new_spo2 > 0) else None,
				wears_glasses=new_wears if "wears_glasses" in visible_fields else None,
				vision_right=new_vr.strip() if (new_vr and "vision_right" in visible_fields) else None,
				vision_left=new_vl.strip() if (new_vl and "vision_left" in visible_fields) else None,
				hair_screening=new_hair if "hair_screening" in visible_fields else None,
				hearing_right=new_hr if "hearing_right" in visible_fields else None,
				hearing_left=new_hl if "hearing_left" in visible_fields else None,
				health_behavior=new_beh if "health_behavior" in visible_fields else None,
				harmful_habits=new_harm if "harmful_habits" in visible_fields else None,
				psychosocial_status=new_psy if "psychosocial_status" in visible_fields else None,
				notes=scr_notes.strip() or None,
				screened_by=scr_nurse.strip() or None,
			)
			st.success(f"✅ {student['full_name']} öğrencisinin sağlık taraması kaydedildi ve tüm profille senkronize edildi.")
			st.rerun()


def build_class_screening_report_docx(class_name: str) -> bytes:
	app_cfg = get_app_settings()
	inst_top = app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI"
	inst_school = app_cfg.get("school_name") or ""
	inst_nurse = app_cfg.get("nurse_name") or "Okul Revir Sorumlusu"
	inst_princ = app_cfg.get("principal_name") or "Okul Müdürü / İdare"
	visible_fields = get_screening_visible_fields()

	with connect_db() as connection:
		students = connection.execute(
			"""SELECT s.id, s.student_no, s.full_name, s.sex, s.birth_date,
					  c.name as class_name, h.last_screened_at, h.wears_glasses,
					  h.vision_right, h.vision_left, h.hair_screening, h.hearing_right,
					  h.hearing_left, h.health_behavior, h.harmful_habits, h.psychosocial_status
			   FROM students s
			   JOIN classes c ON c.id = s.class_id
			   LEFT JOIN health_profiles h ON h.student_id = s.id
			   WHERE c.name = ?
			   ORDER BY s.student_no, s.full_name""",
			(class_name,),
		).fetchall()

	doc = docx.Document()
	section = doc.sections[0]
	section.orientation = WD_ORIENT.LANDSCAPE
	section.page_width = Cm(29.7)
	section.page_height = Cm(21.0)
	section.top_margin = Cm(1.2)
	section.bottom_margin = Cm(1.2)
	section.left_margin = Cm(1.2)
	section.right_margin = Cm(1.2)

	# Header Box
	p_top = doc.add_paragraph()
	p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_top.paragraph_format.space_before = Pt(0)
	p_top.paragraph_format.space_after = Pt(0)
	r_top = p_top.add_run(inst_top)
	r_top.font.name = "Calibri"
	r_top.font.size = Pt(10)
	r_top.font.color.rgb = RGBColor(100, 116, 139)

	if inst_school:
		p_sch = doc.add_paragraph()
		p_sch.alignment = WD_ALIGN_PARAGRAPH.CENTER
		p_sch.paragraph_format.space_before = Pt(0)
		p_sch.paragraph_format.space_after = Pt(2)
		r_sch = p_sch.add_run(inst_school)
		r_sch.font.name = "Calibri"
		r_sch.font.size = Pt(14)
		r_sch.bold = True
		r_sch.font.color.rgb = RGBColor(15, 118, 110)

	p_title = doc.add_paragraph()
	p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_title.paragraph_format.space_before = Pt(2)
	p_title.paragraph_format.space_after = Pt(6)
	r_title = p_title.add_run(f"🏫 {class_name} SAĞLIK TARAMASI VE GELİŞİM İZLEM FORMU")
	r_title.font.name = "Calibri"
	r_title.font.size = Pt(12)
	r_title.bold = True
	r_title.font.color.rgb = RGBColor(30, 41, 59)

	# Stats banner table
	stats = get_class_screening_stats(class_name)
	today_str = date.today().strftime("%d.%m.%Y")

	stats_table = doc.add_table(rows=1, cols=5)
	stats_table.alignment = WD_TABLE_ALIGNMENT.CENTER
	stats_items = [
		("Sınıf Mevcudu", f"{stats['total']} Öğrenci"),
		("Taranan", f"{stats['screened']} Öğrenci"),
		("Bekleyen", f"{stats['pending']} Öğrenci"),
		("Tarama Oranı", f"%{stats['percent']}"),
		("Rapor Tarihi", today_str),
	]
	for idx, (label, val) in enumerate(stats_items):
		cell = stats_table.cell(0, idx)
		set_cell_background(cell, "F0FDF4")
		set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
		cp = cell.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r1 = cp.add_run(f"{label}: ")
		r1.font.name = "Calibri"
		r1.font.size = Pt(9)
		r1.bold = True
		r2 = cp.add_run(val)
		r2.font.name = "Calibri"
		r2.font.size = Pt(9)

	set_table_borders(stats_table, color="BBF7D0", sz="6", val="single")

	doc.add_paragraph().paragraph_format.space_after = Pt(4)

	# Main screening table
	th_cols = ["Sıra", "No", "Öğrenci Adı Soyadı", "Cinsiyet", "Doğum / Yaş", "Son Tarama"]
	for f in visible_fields:
		if f in SCREENING_FIELDS_INFO:
			th_cols.append(SCREENING_FIELDS_INFO[f]["name"])

	table = doc.add_table(rows=1 + len(students), cols=len(th_cols))
	table.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(table, color="CBD5E1", sz="4", val="single")

	# Header row
	hdr_row = table.rows[0]
	hdr_trPr = hdr_row._element.get_or_add_trPr()
	hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

	for idx, col_name in enumerate(th_cols):
		c = hdr_row.cells[idx]
		set_cell_background(c, "0F766E")
		set_cell_margins(c, top=80, bottom=80, left=80, right=80)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(col_name)
		r.font.name = "Calibri"
		r.font.size = Pt(8.5)
		r.bold = True
		r.font.color.rgb = RGBColor(255, 255, 255)

	# Data rows
	for s_idx, s in enumerate(students):
		row_idx = s_idx + 1
		row = table.rows[row_idx]
		bg_color = "F8FAFC" if (s_idx % 2 == 1) else "FFFFFF"

		m_info = get_person_latest_measurements(s["id"])
		latest_scr = get_student_latest_screening(s["id"])
		s_age_m = age_in_months(s["birth_date"], date.today()) if s["birth_date"] else None
		s_age_str = f"{s_age_m // 12}y {s_age_m % 12}a" if s_age_m is not None else "—"
		s_sex = "Kız" if s["sex"] == "K" else ("Erkek" if s["sex"] == "E" else "—")
		scr_date_txt = s["last_screened_at"] or latest_scr.get("screening_date") or "—"

		row_cells_val = [
			str(s_idx + 1),
			str(s["student_no"] or "—"),
			s["full_name"],
			s_sex,
			f"{s['birth_date'] or '—'} ({s_age_str})",
			scr_date_txt,
		]

		for f in visible_fields:
			val = "—"
			if f == "height_cm":
				h_v = latest_scr.get("height_cm") or m_info["height_cm"]
				val = f"{h_v:.1f} cm" if h_v else "—"
			elif f == "weight_kg":
				w_v = latest_scr.get("weight_kg") or m_info["weight_kg"]
				val = f"{w_v:.1f} kg" if w_v else "—"
			elif f == "head_circ_cm":
				hc_v = latest_scr.get("head_circ_cm") or m_info["head_circ_cm"]
				val = f"{hc_v:.1f} cm" if hc_v else "—"
			elif f == "temperature_c":
				t_v = latest_scr.get("temperature_c")
				val = f"{t_v:.1f}°C" if t_v else "—"
			elif f == "pulse_bpm":
				p_v = latest_scr.get("pulse_bpm")
				val = f"{p_v:.0f}/dk" if p_v else "—"
			elif f == "spo2_pct":
				s_v = latest_scr.get("spo2_pct")
				val = f"%{s_v:.0f}" if s_v else "—"
			elif f == "wears_glasses":
				val = latest_scr.get("wears_glasses") or s["wears_glasses"] or "—"
			elif f == "vision_right":
				val = latest_scr.get("vision_right") or s["vision_right"] or "—"
			elif f == "vision_left":
				val = latest_scr.get("vision_left") or s["vision_left"] or "—"
			elif f == "hair_screening":
				val = latest_scr.get("hair_screening") or s["hair_screening"] or "—"
			elif f == "hearing_right":
				val = latest_scr.get("hearing_right") or s["hearing_right"] or "—"
			elif f == "hearing_left":
				val = latest_scr.get("hearing_left") or s["hearing_left"] or "—"
			elif f == "health_behavior":
				val = latest_scr.get("health_behavior") or s["health_behavior"] or "—"
			elif f == "harmful_habits":
				val = latest_scr.get("harmful_habits") or s["harmful_habits"] or "—"
			elif f == "psychosocial_status":
				val = latest_scr.get("psychosocial_status") or s["psychosocial_status"] or "—"

			row_cells_val.append(str(val))

		for col_idx, cell_val in enumerate(row_cells_val):
			c = row.cells[col_idx]
			if bg_color != "FFFFFF":
				set_cell_background(c, bg_color)
			set_cell_margins(c, top=60, bottom=60, left=70, right=70)
			cp = c.paragraphs[0]
			if col_idx in (0, 1, 3, 5):
				cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
			else:
				cp.alignment = WD_ALIGN_PARAGRAPH.LEFT
			r = cp.add_run(cell_val)
			r.font.name = "Calibri"
			r.font.size = Pt(8)
			r.font.color.rgb = RGBColor(30, 41, 59)

	# Signature section table
	doc.add_paragraph().paragraph_format.space_after = Pt(12)
	sig_table = doc.add_table(rows=1, cols=3)
	sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
	sig_data = [
		(f"{inst_nurse}\nOkul Revir Sorumlusu\n\n(İmza / Kaşe)"),
		("Sınıf Rehber Öğretmeni\nİnceleme & Onay\n\n(İmza)"),
		(f"{inst_princ}\nOkul Müdürü\n\n(Mühür / İmza)"),
	]
	for idx, text in enumerate(sig_data):
		c = sig_table.cell(0, idx)
		set_cell_margins(c, top=100, bottom=100, left=150, right=150)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(text)
		r.font.name = "Calibri"
		r.font.size = Pt(9.5)

	# Footer disclaimer
	p_foot = doc.add_paragraph()
	p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_foot.paragraph_format.space_before = Pt(10)
	r_foot = p_foot.add_run("Bu belge okul sağlık taraması ve revir kayıtları doğrultusunda sistemden otomatik üretilmiştir. Resmi tanı ve kurul raporu yerine geçmez.")
	r_foot.font.name = "Calibri"
	r_foot.font.size = Pt(8)
	r_foot.font.color.rgb = RGBColor(148, 163, 184)

	buf = BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_class_screening_report_html(class_name: str) -> str:
	app_cfg = get_app_settings()
	inst_top = html.escape(app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	inst_school = html.escape(app_cfg.get("school_name") or "")
	inst_nurse = html.escape(app_cfg.get("nurse_name") or "Okul Revir Sorumlusu")
	inst_princ = html.escape(app_cfg.get("principal_name") or "Okul Müdürü / İdare")
	visible_fields = get_screening_visible_fields()

	with connect_db() as connection:
		students = connection.execute(
			"""SELECT s.id, s.student_no, s.full_name, s.sex, s.birth_date,
					  c.name as class_name, h.last_screened_at, h.wears_glasses,
					  h.vision_right, h.vision_left, h.hair_screening, h.hearing_right,
					  h.hearing_left, h.health_behavior, h.harmful_habits, h.psychosocial_status
			   FROM students s
			   JOIN classes c ON c.id = s.class_id
			   LEFT JOIN health_profiles h ON h.student_id = s.id
			   WHERE c.name = ?
			   ORDER BY s.student_no, s.full_name""",
			(class_name,),
		).fetchall()

	th_cols = ["Sıra", "No", "Öğrenci Adı Soyadı", "Cinsiyet", "Doğum Tarihi / Yaş", "Son Tarama"]
	for f in visible_fields:
		if f in SCREENING_FIELDS_INFO:
			th_cols.append(SCREENING_FIELDS_INFO[f]["name"])

	thead_html = "".join(f"<th>{html.escape(c)}</th>" for c in th_cols)

	tbody_rows = []
	for idx, s in enumerate(students, 1):
		m_info = get_person_latest_measurements(s["id"])
		latest_scr = get_student_latest_screening(s["id"])
		s_age_m = age_in_months(s["birth_date"], date.today()) if s["birth_date"] else None
		s_age_str = f"{s_age_m // 12}y {s_age_m % 12}a" if s_age_m is not None else "—"
		s_sex = "Kız" if s["sex"] == "K" else ("Erkek" if s["sex"] == "E" else "—")
		scr_date_txt = s["last_screened_at"] or latest_scr.get("screening_date") or "—"

		row_cells = [
			str(idx),
			str(s["student_no"] or "—"),
			html.escape(s["full_name"]),
			s_sex,
			f"{s['birth_date'] or '—'} ({s_age_str})",
			scr_date_txt,
		]

		for f in visible_fields:
			val = "—"
			if f == "height_cm":
				h_v = latest_scr.get("height_cm") or m_info["height_cm"]
				val = f"{h_v:.1f} cm" if h_v else "—"
			elif f == "weight_kg":
				w_v = latest_scr.get("weight_kg") or m_info["weight_kg"]
				val = f"{w_v:.1f} kg" if w_v else "—"
			elif f == "head_circ_cm":
				hc_v = latest_scr.get("head_circ_cm") or m_info["head_circ_cm"]
				val = f"{hc_v:.1f} cm" if hc_v else "—"
			elif f == "temperature_c":
				t_v = latest_scr.get("temperature_c")
				val = f"{t_v:.1f}°C" if t_v else "—"
			elif f == "pulse_bpm":
				p_v = latest_scr.get("pulse_bpm")
				val = f"{p_v:.0f}/dk" if p_v else "—"
			elif f == "spo2_pct":
				s_v = latest_scr.get("spo2_pct")
				val = f"%{s_v:.0f}" if s_v else "—"
			elif f == "wears_glasses":
				val = latest_scr.get("wears_glasses") or s["wears_glasses"] or "—"
			elif f == "vision_right":
				val = latest_scr.get("vision_right") or s["vision_right"] or "—"
			elif f == "vision_left":
				val = latest_scr.get("vision_left") or s["vision_left"] or "—"
			elif f == "hair_screening":
				val = latest_scr.get("hair_screening") or s["hair_screening"] or "—"
			elif f == "hearing_right":
				val = latest_scr.get("hearing_right") or s["hearing_right"] or "—"
			elif f == "hearing_left":
				val = latest_scr.get("hearing_left") or s["hearing_left"] or "—"
			elif f == "health_behavior":
				val = latest_scr.get("health_behavior") or s["health_behavior"] or "—"
			elif f == "harmful_habits":
				val = latest_scr.get("harmful_habits") or s["harmful_habits"] or "—"
			elif f == "psychosocial_status":
				val = latest_scr.get("psychosocial_status") or s["psychosocial_status"] or "—"

			row_cells.append(html.escape(str(val)))

		tr_cells_html = "".join(f"<td>{c}</td>" for c in row_cells)
		tbody_rows.append(f"<tr>{tr_cells_html}</tr>")

	tbody_html = "".join(tbody_rows) if tbody_rows else "<tr><td colspan='20' style='text-align:center;'>Bu sınıfta kayıtlı öğrenci bulunmuyor.</td></tr>"

	stats = get_class_screening_stats(class_name)
	today_str = date.today().strftime("%d.%m.%Y")

	html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>{html.escape(class_name)} - Sınıf Sağlık Taraması Raporu</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; padding: 25px; line-height: 1.4; font-size: 13px; }}
.header-box {{ border-bottom: 2px solid #0f766e; padding-bottom: 12px; margin-bottom: 16px; text-align: center; }}
.header-box h1 {{ color: #0f766e; margin: 0; font-size: 20px; text-transform: uppercase; }}
.header-box h3 {{ color: #475569; margin: 3px 0 0; font-size: 13px; font-weight: normal; }}
.summary-bar {{ display: flex; justify-content: space-between; background: #f0fdfa; border: 1px solid #ccfbf1; padding: 10px 14px; border-radius: 6px; margin-bottom: 16px; font-size: 13px; }}
table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 12px; }}
th, td {{ border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; }}
th {{ background-color: #f1f5f9; color: #0f766e; font-weight: bold; white-space: nowrap; }}
tr:nth-child(even) {{ background-color: #f8fafc; }}
.signature-section {{ display: flex; justify-content: space-between; margin-top: 35px; page-break-inside: avoid; }}
.sig-box {{ border-top: 1px dashed #64748b; width: 30%; text-align: center; padding-top: 8px; font-size: 12px; }}
@media print {{
  body {{ padding: 0; font-size: 11px; }}
  .header-box h1 {{ font-size: 16px; }}
  table {{ font-size: 10px; }}
  th, td {{ padding: 4px 6px; }}
}}
</style>
</head>
<body>
<div class='header-box'>
	<h3>{inst_top}</h3>
	<h1>{inst_school}</h1>
	<h3>🏫 {html.escape(class_name)} SAĞLIK TARAMASI VE GELİŞİM İZLEM FORMU</h3>
</div>

<div class='summary-bar'>
	<div><b>Sınıf Mevcudu:</b> {stats['total']} Öğrenci</div>
	<div><b>Taranan:</b> {stats['screened']} Öğrenci</div>
	<div><b>Bekleyen:</b> {stats['pending']} Öğrenci</div>
	<div><b>Tarama Oranı:</b> %{stats['percent']}</div>
	<div><b>Rapor Tarihi:</b> {today_str}</div>
</div>

<table>
	<thead>
		<tr>{thead_html}</tr>
	</thead>
	<tbody>
		{tbody_html}
	</tbody>
</table>

<div class='signature-section'>
	<div class='sig-box'>
		<b>{inst_nurse}</b><br>
		Okul Revir Sorumlusu<br><br>
		(İmza / Kaşe)
	</div>
	<div class='sig-box'>
		<b>Sınıf Rehber Öğretmeni</b><br>
		İnceleme & Onay<br><br>
		(İmza)
	</div>
	<div class='sig-box'>
		<b>{inst_princ}</b><br>
		Okul Müdürü<br><br>
		(Mühür / İmza)
	</div>
</div>

<p style='margin-top: 25px; font-size: 10px; color: #94a3b8; text-align: center;'>
	Bu belge okul sağlık taraması ve revir kayıtları doğrultusunda sistemden otomatik üretilmiştir. Resmi tanı ve kurul raporu yerine geçmez.
</p>
</body>
</html>"""
	return html_content


def render_school_screening() -> None:
	visible_fields = get_screening_visible_fields()
	all_classes = [c for c in get_classes() if c != "Sivil / Diğer"]

	selected_class = st.session_state.get("selected_screening_class")

	# Overall statistics
	with connect_db() as connection:
		all_students_cnt = connection.execute("SELECT COUNT(*) FROM students WHERE class_id IN (SELECT id FROM classes WHERE name != 'Sivil / Diğer')").fetchone()[0]
		all_screened_cnt = connection.execute(
			"""SELECT COUNT(DISTINCT s.id) FROM students s
			   JOIN classes c ON c.id = s.class_id
			   LEFT JOIN health_profiles h ON h.student_id = s.id
			   WHERE c.name != 'Sivil / Diğer' AND (h.last_screened_at IS NOT NULL OR EXISTS (SELECT 1 FROM school_screenings sc WHERE sc.student_id = s.id))"""
		).fetchone()[0]

	pct_overall = round((all_screened_cnt / all_students_cnt * 100), 1) if all_students_cnt > 0 else 0

	# Top action & summary banner
	ban_c1, ban_c2, ban_c3, ban_c4 = st.columns([2.5, 1, 1, 1])
	with ban_c1:
		st.markdown("### 🏫 Okul Sağlık Taraması Paneli")
		st.caption("Sınıf kutucuklarına tıklayarak sınıf içi taramaları yürütün; boy, kilo, vital, görme, işitme ve hijyen bulgularını tek ekranda güncelleyin.")
	with ban_c2:
		st.metric("Toplam Sınıf", len(all_classes))
	with ban_c3:
		st.metric("Taranan / Toplam", f"{all_screened_cnt} / {all_students_cnt}")
	with ban_c4:
		st.metric("Tarama Oranı", f"%{pct_overall}")

	st.markdown("---")

	# VIEW 1: Sınıf Kutucukları Listesi (Grid of Class Boxes)
	if not selected_class or selected_class not in all_classes:
		st.markdown("#### 🗂️ Taranacak Sınıfı Seçin")
		st.caption("Aşağıdaki sınıf kutucuklarından birine tıklayarak o sınıfın öğrenci listesini ve tarama panelini açabilirsiniz.")

		vis_names = [SCREENING_FIELDS_INFO[f]["name"] for f in visible_fields if f in SCREENING_FIELDS_INFO]
		st.info(f"⚙️ **Aktif Görünür Tarama Alanları ({len(vis_names)}):** {', '.join(vis_names)} · *(Ayarlar > Okul Taraması Alanları sekmesinden değiştirebilirsiniz)*")

		# Render class cards in 3 columns
		grid_cols = st.columns(3)
		for idx, c_name in enumerate(all_classes):
			stats = get_class_screening_stats(c_name)
			col = grid_cols[idx % 3]
			with col:
				with st.container(border=True):
					st.markdown(f"### 🏫 {c_name}")
					c_stat1, c_stat2 = st.columns(2)
					c_stat1.markdown(f"👥 **Mevcut:** {stats['total']}")
					c_stat2.markdown(f"🩺 **Taranan:** {stats['screened']}")

					pct = stats["percent"]
					st.progress(pct / 100.0)

					if stats["is_complete"]:
						st.markdown("<span style='background:#dcfce7;color:#15803d;padding:2px 8px;border-radius:6px;font-size:12px;font-weight:600;'>🟢 %100 Tamamlandı</span>", unsafe_allow_html=True)
					elif stats["screened"] > 0:
						st.markdown(f"<span style='background:#fef3c7;color:#b45309;padding:2px 8px;border-radius:6px;font-size:12px;font-weight:600;'>🟡 %{pct} Devam Ediyor ({stats['pending']} bekliyor)</span>", unsafe_allow_html=True)
					else:
						st.markdown(f"<span style='background:#f1f5f9;color:#64748b;padding:2px 8px;border-radius:6px;font-size:12px;font-weight:600;'>⚪ Henüz Başlanmadı ({stats['total']} bekliyor)</span>", unsafe_allow_html=True)

					st.write("")
					if st.button(f"📋 Sınıfı Aç & Tara ➔", key=f"btn_sel_cls_{idx}_{c_name}", type="primary", use_container_width=True):
						st.session_state["selected_screening_class"] = c_name
						st.rerun()

	# VIEW 2: Seçili Sınıfın Tarama Paneli ve Öğrenci Listesi
	else:
		hdr_c1, hdr_c2, hdr_c3 = st.columns([1.5, 3.5, 2])
		with hdr_c1:
			if st.button("⬅️ Tüm Sınıflara Dön", key="btn_back_to_all_classes", use_container_width=True):
				st.session_state["selected_screening_class"] = None
				st.rerun()
		with hdr_c2:
			c_stats = get_class_screening_stats(selected_class)
			st.markdown(f"### 🏫 **{selected_class}** Sağlık Taraması Paneli")
			st.caption(f"Sınıf Mevcudu: **{c_stats['total']}** · Taranan: **{c_stats['screened']}** · Bekleyen: **{c_stats['pending']}** · Başarı: **%{c_stats['percent']}**")
		with hdr_c3:
			rep_docx = build_class_screening_report_docx(selected_class)
			st.download_button(
				"📝 Sınıf Tarama Raporunu İndir (.docx)",
				data=rep_docx,
				file_name=f"sinif_saglik_taramasi_{selected_class}_{date.today().isoformat()}.docx",
				mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				type="primary",
				key=f"dl_class_scr_rep_{selected_class}",
				use_container_width=True,
			)

		st.progress(c_stats["percent"] / 100.0)

		# Fetch students for this class
		with connect_db() as connection:
			stu_rows = connection.execute(
				"""SELECT s.id, s.student_no, s.full_name, s.sex, s.birth_date,
						  h.last_screened_at, h.wears_glasses, h.vision_right, h.vision_left,
						  h.hair_screening, h.hearing_right, h.hearing_left, h.health_behavior,
						  h.harmful_habits, h.psychosocial_status
				   FROM students s
				   JOIN classes c ON c.id = s.class_id
				   LEFT JOIN health_profiles h ON h.student_id = s.id
				   WHERE c.name = ?
				   ORDER BY s.student_no, s.full_name""",
				(selected_class,),
			).fetchall()

		if not stu_rows:
			st.warning(f"{selected_class} sınıfında henüz kayıtlı öğrenci bulunmuyor.")
			return

		# Filters toolbar
		filt_c1, filt_c2 = st.columns([2.5, 1.5])
		with filt_c1:
			search_kw = st.text_input("🔍 Öğrenci Ara (İsim veya Okul No)", placeholder="Aramak için yazın...", key=f"scr_search_{selected_class}")
		with filt_c2:
			status_filt = st.selectbox("Durum Filtresi", ["Tümü", "⏳ Taranmamış / Bekleyenler", "🟢 Taranmış Olanlar"], key=f"scr_filt_{selected_class}")

		# Filter student list
		filtered_students = []
		for s in stu_rows:
			is_screened = bool(s["last_screened_at"])
			if status_filt == "⏳ Taranmamış / Bekleyenler" and is_screened:
				continue
			if status_filt == "🟢 Taranmış Olanlar" and not is_screened:
				continue
			if search_kw.strip():
				kw = search_kw.strip().lower()
				matches_name = kw in str(s["full_name"]).lower()
				matches_no = kw in str(s["student_no"] or "").lower()
				if not (matches_name or matches_no):
					continue
			filtered_students.append(s)

		st.caption(f"Listelenen: **{len(filtered_students)}** / {len(stu_rows)} öğrenci · Her öğrencinin sağındaki **🩺 Taramayı Gir / Düzenle ✏️** butonuna basarak açılır pencerede taramasını yapabilirsiniz.")

		# Render each student card
		for s in filtered_students:
			s_id = int(s["id"])
			latest_scr = get_student_latest_screening(s_id)
			m_info = get_person_latest_measurements(s_id)
			s_age_m = age_in_months(s["birth_date"], date.today()) if s["birth_date"] else None
			s_age_str = f"{s_age_m // 12}y {s_age_m % 12}a" if s_age_m is not None else "Doğum tarihi yok"
			s_sex_ico = "♀ Kız" if s["sex"] == "K" else ("♂ Erkek" if s["sex"] == "E" else "—")
			sno_txt = f"No: {s['student_no']} · " if s["student_no"] else ""
			is_done = bool(s["last_screened_at"] or latest_scr.get("screening_date"))

			with st.container(border=True):
				row_l, row_r = st.columns([3.2, 1.8])
				with row_l:
					st_badge = "<span style='background:#dcfce7;color:#15803d;padding:2px 7px;border-radius:5px;font-size:12px;font-weight:600;'>🟢 Taranmış</span>" if is_done else "<span style='background:#fee2e2;color:#b91c1c;padding:2px 7px;border-radius:5px;font-size:12px;font-weight:600;'>⏳ Bekliyor</span>"
					st.markdown(f"**👤 {html.escape(s['full_name'])}** &nbsp; {st_badge}", unsafe_allow_html=True)
					st.caption(f"{sno_txt}{s_sex_ico} · Doğum: {s['birth_date'] or '—'} ({s_age_str}) &nbsp;·&nbsp; Son Tarama: **{latest_scr.get('screening_date') or s['last_screened_at'] or 'Yapılmadı'}**")

					preview_chips = []
					if "height_cm" in visible_fields or "weight_kg" in visible_fields:
						h_v = latest_scr.get("height_cm") or m_info["height_cm"]
						w_v = latest_scr.get("weight_kg") or m_info["weight_kg"]
						bmi_v = m_info["bmi"]
						hw_parts = []
						if "height_cm" in visible_fields and h_v: hw_parts.append(f"{h_v:.1f} cm")
						if "weight_kg" in visible_fields and w_v: hw_parts.append(f"{w_v:.1f} kg")
						if bmi_v: hw_parts.append(f"VKİ: {bmi_v:.1f}")
						if hw_parts:
							preview_chips.append(f"📏 {' / '.join(hw_parts)}")

					if "head_circ_cm" in visible_fields:
						hc_v = latest_scr.get("head_circ_cm") or m_info["head_circ_cm"]
						if hc_v: preview_chips.append(f"👶 Baş Çevresi: {hc_v:.1f} cm")

					if any(f in visible_fields for f in ("temperature_c", "pulse_bpm", "spo2_pct")):
						vit_parts = []
						if "temperature_c" in visible_fields and latest_scr.get("temperature_c"):
							vit_parts.append(f"🌡️ {latest_scr['temperature_c']:.1f}°C")
						if "pulse_bpm" in visible_fields and latest_scr.get("pulse_bpm"):
							vit_parts.append(f"💓 {latest_scr['pulse_bpm']:.0f}/dk")
						if "spo2_pct" in visible_fields and latest_scr.get("spo2_pct"):
							vit_parts.append(f"🫁 %{latest_scr['spo2_pct']:.0f}")
						if vit_parts:
							preview_chips.append(" · ".join(vit_parts))

					if any(f in visible_fields for f in ("wears_glasses", "vision_right", "vision_left")):
						gl_txt = latest_scr.get("wears_glasses") or s["wears_glasses"]
						vr_txt = latest_scr.get("vision_right") or s["vision_right"]
						vl_txt = latest_scr.get("vision_left") or s["vision_left"]
						vis_parts = []
						if "wears_glasses" in visible_fields and gl_txt: vis_parts.append(f"Gözlük: {gl_txt}")
						if "vision_right" in visible_fields and vr_txt: vis_parts.append(f"Sağ: {vr_txt}")
						if "vision_left" in visible_fields and vl_txt: vis_parts.append(f"Sol: {vl_txt}")
						if vis_parts:
							preview_chips.append(f"👓 {' · '.join(vis_parts)}")

					if "hair_screening" in visible_fields:
						hr_txt = latest_scr.get("hair_screening") or s["hair_screening"]
						if hr_txt: preview_chips.append(f"💇 Saç: {hr_txt}")

					if any(f in visible_fields for f in ("hearing_right", "hearing_left")):
						hear_r = latest_scr.get("hearing_right") or s["hearing_right"]
						hear_l = latest_scr.get("hearing_left") or s["hearing_left"]
						if hear_r or hear_l:
							preview_chips.append(f"👂 Kulak: Sağ: {hear_r or '—'} · Sol: {hear_l or '—'}")

					if "health_behavior" in visible_fields:
						beh_txt = latest_scr.get("health_behavior") or s["health_behavior"]
						if beh_txt: preview_chips.append(f"🍎 {beh_txt}")

					if "harmful_habits" in visible_fields:
						harm_txt = latest_scr.get("harmful_habits") or s["harmful_habits"]
						if harm_txt: preview_chips.append(f"⚠️ {harm_txt}")

					if "psychosocial_status" in visible_fields:
						psy_txt = latest_scr.get("psychosocial_status") or s["psychosocial_status"]
						if psy_txt: preview_chips.append(f"🧠 {psy_txt}")

					if preview_chips:
						chip_html = "".join([f"<span style='background:#f1f5f9;color:#334155;border:1px solid #cbd5e1;padding:2px 7px;border-radius:4px;font-size:11px;margin-right:6px;display:inline-block;margin-top:4px;'>{html.escape(c)}</span>" for c in preview_chips])
						st.markdown(chip_html, unsafe_allow_html=True)
					else:
						st.caption("Tarama verisi henüz girilmedi.")

				with row_r:
					st.write("")
					if st.button("🩺 Taramayı Gir / Düzenle ✏️", key=f"btn_open_scr_dialog_{s_id}", type="primary", use_container_width=True):
						open_screening_dialog(s_id, selected_class)
					if st.button("📂 Kişi Dosyası", key=f"btn_open_dossier_from_scr_{s_id}", use_container_width=True):
						open_person_dossier_dialog(s_id)


def append_student_report_docx_section(doc: docx.Document, student_id: int, is_multi: bool = False) -> None:
	app_cfg = get_app_settings()
	inst_top = app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI"
	inst_school = app_cfg.get("school_name") or ""
	inst_nurse = app_cfg.get("nurse_name") or "Okul Revir Sorumlusu"
	inst_princ = app_cfg.get("principal_name") or "Okul Müdürü / İdare"

	student = get_person_by_id(student_id)
	if not student:
		doc.add_paragraph(f"Öğrenci (ID: {student_id}) bulunamadı.")
		return

	stu_dict = dict(student) if hasattr(student, "keys") else (student if isinstance(student, dict) else {})
	stu_name = stu_dict.get("full_name", "")
	stu_type = stu_dict.get("person_type") or "Öğrenci"
	stu_sex = stu_dict.get("sex") or ""
	stu_birth = stu_dict.get("birth_date") or ""
	stu_no = stu_dict.get("student_no") or ""
	stu_class = stu_dict.get("class_name") or ""

	s_age_m = age_in_months(stu_birth, date.today()) if stu_birth else None
	s_age_str = f"{s_age_m // 12} yaş {s_age_m % 12} ay" if s_age_m is not None else "—"
	s_sex_str = "Kız" if stu_sex == "K" else ("Erkek" if stu_sex == "E" else "Belirtilmedi")

	# Header Box
	p_top = doc.add_paragraph()
	p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_top.paragraph_format.space_before = Pt(4)
	p_top.paragraph_format.space_after = Pt(0)
	r_top = p_top.add_run(f"{inst_top}{f' · {inst_school}' if inst_school else ''}")
	r_top.font.name = "Calibri"
	r_top.font.size = Pt(10)
	r_top.font.color.rgb = RGBColor(100, 116, 139)

	p_title = doc.add_paragraph()
	p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_title.paragraph_format.space_before = Pt(2)
	p_title.paragraph_format.space_after = Pt(2)
	r_title = p_title.add_run("ÖĞRENCİ REVİR & SAĞLIK GELİŞİM RAPORU")
	r_title.font.name = "Calibri"
	r_title.font.size = Pt(13)
	r_title.bold = True
	r_title.font.color.rgb = RGBColor(15, 118, 110)

	p_dt = doc.add_paragraph()
	p_dt.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_dt.paragraph_format.space_before = Pt(0)
	p_dt.paragraph_format.space_after = Pt(6)
	r_dt = p_dt.add_run(f"Rapor Tarihi: {date.today().strftime('%d.%m.%Y')}")
	r_dt.font.name = "Calibri"
	r_dt.font.size = Pt(9)
	r_dt.font.color.rgb = RGBColor(100, 116, 139)

	# Student Info Box (table)
	info_tbl = doc.add_table(rows=2, cols=3)
	info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
	set_table_borders(info_tbl, color="CBD5E1", sz="4", val="single")
	info_data = [
		[("Adı Soyadı", stu_name), ("Sınıf / Şube", stu_class or "—"), ("Okul No", str(stu_no or "—"))],
		[("Kayıt Türü", stu_type), ("Cinsiyet", s_sex_str), ("Doğum Tarihi / Yaş", f"{stu_birth or '—'} ({s_age_str})")],
	]
	for r_i, r_data in enumerate(info_data):
		for c_i, (lbl, val) in enumerate(r_data):
			cell = info_tbl.cell(r_i, c_i)
			set_cell_background(cell, "F8FAFC")
			set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
			cp = cell.paragraphs[0]
			r1 = cp.add_run(f"{lbl}: ")
			r1.font.name = "Calibri"
			r1.font.size = Pt(9)
			r1.font.color.rgb = RGBColor(100, 116, 139)
			r2 = cp.add_run(str(val))
			r2.font.name = "Calibri"
			r2.font.size = Pt(9)
			r2.bold = True
			r2.font.color.rgb = RGBColor(30, 41, 59)

	# Fetch all data from DB
	with connect_db() as connection:
		visits = pd.read_sql_query(
			"SELECT visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru, symptoms AS Belirtiler, observation AS Gözlem, action AS İşlem, guardian_contacted AS Veli_Bilgilendirildi, referral AS Yönlendirme, follow_up AS İzlem FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC",
			connection, params=(student_id,),
		)
		measurements = pd.read_sql_query(
			"SELECT id, measured_on AS Tarih, height_cm AS Boy_cm, weight_kg AS Kilo_kg, head_circ_cm AS Bas_Cevresi_cm, age_months AS Yas_ay, bmi AS BMI FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC",
			connection, params=(student_id,),
		)
		vitals = pd.read_sql_query(
			"SELECT measured_at AS Tarih_Saat, systolic_mmhg AS Buyuk_Tansiyon, diastolic_mmhg AS Kucuk_Tansiyon, pulse_bpm AS Nabiz, temperature_c AS Ates_C, spo2_pct AS Saturasyon_SpO2, glucose_mg_dl AS Kan_Sekeri, glucose_context AS Olcum_Kosulu, notes AS Notlar FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC",
			connection, params=(student_id,),
		)
		lab_results = pd.read_sql_query(
			"SELECT r.report_date AS Tarih, r.file_name AS Dosya, x.test_name AS Test, x.value AS Sonuç, x.unit AS Birim, x.ref_low AS Referans_Alt, x.ref_high AS Referans_Üst, x.lab_flag AS Rapor_İşareti FROM lab_results x JOIN lab_reports r ON r.id=x.report_id WHERE r.student_id=? ORDER BY r.report_date DESC, x.id",
			connection, params=(student_id,),
		)
		health_profile = pd.read_sql_query(
			"""SELECT COALESCE(h.guardian_name, s.guardian_name) AS Veli,
			   COALESCE(h.guardian_phone, s.guardian_phone) AS Telefon, h.allergies AS Alerjiler,
			   h.chronic_conditions AS Kronik_Hastalıklar, h.medical_history AS Sağlık_Geçmişi,
			   h.medications AS İlaçlar, h.blood_group AS Kan_Grubu,
			   h.wears_glasses AS Gözlük, h.vision_right AS Sağ_Göz, h.vision_left AS Sol_Göz,
			   h.hair_screening AS Saç_Taraması, h.hearing_right AS Sağ_Kulak, h.hearing_left AS Sol_Kulak,
			   h.health_behavior AS Sağlık_Davranışı, h.harmful_habits AS Zararlı_Alışkanlıklar, h.psychosocial_status AS Psikososyal_Durum,
			   h.emergency_notes AS Acil_Notlar, h.updated_at AS Güncelleme, h.reviewed_by AS Doğrulayan
		   FROM students s LEFT JOIN health_profiles h ON h.student_id=s.id
		   WHERE s.id=? AND (h.student_id IS NOT NULL OR s.guardian_name IS NOT NULL OR s.guardian_phone IS NOT NULL)""",
			connection, params=(student_id,),
		)
		care_plans = pd.read_sql_query(
			"SELECT created_at AS Tarih, nursing_diagnosis AS Hemşirelik_Değerlendirmesi, observed_facts AS Olgular, goals AS Hedefler, interventions AS Girişimler, monitoring AS İzlem, escalation AS Yönlendirme, status AS Durum, reviewed_by AS Sorumlu FROM care_plans WHERE student_id=? ORDER BY created_at DESC",
			connection, params=(student_id,),
		)
		vac_rows = connection.execute(
			"""SELECT vaccine_name AS Aşı, dose_name AS Doz, period_label AS Takvim_Dönemi,
					  CASE WHEN is_done = 1 THEN 'Yapıldı ✔' ELSE 'Yapılmadı / Eksik ✖' END AS Durum,
					  COALESCE(administered_date, '—') AS Uygulama_Tarihi,
					  COALESCE(facility, '—') AS Kurum
			   FROM student_vaccinations WHERE student_id = ? ORDER BY period_months, id""",
			(student_id,)
		).fetchall()

	# Alert Box if allergies, chronic conditions, etc.
	if not health_profile.empty:
		p_row = health_profile.iloc[0]
		alerji = clean_text(p_row.get("Alerjiler"))
		kronik = clean_text(p_row.get("Kronik_Hastalıklar"))
		gecmis = clean_text(p_row.get("Sağlık_Geçmişi"))
		ilac = clean_text(p_row.get("İlaçlar"))
		kan = clean_text(p_row.get("Kan_Grubu"))
		acil = clean_text(p_row.get("Acil_Notlar"))

		if alerji or kronik or gecmis or ilac or kan or acil:
			alert_tbl = doc.add_table(rows=1, cols=1)
			alert_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
			set_table_borders(alert_tbl, color="EF4444", sz="8", val="single")
			c = alert_tbl.cell(0, 0)
			set_cell_background(c, "FEF2F2")
			set_cell_margins(c, top=80, bottom=80, left=120, right=120)
			cp = c.paragraphs[0]
			r_at = cp.add_run("⚠️ DİKKAT EDİLECEK SAĞLIK BİLGİLERİ\n")
			r_at.font.name = "Calibri"
			r_at.font.size = Pt(10)
			r_at.bold = True
			r_at.font.color.rgb = RGBColor(185, 28, 28)

			alert_lines = []
			if kan: alert_lines.append(f"🩸 Kan Grubu: {kan}")
			if alerji: alert_lines.append(f"🔴 Alerjiler: {alerji}")
			if kronik: alert_lines.append(f"🟠 Kronik Hastalıklar: {kronik}")
			if ilac: alert_lines.append(f"💊 Düzenli İlaç: {ilac}")
			if gecmis: alert_lines.append(f"📋 Sağlık Geçmişi: {gecmis}")
			if acil: alert_lines.append(f"🚨 Acil Notu: {acil}")

			r_al = cp.add_run("\n".join(alert_lines))
			r_al.font.name = "Calibri"
			r_al.font.size = Pt(8.5)
			r_al.font.color.rgb = RGBColor(51, 65, 85)

	# 1. Sağlık Profili Tablosu
	add_styled_df_table(doc, health_profile, title="📋 Sağlık Profili ve Acil Bilgiler", empty_msg="Sağlık profili kaydı bulunmuyor.")

	# 2. Büyüme ve Persantil Özeti
	person_meas = get_person_latest_measurements(student_id)
	report_growth_df = None
	can_eval_rep_age = False
	can_eval_rep_hc = False
	if person_meas["has_data"] and stu_sex:
		latest_m = None
		if person_meas["latest_measured_on"] and stu_birth:
			try:
				latest_m = age_in_months(stu_birth, date.fromisoformat(person_meas["latest_measured_on"]))
			except Exception:
				latest_m = None
		if latest_m is not None and 0 <= latest_m <= 228:
			can_eval_rep_age = True
			can_eval_rep_hc = (latest_m <= 60)
			report_growth_df = get_clinical_growth_summary_table(
				person_meas["height_cm"], person_meas["weight_kg"], person_meas["bmi"],
				stu_sex, latest_m,
				head_circ_cm=person_meas["head_circ_cm"] if can_eval_rep_hc else None,
				always_include_head_circ=can_eval_rep_hc
			)

	if report_growth_df is not None and not report_growth_df.empty:
		growth_disp = report_growth_df.copy()
		growth_disp["Z-Skoru (SDS)"] = growth_disp["Z-Skoru (SDS)"].map(lambda x: f"{x:+.2f} SDS" if pd.notna(x) else "—")
		if "Persentil (%)" in growth_disp.columns:
			growth_disp["Persentil (%)"] = growth_disp["Persentil (%)"].map(lambda x: f"%{x:.1f}p" if pd.notna(x) else "—")
		add_styled_df_table(doc, growth_disp, title="📏 Büyüme ve Persantil Taraması (DSÖ)")
	else:
		p_gr_e = doc.add_paragraph()
		p_gr_e.paragraph_format.space_before = Pt(8)
		p_gr_e.paragraph_format.space_after = Pt(2)
		r_t = p_gr_e.add_run("📏 Büyüme ve Persantil Taraması\n")
		r_t.font.name = "Calibri"
		r_t.font.size = Pt(10.5)
		r_t.bold = True
		r_t.font.color.rgb = RGBColor(15, 118, 110)
		r_m = p_gr_e.add_run("Ölçüm kaydı veya doğum tarihi bulunmadığından persantil tablosu oluşturulmadı.")
		r_m.font.size = Pt(8.5)
		r_m.font.italic = True
		r_m.font.color.rgb = RGBColor(148, 163, 184)

	# 3. Vital Bulgular
	add_styled_df_table(doc, vitals, title="💓 Vital Ölçüm Geçmişi", empty_msg="Vital bulgu kaydı bulunmuyor.")

	# 4. Laboratuvar Tahlil Sonuçları
	lab_results_assessed = assess_lab_history(lab_results)
	add_styled_df_table(doc, lab_results_assessed, title="🧪 Laboratuvar Tahlil Sonuçları", empty_msg="Tahlil kaydı bulunmuyor.")

	# 5. Boy, Kilo ve Baş Çevresi Ölçüm Geçmişi
	if not measurements.empty:
		rep_enriched = enrich_measurements_with_percentiles(measurements, stu_sex, stu_birth)
		rep_cols = [
			"Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Boy_Dilimi",
			"Kilo_kg", "Kilo_Z", "Kilo_Persantil", "Kilo_Dilimi",
			"Ideal_Kilo_Araligi", "BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"
		]
		has_hc_meas = (rep_enriched["Bas_Cevresi_cm"].notna() & (rep_enriched["Bas_Cevresi_cm"] > 0)).any()
		if can_eval_rep_hc or has_hc_meas:
			rep_cols.insert(10, "Bas_Cevresi_cm")
		rep_disp = rep_enriched[[c for c in rep_cols if c in rep_enriched.columns]].rename(columns={
			"Yas_ay": "Yaş (Ay)", "Boy_cm": "Boy (cm)", "Boy_Z": "Boy Z", "Boy_Persantil": "Boy %",
			"Boy_Dilimi": "Boy Dilimi", "Kilo_kg": "Kilo (kg)", "Kilo_Z": "Kilo Z", "Kilo_Persantil": "Kilo %",
			"Kilo_Dilimi": "Kilo Dilimi", "Ideal_Kilo_Araligi": "İdeal Kilo Aralığı",
			"Bas_Cevresi_cm": "Baş Çev (cm)", "BMI": "VKİ", "BMI_Z": "VKİ Z", "BMI_Persantil": "VKİ %",
			"Gelisim_Degerlendirmesi": "Gelişim Değerlendirmesi"
		})
		formatted_m = format_measurements_for_html(rep_disp)
		add_styled_df_table(doc, formatted_m, title="⚖️ Boy, Kilo ve Baş Çevresi Ölçüm Geçmişi")
	else:
		add_styled_df_table(doc, None, title="⚖️ Boy, Kilo ve Baş Çevresi Ölçüm Geçmişi", empty_msg="Ölçüm kaydı bulunmuyor.")

	# 6. Revir Başvuru ve Müdahale Kayıtları
	add_styled_df_table(doc, visits, title="🩺 Revir Başvuru ve Müdahale Kayıtları", empty_msg="Revir ziyareti kaydı bulunmuyor.")

	# 7. Bakım Planı Kayıtları
	add_styled_df_table(doc, care_plans, title="📝 Bakım Planı Kayıtları", empty_msg="Bakım planı kaydı bulunmuyor.")

	# 8. Aşı ve Bağışıklama Kayıtları
	if vac_rows:
		vac_df = pd.DataFrame([dict(r) for r in vac_rows])
		add_styled_df_table(doc, vac_df, title="💉 Aşı ve Bağışıklama Kayıtları (T.C. Sağlık Bakanlığı)")
	else:
		add_styled_df_table(doc, None, title="💉 Aşı ve Bağışıklama Kayıtları (T.C. Sağlık Bakanlığı)", empty_msg="Kayıtlı aşı bilgisi bulunmuyor.")

	# Signature section table
	doc.add_paragraph().paragraph_format.space_after = Pt(10)
	sig_table = doc.add_table(rows=1, cols=3)
	sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
	sig_data = [
		(f"{inst_nurse}\nOkul Revir Sorumlusu\n\n(İmza / Kaşe)"),
		("Veli Bilgilendirme Onayı\n\n\n(Tarih / İmza)"),
		(f"{inst_princ}\nOkul Müdürü\n\n(Mühür / İmza)"),
	]
	for idx, text in enumerate(sig_data):
		c = sig_table.cell(0, idx)
		set_cell_margins(c, top=80, bottom=80, left=120, right=120)
		cp = c.paragraphs[0]
		cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
		r = cp.add_run(text)
		r.font.name = "Calibri"
		r.font.size = Pt(9)

	# Footer disclaimer
	p_foot = doc.add_paragraph()
	p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
	p_foot.paragraph_format.space_before = Pt(8)
	p_foot.paragraph_format.space_after = Pt(4)
	r_foot = p_foot.add_run("Bu belge okul reviri dahili izlem ve sağlık bilgilendirme tutanağıdır; tek başına resmi sağlık kurulu raporu veya tıbbi tanı yerine geçmez.")
	r_foot.font.name = "Calibri"
	r_foot.font.size = Pt(7.5)
	r_foot.font.color.rgb = RGBColor(148, 163, 184)


def build_student_report_docx(student_or_id: int | sqlite3.Row | dict) -> bytes:
	stu_id = student_or_id["id"] if isinstance(student_or_id, (sqlite3.Row, dict)) else int(student_or_id)
	doc = docx.Document()
	section = doc.sections[0]
	section.orientation = WD_ORIENT.PORTRAIT
	section.page_width = Cm(21.0)
	section.page_height = Cm(29.7)
	section.top_margin = Cm(1.5)
	section.bottom_margin = Cm(1.5)
	section.left_margin = Cm(1.5)
	section.right_margin = Cm(1.5)

	append_student_report_docx_section(doc, stu_id, is_multi=False)

	buf = BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_class_batch_report_docx(class_name: str) -> bytes:
	class_students = student_options(class_name, person_type="Öğrenci")
	doc = docx.Document()
	section = doc.sections[0]
	section.orientation = WD_ORIENT.PORTRAIT
	section.page_width = Cm(21.0)
	section.page_height = Cm(29.7)
	section.top_margin = Cm(1.5)
	section.bottom_margin = Cm(1.5)
	section.left_margin = Cm(1.5)
	section.right_margin = Cm(1.5)

	for idx, st_row in enumerate(class_students):
		append_student_report_docx_section(doc, int(st_row["id"]), is_multi=True)
		if idx < len(class_students) - 1:
			doc.add_page_break()

	buf = BytesIO()
	doc.save(buf)
	return buf.getvalue()


def build_student_report_section_html(student: sqlite3.Row, inst_top: str, inst_school: str, inst_nurse: str, inst_princ: str, is_multi: bool = False) -> str:
	stu_dict = dict(student) if hasattr(student, "keys") else (student if isinstance(student, dict) else {})
	stu_id = stu_dict.get("id")
	stu_name = stu_dict.get("full_name", "")
	stu_type = stu_dict.get("person_type", "Öğrenci")
	stu_sex = stu_dict.get("sex", "")
	stu_birth = stu_dict.get("birth_date", "")
	stu_no = stu_dict.get("student_no", "")
	stu_class = stu_dict.get("class_name", "")
	if not stu_class and stu_dict.get("class_id"):
		try:
			with connect_db() as c_conn:
				c_row = c_conn.execute("SELECT name FROM classes WHERE id=?", (stu_dict["class_id"],)).fetchone()
				if c_row and c_row["name"]:
					stu_class = c_row["name"]
		except Exception:
			pass

	with connect_db() as connection:
		visits = pd.read_sql_query(
			"SELECT visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru, symptoms AS Belirtiler, observation AS Gözlem, action AS İşlem, guardian_contacted AS Veli_Bilgilendirildi, referral AS Yönlendirme, follow_up AS İzlem FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC",
			connection, params=(stu_id,),
		)
		measurements = pd.read_sql_query(
			"SELECT id, measured_on AS Tarih, height_cm AS Boy_cm, weight_kg AS Kilo_kg, head_circ_cm AS Bas_Cevresi_cm, age_months AS Yas_ay, bmi AS BMI FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC",
			connection, params=(stu_id,),
		)
		vitals = pd.read_sql_query(
			"SELECT measured_at AS Tarih_Saat, systolic_mmhg AS Buyuk_Tansiyon, diastolic_mmhg AS Kucuk_Tansiyon, pulse_bpm AS Nabiz, temperature_c AS Ates_C, spo2_pct AS Saturasyon_SpO2, glucose_mg_dl AS Kan_Sekeri, glucose_context AS Olcum_Kosulu, notes AS Notlar FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC",
			connection, params=(stu_id,),
		)
		lab_results = pd.read_sql_query(
			"SELECT r.report_date AS Tarih, r.file_name AS Dosya, x.test_name AS Test, x.value AS Sonuç, x.unit AS Birim, x.ref_low AS Referans_Alt, x.ref_high AS Referans_Üst, x.lab_flag AS Rapor_İşareti FROM lab_results x JOIN lab_reports r ON r.id=x.report_id WHERE r.student_id=? ORDER BY r.report_date DESC, x.id",
			connection, params=(stu_id,),
		)
		health_profile = pd.read_sql_query(
			"""SELECT COALESCE(h.guardian_name, s.guardian_name) AS Veli,
			   COALESCE(h.guardian_phone, s.guardian_phone) AS Telefon, h.allergies AS Alerjiler,
			   h.chronic_conditions AS Kronik_Hastalıklar, h.medical_history AS Sağlık_Geçmişi,
			   h.medications AS İlaçlar, h.blood_group AS Kan_Grubu,
			   h.wears_glasses AS Gözlük, h.vision_right AS Sağ_Göz, h.vision_left AS Sol_Göz,
			   h.hair_screening AS Saç_Taraması, h.hearing_right AS Sağ_Kulak, h.hearing_left AS Sol_Kulak,
			   h.health_behavior AS Sağlık_Davranışı, h.harmful_habits AS Zararlı_Alışkanlıklar, h.psychosocial_status AS Psikososyal_Durum,
			   h.emergency_notes AS Acil_Notlar, h.updated_at AS Güncelleme, h.reviewed_by AS Doğrulayan
		   FROM students s LEFT JOIN health_profiles h ON h.student_id=s.id
		   WHERE s.id=? AND (h.student_id IS NOT NULL OR s.guardian_name IS NOT NULL OR s.guardian_phone IS NOT NULL)""",
			connection, params=(stu_id,),
		)
		care_plans = pd.read_sql_query(
			"SELECT created_at AS Tarih, nursing_diagnosis AS Hemşirelik_Değerlendirmesi, observed_facts AS Olgular, goals AS Hedefler, interventions AS Girişimler, monitoring AS İzlem, escalation AS Yönlendirme, status AS Durum, reviewed_by AS Sorumlu FROM care_plans WHERE student_id=? ORDER BY created_at DESC",
			connection, params=(stu_id,),
		)
		vac_rows = connection.execute(
			"""SELECT vaccine_name AS Aşı, dose_name AS Doz, period_label AS Takvim_Dönemi,
					  CASE WHEN is_done = 1 THEN 'Yapıldı ✔' ELSE 'Yapılmadı / Eksik ✖' END AS Durum,
					  COALESCE(administered_date, '—') AS Uygulama_Tarihi,
					  COALESCE(facility, '—') AS Kurum
			   FROM student_vaccinations WHERE student_id = ? ORDER BY period_months, id""",
			(stu_id,)
		).fetchall()

	lab_results = assess_lab_history(lab_results)
	report_growth_df = None
	person_meas = get_person_latest_measurements(stu_id)
	can_eval_rep_age = False
	can_eval_rep_hc = False
	if person_meas["has_data"] and stu_sex:
		latest_m = None
		if person_meas["latest_measured_on"] and stu_birth:
			try:
				latest_m = age_in_months(stu_birth, date.fromisoformat(person_meas["latest_measured_on"]))
			except Exception:
				latest_m = None
		if latest_m is not None and 0 <= latest_m <= 228:
			can_eval_rep_age = True
			can_eval_rep_hc = (latest_m <= 60)
			report_growth_df = get_clinical_growth_summary_table(
				person_meas["height_cm"], person_meas["weight_kg"], person_meas["bmi"],
				stu_sex, latest_m,
				head_circ_cm=person_meas["head_circ_cm"] if can_eval_rep_hc else None,
				always_include_head_circ=can_eval_rep_hc
			)

	safe_name = html.escape(str(stu_name or ""))
	safe_class = html.escape(str(stu_class or "—"))
	visits_html = visits.to_html(index=False, border=0, na_rep="—") if not visits.empty else "<p>Revir kaydı yok.</p>"
	vitals_html = vitals.to_html(index=False, border=0, na_rep="—") if not vitals.empty else "<p>Vital kaydı yok.</p>"
	labs_html = lab_results.to_html(index=False, border=0, na_rep="—") if not lab_results.empty else "<p>Tahlil kaydı yok.</p>"
	profile_html = health_profile.to_html(index=False, border=0, na_rep="—") if not health_profile.empty else "<p>Sağlık profili yok.</p>"
	care_plans_html = care_plans.to_html(index=False, border=0, na_rep="—") if not care_plans.empty else "<p>Bakım planı kaydı yok.</p>"
	if vac_rows:
		vac_df = pd.DataFrame([dict(r) for r in vac_rows])
		vaccinations_html = vac_df.to_html(index=False, border=0, na_rep="—")
	else:
		vaccinations_html = "<p style='color:#64748b;font-style:italic;'>Kayıtlı aşı bilgisi bulunmuyor.</p>"

	if not measurements.empty:
		rep_enriched = enrich_measurements_with_percentiles(measurements, stu_sex, stu_birth)
		rep_cols = [
			"Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Boy_Dilimi", "Boy_Farki",
			"Kilo_kg", "Kilo_Z", "Kilo_Persantil", "Kilo_Dilimi", "Kilo_Farki",
			"Ideal_Kilo_Araligi",
			"BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"
		]
		has_hc_meas = (rep_enriched["Bas_Cevresi_cm"].notna() & (rep_enriched["Bas_Cevresi_cm"] > 0)).any()
		if can_eval_rep_hc or has_hc_meas:
			rep_cols[12:12] = ["Bas_Cevresi_cm", "Bas_Cevresi_Z", "Bas_Cevresi_Persantil"]
		if not can_eval_rep_age:
			rep_cols = [c for c in rep_cols if "_Z" not in c and "Persantil" not in c and "Dilimi" not in c and c != "Gelisim_Degerlendirmesi"]

		rep_disp = rep_enriched[[c for c in rep_cols if c in rep_enriched.columns]].rename(columns={
			"Yas_ay": "Yaş (Ay)", "Boy_cm": "Boy (cm)", "Boy_Z": "Boy Z-Skor", "Boy_Persantil": "Boy %",
			"Boy_Dilimi": "Boy Dilimi", "Boy_Farki": "Boy Değişimi", "Kilo_kg": "Kilo (kg)",
			"Kilo_Z": "Kilo Z-Skor", "Kilo_Persantil": "Kilo %", "Kilo_Dilimi": "Kilo Dilimi",
			"Kilo_Farki": "Kilo Değişimi", "Ideal_Kilo_kg": "İdeal Kilo (kg)", "Ideal_Kilo_Araligi": "İdeal Kilo Aralığı", "Bas_Cevresi_cm": "Baş Çevresi (cm)",
			"Bas_Cevresi_Z": "Baş Çevresi Z-Skor", "Bas_Cevresi_Persantil": "Baş Çevresi %",
			"BMI": "VKİ (BMI)", "BMI_Z": "VKİ Z-Skor", "BMI_Persantil": "VKİ %",
			"Gelisim_Degerlendirmesi": "Gelişim Değerlendirmesi"
		})
		measurements_html = format_measurements_for_html(rep_disp).to_html(index=False, border=0, na_rep="—")
	else:
		measurements_html = "<p>Ölçüm kaydı yok.</p>"

	if report_growth_df is not None and not report_growth_df.empty and can_eval_rep_age:
		growth_html_df = report_growth_df.copy()
		growth_html_df["Z-Skoru (SDS)"] = growth_html_df["Z-Skoru (SDS)"].map(lambda x: f"{x:+.2f} SDS" if pd.notna(x) else "—")
		if "Persentil (%)" in growth_html_df.columns:
			growth_html_df["Persentil (%)"] = growth_html_df["Persentil (%)"].map(lambda x: f"%{x:.1f}p" if pd.notna(x) else "—")
		growth_html = growth_html_df.to_html(index=False, border=0, na_rep="—")
	else:
		if not stu_birth:
			growth_html = "<p style='color:#64748b;font-style:italic;'>Öğrencinin doğum tarihi belirtilmediği için yaşa göre persantil ve Z-skoru değerlendirmesi raporda yer almamaktadır.</p>"
		elif not can_eval_rep_age:
			growth_html = "<p style='color:#64748b;font-style:italic;'>Kişi 19 yaş üstü (yetişkin) olduğu için çocukluk çağı persantil değerlendirmesi raporda yer almamaktadır.</p>"
		else:
			growth_html = "<p style='color:#64748b;font-style:italic;'>Ölçüm kaydı bulunmadığından persantil özeti oluşturulmadı.</p>"

	alert_box = ""
	if not health_profile.empty:
		p_row = health_profile.iloc[0]
		alerji = clean_text(p_row.get("Alerjiler"))
		kronik = clean_text(p_row.get("Kronik_Hastalıklar"))
		gecmis = clean_text(p_row.get("Sağlık_Geçmişi"))
		ilac = clean_text(p_row.get("İlaçlar"))
		kan = clean_text(p_row.get("Kan_Grubu"))
		if alerji or kronik or gecmis or ilac or kan:
			alert_box = f"""
			<div style='background:#fef2f2;border:2px solid #ef4444;border-radius:6px;padding:12px;margin:14px 0;'>
				<strong style='color:#b91c1c;font-size:15px;'>⚠️ DİKKAT EDİLECEK SAĞLIK BİLGİLERİ:</strong><br>
				{f"<b>Kan Grubu:</b> {html.escape(kan)}<br>" if kan else ""}
				{f"<b>Alerjiler:</b> {html.escape(alerji)}<br>" if alerji else ""}
				{f"<b>Kronik Hastalıklar:</b> {html.escape(kronik)}<br>" if kronik else ""}
				{f"<b>Sağlık Geçmişi / Diğer Durumlar:</b> {html.escape(gecmis)}<br>" if gecmis else ""}
				{f"<b>Düzenli İlaç:</b> {html.escape(ilac)}" if ilac else ""}
			</div>
			"""

	page_break_style = "page-break-after: always; break-after: page;" if is_multi else ""
	return f"""
<div class='student-report-page' style='{page_break_style}'>
	<div class='header-box'>
		<h3>{inst_top}{f" · {inst_school}" if inst_school else ""}</h3>
		<h1>ÖĞRENCİ REVİR & SAĞLIK GELİŞİM RAPORU</h1>
		<h3>Rapor Tarihi: {date.today().strftime('%d.%m.%Y')}</h3>
	</div>

	<div class='student-card'>
		<div class='card-row'>
			<div class='card-item'><b>Adı Soyadı:</b> {safe_name}</div>
			<div class='card-item'><b>Sınıf / Şube:</b> {safe_class}</div>
			<div class='card-item'><b>Okul No:</b> {html.escape(str(stu_no or '—'))}</div>
			<div class='card-item'><b>Kayıt Türü:</b> {html.escape(str(stu_type or 'Öğrenci'))}</div>
			<div class='card-item'><b>Cinsiyet:</b> {html.escape(str(stu_sex or 'Belirtilmedi'))}</div>
			<div class='card-item'><b>Doğum Tarihi:</b> {html.escape(str(stu_birth or '—'))}</div>
		</div>
	</div>

	{alert_box}

	<h2>📋 Sağlık Profili ve Acil Bilgiler</h2>
	{profile_html}

	<h2>📏 Büyüme ve Persantil Taraması</h2>
	{growth_html}

	<h2>💓 Vital Ölçüm Geçmişi</h2>
	{vitals_html}

	<h2>🧪 Laboratuvar Tahlil Sonuçları</h2>
	{labs_html}

	<h2>⚖️ Boy, Kilo ve Baş Çevresi Ölçüm Geçmişi</h2>
	{measurements_html}

	<h2>🩺 Revir Başvuru ve Müdahale Kayıtları</h2>
	{visits_html}

	<h2>📝 Bakım Planı Kayıtları</h2>
	{care_plans_html}

	<h2>💉 Aşı ve Bağışıklama Kayıtları (T.C. Sağlık Bakanlığı)</h2>
	{vaccinations_html}

	<div class='signature-section'>
		<div class='sig-box'>
			<b>{inst_nurse}</b><br><br><br>
			(İmza / Kaşe)
		</div>
		<div class='sig-box'>
			<b>Veli Bilgilendirme Onayı</b><br><br><br>
			(Tarih / İmza)
		</div>
		<div class='sig-box'>
			<b>{inst_princ}</b><br><br><br>
			(Mühür / İmza)
		</div>
	</div>

	<p style='margin-top: 30px; font-size: 11px; color: #94a3b8; text-align: center;'>
		Bu belge okul reviri dahili izlem ve sağlık bilgilendirme tutanağıdır; tek başına resmi sağlık kurulu raporu veya tıbbi tanı yerine geçmez.
	</p>
</div>
"""


def render_reports() -> None:
	st.markdown("## Kişi ve Sınıf Raporları (Yazdırma & İndirme)")
	all_classes = [c for c in get_classes() if c != "Sivil / Diğer"]

	# Report mode selection: Single Person or Whole Class
	col_m1, col_m2 = st.columns([1.2, 2.5])
	with col_m1:
		report_mode = st.radio(
			"Yazdırma Kapsamı",
			["👤 Tek Kişi Raporu", "🏫 Tüm Sınıf Raporu (Toplu)"],
			horizontal=True,
			key="report_scope_mode",
		)

	app_cfg = get_app_settings()
	inst_top = html.escape(app_cfg.get("institution_top_header") or "T.C. MİLLÎ EĞİTİM BAKANLIĞI")
	inst_school = html.escape(app_cfg.get("school_name") or "")
	inst_nurse = html.escape(app_cfg.get("nurse_name") or "Okul Revir Sorumlusu")
	inst_princ = html.escape(app_cfg.get("principal_name") or "Okul Müdürü / İdare")

	css_header = f"""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>Sağlık ve Revir Raporu</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; padding: 30px; line-height: 1.5; }}
.header-box {{ border-bottom: 2px solid #176b68; padding-bottom: 15px; margin-bottom: 20px; text-align: center; }}
.header-box h1 {{ color: #176b68; margin: 0; font-size: 24px; text-transform: uppercase; }}
.header-box h3 {{ color: #64748b; margin: 4px 0 0; font-size: 14px; font-weight: normal; }}
.student-card {{ background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 20px; }}
.card-row {{ display: flex; justify-content: space-between; flex-wrap: wrap; font-size: 14px; }}
.card-item {{ margin: 4px 12px 4px 0; }}
h2 {{ color: #176b68; font-size: 16px; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px; margin-top: 24px; }}
table {{ border-collapse: collapse; width: 100%; margin: 10px 0 20px; font-size: 13px; page-break-inside: avoid; }}
th, td {{ border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }}
th {{ background-color: #f1f5f9; color: #334155; font-weight: 600; }}
tr:nth-child(even) {{ background-color: #f8fafc; }}
.signature-section {{ margin-top: 40px; display: flex; justify-content: space-between; page-break-inside: avoid; }}
.sig-box {{ width: 30%; text-align: center; border-top: 1px solid #94a3b8; padding-top: 10px; font-size: 12px; color: #475569; }}
button {{ padding: 12px 20px; background: #176b68; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; margin-bottom: 20px; }}
@media print {{
	button {{ display: none; }}
	body {{ padding: 0; font-size: 12px; }}
	.student-report-page {{ page-break-after: always; break-after: page; }}
	.header-box h1 {{ font-size: 20px; }}
	h2 {{ font-size: 14px; margin-top: 16px; }}
	th, td {{ padding: 6px 8px; }}
}}
</style>
</head>
<body>
<button onclick='window.print()'>🖨️ Tümünü Yazdır / PDF Olarak Kaydet</button>
"""

	if report_mode == "🏫 Tüm Sınıf Raporu (Toplu)":
		with col_m2:
			sel_cls_for_report = st.selectbox(
				"Raporu Yazdırılacak Sınıfı Seçin",
				all_classes,
				key="report_chosen_class",
			)

		class_students = student_options(sel_cls_for_report, person_type="Öğrenci")
		if not class_students:
			st.warning(f"**{sel_cls_for_report}** sınıfında kayıtlı öğrenci bulunamadı.")
		else:
			st.success(f"**{sel_cls_for_report}** sınıfında toplam **{len(class_students)} öğrenci** için rapor hazırlandı. Her öğrencinin raporu yazdırma sırasında ayrı bir sayfa olarak başlar.")

			cls_st_preview = [
				{
					"No": r["student_no"] or "—",
					"Öğrenci Adı": r["full_name"],
					"Cinsiyet": r["sex"] or "—",
					"Doğum Tarihi": r["birth_date"] or "—",
					"Telefon": r["guardian_phone"] or "—",
				}
				for r in class_students
			]
			st.dataframe(pd.DataFrame(cls_st_preview), width="stretch", hide_index=True)

			# Generate multi-student combined HTML report
			sections = []
			for idx, st_row in enumerate(class_students):
				is_last = (idx == len(class_students) - 1)
				sections.append(
					build_student_report_section_html(
						st_row, inst_top, inst_school, inst_nurse, inst_princ, is_multi=not is_last
					)
				)

			class_print_html = css_header + "\n".join(sections) + "</body>\n</html>"
			clean_cls_slug = re.sub(r"[^A-Za-z0-9_-]+", "_", sel_cls_for_report).strip("_")

			down_c1, down_c2 = st.columns([2, 1])
			with down_c1:
				batch_docx = build_class_batch_report_docx(sel_cls_for_report)
				st.download_button(
					f"📝 {sel_cls_for_report} - Tüm Sınıfın Word Raporunu İndir ({len(class_students)} Öğrenci, .docx)",
					data=batch_docx,
					file_name=f"{clean_cls_slug}_tum_sinif_saglik_raporu.docx",
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					type="primary",
					use_container_width=True,
				)
			with down_c2:
				st.caption("İndirilen Microsoft Word (.docx) dosyasında her öğrencinin sağlık ve revir dosyası ayrı bir sayfa olarak başlar; doğrudan Word'de düzenleyebilir ve yazdırabilirsiniz.")

	else:
		# Single person report
		with col_m2:
			student = select_shared_person("rep", label="Kişi raporu seçin")

		if student:
			stu_sex = student["sex"] if student else None
			stu_birth = student["birth_date"] if student else None
			render_shared_person_summary(student, section_key="rep")
			with connect_db() as connection:
				visits = pd.read_sql_query(
					"SELECT visit_date AS Tarih, visit_time AS Saat, complaint AS Başvuru, symptoms AS Belirtiler, observation AS Gözlem, action AS İşlem, guardian_contacted AS Veli_Bilgilendirildi, referral AS Yönlendirme, follow_up AS İzlem FROM visits WHERE student_id=? ORDER BY visit_date DESC, visit_time DESC, id DESC",
					connection, params=(student["id"],),
				)
				measurements = pd.read_sql_query(
					"SELECT id, measured_on AS Tarih, height_cm AS Boy_cm, weight_kg AS Kilo_kg, head_circ_cm AS Bas_Cevresi_cm, age_months AS Yas_ay, bmi AS BMI FROM measurements WHERE student_id=? ORDER BY measured_on ASC, id ASC",
					connection, params=(student["id"],),
				)
				vitals = pd.read_sql_query(
					"SELECT measured_at AS Tarih_Saat, systolic_mmhg AS Buyuk_Tansiyon, diastolic_mmhg AS Kucuk_Tansiyon, pulse_bpm AS Nabiz, temperature_c AS Ates_C, spo2_pct AS Saturasyon_SpO2, glucose_mg_dl AS Kan_Sekeri, glucose_context AS Olcum_Kosulu, notes AS Notlar FROM vital_readings WHERE student_id=? ORDER BY measured_at DESC, id DESC",
					connection, params=(student["id"],),
				)
				lab_results = pd.read_sql_query(
					"SELECT r.report_date AS Tarih, r.file_name AS Dosya, x.test_name AS Test, x.value AS Sonuç, x.unit AS Birim, x.ref_low AS Referans_Alt, x.ref_high AS Referans_Üst, x.lab_flag AS Rapor_İşareti FROM lab_results x JOIN lab_reports r ON r.id=x.report_id WHERE r.student_id=? ORDER BY r.report_date DESC, x.id",
					connection, params=(student["id"],),
				)
				health_profile = pd.read_sql_query(
					"""SELECT COALESCE(h.guardian_name, s.guardian_name) AS Veli,
					   COALESCE(h.guardian_phone, s.guardian_phone) AS Telefon, h.allergies AS Alerjiler,
					   h.chronic_conditions AS Kronik_Hastalıklar, h.medical_history AS Sağlık_Geçmişi,
					   h.medications AS İlaçlar, h.blood_group AS Kan_Grubu,
					   h.emergency_notes AS Acil_Notlar, h.updated_at AS Güncelleme, h.reviewed_by AS Doğrulayan
				   FROM students s LEFT JOIN health_profiles h ON h.student_id=s.id
				   WHERE s.id=? AND (h.student_id IS NOT NULL OR s.guardian_name IS NOT NULL OR s.guardian_phone IS NOT NULL)""",
					connection, params=(student["id"],),
				)
				care_plans = pd.read_sql_query(
					"SELECT created_at AS Tarih, nursing_diagnosis AS Hemşirelik_Değerlendirmesi, observed_facts AS Olgular, goals AS Hedefler, interventions AS Girişimler, monitoring AS İzlem, escalation AS Yönlendirme, status AS Durum, reviewed_by AS Sorumlu FROM care_plans WHERE student_id=? ORDER BY created_at DESC",
					connection, params=(student["id"],),
				)
			lab_results = assess_lab_history(lab_results)
			report_growth_df = None
			person_meas = get_person_latest_measurements(student["id"])
			if person_meas["has_data"] and student["sex"]:
				latest_m = None
				if person_meas["latest_measured_on"] and student["birth_date"]:
					try:
						latest_m = age_in_months(student["birth_date"], date.fromisoformat(person_meas["latest_measured_on"]))
					except Exception:
						latest_m = None
				if latest_m is not None:
					report_growth_df = get_clinical_growth_summary_table(
						person_meas["height_cm"], person_meas["weight_kg"], person_meas["bmi"],
						student["sex"], latest_m, head_circ_cm=person_meas["head_circ_cm"]
					)
			st.markdown(f"### {student['person_type']} · {student['full_name']} · {student['class_name']}")
			st.write(f"Okul no: {student['student_no'] or '—'} · Cinsiyet: {student['sex'] or 'Belirtilmedi'} · Doğum tarihi: {student['birth_date'] or '—'}")
			st.markdown("#### Sağlık profili")
			if health_profile.empty:
				st.caption("Sağlık profili kaydı bulunmuyor.")
			else:
				st.dataframe(health_profile, width="stretch", hide_index=True)
			st.markdown("#### Bakım planı taslakları")
			if care_plans.empty:
				st.caption("Bakım planı kaydı bulunmuyor.")
			else:
				st.dataframe(care_plans, width="stretch", hide_index=True)
			st.markdown("#### 🩺 Klinik Büyüme ve Persantil Değerlendirmesi (WHO / ÇEDD Standardı)")
			if report_growth_df is not None and not report_growth_df.empty:
				st.dataframe(
					report_growth_df,
					width="stretch",
					hide_index=True,
					column_config=get_clinical_growth_column_config(),
				)
				st.caption("Dünya Sağlık Örgütü (WHO) LMS büyüme referansına dayalı antropometrik tarama ve Z-skoru (SDS) raporu.")
			else:
				st.caption("Persantil ve Z-skoru raporu için öğrenci profilinde doğum tarihi, cinsiyet ve en az bir ölçüm kaydı bulunmalıdır.")
			st.markdown("#### Vital geçmişi")
			if vitals.empty:
				st.caption("Vital kayıt bulunmuyor.")
			else:
				st.dataframe(vitals, width="stretch", hide_index=True)
			st.markdown("#### Tahlil geçmişi")
			if lab_results.empty:
				st.caption("Tahlil kaydı bulunmuyor.")
			else:
				st.dataframe(lab_results, width="stretch", hide_index=True)
			st.markdown("#### Ölçüm ve persantil geçmişi")
			if measurements.empty:
				st.caption("Ölçüm kaydı bulunmuyor.")
			else:
				rep_enriched = enrich_measurements_with_percentiles(measurements, stu_sex, stu_birth)
				m_cols = [
					"Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Boy_Dilimi", "Boy_Farki",
					"Kilo_kg", "Kilo_Z", "Kilo_Persantil", "Kilo_Dilimi", "Kilo_Farki",
					"Ideal_Kilo_Araligi",
					"Bas_Cevresi_cm", "Bas_Cevresi_Z", "Bas_Cevresi_Persantil",
					"BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"
				]
				avail_cols = [c for c in m_cols if c in rep_enriched.columns]
				rep_disp = rep_enriched[avail_cols].rename(columns={
					"Yas_ay": "Yaş (Ay)", "Boy_cm": "Boy (cm)", "Boy_Z": "Boy Z-Skor", "Boy_Persantil": "Boy %",
					"Boy_Dilimi": "Boy Dilimi", "Boy_Farki": "Boy Değişimi", "Kilo_kg": "Kilo (kg)",
					"Kilo_Z": "Kilo Z-Skor", "Kilo_Persantil": "Kilo %", "Kilo_Dilimi": "Kilo Dilimi",
					"Kilo_Farki": "Kilo Değişimi", "Ideal_Kilo_kg": "İdeal Kilo (kg)", "Ideal_Kilo_Araligi": "İdeal Kilo Aralığı", "Bas_Cevresi_cm": "Baş Çevresi (cm)",
					"Bas_Cevresi_Z": "Baş Çevresi Z-Skor", "Bas_Cevresi_Persantil": "Baş Çevresi %",
					"BMI": "VKİ (BMI)", "BMI_Z": "VKİ Z-Skor", "BMI_Persantil": "VKİ %",
					"Gelisim_Degerlendirmesi": "Gelişim Değerlendirmesi"
				})
				st.dataframe(
					rep_disp,
					width="stretch",
					hide_index=True,
					column_config=get_measurement_history_column_config(),
				)
				if len(measurements) > 1:
					first, last = measurements.iloc[0], measurements.iloc[-1]
					if pd.notna(last["Boy_cm"]) and pd.notna(first["Boy_cm"]) and pd.notna(last["Kilo_kg"]) and pd.notna(first["Kilo_kg"]):
						st.info(f"İlk ölçümden son ölçüme değişim: boy {last['Boy_cm'] - first['Boy_cm']:+.1f} cm, kilo {last['Kilo_kg'] - first['Kilo_kg']:+.1f} kg.")
			st.markdown("#### Revir geçmişi")
			if visits.empty:
				st.caption("Revir kaydı bulunmuyor.")
			else:
				st.dataframe(visits, width="stretch", hide_index=True)

			single_body = build_student_report_section_html(student, inst_top, inst_school, inst_nurse, inst_princ, is_multi=False)
			print_html = css_header + single_body + "</body>\n</html>"
			print_name = re.sub(r"[^A-Za-z0-9_-]+", "_", student["full_name"]).strip("_") or "kisi"
			single_docx = build_student_report_docx(student)
			st.download_button(
				"📝 Kişi Sağlık Raporunu İndir (.docx)",
				data=single_docx,
				file_name=f"{print_name}_saglik_raporu.docx",
				mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
				type="primary",
				use_container_width=True,
			)
			st.caption("İndirilen Microsoft Word (.docx) dosyasını Microsoft Office Word uygulamasında açıp doğrudan düzenleyebilir veya yazdırabilirsiniz.")

	with connect_db() as connection:
		all_students = get_students_enriched().drop(columns=["id"], errors="ignore")
		all_visits = pd.read_sql_query(
			"SELECT c.name AS sinif, s.full_name AS ogrenci, v.visit_date AS tarih, v.visit_time AS saat, v.complaint AS basvuru, v.symptoms AS belirtiler, v.observation AS gozlem, v.action AS islem, v.guardian_contacted AS veli_bilgilendirildi, v.referral AS yonlendirme, v.follow_up AS izlem FROM visits v JOIN students s ON s.id=v.student_id JOIN classes c ON c.id=s.class_id ORDER BY v.visit_date DESC, v.visit_time DESC", connection
		)
		raw_all_meas = pd.read_sql_query(
			"""SELECT c.name AS sinif, s.full_name AS ogrenci, s.sex AS cinsiyet, s.birth_date AS dogum_tarihi,
					  m.measured_on AS tarih, m.height_cm AS boy_cm, m.weight_kg AS kilo_kg, m.head_circ_cm AS bas_cevresi_cm, m.age_months AS yas_ay, m.bmi
			   FROM measurements m JOIN students s ON s.id=m.student_id JOIN classes c ON c.id=s.class_id
			   ORDER BY m.measured_on DESC, m.id DESC""", connection
		)
		all_vitals = pd.read_sql_query(
			"SELECT c.name AS sinif, s.full_name AS ogrenci, v.measured_at AS tarih_saat, v.systolic_mmhg AS buyuk_tansiyon, v.diastolic_mmhg AS kucuk_tansiyon, v.pulse_bpm AS nabiz, v.temperature_c AS ates_c, v.spo2_pct AS saturasyon_spo2, v.glucose_mg_dl AS kan_sekeri, v.glucose_context AS olcum_kosulu, v.notes AS notlar FROM vital_readings v JOIN students s ON s.id=v.student_id JOIN classes c ON c.id=s.class_id ORDER BY v.measured_at DESC", connection
		)
		all_labs = pd.read_sql_query(
			"SELECT c.name AS sinif, s.full_name AS ogrenci, r.report_date AS tarih, r.file_name AS dosya, x.test_name AS test, x.value AS sonuc, x.unit AS birim, x.ref_low AS referans_alt, x.ref_high AS referans_ust, x.lab_flag AS rapor_isareti FROM lab_results x JOIN lab_reports r ON r.id=x.report_id JOIN students s ON s.id=r.student_id JOIN classes c ON c.id=s.class_id ORDER BY r.report_date DESC", connection
		)
		all_profiles = pd.read_sql_query(
			"SELECT s.person_type AS kayit_turu, c.name AS sinif, s.full_name AS kisi, COALESCE(h.guardian_name,s.guardian_name) AS veli, COALESCE(h.guardian_phone,s.guardian_phone) AS telefon, h.allergies AS alerjiler, h.chronic_conditions AS kronik_hastaliklar, h.medical_history AS saglik_gecmisi, h.medications AS ilaclar, h.blood_group AS kan_grubu, h.emergency_notes AS acil_notlar, h.updated_at AS guncelleme, h.reviewed_by AS dogrulayan FROM students s JOIN classes c ON c.id=s.class_id LEFT JOIN health_profiles h ON h.student_id=s.id WHERE h.student_id IS NOT NULL OR s.guardian_name IS NOT NULL OR s.guardian_phone IS NOT NULL ORDER BY c.name, s.full_name", connection
		)
		all_plans = pd.read_sql_query(
			"SELECT s.person_type AS kayit_turu, c.name AS sinif, s.full_name AS kisi, p.created_at AS tarih, p.nursing_diagnosis AS hemsirelik_degerlendirmesi, p.observed_facts AS olgular, p.goals AS hedefler, p.interventions AS girisimler, p.monitoring AS izlem, p.escalation AS yonlendirme, p.status AS durum, p.reviewed_by AS sorumlu FROM care_plans p JOIN students s ON s.id=p.student_id JOIN classes c ON c.id=s.class_id ORDER BY p.created_at DESC", connection
		)

	meas_export_rows = []
	for _, mr in raw_all_meas.iterrows():
		m_ay = int(mr["yas_ay"]) if pd.notna(mr["yas_ay"]) else (
			age_in_months(mr["dogum_tarihi"], date.fromisoformat(str(mr["tarih"])[:10])) if mr["dogum_tarihi"] else None
		)
		h_v = float(mr["boy_cm"]) if pd.notna(mr["boy_cm"]) else None
		w_v = float(mr["kilo_kg"]) if pd.notna(mr["kilo_kg"]) else None
		hc_v = float(mr["bas_cevresi_cm"]) if (pd.notna(mr.get("bas_cevresi_cm")) and float(mr["bas_cevresi_cm"]) > 0) else None
		b_v = float(mr["bmi"]) if pd.notna(mr["bmi"]) else ((w_v / ((h_v / 100.0) ** 2)) if (h_v and w_v) else None)
		h_ev = evaluate_measurement(h_v, "height", mr["cinsiyet"], m_ay) if (mr["cinsiyet"] and m_ay is not None and h_v) else None
		w_ev = evaluate_measurement(w_v, "weight", mr["cinsiyet"], m_ay) if (mr["cinsiyet"] and m_ay is not None and w_v) else None
		b_ev = evaluate_measurement(b_v, "bmi", mr["cinsiyet"], m_ay) if (mr["cinsiyet"] and m_ay is not None and b_v) else None
		hc_ev = evaluate_measurement(hc_v, "head_circ", mr["cinsiyet"], m_ay) if (mr["cinsiyet"] and m_ay is not None and hc_v) else None
		meas_export_rows.append({
			"sinif": mr["sinif"],
			"ogrenci": mr["ogrenci"],
			"cinsiyet": mr["cinsiyet"],
			"tarih": mr["tarih"],
			"yas_ay": m_ay,
			"boy_cm": round(h_v, 1) if h_v is not None else None,
			"boy_z_sds": round(h_ev["z"], 2) if h_ev else None,
			"boy_persantil": round(h_ev["percentile"], 1) if h_ev else None,
			"kilo_kg": round(w_v, 1) if w_v is not None else None,
			"kilo_z_sds": round(w_ev["z"], 2) if w_ev else None,
			"kilo_persantil": round(w_ev["percentile"], 1) if w_ev else None,
			"bas_cevresi_cm": round(hc_v, 1) if hc_v is not None else None,
			"bas_cevresi_z_sds": round(hc_ev["z"], 2) if hc_ev else None,
			"bas_cevresi_persantil": round(hc_ev["percentile"], 1) if hc_ev else None,
			"vki_bmi": round(b_v, 2) if b_v is not None else None,
			"vki_z_sds": round(b_ev["z"], 2) if b_ev else None,
			"vki_persantil": round(b_ev["percentile"], 1) if b_ev else None,
			"gelisim_durumu": b_ev["status"] if b_ev else "—",
		})
	all_measurements = pd.DataFrame(meas_export_rows) if meas_export_rows else raw_all_meas

	st.download_button("Tüm kişi, persantil, revir, vital, tahlil ve bakım planlarını Excel indir", workbook_download(all_students, all_visits, all_measurements, all_vitals, all_labs, all_profiles, all_plans),
					   "revir_raporu.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_academic_year_settings() -> None:
	st.markdown("### 📅 Eğitim Öğretim Yılı ve Arşiv Yönetimi")
	st.caption("Eğitim öğretim yıllarını yönetebilir, arşive kaldırılan geçmiş yılların kayıtlarına tam bağımsız olarak göz atabilir ve yeni eğitim öğretim yılları başlatabilirsiniz. Her eğitim yılının öğrenci kadrosu ve revir kayıtları tamamen bağımsız klasör ve veritabanında saklanır; mezuniyet ve yeni öğrenci kayıtları yıllar arasında asla birbirine karışmaz.")

	active_ay = get_active_academic_year()
	all_years = get_all_academic_years()
	today_iso = date.today().isoformat()
	is_arch = active_ay["end_date"] < today_iso

	# 1. Quick Year Selector & Instant Activation Panel
	with st.container(border=True):
		st.markdown("#### 🎓 Çalışılacak / İncelenecek Eğitim Yılı Seçimi")
		st.caption("Aşağıdaki listeden dilediğiniz eğitim öğretim yılını seçtiğinizde, sistem anında o yıla geçer ve tüm sayfalar (Öğrenciler, Revir Kaydı, Vital, Ölçüm vb.) yalnızca seçilen yılın veritabanını gösterir:")

		# Format options with badges
		year_options = []
		for y in all_years:
			y_is_arch = y["end_date"] < today_iso
			if y["id"] == active_ay["id"]:
				badge = "🟢 ŞU ANDA SEÇİLİ"
			elif y_is_arch:
				badge = "🏛️ GEÇMİŞ ARŞİV"
			else:
				badge = "⚪ AKTİF / GELECEK"
			year_options.append({
				"id": y["id"],
				"name": y["name"],
				"label": f"{y['name']}  ({badge})  [{y['start_date']} — {y['end_date']}]",
				"is_arch": y_is_arch,
			})

		cur_opt_idx = 0
		for i, opt in enumerate(year_options):
			if opt["id"] == active_ay["id"]:
				cur_opt_idx = i
				break

		col_sel_y, col_info_y = st.columns([2.2, 1.8], gap="medium")
		with col_sel_y:
			sel_label = st.radio(
				"Eğitim Yılı Seçin",
				options=[opt["label"] for opt in year_options],
				index=cur_opt_idx,
				key="ay_active_year_radio",
				label_visibility="collapsed"
			)
			picked_opt = next(opt for opt in year_options if opt["label"] == sel_label)
			if picked_opt["id"] != active_ay["id"]:
				set_active_academic_year(picked_opt["id"])
				st.success(f"'{picked_opt['name']}' eğitim yılına geçildi. Tüm veriler güncellendi!")
				st.rerun()

		with col_info_y:
			with connect_db() as cur_conn:
				stu_cnt = cur_conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
				cls_cnt = cur_conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0]
				vis_cnt = cur_conn.execute("SELECT COUNT(*) FROM visits").fetchone()[0]
				meas_cnt = cur_conn.execute("SELECT COUNT(*) FROM measurements").fetchone()[0]
			k_dir = get_current_person_data_dir()
			j_cnt = len(list(k_dir.glob("kisi_*.json"))) if k_dir.exists() else 0

			if is_arch:
				st.warning(f"🏛️ **Arşiv Modundasınız:** Şu anda geçmiş **{active_ay['name']}** eğitim yılı arşivini inceliyorsunuz.")
			else:
				st.success(f"🟢 **Yürürlükteki Yıl:** **{active_ay['name']}** eğitim yılı aktif durumdadır.")

			m1, m2, m3 = st.columns(3)
			m1.metric("Kayıtlı Öğrenci", stu_cnt)
			m2.metric("Sınıf Sayısı", cls_cnt)
			m3.metric("Revir Viziti", vis_cnt)
			st.caption(f"📁 **Veri Klasörü:** `{active_ay['folder_name']}/`")

	# 2. Start New Academic Year (New Folder, Isolated DB, Student Policies)
	with st.expander("➕ Yeni Eğitim Öğretim Yılı Başlat (Yeni Klasör & Temiz Sayfa)", expanded=True):
		st.markdown("##### 🚀 Yeni Bir Eğitim Öğretim Yılı Oluştur")
		st.caption("Her yeni eğitim öğretim yılında yeni öğrenciler geldiği ve son sınıflar mezun olduğu için yeni bir bağımsız veritabanı açılır. Böylece önceki yılların verileriyle asla karışmaz. WHO persantil referansları ve revir şablonları aynen korunur.")

		ny_c1, ny_c2, ny_c3 = st.columns([1.5, 1.2, 1.2])
		with ny_c1:
			new_y_name = st.text_input("Eğitim Yılı Adı", placeholder="Örn: 2026-2027", key="new_ay_name_inp")
		with ny_c2:
			new_y_start = st.date_input("Başlangıç Tarihi", value=date.today(), key="new_ay_start_inp")
		with ny_c3:
			new_y_end = st.date_input("Bitiş Tarihi", value=date.today() + timedelta(days=280), key="new_ay_end_inp")

		st.markdown("###### 👥 Öğrenci ve Sınıf Başlangıç Politikası:")
		policy_options = {
			"fresh": "🟢 Sıfırdan Temiz Sayfa (Tavsiye Edilen — Yeni Öğrenci Girişi): Eski yıla ait hiçbir öğrenci ve revir kaydı yeni yıla aktarılmaz. Yeni kayıt olacak öğrenciler için tertemiz sıfır kilometre bir revir defteri açılır.",
			"promote": "🎓 Sınıf Atlatma & Mezuniyet (Terfi Modu): Öğrencileri bir üst sınıfa aktarır (Örn: 1-A -> 2-A, 4-A -> Mezun). Vizit ve muayene kayıtları temizlenerek yeni yıl için sıfırlanır.",
			"clone": "📋 Mevcut Öğrenci Kadrosunu Klonla: Öğrenci ve sınıf kadrosu aynen korunur, revir vizitleri ve muayene kayıtları sıfırlanır.",
			"exact": "💾 Tam Kopya: Tüm öğrenci ve muayene kayıtlarını olduğu gibi yeni yıla kopyalar."
		}
		selected_policy_key = st.radio(
			"Öğrenci Politikası Seçimi",
			list(policy_options.keys()),
			format_func=lambda k: policy_options[k],
			index=0,
			key="new_ay_policy_radio"
		)

		if st.button("🚀 Yeni Eğitim Öğretim Yılını Başlat ve Klasörünü Aç", type="primary", use_container_width=True, key="btn_create_new_ay"):
			if not new_y_name.strip():
				st.error("Lütfen bir eğitim yılı adı girin (Örn: 2026-2027).")
			elif any(y["name"] == new_y_name.strip() for y in all_years):
				st.error("Bu isimde bir eğitim öğretim yılı zaten kayıtlı!")
			else:
				new_id = create_academic_year(
					new_y_name.strip(),
					new_y_start.isoformat(),
					new_y_end.isoformat(),
					student_policy=selected_policy_key
				)
				set_active_academic_year(new_id)
				st.success(f"'{new_y_name.strip()}' eğitim yılı başarıyla oluşturuldu, bağımsız veritabanı hazırlandı ve aktif eğitim yılı yapıldı!")
				st.rerun()

	# 3. Student Transfer Tool
	with st.expander("👥 Önceki Eğitim Yılından Öğrenci Aktar (Kayıt Yenileme / Transfer)", expanded=False):
		st.caption("Yeni bir eğitim yılına sıfırdan başlasanız bile, okulda eğitime devam eden öğrencileri önceki yıldan tek tıkla bu yıla aktarabilirsiniz. Vizit geçmişleri eski yılda güvenle saklanır, yeni yıla temiz bir sağlık profiliyle aktarılır.")
		other_years = [y for y in all_years if y["id"] != active_ay["id"]]
		if not other_years:
			st.info("Aktarım yapılabilecek başka bir eğitim yılı bulunmuyor.")
		else:
			src_y_col1, src_y_col2 = st.columns(2)
			with src_y_col1:
				src_y_pick = st.selectbox(
					"Kaynak Eğitim Yılı Seçin",
					options=[y["name"] for y in other_years],
					key="transfer_src_year_select"
				)
				src_y_obj = next(y for y in other_years if y["name"] == src_y_pick)
			with src_y_col2:
				src_db_file = APP_DIR / src_y_obj["folder_name"] / "revir_defteri.db"
				src_students = []
				if src_db_file.exists():
					s_conn = sqlite3.connect(src_db_file)
					s_conn.row_factory = sqlite3.Row
					try:
						rows = s_conn.execute("SELECT s.id, s.student_no, s.full_name, c.name as class_name FROM students s JOIN classes c ON c.id=s.class_id ORDER BY c.name, s.full_name").fetchall()
						src_students = [dict(r) for r in rows]
					finally:
						s_conn.close()
				st.caption(f"ℹ️ '{src_y_pick}' yılında toplam **{len(src_students)}** öğrenci bulundu.")

			if src_students:
				all_src_classes = sorted(list({s["class_name"] for s in src_students}))
				sel_src_cls = st.multiselect("Sınıfa Göre Filtrele (İsteğe Bağlı):", options=all_src_classes, default=[], key="transfer_filter_classes")
				filtered_src_students = [s for s in src_students if not sel_src_cls or s["class_name"] in sel_src_cls]

				stu_options = {s["id"]: f"{s['full_name']} ({s['class_name']}{f' - No: {s["student_no"]}' if s['student_no'] else ''})" for s in filtered_src_students}
				sel_stu_ids = st.multiselect(
					f"Aktarılacak Öğrencileri Seçin ({len(filtered_src_students)} öğrenci listeleniyor):",
					options=list(stu_options.keys()),
					format_func=lambda x: stu_options[x],
					default=list(stu_options.keys()) if len(filtered_src_students) <= 20 else [],
					key="transfer_picked_students"
				)

				t_col1, t_col2 = st.columns([2, 1])
				with t_col1:
					st.caption(f"Seçilen öğrenci sayısı: **{len(sel_stu_ids)}**")
				with t_col2:
					if st.button("👥 Seçili Öğrencileri Bu Yıla Aktar", type="primary", disabled=(len(sel_stu_ids) == 0), key="btn_exec_student_transfer"):
						done_cnt = transfer_students_from_year(src_y_obj["id"], sel_stu_ids)
						st.success(f"✅ {done_cnt} öğrenci ve sağlık profilleri '{active_ay['name']}' eğitim yılına başarıyla aktarıldı!")
						st.rerun()

	# 4. Edit Dates of Active Year
	with st.expander("✏️ Seçili Eğitim Yılının Tarihlerini Düzenle", expanded=False):
		st.caption(f"'{active_ay['name']}' eğitim öğretim yılı için resmi başlangıç ve bitiş tarihlerini güncelleyebilirsiniz:")
		ed_c1, ed_c2 = st.columns(2)
		with ed_c1:
			try:
				def_sd = date.fromisoformat(active_ay["start_date"][:10])
			except Exception:
				def_sd = date(2025, 9, 8)
			new_start_d = st.date_input("🗓️ Başlangıç Tarihi", value=def_sd, key="ay_edit_start_date")
		with ed_c2:
			try:
				def_ed = date.fromisoformat(active_ay["end_date"][:10])
			except Exception:
				def_ed = date(2026, 6, 19)
			new_end_d = st.date_input("🏁 Bitiş Tarihi", value=def_ed, key="ay_edit_end_date")
		if st.button("💾 Tarihleri Güncelle", type="primary", key="btn_save_ay_dates"):
			update_academic_year(active_ay["id"], active_ay["name"], new_start_d.isoformat(), new_end_d.isoformat())
			st.success("Eğitim öğretim yılı tarihleri güncellendi.")
			st.rerun()

	# 5. Year Database Backup & Download
	with st.expander("💾 Bu Eğitim Yılının Veritabanını İndir / Yedekle", expanded=False):
		st.caption("Seçili eğitim yılına ait bağımsız SQLite veritabanını (`revir_defteri.db`) tek tıkla bilgisayarınıza indirebilirsiniz:")
		cur_db_p = get_current_db_path()
		if cur_db_p.exists():
			db_bytes = cur_db_p.read_bytes()
			st.download_button(
				f"💾 {active_ay['name']} Veritabanı Yedeğini İndir ({format_file_size(len(db_bytes))})",
				data=db_bytes,
				file_name=f"revir_defteri_{re.sub(r'[^a-zA-Z0-9]+', '_', active_ay['name'])}.db",
				mime="application/octet-stream",
				type="primary",
				key="btn_download_year_db"
			)

	# 6. Reset Current Year Data
	with st.expander("⚠️ Bu Eğitim Yılının Verilerini Sıfırla (Temiz Sayfa Aç)", expanded=False):
		st.warning(f"🚨 **DİKKAT:** Bu işlem yalnızca seçili eğitim yılına (**{active_ay['name']}**) ait öğrencileri, sınıfları, revir başvurularını ve ölçümleri kalıcı olarak siler. Diğer eğitim yıllarının veya arşivlerin verileri bundan ETKİLENMEZ.")
		confirm_reset = st.checkbox(f"Evet, '{active_ay['name']}' eğitim yılındaki tüm öğrenci ve revir kayıtlarını temizlemek ve sıfırdan başlamak istiyorum.", key="chk_confirm_reset_ay")
		if st.button(f"🗑️ '{active_ay['name']}' Verilerini Kalıcı Olarak Sıfırla", type="primary", disabled=not confirm_reset, key="btn_exec_reset_ay"):
			reset_current_year_students_and_records()
			st.success(f"'{active_ay['name']}' eğitim yılı başarıyla sıfırlandı. Yeni öğrenci kayıtları için hazır!")
			st.rerun()

	# 7. Delete Academic Year and All Data Permanently
	with st.expander("🗑️ Eğitim Öğretim Yılını ve Tüm Verilerini Sil (Kalıcı Olarak Kaldır)", expanded=False):
		st.caption("Kullanılmayan, yanlış açılan veya arşivi artık saklanmak istenmeyen bir eğitim öğretim yılını, bağımsız veritabanını (`revir_defteri.db`) ve o yıla ait tüm öğrenci dosyalarını sistemden kalıcı olarak silebilirsiniz.")
		if len(all_years) <= 1:
			st.info("ℹ️ Sistemde yalnızca bir eğitim öğretim yılı kayıtlı olduğundan silme işlemi yapılamaz. Yeni bir eğitim yılı başlattıktan sonra bu yılı silebilirsiniz.")
		else:
			del_y_names = [f"{y['name']}{' (Şu Anda Seçili)' if y['id'] == active_ay['id'] else ''}{' [Arşiv]' if y['end_date'] < today_iso else ''}" for y in all_years]
			del_sel_name = st.selectbox("Silinecek Eğitim Yılını Seçin:", del_y_names, key="sel_delete_academic_year")
			del_target_y = all_years[del_y_names.index(del_sel_name)]

			# Query stats of year to be deleted
			del_db_path = APP_DIR / del_target_y["folder_name"] / "revir_defteri.db"
			del_stu_cnt, del_vis_cnt = 0, 0
			if del_db_path.exists():
				try:
					dconn = sqlite3.connect(del_db_path)
					del_stu_cnt = dconn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
					del_vis_cnt = dconn.execute("SELECT COUNT(*) FROM visits").fetchone()[0]
					dconn.close()
				except Exception:
					pass
			del_f_path = APP_DIR / del_target_y["folder_name"] / "kisi_verileri"
			del_json_cnt = len(list(del_f_path.glob("kisi_*.json"))) if del_f_path.exists() else 0

			st.error(
				f"⚠️ **Kalıcı Silme Uyarısı:** '{del_target_y['name']}' eğitim yılını sildiğinizde, `{del_target_y['folder_name']}/` klasörü, "
				f"**{del_stu_cnt}** öğrenci kaydı, **{del_vis_cnt}** revir viziti ve **{del_json_cnt}** kişi dosyası geri döndürülemez şekilde silinecektir."
			)
			if del_target_y["id"] == active_ay["id"]:
				st.warning(f"📌 Şu anda aktif olan yılı siliyorsunuz. Silme işlemi tamamlandığında sistem otomatik olarak kalan başka bir eğitim yılına geçiş yapacaktır.")

			confirm_del_ay = st.checkbox(
				f"Evet, '{del_target_y['name']}' eğitim öğretim yılını, veritabanını ve klasörünü kalıcı olarak silmek istiyorum.",
				key="chk_confirm_delete_academic_year"
			)
			if st.button(f"🗑️ '{del_target_y['name']}' Eğitim Yılını Kalıcı Olarak Sil", type="primary", disabled=not confirm_del_ay, key="btn_exec_delete_academic_year"):
				ok, msg = delete_academic_year(del_target_y["id"])
				if ok:
					st.success(msg)
					st.rerun()
				else:
					st.error(msg)

	# 8. Comparative Statistics Table
	st.markdown("##### 📊 Kayıtlı Eğitim Öğretim Yılları ve Karşılaştırmalı Veri İstatistikleri")
	ay_table_data = []
	for y in all_years:
		y_start = y["start_date"]
		y_end = y["end_date"]
		ydb_path = APP_DIR / y["folder_name"] / "revir_defteri.db"
		if ydb_path.exists():
			try:
				yconn = sqlite3.connect(ydb_path)
				y_stu = yconn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
				y_vis = yconn.execute("SELECT COUNT(*) FROM visits").fetchone()[0]
				y_meas = yconn.execute("SELECT COUNT(*) FROM measurements").fetchone()[0]
				y_scr = yconn.execute("SELECT COUNT(*) FROM school_screenings").fetchone()[0]
				yconn.close()
			except Exception:
				y_stu, y_vis, y_meas, y_scr = 0, 0, 0, 0
		else:
			y_stu, y_vis, y_meas, y_scr = 0, 0, 0, 0

		y_folder = APP_DIR / y["folder_name"] / "kisi_verileri"
		p_cnt = len(list(y_folder.glob("kisi_*.json"))) if y_folder.exists() else 0
		is_cur = "🟢 SEÇİLİ" if y["id"] == active_ay["id"] else ("🏛️ Arşiv" if y_end < today_iso else "⚪ Gelecek")
		ay_table_data.append({
			"Durum": is_cur,
			"Eğitim Yılı": y["name"],
			"Başlangıç": y_start,
			"Bitiş": y_end,
			"Öğrenci Sayısı": y_stu,
			"Revir Viziti": y_vis,
			"Ölçüm": y_meas,
			"Tarama": y_scr,
			"Kişi Dosyası (JSON)": p_cnt,
			"Veri Klasörü": y["folder_name"]
		})
	st.dataframe(pd.DataFrame(ay_table_data), use_container_width=True, hide_index=True)


def get_evrak_cover_font(size: int = 14, bold: bool = False):
	font_paths = [
		r"C:\Windows\Fonts\segoeui.ttf",
		r"C:\Windows\Fonts\arial.ttf",
		r"C:\Windows\Fonts\tahoma.ttf",
	]
	if bold:
		font_paths = [
			r"C:\Windows\Fonts\segouib.ttf",
			r"C:\Windows\Fonts\arialbd.ttf",
			r"C:\Windows\Fonts\tahomabd.ttf",
		] + font_paths
	from PIL import ImageFont
	for p in font_paths:
		if Path(p).exists():
			try:
				return ImageFont.truetype(p, size)
			except Exception:
				pass
	return ImageFont.load_default()


def get_document_cover_path(file_path: Path) -> Path:
	thumb_dir = EVRAKLAR_DIR / ".thumbnails"
	thumb_dir.mkdir(parents=True, exist_ok=True)
	f_hash = sha256(file_path.name.encode("utf-8")).hexdigest()[:16]
	return thumb_dir / f"cov_{f_hash}.jpg"


def get_or_create_document_cover(file_path: Path, force: bool = False) -> Path | None:
	if not file_path.exists() or not file_path.is_file():
		return None
	thumb_path = get_document_cover_path(file_path)
	if not force and thumb_path.exists():
		try:
			if thumb_path.stat().st_mtime >= file_path.stat().st_mtime and thumb_path.stat().st_size > 0:
				return thumb_path
		except Exception:
			pass

	ext = file_path.suffix.lower()
	import pymupdf
	from PIL import Image, ImageDraw
	import zipfile

	try:
		# 1. PDF: Render exact 1st page at 72 DPI, save as compact JPEG
		if ext == ".pdf":
			doc = pymupdf.open(file_path)
			if len(doc) > 0:
				page = doc[0]
				pix = page.get_pixmap(dpi=72)
				pix.pil_save(str(thumb_path), format="JPEG", quality=80)
				doc.close()
				del doc, page, pix
				return thumb_path
			doc.close()
			del doc

		# 2. Images: Thumbnail as JPEG
		elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
			img = Image.open(file_path)
			img.thumbnail((320, 440), Image.Resampling.LANCZOS)
			img.convert("RGB").save(str(thumb_path), format="JPEG", quality=80)
			img.close()
			del img
			return thumb_path

		# 3. Word Documents (.docx)
		elif ext == ".docx":
			has_embedded = False
			try:
				with zipfile.ZipFile(file_path) as z:
					for name in z.namelist():
						if name.startswith("word/media/image1.") and any(name.endswith(e) for e in [".png", ".jpg", ".jpeg"]):
							img_data = z.read(name)
							test_im = Image.open(io.BytesIO(img_data))
							test_im.thumbnail((320, 440), Image.Resampling.LANCZOS)
							test_im.convert("RGB").save(str(thumb_path), format="JPEG", quality=80)
							test_im.close()
							del test_im
							has_embedded = True
							break
			except Exception:
				pass

			if not has_embedded:
				w, h = 320, 440
				img = Image.new("RGB", (w, h), "#ffffff")
				draw = ImageDraw.Draw(img)
				draw.rectangle([(0, 0), (w - 1, h - 1)], outline="#cbd5e1", width=2)
				draw.rectangle([(0, 0), (w, 32)], fill="#0284c7")
				draw.text((12, 8), "📘 MICROSOFT WORD BELGESİ", fill="#ffffff", font=get_evrak_cover_font(11, bold=True))

				title_text = file_path.stem if len(file_path.stem) <= 30 else file_path.stem[:27] + "..."
				draw.text((14, 42), title_text, fill="#0f172a", font=get_evrak_cover_font(13, bold=True))
				draw.line([(14, 64), (w - 14, 64)], fill="#e2e8f0", width=1)

				content_lines = []
				try:
					d = docx.Document(file_path)
					for p in d.paragraphs:
						t = p.text.strip()
						if t and t not in content_lines:
							content_lines.append(t)
					for table in d.tables:
						for row in table.rows:
							row_txt = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
							if row_txt and row_txt not in content_lines:
								content_lines.append(row_txt)
					del d
				except Exception:
					pass

				f_body = get_evrak_cover_font(9, bold=False)
				y_pos = 75
				for line in content_lines[:12]:
					if y_pos > h - 45:
						break
					disp_line = line if len(line) <= 42 else line[:39] + "..."
					draw.text((16, y_pos), disp_line, fill="#334155", font=f_body)
					y_pos += 18

				while y_pos < h - 45:
					draw.line([(16, y_pos + 5), (w - 24, y_pos + 5)], fill="#f1f5f9", width=1)
					y_pos += 16

				draw.rectangle([(0, h - 26), (w, h)], fill="#f8fafc")
				draw.line([(0, h - 26), (w, h - 26)], fill="#e2e8f0", width=1)
				draw.text((12, h - 19), "📄 1. Sayfa Önizlemesi · Resmi Revir Belgesi", fill="#64748b", font=get_evrak_cover_font(9, bold=False))
				img.save(str(thumb_path), format="JPEG", quality=80)
				img.close()
				del img
			return thumb_path

		# 4. Excel Spreadsheets (.xlsx / .xls)
		elif ext in [".xlsx", ".xls"]:
			w, h = 320, 440
			img = Image.new("RGB", (w, h), "#ffffff")
			draw = ImageDraw.Draw(img)
			draw.rectangle([(0, 0), (w - 1, h - 1)], outline="#cbd5e1", width=2)
			draw.rectangle([(0, 0), (w, 32)], fill="#16a34a")
			draw.text((12, 8), "📗 MICROSOFT EXCEL TABLOSU", fill="#ffffff", font=get_evrak_cover_font(11, bold=True))

			title_text = file_path.stem if len(file_path.stem) <= 30 else file_path.stem[:27] + "..."
			draw.text((14, 42), title_text, fill="#0f172a", font=get_evrak_cover_font(13, bold=True))
			draw.line([(14, 64), (w - 14, 64)], fill="#e2e8f0", width=1)

			sheet_rows = []
			try:
				import openpyxl
				wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
				ws = wb.active
				for r in ws.iter_rows(values_only=True):
					row_vals = [str(c).strip() for c in r if c is not None and str(c).strip()]
					if row_vals:
						sheet_rows.append(row_vals[:4])
					if len(sheet_rows) >= 12:
						break
				wb.close()
				del wb
			except Exception:
				pass

			col_widths = [30, 85, 95, 85]
			col_letters = ["A", "B", "C", "D"]
			y_tbl = 75
			draw.rectangle([(14, y_tbl), (w - 14, y_tbl + 20)], fill="#f1f5f9", outline="#cbd5e1")
			f_col = get_evrak_cover_font(9, bold=True)
			x_cur = 16
			for i, cw in enumerate(col_widths):
				draw.text((x_cur + 3, y_tbl + 4), col_letters[i], fill="#475569", font=f_col)
				x_cur += cw
				if i < 3:
					draw.line([(x_cur, y_tbl), (x_cur, y_tbl + 20)], fill="#cbd5e1", width=1)

			y_tbl += 20
			f_cell = get_evrak_cover_font(9, bold=False)
			for r_idx, row in enumerate(sheet_rows):
				if y_tbl > h - 40:
					break
				row_bg = "#ffffff" if r_idx % 2 == 0 else "#f8fafc"
				draw.rectangle([(14, y_tbl), (w - 14, y_tbl + 18)], fill=row_bg, outline="#e2e8f0")
			x_cur = 16
			for c_idx, cw in enumerate(col_widths):
				val = row[c_idx] if c_idx < len(row) else ""
				val_disp = val if len(val) <= 12 else val[:10] + ".."
				draw.text((x_cur + 3, y_tbl + 3), val_disp, fill="#1e293b", font=f_cell)
				x_cur += cw
				if c_idx < 3:
					draw.line([(x_cur, y_tbl), (x_cur, y_tbl + 18)], fill="#e2e8f0", width=1)
				y_tbl += 18

			draw.rectangle([(0, h - 26), (w, h)], fill="#f8fafc")
			draw.line([(0, h - 26), (w, h - 26)], fill="#e2e8f0", width=1)
			draw.text((12, h - 19), "📊 Tablo Veri Önizlemesi · 1. Sayfa", fill="#64748b", font=get_evrak_cover_font(9, bold=False))
			img.save(str(thumb_path), format="JPEG", quality=80)
			img.close()
			del img
			return thumb_path

		# 5. Default/Generic
		else:
			w, h = 320, 440
			img = Image.new("RGB", (w, h), "#ffffff")
			draw = ImageDraw.Draw(img)
			draw.rectangle([(0, 0), (w - 1, h - 1)], outline="#cbd5e1", width=2)
			draw.rectangle([(0, 0), (w, 32)], fill="#64748b")
			draw.text((12, 8), "📄 RESMİ EVRAK DOSYASI", fill="#ffffff", font=get_evrak_cover_font(11, bold=True))
			draw.text((14, 50), file_path.name[:28], fill="#0f172a", font=get_evrak_cover_font(13, bold=True))
			img.save(str(thumb_path), format="JPEG", quality=80)
			img.close()
			del img
			return thumb_path

	except Exception:
		return None


@st.dialog("👁️ Evrak / Belge Önizleme", width="large")
def open_document_preview_dialog(file_path: Path) -> None:
	if not file_path.exists():
		st.error("Dosya bulunamadı.")
		return

	ext = file_path.suffix.lower()
	stat = file_path.stat()
	f_size_str = format_file_size(stat.st_size)
	f_date_str = datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M")

	h_c1, h_c2 = st.columns([7, 3])
	with h_c1:
		st.markdown(f"#### 📄 {file_path.name}")
		st.caption(f"📏 Boyut: **{f_size_str}** · 🕒 Tarih: **{f_date_str}** · 🏷️ Tür: **{ext.upper().strip('.')}**")
	with h_c2:
		mime_t = mimetypes.guess_type(file_path)[0] or "application/octet-stream"
		with open(file_path, "rb") as f_dl:
			dl_data = f_dl.read()
		st.download_button("📥 Bilgisayara İndir", data=dl_data, file_name=file_path.name, mime=mime_t, type="primary", use_container_width=True)

	st.divider()

	if ext == ".pdf":
		import pymupdf
		from PIL import Image
		try:
			doc = pymupdf.open(file_path)
			num_pages = len(doc)
			st.caption(f"📑 Toplam **{num_pages}** sayfa bulundu. Önizleme aşağıda gösterilmektedir:")
			max_preview = min(num_pages, 8)
			for p_num in range(max_preview):
				page = doc[p_num]
				pix = page.get_pixmap(dpi=110)
				page_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
				st.markdown(f"**Sayfa {p_num + 1} / {num_pages}**")
				st.image(page_img, use_container_width=True)
				del page, pix, page_img
			if num_pages > max_preview:
				st.info(f"ℹ️ Belgenin devamı ({num_pages - max_preview} sayfa daha var) için sağ üstteki **📥 Bilgisayara İndir** butonunu kullanabilirsiniz.")
			doc.close()
			del doc
		except Exception as e:
			st.error(f"PDF önizleme hatası: {e}")

	elif ext in [".png", ".jpg", ".jpeg", ".webp"]:
		st.image(str(file_path), use_container_width=True)

	elif ext in [".xlsx", ".xls"]:
		try:
			import openpyxl
			wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
			sheet_names = wb.sheetnames
			sel_sheet = st.selectbox("📑 Görüntülenecek Sayfa (Sheet)", sheet_names, key=f"prev_sheet_{sha256(file_path.name.encode()).hexdigest()[:8]}")
			ws = wb[sel_sheet]
			rows_data = []
			for r in ws.iter_rows(values_only=True):
				if any(c is not None for c in r):
					rows_data.append([str(c) if c is not None else "" for c in r])
				if len(rows_data) >= 80:
					break
			wb.close()
			del wb
			if rows_data:
				headers = rows_data[0]
				body = rows_data[1:] if len(rows_data) > 1 else []
				df_preview = pd.DataFrame(body, columns=[h or f"Kolon {i+1}" for i, h in enumerate(headers)])
				st.dataframe(df_preview, use_container_width=True, height=420)
			else:
				st.info("Sayfa boş.")
		except Exception as e:
			st.warning(f"Excel önizleme oluşturulamadı: {e}")

	elif ext == ".docx":
		try:
			d = docx.Document(file_path)
			st.markdown("##### 📝 Belge İçeriği Metin Önizlemesi:")
			with st.container(border=True, height=420):
				for p in d.paragraphs:
					t = p.text.strip()
					if t:
						st.markdown(f"• {html.escape(t)}")
				if d.tables:
					st.markdown("---")
					st.markdown("##### 📊 Belgedeki Tablolar:")
					for t_idx, table in enumerate(d.tables):
						t_rows = []
						for row in table.rows:
							r_vals = [c.text.strip() for c in row.cells]
							if any(r_vals):
								t_rows.append(r_vals)
						if t_rows:
							st.caption(f"Tablo #{t_idx + 1}")
							df_t = pd.DataFrame(t_rows[1:], columns=t_rows[0]) if len(t_rows) > 1 else pd.DataFrame(t_rows)
							st.dataframe(df_t, use_container_width=True)
			del d
		except Exception as e:
			st.warning(f"Word içeriği okunamadı: {e}")

	elif ext in [".txt", ".csv"]:
		try:
			txt_content = file_path.read_text(encoding="utf-8", errors="replace")[:10000]
			st.code(txt_content, language="csv" if ext == ".csv" else "text")
		except Exception as e:
			st.warning(f"Metin okunamadı: {e}")

	else:
		cov = get_or_create_document_cover(file_path)
		if cov and cov.exists():
			st.image(str(cov), use_container_width=True)


@st.dialog("📤 Yeni Evrak / Belge Yükle", width="large")
def open_upload_document_dialog() -> None:
	st.caption("Revir ile ilgili resmi belgeleri (Word, Excel, PowerPoint), PDF yönergeleri, formları veya fotoğrafları sürükleyip bırakabilir ya da dosya seçebilirsiniz:")
	uploaded_files = st.file_uploader(
		"Evrak Dosyaları Seçin",
		type=["pdf", "docx", "doc", "xlsx", "xls", "pptx", "ppt", "jpg", "jpeg", "png", "webp", "txt", "csv"],
		accept_multiple_files=True,
		key="depot_modal_file_uploader"
	)
	st.markdown("")
	col_close, col_save = st.columns([1, 1.2])
	with col_close:
		if st.button("İptal / Kapat", use_container_width=True, key="btn_cancel_depot_modal"):
			st.rerun()
	with col_save:
		if st.button("💾 Dosyaları Depoya Kaydet", type="primary", disabled=not uploaded_files, use_container_width=True, key="btn_save_depot_modal"):
			saved_cnt = 0
			for uf in uploaded_files:
				safe_filename = re.sub(r'[\\/*?:"<>|]', "_", uf.name)
				target_path = EVRAKLAR_DIR / safe_filename
				target_path.write_bytes(uf.read())
				get_or_create_document_cover(target_path, force=True)
				saved_cnt += 1
			st.success(f"✅ {saved_cnt} evrak başarıyla depoya yüklendi ve 1. sayfa kapak fotoğrafları hazırlandı!")
			st.rerun()


def render_documents_depot() -> None:
	EVRAKLAR_DIR.mkdir(parents=True, exist_ok=True)
	(EVRAKLAR_DIR / ".thumbnails").mkdir(parents=True, exist_ok=True)

	# Clean top header with upload button on top right
	head_c, btn_c = st.columns([4, 1.3])
	with head_c:
		st.markdown("## 📁 Revir Evrak & Dosya Deposu")
		st.caption("Okul reviri ile ilgili tüm resmi evrakları, Word/Excel formlarını, PDF yönergeleri, sağlık raporlarını ve fotoğrafları güvenle depolayın, düzenleyin ve indirin.")
	with btn_c:
		if st.button("📤 Yeni Evrak Yükle", type="primary", use_container_width=True, key="btn_trigger_upload_doc"):
			open_upload_document_dialog()

	st.markdown("---")
	st.markdown("#### 🗃️ Depodaki Belgeler ve Pencere Yöneticisi")

	toolbar_c1, toolbar_c2, toolbar_c3, toolbar_c4 = st.columns([1.5, 1.8, 1.4, 2.0])
	with toolbar_c1:
		view_layout = st.selectbox(
			"📐 Kart Boyutu",
			["Büyük Kartlar (3'lü Izgara)", "Orta Kartlar (4'lü Izgara)", "Küçük Kartlar (6'lı Izgara)", "Detaylı Liste"],
			index=1,
			key="depot_layout_choice"
		)
	with toolbar_c2:
		sort_mode = st.selectbox(
			"🔃 Sıralama Biçimi",
			[
				"En Yeni Yüklenenler",
				"En Eski Yüklenenler",
				"İsme Göre (A-Z)",
				"İsme Göre (Z-A)",
				"Boyuta Göre (Büyükten Küçüğe)",
				"Boyuta Göre (Küçükten Büyüğe)",
				"Dosya Türüne Göre"
			],
			index=0,
			key="depot_sort_choice"
		)
	with toolbar_c3:
		type_filter = st.selectbox(
			"🏷️ Tür Filtresi",
			["Tümü", "PDF Belgeleri", "Word Dosyaları", "Excel Dosyaları", "Sunumlar (PPTX)", "Resimler / Görseller", "Diğer"],
			index=0,
			key="depot_type_choice"
		)
	with toolbar_c4:
		search_query = st.text_input(
			"🔍 Evrak Ara...",
			placeholder="Dosya adında ara...",
			key="depot_search_box"
		)

	all_files = [f for f in EVRAKLAR_DIR.iterdir() if f.is_file() and not f.name.startswith(".")]
	files_data = []
	for f in all_files:
		ext = f.suffix.lower()
		stat = f.stat()
		files_data.append({
			"path": f,
			"name": f.name,
			"ext": ext,
			"size": stat.st_size,
			"mtime": stat.st_mtime,
			"mtime_dt": datetime.fromtimestamp(stat.st_mtime)
		})

	if search_query.strip():
		q_lower = search_query.strip().lower()
		files_data = [x for x in files_data if q_lower in x["name"].lower()]

	if type_filter == "PDF Belgeleri":
		files_data = [x for x in files_data if x["ext"] == ".pdf"]
	elif type_filter == "Word Dosyaları":
		files_data = [x for x in files_data if x["ext"] in [".docx", ".doc"]]
	elif type_filter == "Excel Dosyaları":
		files_data = [x for x in files_data if x["ext"] in [".xlsx", ".xls"]]
	elif type_filter == "Sunumlar (PPTX)":
		files_data = [x for x in files_data if x["ext"] in [".pptx", ".ppt"]]
	elif type_filter == "Resimler / Görseller":
		files_data = [x for x in files_data if x["ext"] in [".jpg", ".jpeg", ".png", ".webp"]]
	elif type_filter == "Diğer":
		files_data = [x for x in files_data if x["ext"] not in [".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".jpg", ".jpeg", ".png", ".webp"]]

	if sort_mode == "En Yeni Yüklenenler":
		files_data.sort(key=lambda x: x["mtime"], reverse=True)
	elif sort_mode == "En Eski Yüklenenler":
		files_data.sort(key=lambda x: x["mtime"])
	elif sort_mode == "İsme Göre (A-Z)":
		files_data.sort(key=lambda x: x["name"].lower())
	elif sort_mode == "İsme Göre (Z-A)":
		files_data.sort(key=lambda x: x["name"].lower(), reverse=True)
	elif sort_mode == "Boyuta Göre (Büyükten Küçüğe)":
		files_data.sort(key=lambda x: x["size"], reverse=True)
	elif sort_mode == "Boyuta Göre (Küçükten Büyüğe)":
		files_data.sort(key=lambda x: x["size"])
	elif sort_mode == "Dosya Türüne Göre":
		files_data.sort(key=lambda x: (x["ext"], x["name"].lower()))

	total_cnt = len(files_data)
	total_size_bytes = sum(x["size"] for x in files_data)
	st.caption(f"📁 **Toplam:** {total_cnt} evrak listeleniyor · **Boyut:** {format_file_size(total_size_bytes)} · **Konum:** `revir_evraklari/`")

	if not files_data:
		if search_query.strip() or type_filter != "Tümü":
			st.info("Arama veya filtre kriterlerine uygun dosya bulunamadı.")
		else:
			st.info("Evrak deposunda henüz yüklenmiş bir dosya bulunmuyor. Sağ üstteki '📤 Yeni Evrak Yükle' butonundan belge yükleyebilirsiniz.")
	else:
		PAGE_SIZE = 12
		total_pages = max(1, (total_cnt + PAGE_SIZE - 1) // PAGE_SIZE)
		if "depot_page_num" not in st.session_state:
			st.session_state["depot_page_num"] = 1
		if st.session_state["depot_page_num"] > total_pages:
			st.session_state["depot_page_num"] = total_pages
		if st.session_state["depot_page_num"] < 1:
			st.session_state["depot_page_num"] = 1

		curr_page = st.session_state["depot_page_num"]
		start_idx = (curr_page - 1) * PAGE_SIZE
		end_idx = min(start_idx + PAGE_SIZE, total_cnt)
		page_items = files_data[start_idx:end_idx]

		if total_pages > 1:
			p_col1, p_col2, p_col3 = st.columns([1.5, 3.5, 1.5])
			with p_col1:
				if st.button("◀ Önceki Sayfa", disabled=(curr_page <= 1), key="depot_pg_prev_top", use_container_width=True):
					st.session_state["depot_page_num"] = curr_page - 1
					st.rerun()
			with p_col2:
				st.markdown(
					f"<div style='text-align:center; padding-top:6px; font-weight:600; color:#475569; font-size:13.5px;'>"
					f"📄 Sayfa {curr_page} / {total_pages} &nbsp;·&nbsp; {start_idx + 1} - {end_idx} / {total_cnt} evrak gösteriliyor"
					f"</div>",
					unsafe_allow_html=True
				)
			with p_col3:
				if st.button("Sonraki Sayfa ▶", disabled=(curr_page >= total_pages), key="depot_pg_next_top", use_container_width=True):
					st.session_state["depot_page_num"] = curr_page + 1
					st.rerun()

		# DETAYLI LİSTE GÖRÜNÜMÜ (Windows Dosya Gezgini / Finder Tarzı Yatay Satırlar)
		if view_layout == "Detaylı Liste":
			for item in page_items:
				item_hash = sha256(item['name'].encode()).hexdigest()[:12]
				with st.container(border=True):
					col_icon, col_info, col_actions = st.columns([0.8, 5.8, 3.4])
					with col_icon:
						cover_path = get_or_create_document_cover(item["path"])
						if cover_path and cover_path.exists():
							st.image(str(cover_path), use_container_width=True)
						else:
							st.markdown("<div style='font-size:28px; text-align:center; padding-top:4px;'>📄</div>", unsafe_allow_html=True)
					with col_info:
						st.markdown(f"**{item['name']}**")
						ext_label = item['ext'].upper().strip('.')
						st.caption(f"🏷️ `{ext_label}` · 📏 {format_file_size(item['size'])} · 🕒 {item['mtime_dt'].strftime('%d.%m.%Y %H:%M')}")
					with col_actions:
						la1, la2, la3, la4 = st.columns([1, 1, 1, 1])
						with la1:
							if st.button("👁️", key=f"btn_prev_list_{item_hash}", help="Evrakı Aç / Önizle", use_container_width=True):
								open_document_preview_dialog(item["path"])
						with la2:
							mime_t = mimetypes.guess_type(item["path"])[0] or "application/octet-stream"
							with open(item["path"], "rb") as f_dl:
								f_bytes = f_dl.read()
							st.download_button("📥", data=f_bytes, file_name=item["name"], mime=mime_t, key=f"dl_list_{item_hash}", help="İndir", use_container_width=True)
						with la3:
							with st.popover("✏️", use_container_width=True, help="Ad Değiştir"):
								st.markdown("##### ✏️ Yeniden Adlandır")
								new_stem = st.text_input("Yeni Dosya Adı", value=item["path"].stem, key=f"ren_list_{item_hash}")
								if st.button("Kaydet", type="primary", key=f"btn_save_ren_list_{item_hash}", use_container_width=True):
									if new_stem.strip():
										safe_stem = re.sub(r'[\\/*?:"<>|]', "_", new_stem.strip())
										new_target = EVRAKLAR_DIR / f"{safe_stem}{item['ext']}"
										if new_target != item["path"]:
											if new_target.exists():
												st.error("Bu isimde bir dosya zaten var!")
											else:
												old_cov = get_document_cover_path(item["path"])
												if old_cov.exists():
													old_cov.unlink(missing_ok=True)
												item["path"].rename(new_target)
												get_or_create_document_cover(new_target, force=True)
												st.success("Dosya adı güncellendi.")
												st.rerun()
						with la4:
							with st.popover("🗑️", use_container_width=True, help="Sil"):
								st.markdown("##### ⚠️ Dosyayı Sil")
								st.warning(f"**{item['name']}** dosyası silinecek. Emin misiniz?")
								if st.button("Evet, Sil", type="secondary", key=f"btn_del_list_{item_hash}", use_container_width=True):
									try:
										cov_p = get_document_cover_path(item["path"])
										if cov_p.exists():
											cov_p.unlink(missing_ok=True)
										item["path"].unlink()
										st.success("Dosya silindi.")
										st.rerun()
									except Exception as e:
										st.error(f"Silme hatası: {e}")

		# KART IZGARA GÖRÜNÜMÜ (3'lü, 4'lü, 6'lı)
		else:
			if view_layout == "Büyük Kartlar (3'lü Izgara)":
				grid_cols_cnt = 3
			elif view_layout == "Orta Kartlar (4'lü Izgara)":
				grid_cols_cnt = 4
			else:
				grid_cols_cnt = 6

			rows = [page_items[i:i + grid_cols_cnt] for i in range(0, len(page_items), grid_cols_cnt)]
			for row in rows:
				cols = st.columns(grid_cols_cnt)
				for idx, item in enumerate(row):
					with cols[idx]:
						item_hash = sha256(item['name'].encode()).hexdigest()[:12]
						with st.container(border=True):
							# Render First Page Cover Photo!
							cover_path = get_or_create_document_cover(item["path"])
							if cover_path and cover_path.exists():
								try:
									st.image(str(cover_path), use_container_width=True)
								except Exception:
									st.markdown("📄 Evrak Önizlemesi")
							else:
								# Clean fallback card
								st.markdown(
									f"""<div style='background:#f1f5f9; border:1px solid #cbd5e1; border-radius:8px; padding:22px 10px; text-align:center;'>
										<div style='font-size:32px;'>📄</div>
										<div style='font-weight:700; color:#475569; font-size:12px; margin-top:4px;'>{item['ext'].upper().strip('.')} BELGESİ</div>
									</div>""",
									unsafe_allow_html=True
								)

							short_name = item["name"] if len(item["name"]) <= 30 else item["name"][:27] + "..."
							st.markdown(f"<div style='font-weight:700; font-size:13.5px; margin-top:8px; word-break:break-all;' title='{html.escape(item['name'])}'>{html.escape(short_name)}</div>", unsafe_allow_html=True)
							st.caption(f"📏 {format_file_size(item['size'])} · 🕒 {item['mtime_dt'].strftime('%d.%m.%Y %H:%M')}")

							act_col0, act_col1, act_col2, act_col3 = st.columns([1, 1, 1, 1])
							with act_col0:
								if st.button("👁️", key=f"btn_prev_{item_hash}", help="Evrakı Önizle / Oku", use_container_width=True):
									open_document_preview_dialog(item["path"])
							with act_col1:
								mime_t = mimetypes.guess_type(item["path"])[0] or "application/octet-stream"
								with open(item["path"], "rb") as f_dl:
									f_bytes = f_dl.read()
								st.download_button(
									"📥",
									data=f_bytes,
									file_name=item["name"],
									mime=mime_t,
									key=f"dl_btn_{item_hash}",
									help="İndir",
									use_container_width=True
								)
							with act_col2:
								with st.popover("✏️", use_container_width=True, help="Dosyayı Yeniden Adlandır"):
									st.markdown("##### ✏️ Yeniden Adlandır")
									new_stem = st.text_input("Yeni Dosya Adı", value=item["path"].stem, key=f"rename_inp_{item_hash}")
									if st.button("Kaydet", type="primary", key=f"btn_save_rename_{item_hash}", use_container_width=True):
										if new_stem.strip():
											safe_stem = re.sub(r'[\\/*?:"<>|]', "_", new_stem.strip())
											new_target = EVRAKLAR_DIR / f"{safe_stem}{item['ext']}"
											if new_target != item["path"]:
												if new_target.exists():
													st.error("Bu isimde bir dosya zaten var!")
												else:
													# Remove old cover
													old_cov = get_document_cover_path(item["path"])
													if old_cov.exists():
														old_cov.unlink(missing_ok=True)
													item["path"].rename(new_target)
													get_or_create_document_cover(new_target, force=True)
													st.success("Dosya adı güncellendi.")
													st.rerun()
							with act_col3:
								with st.popover("🗑️", use_container_width=True, help="Dosyayı Sil"):
									st.markdown("##### ⚠️ Dosyayı Sil")
									st.warning(f"**{item['name']}** dosyası kalıcı olarak silinecek. Emin misiniz?")
									if st.button("Evet, Sil", type="secondary", key=f"btn_confirm_del_{item_hash}", use_container_width=True):
										try:
											cov_p = get_document_cover_path(item["path"])
											if cov_p.exists():
												cov_p.unlink(missing_ok=True)
											item["path"].unlink()
											st.success("Dosya silindi.")
											st.rerun()
										except Exception as e:
											st.error(f"Silme hatası: {e}")

		if total_pages > 1:
			st.markdown("")
			bp_col1, bp_col2, bp_col3 = st.columns([1.5, 3.5, 1.5])
			with bp_col1:
				if st.button("◀ Önceki Sayfa", disabled=(curr_page <= 1), key="depot_pg_prev_btm", use_container_width=True):
					st.session_state["depot_page_num"] = curr_page - 1
					st.rerun()
			with bp_col2:
				st.markdown(
					f"<div style='text-align:center; padding-top:6px; font-weight:600; color:#475569; font-size:13.5px;'>"
					f"📄 Sayfa {curr_page} / {total_pages}"
					f"</div>",
					unsafe_allow_html=True
				)
			with bp_col3:
				if st.button("Sonraki Sayfa ▶", disabled=(curr_page >= total_pages), key="depot_pg_next_btm", use_container_width=True):
					st.session_state["depot_page_num"] = curr_page + 1
					st.rerun()

		gc.collect()


def get_git_info() -> dict[str, Any]:
	info = {
		"is_git": False,
		"branch": "",
		"commit": "",
		"commit_date": "",
		"commit_msg": "",
		"remote_url": "",
		"has_remote": False,
	}
	git_dir = APP_DIR / ".git"
	if not git_dir.exists():
		return info
	info["is_git"] = True
	import subprocess
	try:
		r = subprocess.run(["git", "branch", "--show-current"], cwd=APP_DIR, capture_output=True, text=True, timeout=3)
		info["branch"] = r.stdout.strip() or "main"

		r2 = subprocess.run(["git", "log", "-1", "--format=%h|%cd|%s", "--date=short"], cwd=APP_DIR, capture_output=True, text=True, timeout=3)
		if r2.stdout.strip():
			parts = r2.stdout.strip().split("|", 2)
			if len(parts) >= 3:
				info["commit"], info["commit_date"], info["commit_msg"] = parts

		r3 = subprocess.run(["git", "remote", "get-url", "origin"], cwd=APP_DIR, capture_output=True, text=True, timeout=3)
		if r3.returncode == 0 and r3.stdout.strip():
			info["remote_url"] = r3.stdout.strip()
			info["has_remote"] = True
	except Exception:
		pass
	return info


OFFICIAL_REPO_URL = "https://github.com/blue-doggy/revir_defteri.git"


def ensure_official_git_remote() -> bool:
	import subprocess
	try:
		r = subprocess.run(["git", "remote", "get-url", "origin"], cwd=APP_DIR, capture_output=True, text=True, timeout=3)
		if r.returncode != 0:
			subprocess.run(["git", "remote", "add", "origin", OFFICIAL_REPO_URL], cwd=APP_DIR, capture_output=True, timeout=3)
		elif "blue-doggy/revir_defteri" not in r.stdout.strip():
			subprocess.run(["git", "remote", "set-url", "origin", OFFICIAL_REPO_URL], cwd=APP_DIR, capture_output=True, timeout=3)
		return True
	except Exception:
		return False


def check_for_git_updates() -> dict[str, Any]:
	res = {
		"success": False,
		"has_updates": False,
		"ahead": 0,
		"behind": 0,
		"commits": [],
		"message": "",
	}
	import subprocess
	ensure_official_git_remote()

	try:
		fetch_res = subprocess.run(["git", "fetch", "origin"], cwd=APP_DIR, capture_output=True, text=True, timeout=15)
		if fetch_res.returncode != 0:
			res["message"] = "Güncelleme sunucusuna bağlanılamadı. Lütfen internet bağlantınızı kontrol edin."
			return res

		comp_res = subprocess.run(
			["git", "rev-list", "--left-right", "--count", "HEAD...origin/main"],
			cwd=APP_DIR, capture_output=True, text=True, timeout=5
		)
		if comp_res.returncode == 0:
			ahead_str, behind_str = comp_res.stdout.strip().split()
			ahead, behind = int(ahead_str), int(behind_str)
			res["ahead"] = ahead
			res["behind"] = behind
			res["success"] = True

			if behind > 0:
				res["has_updates"] = True
				log_res = subprocess.run(
					["git", "log", "HEAD..origin/main", "--oneline", "-10"],
					cwd=APP_DIR, capture_output=True, text=True, timeout=5
				)
				res["commits"] = [c.strip() for c in log_res.stdout.strip().split("\n") if c.strip()]
				res["message"] = f"{behind} yeni sistem güncellemesi bulundu."
			else:
				res["message"] = "Sisteminiz tamamen güncel! En son sürümü kullanıyorsunuz."
		else:
			res["message"] = "Sürüm denetimi sırasında bir uyumsuzluk oluştu."
	except subprocess.TimeoutExpired:
		res["message"] = "Zaman aşımı: Güncelleme sunucusuna bağlanırken yanıt gecikti."
	except Exception as e:
		res["message"] = f"Güncelleme denetleme hatası: {e}"
	return res


def perform_system_update() -> tuple[bool, str]:
	import subprocess
	import shutil

	ensure_official_git_remote()

	backup_dir = APP_DIR / "otomatik_guncelleme_yedekleri"
	backup_dir.mkdir(parents=True, exist_ok=True)
	ts = datetime.now().strftime("%Y%m%d_%H%M%S")
	zip_path = backup_dir / f"guncelleme_oncesi_yedek_{ts}.zip"
	try:
		import zipfile
		curr_db = get_current_db_path()
		curr_pdir = get_current_person_data_dir()
		with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
			if curr_db.exists():
				zf.write(curr_db, arcname="revir_defteri.db")
			if curr_pdir.exists():
				for pf in curr_pdir.glob("*.json"):
					zf.write(pf, arcname=f"kisi_verileri/{pf.name}")
	except Exception:
		pass

	try:
		pull_res = subprocess.run(["git", "pull", "--ff-only", "origin", "main"], cwd=APP_DIR, capture_output=True, text=True, timeout=30)
		if pull_res.returncode != 0:
			pull_res = subprocess.run(["git", "pull", "origin", "main"], cwd=APP_DIR, capture_output=True, text=True, timeout=30)
			if pull_res.returncode != 0:
				return False, f"Güncelleme indirilemedi: {pull_res.stderr.strip() or pull_res.stdout.strip()}"

		req_file = APP_DIR / "requirements.txt"
		if req_file.exists():
			subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "--quiet"], cwd=APP_DIR, capture_output=True, timeout=40)

		clear_all_data_caches()
		return True, "Güncelleme başarıyla tamamlandı!"
	except Exception as e:
		return False, f"Güncelleme işlemi başarısız oldu: {e}"


def create_or_refresh_desktop_launcher() -> Path:
	desktop_dir = Path.home() / "OneDrive" / "Masaüstü"
	if not desktop_dir.exists():
		desktop_dir = Path.home() / "Desktop"
	desktop_dir.mkdir(parents=True, exist_ok=True)

	pythonw_exe = Path(sys.executable).parent / "pythonw.exe"
	if not pythonw_exe.exists():
		pythonw_exe = Path(sys.executable)

	baslat_pyw = APP_DIR / "baslat.pyw"

	# Create or refresh desktop .lnk shortcut with custom icon and pythonw (NO black screen, standalone window)
	lnk_path = desktop_dir / "Revir Defteri.lnk"
	try:
		import subprocess
		ps_cmd = f"""
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('{lnk_path}')
$Shortcut.TargetPath = '{pythonw_exe}'
$Shortcut.Arguments = '`"{baslat_pyw}`"'
$Shortcut.WorkingDirectory = '{APP_DIR}'
$Shortcut.IconLocation = '{APP_ICON_ICO}, 0'
$Shortcut.Description = 'Revir Defteri - Okul Saglik Takip Sistemi'
$Shortcut.WindowStyle = 1
$Shortcut.Save()
(Get-Item '{lnk_path}').LastWriteTime = Get-Date
ie4uinit.exe -show
"""
		subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, timeout=8)
	except Exception:
		pass
	return lnk_path



def render_settings() -> None:
	st.markdown("## ⚙️ Sistem, Kurum ve Uygulama Ayarları")
	current_cfg = get_app_settings()

	# Master-Detail Dikey Ayarlar Paneli
	nav_col, content_col = st.columns([1.15, 3.4], gap="large")

	with nav_col:
		with st.container(border=True):
			st.markdown("#### ⚙️ Ayarlar Menüsü")
			st.caption("Ayar kategorisini seçin:")
			settings_sections = [
				"🏫 Okul Adı, Kurum & Rapor Anteti",
				"📅 Eğitim Öğretim Yılı & Arşiv Yönetimi",
				"📥 Liste İçe Aktar (Excel / PDF)",
				"🎨 Görünüm, Yazı Tipi & Dil",
				"📍 Varsayılan Konum & İl-İlçe-Köy",
				"🩺 Klinik Şablonlar & WHO",
				"📋 Bakım Planı Şablonları (NANDA-I)",
				"🔍 Okul Taraması Alanları (Görünürlük)",
				"💊 İlaç Takip Alanları (Görünürlük)",
				"🗂️ Şube ayarları",
				"💾 Veritabanı & Yedekleme",
				"🚀 Sürüm & Güncelleme Yönetimi",
			]
			active_sec = st.radio(
				"Ayar Menüsü",
				settings_sections,
				index=0,
				key="settings_vertical_nav",
				label_visibility="collapsed"
			)

	with content_col:
		# 📅 Eğitim Öğretim Yılı Yönetimi
		if active_sec == "📅 Eğitim Öğretim Yılı & Arşiv Yönetimi":
			render_academic_year_settings()

		# 🏫 Okul Adı, Kurum Bilgileri ve Rapor Anteti
		elif active_sec == "🏫 Okul Adı, Kurum & Rapor Anteti":
			st.markdown("### 🏫 Okul / Kurum Bilgileri, Revir Hemşiresi, Müdür ve Rapor Anteti")
			st.caption("Sol menüde, üst bilgi alanında, revir kayıtlarında, sağlık profillerinde ve yazdırılabilir resmi raporlarda kullanılan okul adı, kurum başlığı, revir hemşiresi ve okul müdürü bilgilerini buradan düzenleyebilir ve senkronize edebilirsiniz.")
			in_c1, in_c2 = st.columns(2)
			with in_c1:
				inst_school = st.text_input(
					"🏫 Okul / Kurum Adı",
					value=current_cfg.get("school_name", "Karaman Özel Eğitim Anaokulu"),
					placeholder="Örn: Karaman Özel Eğitim Anaokulu",
					key="cfg_school_name",
				)
				inst_top = st.text_input(
					"🏛️ Üst Kurum Başlığı (Antet 1. Satır)",
					value=current_cfg.get("institution_top_header", "T.C. MİLLÎ EĞİTİM BAKANLIĞI"),
					placeholder="Örn: T.C. MİLLÎ EĞİTİM BAKANLIĞI",
					key="cfg_inst_top",
				)
				inst_phone = st.text_input(
					"📞 Kurum / Revir İletişim Telefonu",
					value=current_cfg.get("institution_phone", ""),
					placeholder="Örn: 0338 000 00 00",
					key="cfg_inst_phone",
				)
			with in_c2:
				inst_nurse = st.text_input(
					"👩‍⚕️ Okul Revir Sorumlusu / Hemşire Adı ve Ünvanı",
					value=current_cfg.get("nurse_name", "Okul Revir Sorumlusu"),
					placeholder="Örn: Okul Revir Sorumlusu Muhammet DOĞANÇAY",
					key="cfg_nurse_name",
				)
				inst_princ = st.text_input(
					"👔 Okul Müdürü / İdareci Adı ve Ünvanı",
					value=current_cfg.get("principal_name", "Okul Müdürü"),
					placeholder="Örn: Okul Müdürü Murat KABAKÇI",
					key="cfg_principal_name",
				)

			ti_c1, ti_c2 = st.columns([1.5, 2])
			with ti_c1:
				if st.button("💾 Bilgileri Kaydet ve Senkronize Et", type="primary", key="btn_save_inst_settings", use_container_width=True):
					save_app_settings({
						"institution_top_header": inst_top.strip(),
						"school_name": inst_school.strip(),
						"institution_phone": inst_phone.strip(),
						"nurse_name": inst_nurse.strip(),
						"principal_name": inst_princ.strip(),
					})
					sync_res = sync_nurse_and_principal_everywhere(inst_nurse.strip(), inst_princ.strip())
					st.success(f"Kurum, hemşire ve müdür bilgileri kaydedildi; {sync_res['health_profiles_updated']} sağlık profili ve {sync_res['json_files_synced']} dosya senkronize edildi.")
					st.rerun()
			with ti_c2:
				if st.button("🔄 Tüm Sağlık Profillerini ve JSON Dosyalarını Eşitle", type="secondary", key="btn_sync_inst_to_all", use_container_width=True):
					sync_res = sync_nurse_and_principal_everywhere(inst_nurse.strip(), inst_princ.strip())
					st.success(f"Tüm sistem tarandı: {sync_res['health_profiles_updated']} sağlık profili ve {sync_res['json_files_synced']} kişi dosyası eşitlendi.")
					st.rerun()

		# 📥 Liste İçe Aktarma & Excel Şablonu
		elif active_sec == "📥 Liste İçe Aktar (Excel / PDF)":
			st.markdown("### 📥 Öğrenci, Servis / Adres Listesi İçe Aktarma ve Şablon")
			st.caption("Excel, CSV veya e-Okul PDF öğrenci listesini sisteme yükleyebilir, yeni öğrencileri kaydedebilir veya var olan öğrencilerin bilgilerini güncelleyebilirsiniz.")

			col_upload, col_template = st.columns([2, 1])
			with col_upload:
				uploaded = st.file_uploader(
					"Öğrenci veya servis/adres listesini Excel/CSV/PDF olarak yükle (e-Okul PDF listesi, Öğrenci Adı, Sınıf, Okul No, Doğum Tarihi, Köy, Mahalle/Adres, Telefon)",
					type=["xlsx", "xls", "csv", "pdf"],
					key="settings_roster_upload",
				)
				if uploaded is not None and st.button("📥 Listeyi İçe Aktar", type="primary", key="settings_import_roster"):
					try:
						imported, skipped, issues = import_roster(uploaded)
						sync_student_measurements()
						st.success(f"{imported} öğrenci eklendi veya güncellendi. {skipped} satır atlandı.")
						for issue in issues[:8]:
							st.warning(issue)
						st.rerun()
					except Exception as error:
						st.error(f"Dosya okunamadı: {error}")
			with col_template:
				st.markdown("##### 📄 Boş Excel Şablonu")
				st.caption("Öğrenci ve servis listesini standart biçimde hazırlamak için hazır Excel şablonunu indirin:")
				st.download_button(
					"📄 Excel Şablonunu İndir",
					make_roster_template(),
					"ogrenci_ve_servis_listesi_sablonu.xlsx",
					"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
					key="settings_download_roster_template",
					use_container_width=True,
				)


		# 🎨 Görünüm, Yazı Tipi, Boyut ve Dil
		elif active_sec == "🎨 Görünüm, Yazı Tipi & Dil":
			st.markdown("### 🎨 Yazı Tipi, Pencere Teması, Punto ve Tablo Ayarları")
			st.caption("Uygulamanın yazı tipini, yazı boyutunu, pencere renk temasını, tablo yüksekliğini ve arayüz dilini özelleştirin.")
			app_c1, app_c2, app_c3, app_c4 = st.columns(4)
			font_keys = list(FONT_OPTIONS.keys())
			cur_font = current_cfg.get("font_family", "DM Sans")
			f_idx = font_keys.index(cur_font) if cur_font in font_keys else 0
			with app_c1:
				sel_font = st.selectbox("🔤 Yazı Tipi (Font)", font_keys, index=f_idx, key="cfg_font_family")

			theme_keys = list(THEME_PRESETS.keys())
			cur_theme = current_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
			t_idx = theme_keys.index(cur_theme) if cur_theme in theme_keys else 0
			with app_c2:
				sel_theme = st.selectbox("🎨 Pencere Teması", theme_keys, index=t_idx, key="cfg_theme_name")

			cur_dark_mode = current_cfg.get("dark_mode_adaptation", "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)")
			dm_idx = DARK_MODE_OPTIONS.index(cur_dark_mode) if cur_dark_mode in DARK_MODE_OPTIONS else 0
			with app_c3:
				sel_dark_mode = st.selectbox(
					"🌓 Dark Mode & Sistem Uyumu",
					DARK_MODE_OPTIONS,
					index=dm_idx,
					key="cfg_dark_mode_adaptation",
					help="Sistem teması veya karanlık mod açıldığında yazıların ve kartların otomatik olarak yüksek kontrastlı ve okunabilir olmasını sağlar ya da sabit açık temada tutar."
				)

			cur_lang = current_cfg.get("language", "Türkçe (TR)")
			l_idx = LANGUAGE_OPTIONS.index(cur_lang) if cur_lang in LANGUAGE_OPTIONS else 0
			with app_c4:
				sel_lang = st.selectbox("🌐 Uygulama Dili", LANGUAGE_OPTIONS, index=l_idx, key="cfg_language")

			if sel_dark_mode == "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)":
				st.caption("🌓 **Otomatik Uyum:** Cihazınız veya tarayıcınız Karanlık Mod'a geçtiğinde arayüz otomatik olarak koyu renklere adapte olur, tüm metin ve başlıklar **yüksek kontrastlı ve net okunabilir açık renklere** çevrilir.")
			elif sel_dark_mode == "Kapalı (Dark Mode'a Adapte Olma - Sabit Aydınlık)":
				st.caption("☀️ **Sabit Aydınlık:** Cihazınız karanlık modda olsa dahi uygulama daima standart açık temada ve koyu renkli net yazılarla kalır.")
			else:
				st.caption("🌙 **Daima Karanlık:** Uygulama her zaman koyu tema ile çalışır ve yüksek kontrastlı beyaz/açık renkli yazılar kullanılır.")

			ui_c1, ui_c2, ui_c3, ui_c4 = st.columns(4)
			fsize_opts = ["Küçük (13.5px)", "Normal (15px)", "Büyük (16.5px)"]
			cur_fsize = current_cfg.get("font_size", "Normal (15px)")
			fs_idx = fsize_opts.index(cur_fsize) if cur_fsize in fsize_opts else 1
			with ui_c1:
				sel_fsize = st.selectbox("🔠 Yazı Boyutu (Punto)", fsize_opts, index=fs_idx, key="cfg_font_size")

			dens_opts = ["Kompakt", "Standart", "Geniş"]
			cur_dens = current_cfg.get("ui_density", "Standart")
			dn_idx = dens_opts.index(cur_dens) if cur_dens in dens_opts else 1
			with ui_c2:
				sel_dens = st.selectbox("📐 Ekran Yoğunluğu", dens_opts, index=dn_idx, key="cfg_ui_density")

			th_opts = ["320", "420", "550", "700"]
			cur_th = current_cfg.get("default_table_height", "420")
			th_idx = th_opts.index(cur_th) if cur_th in th_opts else 1
			with ui_c3:
				sel_th = st.selectbox("📊 Liste Tablosu Yüksekliği (px)", th_opts, index=th_idx, key="cfg_tbl_height")

			with ui_c4:
				st.write("")
				sel_show_addr = st.checkbox(
					"Özet kartta açık adresi göster",
					value=(current_cfg.get("show_address_in_summary", "1") == "1"),
					key="cfg_show_addr",
				)

			if st.button("💾 Görünüm ve Tema Ayarlarını Kaydet", type="primary", key="btn_save_ui_settings"):
				save_app_settings({
					"font_family": sel_font,
					"font_size": sel_fsize,
					"ui_density": sel_dens,
					"theme_name": sel_theme,
					"dark_mode_adaptation": sel_dark_mode,
					"language": sel_lang,
					"default_table_height": sel_th,
					"show_address_in_summary": "1" if sel_show_addr else "0",
				})
				st.success("Görünüm, yazı tipi, tema, karanlık mod ve tablo ayarları kaydedildi.")
				st.rerun()



		# 📍 Varsayılan Konum ve Türkiye İl-İlçe-Köy
		elif active_sec == "📍 Varsayılan Konum & İl-İlçe-Köy":
			st.markdown("### 📍 Kurumun Bulunduğu Ülke, İl, İlçe ve Köy (Varsayılan Konum Ayarı)")
			st.caption("Burada seçtiğiniz ülke, il, ilçe ve köy/mahalle bilgisi yeni öğrenci ve sivil kayıt pencerelerinde **varsayılan seçili** olarak gelir. Kayıt sırasında kişi farklı bir ilçeden veya köyden geliyorsa değiştirilebilir.")

			loc_c1, loc_c2, loc_c3, loc_c4 = st.columns(4)
			cur_country = current_cfg.get("default_country", "Türkiye")
			c_idx = COUNTRY_OPTIONS.index(cur_country) if cur_country in COUNTRY_OPTIONS else 0
			with loc_c1:
				sel_def_country = st.selectbox("🌍 Bulunduğu Ülke", COUNTRY_OPTIONS, index=c_idx, key="cfg_def_country")

			if sel_def_country == "Türkiye":
				provinces = get_tr_provinces()
				cur_prov = current_cfg.get("default_province", "Karaman")
				p_idx = provinces.index(cur_prov) if cur_prov in provinces else 0
				with loc_c2:
					sel_def_prov = st.selectbox("🏙️ Varsayılan İl", provinces, index=p_idx, key="cfg_def_prov")

				districts = get_tr_districts(sel_def_prov)
				cur_dist = current_cfg.get("default_district", "Merkez")
				d_idx = districts.index(cur_dist) if cur_dist in districts else 0
				with loc_c3:
					sel_def_dist = st.selectbox("🏘️ Varsayılan İlçe", districts, index=d_idx, key="cfg_def_dist")

				villages = get_tr_villages(sel_def_prov, sel_def_dist)
				vil_opts = ["İlçe Merkezi / Belirtilmedi", *villages]
				cur_vil = current_cfg.get("default_village", "")
				v_idx = vil_opts.index(cur_vil) if cur_vil in vil_opts else 0
				with loc_c4:
					sel_def_vil_raw = st.selectbox("🏡 Varsayılan Köy / Mahalle", vil_opts, index=v_idx, key="cfg_def_vil")
					sel_def_vil = "" if sel_def_vil_raw == "İlçe Merkezi / Belirtilmedi" else sel_def_vil_raw
			else:
				with loc_c2:
					sel_def_prov = st.text_input("🏙️ Varsayılan İl / Eyalet", value=current_cfg.get("default_province", ""), key="cfg_def_prov_txt")
				with loc_c3:
					sel_def_dist = st.text_input("🏘️ Varsayılan İlçe / Şehir", value=current_cfg.get("default_district", ""), key="cfg_def_dist_txt")
				with loc_c4:
					sel_def_vil = st.text_input("🏡 Varsayılan Köy / Mahalle", value=current_cfg.get("default_village", ""), key="cfg_def_vil_txt")

			if st.button("💾 Varsayılan Konum Ayarlarını Kaydet", type="primary", key="btn_save_loc_settings"):
				save_app_settings({
					"default_country": sel_def_country,
					"default_province": (sel_def_prov or "").strip(),
					"default_district": (sel_def_dist or "").strip(),
					"default_village": (sel_def_vil or "").strip(),
				})
				st.success("Varsayılan ülke, il, ilçe ve köy ayarları kaydedildi.")
				st.rerun()

			st.divider()
			st.markdown("### 🗺️ Türkiye İl, İlçe ve Köy / Mahalle Veritabanı")
			with connect_db() as connection:
				loc_stats = connection.execute(
					"""SELECT COUNT(DISTINCT province) AS prov_cnt,
							  COUNT(DISTINCT province || '|' || district) AS dist_cnt,
							  COUNT(*) AS vil_cnt
					   FROM tr_locations"""
				).fetchone()
			st_c1, st_c2, st_c3, st_c4 = st.columns([1, 1, 1, 1.2])
			st_c1.metric("Kayıtlı İl Sayısı", loc_stats["prov_cnt"] if loc_stats else 0)
			st_c2.metric("Kayıtlı İlçe Sayısı", loc_stats["dist_cnt"] if loc_stats else 0)
			st_c3.metric("Kayıtlı Köy / Mahalle", loc_stats["vil_cnt"] if loc_stats else 0)
			with st_c4:
				st.write("")
				if st.button("🔄 Türkiye İl-İlçe-Köy Veritabanını Yenile", type="secondary", key="btn_reload_tr_locations", use_container_width=True):
					cnt = import_tr_locations_to_db(force=True)
					get_tr_provinces.clear()
					get_tr_districts.clear()
					get_tr_villages.clear()
					st.success(f"Türkiye adres veritabanı güncellendi ({cnt} yerleşim yeri yüklendi).")
					st.rerun()

			with st.expander("🔍 Veritabanındaki İl, İlçe ve Köy / Mahalleleri İncele veya Yeni Köy Ekle", expanded=False):
				exp_c1, exp_c2 = st.columns(2)
				all_provs = get_tr_provinces()
				b_p_idx = all_provs.index(sel_def_prov) if sel_def_prov in all_provs else 0
				with exp_c1:
					browse_prov = st.selectbox("İncelenecek İl", all_provs, index=b_p_idx, key="browse_tr_prov")
				b_dists = get_tr_districts(browse_prov)
				b_d_idx = b_dists.index(sel_def_dist) if sel_def_dist in b_dists else 0
				with exp_c2:
					browse_dist = st.selectbox("İncelenecek İlçe", b_dists, index=b_d_idx, key="browse_tr_dist")
				b_villages = get_tr_villages(browse_prov, browse_dist)
				st.caption(f"**{browse_prov} / {browse_dist}** ilçesine bağlı toplam **{len(b_villages)}** köy/mahalle kayıtlı:")
				st.dataframe(pd.DataFrame({"İl": browse_prov, "İlçe": browse_dist, "Köy / Mahalle / Belde": b_villages}), width="stretch", hide_index=True, height=220)

				add_c1, add_c2 = st.columns([3, 1])
				with add_c1:
					custom_vil = st.text_input("Seçili ilçeye özel köy / mezra / mahalle ekle", placeholder="Örn: Yeni Mezra / Köy", key="custom_tr_vil_input")
				with add_c2:
					st.write("")
					if st.button("➕ Veritabanına Ekle", key="btn_add_custom_tr_vil", use_container_width=True):
						if custom_vil.strip():
							with connect_db() as connection:
								connection.execute(
									"INSERT OR IGNORE INTO tr_locations(province, district, village) VALUES (?, ?, ?)",
									(browse_prov, browse_dist, custom_vil.strip()),
								)
							get_tr_villages.clear()
							st.success(f"{browse_prov} / {browse_dist} -> {custom_vil.strip()} eklendi.")
							st.rerun()


		# 🩺 Klinik Hızlı Şablonlar ve WHO
		elif active_sec == "🩺 Klinik Şablonlar & WHO":
			st.markdown("### ⚡ Revir Kaydı Hızlı Şikayet ve İşlem Şablonları")
			st.caption("Revir Kaydı ve Yetişkinler sayfalarındaki çoklu seçim kutularında çıkan hazır şikayet ve müdahale seçeneklerini virgülle ayırarak özelleştirebilirsiniz.")
			qt_c1, qt_c2 = st.columns(2)
			with qt_c1:
				stu_syms = st.text_area(
					"Öğrenci Sık Şikayet Şablonları (Virgülle ayırın)",
					value=current_cfg.get("custom_student_symptoms", DEFAULT_APP_SETTINGS["custom_student_symptoms"]),
					height=100,
					key="cfg_stu_syms",
				)
				stu_acts = st.text_area(
					"Öğrenci Sık Yapılan İşlem Şablonları (Virgülle ayırın)",
					value=current_cfg.get("custom_student_actions", DEFAULT_APP_SETTINGS["custom_student_actions"]),
					height=100,
					key="cfg_stu_acts",
				)
			with qt_c2:
				civ_syms = st.text_area(
					"Sivil / Personel Sık Şikayet Şablonları (Virgülle ayırın)",
					value=current_cfg.get("custom_civilian_symptoms", DEFAULT_APP_SETTINGS["custom_civilian_symptoms"]),
					height=100,
					key="cfg_civ_syms",
				)
				civ_acts = st.text_area(
					"Sivil / Personel Sık Yapılan İşlem Şablonları (Virgülle ayırın)",
					value=current_cfg.get("custom_civilian_actions", DEFAULT_APP_SETTINGS["custom_civilian_actions"]),
					height=100,
					key="cfg_civ_acts",
				)

			btn_qt_save, btn_qt_reset = st.columns([2, 1])
			with btn_qt_save:
				if st.button("💾 Hızlı Klinik Şablonları Kaydet", type="primary", key="btn_save_clin_templates", use_container_width=True):
					save_app_settings({
						"custom_student_symptoms": stu_syms.strip(),
						"custom_student_actions": stu_acts.strip(),
						"custom_civilian_symptoms": civ_syms.strip(),
						"custom_civilian_actions": civ_acts.strip(),
					})
					st.success("Klinik hızlı şikayet ve işlem şablonları güncellendi.")
					st.rerun()
			with btn_qt_reset:
				if st.button("🔄 Varsayılan Şablonlara Dön", type="secondary", key="btn_reset_clin_templates", use_container_width=True):
					save_app_settings({
						"custom_student_symptoms": DEFAULT_APP_SETTINGS["custom_student_symptoms"],
						"custom_student_actions": DEFAULT_APP_SETTINGS["custom_student_actions"],
						"custom_civilian_symptoms": DEFAULT_APP_SETTINGS["custom_civilian_symptoms"],
						"custom_civilian_actions": DEFAULT_APP_SETTINGS["custom_civilian_actions"],
					})
					st.success("Klinik şablonlar varsayılan değerlere döndürüldü.")
					st.rerun()

			st.divider()
			st.markdown("### 📏 WHO LMS Persantil ve Büyüme Referans Tablosu")
			with connect_db() as connection:
				reference_count = connection.execute("SELECT COUNT(*) FROM lms_reference").fetchone()[0]
			st.metric("Kayıtlı WHO LMS Referans Satırı", reference_count)

			col_load, col_file = st.columns([1, 1])
			with col_load:
				st.markdown("#### 📦 Hazır WHO Büyüme Referansları")
				st.write("`persantil veriler/` klasöründeki resmi WHO genişletilmiş büyüme verilerini (Boy, Kilo, VKİ ve Baş Çevresi) tek tıkla veritabanına aktarın.")
				if st.button("WHO Referanslarını Otomatik Yükle", type="secondary", key="btn_load_who"):
					try:
						cnt_lms, cnt_wfa = import_builtin_who_references()
						st.success(f"Başarılı! {cnt_lms} adet LMS referans satırı ve {cnt_wfa} adet WFA referansı yüklendi.")
						st.rerun()
					except Exception as e:
						st.error(f"Referanslar yüklenirken hata oluştu: {e}")

			with col_file:
				st.markdown("#### 📁 Özel LMS Dosyası Yükle")
				uploaded = st.file_uploader("Resmi LMS referans dosyası", type=["xlsx", "csv"], key="lms_upload")
				if uploaded is not None and st.button("Referansı doğrula ve kaydet", type="primary"):
					try:
						count, issues = parse_lms_file(uploaded)
						st.success(f"{count} LMS satırı kaydedildi.")
						for issue in issues[:8]:
							st.warning(issue)
					except Exception as error:
						st.error(f"Referans dosyası okunamadı: {error}")


		# 📋 NANDA-I Bakım Planı Şablonları
		elif active_sec == "📋 Bakım Planı Şablonları (NANDA-I)":
			st.markdown("### 📋 NANDA-I Hemşirelik Bakım Planı Şablonları & Kısayol Yönetimi")
			st.caption(
				"Kişi Dosyası ve Sağlık Profili sayfalarında sunulan hemşirelik tanıları, gözlenen olgular, bakım hedefleri, "
				"girişim protokolleri, izlem sıklıkları ve acil sevk kriterlerini her alan için ayrı ayrı özelleştirebilir veya yeni şablonlar ekleyebilirsiniz."
			)

			sub_tabs = st.tabs([
				"🩺 1. Hemşirelik Tanıları",
				"🔍 2. Gözlenen Olgular",
				"🎯 3. Bakım Hedefleri",
				"🛠️ 4. Hemşirelik Girişimleri",
				"⏱️ 5. İzlem Sıklıkları",
				"🚨 6. Sevk Kriterleri",
				"📑 Bütünleşik Protokoller (Tüm Plan)",
			])

			section_configs = [
				("diagnoses", "Hemşirelik Tanısı (NANDA-I)", "🩺", "Örn: Doku Bütünlüğünde Bozulma Riski (NANDA: 00047)..."),
				("facts", "Gözlenen Olgu & Klinik Belirti", "🔍", "Örn: Ciltte kuruluk, çatlak, bası izi ve hassasiyet..."),
				("goals", "Bakım Hedefi (NOC)", "🎯", "Örn: 1. Doku bütünlüğünün korunması.\n2. Enfeksiyonun önlenmesi..."),
				("interventions", "Hemşirelik Girişimi (NIC)", "🛠️", "Örn: 1. Günlük cilt muayenesi yapılır.\n2. Nemlendirici losyon sürülür..."),
				("monitorings", "İzlem ve Değerlendirme Sıklığı", "⏱️", "Örn: Günlük pansuman kontrolü ve haftalık değerlendirme"),
				("escalations", "Yönlendirme / Sevk / Acil Eylem Kriteri", "🚨", "Örn: Ciltte pürülan akıntı veya ateş gelişirse derhal hekime sevk"),
			]

			for idx, (sec_key, sec_title, sec_icon, sec_placeholder) in enumerate(section_configs):
				with sub_tabs[idx]:
					st.markdown(f"#### {sec_icon} {sec_title} Şablonları")
					sec_items = get_care_plan_section_templates(sec_key)
					sec_keys = list(sec_items.keys())

					s_col1, s_col2 = st.columns([3, 1.2])
					with s_col1:
						sel_item_key = st.selectbox(
							f"Düzenlenecek / Silinecek {sec_title} Şablonunu Seçin:",
							sec_keys,
							key=f"cfg_sec_sel_{sec_key}",
						)
					if sel_item_key and sel_item_key in sec_items:
						k_hash = sha256(sel_item_key.encode()).hexdigest()[:8]
						with s_col2:
							st.write("")
							st.write("")
						if st.button(f"🗑️ Bu Şablonu Sil", key=f"btn_del_sec_{sec_key}_{k_hash}", type="secondary", use_container_width=True):
							if len(sec_items) <= 1:
								st.error("Bu bölümde en az bir şablon kayıtlı kalmalıdır!")
							else:
								del sec_items[sel_item_key]
								save_care_plan_section_templates(sec_key, sec_items)
								st.success(f"'{sel_item_key}' silindi.")
								st.rerun()

						with st.form(f"form_edit_sec_{sec_key}_{k_hash}"):
							edit_k = st.text_input("Şablon Başlığı", value=sel_item_key)
							edit_v = st.text_area("Şablon Metni (Forma Aktarılacak İçerik)", value=sec_items[sel_item_key], height=110)
							if st.form_submit_button("💾 Bu Şablondaki Değişiklikleri Kaydet", type="primary", use_container_width=True):
								if not edit_k.strip():
									st.error("Şablon başlığı boş bırakılamaz.")
								else:
									if edit_k.strip() != sel_item_key:
										del sec_items[sel_item_key]
									sec_items[edit_k.strip()] = edit_v.strip()
									save_care_plan_section_templates(sec_key, sec_items)
									st.success(f"'{edit_k.strip()}' şablonu güncellendi.")
									st.rerun()

					# Add new snippet to section
					with st.expander(f"➕ Bu Bölüme Yeni {sec_title} Şablonu Ekle", expanded=False):
						with st.form(f"form_add_new_sec_{sec_key}"):
							new_sec_k = st.text_input("Yeni Şablon Başlığı", placeholder=f"Örn: {sec_icon} Yeni {sec_title}")
							new_sec_v = st.text_area("Şablon İçeriği", placeholder=sec_placeholder, height=90)
							if st.form_submit_button(f"➕ Yeni {sec_title} Şablonunu Ekle", type="primary", use_container_width=True):
								if not new_sec_k.strip():
									st.error("Lütfen bir başlık girin.")
								elif new_sec_k.strip() in sec_items:
									st.error("Bu başlıkla bir şablon zaten mevcut!")
								else:
									sec_items[new_sec_k.strip()] = new_sec_v.strip()
									save_care_plan_section_templates(sec_key, sec_items)
									st.success(f"'{new_sec_k.strip()}' başarıyla eklendi.")
									st.rerun()

					# Table of snippets in this section
					sec_df = pd.DataFrame([{"Şablon Başlığı": k, "Şablon İçeriği": v} for k, v in sec_items.items()])
					st.dataframe(sec_df, width="stretch", hide_index=True)

					r_col1, r_col2 = st.columns([2.5, 1.5])
					with r_col2:
						if st.button(f"🔄 Bu Bölümü Varsayılanlara Sıfırla", key=f"btn_reset_sec_{sec_key}", type="secondary", use_container_width=True):
							reset_care_plan_section_templates(sec_key)
							st.success(f"{sec_title} şablonları varsayılanlara sıfırlandı.")
							st.rerun()

			# Subtab 7: Master Complete Protocols
			with sub_tabs[6]:
				st.markdown("#### 📑 Bütünleşik NANDA-I Bakım Protokolleri (Tüm Formu Tek Seferde Dolduran Şablonlar)")
				st.caption("Bu bölümdeki şablonlar, Kişi Dosyası ve Sağlık Profili sayfalarındaki '⚡ Tüm Alanları Dolduracak Hazır NANDA-I Protokolü' kutusundan seçildiğinde tüm formu tek tıkla doldurur.")

				current_templates = get_care_plan_templates()
				all_template_keys = list(current_templates.keys())
				all_categories = sorted(list({t.get("category", "Genel Klinik") for t in current_templates.values() if t.get("category")}))

				m1, m2 = st.columns(2)
				m1.metric("Kayıtlı Protokol Sayısı", len(current_templates))
				m2.metric("Kategori Sayısı", len(all_categories))

				col_sel, col_del = st.columns([3, 1.2])
				with col_sel:
					selected_key = st.selectbox(
						"Düzenlenecek / Silinecek Bütünleşik Protokolü Seçin:",
						all_template_keys,
						key="cfg_cp_selected_tpl_master",
					)

				if selected_key and selected_key in current_templates:
					tpl_data = current_templates[selected_key]
					t_hash = sha256(selected_key.encode("utf-8")).hexdigest()[:8]

					with col_del:
						st.write("")
						st.write("")
						if st.button("🗑️ Bu Protokolü Sil", key=f"cfg_cp_btn_del_{t_hash}", type="secondary", use_container_width=True):
							if len(current_templates) <= 1:
								st.error("En az bir protokol sistemde kayıtlı kalmalıdır!")
							else:
								del current_templates[selected_key]
								save_care_plan_templates(current_templates)
								st.success(f"'{selected_key}' silindi.")
								st.rerun()

					with st.form(key=f"cfg_cp_edit_form_{t_hash}"):
						e_c1, e_c2 = st.columns([2.5, 1.5])
						with e_c1:
							edit_title = st.text_input("📌 Protokol Başlığı", value=selected_key)
						with e_c2:
							cur_cat = tpl_data.get("category", "Genel Klinik")
							cat_options = sorted(list(set(all_categories + ["Genel Klinik", "Büyüme & Gelişme", "Beslenme & Metabolizma", "Solunum & Alerji", "Kronik Durumlar & Endokrin", "Kronik Durumlar & Nöroloji", "Akut Bakım & Enfeksiyon", "Akut Bakım & Travma", "Psikososyal & Yaşam Tarzı"])))
							cat_idx = cat_options.index(cur_cat) if cur_cat in cat_options else 0
							edit_cat = st.selectbox("📂 Kategori", cat_options, index=cat_idx)

						edit_diag = st.text_area("1. 🩺 Hemşirelik Tanısı (NANDA-I)", value=tpl_data.get("diagnosis", ""), height=80)
						edit_facts = st.text_area("2. 🔍 Gözlenen Olgular & Bulgular", value=tpl_data.get("observed_facts", ""), height=80)
						edit_goals = st.text_area("3. 🎯 Bakım Hedefleri", value=tpl_data.get("goals", ""), height=80)
						edit_interv = st.text_area("4. 🛠️ Planlanan Hemşirelik Girişimleri", value=tpl_data.get("interventions", ""), height=100)
						e_c3, e_c4 = st.columns(2)
						with e_c3:
							edit_mon = st.text_input("5. ⏱️ İzlem Sıklığı", value=tpl_data.get("monitoring", ""))
						with e_c4:
							edit_esc = st.text_input("6. 🚨 Sevk Kriteri", value=tpl_data.get("escalation", ""))

						status_opts = ["Aktif - İzlemde", "Taslak - Uzman İncelemesi Gerekli", "Uzman Değerlendirmesi Tamamlandı", "Tamamlandı - Hedefe Ulaşıldı", "İptal / İhtiyaç Kalmadı"]
						cur_stat = tpl_data.get("status", "Aktif - İzlemde")
						stat_idx = status_opts.index(cur_stat) if cur_stat in status_opts else 0
						edit_stat = st.selectbox("Varsayılan Durum", status_opts, index=stat_idx)

						if st.form_submit_button("💾 Protokol Değişikliklerini Kaydet", type="primary", use_container_width=True):
							new_title = edit_title.strip()
							if not new_title:
								st.error("Başlık boş bırakılamaz!")
							else:
								if new_title != selected_key and selected_key in current_templates:
									del current_templates[selected_key]
								current_templates[new_title] = {
									"category": edit_cat.strip() if edit_cat else "Genel Klinik",
									"diagnosis": edit_diag.strip(),
									"observed_facts": edit_facts.strip(),
									"goals": edit_goals.strip(),
									"interventions": edit_interv.strip(),
									"monitoring": edit_mon.strip(),
									"escalation": edit_esc.strip(),
									"status": edit_stat,
								}
								save_care_plan_templates(current_templates)
								st.success(f"'{new_title}' protokolü kaydedildi!")
								st.rerun()

				with st.expander("➕ Sisteme Yeni Bütünleşik NANDA-I Protokolü Ekle", expanded=False):
					with st.form("cfg_cp_add_master_form"):
						an_c1, an_c2 = st.columns([2.5, 1.5])
						with an_c1:
							new_title = st.text_input("Protokol Başlığı (Örn: 🩹 Diyabetik Ayak Bakımı (NANDA: 00047))")
						with an_c2:
							new_cat = st.text_input("Kategori", value="Genel Klinik")
						new_diag = st.text_area("1. 🩺 Hemşirelik Tanısı (NANDA-I)", height=75)
						new_facts = st.text_area("2. 🔍 Gözlenen Olgular", height=75)
						new_goals = st.text_area("3. 🎯 Bakım Hedefleri", height=75)
						new_interv = st.text_area("4. 🛠️ Hemşirelik Girişimleri", height=90)
						an_c3, an_c4 = st.columns(2)
						with an_c3:
							new_mon = st.text_input("5. ⏱️ İzlem Sıklığı", placeholder="Örn: Haftalık takip")
						with an_c4:
							new_esc = st.text_input("6. 🚨 Sevk Kriteri", placeholder="Örn: 112 Acil sevk")
						new_stat = st.selectbox("Varsayılan Durum", ["Aktif - İzlemde", "Taslak - Uzman İncelemesi Gerekli", "Uzman Değerlendirmesi Tamamlandı", "Tamamlandı - Hedefe Ulaşıldı", "İptal / İhtiyaç Kalmadı"])

						if st.form_submit_button("➕ Yeni Protokolü Sisteme Ekle", type="primary", use_container_width=True):
							if not new_title.strip():
								st.error("Protokol başlığı girin.")
							elif new_title.strip() in current_templates:
								st.error("Bu başlıkla bir protokol zaten var!")
							else:
								current_templates[new_title.strip()] = {
									"category": new_cat.strip() if new_cat.strip() else "Genel Klinik",
									"diagnosis": new_diag.strip(),
									"observed_facts": new_facts.strip(),
									"goals": new_goals.strip(),
									"interventions": new_interv.strip(),
									"monitoring": new_mon.strip(),
									"escalation": new_esc.strip(),
									"status": new_stat,
								}
								save_care_plan_templates(current_templates)
								st.success(f"'{new_title.strip()}' başarıyla eklendi!")
								st.rerun()

				summary_rows = [
					{
						"Protokol Adı": k, "Kategori": v.get("category", "Genel Klinik"),
						"Tanı Özeti": (v.get("diagnosis", "")[:80] + "...") if len(v.get("diagnosis", "")) > 80 else v.get("diagnosis", ""),
						"Hedef Özeti": (v.get("goals", "")[:70] + "...") if len(v.get("goals", "")) > 70 else v.get("goals", ""),
						"İzlem Sıklığı": v.get("monitoring", ""), "Sevk Kriteri": v.get("escalation", "")
					}
					for k, v in current_templates.items()
				]
				st.dataframe(pd.DataFrame(summary_rows), width="stretch", hide_index=True)

				r_col1, r_col2 = st.columns([2.5, 1.5])
				with r_col2:
					if st.button("🔄 Tüm Protokolleri Varsayılanlara Sıfırla", type="secondary", key="cfg_cp_btn_reset_all", use_container_width=True):
						reset_care_plan_templates()
						st.success("Tüm protokoller varsayılanlara sıfırlandı!")
						st.rerun()


		# 🔍 Okul Sağlık Taraması Alanları (Görünürlük)
		elif active_sec == "🔍 Okul Taraması Alanları (Görünürlük)":
			st.markdown("### 🔍 Okul Sağlık Taraması Alanları ve Görünürlük Tercihleri")
			st.caption("Okul taraması ekranında ve açılır tarama formunda hangi alanların görüneceğini buradan seçebilirsiniz. Seçtiğiniz alanlar sınıf tarama listesinde ve tarama formunda aktif olacaktır.")

			cur_vis_list = get_screening_visible_fields()

			# Quick preset buttons
			pr_c1, pr_c2, pr_c3 = st.columns([1.5, 1.5, 3])
			with pr_c1:
				if st.button("✅ Tüm Alanları Seç", key="btn_scr_preset_all", use_container_width=True):
					save_app_settings({"screening_visible_fields": ",".join(SCREENING_FIELDS_INFO.keys())})
					st.success("Tüm tarama alanları görünür yapıldı.")
					st.rerun()
			with pr_c2:
				if st.button("🎯 Temel Alanları Seç", key="btn_scr_preset_basic", use_container_width=True):
					basic_f = ["height_cm", "weight_kg", "temperature_c", "pulse_bpm", "wears_glasses", "hair_screening"]
					save_app_settings({"screening_visible_fields": ",".join(basic_f)})
					st.success("Temel alanlar seçildi.")
					st.rerun()

			st.markdown("---")
			with st.form("form_screening_fields_cfg"):
				selected_field_keys = []

				categories = {
					"Antropometri (Büyüme & Fiziksel Gelişim)": ["height_cm", "weight_kg", "head_circ_cm"],
					"Vital Bulgular (Yaşam Bulguları)": ["temperature_c", "pulse_bpm", "spo2_pct"],
					"Göz & Görme Sağlığı": ["wears_glasses", "vision_right", "vision_left"],
					"Hijyen & Parazit Taraması": ["hair_screening"],
					"İşitme Sağlığı": ["hearing_right", "hearing_left"],
					"Gelişim, Beslenme & Risk Faktörleri": ["health_behavior", "harmful_habits", "psychosocial_status"],
				}

				for cat_name, field_keys in categories.items():
					st.markdown(f"##### {cat_name}")
					f_cols = st.columns(len(field_keys))
					for c_idx, fk in enumerate(field_keys):
						f_info = SCREENING_FIELDS_INFO[fk]
						is_checked = fk in cur_vis_list
						with f_cols[c_idx]:
							chk = st.checkbox(
								f"{f_info['icon']} **{f_info['name']}**",
								value=is_checked,
								key=f"scr_cfg_chk_{fk}",
								help=f_info["desc"]
							)
							if chk:
								selected_field_keys.append(fk)

				st.write("")
				if st.form_submit_button("💾 Tarama Alanı Tercihlerini Kaydet", type="primary", use_container_width=True):
					if not selected_field_keys:
						st.error("En az bir tarama alanı seçmelisiniz.")
					else:
						save_app_settings({"screening_visible_fields": ",".join(selected_field_keys)})
						st.success("Okul taraması görünür alan tercihleri başarıyla kaydedildi.")
						st.rerun()


		# 💊 İlaç Takip Sistemi Alanları (Görünürlük)
		elif active_sec == "💊 İlaç Takip Alanları (Görünürlük)":
			st.markdown("### 💊 İlaç Takip Sistemi Alanları ve Görünürlük Tercihleri")
			st.caption("Yeni ilaç ve tedavi planı ekleme / düzenleme penceresinde hangi alanların görüneceğini buradan belirleyebilirsiniz. Kapattığınız alanlar modal pencerede gizlenerek formu daha sade ve hızlı hale getirir.")

			cur_med_vis = get_medication_visible_fields()

			mp_c1, mp_c2, mp_c3 = st.columns([1.5, 1.5, 3])
			with mp_c1:
				if st.button("🔘 Tüm Alanları Göster", key="btn_med_preset_all", use_container_width=True):
					save_app_settings({"medication_visible_fields": ",".join(MEDICATION_FIELDS_INFO.keys())})
					st.success("Tüm ilaç takip alanları görünür yapıldı.")
					st.rerun()
			with mp_c2:
				if st.button("⚡ Temel Alanları Göster", key="btn_med_preset_basic", use_container_width=True):
					basic_med_f = ["med_route", "dosage_pattern", "dose_unit", "meal_relation", "duration_days", "active_days"]
					save_app_settings({"medication_visible_fields": ",".join(basic_med_f)})
					st.success("Temel ilaç alanları seçildi.")
					st.rerun()

			st.markdown("---")
			with st.form("form_medication_fields_cfg"):
				selected_med_field_keys = []

				med_categories = {
					"Temel İlaç & Doz Bilgileri": ["med_route", "dosage_pattern", "dose_unit"],
					"Kullanım Koşulları & Güvenlik": ["meal_relation", "storage_condition", "guardian_consent", "expiry_date", "special_notes"],
					"Stok ve Takvim": ["total_volume", "duration_days", "active_days"],
				}

				for cat_name, f_keys in med_categories.items():
					st.markdown(f"##### {cat_name}")
					f_cols = st.columns(len(f_keys))
					for c_idx, fk in enumerate(f_keys):
						f_info = MEDICATION_FIELDS_INFO[fk]
						is_checked = fk in cur_med_vis
						with f_cols[c_idx]:
							chk = st.checkbox(
								f"{f_info.get('icon', '💊')} **{f_info.get('name') or f_info.get('label', fk)}**",
								value=is_checked,
								key=f"chk_med_field_{fk}",
								help=f_info["desc"]
							)
							if chk:
								selected_med_field_keys.append(fk)

				st.write("")
				if st.form_submit_button("💾 İlaç Alanı Görünürlük Ayarlarını Kaydet", type="primary", use_container_width=True):
					save_app_settings({"medication_visible_fields": ",".join(selected_med_field_keys)})
					st.success("İlaç takip alanı görünürlük tercihleri başarıyla kaydedildi.")
					st.rerun()


		# 🗂️ Okul Sınıfları ve Şubeleri Yönetimi
		elif active_sec == "🗂️ Şube ayarları":
			top_gc1, top_gc2 = st.columns([3, 1.2])
			with top_gc1:
				st.markdown("### 🏫 Şube Ayarları (Sınıflar ve Şubeler)")
				st.caption("Okuldaki mevcut sınıfları inceleyebilir, yeni sınıf/şube ekleyip çıkarabilir veya sağdaki pencere butonundan toplu sınıf ataması yapabilirsiniz. (İl, ilçe ve köy bilgileri sınıf filtresinde yer almaz.)")
			with top_gc2:
				if st.button("🪟 Sınıf Yönetimi Penceresini Aç", type="primary", key="cfg_open_cls_modal", use_container_width=True):
					open_class_manager_dialog()

			with connect_db() as connection:
				grp_df = pd.read_sql_query(
					"""SELECT c.id, c.name AS Sinif_Adi, COUNT(s.id) AS Ogrenci_Sayisi
					   FROM classes c LEFT JOIN students s ON s.class_id = c.id
					   WHERE c.name != 'Sivil / Diğer'
					   GROUP BY c.id, c.name ORDER BY c.name""",
					connection,
				)
			st.dataframe(grp_df.drop(columns=["id"], errors="ignore").rename(columns={"Sinif_Adi": "Sınıf / Şube", "Ogrenci_Sayisi": "Öğrenci Sayısı"}), width="stretch", hide_index=True)

			g_c1, g_c2 = st.columns(2)
			with g_c1:
				st.markdown("##### ➕ Yeni Sınıf / Şube Ekle")
				new_grp_name = st.text_input("Yeni Sınıf Adı (Virgülle birden fazla yazabilirsiniz)", placeholder="Örn: 1-A, 2-A, 5-B, 1. Kademe A", key="cfg_new_grp_name")
				if st.button("➕ Sınıfı Oluştur", type="primary", key="btn_create_grp"):
					parts = [p.strip() for p in new_grp_name.split(",") if p.strip()]
					if parts:
						with connect_db() as connection:
							for p in parts:
								connection.execute("INSERT OR IGNORE INTO classes(name) VALUES (?)", (p,))
						st.success(f"{len(parts)} sınıf eklendi.")
						st.rerun()
			with g_c2:
				st.markdown("##### ✏️ Sınıf Adını Değiştir veya Sınıfı Çıkar")
				editable_grps = [r["Sinif_Adi"] for _, r in grp_df.iterrows() if r["Sinif_Adi"] not in {"Sivil / Diğer", "Sınıf Belirtilmedi"}]
				if editable_grps:
					sel_old_grp = st.selectbox("İşlem Yapılacak Sınıf", editable_grps, key="cfg_old_grp")
					new_rename_grp = st.text_input("Yeni Sınıf Adı", value=sel_old_grp, key="cfg_rename_grp")
					b_ren, b_del = st.columns(2)
					with b_ren:
						if st.button("🔄 Adını Güncelle", type="secondary", key="btn_rename_grp", use_container_width=True):
							target_name = new_rename_grp.strip()
							if target_name and target_name != sel_old_grp:
								with connect_db() as connection:
									old_row = connection.execute("SELECT id FROM classes WHERE name=?", (sel_old_grp,)).fetchone()
									target_row = connection.execute("SELECT id FROM classes WHERE name=?", (target_name,)).fetchone()
									if old_row:
										if target_row:
											connection.execute("UPDATE OR IGNORE students SET class_id=? WHERE class_id=?", (target_row["id"], old_row["id"]))
											connection.execute("DELETE FROM classes WHERE id=?", (old_row["id"],))
										else:
											connection.execute("UPDATE classes SET name=? WHERE id=?", (target_name, old_row["id"]))
								st.success(f"'{sel_old_grp}' sınıfı '{target_name}' olarak güncellendi.")
								st.rerun()
					with b_del:
						if st.button("🗑️ Sınıfı Çıkar", type="secondary", key="btn_delete_grp_cfg", use_container_width=True):
							with connect_db() as connection:
								connection.execute("INSERT OR IGNORE INTO classes(name) VALUES ('Sınıf Belirtilmedi')")
								unassigned_id = connection.execute("SELECT id FROM classes WHERE name='Sınıf Belirtilmedi'").fetchone()["id"]
								old_row = connection.execute("SELECT id FROM classes WHERE name=?", (sel_old_grp,)).fetchone()
								if old_row:
									connection.execute("UPDATE OR IGNORE students SET class_id=? WHERE class_id=?", (unassigned_id, old_row["id"]))
									connection.execute("DELETE FROM classes WHERE id=?", (old_row["id"],))
							st.success(f"'{sel_old_grp}' sınıfı çıkarıldı.")
							st.rerun()

			st.divider()
			st.markdown("### 📋 Sınıf Adı Şablonları & Hazır Okul Kademeleri")
			st.caption("Anaokulu, İlkokul, Ortaokul ve Lise kademeleri için hazırlanmış standart sınıf şablonlarını tek tıkla sisteme ekleyebilir veya okulunuzdaki mevcut sınıfları tek tıkla yeni bir şablon olarak kaydedebilirsiniz.")

			all_templates = get_all_class_templates()
			custom_templates = get_custom_class_templates()

			# Get current active classes from DB
			current_classes_in_db = [
				r["Sinif_Adi"] for _, r in grp_df.iterrows()
				if r["Sinif_Adi"] not in {"Sivil / Diğer", "Sınıf Belirtilmedi"}
			]

			tpl_col1, tpl_col2 = st.columns([1.15, 1.0])

			with tpl_col1:
				st.markdown("##### ⚡ Şablon Seç ve Sisteme Uygula")
				selected_tpl_name = st.selectbox(
					"Kullanılacak Sınıf Şablonu",
					list(all_templates.keys()),
					key="cfg_sel_class_template",
				)
				tpl_classes = all_templates.get(selected_tpl_name, [])
				st.markdown(f"**Şablon İçeriği ({len(tpl_classes)} Sınıf):**")
				chips_html = " ".join([
					f"<span style='display:inline-block; background-color:#e0f2fe; color:#0369a1; padding:3px 10px; border-radius:12px; margin:2px; font-weight:600; font-size:13px;'>{c}</span>"
					for c in tpl_classes
				])
				st.markdown(f"<div style='margin-bottom:12px;'>{chips_html}</div>", unsafe_allow_html=True)

				btn_app1, btn_app2 = st.columns(2)
				with btn_app1:
					if st.button("➕ Şablonu Mevcutlara Ekle", type="primary", key="btn_apply_template_keep", use_container_width=True, help="Var olan sınıfları korur, şablondaki sınıflardan sistemde olmayanları ekler."):
						added_cnt, _ = apply_class_template(tpl_classes, clear_empty=False)
						st.success(f"'{selected_tpl_name}' şablonundan {added_cnt} yeni sınıf eklendi.")
						st.rerun()
				with btn_app2:
					if st.button("🔄 Boş Sınıfları Temizle & Kur", type="secondary", key="btn_apply_template_clean", use_container_width=True, help="İçinde öğrenci OLMAYAN boş sınıfları temizler ve bu şablonu kurar. Öğrencisi olan sınıflar kesinlikle silinmez."):
						added_cnt, del_cnt = apply_class_template(tpl_classes, clear_empty=True)
						st.success(f"{del_cnt} boş sınıf temizlendi, {added_cnt} yeni şablon sınıfı eklendi.")
						st.rerun()

				# Check if selected template is a custom user template
				raw_custom_name = selected_tpl_name.replace("⭐ ", "").replace(" (Özel Şablon)", "").strip()
				if raw_custom_name in custom_templates:
					st.write("")
					if st.button(f"🗑️ '{raw_custom_name}' Özel Şablonunu Sil", type="secondary", key=f"btn_del_cust_tpl_{raw_custom_name}"):
						delete_custom_class_template(raw_custom_name)
						st.success(f"'{raw_custom_name}' özel şablonu silindi.")
						st.rerun()

			with tpl_col2:
				st.markdown("##### 💾 Var Olan Sınıfları Şablon Olarak Kaydet")
				st.caption("Sisteminizde kayıtlı olan mevcut sınıfları tek tıkla özel şablon paketi olarak kaydedebilirsiniz.")

				st.markdown(f"**Sistemdeki Aktif Sınıflar ({len(current_classes_in_db)} adet):**")
				if current_classes_in_db:
					db_chips = " ".join([
						f"<span style='display:inline-block; background-color:#f1f5f9; color:#334155; padding:3px 8px; border-radius:10px; margin:2px; font-size:12px;'>{c}</span>"
						for c in current_classes_in_db
					])
					st.markdown(f"<div style='margin-bottom:12px; max-height:100px; overflow-y:auto; border:1px solid #e2e8f0; padding:6px; border-radius:8px;'>{db_chips}</div>", unsafe_allow_html=True)
				else:
					st.info("Henüz sisteme eklenmiş aktif sınıf bulunmuyor.")

				new_tpl_name_input = st.text_input(
					"Yeni Şablon Adı",
					placeholder="Örn: 2026-2027 Okulumuzun Şubeleri",
					key="cfg_save_current_as_tpl_name"
				)
				if st.button("💾 Mevcut Sınıfları Şablon Olarak Kaydet", type="primary", key="btn_save_current_classes_tpl", use_container_width=True):
					if not new_tpl_name_input.strip():
						st.error("Lütfen kaydedilecek şablon için bir isim yazın.")
					elif not current_classes_in_db:
						st.warning("Veritabanında kaydedilecek aktif sınıf bulunamadı.")
					else:
						save_custom_class_template(new_tpl_name_input.strip(), current_classes_in_db)
						st.success(f"'{new_tpl_name_input.strip()}' şablonu başarıyla kaydedildi! ({len(current_classes_in_db)} sınıf)")
						st.rerun()

				if custom_templates:
					with st.expander(f"📁 Kayıtlı Özel Şablonlarınız ({len(custom_templates)})"):
						for c_name, c_list in custom_templates.items():
							st.markdown(f"**{c_name}** ({len(c_list)} sınıf): `{', '.join(c_list)}`")


		# 💾 Veritabanı Yedekleme ve Bakım
		elif active_sec == "💾 Veritabanı & Yedekleme":
			st.markdown("### 💾 Veritabanı Yedekleme, Arşiv ve Sistem Bakımı")
			active_ay = get_active_academic_year()
			curr_db = get_current_db_path()
			curr_pdir = get_current_person_data_dir()
			pfile_count = len(list(curr_pdir.glob("kisi_*.json"))) if curr_pdir.exists() else 0
			evrak_count = len([f for f in EVRAKLAR_DIR.iterdir() if f.is_file() and not f.name.startswith(".")]) if EVRAKLAR_DIR.exists() else 0

			st.info(f"📌 **Aktif Çalışma Yılı:** `{active_ay['name']}` · **Veritabanı Dosyası:** `{curr_db.name}` · **Kişi Dosyaları:** `{pfile_count} adet` · **Depo Evrakları:** `{evrak_count} adet`")

			# 1. BÜYÜK TAM SİSTEM YEDEĞİ (.ZIP)
			with st.container(border=True):
				st.markdown("#### 📦 Eksiksiz Sistem Yedeği (.ZIP İndir)")
				st.caption("Aktif eğitim yılı veritabanı (`revir_defteri.db`), tüm bireysel öğrenci sağlık dosyaları (`kisi_verileri/*.json`) ve tüm evrak deposu tek bir sıkıştırılmış ZIP arşivinde toplanır.")

				import zipfile
				zip_buffer = io.BytesIO()
				with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
					if curr_db.exists():
						zf.write(curr_db, arcname="revir_defteri.db")
					if curr_pdir.exists():
						for pf in curr_pdir.glob("*.json"):
							zf.write(pf, arcname=f"kisi_verileri/{pf.name}")
					if EVRAKLAR_DIR.exists():
						for ef in EVRAKLAR_DIR.iterdir():
							if ef.is_file() and not ef.name.startswith("."):
								zf.write(ef, arcname=f"revir_evraklari/{ef.name}")
				full_zip_bytes = zip_buffer.getvalue()

				zip_fn = f"Revir_Defteri_Tam_Yedek_{active_ay['name'].replace('-', '_')}_{date.today().isoformat()}.zip"
				st.download_button(
					f"📦 Tüm Sistemi Tek Tıkla Yedekle (.ZIP İndir - {format_file_size(len(full_zip_bytes))})",
					data=full_zip_bytes,
					file_name=zip_fn,
					mime="application/zip",
					type="primary",
					use_container_width=True,
					help="Tüm veritabanını, öğrenci dosyalarını ve evrak deposunu tek arşiv olarak bilgisayarınıza indirir."
				)

			st.markdown("")
			db_c1, db_c2 = st.columns(2)
			with db_c1:
				with st.container(border=True):
					st.markdown("#### 📥 Yalnızca Veritabanı (.db)")
					if curr_db.exists():
						db_bytes = curr_db.read_bytes()
						st.download_button(
							f"💾 SQLite .db İndir ({format_file_size(len(db_bytes))})",
							data=db_bytes,
							file_name=f"revir_defteri_{active_ay['name'].replace('-', '_')}_{date.today().isoformat()}.db",
							mime="application/octet-stream",
							type="secondary",
							use_container_width=True,
						)
					st.divider()
					st.markdown("#### 🔧 Veritabanı & Dosya Senkronizasyonu")
					st.caption(f"Aktif eğitim yılına ait `{pfile_count}` kişinin JSON dosyalarını ve WHO persantillerini yeniden denetler:")
					if st.button("🔄 Kişi Dosyalarını ve Persantilleri Eşitle", use_container_width=True):
						sync_student_measurements()
						cnt_files = sync_all_person_data_files()
						clear_all_data_caches()
						with connect_db() as connection:
							connection.execute("VACUUM")
						st.success(f"{cnt_files} kişinin bireysel veri dosyası (`kisi_verileri/*.json`) ve persantil ölçümleri senkronize edildi.")
						st.rerun()

			with db_c2:
				with st.container(border=True):
					st.markdown("#### 📤 Yedekten Geri Yükle (.db / .zip)")
					restore_file = st.file_uploader("Daha önce aldığınız `.db` veya `.zip` yedek dosyasını seçin", type=["db", "sqlite", "sqlite3", "zip"], key="cfg_restore_db")
					if restore_file is not None:
						if st.button("⚠️ Seçili Yedeği Geri Yükle (Aktif Verilerin Üzerine Yazılır)", type="secondary", use_container_width=True):
							raw_data = restore_file.getvalue()
							if restore_file.name.endswith(".zip"):
								try:
									with zipfile.ZipFile(io.BytesIO(raw_data)) as r_zf:
										for z_item in r_zf.namelist():
											if z_item == "revir_defteri.db":
												curr_db.write_bytes(r_zf.read(z_item))
											elif z_item.startswith("kisi_verileri/") and z_item.endswith(".json"):
												f_name = Path(z_item).name
												(curr_pdir / f_name).write_bytes(r_zf.read(z_item))
											elif z_item.startswith("revir_evraklari/") and not Path(z_item).name.startswith("."):
												f_name = Path(z_item).name
												if f_name:
													(EVRAKLAR_DIR / f_name).write_bytes(r_zf.read(z_item))
									clear_all_data_caches()
									st.success("✅ ZIP yedek arşivi (Veritabanı, kişi dosyaları ve evraklar) başarıyla geri yüklendi!")
									st.rerun()
								except Exception as e:
									st.error(f"ZIP açma hatası: {e}")
							elif raw_data[:15] == b"SQLite format 3":
								curr_db.write_bytes(raw_data)
								clear_all_data_caches()
								st.success("✅ SQLite veritabanı yedeği başarıyla geri yüklendi!")
								st.rerun()
							else:
								st.error("Geçersiz yedek dosyası. Lütfen geçerli bir .db veya .zip dosyası yükleyin.")


		# 🚀 Sürüm & Güncelleme Yönetimi
		elif active_sec == "🚀 Sürüm & Güncelleme Yönetimi":
			st.markdown("### 🚀 Sistem Sürümü & Güncelleme Yönetimi")
			st.caption("Revir Defteri yazılımınızı doğrudan resmi güncelleme sunucusu üzerinden veya yerel yedek paketleriyle tek tıkla en son sürüme güncelleyin.")

			git_info = get_git_info()

			col_v1, col_v2, col_v3 = st.columns([1.5, 1.8, 1.4])
			with col_v1:
				with st.container(border=True):
					st.caption("📌 **Mevcut Sürüm**")
					st.markdown(f"### v{APP_VERSION}")
					st.caption(f"Aktif Dal: `{git_info['branch'] or 'main'}`")
			with col_v2:
				with st.container(border=True):
					st.caption("🔖 **Son Yapılan İyileştirme**")
					st.markdown(f"**`{git_info['commit'] or 'Yerel Sürüm'}`**")
					cmt_clean = git_info['commit_msg'] if len(git_info['commit_msg']) <= 35 else git_info['commit_msg'][:32] + "..."
					st.caption(cmt_clean or "En son commit")
			with col_v3:
				with st.container(border=True):
					st.caption("🖥️ **Masaüstü Başlatıcı**")
					if st.button("🔄 Kısayolu Yenile", key="btn_refresh_desktop_sc", use_container_width=True):
						sc_path = create_or_refresh_desktop_launcher()
						st.success(f"Masaüstü kısayolu yenilendi: `{sc_path.name}`")

			st.markdown("")

			# Online Update Checking
			with st.container(border=True):
				st.markdown("#### ⚡ Sistem Güncellemeleri")
				st.caption("Resmi sunucudaki en son kararlı sürümü ve yeni özellikleri denetler, verilerinizi otomatik yedekleyerek sistemi güvenle günceller.")

				btn_chk_c, _ = st.columns([1.6, 3])
				with btn_chk_c:
					check_clicked = st.button("🔍 Güncellemeleri Denetle", type="primary", key="btn_chk_updates", use_container_width=True)

				if check_clicked:
					with st.spinner("Güncelleme sunucusu ile iletişim kuruluyor ve denetleniyor..."): 
						update_status = check_for_git_updates()
					st.session_state["last_update_check_result"] = update_status

				chk_res = st.session_state.get("last_update_check_result")
				if chk_res:
					if not chk_res["success"]:
						st.warning(f"⚠️ {chk_res['message']}")
					elif chk_res["has_updates"]:
						st.success(f"🟢 **Yeni Güncelleme Mevcut!** ({chk_res['behind']} yeni iyileştirme bulundu)")
						st.markdown("##### 📝 Yeni Geliştirmeler ve Notlar:")
						for c in chk_res["commits"]:
							st.markdown(f"• `{c}`")

						st.markdown("")
						if st.button("⚡ Şimdi Güncelle ve Yeniden Başlat", type="primary", key="btn_apply_git_update", use_container_width=True):
							with st.spinner("Önce güvenlik yedeği alınıyor, ardından yeni kodlar yükleniyor..."):
								ok, msg = perform_system_update()
							if ok:
								st.balloons()
								st.success("✅ Güncelleme başarıyla tamamlandı! Sistem yeniden başlatılıyor...")
								st.session_state.pop("last_update_check_result", None)
								st.rerun()
							else:
								st.error(f"Güncelleme başarısız: {msg}")
					else:
						st.info(f"✅ {chk_res['message']}")

			# Offline / ZIP Manual Update
			st.markdown("")
			with st.expander("📥 Çevrimdışı / Manuel Güncelleme Paketi Yükle (.ZIP)"):
				st.caption("İnternetsiz bir ortamda veya elden aldığınız güncel `.zip` paketini yükleyerek sistemi güncelleyebilirsiniz (Tüm öğrenci verileri, vizitler ve evraklar korunur).")
				pkg_file = st.file_uploader("Güncelleme Paketini Seçin (.zip)", type=["zip"], key="cfg_upload_update_pkg")
				if pkg_file is not None:
					if st.button("🚀 Manuel Paketi Uygula", type="secondary", key="btn_apply_zip_pkg", use_container_width=True):
						try:
							import zipfile
							with zipfile.ZipFile(io.BytesIO(pkg_file.getvalue())) as zf:
								for z_name in zf.namelist():
									if z_name.endswith("/") or z_name.startswith("egitim_yillari/") or z_name.startswith("revir_evraklari/") or "revir_defteri.db" in z_name:
										continue
									zf.extract(z_name, APP_DIR)
							clear_all_data_caches()
							st.success("Manuel güncelleme paketi başarıyla uygulandı!")
							st.rerun()
						except Exception as e:
							st.error(f"Paket yükleme hatası: {e}")


def main() -> None:
	from PIL import Image as _PILImg
	fav_icon = _PILImg.open(APP_ICON_PATH) if APP_ICON_PATH.exists() else "✚"
	st.set_page_config(page_title="Revir Defteri", page_icon=fav_icon, layout="wide")
	ensure_db_initialized()
	app_cfg = get_app_settings()
	font_name = app_cfg.get("font_family", "DM Sans")
	font_css, font_url = FONT_OPTIONS.get(font_name, FONT_OPTIONS["DM Sans"])
	theme_name = app_cfg.get("theme_name", "Klinik Turkuaz (Varsayılan)")
	theme = THEME_PRESETS.get(theme_name, THEME_PRESETS["Klinik Turkuaz (Varsayılan)"])
	dark_mode_mode = app_cfg.get("dark_mode_adaptation", "Otomatik (Sistem Temasına ve Dark Mode'a Uyarla)")
	is_dark = is_dark_mode_active(app_cfg)
	muted_col = "#94a3b8" if is_dark else "#475569"

	fsize_map = {"Küçük (13.5px)": "13.5px", "Normal (15px)": "15px", "Büyük (16.5px)": "16.5px"}
	base_fsize = fsize_map.get(app_cfg.get("font_size", "Normal (15px)"), "15px")
	dens_map = {"Kompakt": "1.2rem", "Standart": "2rem", "Geniş": "2.8rem"}
	pad_top = dens_map.get(app_cfg.get("ui_density", "Standart"), "2rem")

	app_style_tag = build_app_css(font_css, font_url, base_fsize, pad_top, theme, theme_name, dark_mode_mode)
	st.markdown(app_style_tag, unsafe_allow_html=True)

	# Sol Alt Köşe - Menüyü Aç Butonu
	floating_menu_html = """
	<div id="fab-menu-toggle" onclick="(function(){
		var targets = [
			'[data-testid=\'collapsedControl\'] button',
			'[data-testid=\'collapsedControl\']',
			'[data-testid=\'stSidebarCollapsedControl\'] button',
			'[data-testid=\'stSidebarCollapsedControl\']',
			'button[data-testid=\'stSidebarToggle\']',
			'button[aria-label*=\'sidebar\' i]'
		];
		for (var i = 0; i < targets.length; i++) {
			var el = document.querySelector(targets[i]);
			if (el && el.offsetParent !== null) { el.click(); return; }
		}
		for (var i = 0; i < targets.length; i++) {
			var el = document.querySelector(targets[i]);
			if (el) { el.click(); return; }
		}
	})()" title="Menüyü Aç">
		<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
			<line x1="3" y1="12" x2="21" y2="12"></line>
			<line x1="3" y1="6" x2="21" y2="6"></line>
			<line x1="3" y1="18" x2="21" y2="18"></line>
		</svg>
		<span>Menüyü Aç</span>
	</div>
	"""
	st.markdown(floating_menu_html, unsafe_allow_html=True)

	with st.sidebar:
		if APP_ICON_PATH.exists():
			from PIL import Image as _PILImg
			col_logo, col_ttl = st.columns([1, 2.7])
			with col_logo:
				st.image(_PILImg.open(APP_ICON_PATH), use_container_width=True)
			with col_ttl:
				st.markdown("<h2 style='margin:0; padding-top:4px; font-weight:800; font-size:24px;'>Revir Defteri</h2>", unsafe_allow_html=True)
		else:
			st.markdown("# ✚ Revir Defteri")
		school_title = app_cfg.get("school_name", "")
		nurse_title = app_cfg.get("nurse_name", "")
		def_loc_str = " / ".join(filter(None, [app_cfg.get("default_province"), app_cfg.get("default_district")]))
		if school_title:
			st.caption(f"🏫 **{school_title}**" + (f" · 📍 {def_loc_str}" if def_loc_str else ""))
		else:
			st.caption(f"Öğrenci & Okul Sağlık Takip Sistemi{f' · 📍 {def_loc_str}' if def_loc_str else ''}")
		if nurse_title:
			st.caption(f"👩‍⚕️ **Revir:** {nurse_title}")

		target_p = st.session_state.pop("target_page", None)
		if target_p and target_p in PAGES:
			st.session_state["main_menu_page"] = target_p

		cur_page_val = st.session_state.get("main_menu_page", PAGES[0])
		if cur_page_val not in PAGES:
			cur_page_val = PAGES[0]
			st.session_state["main_menu_page"] = cur_page_val

		cur_idx = PAGES.index(cur_page_val)
		page = st.radio("Menü", PAGES, index=cur_idx, key="main_menu_page", label_visibility="collapsed")
		st.divider()
		active_person = get_person_by_id(st.session_state.get("active_person_id"))
		if active_person:
			with st.container(border=True):
				st.caption("📌 **Aktif Seçili Kişi**")
				st.markdown(f"**{active_person['full_name']}**")
				sno_txt = f" · No: {active_person['student_no']}" if active_person['student_no'] else ""
				st.caption(f"{active_person['class_name']}{sno_txt}")
			sb_c1, sb_c2 = st.columns([1.5, 1])
			with sb_c1:
				if st.button("📂 Kişi Dosyası", key="sidebar_goto_dossier_btn", type="primary", use_container_width=True):
					open_person_dossier_dialog(active_person["id"])
			with sb_c2:
				if st.button("Temizle", key="sidebar_clear_active_person", use_container_width=True):
					set_active_person(None)
					st.rerun()
		st.caption("🔒 Veriler bu cihazda saklanır.")
		st.caption(f"🚀 **Sürüm:** v{APP_VERSION}")

		# Sidebar Kapatma Butonu
		st.markdown("""
		<div onclick="(function(){
			var closeBtn = document.querySelector('[data-testid=\'stSidebarCollapseButton\'] button')
						|| document.querySelector('[data-testid=\'stSidebarCollapseButton\']')
						|| document.querySelector('section[data-testid=\'stSidebar\'] button[aria-label*=\'close\' i]')
						|| document.querySelector('section[data-testid=\'stSidebar\'] button[aria-label*=\'collapse\' i]');
			if (closeBtn) closeBtn.click();
		})()" style="
			display: flex;
			align-items: center;
			justify-content: center;
			gap: 6px;
			padding: 6px 10px;
			margin-top: 8px;
			border-radius: 6px;
			border: 1px solid var(--line);
			background: var(--card-bg);
			color: var(--text-muted);
			font-size: 12.5px;
			font-weight: 600;
			cursor: pointer;
			transition: all 0.2s ease;
		" onmouseover="this.style.background='var(--mint)'; this.style.color='var(--teal)';" onmouseout="this.style.background='var(--card-bg)'; this.style.color='var(--text-muted)';">
			<span>◀</span> <span>Menüyü Kapat</span>
		</div>
		""", unsafe_allow_html=True)

	# Extract clean page name regardless of emoji prefix
	page_clean = page.split(" ", 1)[-1] if " " in page else page

	if page_clean == "Genel Bakış":
		render_overview()
	elif page_clean == "Öğrenciler":
		render_students()
	elif page_clean == "Okul Taraması":
		render_school_screening()
	elif page_clean in ["Yetişkinler", "Sivil Kişiler"]:
		render_civilians()
	elif "Kişi Dosyası" in page_clean:
		render_person_dossier()
	elif page_clean == "Revir Kaydı":
		render_visit()
	elif page_clean == "İlaç Takip Sistemi":
		render_medication_tracker()
	elif page_clean == "Vital Takip":
		render_vital_tracking()
	elif page_clean == "Sağlık Profili ve Bakım Planı":
		render_health_profile_and_plan()
	elif page_clean == "Ölçüm ve Gelişim":
		render_measurements()
	elif page_clean == "Raporlar":
		render_reports()
	elif page_clean == "Evraklar":
		render_documents_depot()
	else:
		render_settings()


if __name__ == "__main__":
	main()
