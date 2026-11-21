"""
Revir Defteri - KVKK Korumalı Gemini Klinik Yapay Zekâ Asistanı (gemini_assistant.py)

Özellikler:
- Sıfır Kişisel Veri Sızıntısı: İsim, okul no, telefon, T.C. kimlik ve adres bilgileri
  buluta gönderilmeden önce yerel olarak maskelenir.
- Google Gemini REST API (Multimodal Text + PDF/Görsel Vision) istemcisi ve
  otomatik model yedekleme zinciri (gemini-3.8-flash -> gemini-flash-latest -> gemini-pro-latest).
- 7 Klinik Senaryo:
  1. Akıllı Revir Triyaj, Resmi Gözlem/Müdahale Notu & Veli Bilgilendirme Taslağı
  2. Multimodal Vision ile Laboratuvar Tahlili (PDF/Görsel) Tabloya Aktarma & Belge Özeti
  3. NANDA-I Hemşirelik Bakım Planı Otomatik Kurgulayıcı
  4. WHO Büyüme & Persantil Kanal Kayması Klinik Yorumlayıcı
  5. İlaç-Alerji, Kronik Hastalık & Pediatrik Doz Güvenlik Kalkanı
  6. Okul Salgın Radarı Trend Analizi & MEB Yıl Sonu Resmi Faaliyet Raporu Değerlendirme Yazarı
  7. Doğal Dille Revir Veritabanı Sorgulama (Salt-Okunur Güvenli SQL Asistanı)
"""
from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
import sqlite3
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

import pandas as pd


GEMINI_MODEL_OPTIONS: list[str] = [
	"gemini-3-flash-preview",
	"gemini-3.5-flash-lite",
	"gemini-3.1-flash-lite-preview",
	"gemini-flash-latest",
]

GEMINI_MODEL_LABELS: dict[str, str] = {
	"gemini-3-flash-preview": "🚀 Gemini 3 Flash (Önerilen - Hızlı Klinik, Ses & Görsel Analiz)",
	"gemini-3.5-flash-lite": "⚡ Gemini 3.5 Flash Lite (Ultra Hızlı ~1sn & Yüksek Kota)",
	"gemini-3.1-flash-lite-preview": "💨 Gemini 3.1 Flash Lite (Hafif & Hızlı Yedek)",
	"gemini-flash-latest": "🔄 Gemini Flash Latest (Otomatik Güncel Sürüm)",
}

_FALLBACK_MODELS: list[str] = [
	"gemini-3-flash-preview",
	"gemini-3.5-flash-lite",
	"gemini-3.1-flash-lite-preview",
	"gemini-flash-latest",
]

_LEGACY_MODEL_MAP: dict[str, str] = {
	"gemini-3.1-flash-lite": "gemini-3-flash-preview",
	"gemini-3.8-flash": "gemini-3-flash-preview",
	"gemini-3.1-pro-preview": "gemini-3-flash-preview",
	"gemini-2.5-flash": "gemini-3-flash-preview",
	"gemini-2.5-flash-lite": "gemini-3.5-flash-lite",
}


# ==============================================================================
# 1. ANAHTAR YÖNETİMİ VE KVKK ANONİMLEŞTİRME KATMANI
# ==============================================================================

def _local_key_file(app_dir: Path | None = None) -> Path:
	base = app_dir or Path(__file__).resolve().parent
	return base / ".gemini_api_key"


def get_api_key(app_cfg: dict[str, Any] | None = None, app_dir: Path | None = None) -> str:
	if app_cfg:
		k = str(app_cfg.get("gemini_api_key") or "").strip()
		if k:
			return k
	env_k = os.environ.get("GEMINI_API_KEY", "").strip()
	if env_k:
		return env_k
	kf = _local_key_file(app_dir)
	if kf.exists():
		try:
			return kf.read_text(encoding="utf-8").strip()
		except Exception:
			pass
	return ""


def save_local_api_key(api_key: str, app_dir: Path | None = None) -> None:
	kf = _local_key_file(app_dir)
	clean_k = (api_key or "").strip()
	try:
		if clean_k:
			kf.write_text(clean_k + "\n", encoding="utf-8")
		elif kf.exists():
			kf.unlink(missing_ok=True)
	except Exception:
		pass


def is_ai_enabled(app_cfg: dict[str, Any] | None = None, app_dir: Path | None = None) -> bool:
	if app_cfg and str(app_cfg.get("gemini_enabled", "1")).strip() == "0":
		return False
	return bool(get_api_key(app_cfg, app_dir))


_PHONE_RE = re.compile(r"(?:\+?90|0)?\s*\(?5\d{2}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}")
_TC_RE = re.compile(r"\b[1-9]\d{10}\b")


def anonymize_free_text(text: str, sensitive_names: list[str] | None = None) -> str:
	"""Serbest metin içindeki telefon, T.C. kimlik ve bilinen kişi isimlerini maskeler."""
	if not text:
		return ""
	out = str(text)
	out = _TC_RE.sub("[KVKK-TC]", out)
	out = _PHONE_RE.sub("[KVKK-TEL]", out)
	if sensitive_names:
		for raw_name in sensitive_names:
			if not raw_name:
				continue
			for token in str(raw_name).strip().split():
				if len(token) >= 3:
					out = re.sub(rf"\b{re.escape(token)}\b", "[KVKK-İSİM]", out, flags=re.IGNORECASE)
	return out


def format_disability_profile_for_ai(
	hp_dict: dict[str, Any] | None,
	person_dict: dict[str, Any] | None = None,
	sensitive_names: list[str] | None = None,
) -> str:
	"""Yapılandırılmış engel/özel gereksinim ve uzuv-duyu fonksiyonel haritasını anonim klinik metne dönüştürür."""
	hp = hp_dict or {}
	p = person_dict or {}
	has_flag = bool(int(p.get("has_disability") or 0)) if p.get("has_disability") is not None else False
	raw_json = hp.get("disability_profile_json")
	dp: dict[str, Any] = {}
	if raw_json:
		try:
			parsed = json.loads(str(raw_json))
			if isinstance(parsed, dict):
				dp = parsed
		except Exception:
			dp = {}
	if dp.get("has_disability"):
		has_flag = True
	if not has_flag:
		return ""

	bits: list[str] = []
	rate = dp.get("disability_rate")
	if rate is None:
		rate = hp.get("disability_rate")
	try:
		rate_int = int(rate or 0)
	except Exception:
		rate_int = 0
	if rate_int > 0:
		bits.append(f"Engel Oranı: %{rate_int}")
	rep_type = str(dp.get("report_type") or "").strip()
	cozger = str(dp.get("cozger_level") or "").strip()
	dep = str(dp.get("dependency_level") or "").strip()
	if rep_type and rep_type != "Belirtilmedi":
		bits.append(f"Rapor: {rep_type}" + (f" ({cozger})" if cozger and cozger != "Belirtilmedi" else ""))
	if dep and dep != "Bağımsız":
		bits.append(f"Bağımlılık Düzeyi: {dep}")

	d_types = [str(x).strip() for x in (dp.get("disability_types") or []) if str(x).strip()]
	if d_types:
		bits.append("Engel Türleri: " + ", ".join(d_types))

	mob = str(dp.get("mobility_status") or "").strip()
	if mob and mob != "Bağımsız yürüyebiliyor":
		bits.append(f"Hareket/Mobilite: {mob}")

	limb_labels = {
		"right_arm": "Sağ Kol",
		"left_arm": "Sol Kol",
		"right_hand": "Sağ El",
		"left_hand": "Sol El",
		"right_leg": "Sağ Bacak",
		"left_leg": "Sol Bacak",
	}
	limb_map = dp.get("limb_status") if isinstance(dp.get("limb_status"), dict) else {}
	limb_issues: list[str] = []
	for k, lbl in limb_labels.items():
		val = str(limb_map.get(k) or "").strip()
		if val and not val.startswith("Normal"):
			limb_issues.append(f"{lbl}: {val}")
	if limb_issues:
		bits.append("Uzuv Fonksiyon Durumu: " + "; ".join(limb_issues))

	sens_labels = {
		"right_ear": "Sağ Kulak",
		"left_ear": "Sol Kulak",
		"right_eye": "Sağ Göz",
		"left_eye": "Sol Göz",
		"speech": "Konuşma/İletişim",
	}
	sens_map = dp.get("sensory_status") if isinstance(dp.get("sensory_status"), dict) else {}
	sens_issues: list[str] = []
	for k, lbl in sens_labels.items():
		val = str(sens_map.get(k) or "").strip()
		if val and not val.startswith("Normal"):
			sens_issues.append(f"{lbl}: {val}")
	if sens_issues:
		bits.append("Duyu/İletişim Durumu: " + "; ".join(sens_issues))

	aut_notes = [str(x).strip() for x in (dp.get("autism_cognitive_notes") or []) if str(x).strip()]
	if aut_notes:
		bits.append("Otizm/Bilişsel/Duyusal Özellikler: " + ", ".join(aut_notes))

	devs = [str(x).strip() for x in (dp.get("assistive_devices") or []) if str(x).strip()]
	if devs:
		bits.append("Yardımcı Cihaz/Protez: " + ", ".join(devs))

	self_care = [str(x).strip() for x in (dp.get("self_care_needs") or []) if str(x).strip()]
	if self_care:
		bits.append("Özbakım/Destek İhtiyacı: " + ", ".join(self_care))

	prec = str(dp.get("clinical_revir_precautions") or "").strip()
	if prec:
		bits.append("Revir & Acil Özel Dikkat: " + anonymize_free_text(prec, sensitive_names))

	det = str(dp.get("detailed_notes") or "").strip()
	if det:
		bits.append("Engel Ek Detay: " + anonymize_free_text(det, sensitive_names))

	return " ; ".join(bits) if bits else "Engel / Özel Gereksinim: Var (Detay girilmedi)"


def build_anonymous_clinical_context(
	person: dict[str, Any] | sqlite3.Row | None,
	health_profile: dict[str, Any] | sqlite3.Row | None = None,
	latest_measurement: dict[str, Any] | sqlite3.Row | None = None,
	active_medications: list[Any] | None = None,
	age_months: int | None = None,
) -> str:
	"""Kişinin kimlik verilerini tamamen çıkararak yalnızca anonim klinik profil özeti üretir."""
	if not person:
		return "Kişi Profili: Belirtilmedi"
	p_dict = dict(person)
	p_type = str(p_dict.get("person_type") or "Öğrenci")
	sex_raw = str(p_dict.get("sex") or "").upper()
	sex_lbl = "Erkek" if sex_raw == "E" else ("Kız" if sex_raw == "K" else "Belirtilmedi")

	age_str = "Bilinmiyor"
	if age_months is not None and age_months >= 0:
		yrs, mos = divmod(int(age_months), 12)
		age_str = f"{yrs} yaş {mos} ay ({int(age_months)} ay)"
	elif p_dict.get("birth_date"):
		try:
			from datetime import date
			bd = date.fromisoformat(str(p_dict["birth_date"])[:10])
			today = date.today()
			mos_total = (today.year - bd.year) * 12 + (today.month - bd.month)
			yrs, mos = divmod(max(0, mos_total), 12)
			age_str = f"{yrs} yaş {mos} ay ({mos_total} ay)"
		except Exception:
			pass

	parts = [
		f"Kişi Grubu: {p_type}",
		f"Yaş: {age_str}",
		f"Cinsiyet: {sex_lbl}",
	]

	sensitive = [
		str(p_dict.get("full_name") or ""),
		str(p_dict.get("guardian_name") or ""),
	]

	hp_dict = dict(health_profile) if health_profile else {}
	if hp_dict:
		if hp_dict.get("blood_group"):
			parts.append(f"Kan Grubu: {hp_dict['blood_group']}")
		if hp_dict.get("allergies"):
			parts.append(f"Kayıtlı Alerjiler: {anonymize_free_text(str(hp_dict['allergies']), sensitive)}")
		else:
			parts.append("Kayıtlı Alerjiler: Bilinen alerji yok")
		if hp_dict.get("chronic_conditions"):
			parts.append(f"Kronik Rahatsızlıklar: {anonymize_free_text(str(hp_dict['chronic_conditions']), sensitive)}")
		if hp_dict.get("medical_history"):
			parts.append(f"Tıbbi Öykü: {anonymize_free_text(str(hp_dict['medical_history']), sensitive)}")
		reg_meds = hp_dict.get("medications") or hp_dict.get("regular_medications")
		if reg_meds:
			parts.append(f"Sürekli İlaçlar: {anonymize_free_text(str(reg_meds), sensitive)}")
		if hp_dict.get("emergency_notes"):
			parts.append(f"Acil Klinik Not: {anonymize_free_text(str(hp_dict['emergency_notes']), sensitive)}")
		sn_val = hp_dict.get("special_needs_codes") or hp_dict.get("special_needs")
		if sn_val:
			parts.append(f"Özel Gereksinim/Tanı: {anonymize_free_text(str(sn_val), sensitive)}")
		if hp_dict.get("medical_devices"):
			parts.append(f"Tıbbi Cihaz/Protez: {anonymize_free_text(str(hp_dict['medical_devices']), sensitive)}")
		if hp_dict.get("dietary_swallowing_notes"):
			parts.append(f"Beslenme/Yutma Güvenliği: {anonymize_free_text(str(hp_dict['dietary_swallowing_notes']), sensitive)}")
		if hp_dict.get("seizure_emergency_protocol"):
			parts.append(f"Acil Nöbet Protokolü: {anonymize_free_text(str(hp_dict['seizure_emergency_protocol']), sensitive)}")

	dis_ctx = format_disability_profile_for_ai(hp_dict, p_dict, sensitive)
	if dis_ctx:
		parts.append(f"ENGELLİLİK & FONKSİYONEL DURUM PROFİLİ: {dis_ctx}")

	if latest_measurement:
		lm = dict(latest_measurement)
		m_bits = []
		if lm.get("height_cm") or lm.get("Boy_cm"):
			m_bits.append(f"Boy: {lm.get('height_cm') or lm.get('Boy_cm')} cm")
		if lm.get("weight_kg") or lm.get("Kilo_kg"):
			m_bits.append(f"Kilo: {lm.get('weight_kg') or lm.get('Kilo_kg')} kg")
		if lm.get("bmi") or lm.get("BMI"):
			m_bits.append(f"VKİ: {lm.get('bmi') or lm.get('BMI')}")
		if m_bits:
			parts.append("Son Antropometri: " + ", ".join(m_bits))

	if active_medications:
		med_names = []
		for m in active_medications:
			md = dict(m) if not isinstance(m, dict) else m
			mname = md.get("medication_name")
			if mname:
				med_names.append(f"{mname} ({md.get('dosage_pattern') or ''})")
		if med_names:
			parts.append("Revirde Aktif Takip Edilen İlaçlar: " + "; ".join(med_names))

	return " | ".join(parts)


# ==============================================================================
# 2. GEMINI REST API İSTEMCİSİ (Metin + JSON + Multimodal PDF/Görsel)
# ==============================================================================

def call_gemini(
	prompt: str,
	*,
	api_key: str,
	model: str = "gemini-3-flash-preview",
	system_instruction: str | None = None,
	expect_json: bool = False,
	inline_files: list[tuple[bytes, str]] | None = None,
	temperature: float = 0.25,
	timeout: float = 16.0,
) -> tuple[bool, str, str]:
	"""
	Gemini generateContent çağrısı yapar.
	Dönüş: (başarılı_mı, yanıt_metni_veya_hata, kullanılan_model)
	"""
	clean_key = (api_key or "").strip()
	if not clean_key:
		return False, "Gemini API anahtarı tanımlanmamış. Ayarlar > Yapay Zekâ (Gemini API) menüsünden anahtarınızı kaydedin.", ""

	raw_pref = (model or "gemini-3-flash-preview").strip().replace("models/", "")
	preferred = _LEGACY_MODEL_MAP.get(raw_pref, raw_pref)
	models_to_try = [preferred] + [m for m in _FALLBACK_MODELS if m != preferred]

	parts: list[dict[str, Any]] = []
	if inline_files:
		for raw_bytes, mime_type in inline_files:
			parts.append({
				"inline_data": {
					"mime_type": mime_type or "application/pdf",
					"data": base64.b64encode(raw_bytes).decode("ascii"),
				}
			})
	parts.append({"text": prompt})

	payload: dict[str, Any] = {
		"contents": [{"parts": parts}],
		"generationConfig": {
			"temperature": temperature,
		},
	}
	if expect_json:
		payload["generationConfig"]["responseMimeType"] = "application/json"
	if system_instruction:
		payload["systemInstruction"] = {"parts": [{"text": system_instruction}]}

	last_error = ""
	for cand_model in models_to_try:
		gen_cfg: dict[str, Any] = {"temperature": temperature}
		if expect_json:
			gen_cfg["responseMimeType"] = "application/json"
		if cand_model.startswith("gemini-3"):
			gen_cfg["thinkingConfig"] = {"thinkingLevel": "LOW"}
		req_payload = dict(payload)
		req_payload["generationConfig"] = gen_cfg
		body_bytes = json.dumps(req_payload, ensure_ascii=False).encode("utf-8")

		url = f"https://generativelanguage.googleapis.com/v1beta/models/{urllib.parse.quote(cand_model)}:generateContent?key={urllib.parse.quote(clean_key)}"
		req = urllib.request.Request(
			url,
			data=body_bytes,
			headers={"Content-Type": "application/json; charset=utf-8"},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw_resp = resp.read().decode("utf-8")
				data = json.loads(raw_resp)
				candidates = data.get("candidates") or []
				if not candidates:
					last_error = "Model yanıt üretmedi (candidates boş)."
					continue
				c_parts = (candidates[0].get("content") or {}).get("parts") or []
				texts = [str(p.get("text") or "") for p in c_parts if "text" in p and not p.get("thought")]
				if not texts:
					texts = [str(p.get("text") or "") for p in c_parts if "text" in p]
				full_text = "\n".join(t for t in texts if t).strip()
				if full_text:
					return True, full_text, cand_model
				last_error = "Model boş metin döndürdü."
		except urllib.error.HTTPError as he:
			try:
				err_json = json.loads(he.read().decode("utf-8", errors="replace"))
				msg = (err_json.get("error") or {}).get("message") or str(he)
			except Exception:
				msg = str(he)
			last_error = f"HTTP {he.code}: {msg}"
			if he.code in (401, 403) and "api key" in msg.lower():
				break
			continue
		except Exception as exc:
			last_error = f"Bağlantı/Zaman aşımı ({cand_model}): {exc}"
			continue

	return False, last_error or "Gemini API yanıt vermedi.", ""


def _parse_json_response(raw_text: str) -> dict[str, Any]:
	cleaned = raw_text.strip()
	if cleaned.startswith("```"):
		cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
		cleaned = re.sub(r"\s*```$", "", cleaned)
	try:
		obj = json.loads(cleaned)
		return obj if isinstance(obj, dict) else {}
	except Exception:
		m = re.search(r"\{.*\}", cleaned, re.S)
		if m:
			try:
				obj = json.loads(m.group(0))
				return obj if isinstance(obj, dict) else {}
			except Exception:
				pass
	return {}


def _stringify_ai_field(val: Any) -> str:
	"""Gemini JSON alanlarında bazen liste veya sözlük dönen değerleri temiz Türkçe metne çevirir."""
	if val is None:
		return ""
	if isinstance(val, list):
		parts = [str(x).strip() for x in val if x is not None and str(x).strip()]
		if not parts:
			return ""
		if len(parts) == 1:
			return parts[0]
		return "\n".join(f"• {p}" if not p.startswith(("•", "-", "*")) else p for p in parts)
	if isinstance(val, dict):
		return "\n".join(f"{k}: {_stringify_ai_field(v)}" for k, v in val.items() if v not in (None, ""))
	return str(val).strip()


def test_gemini_connection(api_key: str, model: str = "gemini-3-flash-preview") -> tuple[bool, str, str]:
	ok, txt, used_model = call_gemini(
		"Merhaba! Okul revir sistemi bağlantı testi için tek cümleyle Türkçe 'Gemini Klinik Asistan bağlantısı başarılı.' yaz.",
		api_key=api_key,
		model=model,
		temperature=0.1,
		timeout=12.0,
	)
	return ok, txt, used_model


# ==============================================================================
# 3. SENARYO 1: AKILLI REVİR TRİYAJ, KLİNİK NOT & VELİ MESAJI ASİSTANI
# ==============================================================================

def analyze_visit_triage(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	complaint: str,
	vitals_summary: str,
	symptoms_draft: str,
	observation_draft: str,
	action_draft: str,
	sensitive_names: list[str] | None = None,
	school_name: str = "Okul Reviri",
	nurse_name: str = "Okul Revir Sorumlusu",
) -> tuple[bool, dict[str, Any], str]:
	sys_inst = (
		"Sen Türkiye'deki bir okul revirinde görev yapan deneyimli bir okul sağlığı hemşiresi ve pediatrik triyaj uzmanısın. "
		"Kısa veya günlük dille yazılmış revir başvuru notlarını resmi, tıbbi açıdan doğru, net ve KVKK uyumlu klinik kayıtlara dönüştürürsün. "
		"Tanı koymak yerine hemşirelik gözlemi, triyaj değerlendirmesi ve güvenli müdahale dilini kullanırsın."
	)
	safe_sym = anonymize_free_text(symptoms_draft, sensitive_names)
	safe_obs = anonymize_free_text(observation_draft, sensitive_names)
	safe_act = anonymize_free_text(action_draft, sensitive_names)

	prompt = f"""Aşağıdaki anonim öğrenci/kişi revir başvurusunu değerlendir ve JSON formatında yanıt üret:

[ANONİM PROFİL]: {anon_context}
[BAŞVURU NEDENİ]: {complaint}
[ÖLÇÜLEN VİTAL BULGULAR]: {vitals_summary or 'Vital ölçüm girilmedi'}
[BELİRTİ / ŞİKAYET TASLAĞI]: {safe_sym or 'Belirtilmedi'}
[FİZİK MUAYENE / GÖZLEM TASLAĞI]: {safe_obs or 'Belirtilmedi'}
[YAPILAN MÜDAHALE TASLAĞI]: {safe_act or 'Belirtilmedi'}

Lütfen tam olarak şu JSON anahtarlarıyla Türkçe yanıt ver:
{{
  "risk_level": "Normal | Dikkat | Acil / Kritik",
  "clinical_warning": "Yaşa göre vital bulgu sapması, kayıtlı alerji/kronik hastalık çakışması veya kırmızı bayrak varsa 1-2 cümlelik net uyarı (yoksa 'Belirgin akut kırmızı bayrak saptanmadı.')",
  "formal_symptoms": "Resmi revir defterine yazılacak profesyonel 'Belirtiler / Şikayet' metni (1-2 cümle)",
  "formal_observation": "Resmi revir defterine yazılacak profesyonel 'Fizik Muayene / Klinik Gözlem' metni (vital bulguları da içeren 1-2 cümle)",
  "formal_action": "Resmi revir defterine yazılacak profesyonel 'Yapılan İşlem / Müdahale' metni (1-2 cümle)",
  "recommended_referral": "Sevk edilmedi | Sağlık ocağı / ASM | Devlet Hastanesi | 112 Acil Sevk | Eve Gönderildi",
  "guardian_message": "Veliye gönderilecek nazik, sakinleştirici ve bilgilendirici mesaj taslağı. Öğrenci adı yerine [ÖĞRENCİ_ADI] yer tutucusunu kullan. Kurum adı: {school_name}, İmza: {nurse_name}."
}}"""

	ok, raw, used_model = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		expect_json=True,
	)
	if not ok:
		return False, {"error": raw}, used_model
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Model yanıtı JSON olarak ayrıştırılamadı."}, used_model
	return True, {k: _stringify_ai_field(v) for k, v in parsed.items()}, used_model


# ==============================================================================
# 4. SENARYO 2: MULTIMODAL VISION İLE TAHLİL & EVRAK ÇÖZÜMLEYİCİ
# ==============================================================================

def parse_lab_report_with_gemini(
	*,
	api_key: str,
	model: str,
	file_bytes: bytes,
	file_name: str,
	anon_context: str = "",
) -> tuple[bool, pd.DataFrame, str, str]:
	"""Yüklenen PDF veya görsel laboratuvar raporunu Gemini Vision ile yapılandırılmış tabloya dönüştürür."""
	mime = mimetypes.guess_type(file_name)[0] or ("application/pdf" if file_name.lower().endswith(".pdf") else "image/png")
	prompt = f"""Ekli laboratuvar / tahlil belgesindeki tüm biyokimya, hemogram, hormon veya idrar testi sonuçlarını dikkatle oku.
{f'Anonim Kişi Bağlamı: {anon_context}' if anon_context else ''}

Yanıtı yalnızca aşağıdaki JSON şemasıyla döndür:
{{
  "clinical_summary": "Referans dışı (yüksek/düşük) değerlerin ve genel tablonun 2-3 cümlelik Türkçe hemşirelik özeti",
  "results": [
    {{
      "Test": "Test adı (örn. Hemoglobin, Ferritin, TSH, Glukoz)",
      "Sonuç": 12.4,
      "Birim": "g/dL",
      "Referans alt": 11.5,
      "Referans üst": 15.5,
      "Laboratuvar işareti": "N veya H veya L"
    }}
  ]
}}"""

	ok, raw, used_model = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		inline_files=[(file_bytes, mime)],
		timeout=40.0,
	)
	cols = ["Test", "Sonuç", "Birim", "Referans alt", "Referans üst", "Laboratuvar işareti"]
	if not ok:
		return False, pd.DataFrame(columns=cols), raw, used_model

	parsed = _parse_json_response(raw)
	rows = parsed.get("results") or []
	summary = str(parsed.get("clinical_summary") or "Laboratuvar sonuçları Gemini Vision ile çözümlendi.")
	clean_rows = []
	for r in rows:
		if not isinstance(r, dict):
			continue
		t_name = str(r.get("Test") or "").strip()
		if not t_name:
			continue
		def _to_flt(v: Any) -> float | None:
			if v is None or v == "":
				return None
			try:
				return float(str(v).replace(",", ".").strip())
			except Exception:
				return None
		s_val = _to_flt(r.get("Sonuç"))
		if s_val is None:
			continue
		clean_rows.append({
			"Test": t_name,
			"Sonuç": s_val,
			"Birim": str(r.get("Birim") or "").strip(),
			"Referans alt": _to_flt(r.get("Referans alt")),
			"Referans üst": _to_flt(r.get("Referans üst")),
			"Laboratuvar işareti": str(r.get("Laboratuvar işareti") or "").strip(),
		})
	return True, pd.DataFrame(clean_rows, columns=cols), summary, used_model


def summarize_document_with_gemini(
	*,
	api_key: str,
	model: str,
	file_path: Path,
	extracted_text: str = "",
) -> tuple[bool, str, str]:
	"""Evrak deposundaki PDF, görsel veya metin belgesini okul reviri açısından özetler."""
	ext = file_path.suffix.lower()
	inline_files: list[tuple[bytes, str]] = []
	if ext in (".pdf", ".png", ".jpg", ".jpeg", ".webp") and file_path.stat().st_size <= 15 * 1024 * 1024:
		mime = mimetypes.guess_type(str(file_path))[0] or ("application/pdf" if ext == ".pdf" else "image/png")
		inline_files.append((file_path.read_bytes(), mime))

	safe_txt = anonymize_free_text(extracted_text[:10000]) if extracted_text else ""
	txt_block = f"Çıkarılan Metin Özeti:\n{safe_txt}" if safe_txt else ""
	prompt = f"""Bu belgeyi okul reviri ve öğrenci/personel sağlığı yönetimi açısından Türkçe olarak özetle.
Belge Adı: {file_path.name}
{txt_block}

Lütfen şu başlıklarla net ve maddeli bir özet sun:
1. **Belge Türü ve Amacı**
2. **Önemli Tıbbi / İdari Bulgular ve Tarihler**
3. **Okul Reviri & Hemşire İçin Aksiyon / Takip Önerisi**"""

	return call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		inline_files=inline_files if inline_files else None,
		timeout=35.0,
	)


# ==============================================================================
# 5. SENARYO 3: NANDA-I HEMŞİRELİK BAKIM PLANI OTOMATİK KURGULAYICI
# ==============================================================================

def generate_nanda_care_plan(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	compiled_facts: str,
	focus_note: str = "",
	sensitive_names: list[str] | None = None,
) -> tuple[bool, dict[str, str], str]:
	safe_facts = anonymize_free_text(compiled_facts, sensitive_names)
	safe_focus = anonymize_free_text(focus_note, sensitive_names)

	prompt = f"""Bir okul sağlığı hemşiresi olarak aşağıdaki anonim öğrenci profili için NANDA-I, NIC ve NOC standartlarına uygun, okul revirinde uygulanabilir kapsamlı bir Hemşirelik Bakım Planı oluştur:

[ANONİM PROFİL]: {anon_context}
[MEVCUT KLİNİK OLGULAR]: {safe_facts}
[EK ODAK / HEMŞİRE NOTU]: {safe_focus or 'Genel okul içi sağlık güvenliği ve kronik durum yönetimi'}

Yanıtı yalnızca aşağıdaki JSON şemasıyla Türkçe döndür:
{{
  "diagnosis": "NANDA-I Hemşirelik Tanısı (Örn: Hava Yolu Temizliğinde Etkisizlik Riski / Kan Şekeri Dengesizliği Riski / Yaralanma Riski vb.)",
  "observed_facts": "Tanıyı destekleyen klinik bulgular, kronik öykü ve okul içi risk faktörleri (2-3 cümle)",
  "goals": "Beklenen Hasta Sonuçları (NOC) - Okul saatleri içinde hedeflenen somut sağlık göstergeleri",
  "interventions": "Hemşirelik Girişimleri (NIC) - Revirde ve sınıfta uygulanacak somut adımlar (maddeli veya net paragraf)",
  "monitoring": "İzlem Sıklığı ve Değerlendirme Parametreleri (Örn: Günlük/Haftalık vital, semptom ve ilaç uyumu kontrolü)",
  "escalation": "Acil Sevk, 112 ve Veli Bildirim Kırmızı Bayrak Kriterleri (Hangi bulguda derhal acil müdahale yapılacağı)"
}}"""

	ok, raw, used_model = call_gemini(prompt, api_key=api_key, model=model, expect_json=True)
	if not ok:
		return False, {"error": raw}, used_model
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "NANDA bakım planı JSON olarak ayrıştırılamadı."}, used_model
	return True, {k: _stringify_ai_field(v) for k, v in parsed.items()}, used_model


# ==============================================================================
# 6. SENARYO 4: WHO BÜYÜME & PERSANTİL SEYRİ KLİNİK YORUMLAYICI
# ==============================================================================

def interpret_growth_trajectory(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	measurements_df: pd.DataFrame,
	school_name: str = "Okul Reviri",
	nurse_name: str = "Okul Revir Sorumlusu",
	focus_indicator: str = "",
) -> tuple[bool, dict[str, str], str]:
	cols = [
		c for c in [
			"Tarih", "Yas_ay",
			"Boy_cm", "Boy_Z", "Boy_Persantil",
			"Kilo_kg", "Kilo_Z", "Kilo_Persantil",
			"BMI", "BMI_Z", "BMI_Persantil",
			"Bas_Cevresi_cm", "Bas_Cevresi_Z", "Bas_Cevresi_Persantil",
			"Gelisim_Degerlendirmesi",
		]
		if c in measurements_df.columns
	]
	history_records: list[dict[str, Any]] = []
	if not measurements_df.empty and cols:
		for raw_rec in measurements_df[cols].head(12).to_dict(orient="records"):
			clean_rec: dict[str, Any] = {}
			for rk, rv in raw_rec.items():
				try:
					if pd.isna(rv):
						clean_rec[rk] = None
						continue
				except Exception:
					pass
				clean_rec[rk] = rv
			history_records.append(clean_rec)

	focus_line = f"\n[ÖNCELİKLİ ODAK PARAMETRESİ]: Özellikle **{focus_indicator}** persantil ve Z-skoru seyrine odaklan; diğer büyüme parametreleriyle uyumunu da kısaca belirt." if str(focus_indicator or "").strip() else ""

	prompt = f"""Dünya Sağlık Örgütü (WHO) pediatrik büyüme standartlarına göre aşağıdaki anonim öğrencinin kronolojik boy, kilo, VKİ ve (varsa) baş çevresi persantil/Z-skoru geçmişini değerlendir:{focus_line}

[ANONİM PROFİL]: {anon_context}
[KRONOLOJİK ÖLÇÜM GEÇMİŞİ (En yeniden eskiye)]:
{json.dumps(history_records, ensure_ascii=False, default=str)}

Lütfen yanıtı aşağıdaki JSON anahtarlarıyla Türkçe olarak döndür:
{{
  "channel_status": "Normal Kanalda Stabil | Persantil Kayması / İzlem Önerilir | Kritik Sınır (<3p veya >97p - Hekim Değerlendirmesi Önerilir)",
  "clinical_summary": "Hemşire için büyüme hızı, seçili parametrenin persantil/Z-skoru değişimi, boy-kilo uyumu ve varsa kanal kaymasının 2-3 cümlelik teknik analizi",
  "recommendations": "Okul revirinde izlem sıklığı, beslenme/fiziksel aktivite veya pediatri/aile hekimi yönlendirme önerisi",
  "guardian_letter": "Veliye gönderilecek resmi, anlaşılır ve yapıcı büyüme-gelişme bilgilendirme notu ([ÖĞRENCİ_ADI] yer tutucusunu kullan, Kurum: {school_name}, İmza: {nurse_name})"
}}"""

	ok, raw, used_model = call_gemini(prompt, api_key=api_key, model=model, expect_json=True)
	if not ok:
		return False, {"error": raw}, used_model
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Büyüme analizi JSON yanıtı ayrıştırılamadı."}, used_model
	return True, {k: _stringify_ai_field(v) for k, v in parsed.items()}, used_model


# ==============================================================================
# 7. SENARYO 5: İLAÇ-ALERJİ & PEDİATRİK DOZ GÜVENLİK KALKANI
# ==============================================================================

def check_medication_safety(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	medication_name: str,
	route: str,
	dosage_pattern: str,
	dose_unit: str,
	meal_relation: str,
	storage_condition: str = "",
) -> tuple[bool, dict[str, str], str]:
	prompt = f"""Okul revirinde öğrenciye uygulanmak üzere kaydedilen aşağıdaki ilacı, öğrencinin anonim sağlık profili (yaş, kilo, alerjiler, kronik hastalıklar ve mevcut ilaçlar) ile çapraz kontrol et:

[ANONİM ÖĞRENCİ PROFİLİ]: {anon_context}
[YENİ İLAÇ ADI VE FORMU]: {medication_name}
[UYGULAMA YOLU]: {route}
[DOZ ŞABLONU VE BİRİMİ]: {dosage_pattern} ({dose_unit})
[AÇ/TOK DURUMU]: {meal_relation}
[SAKLAMA KOŞULU]: {storage_condition}

Yanıtı yalnızca aşağıdaki JSON şemasıyla Türkçe döndür:
{{
  "safety_status": "Güvenli Görünüyor | Dikkat / Yakın İzlem | Kritik Uyarı (Alerji veya Etkileşim Riski!)",
  "allergy_check": "Öğrencinin kayıtlı alerjileri ile ilacın etken maddesi/grubu arasında çapraz reaksiyon riski değerlendirmesi",
  "interaction_check": "Öğrencinin kronik hastalıkları veya diğer ilaçlarıyla etkileşim ve yaş/kilo uyum değerlendirmesi",
  "administration_tips": "Okul hemşiresi için uygulama, yan etki takibi ve saklama önerisi (1-2 cümle)"
}}"""

	ok, raw, used_model = call_gemini(prompt, api_key=api_key, model=model, expect_json=True)
	if not ok:
		return False, {"error": raw}, used_model
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "İlaç güvenlik analizi ayrıştırılamadı."}, used_model
	return True, {k: _stringify_ai_field(v) for k, v in parsed.items()}, used_model


# ==============================================================================
# 8. SENARYO 6: SALGIN RADARI TREND ANALİZİ & MEB FAALİYET RAPORU YAZARI
# ==============================================================================

def analyze_outbreak_trends(
	*,
	api_key: str,
	model: str,
	anon_visit_summary: list[dict[str, Any]],
	active_isolations_count: int = 0,
) -> tuple[bool, str, str]:
	prompt = f"""Okul revirinin son 14 günlük anonim başvuru dağılımını (sınıf, tarih, başvuru nedeni ve belirti kümeleri) epidemiyolojik açıdan analiz et:

[AKTİF İZOLASYONDAKİ ÖĞRENCİ SAYISI]: {active_isolations_count}
[SON 14 GÜNLÜK ANONİM BAŞVURU ÖZETİ]:
{json.dumps(anon_visit_summary[:80], ensure_ascii=False, default=str)}

Lütfen Türkçe olarak şu 3 kısa başlık altında raporla:
1. **📊 Kümeleşme ve Salgın Riski Değerlendirmesi** (Belirli sınıflarda ateş, solunum, gastrointestinal veya bulaşıcı hastalık kümelenmesi var mı?)
2. **🛡️ Okul İdaresi ve Sınıflar İçin Koruyucu Önlemler** (Havalandırma, dezenfeksiyon, el hijyeni, veli bilgilendirmesi)
3. **👩‍⚕️ Revir İzlem Önceliği**"""

	return call_gemini(prompt, api_key=api_key, model=model)


def generate_meb_activity_narrative(
	*,
	api_key: str,
	model: str,
	stats: dict[str, Any],
	start_date_str: str,
	end_date_str: str,
	school_name: str,
	missing_data_notes: list[str] | None = None,
	custom_instruction: str = "",
) -> tuple[bool, str, str]:
	safe_stats = {
		k: v for k, v in stats.items()
		if k in (
			"students_total", "adults_total", "civilians_total", "classes_total", "visits_total",
			"student_visits", "civilian_visits", "guardian_contacted_cnt", "guardian_notified_cnt", "referral_cnt",
			"screened_students", "screened_cohort_cnt", "screening_coverage_pct", "meas_total", "bmi_eval_total",
			"bmi_normal_cnt", "bmi_normal", "bmi_under_cnt", "bmi_underweight", "bmi_over_cnt", "bmi_overweight",
			"bmi_obese_cnt", "bmi_obese", "stunted_cnt", "height_stunted", "chronic_or_allergy_cnt",
			"chronic_cnt", "allergy_cnt", "special_needs_cnt", "disability_cnt", "disability_summary_counts",
			"care_plans_total", "care_plans_period", "care_plans_active",
			"active_meds_cnt", "active_med_plans", "doses_given_cnt", "isolation_total", "isolation_period_cnt",
			"seizure_period_cnt", "vax_students_cnt", "vax_done_cnt", "vax_pending_cnt",
			"inventory_critical_cnt", "fridge_logs_cnt", "fridge_alerts_cnt", "top_complaints", "complaint_distribution",
		)
	}
	miss_block = ""
	if missing_data_notes:
		miss_block = "\n[OTOMATİK TESPİT EDİLEN EKSİK / TAMAMLANMASI GEREKEN VERİLER]:\n- " + "\n- ".join(missing_data_notes)
	inst_block = f"\n[KULLANICININ ÖZEL TALİMATI / ODAK]: {custom_instruction.strip()}" if custom_instruction and custom_instruction.strip() else ""

	prompt = f"""T.C. Millî Eğitim Bakanlığı (MEB) standartlarına uygun olarak '{school_name}' revirinin {start_date_str} - {end_date_str} dönemi resmi faaliyet raporu için 'Dönem İçi Yapılan Sağlık Hizmetleri, Eksik Veri Analizi ve Genel Değerlendirme Metni' hazırla.

[DÖNEM İSTATİSTİKLERİ]:
{json.dumps(safe_stats, ensure_ascii=False, default=str)}{miss_block}{inst_block}

Kurumsal, resmi ve akıcı bir Türkçe ile (markdown kod bloğu kullanmadan, doğrudan resmi rapora basılacak temiz başlık ve paragraflar halinde):
1. Poliklinik/revir başvuru, triyaj ve sevk hizmetlerini,
2. Periyodik okul taramaları ve DSÖ (WHO) büyüme-gelişme/persantil izlemlerini,
3. Kronik hastalığı ve engellilik/özel gereksinimi olan öğrencilerin bakım, uyarlama ve ilaç takip süreçlerini,
4. Dönemde girilmemiş / eksik kalan kayıt ve tarama alanlarına yönelik tespitleri ve bir sonraki dönem için öncelikli eylem önerilerini
kapsayıcı ve net biçimde yaz."""

	return call_gemini(prompt, api_key=api_key, model=model, timeout=22.0)


def generate_student_report_ai_analysis(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	clinical_snapshot: dict[str, Any],
	missing_data_items: list[str],
	has_disability: bool = False,
	custom_instruction: str = "",
) -> tuple[bool, str, str]:
	"""
	Öğrenci sağlık raporu için KVKK uyumlu Gemini Klinik Değerlendirme, Eksik Veri Tespiti ve
	(varsa) Engellilik/Özel Gereksinim Uyum Raporu metni üretir.
	"""
	sys_inst = (
		"Sen deneyimli bir okul sağlığı hemşiresi ve pediatrik sağlık raporlama uzmanısın. "
		"Anonim öğrenci verilerinden yola çıkarak resmi okul sağlık raporuna eklenecek net, nesnel, "
		"tıbbi açıdan doğru, gereksiz süslemeden uzak, doğrudan yazdırılabilir Türkçe değerlendirme bölümleri hazırlarsın. "
		"Öğrencinin gerçek adını uydurma; gerektiğinde 'Öğrenci' ifadesini kullan. Tanı koyma, hemşirelik değerlendirmesi ve izlem/uyarlama önerisi sun."
	)
	miss_txt = "\n- ".join(missing_data_items) if missing_data_items else "Temel profil ve izlem verileri büyük ölçüde girilmiş görünüyor."
	dis_rule = (
		"2. **Engellilik / Özel Gereksinim & Okul İçi Uyarlama Değerlendirmesi:** Öğrencinin profilinde tanımlı engel oranı, engel türleri, hareket/uzuv/duyu fonksiyon durumu ve revir dikkat notlarına göre sınıf, beden eğitimi, acil tahliye, iletişim ve revir muayene uyarlamalarını somut olarak yaz."
		if has_disability
		else "(Öğrencide engellilik/özel gereksinim kaydı bulunmadığından engellilik başlığı AÇMA.)"
	)
	extra_inst = f"\n[HEMŞİRENİN ÖZEL TALİMATI]: {custom_instruction.strip()}" if custom_instruction and custom_instruction.strip() else ""

	prompt = f"""Aşağıdaki anonim öğrenci klinik verilerini inceleyerek resmi öğrenci sağlık raporuna eklenecek değerlendirme metnini hazırla.

[ANONİM ÖĞRENCİ KLİNİK PROFİLİ]:
{anon_context}

[ÖZET KAYIT İSTATİSTİKLERİ VE SON BULGULAR]:
{json.dumps(clinical_snapshot, ensure_ascii=False, default=str)}

[SİSTEMDE GİRİLMEMİŞ / EKSİK TESPİT EDİLEN VERİLER]:
- {miss_txt}{extra_inst}

Lütfen doğrudan Word ve PDF raporuna basılabilecek temiz, profesyonel bir Türkçe ile (kod bloğu veya selamlama olmadan) şu bölümleri oluştur:
1. **Genel Klinik ve Gelişim Değerlendirmesi:** Büyüme/persantil, vital bulgular, tahliller, revir başvuru sıklığı ve kronik/alerji durumunun kısa ve bütüncül özeti.
{dis_rule}
3. **Girilmemiş / Eksik Veriler ve Tamamlanması Gereken Kayıtlar:** Yukarıda saptanan eksik verilere dayanarak hangi bilgilerin/taramaların tamamlanması gerektiğini maddeler halinde belirt.
4. **Ekstra Değinilmesi Gereken Noktalar ve İzlem Önerileri:** Verilerdeki örüntülere göre okul hemşiresi, öğretmen ve veli için öne çıkan ek dikkat noktaları ve takip önerileri."""

	return call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		temperature=0.25,
		timeout=24.0,
	)


def generate_screening_batch_ai_analysis(
	*,
	api_key: str,
	model: str,
	cohort_summary: dict[str, Any],
	missing_audit: list[str],
	custom_instruction: str = "",
) -> tuple[bool, str, str]:
	"""Toplu tarama ve persantil raporu için anonim epidemiyolojik analiz, eksik veri denetimi ve ekstra öneriler üretir."""
	sys_inst = (
		"Sen okul sağlığı epidemiyolojisi ve toplu tarama raporlaması konusunda uzman bir okul sağlığı hemşiresisin. "
		"Anonim sınıf ve tarama istatistiklerinden yola çıkarak resmi okul tarama raporuna eklenecek net, kurumsal Türkçe değerlendirme metni yazarsın."
	)
	miss_txt = "\n- ".join(missing_audit) if missing_audit else "Seçili kapsamdaki temel tarama sütunları büyük ölçüde doldurulmuş."
	extra_inst = f"\n[KULLANICI ÖZEL TALİMATI]: {custom_instruction.strip()}" if custom_instruction and custom_instruction.strip() else ""

	prompt = f"""Aşağıdaki anonim toplu okul taraması ve persantil istatistiklerini analiz ederek rapora eklenecek değerlendirme bölümünü hazırla:

[TARAMA KAPSAMI VE ANONİM İSTATİSTİKLER]:
{json.dumps(cohort_summary, ensure_ascii=False, default=str)}

[TESPİT EDİLEN EKSİK / GİRİLMEMİŞ TARAMA VERİLERİ]:
- {miss_txt}{extra_inst}

Lütfen doğrudan Word / Excel / PDF raporuna eklenecek şekilde (selamlama ve kod bloğu olmadan) şu 3 başlıkla Türkçe yanıt ver:
1. **Genel Tarama ve Gelişim Değerlendirmesi**
2. **Eksik / Girilmemiş Veriler ve Tamamlama Planı**
3. **Ekstra Değinilmesi Gereken Klinik Noktalar ve Okul İçi Öneriler** (risk kümeleri, özel gereksinimli/engelli öğrenci destekleri, veli bilgilendirme öncelikleri)"""

	return call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		temperature=0.25,
		timeout=24.0,
	)


def revise_report_ai_text(
	*,
	api_key: str,
	model: str,
	current_text: str,
	instruction: str,
	anon_context: str = "",
) -> tuple[bool, str, str]:
	"""Mevcut rapor değerlendirme metnini kullanıcının özel talimatına göre yeniden düzenler."""
	sys_inst = (
		"Sen resmi okul sağlığı raporlarını düzenleyen uzman bir editör ve okul sağlığı hemşiresisin. "
		"Kullanıcının verdiği düzenleme talimatını eksiksiz uygular, yalnızca doğrudan rapora konacak güncellenmiş temiz Türkçe metni döndürürsün."
	)
	ctx_line = f"\n[ANONİM BAĞLAM]: {anon_context}\n" if anon_context else ""
	prompt = f"""Aşağıdaki mevcut rapor metnini kullanıcının talimatına göre yeniden düzenle.{ctx_line}
[MEVCUT RAPOR METNİ]:
{current_text}

[DÜZENLEME TALİMATI]:
{instruction}

Yalnızca güncellenmiş nihai rapor metnini döndür (giriş cümlesi veya açıklama ekleme)."""
	return call_gemini(prompt, api_key=api_key, model=model, system_instruction=sys_inst, temperature=0.25, timeout=22.0)


# ==============================================================================
# 9. SENARYO 7: DOĞAL DİLLE REVİR VERİTABANI SORGULAMA (SALT-OKUNUR SQL)
# ==============================================================================

_DB_SCHEMA_HINT = """
Tablolar:
- classes(id, name, teacher_name)
- students(id, class_id, person_type, student_no, full_name, sex, birth_date, guardian_name, guardian_phone, notes, country, province, district, village, has_disability)
- health_profiles(student_id, guardian_name, guardian_phone, allergies, chronic_conditions, medical_history, medications, blood_group, emergency_notes, special_needs_codes, medical_devices, dietary_swallowing_notes, seizure_emergency_protocol, wears_glasses, vision_right, vision_left, hair_screening, hearing_right, hearing_left, health_behavior, harmful_habits, psychosocial_status, last_screened_at, disability_rate, disability_profile_json, ai_report_notes, ai_report_include)
- visits(id, student_id, visit_date, visit_time, complaint, symptoms, observation, action, guardian_contacted, referral, follow_up, follow_up_status, follow_up_result, incident_location, witness_teacher)
- vital_readings(id, student_id, measured_at, systolic_mmhg, diastolic_mmhg, pulse_bpm, temperature_c, spo2_pct, glucose_mg_dl, glucose_context, notes)
- measurements(id, student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
- school_screenings(id, student_id, screening_date, height_cm, weight_kg, head_circ_cm, temperature_c, pulse_bpm, spo2_pct, wears_glasses, vision_right, vision_left, hair_screening, hearing_right, hearing_left, health_behavior, harmful_habits, psychosocial_status, notes)
- student_medications(id, student_id, medication_name, med_type, route, dosage_pattern, dose_amount, dose_unit, meal_relation, remaining_volume, start_date, end_date, expiry_date, status)
- medication_logs(id, medication_id, student_id, log_date, scheduled_time, given_time, dose_given, dose_unit, status, administered_by, notes)
- care_plans(id, student_id, created_at, reviewed_by, nursing_diagnosis, observed_facts, goals, interventions, monitoring, escalation, status)
- isolation_records(id, student_id, disease_name, detected_date, expected_return_date, actual_return_date, status, notes)
- infirmary_inventory(id, item_name, category, quantity, unit, min_threshold, expiry_date, storage_location, notes)
- lab_reports(id, student_id, report_date, file_name, extracted_text)
- lab_results(id, report_id, test_name, value, unit, ref_low, ref_high, lab_flag)
- diabetes_profiles(student_id, is_diabetic, diabetes_type, insulin_brand, basal_insulin, delivery_method, pen_increment, icr_breakfast, icr_lunch, icr_snack, isf, target_bg, correction_threshold, hypo_threshold, hyper_ketone_threshold, insulin_duration_hours, tdd_units, max_single_bolus, cgm_device, glucagon_location, physician_name, emergency_protocol)
- school_meal_menus(id, year_month, menu_date, meal_type, items_json, total_carbs_g, source_note)
- insulin_bolus_logs(id, student_id, log_date, log_time, meal_type, bg_mg_dl, ketone_status, menu_carbs_g, menu_consumed_pct, extra_carbs_g, total_carbs_g, activity_level, activity_duration_min, activity_reduction_pct, icr_used, isf_used, iob_units, carb_dose_units, correction_dose_units, calculated_dose_units, applied_dose_units, injection_site, meal_summary, ai_commentary, administered_by, notes)
"""

_FORBIDDEN_SQL_RE = re.compile(
	r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|REPLACE|TRUNCATE|ATTACH|DETACH|PRAGMA|VACUUM|REINDEX)\b",
	re.IGNORECASE,
)


def natural_language_db_query(
	*,
	api_key: str,
	model: str,
	question: str,
	connection: sqlite3.Connection,
) -> tuple[bool, str, pd.DataFrame, str]:
	"""
	Türkçe doğal dil sorusunu yalnızca güvenli salt-okunur SELECT sorgusuna çevirir ve çalıştırır.
	Dönüş: (başarılı_mı, açıklama_veya_hata, sonuç_df, üretilen_sql)
	"""
	prompt = f"""Aşağıdaki SQLite veritabanı şemasını kullanarak kullanıcının Türkçe sorusunu yanıtlayan TEK BİR salt-okunur SQLite `SELECT` sorgusu üret.

[ŞEMA]:
{_DB_SCHEMA_HINT}

[KULLANICI SORUSU]: {question}

Kurallar:
1. Yalnızca `SELECT` sorgusu yaz. Asla veri değiştiren komut kullanma.
2. Sütun adlarını Türkçe ve anlaşılır `AS "Ad Soyad"`, `AS "Sınıf"` gibi takma adlarla döndür.
3. `students` tablosunu `LEFT JOIN classes c ON c.id = students.class_id` ile bağla.
4. Vital ölçümler (`temperature_c`, `spo2_pct`, `pulse_bpm`, `systolic_mmhg`, `diastolic_mmhg`, `glucose_mg_dl`) `vital_readings` tablosundadır (`visits` tablosunda `temperature_c` sütunu yoktur; `visits.observation` metin sütunudur). Ateş veya vital ölçüm sorgularında `vital_readings` tablosunu kullan.
5. En fazla `LIMIT 100` satır döndür.

Yanıtı yalnızca şu JSON formatında ver:
{{
  "sql": "SELECT ...",
  "explanation": "Sorgunun neyi listelediğinin 1 cümlelik Türkçe açıklaması"
}}"""

	ok, raw, _used_model = call_gemini(prompt, api_key=api_key, model=model, expect_json=True)
	if not ok:
		return False, raw, pd.DataFrame(), ""

	parsed = _parse_json_response(raw)
	sql_q = str(parsed.get("sql") or "").strip().rstrip(";")
	explanation = str(parsed.get("explanation") or "Sorgu çalıştırıldı.")

	if not sql_q or not sql_q.upper().startswith("SELECT") or _FORBIDDEN_SQL_RE.search(sql_q) or ";" in sql_q:
		return False, "Güvenlik politikası gereği yalnızca tekil salt-okunur SELECT sorgularına izin verilir.", pd.DataFrame(), sql_q

	try:
		df = pd.read_sql_query(sql_q, connection)
		return True, explanation, df, sql_q
	except Exception as exc:
		return False, f"SQL çalıştırma hatası: {exc}", pd.DataFrame(), sql_q


# ==============================================================================
# 10. YÜZEN SOHBET BALONU (MULTIMODAL: METİN + DOSYA + SESLİ KONUŞMA)
# ==============================================================================

def _extract_attachment_context(name: str, mime: str, raw_bytes: bytes, sensitive_names: list[str] | None = None) -> tuple[tuple[bytes, str] | None, str]:
	"""
	Dosya ekini Gemini için hazırlar.
	- PDF, Görsel (PNG/JPG/WEBP) ve Ses (WEBM/WAV/MP3/OGG/MP4) dosyalarını doğrudan inline_data olarak döndürür.
	- Excel, CSV, Word (.docx) ve Metin dosyalarını yerel olarak okuyup KVKK maskeli metin özetine çevirir.
	"""
	low = (name or "").lower()
	eff_mime = (mime or "").lower() or (mimetypes.guess_type(name)[0] or "application/octet-stream").lower()

	if low.endswith((".pdf", ".png", ".jpg", ".jpeg", ".webp", ".gif", ".webm", ".wav", ".mp3", ".ogg", ".m4a", ".mp4")) or eff_mime.startswith(("image/", "audio/", "application/pdf")):
		if low.endswith(".pdf"):
			eff_mime = "application/pdf"
		elif low.endswith((".webm",)) and "audio" not in eff_mime:
			eff_mime = "audio/webm"
		return (raw_bytes, eff_mime), f"[Ekli Dosya: {name} ({eff_mime})]"

	import io as _io
	extracted = ""
	try:
		if low.endswith((".xlsx", ".xls")):
			df = pd.read_excel(_io.BytesIO(raw_bytes)).head(60)
			extracted = df.to_csv(index=False)
		elif low.endswith(".csv"):
			df = pd.read_csv(_io.BytesIO(raw_bytes)).head(80)
			extracted = df.to_csv(index=False)
		elif low.endswith(".docx"):
			import docx as _docx
			doc = _docx.Document(_io.BytesIO(raw_bytes))
			extracted = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
		else:
			extracted = raw_bytes.decode("utf-8", errors="replace")
	except Exception as exc:
		extracted = f"(Dosya içeriği okunamadı: {exc})"

	safe_txt = anonymize_free_text(extracted[:10000], sensitive_names)
	return None, f"[Ekli Belge ({name}) İçeriği]:\n{safe_txt}"


def chat_with_gemini_assistant(
	*,
	api_key: str,
	model: str,
	user_message: str,
	history: list[dict[str, str]] | None = None,
	attachments: list[dict[str, str]] | None = None,
	voice_audio_b64: str = "",
	voice_audio_mime: str = "audio/wav",
	anon_person_context: str = "",
	school_context: str = "",
	sensitive_names: list[str] | None = None,
	available_classes: list[str] | None = None,
) -> tuple[bool, str, str, str, dict[str, Any] | None, dict[str, Any] | None]:
	"""
	Yüzen sohbet balonundan gelen metin, çoklu dosya eki ve/veya ses kaydını yanıtlar.
	Kullanıcı Word/Excel/rapor veya sınıf oluşturma isterse `export_spec`,
	öğrenciye ölçüm/vital/profil/belge işleme isterse `write_spec` döndürür.
	Dönüş: (başarılı_mı, yanıt_metni, kullanılan_model, ses_çözümü, export_spec, write_spec)
	"""
	cls_hint = ", ".join(available_classes[:30]) if available_classes else "Belirtilmedi"
	sys_inst = (
		"Sen 'Revir Defteri' masaüstü okul sağlığı yazılımının içinde yer alan KVKK korumalı 'Gemini Klinik Asistan'sın. "
		"Okul revir hemşiresine pediatrik triyaj, büyüme-gelişme (WHO persantil/Z-skoru), NANDA bakım planları, "
		"engellilik/özel gereksinim uyarlamaları, ilaç etkileşimleri, tahlil/evrak yorumlama ve okul sağlığı yönetimi konularında hızlı, doğru, Türkçe ve öz yanıtlar verirsin.\n\n"
		"ÖNEMLİ YETENEK — GERÇEK WORD (.docx) VE EXCEL (.xlsx) RAPOR OLUŞTURMA:\n"
		"Programın yerel Word ve Excel oluşturma motorlarını doğrudan tetikleyebilirsin! "
		"Kullanıcı senden herhangi bir raporu, listeyi, tabloyu, tutanağı veya belgeyi Word (.docx), Excel (.xlsx), PDF/HTML olarak hazırlamanı/vermeni isterse, "
		"önce kullanıcıya kısa ve net bir Türkçe açıklama yaz, ardından yanıtının EN SONUNA tek satırda şu JSON komut etiketini ekle:\n"
		"[[REVIR_EXPORT:{...}]]\n\n"
		"Kullanabileceğin `action` değerleri:\n"
		"1) Seçili kişinin tam resmi sağlık raporu (Word): {\"action\":\"student_report\",\"format\":\"docx\"}\n"
		"2) Belirli bir sınıfın toplu sağlık raporu (Word): {\"action\":\"class_report\",\"class_name\":\"5-A\",\"format\":\"docx\"}\n"
		"3) MEB Dönem/Yıl Sonu resmi faaliyet raporu (Word): {\"action\":\"meb_activity\",\"format\":\"docx\",\"extra_notes\":\"...\"}\n"
		"4) Seçili kişinin aşı kartı (Word): {\"action\":\"vaccination_card\",\"format\":\"docx\"}\n"
		"5) Seçili kişinin nöbet & özel gereksinim acil protokolü (Word): {\"action\":\"seizure_protocol\",\"format\":\"docx\"}\n"
		"6) Revir stok & malzeme envanter raporu (Word): {\"action\":\"inventory_report\",\"format\":\"docx\"}\n"
		"7) Buzdolabı soğuk zincir aylık ısı çizelgesi (Word): {\"action\":\"fridge_report\",\"format\":\"docx\"}\n"
		"8) Tüm veritabanı tam Excel yedeği/kitabı: {\"action\":\"all_workbook\",\"format\":\"xlsx\"}\n"
		"9) Veritabanından özel filtreli liste/rapor (Excel ve/veya Word, KVKK gereği yerel SQLite'da çalışır):\n"
		"   {\"action\":\"sql_export\",\"format\":\"xlsx\" veya \"docx\" veya \"both\",\"title\":\"Rapor Başlığı\",\"sql\":\"SELECT ...\"}\n"
		"   Not: SQL yalnızca güvenli salt-okunur SELECT olmalı, Türkçe sütun takma adları kullanmalı ve en fazla LIMIT 200 olmalı.\n"
		"10) Özel metin/dilekçe/veli mektubu/BEP uyarlama/özel tablo belgesi (Word veya Excel):\n"
		"   {\"action\":\"custom_doc\",\"format\":\"docx\" veya \"xlsx\" veya \"both\",\"title\":\"Belge Başlığı\",\"subtitle\":\"Alt Başlık\",\"sections\":[{\"heading\":\"Başlık\",\"body\":\"İçerik...\"}],\"table\":{\"headers\":[\"Sütun1\",\"Sütun2\"],\"rows\":[[\"Değer1\",\"Değer2\"]]}}\n"
		"11) Seçili öğrencinin aylık yemek menüsüne göre tahmini insülin planı (Word/Excel): {\"action\":\"insulin_monthly_plan\",\"format\":\"docx\" veya \"xlsx\" veya \"both\",\"year_month\":\"YYYY-MM\"}\n"
		"12) Seçili öğrencinin insülin & kan şekeri izlem defteri raporu (Word/Excel): {\"action\":\"insulin_log_report\",\"format\":\"docx\" veya \"xlsx\" veya \"both\"}\n"
		"13) Okul sınıfları / şubeleri oluşturma (kullanıcı açıkça 'şu sınıfları oluştur / ekle' derse): {\"action\":\"create_classes\",\"classes\":[{\"name\":\"5-A\",\"teacher_name\":\"\"}]}\n\n"
		"ÖNEMLİ YETENEK — ÖĞRENCİ / KİŞİ DOSYASINA ÖLÇÜM, VİTAL, TAHLİL, İLAÇ, PROFİL VEYA BELGE İŞLEME:\n"
		"Kullanıcı sohbet üzerinden yazılı/sesli bir ölçüm, vital, alerji, kronik hastalık, ilaç, revir notu söylerse "
		"VEYA bir belge/tahlil/rapor/reçete/fotoğraf ekleyip 'öğrenciye işle / kaydet / profiline ekle / ölçümünü gir' derse, "
		"önce hemşireye kısa ve net bir Türkçe özet yaz (örn: 'Çıkarılan veriler onay kartında hazırlandı'), ardından yanıtının EN SONUNA tek satırda şu JSON komut etiketini ekle:\n"
		"[[REVIR_WRITE:{...}]]\n"
		"Şema (sadece gerçekten geçen/okunan alanları doldur; diğerlerini boş bırak veya ekleme):\n"
		"{\"person_hint\":\"Kullanıcının söylediği veya belgede okunan ad soyad (yoksa boş)\",\"student_no_hint\":\"\",\"class_hint\":\"\",\"document_type\":\"Ölçüm / Tahlil / Rapor / Profil\",\"document_date\":\"YYYY-MM-DD\",\"clinical_summary\":\"Kısa özet\","
		"\"measurements\":{\"measured_on\":\"YYYY-MM-DD\",\"height_cm\":null,\"weight_kg\":null,\"head_circ_cm\":null},"
		"\"vitals\":{\"measured_at\":\"YYYY-MM-DD HH:MM\",\"temperature_c\":null,\"pulse_bpm\":null,\"spo2_pct\":null,\"systolic_bp\":null,\"diastolic_bp\":null,\"glucose_mg_dl\":null},"
		"\"health_profile\":{\"allergies\":\"\",\"chronic_conditions\":\"\",\"blood_group\":\"\",\"regular_medications\":\"\",\"emergency_notes\":\"\"},"
		"\"lab_results\":[{\"test_name\":\"\",\"value_num\":null,\"value_text\":\"\",\"unit\":\"\",\"ref_range\":\"\",\"flag\":\"Normal\",\"measured_on\":\"YYYY-MM-DD\"}],"
		"\"medications\":[{\"medication_name\":\"\",\"med_type\":\"Düzenli İlaç\",\"dosage\":\"\",\"schedule_times\":\"\",\"notes\":\"\"}],"
		"\"disability\":{\"has_disability\":false,\"disability_rate\":null,\"categories\":[],\"school_adaptations\":\"\",\"report_summary\":\"\"},"
		"\"diabetes\":{\"is_diabetic\":false,\"diabetes_type\":\"\",\"icr_breakfast\":null,\"icr_lunch\":null,\"icr_snack\":null,\"isf\":null,\"target_bg\":null,\"insulin_brand\":\"\"},"
		"\"visit_note_suggestion\":\"\"}\n\n"
		"Kullanıcı kayıt/veri işleme istemiyorsa [[REVIR_WRITE:...]] EKLEME; dosya/rapor/sınıf istemiyorsa [[REVIR_EXPORT:...]] EKLEME. "
		"Kısa, maddeli ve okunması kolay yanıtlar üret."
	)

	inline_files: list[tuple[bytes, str]] = []
	extra_doc_blocks: list[str] = []

	if attachments:
		for att in attachments[:5]:
			fname = str(att.get("name") or "dosya")
			fmime = str(att.get("type") or "")
			fb64 = str(att.get("data_b64") or "")
			if not fb64:
				continue
			try:
				fbytes = base64.b64decode(fb64)
				inline_item, txt_block = _extract_attachment_context(fname, fmime, fbytes, sensitive_names)
				if inline_item is not None:
					inline_files.append(inline_item)
				if txt_block:
					extra_doc_blocks.append(txt_block)
			except Exception:
				continue

	has_voice = False
	if voice_audio_b64:
		try:
			vbytes = base64.b64decode(voice_audio_b64)
			if vbytes:
				inline_files.append((vbytes, voice_audio_mime or "audio/wav"))
				has_voice = True
		except Exception:
			pass

	hist_lines: list[str] = []
	if history:
		for msg in history[-10:]:
			role_lbl = "Hemşire" if msg.get("role") == "user" else "Asistan"
			content_s = anonymize_free_text(str(msg.get("text") or ""), sensitive_names)
			if content_s:
				hist_lines.append(f"{role_lbl}: {content_s}")

	safe_user_msg = anonymize_free_text(user_message, sensitive_names)
	ctx_parts: list[str] = []
	if school_context:
		ctx_parts.append(f"[OKUL & REVİR DURUMU]: {school_context} | Mevcut Sınıflar: {cls_hint}")
	if anon_person_context:
		ctx_parts.append(f"[SEÇİLİ KİŞİ ANONİM KLİNİK PROFİLİ]: {anon_person_context}")
	ctx_parts.append(f"[VERİTABANI ŞEMASI (Sadece sql_export gerekirse kullan)]:\n{_DB_SCHEMA_HINT}")
	if hist_lines:
		hist_joined = "\n".join(hist_lines)
		ctx_parts.append(f"[ÖNCEKİ SOHBET GEÇMİŞİ]:\n{hist_joined}")
	if extra_doc_blocks:
		docs_joined = "\n\n".join(extra_doc_blocks)
		ctx_parts.append(docs_joined)

	if has_voice and not safe_user_msg.strip():
		ctx_parts.append(
			"[KULLANICI SESLİ MESAJ GÖNDERDİ]: Ekli Türkçe ses kaydını dikkatle dinle. "
			"Eğer ses kaydı sessizse, yalnızca ortam gürültüsü varsa veya anlaşılır bir insan konuşması yoksa ASLA metin uydurma; "
			"ilk satıra `[SES_ÇÖZÜMÜ]: (Anlaşılamadı)` yaz ve nazikçe tekrar konuşmasını iste. "
			"Konuşma anlaşılırsa ilk satıra kullanıcının söylediği tam cümleyi `[SES_ÇÖZÜMÜ]: <söylenen cümle>` biçiminde yaz, "
			"alt satırda ise sorusuna veya rapor/Word/Excel talebine yanıt ver."
		)
	else:
		ctx_parts.append(f"[HEMŞİRENİN MESAJI]: {safe_user_msg or 'Ekli dosyayı/sesi inceleyip yorumla.'}")

	full_prompt = "\n\n".join(ctx_parts)
	ok, raw_reply, used_model = call_gemini(
		full_prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		inline_files=inline_files if inline_files else None,
		timeout=26.0,
	)
	if not ok:
		return False, raw_reply, used_model, "", None, None

	transcribed = ""
	clean_reply = raw_reply
	m_tr = re.match(r"^\s*\[SES_ÇÖZÜMÜ\]\s*:\s*(.+?)(?:\n+|$)(.*)", raw_reply, re.S)
	if m_tr:
		transcribed = m_tr.group(1).strip()
		clean_reply = m_tr.group(2).strip() or transcribed

	def _extract_directive_tag(text_in: str, tag_prefix: str) -> tuple[str, dict[str, Any] | None]:
		idx_s = text_in.find(tag_prefix)
		if idx_s == -1:
			return text_in, None
		# Find matching closing ]]
		idx_e = text_in.find("]]", idx_s + len(tag_prefix))
		if idx_e == -1:
			return text_in, None
		# Handle nested ]] inside JSON arrays by balancing braces
		json_start = text_in.find("{", idx_s + len(tag_prefix))
		if json_start != -1 and json_start < idx_e:
			depth = 0
			in_str = False
			esc = False
			end_brace = -1
			for pos in range(json_start, len(text_in)):
				ch = text_in[pos]
				if esc:
					esc = False
					continue
				if ch == "\\":
					esc = True
					continue
				if ch == '"':
					in_str = not in_str
					continue
				if not in_str:
					if ch == "{":
						depth += 1
					elif ch == "}":
						depth -= 1
						if depth == 0:
							end_brace = pos
							break
			if end_brace != -1:
				close_tag = text_in.find("]]", end_brace)
				if close_tag != -1:
					parsed_obj = _parse_json_response(text_in[json_start:end_brace + 1])
					stripped = (text_in[:idx_s] + text_in[close_tag + 2:]).strip()
					return stripped, parsed_obj if isinstance(parsed_obj, dict) and parsed_obj else None
		chunk = text_in[idx_s + len(tag_prefix):idx_e].strip()
		parsed_obj = _parse_json_response(chunk)
		stripped = (text_in[:idx_s] + text_in[idx_e + 2:]).strip()
		return stripped, parsed_obj if isinstance(parsed_obj, dict) and parsed_obj else None

	clean_reply, export_spec = _extract_directive_tag(clean_reply, "[[REVIR_EXPORT:")
	clean_reply, write_spec = _extract_directive_tag(clean_reply, "[[REVIR_WRITE:")

	return True, clean_reply, used_model, transcribed, export_spec, write_spec


# ==============================================================================
# 11. PEDİATRİK DİYABET & İNSÜLİN DOZ HESAPLAMA YAPAY ZEKÂ MODÜLLERİ
# ==============================================================================

def parse_monthly_school_menu_with_gemini(
	*,
	api_key: str,
	model: str,
	year_month: str,
	file_bytes: bytes | None = None,
	file_name: str = "",
	raw_text: str = "",
	default_meal_type: str = "Öğle Yemeği",
) -> tuple[bool, list[dict[str, Any]], str]:
	"""
	Aylık okul yemek programını (PDF, Fotoğraf, Excel, Word veya düz metin) Gemini ile okuyup
	gün gün yemek kalemlerine ve Türk okul porsiyonu tahmini karbonhidrat (g) değerlerine ayrıştırır.
	"""
	inline_files: list[tuple[bytes, str]] = []
	extra_txt = (raw_text or "").strip()

	if file_bytes and file_name:
		item, txt_blk = _extract_attachment_context(file_name, "", file_bytes, sensitive_names=None)
		if item is not None:
			inline_files.append(item)
		if txt_blk and not item:
			extra_txt = (extra_txt + "\n\n" + txt_blk).strip()

	if not inline_files and not extra_txt:
		return False, [], "Lütfen aylık yemek listesi dosyası yükleyin veya menü metnini yapıştırın."

	sys_inst = (
		"Sen çocuk endokrinolojisi ve okul beslenmesi alanında uzman bir Türk pediatrik diyetisyenisin. "
		"Türk okul yemekhanelerindeki standart çocuk porsiyonlarına (1 kase çorba, 1 porsiyon ana yemek, 1 küçük kase pilav/makarna, 1 dilim ekmek, 1 küçük meyve/yoğurt) "
		"göre gerçekçi, güvenilir karbonhidrat (gram) tahminleri ve glisemik indeks (Düşük/Orta/Yüksek) sınıflandırması yaparsın."
	)
	txt_section = f"\n[YEMEK LİSTESİ METNİ / TABLOSU]:\n{extra_txt[:15000]}" if extra_txt else ""
	prompt = f"""Ekli veya aşağıda verilen aylık okul yemek listesini dikkatlice incele ve '{year_month}' ayı için gün gün yemekleri ve her yemeğin çocuk porsiyonu tahmini karbonhidrat (gram) değerini çıkar.
Varsayılan öğün türü: '{default_meal_type}' (belgede Kahvaltı / İkindi ayrı verilmişse ilgili `meal_type` değerini yaz).{txt_section}

Yanıtı YALNIZCA aşağıdaki JSON formatında döndür:
{{
  "days": [
    {{
      "menu_date": "{year_month}-01",
      "meal_type": "{default_meal_type}",
      "items": [
        {{
          "name": "Mercimek Çorbası",
          "portion": "1 orta kase (200 ml)",
          "carbs_g": 18,
          "gi": "Orta",
          "fat_protein_note": ""
        }},
        {{
          "name": "Pirinç Pilavı",
          "portion": "4-5 yemek kaşığı",
          "carbs_g": 30,
          "gi": "Yüksek",
          "fat_protein_note": ""
        }}
      ]
    }}
  ]
}}

Kurallar:
- `menu_date` mutlaka `YYYY-MM-DD` biçiminde olsun (belgede yalnızca gün numarası varsa `{year_month}-GG` olarak tamamla).
- Ekmek menüde yazmasa bile listede yer alan her kalemi ayrı öğe olarak gir; su/çay gibi 0g karb içerenleri 0 yaz.
- Yağlı/hamur işi/pizza/börek/kızartma/mantı gibi gecikmeli kan şekeri yükselişi yapan yemeklerde `fat_protein_note` alanına kısa uyarı yaz (örn: "Yağ/protein yüksek - geç yükseliş riski").
- Sadece belgede yer alan günleri döndür."""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		expect_json=True,
		inline_files=inline_files if inline_files else None,
		temperature=0.15,
		timeout=40.0,
	)
	if not ok:
		return False, [], raw

	parsed = _parse_json_response(raw)
	days = parsed.get("days") if isinstance(parsed, dict) else None
	if not isinstance(days, list) or not days:
		return False, [], "Belgeden gün bazlı yemek listesi ayrıştırılamadı. Lütfen daha net bir dosya veya metin deneyin."

	cleaned_days: list[dict[str, Any]] = []
	for d in days:
		if not isinstance(d, dict):
			continue
		mdate = str(d.get("menu_date") or "").strip()[:10]
		if not re.match(r"^\d{4}-\d{2}-\d{2}$", mdate):
			continue
		mtype = str(d.get("meal_type") or default_meal_type).strip() or default_meal_type
		raw_items = d.get("items") or []
		clean_items: list[dict[str, Any]] = []
		for it in raw_items:
			if not isinstance(it, dict):
				continue
			nm = str(it.get("name") or "").strip()
			if not nm:
				continue
			try:
				cg = max(0.0, round(float(str(it.get("carbs_g") or 0).replace(",", ".")), 1))
			except Exception:
				cg = 0.0
			clean_items.append({
				"name": nm,
				"portion": str(it.get("portion") or "1 porsiyon").strip(),
				"carbs_g": cg,
				"gi": str(it.get("gi") or "Orta").strip(),
				"fat_protein_note": str(it.get("fat_protein_note") or "").strip(),
			})
		if clean_items:
			cleaned_days.append({
				"menu_date": mdate,
				"meal_type": mtype,
				"items": clean_items,
				"total_carbs_g": round(sum(x["carbs_g"] for x in clean_items), 1),
			})
	return True, cleaned_days, used_m


def estimate_extra_foods_carbs_with_gemini(
	*,
	api_key: str,
	model: str,
	food_description: str = "",
	file_bytes: bytes | None = None,
	file_name: str = "",
	student_age_years: float | None = None,
) -> tuple[bool, dict[str, Any], str]:
	"""Ekstra yiyecek metni veya tabak/atıştırmalık fotoğrafından karbonhidrat (g) ve glisemik etki hesaplar."""
	inline_files: list[tuple[bytes, str]] = []
	if file_bytes and file_name:
		mime = mimetypes.guess_type(file_name)[0] or "image/jpeg"
		inline_files.append((file_bytes, mime))
	desc = (food_description or "").strip()
	if not desc and not inline_files:
		return False, {"error": "Yiyecek açıklaması yazın veya yemek fotoğrafı yükleyin."}, ""

	age_hint = f"Öğrenci yaşı: yaklaşık {student_age_years:.1f} yaş." if student_age_years else ""
	prompt = f"""Bir pediatrik diyabet diyetisyeni olarak aşağıdaki ekstra yiyecek / atıştırmalık açıklamasını ve/veya ekli yemek fotoğrafını analiz et, her yiyeceğin karbonhidrat miktarını (gram) hesapla.
{age_hint}
[YİYECEK AÇIKLAMASI]: {desc or 'Ekli fotoğraftaki yiyecekleri ve porsiyonlarını tahmin et.'}

Yalnızca şu JSON şemasıyla Türkçe yanıt ver:
{{
  "items": [
    {{"name": "Yiyecek adı", "portion": "Tahmini porsiyon", "carbs_g": 15.0, "gi": "Düşük | Orta | Yüksek"}}
  ],
  "total_carbs_g": 15.0,
  "glycemic_note": "Glisemik hız, yağ/protein gecikmeli yükseliş riski veya porsiyon uyarısı (1-2 cümle)"
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		inline_files=inline_files if inline_files else None,
		temperature=0.15,
		timeout=25.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	raw_items = parsed.get("items") or []
	clean_items: list[dict[str, Any]] = []
	for it in raw_items:
		if not isinstance(it, dict):
			continue
		nm = str(it.get("name") or "").strip()
		if not nm:
			continue
		try:
			cg = max(0.0, round(float(str(it.get("carbs_g") or 0).replace(",", ".")), 1))
		except Exception:
			cg = 0.0
		clean_items.append({
			"name": nm,
			"portion": str(it.get("portion") or "1 porsiyon").strip(),
			"carbs_g": cg,
			"gi": str(it.get("gi") or "Orta").strip(),
		})
	tot = round(sum(x["carbs_g"] for x in clean_items), 1)
	return True, {
		"items": clean_items,
		"total_carbs_g": tot,
		"glycemic_note": str(parsed.get("glycemic_note") or "").strip(),
	}, used_m


def analyze_insulin_dose_with_gemini(
	*,
	api_key: str,
	model: str,
	anon_student_context: str,
	growth_info: dict[str, Any],
	diabetes_profile: dict[str, Any],
	calc_result: dict[str, Any],
	meal_details: dict[str, Any],
	recent_bg_trend: list[dict[str, Any]] | None = None,
	school_name: str = "Okul Reviri",
	nurse_name: str = "Okul Revir Sorumlusu",
	extra_instruction: str = "",
) -> tuple[bool, dict[str, str], str]:
	"""
	Python tarafından kesin matematikle hesaplanan insülin dozunu; öğrencinin yaşı, boyu, kilosu,
	WHO persantil/Z-skoru, TDD/kg oranı, özel gereksinim/alerji profili, yemek menüsü, ekstra yiyecekleri,
	fiziksel aktivitesi ve son kan şekeri eğilimi ışığında pediatrik hemşirelik açısından yorumlar.
	"""
	sys_inst = (
		"Sen çocuk endokrinolojisi ve okulda diyabet hemşireliği (ISPAD / T.C. Sağlık Bakanlığı Okulda Diyabet Programı) "
		"konusunda uzman bir klinik karar destek asistanısın. "
		"Matematiksel doz hesabını değiştirme veya yeni reçete yazma; sistemin hekim parametreleriyle (İKO/İDF) hesapladığı "
		"dozu öğrencinin yaşı, boyu, kilosu, büyüme persantili, günlük insülin/kg (TDD/kg) oranı, yemek içeriğinin glisemik/yağ-protein profili, "
		"fiziksel aktivitesi, aktif insülini (IOB) ve varsa engellilik/alerji/çölyak durumuna göre klinik güvenlik açısından değerlendir."
	)
	inst_line = f"\n[HEMŞİRENİN ÖZEL SORUSU / NOTU]: {extra_instruction.strip()}" if extra_instruction and extra_instruction.strip() else ""
	prompt = f"""Aşağıdaki anonim pediatrik diyabet öğrencisinin güncel öğün, aktivite ve insülin hesaplama verilerini değerlendir:

[ANONİM ÖĞRENCİ KLİNİK & ENGEL/ALERJİ PROFİLİ]:
{anon_student_context}

[YAŞ, BOY, KİLO, VKİ & WHO PERSANTİL BİLGİSİ]:
{json.dumps(growth_info, ensure_ascii=False, default=str)}

[HEKİM DİYABET & İNSÜLİN PARAMETRELERİ]:
{json.dumps(diabetes_profile, ensure_ascii=False, default=str)}

[ÖĞÜN MENÜSÜ, EKSTRA YİYECEKLER & FİZİKSEL AKTİVİTE]:
{json.dumps(meal_details, ensure_ascii=False, default=str)}

[SİSTEMİN HESAPLADIĞI KESİN DOZ & GÜVENLİK BAYRAKLARI]:
{json.dumps(calc_result, ensure_ascii=False, default=str)}

[SON KAN ŞEKERİ & İNSÜLİN KAYIT EĞİLİMİ]:
{json.dumps((recent_bg_trend or [])[:10], ensure_ascii=False, default=str)}{inst_line}

Yanıtı YALNIZCA şu JSON şemasıyla Türkçe döndür:
{{
  "safety_badge": "✅ Uygun / Güvenli Aralıkta | ⚠️ Yakın İzlem Önerilir | 🚨 Kritik Klinik Dikkat!",
  "anthropometry_tdd_comment": "Öğrencinin yaşı, boyu, kilosu, VKİ/persantil durumu ve (varsa) TDD/kg oranına göre insülin gereksinimi ve güvenlik uyumu yorumu (2-3 cümle)",
  "meal_glycemic_comment": "Menüdeki yemekler + ekstra yediklerinin karbonhidrat yükü, glisemik hızı, yağ/protein gecikmeli yükseliş olasılığı, varsa çölyak/alerji/tüketim yüzdesi ve enjeksiyon zamanlaması önerisi (2-3 cümle)",
  "activity_hypo_comment": "Planlanan veya yapılan fiziksel aktivite/beden eğitiminin kan şekerine etkisi, IOB (aktif insülin) etkileşimi ve hipoglisemi/keton koruma uyarısı (2 cümle)",
  "nurse_action_plan": "Okul hemşiresi için somut adımlar: enjeksiyon öncesi/sonrası kontrol saati (örn. 2. saat KŞ ölçümü), enjeksiyon bölgesi rotasyonu veya engellilik/özel durum uyarlaması (maddeli, kısa)",
  "guardian_sms_draft": "Veliye gönderilecek sakin, net bilgilendirme mesajı (Ölçülen KŞ, tüketilen karb, aktivite notu ve uygulanan/planlanan doz özeti; isim yerine [ÖĞRENCİ_ADI] kullan; Kurum: {school_name}, İmza: {nurse_name})"
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		system_instruction=sys_inst,
		expect_json=True,
		temperature=0.2,
		timeout=24.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Gemini diyabet klinik yorumu JSON olarak ayrıştırılamadı."}, used_m
	return True, {k: _stringify_ai_field(v) for k, v in parsed.items()}, used_m


def analyze_glycemic_trend_with_gemini(
	*,
	api_key: str,
	model: str,
	anon_student_context: str,
	growth_info: dict[str, Any],
	tir_stats: dict[str, Any],
	logs_summary: list[dict[str, Any]],
) -> tuple[bool, str, str]:
	"""Öğrencinin son 14-30 günlük okul içi kan şekeri ve insülin kayıtlarını analiz edip hekim/veli kontrol özeti üretir."""
	prompt = f"""Çocuk endokrinolojisi kontrolüne veya veli bilgilendirmesine destek olmak üzere aşağıdaki anonim öğrencinin okul içi kan şekeri, Hedef Aralıkta Kalma (TIR) ve insülin uygulama kayıtlarını analiz et:

[ANONİM PROFİL]: {anon_student_context}
[YAŞ / BOY / KİLO / PERSANTİL]: {json.dumps(growth_info, ensure_ascii=False, default=str)}
[GLİSEMİK İSTATİSTİKLER (TIR)]: {json.dumps(tir_stats, ensure_ascii=False, default=str)}
[SON KAYITLAR]: {json.dumps(logs_summary[:30], ensure_ascii=False, default=str)}

Türkçe olarak şu 4 kısa başlıkla raporla (kod bloğu kullanma):
1. **📊 Glisemik Seyir & Hedef Aralık (TIR) Değerlendirmesi**
2. **⚠️ Tekrarlayan Hipoglisemi / Hiperglisemi Saatleri ve Olası Tetikleyiciler** (öğünler, ekstra yiyecekler veya beden eğitimi ilişkisi)
3. **💉 Enjeksiyon Rotasyonu, Doz Uyumu & Büyüme (Yaş/Kilo) İlişkisi**
4. **👩‍⚕️ Hekim Kontrolü & Okul İçi İzlem İçin Önerilen Notlar**"""

	return call_gemini(prompt, api_key=api_key, model=model, temperature=0.25, timeout=24.0)


# ==============================================================================
# 12. AKILLI SINIF/ÖĞRENCİ OLUŞTURMA, EVRAK OKUYUCU, ÇIKARIM & OFİS/PDF YARDIMCISI
# ==============================================================================

def parse_classes_and_students_with_gemini(
	*,
	api_key: str,
	model: str,
	instruction_text: str = "",
	file_bytes: bytes | None = None,
	file_name: str = "",
	existing_classes: list[str] | None = None,
) -> tuple[bool, dict[str, Any], str]:
	"""
	Serbest Türkçe talimattan veya yüklenen okul şube/öğrenci listesinden (PDF, Word, Excel, CSV, fotoğraf)
	oluşturulacak sınıfları ve (varsa) öğrencileri yapılandırılmış JSON olarak çıkarır.
	"""
	inline_files: list[tuple[bytes, str]] = []
	extracted_txt = ""
	if file_bytes and file_name:
		item, txt_blk = _extract_attachment_context(file_name, "", file_bytes, sensitive_names=None)
		if item is not None:
			inline_files.append(item)
		if txt_blk:
			extracted_txt = txt_blk[:14000]

	cls_str = ", ".join((existing_classes or [])[:40]) or "Henüz sınıf yok"
	prompt = f"""Okul reviri yönetim yazılımı için aşağıdaki talimat ve/veya belgeden oluşturulacak okul sınıflarını/şubelerini ve (varsa) öğrenci listesini çıkar.

[MEVCUT SINIFLAR]: {cls_str}
[KULLANICI TALİMATI]: {instruction_text or 'Ekli belgedeki tüm sınıfları, sınıf öğretmenlerini ve varsa öğrenci listesini çıkar.'}
{('[BELGE İÇERİĞİ]: ' + extracted_txt) if extracted_txt else ''}

SADECE aşağıdaki JSON şemasını döndür (başka metin yazma):
{{
  "classes": [
    {{"name": "5-A", "teacher_name": "Varsa öğretmen adı yoksa boş string"}}
  ],
  "students": [
    {{
      "full_name": "Öğrenci Adı Soyadı",
      "class_name": "5-A",
      "student_no": "Varsa okul no",
      "sex": "K veya E veya boş",
      "birth_date": "YYYY-MM-DD veya boş",
      "person_type": "Öğrenci veya Sivil / diğer kişi"
    }}
  ],
  "summary": "Çıkarılan sınıf ve öğrenci sayısının kısa Türkçe özeti"
}}

KURALLAR:
- İl, ilçe, köy, mahalle veya okul adlarını ASLA sınıf olarak ekleme.
- Kullanıcı '1'den 4'e kadar A, B, C şubeleri oluştur' gibi bir talimat verdiyse tüm kombinasyonları tek tek üret (1-A, 1-B, 1-C, ..., 4-C).
- Belgede sadece sınıf listesi varsa `students` boş dizi `[]` olsun.
- Doğum tarihlerini mümkünse `YYYY-MM-DD` biçimine çevir."""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		inline_files=inline_files if inline_files else None,
		expect_json=True,
		temperature=0.1,
		timeout=30.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Gemini yanıtı JSON olarak ayrıştırılamadı."}, used_m
	classes_out = []
	for c in (parsed.get("classes") or []):
		if isinstance(c, dict) and str(c.get("name") or "").strip():
			classes_out.append({
				"name": str(c.get("name") or "").strip(),
				"teacher_name": str(c.get("teacher_name") or "").strip(),
			})
		elif isinstance(c, str) and c.strip():
			classes_out.append({"name": c.strip(), "teacher_name": ""})
	students_out = []
	for s in (parsed.get("students") or []):
		if isinstance(s, dict) and str(s.get("full_name") or "").strip():
			students_out.append({
				"full_name": str(s.get("full_name") or "").strip(),
				"class_name": str(s.get("class_name") or "").strip(),
				"student_no": str(s.get("student_no") or "").strip(),
				"sex": str(s.get("sex") or "").strip().upper()[:1] if str(s.get("sex") or "").strip() else "",
				"birth_date": str(s.get("birth_date") or "").strip(),
				"person_type": str(s.get("person_type") or "Öğrenci").strip() or "Öğrenci",
			})
	return True, {
		"classes": classes_out,
		"students": students_out,
		"summary": str(parsed.get("summary") or f"{len(classes_out)} sınıf, {len(students_out)} kişi algılandı."),
	}, used_m


def extract_clinical_record_from_document_with_gemini(
	*,
	api_key: str,
	model: str,
	file_bytes: bytes,
	file_name: str,
	extracted_text: str = "",
	extra_instruction: str = "",
) -> tuple[bool, dict[str, Any], str]:
	"""
	Herhangi bir sağlık raporu, tahlil, reçete, ÇÖZGER/engellilik raporu, endokrin/diyabet raporu,
	ölçüm veya okul sağlık evrakını (PDF, fotoğraf, Word, Excel, TXT) okuyarak hangi öğrenciye ait olduğuna
	dair eşleştirme ipuçları ve veritabanına işlenebilecek tüm yapılandırılmış klinik alanları çıkarır.
	"""
	inline_files: list[tuple[bytes, str]] = []
	doc_txt = extracted_text or ""
	if file_bytes and file_name:
		item, txt_blk = _extract_attachment_context(file_name, "", file_bytes, sensitive_names=None)
		if item is not None:
			inline_files.append(item)
		if txt_blk and not doc_txt:
			doc_txt = txt_blk[:14000]

	prompt = f"""Okul revir hemşiresinin yüklediği aşağıdaki belgeyi dikkatle incele ve öğrencinin revir dosyasına işlenebilecek TÜM klinik, sağlık profili, laboratuvar, ilaç, vital, ölçüm, engellilik/özel gereksinim ve diyabet bilgilerini çıkar.

[DOSYA ADI]: {file_name}
[HEMŞİRE NOTU]: {extra_instruction or 'Tüm klinik verileri eksiksiz oku ve eşleştirme için belgedeki kişi/sınıf/numara bilgisini çıkar.'}
{('[YEREL METİN ÖNİZLEMESİ]: ' + doc_txt[:12000]) if doc_txt else ''}

SADECE aşağıdaki JSON yapısını döndür (bulunmayan alanları boş string, null veya [] bırak; ASLA veri uydurma!):
{{
  "document_type": "Sağlık Raporu | Laboratuvar Tahlili | Reçete / İlaç | Engellilik / ÇÖZGER / RAM | Diyabet / Endokrin | Ölçüm / Büyüme | Aşı Kartı | Revir / Acil Epikriz | Diğer",
  "detected_person_hints": {{
    "name_hint": "Belgede yazan hasta/öğrenci adı soyadı (yoksa boş)",
    "student_no_hint": "Varsa okul/öğrenci numarası",
    "class_hint": "Varsa sınıf/şube (örn 5-A)",
    "birth_date_hint": "Varsa doğum tarihi YYYY-MM-DD",
    "sex_hint": "K veya E veya boş"
  }},
  "document_date": "YYYY-MM-DD veya boş",
  "clinical_summary": "Belgenin okul hemşiresi için 2-4 cümlelik net Türkçe klinik özeti",
  "health_profile": {{
    "allergies": "Belgede geçen alerjiler (yoksa boş)",
    "chronic_conditions": "Kronik hastalıklar / tanılar (yoksa boş)",
    "blood_group": "Kan grubu örn A Rh(+) (yoksa boş)",
    "regular_medications": "Düzenli kullanılan ilaçlar özeti (yoksa boş)",
    "emergency_notes": "Okulda acil dikkat/müdahale notu (yoksa boş)"
  }},
  "lab_results": [
    {{"test_name": "Hemoglobin", "value_num": 11.2, "value_text": "11.2", "unit": "g/dL", "ref_range": "11.5-15.5", "flag": "Düşük", "measured_on": "YYYY-MM-DD"}}
  ],
  "vitals": {{
    "measured_at": "YYYY-MM-DD HH:MM veya boş",
    "temperature_c": null,
    "pulse_bpm": null,
    "spo2_pct": null,
    "systolic_bp": null,
    "diastolic_bp": null,
    "glucose_mg_dl": null
  }},
  "measurements": {{
    "measured_on": "YYYY-MM-DD veya boş",
    "height_cm": null,
    "weight_kg": null,
    "head_circ_cm": null
  }},
  "medications": [
    {{"medication_name": "İlaç Adı", "med_type": "Düzenli İlaç", "dosage": "1x1 / 5 ml", "schedule_times": "09:00, 13:00", "notes": "Kullanım notu"}}
  ],
  "disability": {{
    "has_disability": false,
    "disability_rate": null,
    "categories": [],
    "mobility_status": "",
    "school_adaptations": "",
    "report_summary": ""
  }},
  "diabetes": {{
    "is_diabetic": false,
    "diabetes_type": "",
    "icr_breakfast": null,
    "icr_lunch": null,
    "icr_snack": null,
    "isf": null,
    "target_bg": null,
    "insulin_brand": ""
  }},
  "visit_note_suggestion": "Belge acil başvuru/istirahat/rapor içeriyorsa kısa revir izlem notu (yoksa boş)",
  "inferred_risks_and_actions": [
    "Bu belgeden hemşirenin dikkat etmesi gereken çıkarım veya aksiyon 1"
  ]
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		inline_files=inline_files if inline_files else None,
		expect_json=True,
		temperature=0.1,
		timeout=32.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Belge analiz sonucu JSON olarak ayrıştırılamadı."}, used_m
	return True, parsed, used_m


def infer_insights_from_student_reports_with_gemini(
	*,
	api_key: str,
	model: str,
	anon_context: str,
	clinical_snapshot: dict[str, Any],
	missing_data_items: list[str],
	matched_depot_summaries: list[dict[str, str]] | None = None,
	focus_question: str = "",
) -> tuple[bool, dict[str, Any], str]:
	"""
	Öğrencinin veritabanındaki tüm geçmiş raporlarını (vizit, tahlil, vital, büyüme, ilaç, tarama, engel, diyabet)
	ve varsa Evrak Deposundaki belgelerini birlikte değerlendirip gizli klinik örüntüleri, eksikleri ve
	profile yansıtılabilecek önerileri çıkarır. (KVKK tam anonim)
	"""
	depot_txt = json.dumps(matched_depot_summaries or [], ensure_ascii=False)[:4000]
	prompt = f"""Bir pediatri ve okul sağlığı uzmanı olarak aşağıdaki anonim öğrencinin TÜM geçmiş kayıtlarını, tahlillerini, başvurularını ve ilişkili evrak metinlerini çapraz analiz ederek derin klinik çıkarımlar üret.

[ANONİM ÖĞRENCİ BAĞLAMI]: {anon_context}
[KLİNİK VERİ ÖZETİ]: {json.dumps(clinical_snapshot, ensure_ascii=False, default=str)[:5500]}
[EKSİK VERİ LİSTESİ]: {json.dumps(missing_data_items, ensure_ascii=False)}
[EVRAK DEPOSU İLİŞKİLİ BELGE ÖZETLERİ]: {depot_txt}
[ÖZEL ODAK SORUSU]: {focus_question or 'Tekrarlayan semptom örüntülerini, kayıtlar arası çelişkileri/eksikleri, okul içi riskleri ve profile eklenmesi gerekenleri çıkar.'}

SADECE aşağıdaki JSON şemasını döndür (kesin tıbbi tanı koyma, hemşire karar desteği ve yönlendirme dili kullan):
{{
  "executive_summary": "Öğrencinin genel sağlık tablosunun 2-3 cümlelik bütüncül özeti",
  "hidden_patterns": [
    "Kayıtlar arası tespit edilen tekrarlayan örüntü veya ilişkili bulgu 1",
    "Bulgu 2"
  ],
  "inconsistencies_or_gaps": [
    "Bir kayıtta geçen ama profilde eksik kalmış bulgu veya tamamlanması gereken veri 1"
  ],
  "suggested_profile_updates": {{
    "allergies_add": "Profile eklenmesi önerilen alerji (yoksa boş)",
    "chronic_add": "Profile eklenmesi önerilen kronik durum/izlem (yoksa boş)",
    "emergency_notes_add": "Acil notlara eklenmesi önerilen okul önlemi (yoksa boş)",
    "care_plan_title": "Önerilen kısa bakım planı başlığı (gerekmiyorsa boş)",
    "care_plan_interventions": "Önerilen okul hemşireliği girişimleri (gerekmiyorsa boş)"
  }},
  "follow_up_recommendations": [
    "Öncelikli hemşire/veli/hekim takip adımı 1",
    "Adım 2"
  ],
  "report_ready_paragraph": "Öğrencinin resmi Word/PDF sağlık raporundaki 'AI Değerlendirme' bölümüne doğrudan eklenebilecek profesyonel, akıcı Türkçe paragraf"
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		temperature=0.22,
		timeout=28.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Klinik çıkarım yanıtı JSON olarak ayrıştırılamadı."}, used_m
	return True, parsed, used_m


def scan_pdf_form_fields_with_gemini(
	*,
	api_key: str,
	model: str,
	page_png_bytes: bytes,
	page_number: int = 1,
	school_context: str = "",
	anon_student_context: str = "",
	user_instruction: str = "",
) -> tuple[bool, dict[str, Any], str]:
	"""
	PDF sayfasının görüntüsünü (taranmış veya statik formlar dahil) Gemini Vision ile tarar,
	yazı yazılabilecek boş alanların etiketlerini, yüzde cinsinden konumlarını (x_pct, y_pct, w_pct, h_pct)
	ve varsa bağlama göre önerilen doldurma değerlerini döndürür.
	"""
	prompt = f"""Bu görüntü bir okul reviri / resmi form PDF sayfasıdır (Sayfa {page_number}).
Sayfada el ile veya bilgisayarla YAZI YAZILABİLECEK veya İŞARETLENEBİLECEK TÜM boş form alanlarını (isim, sınıf, tarih, okul adı, açıklama satırları, tablo boş kutuları, onay kutuları, imza üstü ad-soyad vb.) tespit et.

[OKUL BİLGİSİ]: {school_context or 'Okul Reviri'}
[AKTİF KİŞİ ANONİM BAĞLAMI]: {anon_student_context or 'Yok'}
[HEMŞİRE TALİMATI]: {user_instruction or 'Sayfadaki tüm yazılabilir boş alanları koordinatlarıyla bul ve uygun olanlara öneri metin yaz.'}

SADECE aşağıdaki JSON yapısını döndür:
{{
  "fields": [
    {{
      "label": "Alan Etiketi (örn: Adı Soyadı / Tarih / Okul Adı / Açıklama)",
      "field_type": "text | date | checkbox",
      "x_pct": 25.0,
      "y_pct": 18.5,
      "w_pct": 40.0,
      "h_pct": 3.5,
      "suggested_value": "Varsa otomatik doldurulabilecek öneri (yoksa boş string)",
      "font_size": 10
    }}
  ],
  "notes": "Sayfa hakkında kısa Türkçe not"
}}

KOORDİNAT KURALLARI:
- x_pct (0-100): Yazının başlayacağı SOL kenar (etiketin hemen sağı veya boş çizginin başı).
- y_pct (0-100): Yazı kutusunun ÜST kenarı (0=sayfa en üstü, 100=sayfa en altı).
- w_pct (5-90): Boş alanın genişliği yüzde cinsinden.
- h_pct (2-25): Boş alanın yüksekliği yüzde cinsinden.
- En fazla 35 en önemli yazılabilir alanı listele; üst üste binen kopya alan üretme."""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		inline_files=[(page_png_bytes, "image/png")],
		expect_json=True,
		temperature=0.1,
		timeout=32.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "PDF alan tarama yanıtı JSON olarak ayrıştırılamadı."}, used_m
	fields_clean = []
	for i, f in enumerate(parsed.get("fields") or []):
		if not isinstance(f, dict):
			continue
		try:
			xp = max(1.0, min(95.0, float(f.get("x_pct") or 10.0)))
			yp = max(1.0, min(95.0, float(f.get("y_pct") or 10.0)))
			wp = max(4.0, min(96.0 - xp, float(f.get("w_pct") or 28.0)))
			hp = max(2.0, min(40.0, float(f.get("h_pct") or 3.2)))
		except Exception:
			continue
		fields_clean.append({
			"label": str(f.get("label") or f"Alan {i + 1}").strip(),
			"field_type": str(f.get("field_type") or "text").strip(),
			"x_pct": round(xp, 2),
			"y_pct": round(yp, 2),
			"w_pct": round(wp, 2),
			"h_pct": round(hp, 2),
			"suggested_value": str(f.get("suggested_value") or "").strip(),
			"font_size": int(max(7, min(20, float(f.get("font_size") or 10)))),
		})
	return True, {"fields": fields_clean, "notes": str(parsed.get("notes") or "")}, used_m


def assist_office_document_with_gemini(
	*,
	api_key: str,
	model: str,
	doc_type: str,
	doc_payload: dict[str, Any],
	instruction: str,
	school_context: str = "",
	anon_student_context: str = "",
) -> tuple[bool, dict[str, Any], str]:
	"""
	Evrak Deposundaki Word (.docx), Excel (.xlsx) veya PDF form alanlarında hemşirenin talimatına göre
	veri girme, metin iyileştirme veya tablo düzenleme önerileri üretir.
	doc_type: 'docx' | 'xlsx' | 'pdf_fields'
	"""
	payload_json = json.dumps(doc_payload, ensure_ascii=False, default=str)[:12000]
	if doc_type == "docx":
		schema_hint = """{
  "explanation": "Yapılan/önerilen değişikliklerin kısa Türkçe açıklaması",
  "paragraph_edits": [{"index": 0, "new_text": "Güncellenmiş paragraf metni"}],
  "append_paragraphs": ["Belge sonuna eklenecek yeni paragraf (gerekirse)"],
  "table_edits": [{"table_index": 0, "row_index": 1, "col_index": 1, "new_text": "Yeni hücre değeri"}]
}"""
	elif doc_type == "xlsx":
		schema_hint = """{
  "explanation": "Yapılan/önerilen değişikliklerin kısa Türkçe açıklaması",
  "cell_edits": [{"row_index": 0, "col_name": "Sütun Adı", "value": "Yeni Değer"}],
  "append_rows": [{"Sütun1": "Değer1", "Sütun2": "Değer2"}]
}"""
	else:
		schema_hint = """{
  "explanation": "PDF alanları için önerilen doldurmaların kısa açıklaması",
  "field_values": {"f_1": "Doldurulacak metin", "f_2": "Değer"}
}"""

	prompt = f"""Okul revir hemşiresinin düzenlediği aşağıdaki `{doc_type}` belgesinde, verilen talimata göre yalnızca gerekli alanları profesyonel ve doğru Türkçe ile düzenle veya doldur.

[OKUL BİLGİSİ]: {school_context or 'Okul Reviri'}
[AKTİF KİŞİ ANONİM BAĞLAMI]: {anon_student_context or 'Seçili kişi yok'}
[HEMŞİRENİN DÜZENLEME TALİMATI]: {instruction}
[MEVCUT BELGE İÇERİĞİ]:
{payload_json}

SADECE aşağıdaki JSON yapısını döndür (değişmeyecek satır/paragrafları tekrarlama, yalnızca değişecek veya eklenecek öğeleri yaz):
{schema_hint}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		temperature=0.18,
		timeout=30.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Belge yardımcı yanıtı JSON olarak ayrıştırılamadı."}, used_m
	return True, parsed, used_m


def generate_calendar_events_with_ai(
	*,
	api_key: str,
	model: str,
	year: int,
	month: int | None = None,
	preset_type: str = "custom",
	custom_prompt: str = "",
	selected_exams: list[str] | None = None,
) -> tuple[bool, dict[str, Any], str]:
	"""
	Gemini AI kullanarak resmi tatilleri, yarım gün arife/28 Ekim çalışmalarını,
	dini bayram ve kandilleri, seçim/idari izin tatillerini, ÖSYM/MEB sınavlarını,
	belirli gün ve haftaları veya meslek günlerini yapılandırılmış takvim kayıtları olarak üretir.
	"""
	period_txt = f"{year} yılı {month}. ayı" if month else f"{year} yılı tamamı"
	exams_txt = ", ".join(selected_exams) if selected_exams else "YKS (TYT-AYT-YDT), LGS, KPSS, MSÜ, ALES, DGS, YDS, İOKBS Bursluluk"

	preset_instructions = {
		"official_holidays": (
			f"{period_txt} için Türkiye Cumhuriyeti resmi tatillerini (1 Ocak, 23 Nisan, 1 Mayıs, 19 Mayıs, "
			"15 Temmuz, 30 Ağustos, 29 Ekim), devlet memurları için yarım gün mesai olan günleri (28 Ekim öğleden sonra, "
			"Ramazan Bayramı Arifesi, Kurban Bayramı Arifesi), Ramazan ve Kurban Bayramı günlerini ve önemli kandil/dini günleri "
			"eksiksiz çıkar. Milli bayramlarda symbol_type='tr_flag', dini bayram ve kandillerde symbol_type='crescent' kullan."
		),
		"osym_exams": (
			f"{period_txt} içinde gerçekleşen veya planlanan ÖSYM ve MEB merkezi sınavlarını ({exams_txt}) "
			"tarihleri, sınav saatleri (ör. 10:15) ve kısa açıklamalarıyla çıkar. symbol_type='exam', category='osym_sinav' kullan."
		),
		"special_days": (
			f"{period_txt} içindeki MEB Belirli Gün ve Haftaları, sağlık takvimi günlerini (Tıp Bayramı, Hemşireler Günü, "
			"Eczacılar Günü, Diş Hekimleri Günü, Dünya Diyabet Günü, Yeşilay Haftası, Otizm Farkındalık Günü vb.) ve "
			"meslek grubu günlerini (Öğretmenler Günü, Avukatlar Günü, Gazeteciler Günü, Mühendisler Günü, Hemşireler Günü vb.) çıkar."
		),
		"meb_belirli_gunler": (
			f"{period_txt} için T.C. Millî Eğitim Bakanlığı (MEB) Eğitim Kurumları Sosyal Etkinlikler Yönetmeliği'nde yer alan "
			"resmi 'MEB Belirli Gün ve Haftalar Çizelgesi'ndeki tüm belirli gün ve haftaları (İlköğretim Haftası, Gaziler Günü, "
			"Hayvanları Koruma Günü, Kızılay Haftası, Lösemili Çocuklar Haftası, Organ Bağışı Haftası, Atatürk Haftası, "
			"Dünya Diyabet Günü, Çocuk Hakları Günü, Ağız ve Diş Sağlığı Haftası, Öğretmenler Günü, Dünya Engelliler Günü, "
			"İnsan Hakları ve Demokrasi Haftası, Tutum Yatırım ve Türk Malları / Yerli Malı Haftası, Enerji Tasarrufu Haftası, "
			"Verem Eğitimi ve Propaganda Haftası, Sivil Savunma Günü, Yeşilay Haftası, Bilim ve Teknoloji Haftası, "
			"İstiklâl Marşı'nın Kabulü ve Mehmet Akif Ersoy'u Anma Günü, Tüketiciyi Koruma Haftası, Şehitler Günü, "
			"Yaşlılara Saygı Haftası, Orman Haftası, Kütüphaneler Haftası, Kanser Haftası, Dünya Otizm Farkındalık Günü, "
			"Sağlık Haftası, Turizm Haftası, Trafik ve İlkyardım Haftası, Bilişim Haftası, Vakıflar Haftası, Engelliler Haftası, "
			"Hemşirelik Haftası, Müzeler Haftası, Etik Günü, İstanbul'un Fethi, Çevre Koruma Haftası vb.) tarihleriyle çıkar. "
			"category='belirli_gun', symbol_type='event' (milli olanlarda 'tr_flag', sağlık olanlarda 'health') kullan."
		),
		"election_or_custom": (
			f"{period_txt} bağlamında kullanıcının şu özel takvim talimatını yerine getir: {custom_prompt or 'Okul ve revir takvim planı oluştur'}"
		),
	}
	task_desc = preset_instructions.get(preset_type, preset_instructions["election_or_custom"])
	if custom_prompt and preset_type != "election_or_custom":
		task_desc += f"\nEk kullanıcı isteği: {custom_prompt}"

	prompt = f"""Sen Türkiye'deki devlet okulları, kamu kurumları ve okul revirleri için çalışan uzman bir Resmi Takvim ve Planlama Asistanısın.
Hedef Dönem: {period_txt}
Görev: {task_desc}

Kurallar:
1. Tarihleri kesinlikle "YYYY-MM-DD" formatında (`event_date`) yaz.
2. Devlet memurları için resmi tatil olan günlerde `is_civil_holiday: true` yap (kırmızı tatil günü). Kandil geceleri veya kutlama haftaları resmi tatil değildir (`is_civil_holiday: false`).
3. 28 Ekim ve Bayram Arifeleri gibi yarım gün mesailerde `is_half_day: true` ve `is_civil_holiday: true` yap.
4. `symbol_type` alanı yalnızca şu değerlerden biri olabilir:
   - "tr_flag" (29 Ekim, 23 Nisan, 19 Mayıs, 30 Ağustos, 15 Temmuz, 10 Kasım, 18 Mart gibi Milli Bayramlar ve Ulusal Günler)
   - "crescent" (Ramazan Bayramı, Kurban Bayramı, Arife, Kandiller, Kadir Gecesi, Aşure Günü gibi Dini Günler)
   - "half_day" (Yarım gün mesai)
   - "exam" (ÖSYM ve MEB Sınavları)
   - "profession" (Meslek grubu günleri: Tıp Bayramı, Hemşireler Günü, Öğretmenler Günü vb.)
   - "ballot" (Seçim tatili / idari izin)
   - "health" (Sağlık haftası / revir tarama / aşı günü)
   - "event" (Diğer belirli gün, hafta veya etkinlik)
5. `category` alanı yalnızca şu değerlerden biri olabilir:
   "milli_bayram", "dini_bayram", "resmi_tatil", "yarim_gun", "secim_idari_tatil", "osym_sinav", "meslek_gunu", "belirli_gun", "revir_etkinlik"

SADECE aşağıdaki JSON yapısını döndür:
{{
  "summary": "Üretilen takvim kayıtlarının kısa Türkçe özeti",
  "events": [
    {{
      "event_date": "{year}-10-29",
      "title": "29 Ekim Cumhuriyet Bayramı",
      "category": "milli_bayram",
      "symbol_type": "tr_flag",
      "is_civil_holiday": true,
      "is_half_day": false,
      "profession_group": "",
      "start_time": "09:00",
      "end_time": "17:00",
      "description": "Cumhuriyetin ilanı - Devlet memurları ve okullar için tam gün resmi tatil."
    }}
  ]
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		temperature=0.2,
		timeout=35.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed or not isinstance(parsed.get("events"), list):
		return False, {"error": "Gemini takvim yanıtı JSON olarak ayrıştırılamadı."}, used_m
	return True, parsed, used_m


def generate_daily_leaf_content_with_ai(
	*,
	api_key: str,
	model: str,
	target_date: str,
	quote_mode: str = "Karışık (Hadis + Özlü Söz)",
	day_context: str = "",
	province: str = "Ankara",
) -> tuple[bool, dict[str, Any], str]:
	"""
	Seçili gün için klasik Türk 'Takvim Yaprağı' içeriği üretir:
	- O günün hangi meslek grubunun günü olduğu / hangi tatil veya belirli gün olduğu
	- Ayarlarda seçilen tercihe göre sahih kaynaklı Hadis-i Şerif, Özlü Söz veya Karışık Takvim Yaprağı yazısı.
	"""
	prompt = f"""Sen geleneksel ve modern Türk 'Takvim Yaprağı' kültürünü çok iyi bilen bir asistansın.
Seçili Tarih: {target_date}
İl: {province}
Günün Bilinen Tatil / Özel Gün Bağlamı: {day_context or 'Standart gün'}
Kullanıcının Söz/Hadis Gösterim Tercihi: {quote_mode}

Lütfen bu tarihe özel olarak şunları üret:
1. O gün Türkiye'de veya dünyada hangi meslek grubunun günü (ör. Hemşireler, Hekimler, Öğretmenler, Eczacılar, Mühendisler, İtfaiyeciler, Çiftçiler, Madenciler vb. veya o güne en yakın meslek/topluluk) ve hangi tatil/belirli gün olduğu.
2. Sahih kaynaklı (Buhârî, Müslim, Tirmizî, Ebû Dâvûd, Riyâzü's-Sâlihîn vb.) ahlak, sağlık, merhamet, ilim, sabır veya yardımlaşma temalı bir Hadis-i Şerif ve kaynağı.
3. Günün anlam ve önemine uygun bir Özlü Söz / Atasözü ve söyleyeni.
4. Klasik takvim yaprağı arkası tadında 2-3 cümlelik "Günün Takvim Yaprağı Notu".

SADECE aşağıdaki JSON yapısını döndür:
{{
  "day_significance": "Günün anlam ve önemi / tatil ve belirli gün özeti",
  "profession_group": "O günün meslek grubu veya 'Tüm Kamu & Sağlık Çalışanları'",
  "profession_note": "Meslek grubuna veya günün çalışanlarına özel kısa tebrik/anlam notu",
  "hadith_text": "Hadis-i Şerif metni",
  "hadith_source": "Kaynak (Örn: Buhârî, İlim, 10)",
  "quote_text": "Özlü söz veya atasözü metni",
  "quote_author": "Söyleyen / Kaynak",
  "leaf_back_note": "Takvim yaprağı arkası 2-3 cümlelik faydalı/hikmetli günün yazısı"
}}"""

	ok, raw, used_m = call_gemini(
		prompt,
		api_key=api_key,
		model=model,
		expect_json=True,
		temperature=0.35,
		timeout=25.0,
	)
	if not ok:
		return False, {"error": raw}, used_m
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Takvim yaprağı yanıtı JSON olarak ayrıştırılamadı."}, used_m
	return True, parsed, used_m




