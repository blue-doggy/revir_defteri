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
	"gemini-3.1-flash-lite",
	"gemini-3-flash-preview",
	"gemini-3.8-flash",
	"gemini-3.1-pro-preview",
]

GEMINI_MODEL_LABELS: dict[str, str] = {
	"gemini-3.1-flash-lite": "⚡ Gemini 3.1 Flash Lite (Önerilen - Ultra Hızlı ~1.5sn & Yüksek Kota)",
	"gemini-3-flash-preview": "🚀 Gemini 3 Flash (Hızlı Klinik & Görsel Analiz)",
	"gemini-3.8-flash": "🧠 Gemini 3.8 Flash (Yeni Nesil Klinik Akıl Yürütme)",
	"gemini-3.1-pro-preview": "🔬 Gemini 3.1 Pro (Derin Klinik & Rapor Analizi)",
}

_FALLBACK_MODELS: list[str] = [
	"gemini-3.1-flash-lite",
	"gemini-3-flash-preview",
	"gemini-3.8-flash",
	"gemini-3.1-pro-preview",
]


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


_PHONE_RE = re.compile(r"(?:\+?90|0)?\s*5\d{2}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}")
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

	if health_profile:
		hp = dict(health_profile)
		if hp.get("blood_group"):
			parts.append(f"Kan Grubu: {hp['blood_group']}")
		if hp.get("allergies"):
			parts.append(f"Kayıtlı Alerjiler: {anonymize_free_text(str(hp['allergies']), sensitive)}")
		else:
			parts.append("Kayıtlı Alerjiler: Bilinen alerji yok")
		if hp.get("chronic_conditions"):
			parts.append(f"Kronik Rahatsızlıklar: {anonymize_free_text(str(hp['chronic_conditions']), sensitive)}")
		if hp.get("regular_medications"):
			parts.append(f"Sürekli İlaçlar: {anonymize_free_text(str(hp['regular_medications']), sensitive)}")
		if hp.get("emergency_notes"):
			parts.append(f"Acil Klinik Not: {anonymize_free_text(str(hp['emergency_notes']), sensitive)}")
		if hp.get("special_needs"):
			parts.append(f"Özel Gereksinim/Cihaz: {anonymize_free_text(str(hp['special_needs']), sensitive)}")

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

	preferred = (model or "gemini-3-flash-preview").strip().replace("models/", "")
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
	return True, parsed, used_model


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
	return True, {k: str(v or "") for k, v in parsed.items()}, used_model


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
) -> tuple[bool, dict[str, str], str]:
	cols = [c for c in ["Tarih", "Yas_ay", "Boy_cm", "Boy_Z", "Boy_Persantil", "Kilo_kg", "Kilo_Z", "Kilo_Persantil", "BMI", "BMI_Z", "BMI_Persantil", "Gelisim_Degerlendirmesi"] if c in measurements_df.columns]
	history_records = measurements_df[cols].head(12).to_dict(orient="records") if not measurements_df.empty else []

	prompt = f"""Dünya Sağlık Örgütü (WHO) pediatrik büyüme standartlarına göre aşağıdaki anonim öğrencinin kronolojik boy, kilo ve VKİ persantil/Z-skoru geçmişini değerlendir:

[ANONİM PROFİL]: {anon_context}
[KRONOLOJİK ÖLÇÜM GEÇMİŞİ (En yeniden eskiye)]:
{json.dumps(history_records, ensure_ascii=False, default=str)}

Lütfen yanıtı aşağıdaki JSON anahtarlarıyla Türkçe olarak döndür:
{{
  "channel_status": "Normal Kanalda Stabil | Persantil Kayması / İzlem Önerilir | Kritik Sınır (<3p veya >97p - Hekim Değerlendirmesi Önerilir)",
  "clinical_summary": "Hemşire için büyüme hızı, boy-kilo uyumu, Z-skoru değişimi ve varsa persantil kanal kaymasının 2-3 cümlelik teknik analizi",
  "recommendations": "Okul revirinde izlem sıklığı, beslenme/fiziksel aktivite veya pediatri/aile hekimi yönlendirme önerisi",
  "guardian_letter": "Veliye gönderilecek resmi, anlaşılır ve yapıcı büyüme-gelişme bilgilendirme notu ([ÖĞRENCİ_ADI] yer tutucusunu kullan, Kurum: {school_name}, İmza: {nurse_name})"
}}"""

	ok, raw, used_model = call_gemini(prompt, api_key=api_key, model=model, expect_json=True)
	if not ok:
		return False, {"error": raw}, used_model
	parsed = _parse_json_response(raw)
	if not parsed:
		return False, {"error": "Büyüme analizi JSON yanıtı ayrıştırılamadı."}, used_model
	return True, {k: str(v or "") for k, v in parsed.items()}, used_model


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
	return True, {k: str(v or "") for k, v in parsed.items()}, used_model


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
) -> tuple[bool, str, str]:
	safe_stats = {
		k: v for k, v in stats.items()
		if k in (
			"students_total", "civilians_total", "classes_total", "visits_total",
			"student_visits", "civilian_visits", "guardian_contacted_cnt", "referral_cnt",
			"screened_students", "screening_coverage_pct", "meas_total", "bmi_eval_total",
			"bmi_normal_cnt", "bmi_under_cnt", "bmi_over_cnt", "bmi_obese_cnt",
			"stunted_cnt", "chronic_or_allergy_cnt", "care_plans_total", "care_plans_active",
			"active_meds_cnt", "doses_given_cnt", "isolation_total", "top_complaints",
		)
	}
	prompt = f"""T.C. Millî Eğitim Bakanlığı (MEB) standartlarına uygun olarak '{school_name}' revirinin {start_date_str} - {end_date_str} dönemi resmi faaliyet raporu için 'Dönem İçi Yapılan Sağlık Hizmetleri ve Genel Değerlendirme Metni' hazırla.

[DÖNEM İSTATİSTİKLERİ]:
{json.dumps(safe_stats, ensure_ascii=False, default=str)}

Kurumsal, resmi ve akıcı bir Türkçe ile (yaklaşık 2-3 paragraf halinde):
- Poliklinik/revir başvuru ve triyaj hizmetlerini,
- Periyodik okul taramaları ve DSÖ (WHO) büyüme-gelişme/persantil izlemlerini,
- Kronik hastalığı olan öğrencilerin NANDA bakım planı ve ilaç takip süreçlerini,
- Bulaşıcı hastalık/salgın kontrolü ve sağlık eğitimi önerilerini
özetleyen resmi rapor değerlendirme metnini doğrudan yaz (başlık veya markdown kod bloğu olmadan doğrudan rapora basılacak resmi metin olarak döndür)."""

	return call_gemini(prompt, api_key=api_key, model=model)


# ==============================================================================
# 9. SENARYO 7: DOĞAL DİLLE REVİR VERİTABANI SORGULAMA (SALT-OKUNUR SQL)
# ==============================================================================

_DB_SCHEMA_HINT = """
Tablolar:
- classes(id, name, teacher_name)
- students(id, person_type, full_name, class_id, sex, birth_date, student_no, guardian_name, guardian_phone, notes, province, district, village)
- health_profiles(student_id, blood_group, allergies, chronic_conditions, regular_medications, emergency_notes, special_needs)
- visits(id, student_id, visit_date, visit_time, complaint, symptoms, observation, action, guardian_contacted, referral, follow_up, temperature_c, spo2_pct, pulse_bpm, systolic_mmhg, diastolic_mmhg, incident_location)
- measurements(id, student_id, measured_on, height_cm, weight_kg, head_circ_cm, age_months, bmi)
- school_screenings(id, student_id, screening_date, height_cm, weight_kg, wears_glasses, vision_right, vision_left, hearing_right, hearing_left, hair_screening, health_behavior, psychosocial_status, referral_needed, notes)
- student_medications(id, student_id, medication_name, med_type, route, dosage_pattern, status, remaining_volume, expiry_date)
- care_plans(id, student_id, nursing_diagnosis, observed_facts, goals, interventions, status, updated_at)
- isolation_records(id, student_id, condition_name, detected_date, expected_return_date, status)
- infirmary_inventory(id, item_name, category, quantity, unit, min_threshold, expiry_date)
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
4. En fazla `LIMIT 100` satır döndür.

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
	voice_audio_mime: str = "audio/webm",
	anon_person_context: str = "",
	school_context: str = "",
	sensitive_names: list[str] | None = None,
) -> tuple[bool, str, str, str]:
	"""
	Yüzen sohbet balonundan gelen metin, çoklu dosya eki ve/veya ses kaydını yanıtlar.
	Dönüş: (başarılı_mı, yanıt_metni, kullanılan_model, ses_kaydı_varsa_çözümlenen_kullanıcı_metni)
	"""
	sys_inst = (
		"Sen 'Revir Defteri' masaüstü okul sağlığı yazılımının içinde yer alan KVKK korumalı 'Gemini Klinik Asistan'sın. "
		"Okul revir hemşiresine pediatrik triyaj, büyüme-gelişme (WHO persantil/Z-skoru), NANDA bakım planları, "
		"ilaç etkileşimleri, tahlil/evrak yorumlama ve okul sağlığı yönetimi konularında hızlı, doğru, Türkçe ve öz yanıtlar verirsin. "
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
				inline_files.append((vbytes, voice_audio_mime or "audio/webm"))
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
		ctx_parts.append(f"[OKUL & REVİR DURUMU]: {school_context}")
	if anon_person_context:
		ctx_parts.append(f"[SEÇİLİ KİŞİ ANONİM KLİNİK PROFİLİ]: {anon_person_context}")
	if hist_lines:
		hist_joined = "\n".join(hist_lines)
		ctx_parts.append(f"[ÖNCEKİ SOHBET GEÇMİŞİ]:\n{hist_joined}")
	if extra_doc_blocks:
		docs_joined = "\n\n".join(extra_doc_blocks)
		ctx_parts.append(docs_joined)

	if has_voice and not safe_user_msg.strip():
		ctx_parts.append(
			"[KULLANICI SESLİ MESAJ GÖNDERDİ]: Ekli ses kaydını dinle. "
			"Yanıtının ilk satırına kullanıcının ne söylediğini `[SES_ÇÖZÜMÜ]: <söylenen cümle>` biçiminde yaz, "
			"alt satırda ise sorusuna Türkçe klinik yanıtını ver."
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
		timeout=30.0,
	)
	if not ok:
		return False, raw_reply, used_model, ""

	transcribed = ""
	clean_reply = raw_reply
	m_tr = re.match(r"^\s*\[SES_ÇÖZÜMÜ\]\s*:\s*(.+?)(?:\n+|$)(.*)", raw_reply, re.S)
	if m_tr:
		transcribed = m_tr.group(1).strip()
		clean_reply = m_tr.group(2).strip() or transcribed

	return True, clean_reply, used_model, transcribed

