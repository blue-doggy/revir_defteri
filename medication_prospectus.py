# -*- coding: utf-8 -*-
"""
Okul Revir Defteri ve Sağlık Takip Sistemi - İlaç Prospektüsü,
Türkiye Piyasası Genel İlaç Veri Bankası ve Yapay Zekâ Destekli Sorgulama Modülü.
"""
from __future__ import annotations

import html
import re
import sqlite3
from datetime import datetime
from io import BytesIO
from typing import Any, Callable

import revir_ui as ui
import gemini_assistant as gai


PROSPECTUS_DISCLAIMER = (
	"⚠️ Bu prospektüs özeti okul reviri hızlı klinik referansı ve bilgilendirme amaçlıdır. "
	"Resmî doz, kontrendikasyon ve tedavi kararlarında TİTCK Onaylı Kısa Ürün Bilgisi (KÜB/KT) "
	"ve hekim reçetesi esas alınmalıdır."
)

ROUTE_OPTIONS = [
	"Ağızdan (Oral)",
	"Solunum (İnhalasyon / Hava)",
	"Göz Damlası",
	"Kulak Damlası",
	"Burun Spreyi",
	"Cilt / Dermal (Merhem/Krem)",
	"Enjeksiyon (Deri Altı/İnsülin)",
	"Rektal",
	"Diğer",
]

MEAL_OPTIONS = [
	"Tok karnına",
	"Aç karnına",
	"Yemekle birlikte",
	"Yemekten 30 dk önce",
	"Fark etmez",
]

STORAGE_OPTIONS = [
	"Oda Sıcaklığı (15-25 °C)",
	"Buzdolabı (+2 / +8 °C)",
	"Kilitli Revir Dolabı",
	"Işıktan Korunan Dolap",
]

DOSAGE_FORM_OPTIONS = [
	"Şurup / Süspansiyon",
	"Tablet / Kapsül",
	"Efervesan / Saşe",
	"İnhaler / Nebül",
	"Göz / Kulak Damlası",
	"Burun Spreyi / Damla",
	"Krem / Merhem / Jel",
	"Enjeksiyon / Kalem / Oto-Enjektör",
	"Boğaz Spreyi / Gargara",
	"Rektal Tüp / Fitil",
	"Oral Damla / Solüsyon",
	"Diğer",
]


def _fold_med_text(text: Any) -> str:
	"""Türkçe karakterleri sadeleştirerek ilaç eşleştirmesi için normalize metin üretir."""
	raw = str(text or "").strip()
	if not raw:
		return ""
	tr_map = str.maketrans({
		"İ": "i", "I": "i", "ı": "i",
		"Ğ": "g", "ğ": "g",
		"Ü": "u", "ü": "u",
		"Ş": "s", "ş": "s",
		"Ö": "o", "ö": "o",
		"Ç": "c", "ç": "c",
		"Â": "a", "â": "a",
		"Î": "i", "î": "i",
	})
	s = raw.translate(tr_map).lower()
	s = re.sub(r"[^a-z0-9]+", " ", s).strip()
	return re.sub(r"\s+", " ", s)


_MED_STOPWORDS = {
	"mg", "mcg", "ug", "ml", "gr", "g", "iu", "bid", "fort", "forte", "plus", "ped",
	"pediatrik", "cocuk", "yetiskin", "surup", "suspansiyon", "tablet", "film", "kapli",
	"kapsul", "sase", "damla", "sprey", "spreyi", "krem", "merhem", "jel", "solusyon",
	"cozelti", "inhaler", "nebul", "ampul", "flakon", "kalem", "oral", "rektal", "topikal",
	"toz", "efervesan", "tb", "cr", "sr", "xr", "zok", "junior", "baby", "inf", "ve", "ile",
}


def _extract_core_med_tokens(text: Any) -> set[str]:
	folded = _fold_med_text(text)
	tokens: set[str] = set()
	for tok in folded.split():
		if not tok or tok.isdigit():
			continue
		if re.match(r"^\d+(mg|ml|mcg|gr|g)$", tok):
			continue
		if tok in _MED_STOPWORDS or len(tok) < 3:
			continue
		tokens.add(tok)
	return tokens


def _make_norm_key(med_name: str, generic_name: str = "") -> str:
	core = _fold_med_text(med_name)
	if not core:
		core = _fold_med_text(generic_name)
	return core.replace(" ", "_")[:140] or "ilac"


def _d(
	name: str,
	generic: str,
	brands: str,
	t_class: str,
	form: str,
	route: str,
	meal: str,
	storage: str,
	purpose: str,
	admin_meal: str,
	max_dose: str,
	side_fx: str,
	contra: str,
	ped_note: str = "",
	full_notes: str = "",
) -> dict[str, str]:
	return {
		"medication_name": name,
		"generic_name": generic,
		"brand_examples": brands,
		"therapeutic_class": t_class,
		"dosage_form": form,
		"route_default": route,
		"meal_relation_default": meal,
		"storage_condition_default": storage,
		"general_purpose": purpose,
		"administration_and_meal": admin_meal,
		"max_daily_dose": max_dose,
		"side_effects": side_fx,
		"contraindicated_drugs": contra,
		"pediatric_notes": ped_note,
		"full_prospectus_notes": full_notes,
	}


# Türkiye ilaç piyasasında okul revirlerinde, pediatride ve yetişkin personelde en sık karşılaşılan 60 temel ilaç
BUILTIN_TR_MEDICATION_BANK: list[dict[str, str]] = [
	# --- 1. ATEŞ DÜŞÜRÜCÜ & AĞRI KESİCİLER ---
	_d(
		"Calpol / Parol (Parasetamol)",
		"Parasetamol (Asetaminofen)",
		"Calpol, Calpol 6 Plus, Parol, Tamol, Minoset, Vermidon",
		"Analjezik & Antipiretik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Hafif-orta şiddette ağrı (baş, diş, boğaz, kas ağrısı) ve ateşin düşürülmesi için ilk basamak ateş düşürücü ve ağrı kesicidir.",
		"Ağızdan (oral) uygulanır; aç veya tok karnına verilebilir (mide hassasiyetinde tok karnına tercih edilir). Şurup/süspansiyon formları her kullanım öncesi iyice çalkalanmalıdır.",
		"Çocuklarda tek seferlik doz 10-15 mg/kg (en az 4-6 saat ara ile), günlük maksimum 60 mg/kg/gün (24 saatte en fazla 4 doz). Yetişkin/12+ yaş (>50 kg) günlük maksimum 4000 mg (4 g).",
		"Nadiren hafif bulantı veya ciltte alerjik döküntü görülür. Yüksek doz aşımında ciddi karaciğer toksisitesi gelişebilir; sarılık veya şiddetli kusma olursa acil sevk gerekir.",
		"Başka parasetamol içeren grip/soğuk algınlığı ilaçlarıyla (A-ferin, Tylol Hot, Katarin, Theraflu vb.) birlikte verilmemelidir (gizli doz aşımı riski). Alkol, izoniazid ve varfarinle etkileşir.",
		"Okul çağında en güvenli birinci basamak ateş düşürücüdür; başka soğuk algınlığı şurubu alıp almadığı mutlaka sorgulanmalıdır.",
		"Vücut ağırlığına göre ölçek ayarı yapılmalı; veli yazılı onamı olmadan rutin dışı verilmemelidir.",
	),
	_d(
		"Dolven / Pedifen / İbufen (İbuprofen)",
		"İbuprofen",
		"Dolven, Pedifen, İbufen, Brufen, Nurofen, Apireks",
		"NSAİİ (Ateş Düşürücü & Antiinflamatuar)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Ateş düşürülmesi, diş/boğaz/kulak ağrısı, kas-eklem ağrıları ve yumuşak doku zedelenmelerine bağlı inflamasyonun giderilmesinde kullanılır.",
		"Mide tahrişini önlemek için mutlaka TOK KARNINA veya bir miktar süt/yemek sonrasında ağızdan verilir. Şişe çalkalanmalıdır.",
		"Çocuklarda 5-10 mg/kg/doz (6-8 saatte bir), günlük maksimum 30-40 mg/kg/gün. Yetişkin günlük maksimum 2400 mg (reçetesiz 1200 mg).",
		"Mide ağrısı, yanma, bulantı; nadir olarak bronkospazm (aspirin/NSAİİ duyarlı astımlılarda), böbrek yüklenmesi veya kurdeşen görülebilir.",
		"Diğer NSAİİ ağrı kesiciler (naproksen, diklofenak, flurbiprofen, aspirin) ve kortikosteroidlerle birlikte kullanılmamalıdır. Lityum, metotreksat, antikoagülanlar (kan sulandırıcılar) ile etkileşir; suçiçeği şüphesinde ve aktif mide ülserinde kaçınılmalıdır.",
		"Dehidrate (kusma/ishal olan) öğrencide böbrek riski nedeniyle önce sıvı durumu değerlendirilmelidir.",
	),
	_d(
		"Arveles (Deksketoprofen)",
		"Deksketoprofen Trometamol",
		"Arveles 25 mg, Dexofen, Dexday, Ketavel",
		"NSAİİ (Kuvvetli Ağrı Kesici)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkinlerde ve 18 yaş üstünde akut kas-iskelet ağrıları, adet sancısı, migren/baş ağrısı ve diş ağrısının kısa süreli semptomatik tedavisinde kullanılır.",
		"Ağızdan bol su ile alınır; hızlı etki isteniyorsa yemekten 30 dk önce alınabilir ancak mide koruması açısından çoğunlukla TOK KARNINA önerilir.",
		"Yetişkinlerde standart doz 8 saatte bir 25 mg; günlük maksimum 75 mg/gün. 18 yaş altı çocuk/ergenlerde kullanımı onaylı değildir.",
		"Mide yanması, dispepsi, bulantı, baş dönmesi; nadiren mide kanaması, cilt döküntüsü veya nefes darlığı.",
		"Başka NSAİİ'ler (ibuprofen, naproksen, diklofenak), yüksek doz aspirin, varfarin/heparin, lityum ve metotreksat ile birlikte kullanılmamalıdır. Mide ülseri ve NSAİİ astımında kontrendikedir.",
		"Öğrencilere değil, yalnızca 18 yaş üstü yetişkin personel için geçerlidir.",
	),
	_d(
		"Majezik (Flurbiprofen)",
		"Flurbiprofen",
		"Majezik 100 mg, Maxaljin, Fortine, Majezik Gargara/Sprey",
		"NSAİİ (Ağrı Kesici & Boğaz İltihabı)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin/12+ yaşta diş ağrısı, kas-eklem ağrısı, dismenore ve sprey formunda boğaz ağrısı/farenjit semptomlarını hafifletmek için kullanılır.",
		"Tablet formu TOK KARNINA bir bardak su ile ağızdan yutulur. Sprey formu doğrudan boğaza uygulanır, yutulmaz.",
		"Tablet: Yetişkinlerde günlük maksimum 300 mg (genellikle 2-3x1 100 mg). Sprey: Günde en fazla 3-5 kez (3 sprey/doz). 12 yaş altında tablet formu kullanılmaz.",
		"Mide ekşimesi, hazımsızlık, karın ağrısı, ağızda geçici yanma/uyuşma (sprey), nadir olarak alerjik ödem.",
		"Diğer NSAİİ'ler, aspirin, kan sulandırıcılar, ACE inhibitörü/diüretik tansiyon ilaçları ve lityum ile birlikte kullanılmamalıdır.",
	),
	_d(
		"Apranax / Naprosyn (Naproksen)",
		"Naproksen Sodyum",
		"Apranax, Apranax Fort, Naprosyn, Aprol",
		"NSAİİ (Ağrı Kesici & Antiinflamatuar)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Orta-şiddetli baş/diş/eklem ağrıları, burkulma travmaları ve adet ağrılarında uzun etki süreli ağrı kesici olarak kullanılır.",
		"Mide mukozasını korumak için mutlaka TOK KARNINA bol su ile ağızdan alınır.",
		"Yetişkinlerde günlük maksimum 1100-1375 mg (Apranax Fort 550 mg günde en fazla 2 tablet). Çocuklarda rutin kullanılmaz.",
		"Mide ağrısı, bulantı, baş dönmesi, kulak çınlaması, ödem veya alerjik döküntü.",
		"Diğer NSAİİ ağrı kesiciler, aspirin, varfarin/DOAC antikoagülanlar ve lityum ile kombine edilmemelidir.",
	),
	_d(
		"Dikloron / Voltaren (Diklofenak)",
		"Diklofenak Sodyum / Potasyum",
		"Dikloron, Voltaren, Cataflam, Diclomec",
		"NSAİİ (Analjezik & Antiinflamatuar)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Eklem, kas-iskelet ağrıları, bel/boyun tutulması ve travmatik şişliklerde kullanılır; jel formu lokal ağrı bölgesine sürülür.",
		"Tablet formu TOK KARNINA bütün yutulur. Jel formu haricen sağlam deriye ince tabaka halinde sürülür, açık yaraya uygulanmaz.",
		"Yetişkinlerde günlük maksimum 150 mg (oral); 14 yaş altı çocuklarda sistemik tablet formu önerilmez.",
		"Mide ağrısı, bulantı, baş ağrısı, nadir olarak tansiyon yükselmesi, karaciğer enzim artışı veya alerji.",
		"Diğer NSAİİ'ler, kan sulandırıcılar, digoksin, lityum ve potasyum tutucu diüretiklerle birlikte kullanılmamalıdır.",
	),
	_d(
		"A-ferin / Katarin / Tylol Hot",
		"Parasetamol + Klorfeniramin / Psödoefedrin",
		"A-ferin Pediatrik, A-ferin Sinus, Katarin, Tylol Hot, İburamin Cold",
		"Soğuk Algınlığı & Grip Kombinasyonu",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Üst solunum yolu enfeksiyonlarına eşlik eden burun akıntısı, tıkanıklık, hapşırma, ateş ve kırgınlık belirtilerini hafifletmek için kullanılır.",
		"Ağızdan şurup, kapsül veya sıcak suyla eritilen saşe halinde TOK KARNINA alınır. Antihistaminik içeriği nedeniyle sersemlik yapabilir.",
		"Çocuk formları yaş/kiloya göre genellikle 6-8 saatte 1 ölçek; içerdiği Parasetamol toplamı 60 mg/kg/gün sınırını ASLA aşmamalıdır. Yetişkin maksimum günde 3-4 doz.",
		"Uyku hali, sersemlik, ağız kuruluğu, çarpıntı (psödoefedrin içerenlerde) veya huzursuzluk.",
		"İçinde PARASETAMOL veya İBUPROFEN bulunduğu için ayrıca Calpol/Parol/Dolven ile birlikte kesinlikle verilmemelidir! MAO inhibitörleri, sedatifler ve tiroid/tansiyon ilaçlarıyla etkileşir.",
		"6 yaş altı çocuklarda dekonjestan (psödoefedrin/fenilefrin) içeren formlar hekim onayı olmadan verilmemelidir.",
	),
	# --- 2. ANTİBİYOTİKLER ---
	_d(
		"Augmentin-BID / Klamoks (Amoksisilin+Klavulanat)",
		"Amoksisilin + Klavulanik Asit",
		"Augmentin-BID 200/400/600/1000, Klamoks, Amoklavin, Croxilex, Bioment",
		"Penisilin Grubu Geniş Spektrumlu Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Yemekle birlikte",
		"Buzdolabı (+2 / +8 °C)",
		"Bademcik/boğaz iltihabı (tonsillit/farenjit), orta kulak iltihabı, sinüzit, alt solunum yolu ve idrar yolu bakteriyel enfeksiyonlarının tedavisinde kullanılır.",
		"Mide-bağırsak yan etkilerini azaltmak için YEMEĞİN BAŞINDA veya hemen tok karnına ağızdan verilir. Sulandırılan şurup mutlaka buzdolabında (+2/+8 °C) saklanmalı ve 7-10 gün içinde kullanılmalıdır.",
		"Çocuklarda standart 25-45 mg/kg/gün (yüksek doz ES formunda 90 mg/kg/gün Amoksisilin bileşeni baz alınır), 12 saatte 1 (2x1) verilir. Yetişkin max 2000 mg/gün amoksisilin (genellikle 2x1 1000 mg).",
		"İshal, yumuşak dışkılama, karın ağrısı, bulantı, oral/bez pamukçuk. PENİSİLİN ALERJİSİ olanlarda ani kurdeşen, yüz-dudak şişmesi ve nefes darlığı (anafilaksi) gelişebilir — ACİL 112!",
		"Allopurinol (döküntü riskini artırır), varfarin ve metotreksat ile etkileşir. Penisilin veya sefalosporin alerjisi olan öğrencilerde KESİNLİKLE kontrendikedir.",
		"Öğrencinin alerji hanesinde 'Penisilin' yazıyorsa doz kesinlikle uygulanmaz; sulandırılmış şişeler her kullanım öncesi çalkalanmalıdır.",
	),
	_d(
		"Largopen / Alfoxil (Amoksisilin)",
		"Amoksisilin",
		"Largopen 250/500/1000, Alfoxil, Amoksilav",
		"Penisilin Grubu Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Beta streptokok boğaz enfeksiyonu, akut otitis media, sinüzit ve diş enfeksiyonlarında kullanılan dar/orta spektrumlu penisilin antibiyotiktir.",
		"Aç veya tok ağızdan alınabilir (tok karnına mide konforu daha iyidir). Sulandırılmış süspansiyon serin yerde/buzdolabında 14 gün saklanabilir; her defasında çalkalanır.",
		"Çocuklarda genellikle 40-90 mg/kg/gün (2 veya 3 bölünmüş doz); günlük maksimum 3000 mg (yetişkin 3 g/gün).",
		"Hafif ishal, bulantı, döküntü; penisilin alerjisi durumunda kaşıntı, kabartı veya anafilaksi.",
		"Penisilin/beta-laktam alerjisinde kontrendikedir. Allopurinol, metotreksat ve probenesid ile etkileşir.",
	),
	_d(
		"Zinnat / Cefaks (Sefuroksim)",
		"Sefuroksim Aksetil",
		"Zinnat 125/250/500, Cefaks, Aksef, Enfüron",
		"2. Kuşak Sefalosporin Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Buzdolabı (+2 / +8 °C)",
		"Üst ve alt solunum yolu enfeksiyonları, kulak iltihabı, sinüzit ve idrar yolu enfeksiyonlarının tedavisinde kullanılır.",
		"Emiliminin tam olması için mutlaka TOK KARNINA (yemekten hemen sonra) ağızdan verilir. Sulandırılmış süspansiyon buzdolabında (+2/+8 °C) saklanır; tadı acı olduğundan hemen ardından meyve suyu/süt verilebilir.",
		"Çocuklarda günlük 20-30 mg/kg/gün (2 doza bölünerek), tek doz max 250-500 mg; yetişkinlerde genellikle günde 2x1 500 mg (max 1000 mg/gün).",
		"Bulantı, karın ağrısı, ishal, baş ağrısı, alerjik döküntü.",
		"Sefalosporin veya şiddetli penisilin alerjisinde dikkat/kontrendikasyon. Güçlü mide asidi baskılayıcılarla (PPI/antiasit) aynı anda alınması emilimi azaltabilir.",
	),
	_d(
		"Suprax (Sefiksim)",
		"Sefiksim",
		"Suprax 100/200/400, Zimaks, Molcef, Cefspan",
		"3. Kuşak Sefalosporin Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"İdrar yolu enfeksiyonları, orta kulak iltihabı ve üst solunum yolu enfeksiyonlarında çoğunlukla günde 1-2 kez kullanılan antibiyotiktir.",
		"Aç veya tok karnına ağızdan alınabilir; mide hassasiyetinde yemekle birlikte verilir. Sulandırıldıktan sonra oda sıcaklığında 14 gün dayanabilir (buzdolabına konmaz).",
		"Çocuklarda 8 mg/kg/gün (tek doz veya 2 bölünmüş doz, günlük max 400 mg); yetişkinlerde max 400 mg/gün.",
		"İshal, karın ağrısı, gaz, bulantı, baş dönmesi, alerjik reaksiyon.",
		"Sefalosporin alerjisi olanlarda kontrendikedir; kan sulandırıcı (varfarin) etkisini artırabilir.",
	),
	_d(
		"Macrol / Klacid (Klaritromisin)",
		"Klaritromisin",
		"Macrol 125/250/500, Klacid, Claricide, Laricid",
		"Makrolid Grubu Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Penisilin alerjisi olanlarda alternatif olarak da kullanılan, boğaz, sinüs, atipik zatürre ve solunum yolu enfeksiyonlarının antibiyotiğidir.",
		"Aç veya tok ağızdan alınır. SÜSPANSİYON FORMU KESİNLİKLE BUZDOLABINA KONULMAZ (granüller sertleşip acılaşır); oda sıcaklığında (15-25 °C) saklanır ve çalkalanır.",
		"Çocuklarda 15 mg/kg/gün (2 doz halinde 12 saatte bir), max 1000 mg/gün; yetişkin max 1000 mg/gün.",
		"Ağızda metalik/acı tat, mide krampı, bulantı, ishal.",
		"Kalp ritmini etkileyen ilaçlar, karbamazepin (Tegretol), teofilin, bazı kolesterol ilaçları (statinler) ve varfarin ile CİDDİ ilaç etkileşimi yapar — epilepsi ilaçlı öğrencilerde doktor onayı teyit edilmelidir.",
	),
	_d(
		"Azitro (Azitromisin)",
		"Azitromisin",
		"Azitro 200/500, Zitromax, Azax, Azeltin",
		"Makrolid Grubu Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Kısa süreli (genellikle 3-5 günlük) solunum yolu, orta kulak iltihabı ve boğaz enfeksiyonu antibiyotiğidir.",
		"Günde 1 kez ağızdan alınır; tablet/şurup genelde aç veya yemekten 1 saat önce/2 saat sonra tercih edilir (mide rahatsızlığında hafif tok alınabilir).",
		"Çocuklarda günde 1 kez 10 mg/kg/gün (genelde max 500 mg/gün); yetişkin 500 mg/gün.",
		"Karın ağrısı, bulantı, ishal, nadir olarak kalp ritmi değişikliği (QT uzaması).",
		"Alüminyum/magnezyum içeren antiasitlerle aynı anda alınmaz (en az 2 saat ara). Kalp ritim ilaçları, digoksin ve varfarinle etkileşir.",
	),
	_d(
		"Flagyl / Biteral (Metronidazol / Ornidazol)",
		"Metronidazol / Ornidazol",
		"Flagyl, Nidazol, Biteral, Ornisid",
		"Nitroimidazol Antibakteriyel / Antiparaziter",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Anaerobik bakteriyel enfeksiyonlar, diş apseleri, giardia/amip gibi paraziter bağırsak enfeksiyonlarında kullanılır.",
		"Mide tahrişini ve metalik tadı azaltmak için TOK KARNINA ağızdan verilir.",
		"Çocuklarda endikasyona göre 15-40 mg/kg/gün (genelde max 2000 mg/gün); yetişkin max 2000 mg/gün.",
		"Ağızda metalik tat, bulantı, koyu renkli idrar (zararsız), baş dönmesi.",
		"Alkol (yetişkinlerde şiddetli disülfiram reaksiyonu: kusma/çarpıntı), lityum, varfarin ve yüksek doz busulfan ile birlikte KESİNLİKLE kullanılmaz.",
	),
	# --- 3. SOLUNUM, ASTIM & ALERJİ ---
	_d(
		"Ventolin (Salbutamol İnhaler / Nebül)",
		"Salbutamol (Albuterol)",
		"Ventolin Evohaler 100 mcg, Ventolin Nebules, Ronkotol, Salbutol",
		"Kısa Etkili Bronkodilatör (Astım Kurtarıcı)",
		"İnhaler / Nebül",
		"Solunum (İnhalasyon / Hava)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Astım atağı, nefes darlığı, hışıltılı solunum (wheezing) ve egzersize bağlı bronkospazmda hava yollarını hızlıca genişletmek için ACİL KURTARICI ilaçtır.",
		"Ölçülü doz inhaler (sprey) ağızdan nefesle çekilir; okul çağı çocuklarda mutlaka aracı tüp (AeroChamber / spacer) ile uygulanmalıdır. Acil solunum sıkıntısı sonrası öğrenci ayakta dolaştırılmamalıdır.",
		"İnhaler (100 mcg/puf): Okul revirinde atağın durumuna göre tek seferde 1-2 puf (acil protokolde hekim planıyla 2-4 puf); günlük normalde en fazla 8 puf (800 mcg) aşılırsa acil tıbbi değerlendirme gerekir.",
		"El titremesi (tremor), kalp çarpıntısı (taşikardi), hafif baş ağrısı veya sinirlilik hali — kısa sürede geçicidir.",
		"Seçici olmayan beta bloker tansiyon/kalp ilaçlarıyla (propranolol vb.) zıt etki gösterir. Şiddetli taşikardi yapan dekonjestanlarla ve düşük potasyum yapabilen diüretiklerle dikkatli izlenmelidir.",
		"Kurtarıcı inhaleri yanında/revirde olan astımlı öğrencinin nefes darlığı 15-20 dk içinde düzelmezse 112 ve veli derhal aranmalıdır.",
	),
	_d(
		"Pulmicort / Budecort (Budesonid Nebül)",
		"Budesonid",
		"Pulmicort Nebül 0.25/0.5 mg, Budecort, Cortair, Miflonide",
		"İnhale Kortikosteroid (Astım / Krup Kontrol)",
		"İnhaler / Nebül",
		"Solunum (İnhalasyon / Hava)",
		"Fark etmez",
		"Işıktan Korunan Dolap",
		"Astım koruyucu idame tedavisi ve çocuklarda krup (akut laringotrakeit / havlar tarzda öksürük) atağının tedavisinde solunum yolu ödemini azaltır.",
		"Nebülizatör cihazı ile maske/ağızlık üzerinden solunur. UYGULAMA SONRASI AĞIZ SU İLE ÇALKALANMALI VE YÜZ YIKANMALIDIR (pamukçuk önlemi). Açılan folyo poşet ışıktan korunmalıdır.",
		"Çocuklarda günlük genellikle 0.25-1 mg/gün (krup atağında tek sefer hekim ordarıyla 2 mg nebül). Yetişkin idame günlük max 1.6-2 mg.",
		"Ağız ve boğazda pamukçuk (kandidiyazis), ses kısıklığı, boğazda hafif tahriş.",
		"Güçlü CYP3A4 inhibitörleri (ketokonazol, itrakonazol, klaritromisin, ritonavir) budesonid düzeyini yükseltebilir.",
	),
	_d(
		"Flixotide (Flutikazon İnhaler)",
		"Flutikazon Propiyonat",
		"Flixotide 50/125/250 mcg Evohaler, Flixotide Nebules",
		"İnhale Kortikosteroid",
		"İnhaler / Nebül",
		"Solunum (İnhalasyon / Hava)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Astımda hava yolu iltihabını baskılayan uzun süreli koruyucu sprey ilaçtır. Ani astım krizini durdurmaz (acil atak için Ventolin gerekir!).",
		"Aracı tüp (spacer) ile ağızdan inhale edilir ve uygulama sonrası mutlaka su ile ağız çalkalanıp tükürülür.",
		"Çocuklarda 100-200 mcg/gün (bölünmüş doz, max 400 mcg/gün); yetişkin/ergen max 1000-2000 mcg/gün.",
		"Ses kısıklığı, ağızda pamukçuk, boğaz kuruluğu.",
		"Ritonavir ve güçlü azol antifungallerle (ketokonazol) sistemik steroid yan etkisi riski nedeniyle birlikte kullanılmamalıdır.",
	),
	_d(
		"Singulair / Onceair (Montelukast)",
		"Montelukast Sodyum",
		"Singulair 4/5/10 mg, Onceair, Zespira, Notta",
		"Lökotrien Reseptör Antagonisti (Astım & Alerji)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Alerjik rinit ve astım ataklarının önlenmesinde günlük idame tedavisi olarak kullanılır (akut nefes darlığı krizi tedavisinde kullanılmaz).",
		"Genellikle günde 1 kez AKŞAM yatmadan önce ağızdan (4/5 mg çiğneme tableti veya granül; 10 mg film tablet) verilir.",
		"Yaşa göre günde 1 kez: 2-5 yaş 4 mg, 6-14 yaş 5 mg, 15+ yaş/yetişkin 10 mg. Günde 1 tableti aşmayınız.",
		"Baş ağrısı, karın ağrısı; nadiren uyku bozukluğu, kabus görme veya davranış/duygu durum değişikliği — aileye bildirilmelidir.",
		"Fenobarbital, fenitoin ve rifampisin ilacın etkisini azaltabilir.",
	),
	_d(
		"Zyrtec / Allerset (Setirizin)",
		"Setirizin Dihidroklorür",
		"Zyrtec Damla/Şurup/Tablet, Allerset, Cetryn, Hitrizin",
		"2. Kuşak Antihistaminik (Alerji İlacı)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Mevsimsel alerjik nezle, kurdeşen (ürtiker), alerjik kaşıntı ve böcek ısırığı sonrası hafif alerjik cilt döküntülerinde kullanılır.",
		"Aç veya tok karnına ağızdan alınır; şurup/damla suya karıştırılabilir. Bazı öğrencilerde hafif uyku hali yapabileceğinden uygulama sonrası sınıfta izlenir.",
		"2-5 yaş 5 mg/gün; 6-11 yaş 10 mg/gün (2x5 mg veya 1x10 mg); 12+ yaş ve yetişkin max 10 mg/gün.",
		"Hafif uyuklama/uyku hali, ağız kuruluğu, baş ağrısı veya yorgunluk.",
		"Alkol ve merkezi sinir sistemini baskılayan sakinleştirici/uyku yapıcı ilaçlarla birlikte alındığında uyku halini artırır.",
	),
	_d(
		"Aerius / Deloday (Desloratadin)",
		"Desloratadin",
		"Aerius Şurup/Tablet, Deloday, Lordes, Desmont",
		"Sedasyonsuz Antihistaminik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Alerjik rinit (hapşırma, burun akıntısı, göz yaşarması) ve kronik/akut ürtiker (kurdeşen) tedavisinde kullanılır.",
		"Günde 1 kez aç veya tok karnına ağızdan verilir. Zyrtec'e kıyasla belirgin uyku yapma riski düşüktür.",
		"1-5 yaş 1.25 mg (2.5 ml); 6-11 yaş 2.5 mg (5 ml); 12+ yaş ve yetişkin 5 mg (10 ml veya 1 tablet) / gün maksimum.",
		"Çok nadiren baş ağrısı, ağız kuruluğu, halsizlik.",
		"Bilinen ciddi ilaç etkileşimi azdır; ağır böbrek/karaciğer yetmezliğinde doz ayarlaması gerektirir.",
	),
	_d(
		"Avil (Feniramin Maleat)",
		"Feniramin Maleat",
		"Avil Tablet, Avil Ampul, Avil Merhem",
		"1. Kuşak Antihistaminik",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Akut alerjik reaksiyon, ürtiker, yoğun kaşıntı ve böcek sokması alerjilerinde kullanılır; ancak anafilakside tek başına İLK SEÇENEK DEĞİLDİR (adrenalin esastır).",
		"Tablet ağızdan tok karnına su ile alınır; merhem kaşıntılı bölgeye haricen sürülür. Belirgin uyuşukluk yapar.",
		"Yetişkinlerde 22.7 mg 1-3x1 (günlük max ~3 mg/kg veya 150 mg). Çocuklarda acil dışı rutin kullanımda 2. kuşak antihistaminikler tercih edilir.",
		"Güçlü uyku hali, dikkat dağınıklığı, ağız kuruluğu, bulanık görme, baş dönmesi.",
		"Alkol, uyku ilaçları, anksiyolitikler, kas gevşeticiler ve MAO inhibitörleriyle birlikte alınmamalıdır (şiddetli sedasyon riski).",
	),
	_d(
		"Zaditen (Ketotifen)",
		"Ketotifen",
		"Zaditen Şurup/Tablet, Zaditen Göz Damlası",
		"Antihistaminik & Mast Hücresi Stabilizatörü",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Alerjik rinit, alerjik konjonktivit (göz damlası) ve çocukluk çağı alerjik astımında koruyucu olarak kullanılır.",
		"Şurup yemekle/tok karnına ağızdan; göz damlası alt göz kapağı kesesine damlatılır.",
		"3 yaş üstü çocuklar ve yetişkinler: Günde 2x1 mg (max 2 mg/gün). Göz damlası günde 2x1 damla.",
		"İlk günlerde uyku hali, ağız kuruluğu, iştah artışı; göz damlasında geçici batma.",
		"Oral diyabet ilaçlarıyla nadiren trombosit düşüklüğü bildirimi; sakinleştiricilerle sedasyon artışı.",
	),
	_d(
		"Nasonex / Nazoster (Mometazon Burun Spreyi)",
		"Mometazon Furoat",
		"Nasonex Aquous, Nazoster, Rinoclenil",
		"Nazal Kortikosteroid Sprey",
		"Burun Spreyi / Damla",
		"Burun Spreyi",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Alerjik nezleye bağlı burun tıkanıklığı, hapşırma ve burun akıntısının uzun süreli kontrolünde kullanılır.",
		"Şişe çalkalanıp baş hafif öne eğilerek sprey ucu burun orta bölmesinden dış yana doğru sıkılır; her uygulama öncesi burun nazikçe temizlenmelidir.",
		"2-11 yaş günde 1 kez her burun deliğine 1 puf (100 mcg); 12+ yaş ve yetişkin günde 1 kez 2 puf (max 200-400 mcg/gün).",
		"Hafif burun kanaması, burun içinde kuruluk veya hapşırma.",
		"Tedavi edilmemiş burun içi mantar/bakteriyel enfeksiyon ve yakın burun ameliyatı sonrasında kullanılmaz.",
	),
	_d(
		"İliadin / Otrivine (Burun Açıcı Sprey)",
		"Oksimetazolin / Ksilometazolin",
		"İliadin Pediatrik/Yetişkin, Otrivine, Rynza",
		"Topikal Nazal Dekonjestan",
		"Burun Spreyi / Damla",
		"Burun Spreyi",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Soğuk algınlığı veya sinüzitte akut burun tıkanıklığını hızla açmak için kısa süreli kullanılır.",
		"Baş dik tutulup her burun deliğine uygulanır. KESİNLİKLE 5 GÜNDEN UZUN KESİNTİSİZ KULLANILMAMALIDIR (ilaç kaynaklı geri tepme tıkanıklığı yapar).",
		"Yaşa uygun pediatrik/yetişkin formundan günde en fazla 2-3 kez 1'er puf/damla.",
		"Burunda yanma/kuruma, hapşırma; uzun kullanımda kalıcı burun tıkanıklığı (rinitis medikamentoza).",
		"MAO inhibitörleri, trisiklik antidepresanlar, kontrolsüz yüksek tansiyon ve 2 yaş altı bebeklerde yetişkin formları kontrendikedir.",
	),
	_d(
		"Asist (Asetilsistein)",
		"Asetilsistein (NAC)",
		"Asist Şurup 200/4, Asist Plus 600, Mucinac, Oxxa",
		"Mukolitik (Balgam Söktürücü)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yoğun, yapışkan balgamlı öksürük ve bronşit durumlarında balgamı sulandırıp atılımını kolaylaştırır.",
		"Efervesan tablet bir bardak suda eritilerek, şurup ise tok karnına içilir. Etkisini göstermesi için öğrencinin gün boyu bol su içmesi teşvik edilmelidir.",
		"2-5 yaş 200 mg/gün; 6-14 yaş 400 mg/gün; 14+ yaş ve yetişkin günde max 600 mg/gün (rutin mukolitik kullanım).",
		"Hafif bulantı, kusma, kükürt/yumurta kokusuna bağlı mide bulantısı; astımlılarda nadiren bronkospazm.",
		"Öksürük kesici (antitüsif) şuruplarla birlikte verilmemelidir (sulanan balgamın akciğerde birikmesine yol açar!). Aktif kömür ve bazı oral antibiyotiklerle (tetrasiklin, sefalosporin) en az 2 saat ara ile alınmalıdır.",
	),
	_d(
		"Levopront / Perebron (Öksürük Şurubu)",
		"Levodropropizin / Oksolamin",
		"Levopront, Perebron, Sekodin, Sinecod",
		"Antitüsif (Kuru Öksürük Kesici)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Balgamsız, boğazı tahriş eden kuru ve gıcık şeklindeki öksürüğü yatıştırmak amacıyla kullanılır.",
		"Levopront tercihen yemeklerden 1 saat önce veya 2 saat sonra ağızdan verilir; Perebron tok karnına da verilebilir.",
		"Çocuklarda kiloya göre (Levopront 1 mg/kg dozunda günde en fazla 3 kez); 2 yaş altında rutin kullanılmaz.",
		"Hafif yorgunluk, uyuklama, bulantı.",
		"Balgam söktürücü (Asist/mukolitik) ilaçlarla aynı anda verilmemelidir! Bol balgamlı öksürükte kullanılmaz.",
	),
	# --- 4. ACİL ANAFİLAKSİ, EPİLEPSİ & NÖBET ---
	_d(
		"EpiPen / Penepin (Adrenalin Oto-Enjektör)",
		"Adrenalin (Epinefrin)",
		"EpiPen 0.3 mg, EpiPen Jr 0.15 mg, Penepin",
		"Acil Anafilaksi Kurtarıcı Sempatomimetik",
		"Enjeksiyon / Kalem / Oto-Enjektör",
		"Enjeksiyon (Deri Altı/İnsülin)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Ağır gıda/arı/ilaç alerjisine bağlı ANAFİLAKTİK ŞOK (nefes darlığı, hırıltı, dil-boğaz şişmesi, tansiyon düşmesi, bilinç bulanıklığı) tablosunda HAYAT KURTARICI birinci basamak acil müdahale ilacıdır.",
		"Koruyucu kapak çıkarılır; KALÇA-UYLUK DIŞ YAN KASA (giysi üzerinden de uygulanabilir) 90 derece dik açıyla bastırılıp klik sesi duyulduktan sonra 3-10 saniye beklenir. Dondurulmaz ve buzdolabına KONULMAZ (oda sıcaklığında kilitli/erişilebilir acil alanda kalır).",
		"15-30 kg çocuk: 0.15 mg (Junior); ≥30 kg çocuk ve yetişkin: 0.30 mg. Semptomlar düzelmezse hekim/112 talimatıyla 5-15 dk sonra ikinci doz uygulanabilir.",
		"Geçici kalp hızlanması, çarpıntı, titreme, solukluk, kaygı hissi — hayat kurtarıcı etki yanında beklenen etkilerdir.",
		"ANAfilaksi acilinde MUTLAK KONTRENDİKASYON YOKTUR — derhal uygulanmalı ve aynı anda 112 ACİL SERVİS ile VELİ aranmalıdır! Renkli solüsyon bulanık/kahverengi ise değiştirilmelidir.",
		"Uygulama sonrası çocuk sırtüstü yatırılıp bacakları hafif yükseltilir (solunum sıkıntısı varsa yarı oturur) ve mutlaka ambulansla hastaneye sevk edilir.",
	),
	_d(
		"Stesolid (Rektal Diazepam)",
		"Diazepam",
		"Stesolid 5 mg / 10 mg Rektal Tüp, Diazem",
		"Acil Nöbet Durdurucu Benzodiazepin",
		"Rektal Tüp / Fitil",
		"Rektal",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Epilepsi veya ateşli havale (febril konvülsiyon) sırasında 5 dakikayı aşan uzun süreli nöbetleri durdurmak için okul acil protokolü kapsamında kullanılır.",
		"Öğrenci yan yatırılır; rektal tüpün ucu anüsten dikkatlice yerleştirilip tüp tamamen sıkılır ve tüp sıkılı halde çıkarılır, ardından kalça yanakları 30 sn kapalı tutulur.",
		"<15 kg: 5 mg tek doz; ≥15 kg (veya >3 yaş hekim planına göre): 10 mg. Hekim onayı olmadan okulda peş peşe tekrarlanmaz; solunum izlenmelidir.",
		"Belirgin uyku hali, kas gevşekliği; nadir ancak hayati önemde SOLUNUM BASKILANMASI (yüzeyel nefes alma).",
		"Alkol, diğer sakinleştiriciler ve opioidlerle birlikte solunum durması riskini ciddi artırır. Uygulama sonrası yan yatış (recovery) pozisyonunda solunum takip edilir ve acil nöbet protokolüne göre 112/veli bilgilendirilir.",
	),
	_d(
		"Keppra (Levetirasetam)",
		"Levetirasetam",
		"Keppra Şurup/Tablet, Levetirasetam, Epixx, Keppra 500",
		"Antiepileptik (Nöbet Önleyici)",
		"Oral Damla / Solüsyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Epilepsi (sara) nöbetlerinin önlenmesinde günde iki kez düzenli saatlerde alınan antiepileptik ilaçtır.",
		"Aç veya tok karnına ağızdan alınır; şurup formu ağıza doğrudan veya bir bardak suyla karıştırılıp verilir. Saat aksatılmamalıdır.",
		"Çocuklarda günlük ortalama 20-60 mg/kg/gün (2 doz halinde); yetişkin günlük maksimum 3000 mg/gün.",
		"Uyku hali, baş dönmesi, nadiren sinirlilik/huysuzluk veya davranış değişikliği.",
		"Bilinen ciddi ilaç etkileşimi azdır; metotreksat düzeyini artırabilir. Nöbet ilaçları ANİDEN KESİLMEZ (status epileptikus riski).",
	),
	_d(
		"Depakin (Valproik Asit)",
		"Valproik Asit / Sodyum Valproat",
		"Depakin Şurup 200 mg/ml, Depakin Chrono 500, Convulex",
		"Geniş Spektrumlu Antiepileptik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Epilepsi nöbetlerinin ve bazı durumlarda migren/duygudurum bozukluklarının kontrolünde kullanılır.",
		"Mide bulantısını önlemek için TOK KARNINA ağızdan verilir. Chrono uzatılmış salımlı tabletler ÇİĞNENMEDEN ve EZİLMEDEN yutulur.",
		"Çocuklarda idame genellikle 20-30 mg/kg/gün; yetişkin max 2500-3000 mg/gün (kan düzeyi hekimce izlenir).",
		"Bulantı, iştah/kilo değişimi, elde hafif titreme, uyku hali; karın ağrısı+kusma+sarılık olursa pankreas/karaciğer açısından acil değerlendirme gerekir.",
		"Karbapenem antibiyotikler (meropenem vb. ilacı etkisizleştirir), aspirin (küçük çocuklarda Reye sendromu ve toksisite riski!), lamotrijin ve topiramat ile kritik etkileşime girer.",
	),
	_d(
		"Tegretol (Karbamazepin)",
		"Karbamazepin",
		"Tegretol 200/400 CR, Tegretol Şurup, Karazepin",
		"Antiepileptik & Nöropatik Ağrı İlacı",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Fokal ve jeneralize epilepsi nöbetlerinin önlenmesinde kullanılan düzenli antiepileptik ilaçtır.",
		"Yemeklerle/TOK KARNINA alınır; CR tabletler çiğnenmez/ezilmez. Şurup her kullanım öncesi çalkalanmalıdır. Greyfurt suyuyla alınmamalıdır.",
		"Çocuklarda ortalama 10-20 mg/kg/gün; yetişkin günlük max 1200-1600 mg.",
		"Baş dönmesi, çift görme, uyku hali, dengesizlik; nadir ama kritik olarak deride soyulmalı yaygın döküntü veya kan hücresi düşüklüğü.",
		"MAO inhibitörleri, klaritromisin/eritromisin (makrolidler Tegretol zehirlenmesi yapar!), valproat ve doğum kontrol haplarıyla ciddi etkileşir.",
	),
	# --- 5. DİYABET / ENDOKRİN ---
	_d(
		"NovoRapid / Humalog (Hızlı Etkili İnsülin)",
		"İnsülin Aspart / İnsülin Lispro / İnsülin Glulisin",
		"NovoRapid FlexPen, Humalog KwikPen, Fiasp, Apidra",
		"Hızlı Etkili Öğün / Düzeltme İnsülini",
		"Enjeksiyon / Kalem / Oto-Enjektör",
		"Enjeksiyon (Deri Altı/İnsülin)",
		"Yemekten 30 dk önce",
		"Buzdolabı (+2 / +8 °C)",
		"Tip 1 diyabetli öğrencilerde öğün karbonhidratını karşılamak ve yüksek kan şekerini düzeltmek için yemekten hemen önce (0-15 dk önce) veya karbonhidrat sayımına göre uygulanan hızlı insülindir.",
		"Karın, kol, uyluk veya kalça deri altı dokusuna kalem iğnesiyle (4-6 mm) dik açıyla yapılır; enjeksiyon sonrası iğne içeride 10 sn tutulur. Açılmamış kalemler buzdolabında (+2/+8 °C, dondurulmaz!), kullanımdaki kalem 4 hafta oda sıcaklığında saklanabilir.",
		"Kişiye özel İKO (İnsülin/Karbonhidrat Oranı) ve İDF (Düzeltme Faktörü) üzerinden belirlenir; tek seferlik maksimum diye sabit sayı yoktur — MUTLAKA kan şekeri ölçülüp diyabet planına göre hesaplanır.",
		"HİPOGLİSEMİ (kan şekeri <70 mg/dL: titreme, terleme, solukluk, baş dönmesi, dalgınlık) EN ÖNEMLİ RİSKTİR; enjeksiyon bölgesinde lipohipertrofi (sertlik).",
		"Beta blokerler hipoglisemi belirtilerini maskeleyebilir; kortikosteroidler ve enfeksiyonlar kan şekerini yükseltir. İnsülin yapılıp öğün yenmezse ağır hipoglisemi gelişir!",
		"Tip 1 diyabetli öğrencide öğün yeme durumu kontrol edilmeden bolus insülin yapılmamalı, şüphede kan şekeri 15 dk sonra tekrar bakılmalıdır.",
	),
	_d(
		"Lantus / Tresiba (Uzun Etkili Bazal İnsülin)",
		"İnsülin Glarjin / İnsülin Degludek / Detemir",
		"Lantus SoloStar, Toujeo, Tresiba, Levemir",
		"Uzun Etkili Bazal İnsülin",
		"Enjeksiyon / Kalem / Oto-Enjektör",
		"Enjeksiyon (Deri Altı/İnsülin)",
		"Fark etmez",
		"Buzdolabı (+2 / +8 °C)",
		"Diyabet hastalarında gün boyu ve gece boyunca temel insülin ihtiyacını sağlayan 24+ saat etkili bazal insülindir (yemekten bağımsız sabit saatte uygulanır).",
		"Günün aynı saatinde deri altına yapılır; ASLA damar içine uygulanmaz ve diğer insülinlerle karıştırılmaz. Açılmamış kalemler buzdolabında (+2/+8 °C), kullanımdaki kalem oda sıcaklığında 28 güne kadar saklanır.",
		"Hastanın günlük bazal ünite reçetesine göre sabit saatte verilir; hızlı etkili NovoRapid/Humalog ile KARIŞTIRILMAMALIDIR!",
		"Hipoglisemi, enjeksiyon bölgesinde hafif kızarıklık veya doku sertleşmesi.",
		"Pioglitazon/kortizon ile ödem veya glukoz dengesi değişikliği; hızlı etkili insülin kalemleriyle yanlışlıkla karıştırılması ağır şoka yol açar.",
	),
	_d(
		"GlucaGen HypoKit (Glukagon 1 mg)",
		"Glukagon",
		"GlucaGen HypoKit 1 mg Flakon + Enjektör",
		"Acil Ağır Hipoglisemi Kurtarıcı Hormonu",
		"Enjeksiyon / Kalem / Oto-Enjektör",
		"Enjeksiyon (Deri Altı/İnsülin)",
		"Fark etmez",
		"Buzdolabı (+2 / +8 °C)",
		"Tip 1 diyabetli öğrencide BİLİNÇ KAYBI, NÖBET veya YUTAMAMA gelişen ağır hipoglisemide (<54 mg/dL) ağızdan şeker verilemediğinde derhal uygulanan hayat kurtarıcı hormondur.",
		"Şırıngadaki sıvı toz flakona sıkılıp köpürtmeden çalkalanır, çekilip uyluk/karın kas içi veya deri altına enjekte edilir. Ağızdan kesinlikle su/şeker verilmeye çalışılmaz; hasta yan yatırılır ve 112 aranır.",
		"<25 kg çocuk: 0.5 mg (yarım flakon); ≥25 kg veya ≥6-8 yaş: 1.0 mg (tam flakon) tek acil doz.",
		"Enjeksiyon sonrası uyandıktan sonra bulantı ve kusma çok sıktır — yan yatış pozisyonu aspirasyonu önler.",
		"İnsülinoma veya feokromositomada dikkat; ağır hipoglisemi acilinde beklemeden uygulanır, sonra mutlaka 112 ve veli bilgilendirilir.",
	),
	_d(
		"Glukoz Tableti / Dekstroz %30 (Hipoglisemi Müdahale)",
		"Dekstroz / Saf Glukoz",
		"Glukoz Tablet, Dekstroz Jel, %30 Dekstroz Ampul (Oral)",
		"Hızlı Karbonhidrat / Acil Şeker Desteği",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Bilinci açık olan diyabetli öğrencide kan şekeri <70 mg/dL ölçüldüğünde 15/15 kuralı gereği şekeri en hızlı yükseltmek için verilir.",
		"Bilinci tam açık ve yutabilen öğrenciye 15 gram hızlı karbonhidrat (genelde 3-4 tablet veya 1 küçük kutu meyve suyu / 3-4 küp şeker) ağızdan çiğnetilerek verilir; 15 dk sonra tekrar şeker ölçülür.",
		"Tek atak başına 15 gram (küçük çocukta 10-15 g) glukoz; 15 dk sonra hâlâ <70 ise tekrarlanır. Bilinci kapalı hastaya AĞIZDAN VERİLMEZ!",
		"Fazla alımda geçici yüksek kan şekeri (hiperglisemi).",
		"Bilinci kapalı veya yutma refleksi bozuk hastada aspirasyon tehlikesi nedeniyle ağızdan verilmez (Glucagen uygulanır).",
	),
	_d(
		"Euthyrox / Levotiron (Levotiroksin)",
		"Levotiroksin Sodyum",
		"Euthyrox 25/50/75/100 mcg, Levotiron, Tefor",
		"Tiroid Hormon Replasmanı",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Hipotiroidi ve konjenital tiroid yetmezliğinde eksik tiroid hormonunu yerine koymak için günde bir kez düzenli kullanılır.",
		"Mutlaka SABAH AÇ KARNINA (kahvaltıdan en az 30-60 dakika önce) tek bir yudum suyla alınmalıdır. Süt, kalsiyum ve demir şuruplarıyla AYNI ANDA ALINMAZ.",
		"Kişiye özel mcg dozunda günde 1 kez (genellikle okul saatinden önce evde alınır; unutulup okulda alınacaksa öğünden ayrı aç alınır).",
		"Yüksek dozda çarpıntı, terleme, sinirlilik, kilo kaybı.",
		"Demir şurupları, kalsiyum, magnezyum/alüminyum antiasitler, sukralfat ve PPI mide ilaçlarıyla EN AZ 4 SAAT ARA İLE alınmalıdır (emilimini tamamen bozar!).",
	),
	_d(
		"Glukofaj / Diaformin (Metformin)",
		"Metformin Hidroklorür",
		"Glukofaj 500/850/1000, Diaformin, Matofin, Glucophage",
		"Oral Antidiyabetik (Yetişkin & İnsülin Direnci)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Tip 2 diyabet, prediyabet ve insülin direnci tedavisinde kan şekerini düzenleyen temel ilaçtır.",
		"Mide bulantısı ve ishali önlemek için mutlaka YEMEKLE BİRLİKTE veya YEMEKTEN HEMEN SONRA (tok) alınır.",
		"Yetişkinlerde günlük maksimum 2000-2550 mg/gün (bölünmüş doz). 10 yaş altı çocuklarda kullanılmaz.",
		"Başlangıçta mide ağrısı, şişkinlik, ishal, B12 vitamini azalması; nadir ancak ciddi laktik asidoz riski (şiddetli halsizlik+hızlı solunum).",
		"İyotlu kontrast maddeli BT/tomografi öncesi geçici kesilir; alkol, simetidin ve böbrek fonksiyonunu bozan ilaçlarla etkileşir.",
	),
	# --- 6. MİDE, BAĞIRSAK & BULANTI ---
	_d(
		"Metpamid / Emedur (Bulantı & Kusma İlacı)",
		"Metoklopramid / Trimetobenzamid",
		"Metpamid Şurup/Tablet/Ampul, Emedur Kapsül/Fitil",
		"Antiemetik (Bulantı & Kusma Kesici)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Yemekten 30 dk önce",
		"Oda Sıcaklığı (15-25 °C)",
		"Şiddetli mide bulantısı, kusma ve reflüye bağlı mide boşalması yavaşlığında kullanılır.",
		"Yemeklerden yaklaşık 30 dakika önce alınır; Emedur fitil formu rektal yoldan uygulanabilir.",
		"Metoklopramid: Yetişkinlerde günde en fazla 3x10 mg (max 30 mg/gün, en fazla 5 gün). 18 yaş altı çocuklarda akut ekstrapiramidal distoni (boyun/çene kasılması!) riski yüksek olduğu için yalnızca doktor tavsiyesiyle verilir.",
		"Boyun/yüz/çenede istemsiz kasılma (akut distoni!), uyku hali, huzursuzluk veya ishal.",
		"Metoklopramid 1 yaş altında kontrendike, çocuk/gençlerde rutin önerilmez; epilepsi, bağırsak tıkanıklığı ve bazı antipsikotik/antidepresan ilaçlarla kullanılmaz.",
	),
	_d(
		"Zofran / Zofer (Ondansetron)",
		"Ondansetron",
		"Zofran 4/8 mg, Zofer, Onda",
		"5-HT3 Antagonisti Antiemetik",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Akut gastroenterit (mide üşütmesi) veya migrene bağlı dirençli kusmalarda oral sıvı alımını sağlamak için doktor reçetesiyle kullanılır.",
		"Ağızda dağılan tablet formu su gerekmeden dil üzerine konur; diğer formlar az su ile yutulur.",
		"Çocuklarda tek doz 0.1-0.15 mg/kg (max 4-8 mg/doz, günde 24 mg'ı aşmaz); yetişkin max 16-24 mg/gün.",
		"Baş ağrısı, kabızlık, geçici sıcaklık hissi; nadir QT uzaması.",
		"Apomorfinle kesinlikle birlikte kullanılmaz; kalp ritim ilaçları (amiodaron) ve SSRI/tramadol ile etkileşebilir.",
	),
	_d(
		"Nexium / Lansor / Pulcet (Mide Koruyucu PPI)",
		"Esomeprazol / Lansoprazol / Pantoprazol",
		"Nexium 20/40, Lansor 30, Pulcet, Pantpas",
		"Proton Pompa İnhibitörü (Mide Asit Baskılayıcı)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Reflü (GÖRH), gastrit, mide ülseri ve ağrı kesici ilaçların mideye zararını önleme amacıyla kullanılır.",
		"En yüksek etki için SABAH AÇ KARNINA (kahvaltıdan 30-60 dk önce) bütün olarak yutulmalıdır; kapsüller çiğnenmez/ezilmez.",
		"Yaşa göre günde 1 kez 10-40 mg; ağır reflüde hekim kararıyla max 80 mg/gün.",
		"Baş ağrısı, karın ağrısı, gaz, kabızlık veya hafif ishal.",
		"Klopidogrel (kan sulandırıcı etkisini azaltabilir), metotreksat yüksek dozu, ketokonazol ve demir emilimi ile etkileşir.",
	),
	_d(
		"Gaviscon / Rennie (Reflü & Mide Yanması Şurubu)",
		"Sodyum Aljinat + Kalsiyum/Sodyum Bikarbonat",
		"Gaviscon Likit / Double Action, Gaviscon Infant, Rennie",
		"Antiasit & Reflü Bariyer Oluşturucu",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yemek sonrası mide yanması, ekşime ve asidin yemek borusuna kaçmasını mekanik köpük bariyeriyle hızlıca yatıştırır.",
		"Yemeklerden sonra ve yatmadan önce ağızdan içilir/çiğnenir. Şişe çalkalanmalıdır.",
		"Yetişkin/12+ yaş günde en fazla 4 kez 10-20 ml; daha küçük çocuklarda yalnızca pediatrik form/hekim dozu.",
		"Çok nadiren kabızlık, şişkinlik veya yüksek sodyum alımına bağlı sıvı tutulumu.",
		"Antibiyotik (tetrasiklin/florokinolon), demir ve tiroid ilaçlarının emilimini bozduğu için diğer ilaçlarla EN AZ 2 SAAT ARA İLE alınmalıdır.",
	),
	_d(
		"Reflor / Maflor (Probiyotik)",
		"Saccharomyces boulardii / Lactobacillus",
		"Reflor 250 mg Saşe/Kapsül, Maflor, Enterogermina",
		"Antidiyareik Probiyotik",
		"Efervesan / Saşe",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Akut ishal ve antibiyotik kullanımı sırasında gelişen bağırsak florası bozulmasını ve ishali düzenlemek için kullanılır.",
		"Saşe oda sıcaklığındaki suya, meyve suyuna veya yoğurda karıştırılıp içilir. KESİNLİKLE SICAK SUYA KARIŞTIRILMAZ (canlı probiyotik ölür).",
		"Çocuk ve yetişkinlerde günde 1-2 kez 1 saşe/kapsül (max 500-1000 mg/gün).",
		"Nadir hafif gaz/şişkinlik; santral kateterli veya ağır bağışıklık yetmezlikli kişilerde dikkat edilmelidir.",
		"Mantar ilaçları (oral antifungaller) Saccharomyces boulardii'yi etkisiz hale getirir; beraber alınmaz.",
	),
	_d(
		"Ge-Oral (Oral Rehidratasyon Tuzları - ORS)",
		"Sodyum Klorür + Potasyum + Glukoz (ORS)",
		"Ge-Oral Toz, Rehidraton",
		"Sıvı ve Elektrolit Replasmanı",
		"Efervesan / Saşe",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"İshal ve kusma nedeniyle kaybedilen su ve tuz/elektrolit açığını gidererek dehidratasyonu (sıvı kaybını) önler.",
		"1 paket toz tam 1 litre temiz içme suyunda tamamen eritilir (süt veya meyve suyunda eritilmez!). Yalnızca su ile hazırlanır ve yudum yudum içirilir.",
		"Hazırlanan solüsyon 24 saat içinde tüketilir; çocuğun her sulu dışkısı sonrası 10 ml/kg veya susuzluğa göre yudum yudum.",
		"Aşırı hızlı içirilirse kusma tetiklenebilir (her 2-3 dakikada 1 kaşık/yudum verilmelidir).",
		"Böbrek yetmezliği, inatçı kontrolsüz kusma ve bilinç kaybında ağızdan verilmez (acil serum gerekir).",
	),
	# --- 7. GÖZ, KULAK, BOĞAZ & CİLT / İLK YARDIM ---
	_d(
		"Tobrex / Tobradex (Göz Damlası)",
		"Tobramisin (+ Deksametazon)",
		"Tobrex Göz Damlası/Merhemi, Tobradex Damla",
		"Oftalmik Antibiyotik (+ Steroid)",
		"Göz / Kulak Damlası",
		"Göz Damlası",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Bakteriyel göz iltihabı (çapaklanma, kızarıklık, konjonktivit) tedavisinde kullanılır (Tobradex kortizonlu olduğu için yalnızca göz doktoru reçetesiyle verilir!).",
		"Eller yıkanır; alt göz kapağı hafif aşağı çekilip damlalık ucu göze DEĞDİRİLMEDEN 1 damla damlatılır. İki göz arasında çapraz bulaş olmamasına dikkat edilir.",
		"Genellikle 4-6 saatte bir (günde 4-6 kez) 1 damla. Şişe açıldıktan sonra 28 gün içinde atılır.",
		"Geçici göz batması, hafif kızarıklık; Tobradex'in kontrolsüz uzun kullanımında göz içi basıncı artışı.",
		"Viral/herpes veya mantar göz enfeksiyonlarında steroid içeren Tobradex KESİNLİKLE kontrendikedir! Lens takılıyken damlatılmaz.",
	),
	_d(
		"Siprogut (Siprofloksasin Göz/Kulak Damlası)",
		"Siprofloksasin %0.3",
		"Siprogut Göz/Kulak Damlası, Ciloxan",
		"Florokinolon Oftalmik/Otik Antibiyotik",
		"Göz / Kulak Damlası",
		"Kulak Damlası",
		"Fark etmez",
		"Işıktan Korunan Dolap",
		"Dış kulak yolu iltihabı (yüzücü kulağı), orta kulak akıntısı ve bakteriyel göz enfeksiyonlarının tedavisinde kullanılır.",
		"Kulak uygulaması öncesi damla avuç içinde 1-2 dakika ısıtılır (soğuk damla baş dönmesi yapar!); kulak yukarı gelecek şekilde yan yatırılarak damlatılıp 3 dakika beklenir.",
		"Kulak için günde 2-3 kez 3-4 damla; göz için 4-6 saatte bir 1-2 damla.",
		"Uygulamada geçici hafif yanma, kaşıntı veya damlanın boğaza inmesiyle kötü tat.",
		"Kulak zarı delikliği şüphesi varsa yalnızca hekim onaylı otik solüsyon kullanılmalıdır; şişe açıldıktan 4 hafta sonra atılır.",
	),
	_d(
		"Refresh Tears / Systane (Suni Gözyaşı)",
		"Karboksimetilselüloz / Polietilen Glikol",
		"Refresh Tears, Refresh Liquigel, Systane Ultra, Tears Naturale",
		"Oküler Kayganlaştırıcı / Suni Gözyaşı",
		"Göz / Kulak Damlası",
		"Göz Damlası",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Göz kuruluğu, ekran yorgunluğu, toz/rüzgara bağlı yanma ve tahriş hissini nemlendirerek rahatlatır.",
		"Damlalık göze değmeden her göze 1-2 damla damlatılır ve birkaç kez göz kırpılır; başka göz damlasıyla arasında en az 5-10 dakika beklenir.",
		"İhtiyaç duyuldukça günde 4-6 kez 1-2 damla (tek kullanımlık flakonlarda daha sık kullanılabilir).",
		"Çok nadiren birkaç saniyelik hafif bulanık görme.",
		"Bilinen ciddi sistemik ilaç etkileşimi yoktur; tek dozluk flakon açıldıktan sonra tek seferde atılır.",
	),
	_d(
		"Andorex / Tantum Verde (Boğaz Spreyi)",
		"Benzidamin HCl (+ Klorheksidin)",
		"Andorex Sprey/Gargara, Tantum Verde, Kloroben, Farhex",
		"Topikal Boğaz Antiseptiği & Ağrı Kesici",
		"Boğaz Spreyi / Gargara",
		"Diğer",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Farenjit, tonsillit, ağız içi aft ve boğaz ağrısı/yutkunma zorluğunda lokal ağrıyı ve mikroorganizma yükünü azaltır.",
		"Doğrudan ağrıyan boğaz bölgesine püskürtülür veya gargara yapılıp TÜKÜRÜLÜR; yutulmaz. Uygulamadan sonraki 15-20 dakika yiyip içilmez.",
		"Yetişkin/12+ yaş günde 2-6 kez (en az 1.5-3 saat ara, 4 püskürtme); 6-12 yaşta pediatrik hekim tavsiyesiyle 2 püskürtme; 6 yaş altında boğaz spreyi refleks spazm riski nedeniyle kullanılmaz.",
		"Ağız/dilde geçici uyuşma ve karıncalanma hissi, klorheksidinli formlarda uzun kullanımda geçici diş lekelenmesi.",
		"Astım veya NSAİİ alerjisi öyküsü olan çocuklarda dikkatli olunmalıdır (sprey bronkospazm tetikleyebilir).",
	),
	_d(
		"Fucidin / Bactroban (Antibiyotikli Krem)",
		"Fusidik Asit / Mupirosin",
		"Fucidin Krem/Pomad, Stafine, Bactroban, Balaban, Thiocilline",
		"Topikal Deri Antibiyotiği",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Küçük kesikler, sıyrıklar, iltihaplı cilt yaraları (impetigo) ve tırnak kenarı enfeksiyonlarının lokal tedavisinde kullanılır.",
		"Yara çevresi serum fizyolojik/temiz suyla temizlenip kurulandıktan sonra ince tabaka halinde sürülür ve gerekirse steril gazlı bezle kapatılır. Yalnızca haricen kullanılır, göze sürülmez.",
		"Günde 2-3 kez ince tabaka (genellikle 7-10 gün); büyük açık yanık veya geniş yaralarda doktor görüşü olmadan uygulanmaz.",
		"Uygulama yerinde hafif kızarıklık veya kaşıntı.",
		"İçeriğindeki antibiyotiğe veya baz maddelere karşı alerjide kullanılmaz. Thiocilline (neomisin) alerjenitesi daha yüksek olabilir; Bactroban burun içi formları göz için uygun değildir.",
	),
	_d(
		"Fenistil Jel (Antihistaminik Jel)",
		"Dimetinden Maleat",
		"Fenistil Jel 30 g, Stilex Jel",
		"Topikal Antipruritik (Kaşıntı Giderici)",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Sivrisinek/böcek ısırıkları, hafif kurdeşen, hafif güneş yanığı ve alerjik kaşıntılarda deriyi serinletip kaşıntıyı azaltır.",
		"Haricen kaşıntılı bölgeye sürülür. Güneş ışığına maruz kalacak geniş alanlara veya açık/kanamalı yaralara sürülmez.",
		"Günde 2-4 kez ince tabaka halinde lokal uygulama.",
		"Çok nadiren ciltte geçici kuruluk veya hafif yanma.",
		"Açık yaralara, mukoza ve göz çevresine sürülmez.",
	),
	_d(
		"Silverdin (Yanık Kremi)",
		"Gümüş Sülfadiazin %1",
		"Silverdin Krem, Silvemed",
		"Topikal Yanık Antiseptik / Antibakteriyel",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Okul kazaları veya sıcak su teması sonrası gelişen yüzeyel/kısmi yanıklarda enfeksiyonu önlemek ve ağrıyı azaltmak için kullanılır.",
		"Yanık ilk olarak 15-20 dk soğuk/ılık akan su altında tutulduktan sonra (buz konulmaz!), temiz pansumanla haricen ince sürülür ve steril kapatılır. SÜLFONAMİD ALERJİSİ sorgulanmalıdır!",
		"Günde 1-2 kez pansuman yenilenerek uygulanır.",
		"Ciltte geçici gri/gümüş renklenme (zararsız), hafif yanma veya sülfonamid alerjik döküntü.",
		"Sülfonamid alerjisi olanlarda ve G6PD eksikliği bulunanlarda KESİNLİKLE kullanılmaz! Yüz ve göz çevresi yanıklarında doğrudan sürülmez.",
	),
	_d(
		"Madecassol / Hametan / Bepanthol (Cilt Onarıcı)",
		"Centella asiatica / Hamamelis / Dekspantenol",
		"Madecassol, Hametan Krem/Pomad, Bepanthol, Fito Krem",
		"Epitelizan & Cilt Bariyeri Onarıcı",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Hafif sıyrık, kızarıklık, dudak/cilt çatlaması, sürtünme tahrişi ve kapanma aşamasındaki yaraların doku iyileşmesini hızlandırır.",
		"Temizlenmiş cilde ince tabaka halinde uygulanır; haricen kullanılır. Derin/iltihaplı kirli yaralarda tek başına antibiyotik yerine kullanılmaz.",
		"Günde 1-3 kez ince tabaka uygulama.",
		"Çok nadiren hafif temas hassasiyeti.",
		"Derin enfekte yaralara veya açık göz içine uygulanmaz.",
	),
	_d(
		"Batticon / Baticon (Povidon İyot %10)",
		"Povidon İyot",
		"Batticon Solüsyon %10, Baticon, Biokadin",
		"Harici Yara Antiseptiği",
		"Oral Damla / Solüsyon",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Işıktan Korunan Dolap",
		"Kesik, sıyrık ve düşme yaralarının etrafını ve yüzeyini mikroorganizmalardan arındırmak için okul revirinin temel ilk yardım antiseptiğidir.",
		"Önce toz/toprak varsa bol temiz su veya serum fizyolojik (%0.9 NaCl) ile yıkanır, ardından steril gazlı bezle haricen pansuman yapılır. ASLA İÇİLMEZ ve kapalı dar pansuman altında aşırı bekletilmez.",
		"Yalnızca pansuman sırasında ihtiyaç kadar harici uygulama.",
		"İYOT ALERJİSİ olanlarda kızarıklık/kaşıntı; yoğun ve sürekli geniş yarada tiroid hormonlarını etkileyebilir.",
		"İyot alerjisi, hipertiroidi/tiroid hastalığı aktif olanlarda ve Oksijenli Su (H2O2) veya gümüşlü/cıvalı ürünlerle eşzamanlı yarada kullanılmaz (Silverdin ile aynı anda sürülmez!).",
	),
	_d(
		"Oksibor / Gliserin Kulak Damlası",
		"Gliserin / Hidrojen Peroksit + Borik Asit",
		"Oksibor Kulak Damlası, Gliserin Kulak Damlası, Kulaş",
		"Otik Yumuşatıcı / Serumenolitik",
		"Göz / Kulak Damlası",
		"Kulak Damlası",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Kulak kiri (buşon) tıkanıklığını yumuşatmak ve kulak kanalını rahatlatmak için kullanılır.",
		"Şişe avuçta ısıtıldıktan sonra baş yana yatırılıp 3-5 damla damlatılır ve 5 dakika beklenir.",
		"Günde 2-3 kez 3-5 damla (genellikle 3-5 gün).",
		"Kulakta hafif köpürme hissi (Oksibor'da normal), nadir hafif tahriş.",
		"Kulak zarında delinme (perforasyon) veya kulaktan kan/cerahatlı akıntı varsa KESİNLİKLE damlatılmaz!",
	),
	# --- 8. DEHB, VİTAMİN & YETİŞKİN KRONİK İLAÇLAR ---
	_d(
		"Ritalin / Concerta / Medikinet (Metilfenidat)",
		"Metilfenidat Hidroklorür",
		"Ritalin 10 mg, Concerta 18/27/36/54 mg, Medikinet Retard",
		"SSS Stimülanı (DEHB Tedavisi - Kırmızı Reçete)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Kilitli Revir Dolabı",
		"Dikkat Eksikliği ve Hiperaktivite Bozukluğu (DEHB) tanılı öğrencilerde dikkat sürdürme ve dürtü kontrolünü desteklemek amacıyla psikiyatrist reçetesiyle kullanılır.",
		"SABAH alınır (öğleden sonra alınırsa uykusuzluk yapar!). Concerta tablet BÜTÜN YUTULUR, kesinlikle kırılmaz veya ezilmez (dış kabuğun dışkıda görünmesi normaldir). Okulda saklanıyorsa KİLİTLİ REVİR DOLABINDA tutulur!",
		"Pediatrik hastaya özel hekim titrasyon dozu (genellikle max 54-72 mg/gün; kilo/yaşa bağlı değişir). Yalnızca reçeteli öğrenciye özel uygulanır.",
		"İştahsızlık (öğle yemeğini az yeme), karın ağrısı, uykusuzluk, kalp hızı/tansiyon artışı, nadiren tik artışı veya içe kapanma.",
		"MAO inhibitörleri (son 14 gün içinde alındıysa ölümcül hipertansif kriz!), kontrolsüz hipertansiyon, aritmi ve glokom ile kontrendikedir. Psödoefedrin içeren soğuk algınlığı ilaçlarıyla tansiyonu/çarpıntıyı artırır.",
		"Kırmızı reçeteye tabi kontrollü ilaçtır; stok sayımı doz bazında titizlikle tutulmalıdır.",
	),
	_d(
		"Strattera / Attex (Atomoksetin)",
		"Atomoksetin Hidroklorür",
		"Strattera 10/18/25/40/60 mg, Attex, Atominex",
		"Seçici Noradrenalin Geri Alım İnhibitörü (DEHB)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Çocuk, ergen ve yetişkinlerde DEHB tedavisinde kullanılan stimülan olmayan günlük idame ilaçtır.",
		"Aç veya tok karnına sabah (veya hekim önerisiyle 2 doz) ağızdan alınır. Kapsül açılmamalıdır (gözü tahriş eder!).",
		"≤70 kg çocuklarda başlangıç 0.5 mg/kg/gün, idame 1.2 mg/kg/gün; mutlak max 1.4 mg/kg/gün veya 100 mg/gün.",
		"İştahsızlık, mide bulantısı, karın ağrısı, hafif uyuklama veya kalp hızı artışı; nadiren duygudurum değişimi veya sarılık.",
		"MAO inhibitörleri ile KESİNLİKLE kontrendikedir; Ventolin nebül/tablet ile çarpıntı artabilir; Fluoksetin/Paroksetin ilaç düzeyini yükseltir.",
	),
	_d(
		"Devit-3 / D-Colefor (D3 Vitamini)",
		"Kolekalsiferol (Vitamin D3)",
		"Devit-3 Damla/Kapsül, D-Colefor, Coledan-D3",
		"Vitamin Takviyesi",
		"Oral Damla / Solüsyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Işıktan Korunan Dolap",
		"D vitamini eksikliğinin önlenmesi ve tedavisi, kemik sağlığı ve bağışıklık desteği için kullanılır.",
		"Yağda eriyen vitamin olduğu için yemekle / TOK KARNINA ağızdan damla veya kapsül halinde alınması emilimi artırır.",
		"Günlük idame 400-1000 IU (çocuk) veya 1000-2000 IU (yetişkin); tedavi edici yüksek dozlar doktor kontrolünde kısa süreli verilir (yüksek doz aşımı hiperkalsemi yapar).",
		"Normal dozlarda yan etki çok nadirdir; yüksek doz birikiminde susama, kabızlık, bulantı.",
		"Tiazid diüretikler, kalsiyum yüksekliği (hiperkalsemi) ve digoksin ile dikkatli olunmalıdır.",
	),
	_d(
		"Ferro Sanol B / Ferrum Hausmann (Demir)",
		"Demir II / Demir III Hidroksit Polimaltoz",
		"Ferro Sanol B/Duodenal, Ferrum Hausmann Damla/Şurup, Sidefer",
		"Antianemik (Demir Takviyesi)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Demir eksikliği anemisi (kansızlık) tedavisinde kullanılır; halsizlik ve solukluğu giderir.",
		"Demir II tuzları (Ferro Sanol) tercihen AÇ KARNINA C vitamini (portakal suyu) ile daha iyi emilir; süt/yoğurt/çay ile aynı anda alınmaz. Ferrum (Demir III) tok alınabilir. Diş boyanmasını önlemek için pipetle içirilip su içilir.",
		"Çocuklarda 3-6 mg/kg/gün elementer demir; tek seferlik yüksek doz kazara içilmesi ÇOCUKLARDA ZEHİRLENME ACİLİDİR!",
		"Dışkının siyaha/koyu yeşile boyanması (tamamen normaldir), kabızlık veya hafif karın ağrısı, dişlerde geçici koyulaşma.",
		"Tiroid ilacı (Euthyrox), bazı antibiyotikler (kinolon/tetrasiklin) ve antiasitlerle EN AZ 2-4 SAAT ARA İLE verilmelidir.",
	),
	_d(
		"Dramamine (Dimenhidrinat - Taşıt Tutması)",
		"Dimenhidrinat",
		"Dramamine 50 mg Tablet, Emedur",
		"Antihistaminik & Antivertigo",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Yemekten 30 dk önce",
		"Oda Sıcaklığı (15-25 °C)",
		"Okul gezileri / servis yolculuklarında araç tutması (hareket hastalığı) kaynaklı bulantı, kusma ve baş dönmesini önler.",
		"Yolculuğa çıkmadan 30-60 dakika önce ağızdan suyla alınır. Veli yazılı izni olmadan gezi öncesi verilmez.",
		"6-12 yaş: 25-50 mg (günlük max 150 mg); 12+ yaş ve yetişkin: 50-100 mg (günlük max 400 mg).",
		"Belirgin uyku hali, ağız kuruluğu, baş dönmesi.",
		"Alkol, diğer antihistaminikler ve uyku yapıcı ilaçlarla sedasyonu artırır; astım ve glokomda dikkat.",
	),
	_d(
		"Coraspin / Ecopirin (Asetilsalisilik Asit 100 mg)",
		"Asetilsalisilik Asit (Düşük Doz Aspirin)",
		"Coraspin 100/300 mg, Ecopirin 100, Aspirin Protect",
		"Antiagregan (Kan Sulandırıcı - Yetişkin Personel)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin öğretmen/personelde kalp-damar koruyucu düşük doz kan sulandırıcı olarak kullanılır. 16 YAŞ ALTI ÇOCUKLARDA VİRAL ATEŞTE KESİNLİKLE KULLANILMAZ (Reye Sendromu!).",
		"Enterik kaplı tabletler TOK KARNINA çiğnenmeden bütün yutulmalıdır.",
		"Yetişkin kalp koruması günde 1x100 mg. 16 yaş altında ateş/grip durumunda KONTRENDİKEDİR!",
		"Mide kanaması/ekşimesi, kolay morarma, uzun kanama süresi.",
		"İbuprofen/Naproksen ile aynı saatte alınması aspirin korumasını bozar ve kanamayı artırır; diğer kan sulandırıcılar ve metotreksat ile dikkat.",
	),
	_d(
		"Beloc ZOK / Norvasc (Yetişkin Tansiyon & Kalp İlaçları)",
		"Metoprolol Süksinat / Amlodipin Besilat",
		"Beloc ZOK 25/50/100, Norvasc 5/10 mg, Vazkor, Delix",
		"Antihipertansif & Kardiyovasküler (Yetişkin)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Okul personeli ve yetişkinlerde hipertansiyon, çarpıntı ve kalp damar hastalıklarının düzenli tedavisinde kullanılır.",
		"Günde bir kez tercihen aynı saatte ağızdan alınır; Beloc ZOK tablet ikiye bölünebilir ancak EZİLMEZ/ÇİĞNENMEZ. Ani tansiyon düşüklüğü/baş dönmesinde kişi oturtulup tansiyonu ölçülür.",
		"Yetişkin hekim dozu (Beloc max 200 mg/gün; Norvasc max 10 mg/gün). Çocuklarda okulda verilmez.",
		"Baş dönmesi, ayakta kararma, yorgunluk, ayak bileğinde şişlik (Norvasc), nabız yavaşlaması (Beloc <55 ise dikkat).",
		"NSAİİ ağrı kesiciler (Voltaren/Apranax) tansiyon ilacı etkisini azaltabilir; beta blokerler astım ataklarını ve hipoglisemi belirtilerini kötüleştirebilir/maskeleyebilir.",
	),
	# --- EK SIK KARŞILAŞILAN TÜRKİYE PİYASA İLAÇLARI (TAMAMLAYICI 20+ KAYIT -> ~60 KAYIT) ---
	_d(
		"Nurofen Cold & Flu (İbuprofen + Psödoefedrin)",
		"İbuprofen + Psödoefedrin HCl",
		"Nurofen Cold & Flu 200/30 mg, İburamin Zero/Cold, Dolorin Cold",
		"Soğuk Algınlığı & Ağrı Kesici Kombinasyonu",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Burun tıkanıklığıyla seyreden soğuk algınlığı, sinüs baş ağrısı ve ateş durumlarında kullanılır.",
		"Mideyi korumak için TOK KARNINA bol suyla yutulur. Kalp çarpıntısı veya tansiyon yapabileceğinden gece yatmadan hemen önce önerilmez.",
		"Yetişkin ve 12+ yaş: Günde en fazla 6 tablet (4-6 saatte bir 1-2 tablet; max 1200 mg ibuprofen / 180 mg psödoefedrin).",
		"Mide yanması, kalp çarpıntısı, huzursuzluk, ağız kuruluğu, uykusuzluk.",
		"Diğer ibuprofen/NSAİİ ilaçlarla veya psödoefedrinli başka şuruplarla KESİNLİKLE birleştirilmez! Hipertansiyon, tiroid ve MAO inhibitörlerinde kontrendike.",
	),
	_d(
		"Vermidon / Minoset Plus (Parasetamol + Kafein)",
		"Parasetamol + Kafein / Klorfeniramin",
		"Vermidon, Minoset Plus, Geralgine-Plus, Parol Plus",
		"Analjezik Kombinasyonu",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin ve 12+ yaşlarda baş ağrısı, migren başlangıcı, diş ve kas ağrılarının giderilmesinde kullanılır.",
		"Ağızdan suyla, tercihen tok karnına alınır.",
		"Günlük toplam Parasetamol 4000 mg sınırını aşmayacak şekilde yetişkin max 6-8 tablet/gün (genelde günde en fazla 3-4 kez 1-2 tablet).",
		"Çarpıntı, uyku kaçması (kafeine bağlı), mide rahatsızlığı.",
		"Diğer parasetamol ürünleriyle birlikte alınmaz!",
	),
	_d(
		"Cefatin / Cefaday (Sefaleksin / Sefaklor)",
		"Sefaklor / Sefaleksin",
		"Ceclor, Kefsid, Sef 125/250/500, Maksef",
		"1-2. Kuşak Sefalosporin Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Buzdolabı (+2 / +8 °C)",
		"Cilt/yumuşak doku enfeksiyonu, orta kulak ve boğaz enfeksiyonlarının tedavisinde kullanılır.",
		"Aç veya tok ağızdan verilir; sulandırılmış süspansiyon buzdolabında saklanır ve çalkalanır.",
		"Çocuklarda 20-40 mg/kg/gün; yetişkin max 1500-4000 mg/gün (etken maddeye göre hekim reçetesi).",
		"Hafif ishal, karın ağrısı, alerjik döküntü.",
		"Sefalosporin veya ağır penisilin alerjisi olanlarda kontrendikedir.",
	),
	_d(
		"İespor / Rocephin / Novosef (Enjeksiyonluk Antibiyotik Notu)",
		"Seftriakson Sodyum",
		"Rocephin, Novosef, İesef, Desefin",
		"3. Kuşak Parenteral Sefalosporin",
		"Enjeksiyon / Kalem / Oto-Enjektör",
		"Enjeksiyon (Deri Altı/İnsülin)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Ağır bakteriyel enfeksiyonlarda hastane veya poliklinik kontrolünde kas içi/damar içi uygulanan antibiyotiktir (okul revirinde hekim ordarı ve anafilaksi kiti olmadan rutin IM enjeksiyon yapılmaz!).",
		"Okul revirinde yalnızca tam teşekküllü acil müdahale donanımı ve yazılı resmi hekim istemi mevcutsa değerlendirilir; damar içi kalsiyum içeren solüsyonlarla karıştırılmaz.",
		"Çocuk 50-75 mg/kg/gün (max 2 g/gün); yetişkin 1-4 g/gün.",
		"Enjeksiyon yerinde ağrı, ishal, nadir safra çamuru veya ciddi penisilin/sefalosporin alerjisi.",
		"Damar içi kalsiyum içeren serumlarla KESİNLİKLE aynı anda verilmez; sefalosporin alerjisinde kontrendike.",
	),
	_d(
		"Fucicort / Advantan / Hipokort (Kortizonlu Krem)",
		"Fusidik Asit + Betametazon / Hidrokortizon / Metilprednizolon",
		"Fucicort, Advantan, Hipokort, Locoid, Elocon",
		"Topikal Kortikosteroid / Kombine Krem",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Egzama, temas alerjisi, şiddetli kaşıntılı böcek ısırığı ve iltihaplı cilt reaksiyonlarında kullanılır.",
		"Yalnızca doktorun önerdiği küçük alana günde 1-2 kez ince tabaka sürülür. Yüz, kasık, bebek cildi ve açık yaralara hekimsiz uygulanmaz.",
		"Kısa süreli (5-7 gün) ince lokal uygulama.",
		"Uzun kullanımda ciltte incelme, beyazlama veya kılcal damar belirginleşmesi.",
		"Viral cilt döküntülerinde (suçiçeği, uçuk) ve mantar enfeksiyonlarında kortizonlu krem durumları ağırlaştırır!",
	),
	_d(
		"Travel / Travazol (Mantar & Kaşıntı Kremi)",
		"İzokonazol + Diflukortolon",
		"Travazol, Travocort, Trosyd, Ketoral",
		"Topikal Antifungal Krem",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Ayak mantarı, kasık/vücut mantar enfeksiyonu ve beraberindeki yoğun kaşıntının tedavisinde kullanılır.",
		"Temiz ve kuru cilde haricen ince sürülür; genellikle en fazla 2 hafta kullanılır, ardından steroid içermeyen antifungalle devam edilir.",
		"Günde 2 kez harici ince uygulama.",
		"Hafif lokal yanma veya kızarıklık.",
		"Açık yaralara, göz çevresine ve suçiçeği/uçuk gibi viral lezyonlara sürülmez.",
	),
	_d(
		"Muscoril / Majezik Jel / Fastjel (Kas Gevşetici & Ağrı Jeli)",
		"Tiyokolşikosid / Ketoprofen / İbuprofen Jel",
		"Muscoril, Tyoflex, Fastjel, Dolgit Krem, Bengay",
		"Kas Gevşetici / Topikal Analjezik Jel",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Beden eğitimi dersi burkulması, kas tutulması, bel/boyun spazmında lokal veya sistemik rahatlama sağlar.",
		"Jel formları sağlam cilde hafifçe sürülür (Fastjel sürdükten sonra güneşe çıkılmamalıdır!). Oral Tiyokolşikosid yalnızca yetişkinlerde kısa süreli tok karnına kullanılır.",
		"Tiyokolşikosid (16+ yaş): Max 16 mg/gün (2x8 mg, en fazla 7 gün); jeller günde 2-3 kez.",
		"Ağızdan muskorilde uyuklama veya nadir ishal; jellerde güneşe bağlı cilt duyarlılığı (fotoallerji).",
		"Tiyokolşikosid 16 yaş altı çocuklarda ve gebelikte KESİNLİKLE kontrendikedir! Fastjel açık yaraya sürülmez.",
	),
	_d(
		"Cold-Mix / Vicks VapoRub (Mentol-Okaliptus Damla/Merhem)",
		"Okaliptus Yağı + Kafur / Mentol",
		"Cold-Mix Damla, Vicks VapoRub, Sterimar Deniz Suyu",
		"Doğal Burun Açıcı / Buhar İnhalant",
		"Oral Damla / Solüsyon",
		"Diğer",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Soğuk algınlığında burun tıkanıklığını rahatlatmak için yaka/yastığa damlatılarak veya göğse sürülerek solunur.",
		"KESİNLİKLE AĞIZDAN İÇİLMEZ ve BURUN DELİĞİ İÇİNE SÜRÜLMEZ/DAMLATILMAZ! Cold-Mix giysi yakasına 1-2 damla damlatılır.",
		"Günde 2-3 kez yakaya/yastığa 1-2 damla; Sterimar izotonik deniz suyu burun içine sınırsız güvenle sıkılabilir.",
		"Burun/göz çevresine temas ederse yakıcı tahriş; küçük çocuklarda mentol/kafur solunum spazmı yapabilir (<2 yaş Vicks kullanılmaz).",
		"Astım atağı eğilimli çocuklarda keskin mentol kokusu öksürüğü tetikleyebilir — yerine izotonik deniz suyu (Sterimar/Tonimer) güvenlidir.",
	),
	_d(
		"Buscopan (Hiyosin-N-Butilbromür - Karın Spazmı)",
		"Hiyosin-N-Butilbromür",
		"Buscopan 10 mg, Buscopan Plus, Spazmol",
		"Antispazmodik (Mide-Bağırsak & Adet Krampı)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Mide-bağırsak krampları, kolik tarzı karın ağrısı ve adet sancısında düz kas spazmını çözer.",
		"Tablet bütün olarak suyla yutulur. ÖNEMLİ: Apandisit şüphesi (sağ alt karın ağrısı + ateş + kusma) varsa tanı maskelenmemesi için ağrı kesici/spazmolitik VERİLMEZ!",
		"6 yaş üzeri ve yetişkinlerde genellikle günde 3-5 kez 10-20 mg; günlük max 100 mg/gün (6 yaş altı tablet kullanılmaz).",
		"Ağız kuruluğu, çarpıntı, kabızlık, geçici görme bulanıklığı.",
		"Mekanik bağırsak tıkanıklığı, megakolon, myastenia gravis, dar açılı glokom ve akut cerrahi karın/apandisit şüphesinde kontrendikedir.",
	),
	_d(
		"Imigran / Relpax / Zomig (Migren Atağı İlacı)",
		"Sumatriptan / Eletriptan / Zolmitriptan",
		"Imigran 50/100 mg, Sumatran, Relpax, Zomig",
		"Triptan (Akut Migren Tedavisi)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin personel veya doktor tanılı ergenlerde akut migren baş ağrısı atağını başlangıcında durdurmak için kullanılır.",
		"Migren atağı başladığı anda mümkün olan en erken sürede suyla yutulur (rutin baş ağrısında kullanılmaz!).",
		"Sumatriptan yetişkin tek doz 50-100 mg; 2 saatte geçmezse toplam max 200-300 mg/24 saat.",
		"Boyun/göğüste geçici sıkışma/basınç hissi, sıcaklık basması, baş dönmesi.",
		"SSRI/SNRI antidepresanlarla (serotonin sendromu riski!), ergotamin türevleriyle (24 saat içinde aynı anda alınmaz!) ve kontrolsüz kalp/hipertansiyon hastalarında kontrendikedir.",
	),
	_d(
		"Dexpass / Dexofen Plus (Ağrı Kesici + Parasetamol)",
		"Deksketoprofen + Parasetamol",
		"Dexpass, Elektra, Arveles + Parol kombinasyonları",
		"Kombine Analjezik (NSAİİ + Parasetamol)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkinlerde şiddetli diş, kas-eklem ve adet ağrılarının kısa süreli tedavisinde iki farklı etken maddeyi tek tablette birleştirir.",
		"Mideyi korumak için mutlaka TOK KARNINA bol su ile yutulur.",
		"Yetişkinlerde günde en fazla 3 tablet (8 saatte bir, max 75 mg deksketoprofen + 900-1500 mg parasetamol); 18 yaş altında kullanılmaz.",
		"Mide yanması, bulantı, uyuşukluk, alerjik reaksiyon.",
		"İçeriğindeki hem PARASETAMOL hem de NSAİİ nedeniyle ek bir Parol veya Arveles/Dolven ile KESİNLİKLE birlikte alınmamalıdır!",
	),
	_d(
		"Bricanyl (Terbütalin)",
		"Terbütalin Sülfat",
		"Bricanyl Şurup, Bricanyl Turbuhaler",
		"Bronkodilatör (Beta-2 Agonist)",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Astım ve bronşitlerde bronşları genişleterek nefes darlığını ve hırıltılı öksürüğü azaltır.",
		"Şurup formu ağızdan; Turbuhaler formu ağızdan derin nefesle çekilir (sonrası durulanır).",
		"Çocuklarda 0.075 mg/kg günde 3 kez; yetişkinlerde günde max 15 mg (şurup).",
		"Ellerde hafif titreme, kalp çarpıntısı, huzursuzluk.",
		"Seçici olmayan beta blokerlerle (Beloc yüksek doz veya Propranolol/Dideral) ve tiroid krizi durumunda dikkat.",
	),
	_d(
		"Cipro / Tavanic (Kinolon Antibiyotik - Yetişkin)",
		"Siprofloksasin / Levofloksasin",
		"Cipro 500 mg, Ciprasid, Tavanic 500, Avelox",
		"Florokinolon Grubu Antibiyotik (Yetişkin)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkinlerde komplike idrar yolu, solunum yolu ve karın enfeksiyonlarında hekim reçetesiyle kullanılır. ÇOCUKLARDA TENDON/EKLEM RİSKİ NEDENİYLE RUTİN KULLANILMAZ!",
		"Bol su ile alınır. SÜT, YOĞURT, KALSİYUM, DEMİR VE MİDE İLAÇLARIYLA (Rennie/Gaviscon) EN AZ 2 SAAT ARA İLE ALINMALIDIR.",
		"Yetişkin max: Siprofloksasin 1500 mg/gün; Levofloksasin 750-1000 mg/gün.",
		"Mide bulantısı, güneş ışığına duyarlılık, nadir ama kritik AŞİL TENDONU AĞRISI / KOPMA RİSKİ (ağrı olursa spor/yürüyüş durdurulur!).",
		"Kortizonlu ilaçlarla birlikte tendon hasarı riskini artırır; antiasit/demir emilimi bozar; teofilin ve tizanidinle ciddi toksisite yapar.",
	),
	_d(
		"Zovirax / Asiviral (Uçuk & Suçiçeği Antivirali)",
		"Asiklovir",
		"Zovirax Krem/Tablet, Asiviral Krem/Şurup, Virosil",
		"Antiviral (Herpes / Varisella)",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Dudak uçuğu (herpes labialis) kremi veya suçiçeği/zona için hekimce verilen oral antiviraldir.",
		"Krem formu uçuğa temas etmeden pamuklu çubukla/eldivenle sürülür ve hemen el yıkanır (göze bulaştırılmaz!). Suçiçeğinde okul izolasyon kuralları da uygulanır.",
		"Krem: Günde 5 kez (4 saatte bir); oral tablet/şurup çocuk max 80 mg/kg/gün (genelde max 3200 mg/gün).",
		"Kremde hafif yanma/kuruma; oral formda bol su içilmezse böbrek yüklenmesi.",
		"Probenesid ve böbreğe toksik ilaçlarla oral form etkileşir.",
	),
	_d(
		"Otrivine / Serum Fizyolojik (Burun Yıkama %0.9 NaCl)",
		"Sodyum Klorür %0.9 İzotonik",
		"Gifrer / Sterimar / Cold-Away / Tonimer / İzotonik Flakon",
		"İzotonik Nazal ve Yara Yıkama Solüsyonu",
		"Burun Spreyi / Damla",
		"Burun Spreyi",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Burun tıkanıklığını güvenli şekilde açmak, göz/yüz/sıyrık temizliği ve steril yara yıkaması yapmak için en güvenli temel okul reviri malzemesidir.",
		"Burun yıkamada flakon veya sprey ucu buruna sıkılır; yara temizliğinde akıtılarak mekanik temizlik yapılır.",
		"İlaç içermez — günde ihtiyaç duyuldukça güvenle kullanılabilir.",
		"Yan etkisi yok denecek kadar azdır.",
		"Bilinen ilaç etkileşimi yoktur; tek kullanımlık plastik ampuller açıldıktan sonra aynı gün tüketilir.",
	),
	_d(
		"Bactrim / Kemoprim (Ko-Trimoksazol)",
		"Sülfametoksazol + Trimetoprim",
		"Bactrim Süspansiyon/Forte, Kemoprim, Bakton",
		"Sülfonamid Antibiyotik",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"İdrar yolu enfeksiyonu ve bazı özel solunum enfeksiyonlarında kullanılır.",
		"Mide konforu için tok karnına ve gün boyu BOL SU ile içilmelidir.",
		"Çocuklarda trimetoprim bileşenine göre 6-12 mg/kg/gün; yetişkin max 320 mg TMP / 1600 mg SMX/gün.",
		"Mide bulantısı, güneşe hassasiyet, SÜLFONAMİD ALERJİSİ (ciltte kabarcık/döküntü olursa derhal kesilip sevk edilmeli!).",
		"Varfarin, metotreksat, fenitoin ve G6PD enzim eksikliğinde kontrendikedir.",
	),
	_d(
		"Lustral / Prozac / Selectra (SSRI Antidepresan Notu)",
		"Sertralin / Fluoksetin",
		"Lustral 50/100, Selectra, Prozac 20, Fulsac",
		"SSRI Antidepresan (Ergen & Yetişkin)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Depresyon, yaygın anksiyete, OKB ve panik bozukluğu tedavisinde çocuk/ergen psikiyatrisi veya psikiyatri takibinde kullanılır.",
		"Genellikle günde 1 kez sabah tok karnına ağızdan alınır; ani bırakılmaz.",
		"Sertralin max 200 mg/gün; Fluoksetin max 60-80 mg/gün (reçeteye göre).",
		"Başlangıçta hafif bulantı, baş ağrısı, uyku değişikliği; ergenlerde ani davranış/duygudurum değişikliği aile ve hekime bildirilmelidir.",
		"MAO inhibitörleri, Triptanlar (migren ilaçları) ve Tramadol ile SEROTONİN SENDROMU riski! NSAİİ ağrı kesicilerle birlikte mide kanama riski artar.",
	),
	_d(
		"Dideral (Propranolol)",
		"Propranolol HCl",
		"Dideral 40 mg Tablet",
		"Seçici Olmayan Beta Bloker",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Çarpıntı, migren profilaksisi, esansiyel tremor ve performans anksiyetesi tedavisinde kullanılır.",
		"Ağızdan suyla yutulur; astım öyküsü olan öğrencide doktor onayı mutlaka sorgulanmalıdır.",
		"Yetişkin günlük 40-320 mg/gün; ani olarak kesilmez.",
		"Nabız düşmesi (bradikardi), tansiyon düşüklüğü, soğuk el-ayak, halsizlik.",
		"ASTIM / BRONŞİTTE KESİNLİKLE KONTRENDİKEDİR (hava yolunu daraltır ve Ventolin etkisini bloke eder!). Diyabetlilerde hipoglisemi titremesini maskeler.",
	),
	_d(
		"Momecon / Dermovate (Dermatolojik Kortizon)",
		"Mometazon / Klobetasol",
		"Momecon Krem/Losyon, Dermovate",
		"Güçlü Topikal Kortikosteroid",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Şiddetli egzama, sedef ve alerjik dermatit plaklarının kısa süreli tedavisinde kullanılır.",
		"Günde 1 kez çok ince sürüp üzeri kapatılmadan açık bırakılır; yüz ve çocuk bez bölgesine sürülmez.",
		"Kısa süreli (max 1-2 hafta) küçük alana lokal kullanım.",
		"Ciltte incelme, çatlaklar, lokal kıllanma.",
		"Enfekte açık yaralar, mantar/suçiçeği lezyonları.",
	),
	_d(
		"Apikobal / Benexol B12 (B Vitamini Kompleksi)",
		"B1 + B6 + B12 Vitaminleri",
		"Apikobal, Benexol B12, Dodex, Neurobion",
		"B Vitamini Kompleksi (Yetişkin & Ergen)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"B vitamini eksiklikleri, sinir ağrıları (nevralji) ve yorgunluk/kansızlık destek tedavisinde kullanılır.",
		"Büyük film tabletler tok karnına bol suyla bütün yutulur.",
		"Günde 1 tablet (nadiren hekim önerisiyle 2).",
		"İdrarın parlak sarı renge dönmesi (B2/riboflavin veya boya nedeniyle normal); çok yüksek B6 dozunda karıncalanma.",
		"Levodopa (Parkinson ilacı) etkisini B6 vitamini azaltabilir.",
	),
	_d(
		"Diclomec / Voltaren Emulgel (Ağrı Kesici Jel)",
		"Diklofenak Dietilamonyum",
		"Voltaren Emulgel, Diclomec Jel, Dikloron Jel",
		"Topikal NSAİİ Jel",
		"Krem / Merhem / Jel",
		"Cilt / Dermal (Merhem/Krem)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Okulda burkulma, incinme, diz/bilek darbesi ve kas ağrılarında şişliği ve lokal ağrıyı azaltır.",
		"Derisi sağlam (açık yara/kesik olmayan!) ağrılı eklem veya kasa haricen sürülür; uygulama sonrası eller yıkanır ve bandajla sıkı sarılmaz.",
		"Günde 3-4 kez fındık büyüklüğünde ince uygulama (14 yaş altı çocuklarda doktor önerisiyle).",
		"Lokal kızarıklık, kaşıntı veya güneşe maruz kalınca cilt hassasiyeti.",
		"Açık yaraya, göze sürülmez; aspirin/NSAİİ alerjisi olan astımlılarda kullanılmaz.",
	),
	_d(
		"Serevent / Seretide / Symbicort (Astım Kombine İnhaler)",
		"Salmeterol+Flutikazon / Formoterol+Budesonid",
		"Seretide Diskus/Evohaler, Symbicort Turbuhaler, Foster",
		"Uzun Etkili Bronkodilatör + Kortikosteroid İnhaler",
		"İnhaler / Nebül",
		"Solunum (İnhalasyon / Hava)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Orta-ağır persistan astımlı öğrencilerde düzenli kontrol edici inhaler olarak kullanılır.",
		"İnhale edilir ve HER UYGULAMA SONRASI AĞIZ BOL SU İLE ÇALKALANIP TÜKÜRÜLÜR. (Not: Yalnızca Symbicort/MART protokolü olan hastalar hekim planında kurtarıcı da kullanabilir; klasik Seretide acil kriz için değildir.)",
		"Hekim reçetesine göre sabah-akşam 1-2 inhalasyon (Symbicort günde max 8-12 inhalasyon erişkin).",
		"Ses kısıklığı, pamukçuk, hafif çarpıntı veya kas titremesi.",
		"Beta blokerler ve güçlü CYP3A4 inhibitörleriyle (ritonavir, ketokonazol) kullanılmaz.",
	),
	_d(
		"Polivit / Supradyn / Zinco (Çocuk Multivitamin & Çinko)",
		"Multivitamin / Çinko Sülfat",
		"Polivit Şurup, Zinco 15 mg Şurup, Supradyn Kids, Sambucol",
		"Vitamin ve Mineral Takviyesi",
		"Şurup / Süspansiyon",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Büyüme çağında vitamin/mineral desteği, iştahsızlık ve özellikle Zinco formunda akut çocukluk ishalinin süresini kısaltmada kullanılır.",
		"Çinko şurupları aç karnına mide bulantısı yapabildiğinden genellikle TOK KARNINA verilir.",
		"Yaşa göre günde 1 ölçek (ishalde çinko 10-20 mg/gün, 10-14 gün).",
		"Aç alınırsa geçici mide bulantısı.",
		"Tetrasiklin/kinolon antibiyotikler ve demirle emilim rekabeti olmaması için 2 saat arayla verilmesi uygundur.",
	),
	_d(
		"Katarin Fort / Theraflu Forte (Yetişkin Soğuk Algınlığı)",
		"Parasetamol + Fenilefrin + Klorfeniramin / Oksolamin",
		"Theraflu Forte, Katarin Forte, Tylol Hot",
		"Yetişkin Grip Kombinasyonu",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Tok karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin öğretmen ve personelde ateş, vücut kırgınlığı ve burun tıkanıklığıyla seyreden soğuk algınlığında kullanılır.",
		"Tok karnına bol suyla alınır; araç/makine veya yoğun ders dikkati gerektiren durumlarda sedasyon (uyku hali) yapabileceği unutulmamalıdır.",
		"Yetişkin max günde 3-4 doz (toplam Parasetamol ≤4000 mg/gün). 12 yaş altına verilmez.",
		"Uyku hali, baş dönmesi, tansiyon yükselmesi, ağız kuruluğu.",
		"Başka Parasetamol ilaçlarıyla, MAO inhibitörleriyle, hipertansiyon ve kalp ilaçlarıyla kombine edilmemelidir.",
	),
	_d(
		"Monurol (Fosfomisin - Tek Doz İYE Saşe)",
		"Fosfomisin Trometamol 3 g",
		"Monurol 3 g Saşe, Üromisin",
		"İdrar Yolu Antibiyotiği",
		"Efervesan / Saşe",
		"Ağızdan (Oral)",
		"Aç karnına",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkin kadınlarda ve 12+ yaş kız öğrencilerde akut komplike olmayan sistit (idrar yolu yanması/enfeksiyonu) tedavisinde tek doz olarak kullanılır.",
		"Mesane boşaltıldıktan sonra gece yatmadan önce, AÇ KARNINA (yemekten 2-3 saat sonra) yarım bardak oda sıcaklığında suda karıştırılıp içilir.",
		"Tek doz 3 g saşe (hekim uygun görmedikçe tekrarlanmaz). 12 yaş altında standart formu kullanılmaz.",
		"Geçici hafif ishal veya karın ağrısı.",
		"Metoklopramid (Metpamid) fosfomisin idrar konsantrasyonunu düşürür, birlikte alınmaz.",
	),
	_d(
		"Loperamid / Lopermid (Yetişkin İshal Durdurucu Uyarısı)",
		"Loperamid HCl",
		"Lopermid 2 mg, Diareks",
		"Antimotilite İshal İlacı (Yetişkin)",
		"Tablet / Kapsül",
		"Ağızdan (Oral)",
		"Fark etmez",
		"Oda Sıcaklığı (15-25 °C)",
		"Yetişkinlerde ateşsiz ve kansız akut basit ishalde bağırsak hareketlerini yavaşlatır. ÇOCUKLARDA VE ATEŞLİ/KANLI İSHALDE KESİNLİKLE VERİLMEZ!",
		"Yalnızca yetişkin personele her gevşek dışkı sonrası bol sıvıyla uygulanır.",
		"Yetişkin başlangıç 4 mg, günde max 16 mg (8 kapsül). 12 yaş altı çocuklarda okulda KESİNLİKLE kullanılmaz!",
		"Kabızlık, karın krampları, şişkinlik; enfeksiyöz ishalde toksinlerin bağırsakta kalmasına yol açar.",
		"Ateşli/kanlı ishalde, bakteri toksini şüphesinde ve küçük çocuklarda kesinlikle kontrendikedir (bunun yerine ORS/Ge-Oral + Reflor tercih edilir).",
	),
]


def ensure_prospectus_tables(connection: sqlite3.Connection) -> None:
	"""İlaç prospektüsü veri bankası tablosunu oluşturur ve eksik yerleşik piyasa ilaçlarını tohumlar."""
	connection.execute(
		"""CREATE TABLE IF NOT EXISTS medication_prospectus (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			medication_name TEXT NOT NULL,
			generic_name TEXT,
			brand_examples TEXT,
			therapeutic_class TEXT,
			dosage_form TEXT,
			route_default TEXT DEFAULT 'Ağızdan (Oral)',
			meal_relation_default TEXT DEFAULT 'Tok karnına',
			storage_condition_default TEXT DEFAULT 'Oda Sıcaklığı (15-25 °C)',
			general_purpose TEXT,
			administration_and_meal TEXT,
			max_daily_dose TEXT,
			side_effects TEXT,
			contraindicated_drugs TEXT,
			pediatric_notes TEXT,
			full_prospectus_notes TEXT,
			source_type TEXT DEFAULT 'builtin',
			ai_enriched INTEGER NOT NULL DEFAULT 0,
			normalized_key TEXT UNIQUE NOT NULL,
			created_at TEXT NOT NULL,
			updated_at TEXT NOT NULL
		)"""
	)
	cols = {row["name"] for row in connection.execute("PRAGMA table_info(medication_prospectus)")}
	migrations = {
		"generic_name": "TEXT",
		"brand_examples": "TEXT",
		"therapeutic_class": "TEXT",
		"dosage_form": "TEXT",
		"route_default": "TEXT DEFAULT 'Ağızdan (Oral)'",
		"meal_relation_default": "TEXT DEFAULT 'Tok karnına'",
		"storage_condition_default": "TEXT DEFAULT 'Oda Sıcaklığı (15-25 °C)'",
		"general_purpose": "TEXT",
		"administration_and_meal": "TEXT",
		"max_daily_dose": "TEXT",
		"side_effects": "TEXT",
		"contraindicated_drugs": "TEXT",
		"pediatric_notes": "TEXT",
		"full_prospectus_notes": "TEXT",
		"source_type": "TEXT DEFAULT 'builtin'",
		"ai_enriched": "INTEGER NOT NULL DEFAULT 0",
		"normalized_key": "TEXT",
		"created_at": "TEXT",
		"updated_at": "TEXT",
	}
	for cname, cdef in migrations.items():
		if cname not in cols:
			connection.execute(f"ALTER TABLE medication_prospectus ADD COLUMN {cname} {cdef}")
	connection.execute(
		"CREATE INDEX IF NOT EXISTS idx_med_prosp_norm_key ON medication_prospectus(normalized_key)"
	)
	seed_default_prospectus_if_needed(connection)


def seed_default_prospectus_if_needed(connection: sqlite3.Connection, force_missing: bool = False) -> int:
	"""Yerleşik Türkiye piyasası genel ilaç veri bankasını idempotent olarak veritabanına yükler."""
	now_iso = datetime.now().isoformat(timespec="seconds")
	cur_cnt = connection.execute("SELECT COUNT(*) FROM medication_prospectus").fetchone()[0]
	seed_row = None
	try:
		seed_row = connection.execute(
			"SELECT value FROM app_settings WHERE key='med_prospectus_builtin_seed_ver'"
		).fetchone()
	except sqlite3.OperationalError:
		seed_row = None
	seed_ver = str(seed_row[0]) if seed_row else ""

	if not force_missing and cur_cnt > 0 and seed_ver == "1.0":
		return 0

	inserted = 0
	for item in BUILTIN_TR_MEDICATION_BANK:
		nkey = _make_norm_key(item["medication_name"], item.get("generic_name", ""))
		exists = connection.execute(
			"SELECT 1 FROM medication_prospectus WHERE normalized_key = ?",
			(nkey,),
		).fetchone()
		if exists:
			continue
		connection.execute(
			"""INSERT INTO medication_prospectus(
				medication_name, generic_name, brand_examples, therapeutic_class, dosage_form,
				route_default, meal_relation_default, storage_condition_default,
				general_purpose, administration_and_meal, max_daily_dose, side_effects,
				contraindicated_drugs, pediatric_notes, full_prospectus_notes,
				source_type, ai_enriched, normalized_key, created_at, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'builtin', 0, ?, ?, ?)""",
			(
				item["medication_name"],
				item["generic_name"],
				item["brand_examples"],
				item["therapeutic_class"],
				item["dosage_form"],
				item["route_default"],
				item["meal_relation_default"],
				item["storage_condition_default"],
				item["general_purpose"],
				item["administration_and_meal"],
				item["max_daily_dose"],
				item["side_effects"],
				item["contraindicated_drugs"],
				item.get("pediatric_notes", ""),
				item.get("full_prospectus_notes", ""),
				nkey,
				now_iso,
				now_iso,
			),
		)
		inserted += 1

	try:
		has_settings = connection.execute(
			"SELECT 1 FROM sqlite_master WHERE type='table' AND name='app_settings'"
		).fetchone()
		if has_settings:
			connection.execute(
				"INSERT INTO app_settings(key, value) VALUES ('med_prospectus_builtin_seed_ver', '1.0') "
				"ON CONFLICT(key) DO UPDATE SET value=excluded.value"
			)
	except sqlite3.OperationalError:
		pass
	return inserted


def _resolve_conn(connect_db_fn: Callable[[], Any]) -> Any:
	return connect_db_fn()


def list_prospectus_records(
	connect_db_fn: Callable[[], Any],
	*,
	search_text: str = "",
	therapeutic_class: str = "",
	source_filter: str = "",
) -> list[dict[str, Any]]:
	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		rows = conn.execute(
			"SELECT * FROM medication_prospectus ORDER BY CASE WHEN source_type='builtin' THEN 1 ELSE 0 END, medication_name ASC"
		).fetchall()
	items = [dict(r) for r in rows]

	if therapeutic_class and therapeutic_class != "Tümü":
		items = [x for x in items if str(x.get("therapeutic_class") or "") == therapeutic_class]

	if source_filter == "Yerleşik Piyasa Veri Bankası":
		items = [x for x in items if x.get("source_type") == "builtin"]
	elif source_filter == "Kuruma Özel / Eklenen":
		items = [x for x in items if x.get("source_type") in ("custom", "ai", "custom_ai")]
	elif source_filter == "AI Destekli / Zenginleştirilmiş":
		items = [x for x in items if int(x.get("ai_enriched") or 0) == 1 or "ai" in str(x.get("source_type") or "")]

	q_fold = _fold_med_text(search_text)
	if q_fold:
		q_tokens = _extract_core_med_tokens(search_text)
		filtered: list[tuple[int, dict[str, Any]]] = []
		for it in items:
			hay = " ".join(
				_fold_med_text(it.get(k))
				for k in (
					"medication_name",
					"generic_name",
					"brand_examples",
					"therapeutic_class",
					"general_purpose",
					"contraindicated_drugs",
				)
			)
			if q_fold in hay:
				score = 100 if _fold_med_text(it.get("medication_name")).startswith(q_fold) else 60
				filtered.append((score, it))
			elif q_tokens and q_tokens.issubset(set(hay.split())):
				filtered.append((40, it))
			elif q_tokens and any(t in hay for t in q_tokens):
				filtered.append((20, it))
		filtered.sort(key=lambda pair: (-pair[0], ui.tr_alpha_sort_key(pair[1].get("medication_name"))))
		items = [pair[1] for pair in filtered]
	else:
		items.sort(key=lambda x: ui.tr_alpha_sort_key(x.get("medication_name")))
	return items


def get_prospectus_by_id(connect_db_fn: Callable[[], Any], prospectus_id: int) -> dict[str, Any] | None:
	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		row = conn.execute(
			"SELECT * FROM medication_prospectus WHERE id = ?",
			(int(prospectus_id),),
		).fetchone()
	return dict(row) if row else None


def get_prospectus_stats(connect_db_fn: Callable[[], Any]) -> dict[str, int]:
	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		tot = conn.execute("SELECT COUNT(*) FROM medication_prospectus").fetchone()[0]
		builtin = conn.execute("SELECT COUNT(*) FROM medication_prospectus WHERE source_type='builtin'").fetchone()[0]
		custom = conn.execute("SELECT COUNT(*) FROM medication_prospectus WHERE source_type != 'builtin'").fetchone()[0]
		ai_cnt = conn.execute("SELECT COUNT(*) FROM medication_prospectus WHERE ai_enriched=1 OR source_type LIKE '%ai%'").fetchone()[0]
	return {"total": int(tot), "builtin": int(builtin), "custom": int(custom), "ai_enriched": int(ai_cnt)}


def match_medication_prospectus(
	connect_db_fn: Callable[[], Any],
	query_name: str,
) -> dict[str, Any] | None:
	"""Öğrencinin kayıtlı serbest ilaç adını (örn. 'Augmentin-BID 400 mg Şurup') yerel prospektüs kaydıyla akıllı eşleştirir."""
	q_fold = _fold_med_text(query_name)
	if not q_fold:
		return None
	q_tokens = _extract_core_med_tokens(query_name)

	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		rows = conn.execute("SELECT * FROM medication_prospectus").fetchall()

	best_score = 0
	best_item: dict[str, Any] | None = None
	for r in rows:
		it = dict(r)
		nkey_fold = str(it.get("normalized_key") or "").replace("_", " ")
		mname_fold = _fold_med_text(it.get("medication_name"))
		gen_fold = _fold_med_text(it.get("generic_name"))
		brands_list = [_fold_med_text(b) for b in str(it.get("brand_examples") or "").split(",") if _fold_med_text(b)]

		score = 0
		if q_fold == mname_fold or q_fold == nkey_fold or q_fold == gen_fold:
			score = 100
		elif any(q_fold == b for b in brands_list):
			score = 95
		elif q_fold in mname_fold or mname_fold in q_fold:
			score = 80
		elif any((len(b) >= 4 and (b in q_fold or q_fold in b)) for b in brands_list):
			score = 75
		elif len(gen_fold) >= 4 and gen_fold in q_fold:
			score = 70
		else:
			item_tokens = (
				_extract_core_med_tokens(it.get("medication_name"))
				| _extract_core_med_tokens(it.get("generic_name"))
				| _extract_core_med_tokens(it.get("brand_examples"))
			)
			overlap = q_tokens & item_tokens
			if overlap:
				# Prefix/full token match strength
				score = 45 + min(30, len(overlap) * 18)
			elif q_tokens and item_tokens:
				for qt in q_tokens:
					if len(qt) >= 4 and any(itk.startswith(qt) or qt.startswith(itk) for itk in item_tokens if len(itk) >= 4):
						score = max(score, 40)

		if score > best_score:
			best_score = score
			best_item = it

	return best_item if best_score >= 40 else None


def upsert_prospectus_record(
	connect_db_fn: Callable[[], Any],
	payload: dict[str, Any],
	*,
	prospectus_id: int | None = None,
) -> int:
	med_name = str(payload.get("medication_name") or "").strip()
	if not med_name:
		raise ValueError("İlaç adı boş olamaz.")

	generic_name = str(payload.get("generic_name") or "").strip()
	brand_examples = str(payload.get("brand_examples") or "").strip()
	therapeutic_class = str(payload.get("therapeutic_class") or "Diğer / Özel İlaç").strip()
	dosage_form = str(payload.get("dosage_form") or "Diğer").strip()
	route_default = str(payload.get("route_default") or "Ağızdan (Oral)").strip()
	meal_rel = str(payload.get("meal_relation_default") or "Tok karnına").strip()
	storage_cond = str(payload.get("storage_condition_default") or "Oda Sıcaklığı (15-25 °C)").strip()
	general_purpose = str(payload.get("general_purpose") or "").strip()
	admin_meal = str(payload.get("administration_and_meal") or "").strip()
	max_dose = str(payload.get("max_daily_dose") or "").strip()
	side_effects = str(payload.get("side_effects") or "").strip()
	contraindicated = str(payload.get("contraindicated_drugs") or "").strip()
	pediatric_notes = str(payload.get("pediatric_notes") or "").strip()
	full_notes = str(payload.get("full_prospectus_notes") or "").strip()
	source_type = str(payload.get("source_type") or "custom").strip()
	ai_enriched = 1 if int(payload.get("ai_enriched") or 0) else 0
	now_iso = datetime.now().isoformat(timespec="seconds")

	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		base_key = _make_norm_key(med_name, generic_name)

		if prospectus_id is not None:
			old_row = conn.execute(
				"SELECT id, normalized_key FROM medication_prospectus WHERE id = ?",
				(int(prospectus_id),),
			).fetchone()
			if old_row:
				final_key = base_key
				clash = conn.execute(
					"SELECT id FROM medication_prospectus WHERE normalized_key = ? AND id != ?",
					(final_key, int(prospectus_id)),
				).fetchone()
				if clash:
					final_key = f"{base_key}_{int(prospectus_id)}"
				conn.execute(
					"""UPDATE medication_prospectus SET
						medication_name=?, generic_name=?, brand_examples=?, therapeutic_class=?, dosage_form=?,
						route_default=?, meal_relation_default=?, storage_condition_default=?,
						general_purpose=?, administration_and_meal=?, max_daily_dose=?, side_effects=?,
						contraindicated_drugs=?, pediatric_notes=?, full_prospectus_notes=?,
						source_type=?, ai_enriched=?, normalized_key=?, updated_at=?
					   WHERE id=?""",
					(
						med_name, generic_name, brand_examples, therapeutic_class, dosage_form,
						route_default, meal_rel, storage_cond,
						general_purpose, admin_meal, max_dose, side_effects,
						contraindicated, pediatric_notes, full_notes,
						source_type, ai_enriched, final_key, now_iso, int(prospectus_id),
					),
				)
				return int(prospectus_id)

		existing_same = conn.execute(
			"SELECT id, source_type, ai_enriched FROM medication_prospectus WHERE normalized_key = ?",
			(base_key,),
		).fetchone()
		if existing_same:
			ex_id = int(existing_same["id"])
			keep_src = source_type if existing_same["source_type"] != "builtin" else ("custom_ai" if ai_enriched else "custom")
			conn.execute(
				"""UPDATE medication_prospectus SET
					medication_name=?, generic_name=?, brand_examples=?, therapeutic_class=?, dosage_form=?,
					route_default=?, meal_relation_default=?, storage_condition_default=?,
					general_purpose=?, administration_and_meal=?, max_daily_dose=?, side_effects=?,
					contraindicated_drugs=?, pediatric_notes=?, full_prospectus_notes=?,
					source_type=?, ai_enriched=?, updated_at=?
				   WHERE id=?""",
				(
					med_name, generic_name, brand_examples, therapeutic_class, dosage_form,
					route_default, meal_rel, storage_cond,
					general_purpose, admin_meal, max_dose, side_effects,
					contraindicated, pediatric_notes, full_notes,
					keep_src, max(ai_enriched, int(existing_same["ai_enriched"] or 0)), now_iso, ex_id,
				),
			)
			return ex_id

		cur = conn.execute(
			"""INSERT INTO medication_prospectus(
				medication_name, generic_name, brand_examples, therapeutic_class, dosage_form,
				route_default, meal_relation_default, storage_condition_default,
				general_purpose, administration_and_meal, max_daily_dose, side_effects,
				contraindicated_drugs, pediatric_notes, full_prospectus_notes,
				source_type, ai_enriched, normalized_key, created_at, updated_at
			) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
			(
				med_name, generic_name, brand_examples, therapeutic_class, dosage_form,
				route_default, meal_rel, storage_cond,
				general_purpose, admin_meal, max_dose, side_effects,
				contraindicated, pediatric_notes, full_notes,
				source_type, ai_enriched, base_key, now_iso, now_iso,
			),
		)
		return int(cur.lastrowid)


def delete_prospectus_record(connect_db_fn: Callable[[], Any], prospectus_id: int) -> bool:
	with _resolve_conn(connect_db_fn) as conn:
		ensure_prospectus_tables(conn)
		cur = conn.execute("DELETE FROM medication_prospectus WHERE id = ?", (int(prospectus_id),))
		return cur.rowcount > 0


def build_prospectus_docx(
	record: dict[str, Any],
	*,
	ensure_docx_fn: Callable[[], bool] | None = None,
	app_cfg: dict[str, Any] | None = None,
) -> bytes:
	"""Seçili ilaç prospektüsünün 5 bölümlü resmi revir bilgilendirme Word (.docx) çıktısını üretir."""
	if ensure_docx_fn is not None and not ensure_docx_fn():
		raise RuntimeError("Word belge oluşturucu (python-docx) yüklü değil.")
	import docx
	from docx.enum.text import WD_ALIGN_PARAGRAPH
	from docx.shared import Cm, Pt, RGBColor

	doc = docx.Document()
	sec = doc.sections[0]
	sec.top_margin = Cm(1.4)
	sec.bottom_margin = Cm(1.4)
	sec.left_margin = Cm(1.6)
	sec.right_margin = Cm(1.6)

	cfg = app_cfg or {}
	school_name = str(cfg.get("school_name") or "Okul / Kurum Reviri").strip()
	nurse_name = str(cfg.get("nurse_name") or "Okul Revir Sorumlusu").strip()

	p_hdr = doc.add_paragraph()
	p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
	r_s = p_hdr.add_run(f"{school_name} — REVİR İLAÇ PROSPEKTÜSÜ VE KLİNİK ÖN BİLGİ FİŞİ\n")
	r_s.bold = True
	r_s.font.size = Pt(11)
	r_s.font.color.rgb = RGBColor(15, 118, 110)
	r_t = p_hdr.add_run(str(record.get("medication_name") or "İlaç Prospektüsü"))
	r_t.bold = True
	r_t.font.size = Pt(15)
	r_t.font.color.rgb = RGBColor(15, 23, 42)

	meta_tbl = doc.add_table(rows=4, cols=2)
	meta_tbl.style = "Table Grid"
	meta_rows = [
		("Etken Madde", str(record.get("generic_name") or "—")),
		("Piyasa / Eşdeğer Adları", str(record.get("brand_examples") or "—")),
		("İlaç Grubu / Form", f"{record.get('therapeutic_class') or '—'} | {record.get('dosage_form') or '—'}"),
		("Standart Yol / Aç-Tok / Saklama", f"{record.get('route_default') or '—'} | {record.get('meal_relation_default') or '—'} | {record.get('storage_condition_default') or '—'}"),
	]
	for idx, (k_lbl, v_txt) in enumerate(meta_rows):
		c0 = meta_tbl.cell(idx, 0)
		c1 = meta_tbl.cell(idx, 1)
		c0.text = k_lbl
		c1.text = v_txt
		if c0.paragraphs and c0.paragraphs[0].runs:
			c0.paragraphs[0].runs[0].bold = True

	doc.add_paragraph()
	sections = [
		("1. GENEL KULLANIM AMACI", str(record.get("general_purpose") or "Bilgi girilmemiş.")),
		("2. KULLANIM ŞEKLİ, AÇ/TOK DURUMU VE UYGULAMA YOLU", str(record.get("administration_and_meal") or "Bilgi girilmemiş.")),
		("3. GÜNLÜK ALINABİLECEK MAKSİMUM DOZ", str(record.get("max_daily_dose") or "Bilgi girilmemiş.")),
		("4. OLASI YAN ETKİLER VE ACİL UYARI BELİRTİLERİ", str(record.get("side_effects") or "Bilgi girilmemiş.")),
		("5. BERABER KULLANILMAMASI GEREKEN İLAÇLAR & KONTRENDİKASYONLAR", str(record.get("contraindicated_drugs") or "Bilgi girilmemiş.")),
	]
	if str(record.get("pediatric_notes") or "").strip():
		sections.append(("6. OKUL REVİRİ & PEDİATRİK İZLEM NOTU", str(record["pediatric_notes"])))
	if str(record.get("full_prospectus_notes") or "").strip():
		sections.append(("7. EK PROSPEKTÜS NOTLARI", str(record["full_prospectus_notes"])))

	for title_txt, body_txt in sections:
		p_sec = doc.add_paragraph()
		p_sec.paragraph_format.space_before = Pt(6)
		p_sec.paragraph_format.space_after = Pt(2)
		r_head = p_sec.add_run(title_txt)
		r_head.bold = True
		r_head.font.size = Pt(11)
		r_head.font.color.rgb = RGBColor(13, 148, 136)

		p_body = doc.add_paragraph(body_txt)
		p_body.paragraph_format.space_after = Pt(4)

	p_ft = doc.add_paragraph()
	p_ft.paragraph_format.space_before = Pt(10)
	r_f = p_ft.add_run(f"Düzenleyen: {nurse_name} | {datetime.now().strftime('%d.%m.%Y %H:%M')}\n{PROSPECTUS_DISCLAIMER}")
	r_f.font.size = Pt(8.5)
	r_f.font.color.rgb = RGBColor(100, 116, 139)

	buf = BytesIO()
	doc.save(buf)
	return buf.getvalue()


def render_prospectus_preview_cards(record: dict[str, Any]) -> None:
	"""İlacın 5 zorunlu prospektüs ön gösterim bölümünü renk kodlu klinik kartlar halinde çizer."""
	med_title = html.escape(str(record.get("medication_name") or "İlaç Prospektüsü"))
	gen = html.escape(str(record.get("generic_name") or "Belirtilmedi"))
	brands = html.escape(str(record.get("brand_examples") or "—"))
	t_class = html.escape(str(record.get("therapeutic_class") or "Genel"))
	form = html.escape(str(record.get("dosage_form") or "—"))
	route = html.escape(str(record.get("route_default") or "—"))
	meal = html.escape(str(record.get("meal_relation_default") or "—"))
	storage = html.escape(str(record.get("storage_condition_default") or "—"))

	src = str(record.get("source_type") or "builtin")
	ai_tag = ""
	if int(record.get("ai_enriched") or 0) == 1 or "ai" in src:
		ai_tag = "<span style='background:#ede9fe;color:#5b21b6;border:1px solid #c4b5fd;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;'>✨ AI Destekli</span>"
	elif src != "builtin":
		ai_tag = "<span style='background:#e0f2fe;color:#0369a1;border:1px solid #bae6fd;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;'>📝 Özel Kayıt</span>"
	else:
		ai_tag = "<span style='background:#ccfbf1;color:#0f766e;border:1px solid #99f6e4;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;'>🏛️ Türkiye Piyasa Bankası</span>"

	ui.markdown(
		f"""<div style="background:linear-gradient(135deg,rgba(13,148,136,0.10),rgba(2,132,199,0.06));
		border:1.5px solid rgba(13,148,136,0.32);border-radius:12px;padding:12px 14px;margin-bottom:10px;">
			<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap;">
				<div>
					<div style="font-size:1.15rem;font-weight:800;color:var(--text-heading);">💊 {med_title}</div>
					<div style="font-size:0.86rem;color:var(--text-muted);margin-top:2px;">
						<b>Etken Madde:</b> {gen} &nbsp;|&nbsp; <b>Grup:</b> {t_class} &nbsp;|&nbsp; <b>Form:</b> {form}
					</div>
				</div>
				<div>{ai_tag}</div>
			</div>
			<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;font-size:0.8rem;">
				<span style="background:var(--card-bg,#ffffff);border:1px solid #cbd5e1;border-radius:6px;padding:3px 8px;">💉 <b>Yol:</b> {route}</span>
				<span style="background:var(--card-bg,#ffffff);border:1px solid #cbd5e1;border-radius:6px;padding:3px 8px;">🍽️ <b>Aç/Tok:</b> {meal}</span>
				<span style="background:var(--card-bg,#ffffff);border:1px solid #cbd5e1;border-radius:6px;padding:3px 8px;">❄️ <b>Saklama:</b> {storage}</span>
			</div>
			<div style="font-size:0.8rem;color:var(--text-muted);margin-top:6px;">
				<b>🏷️ Eşdeğer / Piyasa İsimleri:</b> {brands}
			</div>
		</div>""",
		unsafe_allow_html=True,
	)

	card_specs = [
		(
			"🎯 Genel Kullanım Amacı (Ne İşe Yarar?)",
			record.get("general_purpose"),
			"#0d9488",
			"rgba(13,148,136,0.07)",
		),
		(
			"🍽️ Kullanım Şekli, Aç/Tok Durumu & Uygulama Yolu",
			record.get("administration_and_meal"),
			"#0284c7",
			"rgba(2,132,199,0.07)",
		),
		(
			"⚖️ Günlük Alınabilecek Maksimum Doz",
			record.get("max_daily_dose"),
			"#7c3aed",
			"rgba(124,58,237,0.07)",
		),
		(
			"⚠️ Olası Yan Etkiler & Acil Uyarı Bulguları",
			record.get("side_effects"),
			"#d97706",
			"rgba(217,119,6,0.08)",
		),
		(
			"🚫 Beraber Kullanılmaması Gereken İlaçlar & Kontrendikasyonlar",
			record.get("contraindicated_drugs"),
			"#dc2626",
			"rgba(220,38,38,0.08)",
		),
	]

	for c_title, c_body, accent, bg in card_specs:
		body_clean = html.escape(str(c_body or "Bu bölüm henüz doldurulmamış. 'Düzenle' veya 'AI ile Doldur' üzerinden tamamlayabilirsiniz."))
		ui.markdown(
			f"""<div style="background:{bg};border-left:4px solid {accent};border-radius:8px;padding:9px 12px;margin-bottom:8px;">
				<div style="font-size:0.9rem;font-weight:800;color:{accent};margin-bottom:3px;">{html.escape(c_title)}</div>
				<div style="font-size:0.88rem;line-height:1.45;color:var(--text-body,#1e293b);">{body_clean}</div>
			</div>""",
			unsafe_allow_html=True,
		)

	ped_n = str(record.get("pediatric_notes") or "").strip()
	full_n = str(record.get("full_prospectus_notes") or "").strip()
	if ped_n or full_n:
		with ui.expander("👶 Okul Reviri / Pediatrik Bakım & Tam Prospektüs Notları", expanded=False):
			if ped_n:
				ui.markdown(f"**👶 Okul Çağı & Hemşirelik Notu:** {ped_n}")
			if full_n:
				ui.markdown(f"**📋 Ek Prospektüs Özeti:** {full_n}")

	ui.caption(PROSPECTUS_DISCLAIMER)


def _apply_ai_dict_to_editor_state(
	ed_key_prefix: str,
	ai_data: dict[str, str],
	fill_mode: str = "all",
	*,
	initial_selects: dict[str, str] | None = None,
) -> None:
	"""AI sonuçlarını editör widget session_state anahtarlarına aktarır."""
	text_fields = (
		"medication_name",
		"generic_name",
		"brand_examples",
		"therapeutic_class",
		"general_purpose",
		"administration_and_meal",
		"max_daily_dose",
		"side_effects",
		"contraindicated_drugs",
		"pediatric_notes",
		"full_prospectus_notes",
	)
	for fk in text_fields:
		wk = f"{ed_key_prefix}{fk}"
		cur = str(ui.session_state.get(wk) or "").strip()
		new_v = str(ai_data.get(fk) or "").strip()
		if not new_v:
			continue
		if fill_mode == "missing_only" and cur:
			continue
		ui.session_state[wk] = new_v

	init_sel = initial_selects or {}
	for sel_fk, allowed in (
		("dosage_form", DOSAGE_FORM_OPTIONS),
		("route_default", ROUTE_OPTIONS),
		("meal_relation_default", MEAL_OPTIONS),
		("storage_condition_default", STORAGE_OPTIONS),
	):
		wk = f"{ed_key_prefix}{sel_fk}"
		new_v = str(ai_data.get(sel_fk) or "").strip()
		if new_v in allowed:
			cur_sel = str(ui.session_state.get(wk) or "").strip()
			was_untouched_default = not init_sel.get(sel_fk) and cur_sel in ("", allowed[0])
			if fill_mode != "missing_only" or not cur_sel or was_untouched_default:
				ui.session_state[wk] = new_v


@ui.dialog("📝 İlaç & Prospektüs Ekle / Düzenle", width="large")
def open_prospectus_editor_dialog(
	*,
	connect_db_fn: Callable[[], Any],
	gemini_cfg_fn: Callable[[], tuple[bool, str, str]],
	prospectus_id: int | None = None,
	initial_med_name: str = "",
	prefill_ai_data: dict[str, Any] | None = None,
	on_saved_cb: Callable[[int], None] | None = None,
) -> None:
	rec: dict[str, Any] = {}
	if prospectus_id is not None:
		rec = get_prospectus_by_id(connect_db_fn, prospectus_id) or {}
	elif prefill_ai_data and isinstance(prefill_ai_data, dict):
		rec = dict(prefill_ai_data)

	is_edit = bool(prospectus_id and rec.get("id"))
	k = f"prosp_ed_{prospectus_id or 'new'}_"

	init_sig_k = f"{k}init_sig"
	desired_sig = f"{prospectus_id}:{rec.get('updated_at', '')}:{initial_med_name}:{bool(prefill_ai_data)}"
	if ui.session_state.get(init_sig_k) != desired_sig:
		seed_src = dict(rec)
		if initial_med_name and not seed_src.get("medication_name"):
			seed_src["medication_name"] = initial_med_name
		for fk in (
			"medication_name",
			"generic_name",
			"brand_examples",
			"therapeutic_class",
			"dosage_form",
			"route_default",
			"meal_relation_default",
			"storage_condition_default",
			"general_purpose",
			"administration_and_meal",
			"max_daily_dose",
			"side_effects",
			"contraindicated_drugs",
			"pediatric_notes",
			"full_prospectus_notes",
		):
			ui.session_state[f"{k}{fk}"] = str(seed_src.get(fk) or "")
		ui.session_state[f"{k}init_sels"] = {
			sfk: str(seed_src.get(sfk) or "").strip()
			for sfk in ("dosage_form", "route_default", "meal_relation_default", "storage_condition_default")
		}
		ui.session_state[f"{k}ai_enriched_flag"] = 1 if (prefill_ai_data or int(seed_src.get("ai_enriched") or 0) == 1) else 0
		ui.session_state[init_sig_k] = desired_sig

	if ui.session_state.get(f"{k}ai_toast"):
		ui.success(ui.session_state.pop(f"{k}ai_toast"))
	if ui.session_state.get(f"{k}ai_err"):
		ui.error(ui.session_state.pop(f"{k}ai_err"))

	ui.caption("Veri bankasında bulunmayan piyasa veya reçeteli ilaçları manuel ekleyebilir; eksik veya tüm prospektüs bölümlerini Gemini AI'ya yorumlatarak doldurabilirsiniz.")

	r1_c1, r1_c2 = ui.columns([1.8, 1.4])
	with r1_c1:
		med_name = ui.text_input(
			"💊 İlaç Adı ve Formu *",
			value=ui.session_state.get(f"{k}medication_name", ""),
			placeholder="Örn: Klamoks-BID 400 mg Süspansiyon, Dolarex, Prospan",
			key=f"{k}medication_name",
		)
	with r1_c2:
		generic_name = ui.text_input(
			"🧪 Etken Madde(ler)",
			value=ui.session_state.get(f"{k}generic_name", ""),
			placeholder="Örn: Amoksisilin + Klavulanik Asit",
			key=f"{k}generic_name",
		)

	ai_on, api_key, model = gemini_cfg_fn()
	ai_bar_c1, ai_bar_c2 = ui.columns(2)
	with ai_bar_c1:
		if ui.button("✨ Boş Bölümleri Gemini AI ile Yorumlat & Tamamla", type="secondary", use_container_width=True, key=f"{k}btn_ai_missing"):
			if not str(med_name or "").strip():
				ui.session_state[f"{k}ai_err"] = "Lütfen önce ilaç adını yazın."
				ui.rerun(scope="fragment")
			elif not ai_on or not api_key:
				ui.session_state[f"{k}ai_err"] = "Gemini AI kapalı veya API anahtarı kayıtlı değil (Ayarlar > Yapay Zekâ & KVKK)."
				ui.rerun(scope="fragment")
			else:
				init_sels = ui.session_state.get(f"{k}init_sels") or {}
				_def_sel_map = {
					"dosage_form": DOSAGE_FORM_OPTIONS[0],
					"route_default": ROUTE_OPTIONS[0],
					"meal_relation_default": MEAL_OPTIONS[0],
					"storage_condition_default": STORAGE_OPTIONS[0],
				}
				cur_fields: dict[str, Any] = {}
				for fk in (
					"medication_name", "generic_name", "brand_examples", "therapeutic_class",
					"dosage_form", "route_default", "meal_relation_default", "storage_condition_default",
					"general_purpose", "administration_and_meal", "max_daily_dose",
					"side_effects", "contraindicated_drugs", "pediatric_notes", "full_prospectus_notes",
				):
					v_now = str(ui.session_state.get(f"{k}{fk}", "") or "")
					if fk in _def_sel_map and not init_sels.get(fk) and v_now == _def_sel_map[fk]:
						v_now = ""
					cur_fields[fk] = v_now
				with ui.spinner("Gemini AI yalnızca eksik prospektüs bölümlerini yorumlayıp tamamlıyor..."):
					ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
						api_key=api_key,
						model=model,
						medication_name=med_name,
						existing_fields=cur_fields,
						fill_mode="missing_only",
					)
				if ok:
					_apply_ai_dict_to_editor_state(k, parsed, fill_mode="missing_only", initial_selects=init_sels)
					ui.session_state[f"{k}ai_enriched_flag"] = 1
					ui.session_state[f"{k}ai_toast"] = f"✨ Eksik bölümler Gemini AI ({used_m}) tarafından dolduruldu. Gözden geçirip kaydedebilirsiniz."
				else:
					ui.session_state[f"{k}ai_err"] = str(parsed.get("error") or "AI sorgulaması başarısız oldu.")
				ui.rerun(scope="fragment")

	with ai_bar_c2:
		if ui.button("🧠 Tüm Prospektüsü Gemini AI ile Doldur / Yenile", type="primary", use_container_width=True, key=f"{k}btn_ai_all"):
			if not str(med_name or "").strip():
				ui.session_state[f"{k}ai_err"] = "Lütfen önce ilaç adını yazın."
				ui.rerun(scope="fragment")
			elif not ai_on or not api_key:
				ui.session_state[f"{k}ai_err"] = "Gemini AI kapalı veya API anahtarı kayıtlı değil (Ayarlar > Yapay Zekâ & KVKK)."
				ui.rerun(scope="fragment")
			else:
				with ui.spinner("Gemini AI tüm prospektüs bölümlerini hazırlıyor..."):
					ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
						api_key=api_key,
						model=model,
						medication_name=med_name,
						extra_context=f"Etken Madde Notu: {generic_name}" if generic_name.strip() else "",
						fill_mode="all",
					)
				if ok:
					_apply_ai_dict_to_editor_state(k, parsed, fill_mode="all")
					ui.session_state[f"{k}ai_enriched_flag"] = 1
					ui.session_state[f"{k}ai_toast"] = f"🧠 Tüm prospektüs alanları Gemini AI ({used_m}) ile oluşturuldu. Kontrol edip kaydedin."
				else:
					ui.session_state[f"{k}ai_err"] = str(parsed.get("error") or "AI sorgulaması başarısız oldu.")
				ui.rerun(scope="fragment")

	r2_c1, r2_c2 = ui.columns(2)
	with r2_c1:
		brand_examples = ui.text_input(
			"🏷️ Piyasa / Muadil İsimleri",
			value=ui.session_state.get(f"{k}brand_examples", ""),
			placeholder="Örn: Augmentin, Klamoks, Amoklavin",
			key=f"{k}brand_examples",
		)
	with r2_c2:
		therapeutic_class = ui.text_input(
			"🩺 Terapötik Grup",
			value=ui.session_state.get(f"{k}therapeutic_class", ""),
			placeholder="Örn: Analjezik & Antipiretik, Antibiyotik, Bronkodilatör",
			key=f"{k}therapeutic_class",
		)

	r3_c1, r3_c2, r3_c3, r3_c4 = ui.columns(4)
	with r3_c1:
		cur_form = ui.session_state.get(f"{k}dosage_form", "")
		f_opts = list(DOSAGE_FORM_OPTIONS)
		if cur_form and cur_form not in f_opts:
			f_opts.insert(0, cur_form)
		f_idx = f_opts.index(cur_form) if cur_form in f_opts else 0
		dosage_form = ui.selectbox("Form", f_opts, index=f_idx, key=f"{k}dosage_form")
	with r3_c2:
		cur_rt = ui.session_state.get(f"{k}route_default", "Ağızdan (Oral)")
		rt_opts = list(ROUTE_OPTIONS)
		if cur_rt and cur_rt not in rt_opts:
			rt_opts.insert(0, cur_rt)
		rt_idx = rt_opts.index(cur_rt) if cur_rt in rt_opts else 0
		route_default = ui.selectbox("Uygulama Yolu", rt_opts, index=rt_idx, key=f"{k}route_default")
	with r3_c3:
		cur_ml = ui.session_state.get(f"{k}meal_relation_default", "Tok karnına")
		ml_opts = list(MEAL_OPTIONS)
		if cur_ml and cur_ml not in ml_opts:
			ml_opts.insert(0, cur_ml)
		ml_idx = ml_opts.index(cur_ml) if cur_ml in ml_opts else 0
		meal_default = ui.selectbox("Aç/Tok Durumu", ml_opts, index=ml_idx, key=f"{k}meal_relation_default")
	with r3_c4:
		cur_st = ui.session_state.get(f"{k}storage_condition_default", "Oda Sıcaklığı (15-25 °C)")
		st_opts = list(STORAGE_OPTIONS)
		if cur_st and cur_st not in st_opts:
			st_opts.insert(0, cur_st)
		st_idx = st_opts.index(cur_st) if cur_st in st_opts else 0
		storage_default = ui.selectbox("Saklama Koşulu", st_opts, index=st_idx, key=f"{k}storage_condition_default")

	general_purpose = ui.text_area(
		"🎯 1. Genel Kullanım Amacı",
		value=ui.session_state.get(f"{k}general_purpose", ""),
		placeholder="İlacın başlıca kullanım amaçları ve endikasyonları...",
		key=f"{k}general_purpose",
	)
	admin_and_meal = ui.text_area(
		"🍽️ 2. Kullanım Şekli, Aç/Tok Karnına & Uygulama Yolu",
		value=ui.session_state.get(f"{k}administration_and_meal", ""),
		placeholder="Aç mı tok mu, suyla mı alınır, uygulama sonrası dikkat edilecekler...",
		key=f"{k}administration_and_meal",
	)
	max_daily_dose = ui.text_area(
		"⚖️ 3. Günlük Alınabilecek Maksimum Doz",
		value=ui.session_state.get(f"{k}max_daily_dose", ""),
		placeholder="Pediatrik (mg/kg/gün ve tavan) ile yetişkin/ergen günlük maksimum doz sınırı...",
		key=f"{k}max_daily_dose",
	)
	side_effects = ui.text_area(
		"⚠️ 4. Olası Yan Etkiler",
		value=ui.session_state.get(f"{k}side_effects", ""),
		placeholder="Sık görülen hafif yan etkiler ile acil müdahale/sevk gerektiren uyarı belirtileri...",
		key=f"{k}side_effects",
	)
	contraindicated = ui.text_area(
		"🚫 5. Beraber Kullanılmaması Gereken İlaçlar & Kontrendikasyonlar",
		value=ui.session_state.get(f"{k}contraindicated_drugs", ""),
		placeholder="Kritik ilaç etkileşimleri, birlikte verilmemesi gereken gruplar ve kontrendikasyonlar...",
		key=f"{k}contraindicated_drugs",
	)

	with ui.expander("👶 Okul Reviri / Pediatrik Notlar ve Ek Prospektüs", expanded=False):
		ped_notes = ui.text_area(
			"👶 Pediatrik & Okul Reviri İzlem Notu",
			value=ui.session_state.get(f"{k}pediatric_notes", ""),
			key=f"{k}pediatric_notes",
		)
		full_notes = ui.text_area(
			"📋 Ek Prospektüs Notları",
			value=ui.session_state.get(f"{k}full_prospectus_notes", ""),
			key=f"{k}full_prospectus_notes",
		)

	b_c1, b_c2 = ui.columns([2, 1])
	with b_c1:
		save_lbl = "💾 Prospektüs Değişikliklerini Kaydet" if is_edit else "💾 Yeni İlacı & Prospektüsünü Veri Bankasına Kaydet"
		if ui.button(save_lbl, type="primary", use_container_width=True, key=f"{k}btn_save"):
			if not str(med_name or "").strip():
				ui.error("Lütfen ilaç adını girin.")
				return
			is_ai_used = int(ui.session_state.get(f"{k}ai_enriched_flag") or 0)
			if is_edit:
				src_val = str(rec.get("source_type") or "custom")
				if is_ai_used and "ai" not in src_val:
					src_val = "custom_ai" if src_val != "builtin" else "builtin_ai"
			else:
				src_val = "ai" if is_ai_used else "custom"

			saved_id = upsert_prospectus_record(
				connect_db_fn,
				{
					"medication_name": med_name,
					"generic_name": generic_name,
					"brand_examples": brand_examples,
					"therapeutic_class": therapeutic_class,
					"dosage_form": dosage_form,
					"route_default": route_default,
					"meal_relation_default": meal_default,
					"storage_condition_default": storage_default,
					"general_purpose": general_purpose,
					"administration_and_meal": admin_and_meal,
					"max_daily_dose": max_daily_dose,
					"side_effects": side_effects,
					"contraindicated_drugs": contraindicated,
					"pediatric_notes": ped_notes,
					"full_prospectus_notes": full_notes,
					"source_type": src_val,
					"ai_enriched": is_ai_used,
				},
				prospectus_id=int(rec["id"]) if is_edit else None,
			)
			ui.session_state["med_prospectus_selected_id"] = saved_id
			ui.session_state["med_prospectus_toast"] = f"✅ '{med_name.strip()}' prospektüs kaydı veri bankasına kaydedildi."
			ui.session_state.pop(init_sig_k, None)
			if on_saved_cb:
				on_saved_cb(saved_id)
			ui.close_dialog()
			ui.rerun(scope="fragment")
	with b_c2:
		if ui.button("✖ Vazgeç", type="secondary", use_container_width=True, key=f"{k}btn_cancel"):
			ui.session_state.pop(init_sig_k, None)
			ui.close_dialog()
			ui.rerun(scope="fragment")


@ui.dialog("📖 İlaç Prospektüsü & Klinik Ön Gösterim", width="large")
def open_medication_prospectus_dialog(
	medication_name: str,
	*,
	connect_db_fn: Callable[[], Any],
	gemini_cfg_fn: Callable[[], tuple[bool, str, str]],
	ensure_docx_fn: Callable[[], bool] | None = None,
	app_cfg_fn: Callable[[], dict[str, Any]] | None = None,
	person_info: dict[str, Any] | None = None,
	med_record: dict[str, Any] | None = None,
	prospectus_id: int | None = None,
) -> None:
	"""Tıklanan ilacın prospektüsünü gösterir; veri bankasında yoksa AI ile sorgulama ve kaydetme imkânı sunar."""
	clean_query = str(medication_name or "").strip()
	if person_info:
		p_name = str(person_info.get("full_name") or "").strip()
		allergies = str(person_info.get("allergies") or "").strip()
		chronic = str(person_info.get("chronic_conditions") or "").strip()
		notes_bits = []
		if p_name:
			notes_bits.append(f"👤 **{p_name}**")
		if med_record and med_record.get("dosage_pattern"):
			notes_bits.append(f"📅 Tedavi Planı: **{med_record.get('dose_amount', 1)} {med_record.get('dose_unit', '')}** ({med_record.get('dosage_pattern')})")
		if allergies:
			notes_bits.append(f"🚨 **Alerji:** {allergies}")
		if chronic:
			notes_bits.append(f"🏥 **Kronik Durum:** {chronic}")
		if notes_bits:
			if allergies:
				ui.warning(" | ".join(notes_bits))
			else:
				ui.info(" | ".join(notes_bits))

	record = None
	if prospectus_id is not None:
		record = get_prospectus_by_id(connect_db_fn, int(prospectus_id))
	if not record and clean_query:
		record = match_medication_prospectus(connect_db_fn, clean_query)

	if ui.session_state.get("prosp_modal_toast"):
		ui.success(ui.session_state.pop("prosp_modal_toast"))
	if ui.session_state.get("prosp_modal_err"):
		ui.error(ui.session_state.pop("prosp_modal_err"))

	ai_on, api_key, model = gemini_cfg_fn()

	if record:
		if clean_query and _fold_med_text(clean_query) not in _fold_med_text(record.get("medication_name")):
			ui.caption(f"🔎 Tıklanan ilaç ('{clean_query}') veri bankasındaki **{record['medication_name']}** kaydıyla eşleştirildi.")

		render_prospectus_preview_cards(record)

		act1, act2, act3, act4 = ui.columns([1.2, 1.5, 1.3, 0.9])
		with act1:
			if ui.button("✏️ Düzenle", key=f"btn_modal_prosp_edit_{record['id']}", use_container_width=True):
				open_prospectus_editor_dialog(
					connect_db_fn=connect_db_fn,
					gemini_cfg_fn=gemini_cfg_fn,
					prospectus_id=int(record["id"]),
				)
		with act2:
			if ui.button("✨ AI ile Yorumlat / Güncelle", type="secondary", key=f"btn_modal_prosp_ai_{record['id']}", use_container_width=True):
				if not ai_on or not api_key:
					ui.session_state["prosp_modal_err"] = "Gemini AI kapalı veya API anahtarı kayıtlı değil."
					ui.rerun(scope="fragment")
				else:
					with ui.spinner("Gemini AI ilacın prospektüs bilgilerini güncelliyor..."):
						ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
							api_key=api_key,
							model=model,
							medication_name=clean_query or str(record.get("medication_name") or ""),
							existing_fields=record,
							fill_mode="all",
						)
					if ok:
						upsert_prospectus_record(
							connect_db_fn,
							{**parsed, "source_type": "builtin_ai" if record.get("source_type") == "builtin" else "custom_ai", "ai_enriched": 1},
							prospectus_id=int(record["id"]),
						)
						ui.session_state["prosp_modal_toast"] = f"✨ Prospektüs Gemini AI ({used_m}) ile güncellendi ve kaydedildi."
					else:
						ui.session_state["prosp_modal_err"] = str(parsed.get("error") or "AI yanıt veremedi.")
					ui.rerun(scope="fragment")
		with act3:
			cfg = app_cfg_fn() if app_cfg_fn else {}
			try:
				docx_bytes = build_prospectus_docx(record, ensure_docx_fn=ensure_docx_fn, app_cfg=cfg)
				ui.download_button(
					"📄 Word Fişi (.docx)",
					data=docx_bytes,
					file_name=f"Prospektus_{_make_norm_key(record.get('medication_name', 'ilac'))}.docx",
					mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
					key=f"btn_modal_prosp_docx_{record['id']}",
					use_container_width=True,
				)
			except Exception:
				ui.caption("Word çıktısı hazır değil")
		with act4:
			if ui.button("✖ Kapat", key=f"btn_modal_prosp_close_{record['id']}", use_container_width=True):
				ui.close_dialog()
				ui.rerun(scope="fragment")
	else:
		ui.warning(
			f"🔍 **'{clean_query}'** için yerel ilaç veri bankasında doğrudan eşleşen bir prospektüs kaydı henüz bulunamadı.\n\n"
			"Aşağıdan tek tıkla **Gemini AI** ile bu ilacın 5 bölümlü prospektüsünü sorgulayıp veri bankasına kaydedebilir "
			"veya kendi prospektüsünüzü manuel olarak ekleyebilirsiniz."
		)
		draft_key = f"prosp_modal_draft_{_make_norm_key(clean_query)}"
		draft = ui.session_state.get(draft_key)
		if isinstance(draft, dict) and draft.get("medication_name"):
			render_prospectus_preview_cards(draft)
			d1, d2, d3 = ui.columns([1.8, 1.4, 1.0])
			with d1:
				if ui.button("💾 Bu AI Prospektüsünü Veri Bankasına Kaydet", type="primary", key=f"btn_save_modal_draft_{draft_key}", use_container_width=True):
					new_id = upsert_prospectus_record(
						connect_db_fn,
						{**draft, "source_type": "ai", "ai_enriched": 1},
					)
					ui.session_state.pop(draft_key, None)
					ui.session_state["prosp_modal_toast"] = "✅ Prospektüs kalıcı olarak veri bankasına kaydedildi!"
					ui.session_state["med_prospectus_selected_id"] = new_id
					ui.rerun(scope="fragment")
			with d2:
				if ui.button("✏️ Düzenleyerek Kaydet", key=f"btn_ed_modal_draft_{draft_key}", use_container_width=True):
					open_prospectus_editor_dialog(
						connect_db_fn=connect_db_fn,
						gemini_cfg_fn=gemini_cfg_fn,
						initial_med_name=clean_query,
						prefill_ai_data={**draft, "ai_enriched": 1},
					)
			with d3:
				if ui.button("✖ Kapat", key="btn_close_missing_prosp_draft", use_container_width=True):
					ui.close_dialog()
					ui.rerun(scope="fragment")
		else:
			m1, m2, m3 = ui.columns([1.8, 1.6, 1.0])
			with m1:
				if ui.button("✨ Gemini AI ile Prospektüsünü Sorgula & Kaydet", type="primary", key="btn_modal_ai_fetch_new", use_container_width=True):
					if not ai_on or not api_key:
						ui.session_state["prosp_modal_err"] = "Gemini AI kapalı veya API anahtarı tanımlanmamış (Ayarlar > Yapay Zekâ & KVKK)."
						ui.rerun(scope="fragment")
					else:
						with ui.spinner(f"Gemini AI '{clean_query}' ilacının prospektüsünü analiz ediyor..."):
							ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
								api_key=api_key,
								model=model,
								medication_name=clean_query,
								fill_mode="all",
							)
						if ok:
							save_rec = dict(parsed)
							cur_b = str(save_rec.get("brand_examples") or "").strip()
							if clean_query and _fold_med_text(clean_query) not in _fold_med_text(f"{save_rec.get('medication_name', '')} {cur_b}"):
								save_rec["brand_examples"] = f"{clean_query}, {cur_b}".strip(", ")
							new_id = upsert_prospectus_record(
								connect_db_fn,
								{**save_rec, "source_type": "ai", "ai_enriched": 1},
							)
							ui.session_state[draft_key] = save_rec
							ui.session_state["prosp_modal_toast"] = f"✨ '{save_rec.get('medication_name', clean_query)}' Gemini AI ({used_m}) ile sorgulandı ve veri bankasına kaydedildi!"
							ui.session_state["med_prospectus_selected_id"] = new_id
						else:
							ui.session_state["prosp_modal_err"] = str(parsed.get("error") or "AI sorgulaması başarısız oldu.")
						ui.rerun(scope="fragment")
			with m2:
				if ui.button("📝 Manuel Prospektüs Olarak Ekle", type="secondary", key="btn_modal_manual_add_new", use_container_width=True):
					open_prospectus_editor_dialog(
						connect_db_fn=connect_db_fn,
						gemini_cfg_fn=gemini_cfg_fn,
						initial_med_name=clean_query,
					)
			with m3:
				if ui.button("✖ Kapat", key="btn_close_missing_prosp", use_container_width=True):
					ui.close_dialog()
					ui.rerun(scope="fragment")


def render_med_dialog_prospectus_helper(
	med_name: str,
	*,
	connect_db_fn: Callable[[], Any],
	gemini_cfg_fn: Callable[[], tuple[bool, str, str]],
	ensure_docx_fn: Callable[[], bool] | None = None,
	app_cfg_fn: Callable[[], dict[str, Any]] | None = None,
	k_prefix: str = "dlg_new_",
	is_prn: bool = False,
) -> None:
	"""İlaç ekleme/düzenleme penceresi içinde canlı prospektüs eşleştirme ve tek tıkla bilgi aktarma şeridi."""
	clean_name = str(med_name or "").strip()
	if not clean_name:
		with ui.expander("📖 İlaç Veri Bankasından Hızlı İlaç Seç (Opsiyonel)", expanded=False):
			all_recs = list_prospectus_records(connect_db_fn)
			opts = ["— Seçiniz —"] + [f"{r['medication_name']} ({r.get('generic_name') or ''})" for r in all_recs]
			pick_map = {f"{r['medication_name']} ({r.get('generic_name') or ''})": r for r in all_recs}
			sel_pick = ui.selectbox("Kayıtlı Piyasa İlacı Seçin", opts, index=0, key=f"{k_prefix}prosp_quick_pick")
			if sel_pick in pick_map:
				picked = pick_map[sel_pick]
				if ui.button("⬇️ Bu İlacı Forma Aktar & Yol/Aç-Tok Doldur", type="primary", key=f"{k_prefix}btn_apply_quick_pick", use_container_width=True):
					ui.session_state[f"{k_prefix}name"] = picked["medication_name"]
					if picked.get("route_default") in ROUTE_OPTIONS:
						ui.session_state[f"{k_prefix}route"] = picked["route_default"]
					if picked.get("meal_relation_default") in MEAL_OPTIONS:
						ui.session_state[f"{k_prefix}meal_prn" if is_prn else f"{k_prefix}meal"] = picked["meal_relation_default"]
					if picked.get("storage_condition_default") in STORAGE_OPTIONS:
						ui.session_state[f"{k_prefix}storage"] = picked["storage_condition_default"]
					ui.rerun(scope="fragment")
		return

	matched = match_medication_prospectus(connect_db_fn, clean_name)
	if matched:
		hp_c1, hp_c2, hp_c3 = ui.columns([2.2, 1.3, 1.5])
		with hp_c1:
			ui.markdown(
				f"<div style='background:#f0fdfa;border:1px solid #99f6e4;border-radius:8px;padding:6px 10px;font-size:0.84rem;'>"
				f"📖 <b>Prospektüs Eşleşti:</b> {html.escape(str(matched['medication_name']))} "
				f"<span style='color:#0f766e;'>({html.escape(str(matched.get('meal_relation_default') or ''))} · "
				f"{html.escape(str(matched.get('route_default') or ''))})</span></div>",
				unsafe_allow_html=True,
			)
		with hp_c2:
			if ui.button("📖 Prospektüsü Gör", key=f"{k_prefix}btn_see_prosp", use_container_width=True):
				open_medication_prospectus_dialog(
					clean_name,
					connect_db_fn=connect_db_fn,
					gemini_cfg_fn=gemini_cfg_fn,
					ensure_docx_fn=ensure_docx_fn,
					app_cfg_fn=app_cfg_fn,
					prospectus_id=int(matched["id"]),
				)
		with hp_c3:
			if ui.button("⬇️ Yol / Aç-Tok / Saklama Aktar", key=f"{k_prefix}btn_fill_from_prosp", use_container_width=True):
				if matched.get("route_default") in ROUTE_OPTIONS:
					ui.session_state[f"{k_prefix}route"] = matched["route_default"]
				if matched.get("meal_relation_default") in MEAL_OPTIONS:
					ui.session_state[f"{k_prefix}meal_prn" if is_prn else f"{k_prefix}meal"] = matched["meal_relation_default"]
				if matched.get("storage_condition_default") in STORAGE_OPTIONS:
					ui.session_state[f"{k_prefix}storage"] = matched["storage_condition_default"]
				tip = str(matched.get("pediatric_notes") or "").strip()
				if tip and not str(ui.session_state.get(f"{k_prefix}notes") or "").strip():
					ui.session_state[f"{k_prefix}notes"] = tip
				ui.rerun(scope="fragment")
	else:
		u_c1, u_c2 = ui.columns([2.5, 1.5])
		with u_c1:
			ui.caption(f"ℹ️ '{clean_name}' henüz yerel prospektüs bankasında yok.")
		with u_c2:
			if ui.button("✨ AI ile Prospektüs Sorgula / Ekle", key=f"{k_prefix}btn_missing_prosp_open", use_container_width=True):
				open_medication_prospectus_dialog(
					clean_name,
					connect_db_fn=connect_db_fn,
					gemini_cfg_fn=gemini_cfg_fn,
					ensure_docx_fn=ensure_docx_fn,
					app_cfg_fn=app_cfg_fn,
				)


def render_prospectus_tab(
	*,
	connect_db_fn: Callable[[], Any],
	gemini_cfg_fn: Callable[[], tuple[bool, str, str]],
	ensure_docx_fn: Callable[[], bool] | None = None,
	app_cfg_fn: Callable[[], dict[str, Any]] | None = None,
) -> None:
	"""İlaç Takip Sistemi içinde Yapay Zekâ Destekli İlaç Prospektüsü & Piyasa Veri Bankası sekmesini çizer."""
	stats = get_prospectus_stats(connect_db_fn)

	if ui.session_state.get("med_prospectus_toast"):
		ui.success(ui.session_state.pop("med_prospectus_toast"))
	if ui.session_state.get("med_prospectus_err"):
		ui.error(ui.session_state.pop("med_prospectus_err"))

	hdr_l, hdr_r1, hdr_r2 = ui.columns([2.8, 1.3, 1.1])
	with hdr_l:
		ui.markdown("#### 📖 İlaç Prospektüsü Sorgulama & Piyasa Veri Bankası")
		ui.caption(
			f"Toplam **{stats['total']}** ilaç kayıtlı ({stats['builtin']} piyasa veri bankası, "
			f"{stats['custom']} kuruma özel, {stats['ai_enriched']} AI destekli). "
			"Olmayan ilaçları AI ile anlık sorgulayıp veri bankasına kaydedebilirsiniz."
		)
	with hdr_r1:
		if ui.button("➕ Yeni İlaç & Prospektüs Ekle", type="primary", use_container_width=True, key="btn_prosp_add_new"):
			q_prefill = str(ui.session_state.get("med_prosp_search") or "").strip()
			open_prospectus_editor_dialog(
				connect_db_fn=connect_db_fn,
				gemini_cfg_fn=gemini_cfg_fn,
				initial_med_name=q_prefill,
			)
	with hdr_r2:
		if ui.button("🔄 Piyasa Bankasını Tamamla", type="secondary", use_container_width=True, key="btn_prosp_reseed"):
			with _resolve_conn(connect_db_fn) as conn:
				added = seed_default_prospectus_if_needed(conn, force_missing=True)
			if added > 0:
				ui.success(f"Türkiye piyasa veri bankasından {added} eksik ilaç kaydı yüklendi.")
			else:
				ui.info("Yerleşik Türkiye piyasa ilaç veri bankasındaki tüm kayıtlar güncel.")
			ui.rerun()

	all_for_cats = list_prospectus_records(connect_db_fn)
	cat_list = sorted({str(x.get("therapeutic_class") or "").strip() for x in all_for_cats if str(x.get("therapeutic_class") or "").strip()})

	sf1, sf2, sf3 = ui.columns([2.2, 1.3, 1.3])
	with sf1:
		search_q = ui.text_input(
			"🔎 İlaç Adı, Etken Madde veya Kullanım Alanı Ara",
			placeholder="Örn: Calpol, Augmentin, İbuprofen, Epinefrin, Prospan...",
			key="med_prosp_search",
		)
	with sf2:
		sel_cat = ui.selectbox("Terapötik Grup", ["Tümü"] + cat_list, index=0, key="med_prosp_cat")
	with sf3:
		sel_src = ui.selectbox(
			"Kaynak Filtresi",
			["Tümü", "Yerleşik Piyasa Veri Bankası", "Kuruma Özel / Eklenen", "AI Destekli / Zenginleştirilmiş"],
			index=0,
			key="med_prosp_src",
		)

	ai_on, api_key, model = gemini_cfg_fn()

	# Arama metni için anlık Yapay Zekâ Prospektüs Sorgulama Şeridi
	clean_q = str(search_q or "").strip()
	if clean_q:
		ai_q_c1, ai_q_c2 = ui.columns([2.4, 1.4], vertical_alignment="center")
		with ai_q_c1:
			ui.caption(f"✨ **'{clean_q}'** için yapay zekâdan güncel prospektüs ön gösterimi almak isterseniz sağdaki sorgu butonunu kullanabilirsiniz:")
		with ai_q_c2:
			if ui.button(f"✨ '{clean_q[:24]}' AI Prospektüsü Sorgula", type="primary", use_container_width=True, key="btn_prosp_ai_search"):
				if not ai_on or not api_key:
					ui.session_state["med_prospectus_err"] = "Gemini AI kapalı veya API anahtarı tanımlanmamış (Ayarlar > Yapay Zekâ & KVKK)."
					ui.rerun()
				else:
					with ui.spinner(f"Gemini AI '{clean_q}' için 5 bölümlü klinik prospektüsü hazırlıyor..."):
						ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
							api_key=api_key,
							model=model,
							medication_name=clean_q,
							fill_mode="all",
						)
					if ok:
						ui.session_state["med_prosp_ai_preview"] = {"query": clean_q, "model": used_m, "data": parsed}
					else:
						ui.session_state["med_prospectus_err"] = str(parsed.get("error") or "AI sorgulaması başarısız oldu.")
					ui.rerun()

	ai_preview = ui.session_state.get("med_prosp_ai_preview")
	if isinstance(ai_preview, dict) and isinstance(ai_preview.get("data"), dict):
		with ui.container(border=True):
			ai_d = ai_preview["data"]
			ui.markdown(f"##### 🤖 Gemini AI Prospektüs Sorgu Sonucu (`{ai_preview.get('model', 'gemini')}`)")
			render_prospectus_preview_cards({**ai_d, "source_type": "ai", "ai_enriched": 1})
			ap1, ap2, ap3 = ui.columns([1.8, 1.4, 1.0])
			with ap1:
				if ui.button("💾 Veri Bankasına Kaydet", type="primary", key="btn_prosp_save_ai_preview", use_container_width=True):
					new_id = upsert_prospectus_record(
						connect_db_fn,
						{**ai_d, "source_type": "ai", "ai_enriched": 1},
					)
					ui.session_state.pop("med_prosp_ai_preview", None)
					ui.session_state["med_prospectus_selected_id"] = new_id
					ui.session_state["med_prospectus_toast"] = f"✅ '{ai_d.get('medication_name', clean_q)}' kalıcı olarak ilaç veri bankasına kaydedildi."
					ui.rerun()
			with ap2:
				if ui.button("✏️ Düzenleyerek Kaydet", key="btn_prosp_edit_ai_preview", use_container_width=True):
					open_prospectus_editor_dialog(
						connect_db_fn=connect_db_fn,
						gemini_cfg_fn=gemini_cfg_fn,
						initial_med_name=str(ai_d.get("medication_name") or clean_q),
						prefill_ai_data={**ai_d, "ai_enriched": 1},
					)
			with ap3:
				if ui.button("✖ Önizlemeyi Kapat", key="btn_prosp_close_ai_preview", use_container_width=True):
					ui.session_state.pop("med_prosp_ai_preview", None)
					ui.rerun()

	filtered_items = list_prospectus_records(
		connect_db_fn,
		search_text=clean_q,
		therapeutic_class=sel_cat,
		source_filter=sel_src,
	)

	if not filtered_items:
		ui.info("Yerel veri bankasında bu aramayla eşleşen ilaç bulunamadı. Yukarıdaki **'AI Prospektüsü Sorgula'** veya **'➕ Yeni İlaç & Prospektüs Ekle'** butonunu kullanabilirsiniz.")
		return

	sel_id = ui.session_state.get("med_prospectus_selected_id")
	selected_item = next((x for x in filtered_items if int(x["id"]) == int(sel_id or -1)), filtered_items[0])
	ui.session_state["med_prospectus_selected_id"] = int(selected_item["id"])

	col_list, col_detail = ui.columns([1.3, 2.2])
	with col_list:
		ui.markdown(f"**📋 Kayıtlı İlaçlar ({len(filtered_items)})**")
		max_show = min(len(filtered_items), 80)
		for idx, it in enumerate(filtered_items[:max_show]):
			is_sel = int(it["id"]) == int(selected_item["id"])
			src_icon = "✨" if (int(it.get("ai_enriched") or 0) == 1 or "ai" in str(it.get("source_type") or "")) else ("📝" if it.get("source_type") != "builtin" else "💊")
			btn_type = "primary" if is_sel else "secondary"
			lbl = f"{src_icon} {it['medication_name']}"
			if it.get("generic_name"):
				lbl += f" · {it['generic_name'][:28]}"
			if ui.button(lbl, type=btn_type, key=f"btn_prosp_pick_{it['id']}_{idx}", use_container_width=True):
				ui.session_state["med_prospectus_selected_id"] = int(it["id"])
				ui.rerun()
		if len(filtered_items) > max_show:
			ui.caption(f"İlk {max_show} ilaç gösteriliyor; aramayı daraltarak daha hızlı bulabilirsiniz.")

	with col_detail:
		with ui.container(border=True):
			render_prospectus_preview_cards(selected_item)

			da1, da2, da3, da4 = ui.columns([1.1, 1.5, 1.3, 1.0])
			with da1:
				if ui.button("✏️ Düzenle", key=f"btn_tab_prosp_edit_{selected_item['id']}", use_container_width=True):
					open_prospectus_editor_dialog(
						connect_db_fn=connect_db_fn,
						gemini_cfg_fn=gemini_cfg_fn,
						prospectus_id=int(selected_item["id"]),
					)
			with da2:
				if ui.button("✨ AI ile Yorumlat", key=f"btn_tab_prosp_ai_{selected_item['id']}", use_container_width=True):
					if not ai_on or not api_key:
						ui.session_state["med_prospectus_err"] = "Gemini AI kapalı veya API anahtarı kayıtlı değil."
						ui.rerun()
					else:
						with ui.spinner("Gemini AI seçili ilacın prospektüsünü yeniden yorumluyor..."):
							ok, parsed, used_m = gai.lookup_medication_prospectus_with_ai(
								api_key=api_key,
								model=model,
								medication_name=str(selected_item.get("medication_name") or ""),
								existing_fields=selected_item,
								fill_mode="all",
							)
						if ok:
							new_src = "builtin_ai" if selected_item.get("source_type") == "builtin" else "custom_ai"
							upsert_prospectus_record(
								connect_db_fn,
								{**parsed, "source_type": new_src, "ai_enriched": 1},
								prospectus_id=int(selected_item["id"]),
							)
							ui.session_state["med_prospectus_toast"] = f"✨ '{selected_item['medication_name']}' Gemini AI ({used_m}) ile güncellendi."
						else:
							ui.session_state["med_prospectus_err"] = str(parsed.get("error") or "AI yanıt vermedi.")
						ui.rerun()
			with da3:
				cfg = app_cfg_fn() if app_cfg_fn else {}
				try:
					docx_bytes = build_prospectus_docx(selected_item, ensure_docx_fn=ensure_docx_fn, app_cfg=cfg)
					ui.download_button(
						"📄 Word (.docx)",
						data=docx_bytes,
						file_name=f"Prospektus_{_make_norm_key(selected_item.get('medication_name', 'ilac'))}.docx",
						mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
						key=f"btn_tab_prosp_docx_{selected_item['id']}",
						use_container_width=True,
					)
				except Exception:
					ui.caption("Word çıktısı kapalı")
			with da4:
				if ui.button("🗑️ Sil", key=f"btn_tab_prosp_del_{selected_item['id']}", use_container_width=True):
					del_name = selected_item.get("medication_name")
					delete_prospectus_record(connect_db_fn, int(selected_item["id"]))
					ui.session_state.pop("med_prospectus_selected_id", None)
					ui.session_state["med_prospectus_toast"] = f"🗑️ '{del_name}' veri bankasından silindi."
					ui.rerun()
