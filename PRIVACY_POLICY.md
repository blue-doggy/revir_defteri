# Gizlilik Politikası ve Tıbbi Sorumluluk Beyanı / Privacy Policy & Medical Disclaimer

**Uygulama / Application:** Okul Revir Defteri ve Sağlık Takip Sistemi  
**Son Güncelleme / Last Updated:** 2026-11-27

---

## 🇹🇷 Türkçe — Gizlilik Politikası (KVKK Aydınlatma Metni) ve Tıbbi Sorumluluk Reddi

### 1. Yerel Veri Depolama (Bulutsuz Çalışma Prensibi)
**Okul Revir Defteri**, tüm öğrenci, personel, revir muayene, vital bulgu, tarama, ilaç ve diyabet izlem kayıtlarını **yalnızca kullanıcının kendi bilgisayarındaki yerel SQLite veritabanında** (`%LOCALAPPDATA%\RevirDefteri` veya kullanıcı tarafından seçilen yerel klasörde) saklar. Geliştiriciye veya herhangi bir merkezi sunucuya otomatik veri aktarımı, telemetri, reklam takibi ya da analitik gönderimi **yapılmaz**.

### 2. İsteğe Bağlı Yapay Zekâ (Google Gemini API) ve KVKK Maskelemesi
Uygulamadaki yapay zekâ destekli analiz, menü/belge okuma ve klinik özet özellikleri **tamamen isteğe bağlıdır** ve yalnızca kullanıcının kendi Google Gemini API anahtarını girmesi halinde çalışır:
- Yapay zekâ sorgularında 6698 sayılı **Kişisel Verilerin Korunması Kanunu (KVKK)** kapsamında ad-soyad, T.C. kimlik numarası, öğrenci numarası, telefon ve açık adres gibi doğrudan tanımlayıcı veriler gönderim öncesinde otomatik olarak anonimleştirilir/maskelenir.
- API anahtarı yalnızca kullanıcının yerel cihazında saklanır.

### 3. Cihaz İzinleri (Mikrofon ve Yerel Dosya Sistemi)
- **Mikrofon İzni:** Yalnızca kullanıcı sohbet veya not alanında sesli dikte düğmesine bastığında anlık ses kaydı almak için kullanılır; arka planda dinleme yapılmaz.
- **Dosya Sistemi ve Ağ:** Kullanıcının seçtiği Word/Excel/PDF raporlarını dışa aktarmak, şablon/evrak yüklemek ve isteğe bağlı internet radyosu akışlarını oynatmak için kullanılır.

### 4. Veri Güvenliği, Yedekleme ve Silme
Verilerin yedeklenmesi, korunması ve silinmesi tamamen kurumun/kullanıcının kontrolündedir. Kullanıcı dilediği anda uygulama içindeki **Yedekleme & Sıfırlama** bölümünden tüm kayıtları dışa aktarabilir veya kalıcı olarak silebilir; uygulamayı kaldırdığında yerel uygulama önbelleği temizlenir.

### 5. ⚕️ Tıbbi Sorumluluk Reddi (Medical Disclaimer)
> **ÖNEMLİ UYARI:** Bu yazılım **tıbbi bir cihaz (Medical Device) değildir** ve hekim muayenesinin, tıbbi tanının veya reçetenin yerine geçmez. Uygulama içinde yer alan WHO büyüme persantili hesaplamaları, vital bulgu referans aralıkları ve hekim reçeteli İKO/İDF parametrelerine dayalı insülin/diyabet hesaplamaları yalnızca okul sağlığı personeline yönelik **kayıt ve karar destek** amaçlıdır. Her türlü ilaç veya insülin uygulaması öncesinde ilgili hekimin yazılı tedavi planı ve güncel klinik ölçüm doğrulanmalıdır.

### 6. İletişim
Gizlilik politikası ve uygulama hakkında sorularınız için GitHub deposu üzerinden iletişime geçebilirsiniz:  
https://github.com/blue-doggy/revir_defteri/issues

---

## 🌐 English — Privacy Policy & Medical Disclaimer

### 1. Local-First Data Storage (No Cloud Collection)
**Okul Revir Defteri** stores all student, staff, infirmary visit, vital sign, screening, medication, and diabetes tracking records **exclusively on the user's local device** inside a local SQLite database (`%LOCALAPPDATA%\RevirDefteri`). We **do not** collect, transmit, rent, or sell any personal or health data, telemetry, analytics, or advertising identifiers to external servers.

### 2. Optional AI Features (Google Gemini API) & De-Identification
AI-assisted features (document/menu parsing, nutrition estimation, and clinical note summaries) are **strictly opt-in** and require the user to configure their own Google Gemini API key:
- Direct personal identifiers (full names, national ID numbers, student numbers, phone numbers) are automatically masked/de-identified prior to sending prompts when KVKK anonymization is active.
- The user's API key is stored solely on the local device.

### 3. Device Permissions
- **Microphone:** Used strictly on-demand when the user clicks the voice input button; never recorded in the background or stored remotely.
- **Local File System & Network:** Used only to import/export user-selected `.docx`, `.xlsx`, `.pdf`, and backup files, and to stream optional online radio stations selected by the user.

### 4. Data Retention & Deletion
Users retain 100% ownership and control over their local database. All records can be exported or permanently wiped at any time via **Settings > Factory Reset**, or removed by uninstalling the application.

### 5. ⚕️ Medical Disclaimer
> **IMPORTANT:** This software is **not a certified medical device** and does not diagnose, treat, cure, or prevent any disease, nor does it replace professional medical advice or a physician's prescription. Growth percentile charts (WHO), vital sign classifications, and physician-prescribed ICR/ISF bolus calculations are provided strictly as record-keeping and informational decision-support tools for qualified school health personnel. Always verify measurements and follow the attending physician's written care plan before any clinical action.

### 6. Contact
For privacy inquiries or support, please contact us via GitHub Issues:  
https://github.com/blue-doggy/revir_defteri/issues
