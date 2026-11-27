# 🛍️ Microsoft Store Yayınlama Rehberi (Adım Adım)

Bu projede Microsoft Store için gerekli tüm teknik paketleme, ikon üretimi, `AppxManifest.xml` şeması, mağaza güncelleme modu (`MS_STORE_BUILD`) ve gizlilik politikası altyapısı hazır hale getirilmiştir.

---

## 1. Partner Center Kimlik Bilgilerinizi `msix_config.json` Dosyasına Girin

1. [Microsoft Partner Center](https://partner.microsoft.com/dashboard) hesabınıza girin.
2. **Apps and games -> New product -> MSIX or PWA app** adımıyla uygulama adınızı (`Okul Revir Defteri`) rezerve edin.
3. Sol menüden **Product management -> Product identity** sayfasını açın ve oradaki 3 satırı kopyalayıp proje kökündeki `msix_config.json` dosyasına yapıştırın:
   - `Package/Identity/Name` $\rightarrow$ `"identity_name"`
   - `Package/Identity/Publisher` (`CN=...` ile başlar) $\rightarrow$ `"publisher"`
   - `Package/Properties/PublisherDisplayName` $\rightarrow$ `"publisher_display_name"`

> **Not:** Paketi kendi bilgisayarınızda sertifikayla imzalamanıza **gerek yoktur**; Microsoft Store `.msix` dosyasını yüklediğinizde kendi kök sertifikasıyla ücretsiz imzalar.

---

## 2. Tek Komutla `.msix` Paketini ve Vitrin Görsellerini Üretin

> **Harika Haber (Windows SDK Kurmanıza Gerek Kalmadı!):** MEB / kurum internetinde dev Windows SDK yükleyicisi (`15605` indirme hatası) takılsa bile, `build_msix.py` artık resmi Microsoft NuGet sunucusundan yalnızca ~20 MB boyutundaki Microsoft imzalı `makeappx.exe` aracını yönetici izni gerekmeden otomatik indirip `build\sdk_tools` altına kurar!

1. Doğrudan proje klasöründeki **`msix_olustur.bat`** dosyasına çift tıklayın (veya terminalde `python build_msix.py` çalıştırın).
2. İşlem bittiğinde iki çıktı hazır olur:
   - **Mağazaya yüklenecek paket:** `installer\RevirDefteri_Store_v<SÜRÜM>.msix`
   - **Mağaza vitrin görselleri (300x300 logo, 720x1080 poster, 1920x1080 kapak):** `store_submission_assets\`

---

## 3. Partner Center Başvuru Formuna Yapıştıracağınız Hazır Bilgiler

### 🔗 Privacy Policy URL (Gizlilik Politikası Bağlantısı — Zorunlu)
```text
https://github.com/blue-doggy/revir_defteri/blob/main/PRIVACY_POLICY.md
```

### 📝 Store Description (Mağaza Açıklaması — Kopyala/Yapıştır)
```text
Okul Revir Defteri ve Sağlık Takip Sistemi; okul hemşireleri ve sağlık personeli için geliştirilmiş, hızlı, modern ve tamamen yerel veritabanlı (çevrimdışı çalışabilen) kapsamlı bir okul sağlığı yönetim yazılımıdır.

Öne Çıkan Özellikler:
• Günlük Revir Muayene Defteri, Sevk, İstirahat ve Açık İzlem Takibi
• WHO (Dünya Sağlık Örgütü) Büyüme Eğrileri, Z-Skoru ve Persantil Hesaplama
• Okul Sağlık Taraması (Öğrenci ve Yetişkin Taramaları, Excel/Fotoğraf/Belge Akıllı Aktarım)
• İlaç Takip Sistemi, Soğuk Zincir ve Revir Malzeme/Stok Yönetimi
• Pediatrik Diyabet & İnsülin Karar Destek Kokpiti (Hekim reçeteli İKO/İDF, Aktif İnsülin IOB ve Glisemik Trend TIR analizi)
• Yemekhane Karbonhidrat/Alerjen Planlayıcı ve Okul Yurdu Müşahede/İzolasyon Modülü
• Resmi MEB Antetli Word (.docx), Excel (.xlsx) ve Yazdırılabilir Rapor Stüdyosu
• KVKK Uyumlu Yerel Veri Mimarisi ve İsteğe Bağlı Google Gemini AI Klinik Asistan Desteği

⚕️ Tıbbi Sorumluluk Beyanı: Bu yazılım tıbbi tanı veya tedavi cihazı değildir; hekim muayenesinin veya reçetesinin yerine geçmez. Yalnızca okul sağlığı personeline yönelik kayıt tutma ve karar destek aracıdır.
```

### 🧪 Notes for Certification (Denetçi Notu — Partner Center'da Çok Önemli!)
```text
This desktop application runs completely offline using a local SQLite database (%LOCALAPPDATA%\RevirDefteri) and requires no login credentials or external account to launch and test.

1. Launch the app directly to access the main dashboard (Genel Bakış), add sample students (Öğrenciler), record clinic visits (Revir Kaydı), calculate WHO growth percentiles (Ölçüm & Gelişim), and export Word/Excel reports.
2. AI features (Google Gemini) are strictly optional; the user can optionally supply their own API key under Settings (Ayarlar > Yapay Zekâ & KVKK). All core modules work 100% without an API key.
3. Medical Disclaimer: This software is an administrative record-keeping and educational/clinical decision-support tool for school infirmaries; it is not a medical device and does not replace a physician's prescription.
```

### 🔞 Yaş Derecelendirmesi (Age Ratings / IARC) Anketi İpucu
- Uygulama şiddet, kumar, kullanıcılar arası açık sohbet, reklam veya uygulama içi satın alma **içermez** (`3+ / Everyone` çıkacaktır).
