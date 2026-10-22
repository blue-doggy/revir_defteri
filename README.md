# 🩺 Revir Defteri - Okul Sağlık & Revir Yönetim Sistemi

**Revir Defteri**, anaokulu, ilkokul, ortaokul ve liseler ile her kademedeki eğitim kurumlarının revirleri için geliştirilmiş; öğrenci ve personel sağlık takibini, büyüme-gelişme persantillerini, aşı takvimini, günlük vizitleri ve ilaç uygulamalarını tek merkezden yöneten **bağımsız masaüstü okul sağlığı yazılımıdır**.

Tüm veriler kurumun kendi bilgisayarında yerel olarak saklanır; **hiçbir öğrenci, personel veya okul verisi harici bulut servislerine veya üçüncü taraf yapay zekâ sunucularına aktarılmaz**.

---

## 🌟 Temel Özellikler

- **WHO (Dünya Sağlık Örgütü) Büyüme Standartları:** Boy, kilo, beden kitle indeksi (BKİ) ve baş çevresi için yaşa ve cinsiyete göre tam LMS z-skoru ve persantil hesaplama.
- **T.C. Sağlık Bakanlığı Aşı Takvimi Entegrasyonu:** Hepatit B, BCG, DaBT-İPA-Hib, KKK, Suçiçeği vb. tüm aşıların takibi, tek tıkla işaretleme ve resmi çıktı.
- **Detaylı Vizit & Müdahale Kayıtları:** Tarih/saat bazlı geliş nedeni, vital bulgular (Ateş, Tansiyon, Nabız, SpO2 Satürasyon), yapılan tıbbi müdahaleler ve veli bilgilendirme durumu.
- **İlaç & Alerji Yönetimi:** Sürekli kullanılan reçeteli ilaçlar, saklama koşulları, veli izinleri ve günlük dozaj uygulama defteri.
- **Dijital Evrak & Tahlil Arşivi:** Öğrencilere ait sağlık raporları, heyet belgeleri ve tahlil sonuçlarını PDF/görsel olarak yükleme, yerel OCR desteği ve ilk sayfa kapak önizlemesi.
- **Yetişkinler (Personel & Ziyaretçi) Takibi:** Okul personeli ve dış ziyaretçiler için ayrı vizit ve tansiyon/vital kayıt defteri.
- **Eğitim-Öğretim Yılı İzolasyonu:** Yıllar arası tam bağımsız arşivleme ve anında geçmiş eğitim yıllarına geçiş.
- **Bağımsız Masaüstü Uygulaması:** Tarayıcı sekmeleri veya adres çubuğu olmadan, konsol/siyah ekran açılmadan doğrudan yerel masaüstü penceresi olarak çalışma (`baslat.pyw`).
- **Tek Tıkla Otomatik Güncelleme:** Veriler otomatik yedeklenerek resmi GitHub reposu üzerinden tek tıkla sorunsuz sürüm yükseltme.

---

## 🚀 Kurulum ve Çalıştırma

### 1. Gereksinimler
- Bilgisayarınızda **Python 3.10 veya üzeri** (tercihen 64-bit) kurulu olmalıdır.
- Git (Otomatik güncellemeler için önerilir).

### 2. Bağımlılıkların Yüklenmesi
Uygulama klasöründe terminal / PowerShell açarak gerekli kütüphaneleri yükleyin:
```powershell
pip install -r requirements.txt
```

### 3. Uygulamayı Başlatma

#### 🖥️ Seçenek A: Bağımsız Masaüstü Uygulaması Olarak (Önerilen)
Siyah konsol penceresi olmadan, yerel bir masaüstü yazılımı gibi çalıştırmak için:
```powershell
pythonw baslat.pyw
```
*(Uygulama ilk açıldığında masaüstünüze otomatik olarak özel ikonlu **Revir Defteri** kısayolu oluşturur. Sonraki seferlerde doğrudan masaüstündeki simgeye çift tıklayabilirsiniz.)*

#### 🌐 Seçenek B: Standart Streamlit Modunda
```powershell
streamlit run persantil.py
```

---

## 📋 İlk Kurulumda Doldurulacak / Yapılandırılacak Kısımlar (Adım Adım Kılavuz)

Projeyi yeni bir bilgisayara veya yeni bir okula kurduğunuzda ilk olarak doldurmanız gereken alanlar şunlardır:

### 1. Adım: Kurum ve Revir Bilgilerini Ayarlama
*Uygulama içinde sol menüden veya sağ üstten **⚙️ Ayarlar** sekmesine gidin:*
- **Kurum Adı:** Okulunuzun tam resmi adı (Örn: *Atatürk İlkokulu*, *Mehmet Akif Ersoy Fen Lisesi*). Rapor ve çıktılarda bu isim görünür.
- **Üst Başlık:** Resmi evrak üst anteti (Örn: *T.C. MİLLÎ EĞİTİM BAKANLIĞI*).
- **Okul Müdürü:** Rapor çıktılarındaki onay alanı için müdür adı-soyadı.
- **Revir Sorumlusu:** Revir hemşiresi, sağlık personeli veya okul doktorunun unvanı ve adı.
- **İl / İlçe / Mahalle:** Adres seçimlerinde varsayılan olarak gelecek şehir ve ilçe.
- **Görsel Tema ve Yazı Tipi:** Tercihinize göre klinik, modern veya karanlık tema seçimi.

### 2. Adım: Aktif Eğitim-Öğretim Yılını Belirleme
- **Ayarlar > Eğitim-Öğretim Yılları** alanından içinde bulunduğunuz yılı (Örn: `2026-2027`) aktif eğitim yılı olarak seçin veya gerekiyorsa yeni bir eğitim yılı oluşturun.
- Her eğitim yılı kendi bağımsız veritabanında saklanır; yıl bittiğinde yeni yıla geçildiğinde eski kayıtlar arşivde güvenle korunur.

### 3. Adım: Sınıf ve Öğrenci Listelerini Tanımlama
*Sol menüden **👥 Öğrenciler** ekranına geçin:*
1. **Sınıf Tanımlama:** Okulunuzda bulunan şubeleri (Örn: *Anasınıfı-A*, *1-A*, *5-B*, *12-Fen-C*) sisteme ekleyin.
2. **Öğrenci Yükleme:**
   - **Excel ile Toplu Yükleme:** Ekrandaki Excel şablonunu indirin; öğrencinin **Adı Soyadı**, **Sınıfı**, **Doğum Tarihi** ve **Cinsiyetini** (`E` veya `K`) doldurarak sisteme yükleyin.
   - **Tek Tek Ekleme:** Sağ üstteki *"Yeni Öğrenci Kaydı"* butonu ile öğrenci profili açabilirsiniz.

### 4. Adım: Öğrenci Sağlık Profili ve Özel Bilgilerin Doldurulması
Her öğrencinin detay kartında doldurulabilecek kritik sağlık alanları:
- **Kan Grubu ve Kimlik/Numara Bilgisi**
- **Kronik Rahatsızlıklar:** Varsa astım, epilepsi, diyabet, çölyak vb. hastalıklar.
- **Alerjiler:** İlaç, besin, arı sokması veya mevsimsel alerjiler.
- **Düzenli İlaçlar:** Çocuğun okul saatinde alması gereken ilaçlar, dozajı ve veli muvafakat durumu.
- **Aşı Kartı:** Sağlık Bakanlığı aşı takvimine göre tamamlanan dozların işaretlenmesi.
- **Veli İletişim Bilgileri:** Acil durumda aranacak anne, baba veya vasi telefon numaraları.

### 5. Adım: Yetişkinler (Personel & Ziyaretçi) Kayıtları
*Sol menüden **🧑‍💼 Yetişkinler** sekmesine geçin:*
- Öğretmenler, memurlar, destek personeli ve veli/ziyaretçiler için profil kartı açın.
- Bu profillerde persantil boy/kilo eğrileri yerine yetişkin vital takibi (tansiyon, nabız, şeker) uygulanır.

### 6. Adım: Günlük Revir Vizit Kayıtlarının Tutulması
*Ana ekranda (**🩺 Revir Defteri**):*
- Öğrenci veya personeli seçin.
- Geliş şikayetini ve saatini seçin.
- Vital bulguları girin (Ateş, Tansiyon, Nabız, SpO2).
- Yapılan müdahaleyi seçin (Pansuman, soğuk kompres, dinlendirme, veliye teslim vb.).
- Kaydet butonuna bastığınızda resmi revir defteri satırı otomatik oluşturulur.

---

## 🔒 Veri Güvenliği ve Gizlilik Taahhüdü

1. **Tamamen Çevrimdışı (Offline) Çalışma:** Uygulama internet bağlantısı olmadan %100 işlevseldir.
2. **Kişisel Verilerin Korunması (KVKK):** Öğrenci kimlikleri, sağlık öyküleri, ölçümleri ve evrakları yalnızca kurum bilgisayarındaki yerel `revir_defteri.db` veritabanında ve `kisi_verileri/` klasöründe saklanır.
3. **Açık Kaynak Kod Temizliği:** Bu depoda **kesinlikle hiçbir öğrenci, sınıf veya kuruma ait gerçek veri bulunmaz**. Projeyi klonlayan veya güncelleyen kullanıcılar sadece temiz uygulama motorunu ve kamuya açık DSÖ persantil tablolarını edinirler.

---

## 📦 Yedekleme ve Geri Yükleme

- **Tek Tıkla Tam Sistem Yedeği (ZIP):** **Ayarlar > Veri Güvenliği & Yedekleme** alanından dilediğiniz zaman tüm veritabanını, kişisel sağlık formlarını ve yüklenen tüm evrakları içeren sıkıştırılmış `.zip` yedeği alabilirsiniz.
- **Yedekten Geri Yükleme:** Bilgisayar değişikliğinde aldığınız `.zip` yedeğini aynı ekrandan yükleyerek tüm sistemi saniyeler içinde geri yükleyebilirsiniz.

---

## 🔄 Otomatik Güncellemeler

Yazılımda yeni bir özellik geliştirildiğinde veya hata düzeltmesi yapıldığında:
1. **Ayarlar > 🚀 Sistem Sürümü & Güncelleme Yönetimi** sekmesine gidin.
2. **🔍 Güncellemeleri Denetle** butonuna basın.
3. Sistem arka planda bu resmi GitHub deposunu (`blue-doggy/revir_defteri`) denetler; yeni bir sürüm varsa yapılan yenilikleri listeler.
4. **⚡ Şimdi Güncelle ve Yeniden Başlat** butonuna tıkladığınızda, **önce tüm öğrenci ve kurum verileriniz otomatik olarak yedeklenir**, ardından en güncel kodlar indirilerek uygulama sorunsuz şekilde yenilenir.

---

## 📄 Lisans
Bu proje eğitim kurumlarında öğrenci sağlığını desteklemek amacıyla geliştirilmiştir.
