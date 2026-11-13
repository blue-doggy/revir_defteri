# Proje Sürüm Yönetimi Kuralı

Bu çalışma alanındaki (`persantil.py`, `desktop_app.py`, paketleme veya kurulum dosyaları dahil) **her güncellemede sürüm numarası zorunlu olarak yükseltilmelidir**:

1. Tek sürüm kaynağı proje kökündeki `VERSION` (`X.Y.Z`, SemVer) dosyasıdır.
2. Her kod veya paket güncellemesinin sonunda `python bump_version.py` çalıştırılır:
   - Yeni özellik, modül veya kapsamlı performans/mimari güncelleme için: `python bump_version.py minor`
   - Hata düzeltmesi, küçük iyileştirme veya yama için: `python bump_version.py patch`
   - Kırıcı büyük sürüm geçişi için: `python bump_version.py major`
3. `bump_version.py`, `VERSION` dosyasını ve `persantil.py`, `desktop_app.py`, `RevirDefteri_Setup.iss` ile `kurulum_olustur.iss` içindeki sürüm sabitlerini senkronize tutar.
4. Kurulum paketi yeniden derlenirken (`paketi_olustur.bat` veya Inno Setup `ISCC.exe`) üretilen dosya adı `RevirDefteri_Kurulum_v<VERSION>.exe` biçiminde olmalı ve kökteki eski sürüm `RevirDefteri_Kurulum_v*.exe` kopyaları temizlenmelidir.
5. GitHub'a commit/push yapılırken commit mesajında yeni sürüm (`vX.Y.Z`) belirtilmelidir.
