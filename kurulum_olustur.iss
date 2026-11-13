; ============================================================================
; Okul Revir Defteri ve Sağlık Takip Sistemi - Inno Setup 6 Kurulum Scripti
; Üretilen Çıktı: installer\Revir_Defteri_Kurulum.exe
;
; Ön Koşul:
;   Önce PyInstaller ile "dist\RevirDefteri" klasörünü oluşturun:
;     pyinstaller build_exe.spec --noconfirm --clean
;
; Derleme (Inno Setup Compiler 6 ile):
;   ISCC.exe kurulum_olustur.iss
; ============================================================================

#define MyAppName "Okul Revir Defteri"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "Okul Sağlığı & Revir Bilgi Sistemi"
#define MyAppExeName "RevirDefteri.exe"
#define SourceDistDir "dist\RevirDefteri"

[Setup]
AppId={{8F6C2B9A-4D17-4A8E-9E3A-1D7B5C2F9041}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\RevirDefteri
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=Revir_Defteri_Kurulum
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/normal
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
CloseApplications=yes
CloseApplicationsFilter=*.exe
RestartApplications=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Dirs]
; SQLite veritabanı ve eğitim yılları klasörlerinin güncelleme ve kaldırmada silinmesini engelle
; ve standart kullanıcı hesaplarının yazma iznine sahip olmasını sağla.
Name: "{app}"; Permissions: users-modify
Name: "{app}\egitim_yillari"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\kisi_verileri"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\revir_evraklari"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\oto_yedekler"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\.webview2_cache"; Permissions: users-modify; Flags: uninsneveruninstall

[InstallDelete]
; Yalnızca yeni PyInstaller _internal klasörünün temiz kurulması için eski kütüphane kalıntılarını temizle.
; Kullanıcı verileri (revir_defteri.db, egitim_yillari, kisi_verileri, revir_evraklari, oto_yedekler) KESİNLİKLE silinmez.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
; Ana uygulama dosyaları (veritabanı ve kullanıcı veri klasörleri güncelleme sırasında hariç tutulur)
Source: "{#SourceDistDir}\*"; DestDir: "{app}"; Excludes: "revir_defteri.db,revir_defteri.db-shm,revir_defteri.db-wal,egitim_yillari\*,kisi_verileri\*,revir_evraklari\*,oto_yedekler\*,.webview2_cache\*,.revir_server.log"; Flags: ignoreversion recursesubdirs createallsubdirs

; İlk kurulum için başlangıç revir_defteri.db dosyası dist içinde varsa yalnızca yoksa kopyala; güncellemede asla üzerine yazma
Source: "{#SourceDistDir}\revir_defteri.db"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
