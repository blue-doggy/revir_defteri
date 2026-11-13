; ============================================================================
; Okul Revir Defteri ve Sağlık Takip Sistemi - Akıllı Güncelleme Destekli Kurulum Scripti
; Sabit Kurulum Dizini: C:\RevirDefteri
; Çıktı: installer\RevirDefteri_Kurulum_v1.2.0.exe
; ============================================================================

#define MyAppName "Okul Revir Defteri ve Sağlık Takip Sistemi"
#define MyShortName "Okul Revir Defteri"
#define MyAppPublisher "Okul Sağlığı & Revir Bilgi Sistemi"
#define MyAppExeName "RevirDefteri.exe"
#define SourceDistDir "dist\RevirDefteri"

; Sürüm numarasını doğrudan proje kökündeki VERSION dosyasından oku
#define MyAppVersionFile FileOpen("VERSION")
#if MyAppVersionFile
  #define MyAppVersion Trim(FileRead(MyAppVersionFile))
  #expr FileClose(MyAppVersionFile)
#else
  #define MyAppVersion "1.2.0"
#endif

[Setup]
AppId={{8F6C2B9A-4D17-4A8E-9E3A-1D7B5C2F9041}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyShortName} v{#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName=C:\RevirDefteri
UsePreviousAppDir=yes
DirExistsWarning=no
DefaultGroupName={#MyShortName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=RevirDefteri_Kurulum_v{#MyAppVersion}
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/normal
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
AppMutex=Local\OkulRevirDefteri_SingleInstance_Mutex_v1
SetupMutex=Local\OkulRevirDefteri_Setup_Mutex_v1
CloseApplications=force
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
; C:\RevirDefteri ve tüm veri klasörlerine tüm kullanıcılar için Tam Okuma/Yazma (users-modify) izni ver
; ve kaldırma/güncelleme sırasında asla silinmemesini (uninsneveruninstall) garanti et.
Name: "{app}"; Permissions: users-modify
Name: "{app}\kisi_verileri"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\egitim_yillari"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\revir_evraklari"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\oto_yedekler"; Permissions: users-modify; Flags: uninsneveruninstall
Name: "{app}\.webview2_cache"; Permissions: users-modify; Flags: uninsneveruninstall

[InstallDelete]
; Güncelleme sırasında yalnızca eski PyInstaller _internal kütüphane klasörünü temizle;
; revir_defteri.db ve kullanıcı veri klasörlerine kesinlikle dokunma.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
; Program dosyaları (mevcut veritabanı ve veri klasörleri güncellemede asla ezilmez)
Source: "{#SourceDistDir}\*"; DestDir: "{app}"; Excludes: "revir_defteri.db,revir_defteri.db-shm,revir_defteri.db-wal,revir_defteri.db-journal,egitim_yillari\*,kisi_verileri\*,revir_evraklari\*,oto_yedekler\*,.webview2_cache\*,.revir_server.log"; Flags: ignoreversion recursesubdirs createallsubdirs

; Simgenin ve sürüm dosyasının C:\RevirDefteri içinde doğrudan bulunması için kopyala
Source: "app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "VERSION"; DestDir: "{app}"; Flags: ignoreversion

; Eğer dist içinde başlangıç revir_defteri.db varsa yalnızca hedefte yoksa kopyala; güncellemede ve kaldırmada koru
Source: "{#SourceDistDir}\revir_defteri.db"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall skipifsourcedoesntexist

[Icons]
Name: "{autoprograms}\{#MyShortName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"; WorkingDir: "{app}"
Name: "{autoprograms}\{cm:UninstallProgram,{#MyShortName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyShortName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Okul Revir Defteri Programını Şimdi Başlat"; Flags: nowait postinstall skipifsilent

[Code]
const
  UNINSTALL_SUBKEY = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{8F6C2B9A-4D17-4A8E-9E3A-1D7B5C2F9041}_is1';
  APP_MUTEX_NAME = 'Local\OkulRevirDefteri_SingleInstance_Mutex_v1';

var
  ModePage: TInputOptionWizardPage;
  AutoBackupCheckBox: TNewCheckBox;
  HasExistingInstall: Boolean;
  DetectedVersion: String;
  DetectedInstallDir: String;

function ReadFirstTrimmedLine(const FileName: String): String;
var
  Lines: TArrayOfString;
begin
  Result := '';
  if FileExists(FileName) and LoadStringsFromFile(FileName, Lines) then
  begin
    if GetArrayLength(Lines) > 0 then
      Result := Trim(Lines[0]);
  end;
end;

procedure DetectExistingInstallation();
var
  RegDir, RegVer: String;
begin
  HasExistingInstall := False;
  DetectedVersion := '';
  DetectedInstallDir := 'C:\RevirDefteri';

  if RegQueryStringValue(HKEY_LOCAL_MACHINE, UNINSTALL_SUBKEY, 'Inno Setup: App Path', RegDir) or
     RegQueryStringValue(HKEY_LOCAL_MACHINE, UNINSTALL_SUBKEY, 'InstallLocation', RegDir) then
  begin
    if Trim(RegDir) <> '' then
      DetectedInstallDir := RemoveBackslashUnlessRoot(Trim(RegDir));
    HasExistingInstall := True;
  end
  else if RegQueryStringValue(HKEY_CURRENT_USER, UNINSTALL_SUBKEY, 'Inno Setup: App Path', RegDir) or
          RegQueryStringValue(HKEY_CURRENT_USER, UNINSTALL_SUBKEY, 'InstallLocation', RegDir) then
  begin
    if Trim(RegDir) <> '' then
      DetectedInstallDir := RemoveBackslashUnlessRoot(Trim(RegDir));
    HasExistingInstall := True;
  end;

  if RegQueryStringValue(HKEY_LOCAL_MACHINE, UNINSTALL_SUBKEY, 'DisplayVersion', RegVer) or
     RegQueryStringValue(HKEY_CURRENT_USER, UNINSTALL_SUBKEY, 'DisplayVersion', RegVer) then
  begin
    DetectedVersion := Trim(RegVer);
    HasExistingInstall := True;
  end;

  if FileExists(AddBackslash(DetectedInstallDir) + 'revir_defteri.db') or
     FileExists(AddBackslash(DetectedInstallDir) + 'RevirDefteri.exe') then
  begin
    HasExistingInstall := True;
  end
  else if FileExists('C:\RevirDefteri\revir_defteri.db') or
          FileExists('C:\RevirDefteri\RevirDefteri.exe') then
  begin
    DetectedInstallDir := 'C:\RevirDefteri';
    HasExistingInstall := True;
  end;

  if DetectedVersion = '' then
    DetectedVersion := ReadFirstTrimmedLine(AddBackslash(DetectedInstallDir) + 'VERSION');
  if DetectedVersion = '' then
    DetectedVersion := ReadFirstTrimmedLine(AddBackslash(DetectedInstallDir) + '_internal\VERSION');
end;

procedure ForceStopRunningApp();
var
  ResultCode: Integer;
begin
  if CheckForMutexes(APP_MUTEX_NAME) then
  begin
    Exec(ExpandConstant('{cmd}'), '/C taskkill /F /T /IM RevirDefteri.exe >nul 2>&1', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Sleep(400);
  end;
end;

function MoveDirectoryRecursive(const SrcDir, DstDir: String): Boolean;
var
  FindRec: TFindRec;
  SrcPath, DstPath: String;
begin
  Result := True;
  if not DirExists(SrcDir) then
    Exit;

  if RenameFile(SrcDir, DstDir) then
    Exit;

  if not ForceDirectories(DstDir) then
  begin
    Result := False;
    Exit;
  end;

  if FindFirst(AddBackslash(SrcDir) + '*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          SrcPath := AddBackslash(SrcDir) + FindRec.Name;
          DstPath := AddBackslash(DstDir) + FindRec.Name;
          if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
          begin
            if not MoveDirectoryRecursive(SrcPath, DstPath) then
              Result := False;
          end
          else
          begin
            if not RenameFile(SrcPath, DstPath) then
            begin
              if FileCopy(SrcPath, DstPath, False) then
                DeleteFile(SrcPath)
              else
                Result := False;
            end;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;

  if Result then
    RemoveDir(SrcDir);
end;

procedure ArchiveFileIfExists(const SrcFile, DstFile: String);
begin
  if FileExists(SrcFile) then
  begin
    if not RenameFile(SrcFile, DstFile) then
    begin
      if FileCopy(SrcFile, DstFile, False) then
        DeleteFile(SrcFile);
    end;
  end;
end;

procedure InitializeWizard();
var
  SubTitleText: String;
  DescText: String;
  Option2Text: String;
begin
  DetectExistingInstallation();

  if HasExistingInstall then
  begin
    if DetectedVersion <> '' then
      SubTitleText := 'Bilgisayarınızda kurulu bir Okul Revir Defteri sürümü (v' + DetectedVersion + ') tespit edildi. Lütfen yapmak istediğiniz işlemi seçin:'
    else
      SubTitleText := 'Bilgisayarınızda kurulu bir Okul Revir Defteri sürümü (v1.0) tespit edildi. Lütfen yapmak istediğiniz işlemi seçin:';
    DescText := 'Yeni Paket Sürümü: v{#MyAppVersion}  |  Mevcut Kurulum Konumu: ' + DetectedInstallDir;
    Option2Text := '🆕 Sıfırdan Temiz Kurulum Yap (Eski Verileri Yedek Klasörüne Arşivle) — Mevcut veritabanı ''oto_yedekler'' klasörüne taşınır ve program boş veritabanıyla sıfırdan kurulur.';
  end
  else
  begin
    SubTitleText := 'Bilgisayarınızda önceki bir Okul Revir Defteri kurulumu bulunamadı. İlk kurulum için "Yeni Kurulum Modu" otomatik olarak seçildi:';
    DescText := 'Yeni Kurulum Sürümü: v{#MyAppVersion}  |  Kurulum Dizini: C:\RevirDefteri';
    Option2Text := '🆕 Yeni Kurulum Modu — Sıfırdan Temiz Kurulum Yap (İlk kurulum için boş veritabanı ve veri klasörleri oluşturulur; varsa eski veriler ''oto_yedekler'' klasörüne arşivlenir).';
  end;

  ModePage := CreateInputOptionPage(
    wpWelcome,
    'Kurulum ve Güncelleme Seçenekleri',
    SubTitleText,
    DescText,
    True,
    False
  );

  ModePage.Add('🔄 Mevcut Sürümü Güncelle ve Tüm Kayıtları Koru (Önerilen) — Öğrenci, revir, tarama, ilaç ve eğitim yılı verileriniz korunur; sadece program yeni sürüme yükseltilir.');
  ModePage.Add(Option2Text);

  ModePage.CheckListBox.Height := ScaleY(130);

  if HasExistingInstall then
    ModePage.SelectedValueIndex := 0
  else
    ModePage.SelectedValueIndex := 1;

  AutoBackupCheckBox := TNewCheckBox.Create(ModePage);
  AutoBackupCheckBox.Parent := ModePage.Surface;
  AutoBackupCheckBox.Left := ScaleX(0);
  AutoBackupCheckBox.Top := ModePage.CheckListBox.Top + ModePage.CheckListBox.Height + ScaleY(12);
  AutoBackupCheckBox.Width := ModePage.SurfaceWidth;
  AutoBackupCheckBox.Height := ScaleY(36);
  AutoBackupCheckBox.Caption := '☑️ Güncelleme öncesinde mevcut veritabanının otomatik yedeğini (C:\RevirDefteri\oto_yedekler) oluştur';
  AutoBackupCheckBox.Checked := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = ModePage.ID then
  begin
    if (ModePage.SelectedValueIndex = 1) and HasExistingInstall and (not WizardSilent()) then
    begin
      if MsgBox(
           'Sıfırdan Temiz Kurulum seçildi.' + #13#10 + #13#10 +
           'Mevcut "revir_defteri.db" ve eğitim yılı verileriniz silinmeyecek, güvenlik amacıyla "oto_yedekler\Arsiv_..." klasörüne taşınacak ve uygulama boş veritabanıyla açılacaktır.' + #13#10 + #13#10 +
           'Devam etmek istiyor musunuz?',
           mbConfirmation,
           MB_YESNO
         ) <> IDYES then
      begin
        Result := False;
      end;
    end;
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  ModeLine: String;
  BackupLine: String;
  SummaryBlock: String;
begin
  if ModePage.SelectedValueIndex = 0 then
    ModeLine := 'İşlem Türü: Mevcut Sürümü Güncelle (Veriler Korunacak)'
  else if HasExistingInstall then
    ModeLine := 'İşlem Türü: Sıfırdan Temiz Kurulum (Eski Veriler oto_yedekler İçine Arşivlenecek)'
  else
    ModeLine := 'İşlem Türü: Yeni Kurulum Modu (İlk Kurulum)';

  if AutoBackupCheckBox.Checked then
    BackupLine := 'Otomatik Ön Yedekleme: Açık (oto_yedekler\kurulum_oncesi_yedek_*.db)'
  else
    BackupLine := 'Otomatik Ön Yedekleme: Kapalı';

  SummaryBlock := ModeLine + NewLine + Space + BackupLine + NewLine + NewLine;

  Result := SummaryBlock;
  if MemoDirInfo <> '' then
    Result := Result + MemoDirInfo + NewLine + NewLine;
  if MemoGroupInfo <> '' then
    Result := Result + MemoGroupInfo + NewLine + NewLine;
  if MemoTasksInfo <> '' then
    Result := Result + MemoTasksInfo + NewLine;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  ForceStopRunningApp();
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppDirPath: String;
  DbFilePath: String;
  BackupDirPath: String;
  BackupFilePath: String;
  ArchiveDirPath: String;
  TimeStamp: String;
begin
  if CurStep = ssInstall then
  begin
    ForceStopRunningApp();

    AppDirPath := ExpandConstant('{app}');
    DbFilePath := AddBackslash(AppDirPath) + 'revir_defteri.db';
    BackupDirPath := AddBackslash(AppDirPath) + 'oto_yedekler';
    TimeStamp := GetDateTimeString('yyyymmdd_hhnnss', '', '');

    // 1. Güncelleme / Kurulum başlamadan önce mevcut veritabanının yedeğini al
    if AutoBackupCheckBox.Checked and FileExists(DbFilePath) then
    begin
      ForceDirectories(BackupDirPath);
      BackupFilePath := AddBackslash(BackupDirPath) + 'kurulum_oncesi_yedek_' + TimeStamp + '.db';
      FileCopy(DbFilePath, BackupFilePath, False);
      if FileExists(DbFilePath + '-wal') then
        FileCopy(DbFilePath + '-wal', BackupFilePath + '-wal', False);
      if FileExists(DbFilePath + '-shm') then
        FileCopy(DbFilePath + '-shm', BackupFilePath + '-shm', False);
    end;

    // 2. Seçenek 2 seçildiyse: eski verileri silmek yerine oto_yedekler\Arsiv_<Tarih> altına taşı
    if ModePage.SelectedValueIndex = 1 then
    begin
      if FileExists(DbFilePath) or
         DirExists(AddBackslash(AppDirPath) + 'egitim_yillari') or
         DirExists(AddBackslash(AppDirPath) + 'kisi_verileri') or
         DirExists(AddBackslash(AppDirPath) + 'revir_evraklari') then
      begin
        ArchiveDirPath := AddBackslash(BackupDirPath) + 'Arsiv_' + TimeStamp;
        ForceDirectories(ArchiveDirPath);

        ArchiveFileIfExists(DbFilePath, AddBackslash(ArchiveDirPath) + 'revir_defteri.db');
        ArchiveFileIfExists(DbFilePath + '-wal', AddBackslash(ArchiveDirPath) + 'revir_defteri.db-wal');
        ArchiveFileIfExists(DbFilePath + '-shm', AddBackslash(ArchiveDirPath) + 'revir_defteri.db-shm');
        ArchiveFileIfExists(DbFilePath + '-journal', AddBackslash(ArchiveDirPath) + 'revir_defteri.db-journal');

        MoveDirectoryRecursive(AddBackslash(AppDirPath) + 'egitim_yillari', AddBackslash(ArchiveDirPath) + 'egitim_yillari');
        MoveDirectoryRecursive(AddBackslash(AppDirPath) + 'kisi_verileri', AddBackslash(ArchiveDirPath) + 'kisi_verileri');
        MoveDirectoryRecursive(AddBackslash(AppDirPath) + 'revir_evraklari', AddBackslash(ArchiveDirPath) + 'revir_evraklari');
      end;
    end;
  end;
end;
