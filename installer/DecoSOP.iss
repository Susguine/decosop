; DecoSOP Inno Setup Installer Script
; Requires Inno Setup 6.x (https://jrsoftware.org/isinfo.php)
;
; Wizard pages:
;   1. Welcome
;   2. License Agreement
;   3. Install Directory
;   4. Port Configuration
;   5. Database Setup (Empty / Demo / Import backup / Scan folders)
;   6. Import Database file      (only if Import selected)
;   7. Import SOP files dir      (only if Import selected, optional)
;   8. Import Documents dir      (only if Import selected, optional)
;   9. Scan SOP source dir       (only if Scan selected)
;  10. Scan Documents source dir (only if Scan selected)
;  11. Auto-Update Preference
;  12. Shortcuts (desktop icon)
;  13. Ready to Install
;  14. Installing...
;  15. Finish (open in browser)

#define MyAppName "DecoSOP"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Tyler Sweeney"
#define MyAppURL "https://github.com/Susguine/decosop"
#define MyAppExeName "DecoSOP.exe"

[Setup]
AppId={{D3C0-50F1-4A2B-B8E9-DecoSOP-1000}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
DefaultDirName=C:\DecoSOP
DefaultGroupName={#MyAppName}
LicenseFile=license.txt
OutputDir=..\installer-output
OutputBaseFilename=DecoSOP-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=force

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Published app files — excludes DB and upload dirs (user data)
Source: "..\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "decosop.db,uploads,sop-uploads"

[Dirs]
Name: "{app}\uploads"
Name: "{app}\sop-uploads"

[Icons]
Name: "{group}\DecoSOP"; Filename: "http://localhost:{code:GetPort}"
Name: "{group}\Uninstall DecoSOP"; Filename: "{uninstallexe}"
Name: "{commondesktop}\DecoSOP"; Filename: "http://localhost:{code:GetPort}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
; Stop existing service if upgrading
Filename: "sc.exe"; Parameters: "stop DecoSOP"; Flags: runhidden; StatusMsg: "Stopping existing service..."; Check: ServiceExists
; Wait for service to stop
Filename: "cmd.exe"; Parameters: "/c timeout /t 3 /nobreak >nul"; Flags: runhidden; Check: ServiceExists
; Delete old service before re-creating (handles path changes)
Filename: "sc.exe"; Parameters: "delete DecoSOP"; Flags: runhidden; Check: ServiceExists
; Wait for deletion
Filename: "cmd.exe"; Parameters: "/c timeout /t 2 /nobreak >nul"; Flags: runhidden; Check: ServiceExists

; Register Windows Service
Filename: "sc.exe"; Parameters: "create DecoSOP binPath=""{app}\{#MyAppExeName}"" start=auto displayname=""DecoSOP Document Manager"""; Flags: runhidden waituntilterminated; StatusMsg: "Registering Windows Service..."
Filename: "sc.exe"; Parameters: "description DecoSOP ""DecoSOP - SOP & Document Management System"""; Flags: runhidden waituntilterminated

; Write configuration files
Filename: "cmd.exe"; Parameters: "/c echo {code:GetPortConfigContent}> ""{app}\port.config"""; Flags: runhidden; Check: IsCustomPort; StatusMsg: "Writing port configuration..."
Filename: "cmd.exe"; Parameters: "/c echo {code:GetUpdateConfigContent}> ""{app}\update-config.json"""; Flags: runhidden; StatusMsg: "Writing update configuration..."

; Copy imported database if user chose "Import backup"
Filename: "cmd.exe"; Parameters: "/c copy /y ""{code:GetImportDbPath}"" ""{app}\decosop.db"""; Flags: runhidden; Check: ShouldImportDb; StatusMsg: "Importing database backup..."

; Copy SOP upload files if directory was selected
Filename: "robocopy.exe"; Parameters: """{code:GetImportSopDirPath}"" ""{app}\sop-uploads"" /E /NFL /NDL /NJH /NJS"; Flags: runhidden; Check: ShouldImportSopDir; StatusMsg: "Importing SOP files..."

; Copy Document upload files if directory was selected
Filename: "robocopy.exe"; Parameters: """{code:GetImportDocDirPath}"" ""{app}\uploads"" /E /NFL /NDL /NJH /NJS"; Flags: runhidden; Check: ShouldImportDocDir; StatusMsg: "Importing document files..."

; Add firewall rule
Filename: "netsh.exe"; Parameters: "advfirewall firewall delete rule name=""DecoSOP"""; Flags: runhidden; StatusMsg: "Updating firewall rules..."
Filename: "netsh.exe"; Parameters: "advfirewall firewall add rule name=""DecoSOP"" dir=in action=allow protocol=TCP localport={code:GetPort}"; Flags: runhidden waituntilterminated

; Seed demo data if selected
Filename: "{app}\{#MyAppExeName}"; Parameters: "--seed-demo"; Flags: runhidden waituntilterminated; StatusMsg: "Loading demo data..."; Check: ShouldSeedDemo

; Scan SOP files from directory if selected
Filename: "{app}\{#MyAppExeName}"; Parameters: "--import-sops ""{code:GetScanSopDirPath}"""; Flags: runhidden waituntilterminated; StatusMsg: "Scanning and importing SOP files (this may take a while)..."; Check: ShouldScanSops

; Scan Document files from directory if selected
Filename: "{app}\{#MyAppExeName}"; Parameters: "--import-docs ""{code:GetScanDocDirPath}"""; Flags: runhidden waituntilterminated; StatusMsg: "Scanning and importing document files (this may take a while)..."; Check: ShouldScanDocs

; Start the service
Filename: "sc.exe"; Parameters: "start DecoSOP"; Flags: runhidden waituntilterminated; StatusMsg: "Starting DecoSOP..."

; Open browser
Filename: "http://localhost:{code:GetPort}"; Flags: postinstall shellexec nowait unchecked; Description: "Open DecoSOP in browser"

[UninstallRun]
Filename: "sc.exe"; Parameters: "stop DecoSOP"; Flags: runhidden; RunOnceId: "StopService"
Filename: "cmd.exe"; Parameters: "/c timeout /t 3 /nobreak >nul"; Flags: runhidden; RunOnceId: "WaitStop"
Filename: "sc.exe"; Parameters: "delete DecoSOP"; Flags: runhidden; RunOnceId: "DeleteService"
Filename: "netsh.exe"; Parameters: "advfirewall firewall delete rule name=""DecoSOP"""; Flags: runhidden; RunOnceId: "RemoveFirewall"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\wwwroot"
Type: files; Name: "{app}\{#MyAppExeName}"
Type: files; Name: "{app}\port.config"
Type: files; Name: "{app}\update-config.json"
; NOTE: database and uploads are intentionally preserved

[Code]
var
  PortPage: TInputQueryWizardPage;
  DatabasePage: TInputOptionWizardPage;
  ImportDbPage: TInputFileWizardPage;
  ImportSopDirPage: TWizardPage;
  ImportSopDirEdit: TNewEdit;
  ImportDocDirPage: TWizardPage;
  ImportDocDirEdit: TNewEdit;
  ScanSopDirPage: TInputDirWizardPage;
  ScanDocDirPage: TWizardPage;
  ScanDocDirEdit: TNewEdit;
  UpdatePage: TInputOptionWizardPage;

// ---- Port helpers ----

function GetPort(Param: String): String;
begin
  if PortPage <> nil then
    Result := PortPage.Values[0]
  else
    Result := '5098';
  if Result = '' then
    Result := '5098';
end;

function IsCustomPort: Boolean;
begin
  Result := GetPort('') <> '5098';
end;

function GetPortConfigContent(Param: String): String;
begin
  Result := 'PORT=' + GetPort('');
end;

// ---- Database helpers ----

function ShouldSeedDemo: Boolean;
begin
  if DatabasePage <> nil then
    Result := DatabasePage.SelectedValueIndex = 1
  else
    Result := False;
end;

function ShouldImportDb: Boolean;
begin
  if DatabasePage <> nil then
    Result := DatabasePage.SelectedValueIndex = 2
  else
    Result := False;
end;

function GetImportDbPath(Param: String): String;
begin
  if ImportDbPage <> nil then
    Result := ImportDbPage.Values[0]
  else
    Result := '';
end;

// ---- Import directory helpers ----

function GetImportSopDirPath(Param: String): String;
begin
  if ImportSopDirEdit <> nil then
    Result := ImportSopDirEdit.Text
  else
    Result := '';
end;

function ShouldImportSopDir: Boolean;
begin
  Result := ShouldImportDb and (GetImportSopDirPath('') <> '') and DirExists(GetImportSopDirPath(''));
end;

function GetImportDocDirPath(Param: String): String;
begin
  if ImportDocDirEdit <> nil then
    Result := ImportDocDirEdit.Text
  else
    Result := '';
end;

function ShouldImportDocDir: Boolean;
begin
  Result := ShouldImportDb and (GetImportDocDirPath('') <> '') and DirExists(GetImportDocDirPath(''));
end;

// ---- Scan directory helpers ----

function ShouldScanFiles: Boolean;
begin
  if DatabasePage <> nil then
    Result := DatabasePage.SelectedValueIndex = 3
  else
    Result := False;
end;

function GetScanSopDirPath(Param: String): String;
begin
  if ScanSopDirPage <> nil then
    Result := ScanSopDirPage.Values[0]
  else
    Result := '';
end;

function ShouldScanSops: Boolean;
begin
  Result := ShouldScanFiles and (GetScanSopDirPath('') <> '') and DirExists(GetScanSopDirPath(''));
end;

function GetScanDocDirPath(Param: String): String;
begin
  if ScanDocDirEdit <> nil then
    Result := ScanDocDirEdit.Text
  else
    Result := '';
end;

function ShouldScanDocs: Boolean;
begin
  Result := ShouldScanFiles and (GetScanDocDirPath('') <> '') and DirExists(GetScanDocDirPath(''));
end;

// ---- Update helpers ----

function IsAutoUpdateEnabled: Boolean;
begin
  if UpdatePage <> nil then
    Result := UpdatePage.SelectedValueIndex = 0
  else
    Result := True;
end;

function GetUpdateConfigContent(Param: String): String;
begin
  if IsAutoUpdateEnabled then
    Result := '{ "enabled": true, "repoOwner": "Susguine", "repoName": "DecoSOP", "checkIntervalHours": 24 }'
  else
    Result := '{ "enabled": false, "repoOwner": "Susguine", "repoName": "DecoSOP", "checkIntervalHours": 24 }';
end;

// ---- Service detection ----

function ServiceExists: Boolean;
var
  ResultCode: Integer;
begin
  Exec('sc.exe', 'query DecoSOP', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Result := (ResultCode = 0);
end;

// ---- Existing install detection ----

function GetUninstallString: String;
var
  UninstallKey: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D3C0-50F1-4A2B-B8E9-DecoSOP-1000}}_is1';
  Result := '';
  RegQueryStringValue(HKLM, UninstallKey, 'UninstallString', Result);
end;

function GetInstalledVersion: String;
var
  UninstallKey: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{D3C0-50F1-4A2B-B8E9-DecoSOP-1000}}_is1';
  Result := '';
  RegQueryStringValue(HKLM, UninstallKey, 'DisplayVersion', Result);
end;

function InitializeSetup: Boolean;
var
  InstalledVersion: String;
  UninstallString: String;
  Msg: String;
  Choice: Integer;
  ResultCode: Integer;
begin
  Result := True;

  InstalledVersion := GetInstalledVersion;
  if InstalledVersion <> '' then
  begin
    if InstalledVersion = '{#MyAppVersion}' then
      Msg := 'DecoSOP v' + InstalledVersion + ' is already installed.' + #13#10 + #13#10 +
             'Yes = Reinstall (keep database and files)' + #13#10 +
             'No = Uninstall (remove application)' + #13#10 +
             'Cancel = Exit setup'
    else
      Msg := 'DecoSOP v' + InstalledVersion + ' is already installed.' + #13#10 + #13#10 +
             'Yes = Upgrade to v{#MyAppVersion} (keep database and files)' + #13#10 +
             'No = Uninstall v' + InstalledVersion + ' (remove application)' + #13#10 +
             'Cancel = Exit setup';

    Choice := MsgBox(Msg, mbConfirmation, MB_YESNOCANCEL or MB_DEFBUTTON1);

    if Choice = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;

    UninstallString := GetUninstallString;

    if Choice = IDNO then
    begin
      // Uninstall: run uninstaller visibly, then exit
      if UninstallString <> '' then
      begin
        if (Length(UninstallString) > 1) and (UninstallString[1] = '"') then
          UninstallString := Copy(UninstallString, 2, Length(UninstallString) - 2);

        Exec(UninstallString, '/NORESTART', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
      end;
      Result := False;
      Exit;
    end;

    // Yes = Upgrade/Reinstall: silently remove old service/firewall, then continue
    if UninstallString <> '' then
    begin
      if (Length(UninstallString) > 1) and (UninstallString[1] = '"') then
        UninstallString := Copy(UninstallString, 2, Length(UninstallString) - 2);

      Exec(UninstallString, '/VERYSILENT /NORESTART /SUPPRESSMSGBOXES', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    end;
  end;
end;

// ---- Browse button handlers for optional directory pages ----

procedure BrowseImportSopDir(Sender: TObject);
var
  Dir: String;
begin
  Dir := ImportSopDirEdit.Text;
  if BrowseForFolder('Select the SOP uploads folder:', Dir, False) then
    ImportSopDirEdit.Text := Dir;
end;

procedure BrowseImportDocDir(Sender: TObject);
var
  Dir: String;
begin
  Dir := ImportDocDirEdit.Text;
  if BrowseForFolder('Select the Document uploads folder:', Dir, False) then
    ImportDocDirEdit.Text := Dir;
end;

procedure BrowseScanDocDir(Sender: TObject);
var
  Dir: String;
begin
  Dir := ScanDocDirEdit.Text;
  if BrowseForFolder('Select the Document source folder:', Dir, False) then
    ScanDocDirEdit.Text := Dir;
end;

// ---- Wizard pages ----

procedure InitializeWizard;
begin
  // Page 1: Port configuration (after directory selection)
  PortPage := CreateInputQueryPage(
    wpSelectDir,
    'Port Configuration',
    'Choose the port DecoSOP will listen on.',
    'Enter the TCP port number (default: 5098).' + #13#10 +
    'Make sure this port is not already in use by another application.' + #13#10 + #13#10 +
    'Other computers on your network will access DecoSOP at:' + #13#10 +
    '  http://your-computer-name:PORT');
  PortPage.Add('Port:', False);
  PortPage.Values[0] := '5098';

  // Page 2: Database setup (after port)
  DatabasePage := CreateInputOptionPage(
    PortPage.ID,
    'Database Setup',
    'Choose how to initialize the database.',
    'DecoSOP stores all SOPs, documents, and categories in a local database.' + #13#10 +
    'Select how you would like to start:',
    True, False);
  DatabasePage.Add('Empty database (start fresh — add your own content)');
  DatabasePage.Add('Demo database (sample SOPs and documents to explore the app)');
  DatabasePage.Add('Import a database backup (.db file from a previous installation)');
  DatabasePage.Add('Import from file directories (scan folders and create categories)');
  DatabasePage.SelectedValueIndex := 0;

  // Page 3: Import file picker (after database — only shown if Import selected)
  ImportDbPage := CreateInputFilePage(
    DatabasePage.ID,
    'Import Database',
    'Select a database backup file to restore.',
    'Choose the .db file from a previous DecoSOP installation or backup.' + #13#10 +
    'You can export a backup from Settings in DecoSOP at any time.');
  ImportDbPage.Add('Database file:', '*.db|*.db', '.db');

  // Page 4: SOP files directory (after db import — only shown if Import selected)
  ImportSopDirPage := CreateCustomPage(
    ImportDbPage.ID,
    'Import SOP Files',
    'Optionally import your SOP upload files.');
  with TNewStaticText.Create(ImportSopDirPage) do
  begin
    Parent := ImportSopDirPage.Surface;
    Caption := 'If you have a previous DecoSOP installation, select the sop-uploads folder' + #13#10 +
               'to restore your uploaded SOP files.' + #13#10 + #13#10 +
               'Leave blank to skip this step.';
    Left := 0;
    Top := 0;
    Width := ImportSopDirPage.SurfaceWidth;
    WordWrap := True;
    AutoSize := True;
  end;
  with TNewStaticText.Create(ImportSopDirPage) do
  begin
    Parent := ImportSopDirPage.Surface;
    Caption := 'SOP uploads folder (e.g. C:\DecoSOP\sop-uploads):';
    Left := 0;
    Top := 76;
  end;
  ImportSopDirEdit := TNewEdit.Create(ImportSopDirPage);
  with ImportSopDirEdit do
  begin
    Parent := ImportSopDirPage.Surface;
    Left := 0;
    Top := 96;
    Width := ImportSopDirPage.SurfaceWidth - 90;
    Text := '';
  end;
  with TNewButton.Create(ImportSopDirPage) do
  begin
    Parent := ImportSopDirPage.Surface;
    Caption := 'Browse...';
    Left := ImportSopDirPage.SurfaceWidth - 85;
    Top := 94;
    Width := 85;
    Height := 25;
    OnClick := @BrowseImportSopDir;
  end;

  // Page 5: Documents directory (after SOP dir — only shown if Import selected)
  ImportDocDirPage := CreateCustomPage(
    ImportSopDirPage.ID,
    'Import Document Files',
    'Optionally import your Document upload files.');
  with TNewStaticText.Create(ImportDocDirPage) do
  begin
    Parent := ImportDocDirPage.Surface;
    Caption := 'If you have a previous DecoSOP installation, select the uploads folder' + #13#10 +
               'to restore your uploaded document files.' + #13#10 + #13#10 +
               'Leave blank to skip this step.';
    Left := 0;
    Top := 0;
    Width := ImportDocDirPage.SurfaceWidth;
    WordWrap := True;
    AutoSize := True;
  end;
  with TNewStaticText.Create(ImportDocDirPage) do
  begin
    Parent := ImportDocDirPage.Surface;
    Caption := 'Document uploads folder (e.g. C:\DecoSOP\uploads):';
    Left := 0;
    Top := 76;
  end;
  ImportDocDirEdit := TNewEdit.Create(ImportDocDirPage);
  with ImportDocDirEdit do
  begin
    Parent := ImportDocDirPage.Surface;
    Left := 0;
    Top := 96;
    Width := ImportDocDirPage.SurfaceWidth - 90;
    Text := '';
  end;
  with TNewButton.Create(ImportDocDirPage) do
  begin
    Parent := ImportDocDirPage.Surface;
    Caption := 'Browse...';
    Left := ImportDocDirPage.SurfaceWidth - 85;
    Top := 94;
    Width := 85;
    Height := 25;
    OnClick := @BrowseImportDocDir;
  end;

  // Page 6: Scan SOP source dir (after import dirs — only shown if Scan selected)
  ScanSopDirPage := CreateInputDirPage(
    ImportDocDirPage.ID,
    'Scan SOP Files',
    'Select a folder containing your SOP files.',
    'DecoSOP will scan the selected folder and its subfolders.' + #13#10 +
    'Subfolders become categories, and matching files (PDF, Word, Excel, etc.)' + #13#10 +
    'are imported as SOPs.' + #13#10 + #13#10 +
    'Archive and temporary folders are automatically skipped.',
    False, '');
  ScanSopDirPage.Add('SOP source folder (e.g. S:\SOPs):');
  ScanSopDirPage.Values[0] := '';

  // Page 7: Scan Documents source dir (after scan SOP — only shown if Scan selected)
  ScanDocDirPage := CreateCustomPage(
    ScanSopDirPage.ID,
    'Scan Document Files',
    'Optionally select a folder containing your document files.');
  with TNewStaticText.Create(ScanDocDirPage) do
  begin
    Parent := ScanDocDirPage.Surface;
    Caption := 'DecoSOP will scan the selected folder and its subfolders.' + #13#10 +
               'Subfolders become categories, and matching files are imported as Documents.' + #13#10 + #13#10 +
               'Leave blank to skip document scanning.';
    Left := 0;
    Top := 0;
    Width := ScanDocDirPage.SurfaceWidth;
    WordWrap := True;
    AutoSize := True;
  end;
  with TNewStaticText.Create(ScanDocDirPage) do
  begin
    Parent := ScanDocDirPage.Surface;
    Caption := 'Document source folder (e.g. S:\Documents):';
    Left := 0;
    Top := 76;
  end;
  ScanDocDirEdit := TNewEdit.Create(ScanDocDirPage);
  with ScanDocDirEdit do
  begin
    Parent := ScanDocDirPage.Surface;
    Left := 0;
    Top := 96;
    Width := ScanDocDirPage.SurfaceWidth - 90;
    Text := '';
  end;
  with TNewButton.Create(ScanDocDirPage) do
  begin
    Parent := ScanDocDirPage.Surface;
    Caption := 'Browse...';
    Left := ScanDocDirPage.SurfaceWidth - 85;
    Top := 94;
    Width := 85;
    Height := 25;
    OnClick := @BrowseScanDocDir;
  end;

  // Page 8: Auto-update preference (after scan dirs)
  UpdatePage := CreateInputOptionPage(
    ScanDocDirPage.ID,
    'Automatic Updates',
    'Choose whether DecoSOP should check for updates.',
    'When enabled, DecoSOP will periodically check GitHub for new releases' + #13#10 +
    'and show a notification in the app when an update is available.' + #13#10 + #13#10 +
    'No data is sent — it only checks the public release page.' + #13#10 +
    'You can change this setting later via the update-config.json file.',
    True, False);
  UpdatePage.Add('Enable automatic update checks (recommended)');
  UpdatePage.Add('Disable automatic update checks');
  UpdatePage.SelectedValueIndex := 0;
end;

// ---- Page visibility + validation ----

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;

  // Skip import backup pages unless "Import backup" is selected
  if (ImportDbPage <> nil) and (PageID = ImportDbPage.ID) then
    Result := not ShouldImportDb;

  if (ImportSopDirPage <> nil) and (PageID = ImportSopDirPage.ID) then
    Result := not ShouldImportDb;

  if (ImportDocDirPage <> nil) and (PageID = ImportDocDirPage.ID) then
    Result := not ShouldImportDb;

  // Skip scan directory pages unless "Import from file directories" is selected
  if (ScanSopDirPage <> nil) and (PageID = ScanSopDirPage.ID) then
    Result := not ShouldScanFiles;

  if (ScanDocDirPage <> nil) and (PageID = ScanDocDirPage.ID) then
    Result := not ShouldScanFiles;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Port: String;
  PortNum: Integer;
  DbFile: String;
begin
  Result := True;

  // Validate port number
  if CurPageID = PortPage.ID then
  begin
    Port := PortPage.Values[0];
    if Port = '' then
    begin
      PortPage.Values[0] := '5098';
      Exit;
    end;

    PortNum := StrToIntDef(Port, 0);
    if (PortNum < 1) or (PortNum > 65535) then
    begin
      MsgBox('Please enter a valid port number between 1 and 65535.', mbError, MB_OK);
      Result := False;
    end;
  end;

  // Validate import file selection
  if (ImportDbPage <> nil) and (CurPageID = ImportDbPage.ID) then
  begin
    DbFile := ImportDbPage.Values[0];
    if DbFile = '' then
    begin
      MsgBox('Please select a database file to import.', mbError, MB_OK);
      Result := False;
    end
    else if not FileExists(DbFile) then
    begin
      MsgBox('The selected file does not exist. Please choose a valid .db file.', mbError, MB_OK);
      Result := False;
    end;
  end;

  // Validate optional import SOP dir (non-empty must exist)
  if (ImportSopDirPage <> nil) and (CurPageID = ImportSopDirPage.ID) then
  begin
    if (ImportSopDirEdit.Text <> '') and not DirExists(ImportSopDirEdit.Text) then
    begin
      MsgBox('The selected folder does not exist. Please choose a valid directory or leave blank to skip.', mbError, MB_OK);
      Result := False;
    end;
  end;

  // Validate optional import doc dir (non-empty must exist)
  if (ImportDocDirPage <> nil) and (CurPageID = ImportDocDirPage.ID) then
  begin
    if (ImportDocDirEdit.Text <> '') and not DirExists(ImportDocDirEdit.Text) then
    begin
      MsgBox('The selected folder does not exist. Please choose a valid directory or leave blank to skip.', mbError, MB_OK);
      Result := False;
    end;
  end;

  // Validate scan SOP directory (required when scan option selected)
  if (ScanSopDirPage <> nil) and (CurPageID = ScanSopDirPage.ID) then
  begin
    if (ScanSopDirPage.Values[0] = '') then
    begin
      MsgBox('Please select a folder containing your SOP files to scan.', mbError, MB_OK);
      Result := False;
    end
    else if not DirExists(ScanSopDirPage.Values[0]) then
    begin
      MsgBox('The selected folder does not exist. Please choose a valid directory.', mbError, MB_OK);
      Result := False;
    end;
  end;

  // Validate optional scan doc dir (non-empty must exist)
  if (ScanDocDirPage <> nil) and (CurPageID = ScanDocDirPage.ID) then
  begin
    if (ScanDocDirEdit.Text <> '') and not DirExists(ScanDocDirEdit.Text) then
    begin
      MsgBox('The selected folder does not exist. Please choose a valid directory or leave blank to skip.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;
