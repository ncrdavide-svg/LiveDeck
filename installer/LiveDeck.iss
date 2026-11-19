; Inno Setup script for LiveDeck.
; Build with installer\build-installer.ps1 (publishes the app, then compiles this script).
; The app is published self-contained, so target PCs do not need .NET. They do need PowerPoint
; or LibreOffice to play presentations: the setup checks and points to LibreOffice if neither
; is installed (PDF files work in any case).

#ifndef AppVersion
  #define AppVersion "0.2.3"
#endif
#ifndef PublishDir
  #define PublishDir "..\artifacts\publish"
#endif

[Setup]
; Keep this AppId fixed forever: it is how new versions recognize and upgrade an existing install.
AppId={{2535BD68-0D2F-4EC7-8B40-B7CD44F43CDC}
AppName=LiveDeck
AppVersion={#AppVersion}
AppVerName=LiveDeck {#AppVersion}
AppPublisher=LiveDeck
DefaultDirName={autopf}\LiveDeck
DefaultGroupName=LiveDeck
DisableProgramGroupPage=yes
SetupIconFile=..\src\LiveDeck.App\Assets\LiveDeck.ico
UninstallDisplayIcon={app}\LiveDeck.exe
OutputDir=..\artifacts\installer
OutputBaseFilename=LiveDeck-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Install for all users (asks for admin), or let the user pick "only for me" (no admin needed).
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
; 64-bit Windows 10/11 only (x64, and ARM64 PCs through x64 emulation).
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
CloseApplications=yes

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Excludes: "*.pdb"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Up to 1.3 the app was called LiveSlides. Same AppId, so an update installs over it (in its
; old folder): the old exe and its shortcuts go, or the user would see two apps.
Type: files; Name: "{app}\LiveSlides.exe"
Type: files; Name: "{app}\LiveSlides.dll"
Type: files; Name: "{app}\LiveSlides.Core.dll"
Type: files; Name: "{autoprograms}\LiveSlides.lnk"
Type: files; Name: "{autodesktop}\LiveSlides.lnk"
Type: files; Name: "{userdesktop}\LiveSlides.lnk"
Type: files; Name: "{userprograms}\LiveSlides.lnk"

[Icons]
Name: "{autoprograms}\LiveDeck"; Filename: "{app}\LiveDeck.exe"
Name: "{autodesktop}\LiveDeck"; Filename: "{app}\LiveDeck.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\LiveDeck.exe"; Description: "{cm:LaunchProgram,LiveDeck}"; Flags: nowait postinstall skipifsilent
Filename: "https://www.libreoffice.org/download/download-libreoffice/"; Description: "Apri la pagina per scaricare LibreOffice (gratuito)"; \
    Flags: shellexec postinstall skipifsilent unchecked; Check: NeedsPresentationProgram

[Code]
// LiveDeck plays presentations with PowerPoint or, without it, LibreOffice (the same checks
// the app makes: PresentationPlayers.Create and PresentationConverter.FindSoffice).
function HasLibreOffice: Boolean;
var
  Dir: String;
begin
  Result :=
    (RegQueryStringValue(HKLM64, 'SOFTWARE\LibreOffice\UNO\InstallPath', '', Dir) and FileExists(AddBackslash(Dir) + 'soffice.exe')) or
    (RegQueryStringValue(HKLM32, 'SOFTWARE\LibreOffice\UNO\InstallPath', '', Dir) and FileExists(AddBackslash(Dir) + 'soffice.exe')) or
    FileExists(ExpandConstant('{commonpf64}\LibreOffice\program\soffice.exe')) or
    FileExists(ExpandConstant('{commonpf32}\LibreOffice\program\soffice.exe'));
end;

function HasPowerPoint: Boolean;
begin
  Result := RegKeyExists(HKCR, 'PowerPoint.Application');
end;

function NeedsPresentationProgram: Boolean;
begin
  Result := not HasPowerPoint and not HasLibreOffice;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpFinished) and NeedsPresentationProgram then
    WizardForm.FinishedLabel.Caption := WizardForm.FinishedLabel.Caption + #13#10#13#10 +
      'Attenzione: su questo PC non ci sono né PowerPoint né LibreOffice. ' +
      'LiveDeck manderà in onda i PDF e le immagini, ma per le presentazioni PowerPoint ' +
      'serve uno dei due. LibreOffice è gratuito: spunta la casella qui sotto per scaricarlo.';
end;
