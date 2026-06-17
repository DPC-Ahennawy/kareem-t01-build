; Kareem T-01 - EDECS Contract Processor - Inno Setup installer
; Produces Kareem_T01_Setup.exe. Installs PER-USER to LocalAppData (no admin, writable).
#define AppName "Kareem T-01"
#define AppPub  "DPC Department - EDECS_EGY"
#define AppVer  "1.0.0"

[Setup]
AppName={#AppName}
AppVersion={#AppVer}
AppPublisher={#AppPub}
; ---- per-user, writable, NOT Program Files ----
PrivilegesRequired=lowest
DefaultDirName={localappdata}\EDECS\Kareem T-01
DisableProgramGroupPage=yes
DefaultGroupName=Kareem T-01
OutputBaseFilename=Kareem_T01_Setup
OutputDir=Output
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName=Kareem T-01 - EDECS Contract Processor

[Files]
; The whole PyInstaller one-folder build (KareemT01.exe + _internal\)
Source: "..\dist\KareemT01\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs
; Bundled OCR (real tesseract.exe + DLLs + ara/eng/fra) placed beside the exe
Source: "..\ocr\*";          DestDir: "{app}\ocr"; Flags: recursesubdirs createallsubdirs
; Packaged config that FORCES bundled OCR
Source: "config.packaged.json"; DestDir: "{app}"; DestName: "config.json"; Flags: onlyifdoesntexist
; Reference data + web frontend (also bundled by PyInstaller, but keep on-disk copies)
Source: "..\refs.json";        DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "..\recipients.json";  DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "..\web\frontend\index.html"; DestDir: "{app}\web\frontend"; Flags: onlyifdoesntexist
[Dirs]
Name: "{app}\Inbox";  Permissions: users-modify
Name: "{app}\Output"; Permissions: users-modify

[Icons]
Name: "{group}\Kareem T-01";              Filename: "{app}\KareemT01.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall Kareem T-01";    Filename: "{uninstallexe}"
Name: "{userdesktop}\Kareem T-01";        Filename: "{app}\KareemT01.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Run]
; Launch the app at the end of install; it serves http://127.0.0.1:8000/ and opens the browser.
Filename: "{app}\KareemT01.exe"; Description: "Launch Kareem T-01"; Flags: nowait postinstall skipifsilent

[Code]
// Compile-time guard: refuse to BUILD the installer unless real OCR is present.
#if !FileExists("..\ocr\tesseract\tesseract.exe")
  #error Bundled OCR missing: ..\ocr\tesseract\tesseract.exe not found. Run packaging\fetch_ocr_full.ps1 first (Route B).
#endif
#if !FileExists("..\ocr\tesseract\tessdata\ara.traineddata")
  #error Bundled OCR missing: ara.traineddata. Run packaging\fetch_ocr_full.ps1 first.
#endif

procedure CurStepChanged(CurStep: TSetupStep);
begin
  // Run-time sanity: confirm OCR landed in the install dir.
  if CurStep = ssPostInstall then
  begin
    if not FileExists(ExpandConstant('{app}\ocr\tesseract\tesseract.exe')) then
      MsgBox('Warning: bundled tesseract.exe was not found in the install folder. '
        + 'OCR may not work.', mbError, MB_OK);
  end;
end;
