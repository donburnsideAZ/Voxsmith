; Voxsmith v2.2 - Inno Setup Installer Script
; Compile with Inno Setup 6.x: https://jrsoftware.org/isinfo.php

#define MyAppName "Voxsmith"
#define MyAppVersion "2.2"
#define MyAppPublisher "Don Burnside"
#define MyAppURL "https://donburnside.com/voxsmith"
#define MyAppExeName "Voxsmith.exe"

[Setup]
; Unique app ID - DO NOT CHANGE after first release
AppId={{8F4E9C2A-5B7D-4E1F-A3C8-9D2B6F4E8A1C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Install to LocalAppData to avoid OneDrive/cloud sync issues
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}

; No admin required - installs to user folder
PrivilegesRequired=lowest

; Output settings
OutputDir=installer_output
OutputBaseFilename=VoxsmithSetup_v{#MyAppVersion}
SetupIconFile=vs_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Compression
Compression=lzma2/max
SolidCompression=yes

; Modern installer look
WizardStyle=modern
WizardResizable=no

; Disable program group page (simpler install)
DisableProgramGroupPage=yes

; Show "Launch Voxsmith" checkbox at end
DisableFinishedPage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main application folder from PyInstaller dist output
Source: "dist\Voxsmith\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu shortcut
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Desktop shortcut (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Launch after install (optional checkbox)
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up logs and settings on uninstall (optional - comment out to preserve settings)
; Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\logs"
; Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\temp"

[Code]
// Check if app is running before install/uninstall
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  // Could add check for running instance here if needed
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  // Could add check for running instance here if needed
end;
