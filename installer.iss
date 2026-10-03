; Script generated for Inno Setup 6
; Team PDF Tools - Desktop Installer Script

#define MyAppName "Team PDF Tools"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Team PDF Tools"
#define MyAppExeName "TeamPDFTools.exe"
#define ProjectDir "C:\Users\Deepak Pawar\Desktop\Automation Projects\Team_PDF_Tools"

[Setup]
AppId={{9F82A4C1-4D3B-4E87-92B5-1C6A9E273410}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir={#ProjectDir}\dist_installer
OutputBaseFilename=TeamPDFTools_Setup
SetupIconFile={#ProjectDir}\app_icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
; Allow installation without administrator rights so any team member can install seamlessly
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#ProjectDir}\dist\TeamPDFTools\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
