#define MyAppName "ระบบขอ สด.8 ใหม่"
#define MyAppVersion "0.9.8"
#define MyAppPublisher "หน่วยสัสดีอำเภอวานรนิวาส"
#define MyAppExeName "SD8System.exe"

[Setup]
AppId={{8F2C1A40-6B2E-4C7E-9A6D-SD8SYSTEM001}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=SD8System_Setup
VersionInfoVersion={#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
SetupIconFile=favicon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "thai"; MessagesFile: "compiler:Languages\Thai.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "สร้างไอคอนบนหน้าจอ (Desktop)"; GroupDescription: "ทางลัด:"

[Files]
Source: "dist\SD8System\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "เปิดโปรแกรมทันที"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\__pycache__"