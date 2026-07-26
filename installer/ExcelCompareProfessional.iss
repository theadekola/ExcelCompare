#define AppName "Excel Compare Professional"
#define AppVersion "2.0.4"
#define AppPublisher "AAT-Tech Ltd"
#define AppExe "ExcelCompareProfessional.exe"
[Setup]
AppId={{F6BC73CC-A850-43BA-A17C-D1F7A5F72110}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Installer
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\AAT-Tech Ltd\Excel Compare Professional
DefaultGroupName={#AppName}
OutputDir=..\dist
OutputBaseFilename=ExcelCompareProfessional-Setup-v2.0.4
SetupIconFile=..\assets\excelcompare.ico
UninstallDisplayIcon={app}\{#AppExe}
LicenseFile=..\licenses\LICENCE.txt
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"
[Files]
Source: "..\package\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\docs\Excel_Compare_Professional_User_Manual.pdf"; DestDir: "{app}\Documentation"; Flags: ignoreversion
Source: "..\licenses\LICENCE.txt"; DestDir: "{app}"; Flags: ignoreversion
[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{autoprograms}\{#AppName} User Manual"; Filename: "{app}\Documentation\Excel_Compare_Professional_User_Manual.pdf"
Name: "{autoprograms}\{#AppName}\Install Legacy Format Support"; Filename: "{app}\tools\install-libreoffice.bat"; WorkingDir: "{app}\tools"
[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\tools\install-libreoffice.bat"; Description: "Install optional conversion support (.xlsb, .ods, .fods, .xlt)"; Flags: postinstall skipifsilent unchecked
