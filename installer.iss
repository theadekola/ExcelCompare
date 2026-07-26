#define AppName "Excel Compare"
#define AppVersion "1.3.0"
[Setup]
AppId={{B182DAE1-8267-4EB4-A8DE-70AB8E201AC1}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\Excel Compare
DefaultGroupName=Excel Compare
OutputDir=installer-output
OutputBaseFilename=ExcelCompare-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
[Files]
Source: "package\*"; DestDir: "{app}"; Flags: recursesubdirs
[Icons]
Name: "{autoprograms}\Excel Compare"; Filename: "{app}\ExcelCompare.exe"
Name: "{autodesktop}\Excel Compare"; Filename: "{app}\ExcelCompare.exe"
[Run]
Filename: "{app}\ExcelCompare.exe"; Flags: nowait postinstall skipifsilent
