[Setup]
AppName=UmlagerungsTool
AppVersion=2.0.0
DefaultDirName={pf}\UmlagerungsTool
DefaultGroupName=UmlagerungsTool
OutputBaseFilename=UmlagerungsTool2.0.0_Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableDirPage=no
DisableProgramGroupPage=no
CloseApplications=yes
RestartApplications=no
PrivilegesRequired=admin
UninstallDisplayIcon={app}\UmlagerungsTool.exe

[Tasks]
Name: desktopicon; Description: "Desktop-Verknüpfung erstellen"; GroupDescription: "Optionale Verknüpfungen"
Name: startmenuicon; Description: "Startmenü-Verknüpfung erstellen"

[InstallDelete]
Type: filesandordirs; Name: "{app}\*"

[Files]
Source: "..\dist\UmlagerungsTool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{commondesktop}\UmlagerungsTool"; Filename: "{app}\UmlagerungsTool.exe"; WorkingDir: "{app}"; Tasks: desktopicon
Name: "{group}\UmlagerungsTool"; Filename: "{app}\UmlagerungsTool.exe"; WorkingDir: "{app}"; Tasks: startmenuicon

[Run]
Filename: "{app}\UmlagerungsTool.exe"; Description: "UmlagerungsTool starten"; Flags: nowait postinstall skipifsilent
