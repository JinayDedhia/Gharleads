[Setup]
AppName=GharLeads
AppVersion=0.4
AppPublisher=GharLeads
DefaultDirName={autopf}\GharLeads
DefaultGroupName=GharLeads
OutputDir=installer_output
OutputBaseFilename=GharLeads_Setup_0.4
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\GharLeads.exe
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"

[Files]
Source: "dist\GharLeads\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\GharLeads"; Filename: "{app}\GharLeads.exe"
Name: "{autodesktop}\GharLeads"; Filename: "{app}\GharLeads.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\GharLeads.exe"; Description: "Launch GharLeads"; Flags: nowait postinstall skipifsilent
