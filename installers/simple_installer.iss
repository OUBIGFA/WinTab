[Setup]
AppId={{E1F2A3C4-F5D4-3B2E-1D5A-8F7B2F8D3E7A}
AppName=WinTab Explorer Utility
AppVersion=2.5.0
AppPublisher=w4po
DefaultDirName={autopf}\WinTab
DefaultGroupName=WinTab
DisableDirPage=no
UsePreviousAppDir=yes
OutputDir=..\installers
OutputBaseFilename=WinTab_Setup
SetupIconFile=..\WinTab\Icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\WinTab"; Filename: "{app}\WinTab.exe"
Name: "{autodesktop}\WinTab"; Filename: "{app}\WinTab.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\WinTab.exe"; Description: "{cm:LaunchProgram,WinTab}"; Flags: nowait postinstall skipifsilent
