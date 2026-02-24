; WinTab Inno Setup Installer Script
; Supports: custom install path, .NET 9 runtime detection, bilingual (Chinese/English)

#define AppName "WinTab"
#define AppVersion "1.0.0"
#define AppPublisher "WinTab Contributors"
#define AppURL "https://github.com/user/WinTab"
#define AppExeName "WinTab.exe"

[Setup]
AppId={{B8F3D2A1-7E4C-4D9F-A6B2-1C8E5F0D3A7B}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableDirPage=no
UsePreviousAppDir=yes
AllowNoIcons=yes
LicenseFile=..\LICENSE
OutputDir=..\publish\installer
OutputBaseFilename=WinTab_Setup_{#AppVersion}
; SetupIconFile=..\src\WinTab.App\Assets\wintab.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayName={#AppName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ShowLanguageDialog=yes
LanguageDetectionMethod=locale

[Languages]
Name: "chinesesimplified"; MessagesFile: "Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
chinesesimplified.DotNetRequired=WinTab 需要 .NET 9 桌面运行时。%n%n是否立即下载并安装？
chinesesimplified.DotNetDownloading=正在下载 .NET 9 桌面运行时...
chinesesimplified.LaunchAtStartup=开机自启动
chinesesimplified.RemoveUserDataPrompt=是否同时删除用户配置和日志（AppData\Roaming\WinTab）？%n%n默认保留，选择"是"才会删除。
english.DotNetRequired=WinTab requires .NET 9 Desktop Runtime.%n%nWould you like to download and install it now?
english.DotNetDownloading=Downloading .NET 9 Desktop Runtime...
english.LaunchAtStartup=Run at Windows startup
english.RemoveUserDataPrompt=Do you also want to remove user settings and logs (AppData\Roaming\WinTab)?%n%nDefault is to keep them. Choose Yes to delete.

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startup"; Description: "{cm:LaunchAtStartup}"; Flags: unchecked

[Files]
Source: "..\publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[Registry]
; Add startup entry if task selected
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#AppName}"; ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue; Tasks: startup

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{userappdata}\WinTab"; Check: ShouldRemoveUserData

[UninstallRun]
; Best-effort app-side cleanup for registry handlers/startup entries
Filename: "{app}\{#AppExeName}"; Parameters: "--wintab-cleanup"; Flags: runhidden waituntilterminated skipifdoesntexist; RunOnceId: "WinTabCleanup"
; Remove startup registry entry on uninstall
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v ""{#AppName}"" /f"; Flags: runhidden; RunOnceId: "WinTabStartupRunCleanup"

[Code]
var
  RemoveUserDataOnUninstall: Boolean;

function InitializeUninstall(): Boolean;
var
  Choice: Integer;
begin
  Result := True;
  RemoveUserDataOnUninstall := False;

  if UninstallSilent then
    exit;

  Choice := MsgBox(CustomMessage('RemoveUserDataPrompt'), mbConfirmation, MB_YESNO or MB_DEFBUTTON2);
  RemoveUserDataOnUninstall := Choice = IDYES;
end;

function ShouldRemoveUserData(): Boolean;
begin
  Result := RemoveUserDataOnUninstall;
end;

function IsDotNet9DesktopInstalled: Boolean;
var
  ResultCode: Integer;
begin
  Result := False;
  // Check via dotnet --list-runtimes for Microsoft.WindowsDesktop.App 9.x
  if Exec('dotnet', '--list-runtimes', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    Result := (ResultCode = 0);
  end;

  // Alternative: check registry
  if not Result then
  begin
    Result := RegKeyExists(HKLM, 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App');
  end;
end;

function InitializeSetup: Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;

  if not IsDotNet9DesktopInstalled then
  begin
    if MsgBox(CustomMessage('DotNetRequired'), mbConfirmation, MB_YESNO) = IDYES then
    begin
      ShellExec('open', 'https://dotnet.microsoft.com/en-us/download/dotnet/9.0', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
      Result := False; // Exit setup, let user install .NET first
    end
    else
    begin
      Result := False;
    end;
  end;
end;
