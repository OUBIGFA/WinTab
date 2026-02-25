; WinTab Inno Setup Installer Script
; Supports: custom install path, .NET 9 runtime detection, bilingual (Chinese/English)

#define AppName "WinTab"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
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
SetupIconFile=..\src\WinTab.App\Assets\wintab.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayName={#AppName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ShowLanguageDialog=yes
LanguageDetectionMethod=locale
AppMutex=WinTab_SingleInstance
CloseApplications=yes
CloseApplicationsFilter=WinTab.exe
RestartApplications=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
chinesesimplified.DotNetRequired=WinTab 需要 .NET 9 桌面运行时。%n%n是否立即下载并安装？
chinesesimplified.DotNetDownloading=正在下载 .NET 9 桌面运行时...
chinesesimplified.LaunchAtStartup=开机自启动
chinesesimplified.ReinstallModeTitle=检测到已安装的 WinTab
chinesesimplified.ReinstallModeDescription=请选择安装模式。
chinesesimplified.ReinstallModeClean=卸载完再安装（推荐）
chinesesimplified.ReinstallModeDirect=不卸载直接安装
chinesesimplified.ReinstallRemoveUserData=卸载时删除用户数据（AppData\Roaming\WinTab）
chinesesimplified.ReinstallRemoveUserDataHint=默认不勾选。仅在需要彻底清理时勾选。
chinesesimplified.ReinstallUninstallerMissing=检测到已安装版本，但找不到卸载程序，无法执行“卸载完再安装”。请改选“直接安装”或先手动卸载。
chinesesimplified.ReinstallUninstallLaunchFailed=启动旧版本卸载程序失败：%1
chinesesimplified.ReinstallUninstallFailed=旧版本卸载未完成（退出码 %1）。安装已取消。
chinesesimplified.ReinstallUninstallIncomplete=旧版本卸载后仍检测到安装信息。请稍后重试，或先手动卸载。
english.DotNetRequired=WinTab requires .NET 9 Desktop Runtime.%n%nWould you like to download and install it now?
english.DotNetDownloading=Downloading .NET 9 Desktop Runtime...
english.LaunchAtStartup=Run at Windows startup
english.ReinstallModeTitle=Existing WinTab installation detected
english.ReinstallModeDescription=Choose how to continue.
english.ReinstallModeClean=Uninstall then install (recommended)
english.ReinstallModeDirect=Install directly without uninstall
english.ReinstallRemoveUserData=Remove user data during uninstall (AppData\Roaming\WinTab)
english.ReinstallRemoveUserDataHint=Unchecked by default. Enable only for full cleanup.
english.ReinstallUninstallerMissing=An existing installation was detected, but no uninstaller was found. Cannot continue with "uninstall then install". Choose direct install or uninstall manually first.
english.ReinstallUninstallLaunchFailed=Failed to start existing uninstaller: %1
english.ReinstallUninstallFailed=Existing uninstall did not complete (exit code %1). Setup was cancelled.
english.ReinstallUninstallIncomplete=Uninstall completed but existing install information is still present. Please retry later or uninstall manually first.

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
  ExistingInstallDetected: Boolean;
  ExistingUninstallerPath: String;
  ExistingUninstallArgs: String;
  ReinstallModePage: TWizardPage;
  ReinstallCleanRadio: TRadioButton;
  ReinstallDirectRadio: TRadioButton;
  ReinstallRemoveUserDataCheckbox: TNewCheckBox;
  ReinstallHintText: TNewStaticText;
  SelectedReinstallMode: String;

const
  ExistingUninstallRegSubKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B8F3D2A1-7E4C-4D9F-A6B2-1C8E5F0D3A7B}_is1';

function ContainsTextIgnoreCase(const Source, Part: String): Boolean;
begin
  Result := Pos(Uppercase(Part), Uppercase(Source)) > 0;
end;

function AppendArgument(const ExistingArgs, NewArg: String): String;
begin
  if Trim(NewArg) = '' then
  begin
    Result := Trim(ExistingArgs);
    exit;
  end;

  if ContainsTextIgnoreCase(ExistingArgs, NewArg) then
  begin
    Result := Trim(ExistingArgs);
    exit;
  end;

  if Trim(ExistingArgs) = '' then
    Result := NewArg
  else
    Result := Trim(ExistingArgs) + ' ' + NewArg;
end;

function ExtractCommandPathAndArgs(const CommandLine: String; var FileName, Params: String): Boolean;
var
  S: String;
  NextQuote: Integer;
  SpacePos: Integer;
begin
  Result := False;
  FileName := '';
  Params := '';

  S := Trim(CommandLine);
  if S = '' then
    exit;

  if S[1] = '"' then
  begin
    NextQuote := 2;
    while (NextQuote <= Length(S)) and (S[NextQuote] <> '"') do
      NextQuote := NextQuote + 1;

    if NextQuote > Length(S) then
      exit;

    FileName := Copy(S, 2, NextQuote - 2);
    Params := Trim(Copy(S, NextQuote + 1, MaxInt));
    Result := Trim(FileName) <> '';
    exit;
  end;

  SpacePos := Pos(' ', S);
  if SpacePos > 0 then
  begin
    FileName := Copy(S, 1, SpacePos - 1);
    Params := Trim(Copy(S, SpacePos + 1, MaxInt));
  end
  else
  begin
    FileName := S;
    Params := '';
  end;

  Result := Trim(FileName) <> '';
end;

function TryReadUninstallValue(const ValueName: String; var ValueData: String): Boolean;
begin
  Result :=
    RegQueryStringValue(HKLM, ExistingUninstallRegSubKey, ValueName, ValueData) or
    RegQueryStringValue(HKCU, ExistingUninstallRegSubKey, ValueName, ValueData);
end;

function DetectExistingInstall(var UninstallerPath, UninstallArgs: String): Boolean;
var
  UninstallString: String;
  QuietUninstallString: String;
  ParsedPath: String;
  ParsedArgs: String;
begin
  UninstallerPath := '';
  UninstallArgs := '';

  Result :=
    RegKeyExists(HKLM, ExistingUninstallRegSubKey) or
    RegKeyExists(HKCU, ExistingUninstallRegSubKey);

  if not Result then
    exit;

  QuietUninstallString := '';
  UninstallString := '';

  if TryReadUninstallValue('QuietUninstallString', QuietUninstallString) and
     ExtractCommandPathAndArgs(QuietUninstallString, ParsedPath, ParsedArgs) then
  begin
    UninstallerPath := ParsedPath;
    UninstallArgs := ParsedArgs;
    exit;
  end;

  if TryReadUninstallValue('UninstallString', UninstallString) and
     ExtractCommandPathAndArgs(UninstallString, ParsedPath, ParsedArgs) then
  begin
    UninstallerPath := ParsedPath;
    UninstallArgs := ParsedArgs;
  end;
end;

function WaitForExistingInstallRemoval(const TimeoutMs: Integer): Boolean;
var
  RemainingWaitMs: Integer;
  DummyPath: String;
  DummyArgs: String;
begin
  RemainingWaitMs := TimeoutMs;

  while RemainingWaitMs >= 0 do
  begin
    if not DetectExistingInstall(DummyPath, DummyArgs) then
    begin
      Result := True;
      exit;
    end;

    Sleep(500);
    RemainingWaitMs := RemainingWaitMs - 500;
  end;

  Result := not DetectExistingInstall(DummyPath, DummyArgs);
end;

function ResolveRemoveUserDataChoiceForSetup(): Boolean;
var
  TailUpper: String;
begin
  Result := False;
  TailUpper := Uppercase(GetCmdTail());
  if Pos('/REMOVEUSERDATA=1', TailUpper) > 0 then
    Result := True;
end;

function ResolveReinstallModeForSetup(): String;
var
  TailUpper: String;
begin
  TailUpper := Uppercase(GetCmdTail());

  if Pos('/REINSTALLMODE=CLEAN', TailUpper) > 0 then
  begin
    Result := 'clean';
    exit;
  end;

  if Pos('/REINSTALLMODE=DIRECT', TailUpper) > 0 then
  begin
    Result := 'direct';
    exit;
  end;

  if WizardSilent then
    Result := 'direct'
  else
    Result := 'clean';
end;

procedure UpdateReinstallModePageState();
begin
  if ReinstallRemoveUserDataCheckbox = nil then
    exit;

  ReinstallRemoveUserDataCheckbox.Enabled := ReinstallCleanRadio.Checked;
  if not ReinstallCleanRadio.Checked then
    ReinstallRemoveUserDataCheckbox.Checked := False;
end;

procedure ReinstallModeSelectionChanged(Sender: TObject);
begin
  UpdateReinstallModePageState();
end;

procedure InitializeWizard();
begin
  ExistingInstallDetected := DetectExistingInstall(ExistingUninstallerPath, ExistingUninstallArgs);
  SelectedReinstallMode := ResolveReinstallModeForSetup();
  RemoveUserDataOnUninstall := ResolveRemoveUserDataChoiceForSetup();

  if not ExistingInstallDetected then
    exit;

  ReinstallModePage := CreateCustomPage(
    wpSelectDir,
    ExpandConstant('{cm:ReinstallModeTitle}'),
    ExpandConstant('{cm:ReinstallModeDescription}'));

  ReinstallCleanRadio := TRadioButton.Create(WizardForm);
  ReinstallCleanRadio.Parent := ReinstallModePage.Surface;
  ReinstallCleanRadio.Left := 0;
  ReinstallCleanRadio.Top := ScaleY(8);
  ReinstallCleanRadio.Width := ReinstallModePage.SurfaceWidth;
  ReinstallCleanRadio.Height := ScaleY(20);
  ReinstallCleanRadio.Caption := ExpandConstant('{cm:ReinstallModeClean}');
  ReinstallCleanRadio.Checked := True;
  ReinstallCleanRadio.OnClick := @ReinstallModeSelectionChanged;

  ReinstallDirectRadio := TRadioButton.Create(WizardForm);
  ReinstallDirectRadio.Parent := ReinstallModePage.Surface;
  ReinstallDirectRadio.Left := 0;
  ReinstallDirectRadio.Top := ReinstallCleanRadio.Top + ReinstallCleanRadio.Height + ScaleY(8);
  ReinstallDirectRadio.Width := ReinstallModePage.SurfaceWidth;
  ReinstallDirectRadio.Height := ScaleY(20);
  ReinstallDirectRadio.Caption := ExpandConstant('{cm:ReinstallModeDirect}');
  ReinstallDirectRadio.Checked := False;
  ReinstallDirectRadio.OnClick := @ReinstallModeSelectionChanged;

  ReinstallRemoveUserDataCheckbox := TNewCheckBox.Create(WizardForm);
  ReinstallRemoveUserDataCheckbox.Parent := ReinstallModePage.Surface;
  ReinstallRemoveUserDataCheckbox.Left := ScaleX(20);
  ReinstallRemoveUserDataCheckbox.Top := ReinstallDirectRadio.Top + ReinstallDirectRadio.Height + ScaleY(16);
  ReinstallRemoveUserDataCheckbox.Width := ReinstallModePage.SurfaceWidth - ScaleX(20);
  ReinstallRemoveUserDataCheckbox.Height := ScaleY(20);
  ReinstallRemoveUserDataCheckbox.Caption := ExpandConstant('{cm:ReinstallRemoveUserData}');
  ReinstallRemoveUserDataCheckbox.Checked := False;

  ReinstallHintText := TNewStaticText.Create(WizardForm);
  ReinstallHintText.Parent := ReinstallModePage.Surface;
  ReinstallHintText.Left := ScaleX(40);
  ReinstallHintText.Top := ReinstallRemoveUserDataCheckbox.Top + ReinstallRemoveUserDataCheckbox.Height + ScaleY(2);
  ReinstallHintText.Width := ReinstallModePage.SurfaceWidth - ScaleX(40);
  ReinstallHintText.Height := ScaleY(32);
  ReinstallHintText.AutoSize := False;
  ReinstallHintText.WordWrap := True;
  ReinstallHintText.Caption := ExpandConstant('{cm:ReinstallRemoveUserDataHint}');

  if SelectedReinstallMode = 'direct' then
  begin
    ReinstallDirectRadio.Checked := True;
    ReinstallCleanRadio.Checked := False;
  end;

  ReinstallRemoveUserDataCheckbox.Checked := RemoveUserDataOnUninstall;
  UpdateReinstallModePageState();
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;

  if (ReinstallModePage <> nil) and (PageID = ReinstallModePage.ID) then
    Result := (not ExistingInstallDetected) or WizardSilent;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;

  if (ReinstallModePage <> nil) and (CurPageID = ReinstallModePage.ID) then
  begin
    if ReinstallDirectRadio.Checked then
      SelectedReinstallMode := 'direct'
    else
      SelectedReinstallMode := 'clean';

    RemoveUserDataOnUninstall :=
      ReinstallCleanRadio.Checked and ReinstallRemoveUserDataCheckbox.Checked;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  LaunchParams: String;
  ResultCode: Integer;
begin
  Result := '';

  if not ExistingInstallDetected then
    exit;

  if SelectedReinstallMode <> 'clean' then
    exit;

  if Trim(ExistingUninstallerPath) = '' then
  begin
    Result := CustomMessage('ReinstallUninstallerMissing');
    exit;
  end;

  LaunchParams := ExistingUninstallArgs;

  if WizardSilent then
    LaunchParams := AppendArgument(LaunchParams, '/VERYSILENT')
  else
    LaunchParams := AppendArgument(LaunchParams, '/SILENT');

  LaunchParams := AppendArgument(LaunchParams, '/SUPPRESSMSGBOXES');
  LaunchParams := AppendArgument(LaunchParams, '/NORESTART');

  if RemoveUserDataOnUninstall then
    LaunchParams := AppendArgument(LaunchParams, '/REMOVEUSERDATA=1');

  if not Exec(ExistingUninstallerPath, LaunchParams, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then
  begin
    Result := FmtMessage(CustomMessage('ReinstallUninstallLaunchFailed'), [SysErrorMessage(ResultCode)]);
    exit;
  end;

  if ResultCode <> 0 then
  begin
    Result := FmtMessage(CustomMessage('ReinstallUninstallFailed'), [IntToStr(ResultCode)]);
    exit;
  end;

  if not WaitForExistingInstallRemoval(20000) then
    Result := CustomMessage('ReinstallUninstallIncomplete');
end;

function ShouldRemoveUserData(): Boolean;
var
  TailUpper: String;
begin
  TailUpper := Uppercase(GetCmdTail());
  if Pos('/REMOVEUSERDATA=1', TailUpper) > 0 then
    Result := True
  else
    Result := RemoveUserDataOnUninstall;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  RemoveUserDataOnUninstall := ShouldRemoveUserData;
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
