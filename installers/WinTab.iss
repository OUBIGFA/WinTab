; WinTab Inno Setup Installer Script
; Supports: custom install path, .NET 9 runtime detection, bilingual (Chinese/English)

#define AppName "WinTab"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#define AppPublisher "WinTab Contributors"
#define AppURL "https://github.com/user/WinTab"
#define AppExeName "WinTab.exe"
#define DelegateExecuteClsid "{{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}}"

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
OutputDir=..\publish
OutputBaseFilename=WinTab_Setup_{#AppVersion}
SetupIconFile=..\src\WinTab.App\Assets\wintab.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayName=UninsWinTab
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
Source: "..\publish\win-x64\*"; Excludes: "portable.txt"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\UninsWinTab"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[Registry]
; Add startup entry if task selected
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#AppName}"; ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue; Tasks: startup

; Register DelegateExecute COM host for Explorer open verb interception
; HKLM (machine-wide) - REQUIRED for Windows 11 Start Menu and third-party apps
Root: HKLM; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}"; ValueType: string; ValueData: "WinTab Open Folder DelegateExecute"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueData: "{app}\WinTab.ShellBridge.comhost.dll"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"
Root: HKLM32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}"; ValueType: string; ValueData: "WinTab Open Folder DelegateExecute"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueData: "{app}\x86\WinTab.ShellBridge.comhost.dll"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"

; HKCU (user-only) - for legacy single-user scenarios and same-user Explorer
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}"; ValueType: string; ValueData: "WinTab Open Folder DelegateExecute"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueData: "{app}\WinTab.ShellBridge.comhost.dll"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}"; ValueType: string; ValueData: "WinTab Open Folder DelegateExecute"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueData: "{app}\x86\WinTab.ShellBridge.comhost.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#DelegateExecuteClsid}\InProcServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{userappdata}\WinTab"; Check: ShouldRemoveUserData

[UninstallRun]
; Best-effort app-side cleanup for registry handlers/startup entries
; NOTE: This runs AFTER [UninstallDelete] has already removed files,
; so skipifdoesntexist means the exe-based cleanup is silently skipped.
; The real open-verb restore happens in CurUninstallStepChanged(usUninstall)
; which runs BEFORE file deletion.
Filename: "{app}\{#AppExeName}"; Parameters: "--wintab-cleanup"; Flags: runhidden waituntilterminated skipifdoesntexist; RunOnceId: "WinTabCleanup"
; Remove startup registry entry on uninstall
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v ""{#AppName}"" /f"; Flags: runhidden; RunOnceId: "WinTabStartupRunCleanup"
; Remove HKLM COM registration (machine-wide) - 64-bit and 32-bit
Filename: "reg.exe"; Parameters: "delete ""HKLM\Software\Classes\CLSID\{#DelegateExecuteClsid}"" /f"; Flags: runhidden; RunOnceId: "WinTabDelegateExecuteCleanupHKLM64"
Filename: "reg.exe"; Parameters: "delete ""HKLM\Software\Classes\CLSID\{#DelegateExecuteClsid}"" /f /reg:32"; Flags: runhidden; RunOnceId: "WinTabDelegateExecuteCleanupHKLM32"
; Remove HKCU COM registration (user-only) - 64-bit and 32-bit
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Classes\CLSID\{#DelegateExecuteClsid}"" /f /reg:64"; Flags: runhidden; RunOnceId: "WinTabDelegateExecuteCleanup64"
Filename: "reg.exe"; Parameters: "delete ""HKCU\Software\Classes\CLSID\{#DelegateExecuteClsid}"" /f /reg:32"; Flags: runhidden; RunOnceId: "WinTabDelegateExecuteCleanup32"

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

function GetExistingInstallAppExePath(): String;
begin
  Result := '';

  if Trim(ExistingUninstallerPath) = '' then
    exit;

  Result := AddBackslash(ExtractFileDir(ExistingUninstallerPath)) + '{#AppExeName}';
end;

procedure StopExistingShellBridgeHostsForUpgrade();
var
  ExistingAppExePath: String;
  ResultCode: Integer;
begin
  if not ExistingInstallDetected then
    exit;

  // Try to run cleanup from existing installation if files still exist
  ExistingAppExePath := GetExistingInstallAppExePath();
  if FileExists(ExistingAppExePath) then
  begin
    Exec(
      ExistingAppExePath,
      '--wintab-cleanup',
      '',
      SW_HIDE,
      ewWaitUntilTerminated,
      ResultCode);
  end;

  // Kill any remaining WinTab processes
  Exec(
    ExpandConstant('{sys}\taskkill.exe'),
    '/IM {#AppExeName} /F /T',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode);
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

procedure RestoreExplorerOpenVerbDefaultsViaReg();
var
  ResultCode: Integer;
  DelegateExecuteClsid: String;
begin
  DelegateExecuteClsid := '{FD5BF2CD-0B24-4A80-9AF3-E40F9AFC0001}';

  // Remove WinTab open-verb overrides from HKCU for all target classes.
  // Called during reinstall cleanup to ensure Explorer doesn't try to load
  // deleted WinTab executables when opening folders.
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Folder\shell\open\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Folder\shell\explore\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Folder\shell\opennewwindow\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Directory\shell\open\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Directory\shell\explore\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Directory\shell\opennewwindow\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Drive\shell\open\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Drive\shell\explore\command');
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\Drive\shell\opennewwindow\command');

  // Restore safe default verbs.
  RegWriteStringValue(HKCU, 'Software\Classes\Folder\shell', '', 'open');
  RegWriteStringValue(HKCU, 'Software\Classes\Directory\shell', '', 'none');
  RegWriteStringValue(HKCU, 'Software\Classes\Drive\shell', '', 'none');

  // Remove DelegateExecute CLSID from HKCU (both views via reg.exe).
  Exec('reg.exe', 'delete "HKCU\Software\Classes\CLSID\' + DelegateExecuteClsid + '" /f /reg:64', '', SW_HIDE, ewNoWait, ResultCode);
  Exec('reg.exe', 'delete "HKCU\Software\Classes\CLSID\' + DelegateExecuteClsid + '" /f /reg:32', '', SW_HIDE, ewNoWait, ResultCode);

  // Remove DelegateExecute CLSID from HKLM (both views via reg.exe).
  Exec('reg.exe', 'delete "HKLM\Software\Classes\CLSID\' + DelegateExecuteClsid + '" /f', '', SW_HIDE, ewNoWait, ResultCode);
  Exec('reg.exe', 'delete "HKLM\Software\Classes\CLSID\' + DelegateExecuteClsid + '" /f /reg:32', '', SW_HIDE, ewNoWait, ResultCode);

  // Remove WinTab backup registry entries.
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\WinTab\Backups\ExplorerOpenVerb');
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  LaunchParams: String;
  ResultCode: Integer;
begin
  Result := '';

  if not ExistingInstallDetected then
    exit;

  if SelectedReinstallMode = 'clean' then
  begin
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

  if Result <> '' then
    exit;

  if SelectedReinstallMode = 'direct' then
    StopExistingShellBridgeHostsForUpgrade();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  // Run open-verb registry restore during standalone uninstall.
  // This ensures Explorer doesn't freeze when opening folders after WinTab is removed.
  if CurUninstallStep = usUninstall then
    RestoreExplorerOpenVerbDefaultsViaReg();
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

function StartsWithIgnoreCase(const Source, Prefix: String): Boolean;
begin
  Result := CompareText(Copy(Source, 1, Length(Prefix)), Prefix) = 0;
end;

function DotNetListContainsDesktop9Runtime: Boolean;
var
  ResultCode: Integer;
begin
  Result := False;

  if not Exec(
    ExpandConstant('{cmd}'),
    '/C dotnet --list-runtimes | findstr /R /C:"^Microsoft.WindowsDesktop.App 9\." >nul',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode) then
  begin
    exit;
  end;

  Result := (ResultCode = 0);
end;

function HasDotNet9DesktopRuntimeInRegistry(const RuntimeKeyPath: String): Boolean;
var
  Versions: TArrayOfString;
  Index: Integer;
begin
  Result := False;
  if not RegGetSubkeyNames(HKLM, RuntimeKeyPath, Versions) then
    exit;

  for Index := 0 to GetArrayLength(Versions) - 1 do
  begin
    if StartsWithIgnoreCase(Versions[Index], '9.') then
    begin
      Result := True;
      exit;
    end;
  end;
end;

function IsDotNet9DesktopInstalled: Boolean;
begin
  Result := False;

  // Check via dotnet --list-runtimes for Microsoft.WindowsDesktop.App 9.x
  if DotNetListContainsDesktop9Runtime() then
  begin
    Result := True;
    exit;
  end;

  // Alternative: check registry
  Result :=
    HasDotNet9DesktopRuntimeInRegistry('SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App') or
    HasDotNet9DesktopRuntimeInRegistry('SOFTWARE\dotnet\Setup\InstalledVersions\x86\sharedfx\Microsoft.WindowsDesktop.App');
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
