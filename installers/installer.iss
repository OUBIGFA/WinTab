; The version will be set by the build script
#ifndef MyAppVersion
  #define MyAppVersion "v1.0.0"
#endif

; Extract numeric version for VersionInfoVersion (remove 'v' and any suffix after a dash)
#define MyAppVersionWithoutV StringChange(MyAppVersion, "v", "")
#define DashPos Pos("-", MyAppVersionWithoutV)
#define MyAppNumericVersion (DashPos > 0) ? Copy(MyAppVersionWithoutV, 1, DashPos - 1) : MyAppVersionWithoutV

#define MyAppPublisher "OUBIGFA"
#define MyAppName "WinTab"
#define MyAppExeName MyAppName + ".exe"
#define MyAppRelativePath MyAppName + "\" + MyAppExeName
#define MyAppURL "https://github.com/OUBIGFA/WinTab"

#ifndef PublishRoot
  #define PublishRoot "..\publish\net9.0-windows"
#endif

#ifndef OutputDir
  #define OutputDir "..\dist"
#endif

; When Arch is defined ("x64" | "x86" | "arm64"), build a per-arch installer
; that only includes that architecture's files. When undefined, the installer
; bundles all three and chooses at install time.
#ifdef Arch
  #if (Arch != "x64") && (Arch != "x86") && (Arch != "arm64")
    #error Arch must be one of: x64, x86, arm64
  #endif
  #define OutputSuffix "_" + Arch
  #if Arch == "x64"
    #define ArchInstallIn64Bit "x64"
  #elif Arch == "arm64"
    #define ArchInstallIn64Bit "arm64"
  #else
    #define ArchInstallIn64Bit ""
  #endif
#else
  #define OutputSuffix ""
  #define ArchInstallIn64Bit "x64 arm64"
#endif

[Setup]
AppId={{E1F2A3C4-F5D4-3B2E-1D5A-8F7B2F8D3E7A}
AppName={#MyAppName}
AppMutex=__WinTabHook__Mutex
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={userpf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
PrivilegesRequired=lowest
OutputDir={#OutputDir}
OutputBaseFilename={#MyAppName}_{#MyAppVersion}{#OutputSuffix}_Setup
SetupIconFile=..\{#MyAppName}\Icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
LanguageDetectionMethod=uilanguage
UsePreviousLanguage=no
ShowLanguageDialog=no
ArchitecturesInstallIn64BitMode={#ArchInstallIn64Bit}
UninstallDisplayIcon={app}\{#MyAppRelativePath}
UninstallDisplayName={#MyAppName}
CloseApplications=yes
RestartApplications=no
VersionInfoVersion={#MyAppNumericVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoTextVersion={#MyAppVersion}
VersionInfoCopyright={#MyAppPublisher}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppNumericVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: ".\ChineseSimplified.isl"

[CustomMessages]
english.StartWithWindows=Start with Windows
chinesesimplified.StartWithWindows=开机启动
english.WindowsIntegration=Windows Integration
chinesesimplified.WindowsIntegration=Windows 集成
english.RuntimeNoInternet=No internet connection, WinTab requires .NET 9 Desktop Runtime
chinesesimplified.RuntimeNoInternet=无法连接网络，WinTab 需要 .NET 9 桌面运行时
english.RuntimeDownloadFailed=Could not download .NET 9 Desktop Runtime: %1
chinesesimplified.RuntimeDownloadFailed=无法下载 .NET 9 桌面运行时：%1
english.RuntimeDownloading=Downloading .NET 9 Desktop Runtime (%1)
chinesesimplified.RuntimeDownloading=正在下载 .NET 9 桌面运行时（%1）
english.RuntimeDownloadWait=Please wait while setup downloads the required files
chinesesimplified.RuntimeDownloadWait=请等待安装程序下载所需文件
english.RuntimeInstalling=Installing .NET 9 Desktop Runtime
chinesesimplified.RuntimeInstalling=正在安装 .NET 9 桌面运行时
english.RuntimeInstallWait=This may take a few minutes
chinesesimplified.RuntimeInstallWait=这可能需要几分钟
english.RuntimeNotDetected=Setup finished but .NET 9 Desktop Runtime was not detected
chinesesimplified.RuntimeNotDetected=安装已结束，但未检测到 .NET 9 桌面运行时
english.RuntimeInstallFailed=Could not install .NET 9 Desktop Runtime
chinesesimplified.RuntimeInstallFailed=无法安装 .NET 9 桌面运行时
english.RuntimeLaunchFailed=Could not start the .NET 9 Desktop Runtime installer
chinesesimplified.RuntimeLaunchFailed=无法启动 .NET 9 桌面运行时安装程序
english.RuntimeInstallError=Error installing .NET 9 Desktop Runtime: %1
chinesesimplified.RuntimeInstallError=安装 .NET 9 桌面运行时出错：%1
english.RuntimeRequired=WinTab requires .NET 9 Desktop Runtime, setup will download it if needed
chinesesimplified.RuntimeRequired=WinTab 需要 .NET 9 桌面运行时，安装程序将在缺少时下载
english.WrongArchitecture=This installer is for %1 Windows, download the installer matching your device
chinesesimplified.WrongArchitecture=此安装器适用于 %1 Windows，请下载与设备架构匹配的安装器

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon"; Description: "{cm:StartWithWindows}"; GroupDescription: "{cm:WindowsIntegration}"

[Files]
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\THIRD-PARTY-NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion
#ifdef Arch
Source: "{#PublishRoot}\{#Arch}\*"; DestDir: "{app}\{#MyAppName}"; Flags: ignoreversion recursesubdirs createallsubdirs
#else
Source: "{#PublishRoot}\x86\*"; DestDir: "{app}\{#MyAppName}"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: IsX86
Source: "{#PublishRoot}\x64\*"; DestDir: "{app}\{#MyAppName}"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: IsX64
Source: "{#PublishRoot}\arm64\*"; DestDir: "{app}\{#MyAppName}"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: IsArm64
#endif

[Dirs]
Name: "{app}"; Flags: uninsalwaysuninstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppRelativePath}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppRelativePath}"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppRelativePath}"" --background"; Flags: uninsdeletevalue; Tasks: startupicon
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"; ValueType: binary; ValueName: "{#MyAppName}"; ValueData: 02 00 00 00 00 00 00 00 00 00 00 00; Flags: uninsdeletevalue; Tasks: startupicon

[UninstallDelete]
Type: dirifempty; Name: "{app}"

[Run]
Filename: "{app}\{#MyAppRelativePath}"; Parameters: "{code:GetAppLaunchParameters}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  DotNet9Detected: Boolean;
  DotNet9RestartRequired: Boolean;
  DownloadPage: TDownloadWizardPage;

#include "Runtime.iss"
#include "RecycleBin.iss"
#include "Startup.iss"

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    DetectUpdateInstall(ExpandConstant('{app}\{#MyAppRelativePath}'));
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RestoreRecycleBinOpenInView(HKCU32, 'Software\Classes\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}');
    if IsWin64 then
      RestoreRecycleBinOpenInView(HKCU64, 'Software\Classes\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}');
  end;
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Microsoft\Windows\CurrentVersion\Run', '{#MyAppName}');
    RegDeleteValue(HKEY_CURRENT_USER, 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run', '{#MyAppName}');
  end;
end;

function OnDownloadProgress(const Url, FileName: String; const Progress, ProgressMax: Int64): Boolean;
begin
  if Progress = ProgressMax then
    Log(Format('Successfully downloaded file to %s', [FileName]));
  Result := True;
end;

procedure HandleDownloadError(const ErrorMessage: String);
begin
  Log('Download error: ' + ErrorMessage);

  if Pos('12007', ErrorMessage) > 0 then
    SuppressibleMsgBox(CustomMessage('RuntimeNoInternet'),
                       mbInformation, MB_OK, MB_OK)
  else if Pos('aborted', ErrorMessage) > 0 then
    Log('Download was aborted by user.')
  else if not DownloadPage.AbortedByUser then
    SuppressibleMsgBox(FmtMessage(CustomMessage('RuntimeDownloadFailed'), [ErrorMessage]),
                       mbInformation, MB_OK, MB_OK);
end;

function DownloadAndInstallDotNet9: Boolean;
var
  DotNetInstallerPath: String;
  ResultCode: Integer;
  ErrorMessage: String;
  ArchString: String;
begin
  Result := False;
  ArchString := GetArchitectureString;
  DownloadPage.Clear;
  DownloadPage.Add(GetDotNet9Url(''), GetDotNet9Filename(), GetDotNet9Hash);

  try
    DownloadPage.SetText(FmtMessage(CustomMessage('RuntimeDownloading'), [ArchString]),
                         CustomMessage('RuntimeDownloadWait'));
    DownloadPage.Show;

    try
      DownloadPage.Download;
      DotNetInstallerPath := ExpandConstant('{tmp}\') + GetDotNet9Filename();

      if FileExists(DotNetInstallerPath) then
        Result := True
      else
      begin
        Log('Downloaded file is missing');
        Result := False;
      end;
    except
      ErrorMessage := GetExceptionMessage;
      Log('Download exception: ' + ErrorMessage);
      HandleDownloadError(ErrorMessage);
      Result := False;
    end;

    if Result then
    begin
      DownloadPage.SetText(CustomMessage('RuntimeInstalling'), CustomMessage('RuntimeInstallWait'));
      DownloadPage.SetProgress(0, 100);

      try
        if Exec(DotNetInstallerPath, '/install /passive /norestart', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
        begin
          if IsSuccessfulRuntimeExitCode(ResultCode) then
          begin
            DotNet9RestartRequired := ResultCode = 3010;
            Log('Successfully installed .NET 9 Desktop Runtime');
            DownloadPage.SetProgress(100, 100);
            Result := True;

            DotNet9Detected := IsDotNet9Installed;
            if not DotNet9Detected then
            begin
              Log('Installation completed but .NET 9 Desktop Runtime is still not detected');
              SuppressibleMsgBox(CustomMessage('RuntimeNotDetected'),
                                mbInformation, MB_OK, MB_OK);
              Result := False;
            end;
          end
          else
          begin
            Log(Format('Failed to install .NET 9 Desktop Runtime. Exit code: %d', [ResultCode]));
            SuppressibleMsgBox(CustomMessage('RuntimeInstallFailed'),
                              mbInformation, MB_OK, MB_OK);
            Result := False;
          end;
        end
        else
        begin
          Log('Failed to execute .NET 9 Desktop Runtime installer');
          SuppressibleMsgBox(CustomMessage('RuntimeLaunchFailed'),
                            mbInformation, MB_OK, MB_OK);
          Result := False;
        end;
      except
        ErrorMessage := GetExceptionMessage;
        Log('Installation exception: ' + ErrorMessage);
        SuppressibleMsgBox(FmtMessage(CustomMessage('RuntimeInstallError'), [ErrorMessage]),
                          mbInformation, MB_OK, MB_OK);
        Result := False;
      end;
    end;
  finally
    DownloadPage.Hide;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  ErrorMessage: String;
begin
  Result := True;

  if (CurPageID = wpReady) and (not DotNet9Detected) then
  begin
    try
      Result := DownloadAndInstallDotNet9;
    except
      ErrorMessage := GetExceptionMessage;
      Log('Unexpected error in .NET 9 installation process: ' + ErrorMessage);
      Result := False;
    end;
  end;
end;

function NeedRestart: Boolean;
begin
  Result := DotNet9RestartRequired;
end;

procedure InitializeWizard;
begin
  DotNet9Detected := IsDotNet9Installed;
  DownloadPage := CreateDownloadPage(SetupMessage(msgWizardPreparing),
    SetupMessage(msgPreparingDesc), @OnDownloadProgress);

  if not DotNet9Detected and not WizardSilent then
  begin
    MsgBox(CustomMessage('RuntimeRequired'),
           mbInformation, MB_OK);
  end;
end;

#ifdef Arch
function InitializeSetup(): Boolean;
begin
  Result := True;
  #if Arch == "x64"
  if not IsX64 then
  begin
    SuppressibleMsgBox(FmtMessage(CustomMessage('WrongArchitecture'), ['x64']),
                       mbError, MB_OK, IDOK);
    Result := False;
  end;
  #elif Arch == "arm64"
  if not IsArm64 then
  begin
    SuppressibleMsgBox(FmtMessage(CustomMessage('WrongArchitecture'), ['ARM64']),
                       mbError, MB_OK, IDOK);
    Result := False;
  end;
  #elif Arch == "x86"
  if IsArm64 then
  begin
    SuppressibleMsgBox(FmtMessage(CustomMessage('WrongArchitecture'), ['x86']),
                       mbError, MB_OK, IDOK);
    Result := False;
  end;
  #endif
end;
#endif
