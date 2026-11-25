#ifdef Arch
  #define TestArchitecture Arch
#else
  #define TestArchitecture "combined"
#endif

[Setup]
AppName=WinTab Installer Runtime Tests
AppVersion=1.0
DefaultDirName={tmp}\WinTabInstallerRuntimeTests
PrivilegesRequired=lowest
CreateAppDir=no
Uninstallable=no
OutputDir={#TestOutputDir}
OutputBaseFilename=RuntimeTests_{#TestArchitecture}

#include "..\installers\Runtime.iss"
#include "..\installers\RecycleBin.iss"
#include "..\installers\Startup.iss"

[Code]
procedure Check(Condition: Boolean; Message: String);
begin
  if not Condition then
    RaiseException(Message);
end;

procedure CheckRecycleCleanup(Root: Integer);
var
  Page, Shell, Verb, Command, Value: String;
begin
  { The cleanup receives a test-owned page, never the real Recycle Bin registration. }
  Page := 'Software\WinTab.Tests\InstallerRecycle-' + '{#TestArchitecture}';
  Check(not RegKeyExists(Root, Page), 'The isolated test key must not exist.');
  Shell := Page + '\shell';
  Verb := Shell + '\open';
  Command := Verb + '\command';
  try
    RegWriteStringValue(Root, Shell, '', 'open');
    RegWriteStringValue(Root, Verb, 'WinTab.Owner', 'WinTab.RecycleBinOpen.v1');
    RegWriteStringValue(Root, Verb, 'WinTab.Command', 'test-command');
    RegWriteDWordValue(Root, Verb, 'WinTab.HadDefault', 1);
    RegWriteStringValue(Root, Verb, 'WinTab.Default', 'original-verb');
    RegWriteDWordValue(Root, Verb, 'WinTab.DefaultKind', 1);
    RegWriteDWordValue(Root, Verb, 'WinTab.HadPage', 1);
    RegWriteDWordValue(Root, Verb, 'WinTab.HadShell', 1);
    RegWriteStringValue(Root, Command, '', 'changed-by-another-app');
    RegWriteStringValue(Root, Command, 'DelegateExecute', '');
    RestoreRecycleBinOpenInView(Root, Page);
    Check(RegQueryStringValue(Root, Command, '', Value) and (Value = 'changed-by-another-app'),
      'Uninstall must preserve a changed command.');
    RegWriteStringValue(Root, Command, '', 'test-command');
    RestoreRecycleBinOpenInView(Root, Page);
    Check(RegQueryStringValue(Root, Shell, '', Value) and (Value = 'original-verb'),
      'Uninstall must restore the original default.');
    Check(not RegKeyExists(Root, Verb), 'Uninstall must remove the owned handler.');
  finally
    RegDeleteKeyIncludingSubkeys(Root, Page);
  end;
end;

function InitializeSetup: Boolean;
var
  ExpectedArchitecture: String;
  ExpectedHash: String;
  ResultPath: String;
  Report: String;
  AppPath: String;
begin
  Result := False;
  ResultPath := ExpandConstant('{param:ResultFile}');
  try
    AppPath := ResultPath + '.old-app.exe';
    Check(not FileExists(AppPath), 'The startup probe must own its temporary app path');
    DetectUpdateInstall(AppPath);
    if ExpandConstant('{param:UPDATE|0}') = '1' then
      Check(GetAppLaunchParameters('') = '--background', 'An in-app update must resume in the tray')
    else
      Check(GetAppLaunchParameters('') = '', 'A first install must still open the main window');
    Check(SaveStringToFile(AppPath, 'old version', False), 'Could not prepare the manual upgrade probe');
    DetectUpdateInstall(AppPath);
    Check(GetAppLaunchParameters('') = '--background', 'A manual upgrade must resume in the tray');
    CheckRecycleCleanup(HKCU32);
    if IsWin64 then CheckRecycleCleanup(HKCU64);
#ifdef Arch
    ExpectedArchitecture := '{#Arch}';
#else
    if ProcessorArchitecture = paARM64 then
      ExpectedArchitecture := 'arm64'
    else if ProcessorArchitecture = paX64 then
      ExpectedArchitecture := 'x64'
    else
      ExpectedArchitecture := 'x86';
#endif
    if ExpectedArchitecture = 'x86' then
      ExpectedHash := '0F55FC79174EA8DAA11C339FBA629B286BA6CADBB920A77DE366B408B15315A5'
    else if ExpectedArchitecture = 'arm64' then
      ExpectedHash := 'AB46299E102BDE89C8B8C19F5E53EB47064513F33BFD2D5FC84BF623FBF00B50'
    else
      ExpectedHash := 'E61B030C1673EDFA3C0844502EFF83198C46AF9E7DF70DFC6E2EFC5B11EA513A';

    Check(GetArchitectureString = ExpectedArchitecture, 'Runtime architecture must match the selected installer.');
    Check(GetDotNet10Url('') = 'https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/10.0.15/windowsdesktop-runtime-10.0.15-win-' + ExpectedArchitecture + '.exe', 'Runtime URL must select the official matching installer.');
    Check(GetDotNet10Filename = 'dotnet10-' + ExpectedArchitecture + '.exe', 'Downloaded files must retain their architecture.');
    Check(CompareText(GetDotNet10Hash, ExpectedHash) = 0, 'The download must require the verified runtime checksum.');
    Check(IsSupportedDesktopRuntimeVersion('10.0.0'), 'An installed .NET 10 desktop runtime should be accepted.');
    Check(IsSupportedDesktopRuntimeVersion('10.0.15'), 'Newer .NET 10 patches should be accepted.');
    Check(not IsSupportedDesktopRuntimeVersion('10.0.15-preview.1'), 'Preview runtimes must not satisfy a stable requirement.');
    Check(not IsSupportedDesktopRuntimeVersion('10.0.bad'), 'Malformed runtime versions must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('10.0.-1'), 'Negative runtime patches must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('10.0.15.1'), 'An extra version component must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('9.0.23'), 'An older major runtime must not satisfy .NET 10.');
    Check(not IsSupportedDesktopRuntimeVersion('11.0.0'), 'A newer major runtime must not satisfy .NET 10.');
    Check(IsSuccessfulRuntimeExitCode(0), 'A completed installation should succeed.');
    Check(IsSuccessfulRuntimeExitCode(3010), 'A successful installation requiring restart should succeed.');
    Check(not IsSuccessfulRuntimeExitCode(1603), 'A failed runtime installation must not report success.');
    Report := 'PASS';
  except
    Report := 'FAIL: ' + GetExceptionMessage;
  end;
  if not SaveStringToFile(ResultPath, Report, False) then
    RaiseException('Could not write the installer self-test result.');
end;
