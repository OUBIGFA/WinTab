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

[Code]
procedure Check(Condition: Boolean; Message: String);
begin
  if not Condition then
    RaiseException(Message);
end;

function InitializeSetup: Boolean;
var
  ExpectedArchitecture: String;
  ExpectedHash: String;
  ResultPath: String;
  Report: String;
begin
  Result := False;
  ResultPath := ExpandConstant('{param:ResultFile}');
  try
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
      ExpectedHash := '31AB52DCF6440AE1BAC553B21AD97FD88F2F2354DF71B19946F90FDE426595BA'
    else if ExpectedArchitecture = 'arm64' then
      ExpectedHash := '03CAB3CB3006E0C089CC716EF27FBEC7E8EA8D7C873D99F4DD397BEE64CC7412'
    else
      ExpectedHash := '4BEE05AA0637468A19CD82490858FC69E93FCE8D22C0AEB272A76B71F0DC93E9';

    Check(GetArchitectureString = ExpectedArchitecture, 'Runtime architecture must match the selected installer.');
    Check(GetDotNet9Url('') = 'https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/9.0.19/windowsdesktop-runtime-9.0.19-win-' + ExpectedArchitecture + '.exe', 'Runtime URL must select the official matching installer.');
    Check(GetDotNet9Filename = 'dotnet9-' + ExpectedArchitecture + '.exe', 'Downloaded files must retain their architecture.');
    Check(CompareText(GetDotNet9Hash, ExpectedHash) = 0, 'The download must require the verified runtime checksum.');
    Check(IsSupportedDesktopRuntimeVersion('9.0.3'), 'An installed .NET 9 desktop runtime should be accepted.');
    Check(IsSupportedDesktopRuntimeVersion('9.0.19'), 'Newer .NET 9 patches should be accepted.');
    Check(not IsSupportedDesktopRuntimeVersion('9.0.19-preview.1'), 'Preview runtimes must not satisfy a stable requirement.');
    Check(not IsSupportedDesktopRuntimeVersion('9.0.bad'), 'Malformed runtime versions must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('9.0.-1'), 'Negative runtime patches must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('9.0.19.1'), 'An extra version component must be rejected.');
    Check(not IsSupportedDesktopRuntimeVersion('8.0.19'), 'An older major runtime must not satisfy .NET 9.');
    Check(not IsSupportedDesktopRuntimeVersion('10.0.0'), 'A different major runtime must not satisfy .NET 9.');
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
