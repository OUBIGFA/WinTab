#define DotNet9Version "9.0"
#define DotNet9InstallerUrl "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/9.0.19/windowsdesktop-runtime-9.0.19-win-x64.exe"
#define DotNet9InstallerUrlX86 "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/9.0.19/windowsdesktop-runtime-9.0.19-win-x86.exe"
#define DotNet9InstallerUrlArm64 "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/9.0.19/windowsdesktop-runtime-9.0.19-win-arm64.exe"

[Code]
function IsX86: Boolean;
begin
  Result := not IsWin64;
end;

function IsX64: Boolean;
begin
  Result := IsWin64 and (ProcessorArchitecture = paX64);
end;

function IsArm64: Boolean;
begin
  Result := IsWin64 and (ProcessorArchitecture = paARM64);
end;

function GetArchitectureString: String;
begin
#ifdef Arch
  Result := '{#Arch}';
#else
  if IsX86 then
    Result := 'x86'
  else if IsArm64 then
    Result := 'arm64'
  else
    Result := 'x64';
#endif
end;

function GetDotNet9Url(Param: String): String;
var
  Architecture: String;
begin
  Architecture := GetArchitectureString;
  if Architecture = 'x86' then
    Result := '{#DotNet9InstallerUrlX86}'
  else if Architecture = 'arm64' then
    Result := '{#DotNet9InstallerUrlArm64}'
  else
    Result := '{#DotNet9InstallerUrl}';
end;

function GetDotNet9Filename: String;
begin
  Result := 'dotnet9-' + GetArchitectureString + '.exe';
end;

function GetDotNet9Hash: String;
var
  Architecture: String;
begin
  Architecture := GetArchitectureString;
  if Architecture = 'x86' then
    Result := '31AB52DCF6440AE1BAC553B21AD97FD88F2F2354DF71B19946F90FDE426595BA'
  else if Architecture = 'arm64' then
    Result := '03CAB3CB3006E0C089CC716EF27FBEC7E8EA8D7C873D99F4DD397BEE64CC7412'
  else
    Result := '4BEE05AA0637468A19CD82490858FC69E93FCE8D22C0AEB272A76B71F0DC93E9';
end;

function IsSupportedDesktopRuntimeVersion(Version: String): Boolean;
var
  Prefix: String;
  PatchText: String;
  PatchVersion: Integer;
begin
  Prefix := '{#DotNet9Version}.';
  Result := Pos(Prefix, Version) = 1;
  if not Result then
    Exit;

  PatchText := Copy(Version, Length(Prefix) + 1, Length(Version));
  PatchVersion := StrToIntDef(PatchText, -1);
  Result := (PatchVersion >= 0) and (IntToStr(PatchVersion) = PatchText);
end;

function IsSuccessfulRuntimeExitCode(ExitCode: Integer): Boolean;
begin
  Result := (ExitCode = 0) or (ExitCode = 3010);
end;

function IsDotNet9Installed: Boolean;
var
  Versions: TArrayOfString;
  VersionIndex: Integer;
  Installed: Cardinal;
  RegistryPath: String;
begin
  Result := False;
  RegistryPath := 'SOFTWARE\dotnet\Setup\InstalledVersions\' + GetArchitectureString +
    '\sharedfx\Microsoft.WindowsDesktop.App';
  if not RegGetValueNames(HKEY_LOCAL_MACHINE_32, RegistryPath, Versions) then
    Exit;

  for VersionIndex := 0 to GetArrayLength(Versions) - 1 do
    if IsSupportedDesktopRuntimeVersion(Versions[VersionIndex]) and
       RegQueryDWordValue(HKEY_LOCAL_MACHINE_32, RegistryPath, Versions[VersionIndex], Installed) then
      if Installed = 1 then
      begin
        Result := True;
        Exit;
      end;
end;
