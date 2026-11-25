#define DotNet10Version "10.0"
#define DotNet10InstallerUrl "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/10.0.15/windowsdesktop-runtime-10.0.15-win-x64.exe"
#define DotNet10InstallerUrlX86 "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/10.0.15/windowsdesktop-runtime-10.0.15-win-x86.exe"
#define DotNet10InstallerUrlArm64 "https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/10.0.15/windowsdesktop-runtime-10.0.15-win-arm64.exe"

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

function GetDotNet10Url(Param: String): String;
var
  Architecture: String;
begin
  Architecture := GetArchitectureString;
  if Architecture = 'x86' then
    Result := '{#DotNet10InstallerUrlX86}'
  else if Architecture = 'arm64' then
    Result := '{#DotNet10InstallerUrlArm64}'
  else
    Result := '{#DotNet10InstallerUrl}';
end;

function GetDotNet10Filename: String;
begin
  Result := 'dotnet10-' + GetArchitectureString + '.exe';
end;

function GetDotNet10Hash: String;
var
  Architecture: String;
begin
  Architecture := GetArchitectureString;
  if Architecture = 'x86' then
    Result := '0F55FC79174EA8DAA11C339FBA629B286BA6CADBB920A77DE366B408B15315A5'
  else if Architecture = 'arm64' then
    Result := 'AB46299E102BDE89C8B8C19F5E53EB47064513F33BFD2D5FC84BF623FBF00B50'
  else
    Result := 'E61B030C1673EDFA3C0844502EFF83198C46AF9E7DF70DFC6E2EFC5B11EA513A';
end;

function IsSupportedDesktopRuntimeVersion(Version: String): Boolean;
var
  Prefix: String;
  PatchText: String;
  PatchVersion: Integer;
begin
  Prefix := '{#DotNet10Version}.';
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

function IsDotNet10Installed: Boolean;
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
