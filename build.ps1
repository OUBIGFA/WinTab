#requires -Version 5.1
<#
.SYNOPSIS
    Publish WinTab for x64/x86/arm64 and compile per-arch Inno Setup installers.

.PARAMETER Version
    Version string used in the installer filename, e.g. "v1.0.0".
    Defaults to the <Version> in WinTab/WinTab.csproj prefixed with "v".

.PARAMETER Arch
    Architectures to build. Defaults to x64, x86, arm64.

.PARAMETER SkipPublish
    Skip the dotnet publish step (use existing publish output).

.PARAMETER SkipInstaller
    Skip the Inno Setup compile step.

.PARAMETER Combined
    Also build the combined (auto-detect) installer in addition to per-arch installers.

.PARAMETER IsccPath
    Path to ISCC.exe. Auto-detected if not provided.

.EXAMPLE
    .\build.ps1
    .\build.ps1 -Version v1.1.0
    .\build.ps1 -Arch x64
#>
[CmdletBinding()]
param(
    [string]$Version,
    [ValidateSet('x64', 'x86', 'arm64')]
    [string[]]$Arch = @('x64', 'x86', 'arm64'),
    [switch]$SkipPublish,
    [switch]$SkipInstaller,
    [switch]$Combined,
    [string]$IsccPath,
    [string]$MSBuildPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectPath = Join-Path $RepoRoot 'WinTab\WinTab.csproj'
$PublishRoot = Join-Path $RepoRoot 'publish\net9.0-windows'
$DistDir = Join-Path $RepoRoot 'dist'
$InstallerScript = Join-Path $RepoRoot 'installers\installer.iss'

if (-not $Version) {
    [xml]$csproj = Get-Content -Path $ProjectPath
    $rawVersion = $csproj.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    if (-not $rawVersion) { throw "Could not read <Version> from $ProjectPath" }
    $Version = "v$rawVersion"
}

Write-Host "==> WinTab build" -ForegroundColor Cyan
Write-Host "    Version    : $Version"
Write-Host "    Arch       : $($Arch -join ', ')"
Write-Host "    PublishRoot: $PublishRoot"
Write-Host "    DistDir    : $DistDir"

function Resolve-Iscc {
    param([string]$Hint)
    if ($Hint) {
        if (Test-Path $Hint) { return $Hint }
        throw "ISCC.exe not found at: $Hint"
    }
    $cmd = Get-Command -Name 'ISCC.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if ($candidates) { return $candidates }
    throw "Inno Setup 6 not found. Install it or pass -IsccPath."
}

function Resolve-MSBuild {
    param([string]$Hint)
    if ($Hint) {
        if (Test-Path $Hint) { return $Hint }
        throw "MSBuild.exe not found at: $Hint"
    }
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        # Try standard channel first, then prerelease/insiders.
        $msb = & $vswhere -latest -products * -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        if (-not $msb) {
            $msb = & $vswhere -latest -prerelease -products * -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        }
        if ($msb -and (Test-Path $msb)) { return $msb }
    }
    $cmd = Get-Command -Name 'MSBuild.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    throw "MSBuild.exe not found. Install Visual Studio with the '.NET desktop development' workload or pass -MSBuildPath."
}

if (-not $SkipPublish) {
    $msbuild = Resolve-MSBuild -Hint $MSBuildPath
    Write-Host "`n==> Publishing $($Arch.Count) arch(es) via $msbuild" -ForegroundColor Cyan
    foreach ($a in $Arch) {
        $out = Join-Path $PublishRoot $a
        Write-Host "    - win-$a -> $out"
        if (Test-Path $out) { Remove-Item -Recurse -Force $out }
        & $msbuild $ProjectPath `
            /restore `
            /t:Publish `
            /p:Configuration=Release `
            /p:TargetFramework=net9.0-windows `
            /p:RuntimeIdentifier="win-$a" `
            /p:SelfContained=false `
            /p:PublishDir="$out" `
            /p:DebugType=none `
            /nologo `
            /v:minimal
        if ($LASTEXITCODE -ne 0) { throw "MSBuild publish failed for win-$a (exit $LASTEXITCODE)" }
    }
}
else {
    Write-Host "`n==> SkipPublish: using existing publish output" -ForegroundColor Yellow
}

if ($SkipInstaller) {
    Write-Host "`n==> SkipInstaller: stopping after publish" -ForegroundColor Yellow
    return
}

$iscc = Resolve-Iscc -Hint $IsccPath
Write-Host "`n==> Compiling installers with $iscc" -ForegroundColor Cyan

if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }

foreach ($a in $Arch) {
    $publishArchDir = Join-Path $PublishRoot $a
    if (-not (Test-Path (Join-Path $publishArchDir 'WinTab.exe'))) {
        throw "Missing publish output for $a at $publishArchDir. Run without -SkipPublish first."
    }
    Write-Host "    - $a installer"
    & $iscc `
        "/DMyAppVersion=$Version" `
        "/DPublishRoot=..\publish\net9.0-windows" `
        "/DOutputDir=..\dist" `
        "/DArch=$a" `
        $InstallerScript | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "ISCC failed for $a (exit $LASTEXITCODE)" }
}

if ($Combined) {
    Write-Host "    - combined installer"
    & $iscc `
        "/DMyAppVersion=$Version" `
        "/DPublishRoot=..\publish\net9.0-windows" `
        "/DOutputDir=..\dist" `
        $InstallerScript | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "ISCC failed for combined installer (exit $LASTEXITCODE)" }
}

Write-Host "`n==> Done. Artifacts:" -ForegroundColor Green
Get-ChildItem -Path $DistDir -Filter "WinTab_${Version}*_Setup.exe" |
    Sort-Object Name |
    ForEach-Object {
        $sizeMb = '{0:N2}' -f ($_.Length / 1MB)
        Write-Host ("    {0,-40} {1,8} MB" -f $_.Name, $sizeMb)
    }
