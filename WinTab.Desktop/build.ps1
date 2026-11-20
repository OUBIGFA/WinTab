#requires -Version 5.1
<# Builds the native MyGo settings UI renderer after the resident application has been published. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('x64', 'arm64')][string[]]$Arch,
    [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][string]$PublishRoot,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$desktopRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent $desktopRoot
$go = (Get-Command go.exe -ErrorAction Stop).Source
$numericVersion = $Version.TrimStart('v').Split('-')[0]
# Match MyGo's production build: settings need neither development inspection nor
# automatic native-library downloads. Test the same code that the installer will carry.
$productionTags = 'mygo_nowebview,mygo_noinspector,mygo_nodownload'

$sourceLogo = Join-Path $repoRoot 'WinTab/wintab-logo.png'
$desktopLogo = Join-Path $desktopRoot 'assets/wintab-logo.png'
if (Test-Path -LiteralPath $sourceLogo) {
    $assetsDir = Split-Path -Parent $desktopLogo
    if (-not (Test-Path -LiteralPath $assetsDir)) {
        New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $sourceLogo -Destination $desktopLogo -Force
}

$previousGoOS = $env:GOOS
$previousGoArch = $env:GOARCH
$previousCgo = $env:CGO_ENABLED
Push-Location $desktopRoot
try {
    Write-Host '==> Building native MyGo settings UI' -ForegroundColor Cyan
    & $go mod download
    if ($LASTEXITCODE -ne 0) { throw "Desktop dependency restore failed (exit $LASTEXITCODE)" }
    & $go vet -tags $productionTags ./...
    if ($LASTEXITCODE -ne 0) { throw "Desktop static check failed (exit $LASTEXITCODE)" }
    if (-not $SkipTests) {
        & $go test -tags $productionTags -count=1 ./...
        if ($LASTEXITCODE -ne 0) { throw "Desktop UI tests failed (exit $LASTEXITCODE)" }
    }
    & $go tool go-winres make --in windows-resources.json --arch amd64,arm64 --product-version $numericVersion --file-version $numericVersion
    if ($LASTEXITCODE -ne 0) { throw "Desktop Windows resource generation failed (exit $LASTEXITCODE)" }
    $env:GOOS = 'windows'
    $env:CGO_ENABLED = '0'
    foreach ($architecture in $Arch) {
        $env:GOARCH = if ($architecture -eq 'x64') { 'amd64' } else { 'arm64' }
        $destination = Join-Path (Join-Path $PublishRoot $architecture) 'WinTab.UI.exe'
        if (-not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $destination) 'WinTab.exe'))) {
            throw "Publish the resident application before its renderer: $destination"
        }
        Write-Host "    - settings UI: $architecture"
        & $go build -tags $productionTags -trimpath -buildvcs=false -ldflags "-H windowsgui -s -w -X main.version=$numericVersion" -o $destination .
        if ($LASTEXITCODE -ne 0) { throw "Desktop build failed for $architecture (exit $LASTEXITCODE)" }
    }
}
finally {
    $env:GOOS = $previousGoOS
    $env:GOARCH = $previousGoArch
    $env:CGO_ENABLED = $previousCgo
    Pop-Location
}
