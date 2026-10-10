#requires -Version 5.1
<# Builds the bundled settings renderer after the resident application has been published. #>
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
$npm = (Get-Command npm.cmd -ErrorAction Stop).Source
$go = (Get-Command go.exe -ErrorAction Stop).Source
$numericVersion = $Version.TrimStart('v').Split('-')[0]

Push-Location (Join-Path $desktopRoot 'frontend')
try {
    Write-Host '==> Building settings frontend' -ForegroundColor Cyan
    & $npm ci --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { throw "Frontend dependency restore failed (exit $LASTEXITCODE)" }
    & $npm run build
    if ($LASTEXITCODE -ne 0) { throw "Frontend build failed (exit $LASTEXITCODE)" }
    if (-not $SkipTests) {
        & $npm exec -- playwright install chromium
        if ($LASTEXITCODE -ne 0) { throw "Test browser setup failed (exit $LASTEXITCODE)" }
        $testOutput = Join-Path $repoRoot ('_temp/frontend-tests-' + [Guid]::NewGuid().ToString('N'))
        & $npm test -- "--output=$testOutput"
        if ($LASTEXITCODE -ne 0) { throw "Frontend tests failed (exit $LASTEXITCODE), results: $testOutput" }
        # Only recycle this run's test output; failed runs retain their diagnostics for inspection.
        if (Test-Path -LiteralPath $testOutput) {
            $temporaryRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot '_temp'))
            $resolvedOutput = [IO.Path]::GetFullPath($testOutput)
            if (-not $resolvedOutput.StartsWith($temporaryRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Test output escaped the temporary directory: $resolvedOutput"
            }
            try {
                Add-Type -AssemblyName Microsoft.VisualBasic
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($resolvedOutput,
                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
            }
            catch {
                if (Test-Path -LiteralPath $resolvedOutput) {
                    Write-Warning ("Test output could not be recycled and remains at {0}: {1}" -f $resolvedOutput, $_)
                }
            }
        }
    }
}
finally { Pop-Location }

$previousGoOS = $env:GOOS
$previousGoArch = $env:GOARCH
$previousCgo = $env:CGO_ENABLED
Push-Location $desktopRoot
try {
    & $go mod download
    if ($LASTEXITCODE -ne 0) { throw "Desktop dependency restore failed (exit $LASTEXITCODE)" }
    if (-not $SkipTests) {
        & $go test ./...
        if ($LASTEXITCODE -ne 0) { throw "Desktop bridge tests failed (exit $LASTEXITCODE)" }
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
        Write-Host "    - settings renderer: $architecture"
        & $go build -trimpath -buildvcs=false -ldflags "-H windowsgui -s -w -X main.version=$numericVersion" -o $destination .
        if ($LASTEXITCODE -ne 0) { throw "Desktop build failed for $architecture (exit $LASTEXITCODE)" }
    }
}
finally {
    $env:GOOS = $previousGoOS
    $env:GOARCH = $previousGoArch
    $env:CGO_ENABLED = $previousCgo
    Pop-Location
}
