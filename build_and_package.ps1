$ProjectDir = "WinTab"
$PublishDir = "publish"
$MSBuildPath = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"

if (Test-Path $PublishDir) {
    Remove-Item -Path $PublishDir -Recurse -Force
}

New-Item -ItemType Directory -Path $PublishDir -Force

Write-Host "Restoring NuGet packages..." -ForegroundColor Cyan
dotnet restore WinTab\WinTab.csproj

Write-Host "Building WinTab (net481) with MSBuild..." -ForegroundColor Cyan
& $MSBuildPath WinTab\WinTab.csproj /p:Configuration=Release /p:OutputPath="..\$PublishDir" /p:TargetFramework=net481

if ($?) {
    Write-Host "Build successful! Files are in the '$PublishDir' folder." -ForegroundColor Green
    Write-Host "Main executable: $PublishDir\WinTab.exe" -ForegroundColor Green
    
    $ISCCPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path $ISCCPath) {
        Write-Host "Compiling installer with Inno Setup..." -ForegroundColor Cyan
        & $ISCCPath installers\simple_installer.iss
        if ($?) {
            Write-Host "Installer created successfully! Find it in the 'installers' folder: WinTab_Setup.exe" -ForegroundColor Green
        }
        else {
            Write-Host "Installer compilation failed." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Inno Setup compiler not found at $ISCCPath. Skipping installer creation." -ForegroundColor Yellow
    }
}
else {
    Write-Host "Build failed." -ForegroundColor Red
    exit 1
}
