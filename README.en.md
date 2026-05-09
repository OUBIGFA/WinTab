<div align="center">
  <img src="Assets/wintab-logo.png" width="96" alt="WinTab Logo">
  <h1>WinTab</h1>
  <p>File Explorer tab utility for Windows 11</p>
  <p>
    <a href="README.md">简体中文</a> | English
  </p>
  <p>
    <img alt="Platform" src="https://img.shields.io/badge/platform-Windows%2011-2563eb">
    <img alt=".NET" src="https://img.shields.io/badge/.NET-9.0-512bd4">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-111827">
  </p>
</div>

## Features

- Merge newly opened File Explorer windows
- Reuse tabs for paths that are already open
- Close the current tab by double-clicking the tab title
- Hold `Ctrl + Shift` for a separate Explorer window
- System tray runtime
- Chinese and English UI
- Light / dark themes
- Startup registration
- GitHub Release update checks

## Download

- `WinTab_v1.0.0_Setup.exe`

## Requirements

- Windows 11 22H2 or newer
- .NET 9 Desktop Runtime

## Local Build

### Environment

- Visual Studio 2022 or newer
- `.NET desktop development`
- .NET 9 SDK
- Inno Setup 6

### Publish

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1

foreach ($arch in "x64", "x86", "arm64") {
  & $msbuild ".\WinTab\WinTab.csproj" `
    /restore `
    /t:Publish `
    /p:Configuration=Release `
    /p:TargetFramework=net9.0-windows `
    /p:RuntimeIdentifier="win-$arch" `
    /p:PublishDir="..\publish\net9.0-windows\$arch"
}
```

### Installer

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DPublishRoot="..\publish\net9.0-windows" `
  /DOutputDir="..\dist" `
  ".\installers\installer.iss"
```

```text
dist/WinTab_v1.0.0_Setup.exe
```

## Project Structure

```text
WinTab/
├── .github/
├── Assets/
├── installers/
├── WinTab/
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

## Settings

```text
%APPDATA%\WinTab\settings.json
```

## Acknowledgements

- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)

## License

MIT License
