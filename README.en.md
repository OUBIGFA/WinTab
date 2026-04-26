<div align="center">
  <img src="Assets/wintab-logo.png" width="96" alt="WinTab Logo">
  <h1>WinTab</h1>
  <p>File Explorer tab control for Windows 11</p>
  <p>
    <a href="README.md">简体中文</a> | English
  </p>
  <p>
    <img alt="Platform" src="https://img.shields.io/badge/platform-Windows%2011-2563eb">
    <img alt=".NET" src="https://img.shields.io/badge/.NET-9.0-512bd4">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-111827">
  </p>
</div>

## Overview

WinTab is a lightweight Windows 11 utility for a cleaner File Explorer tab workflow. It merges newly opened folder windows into an existing Explorer window and reuses an existing tab when the target path is already open.

WinTab does not replace the system folder association or redirect Shell behavior through the registry. It works through Windows Shell COM and window observation, so folders opened by third-party applications still land on the correct target path.

## Features

- Merge newly opened File Explorer windows into existing tabs
- Reuse an existing tab when the same path is already open
- Preserve correct target paths for third-party folder launches
- Close a File Explorer tab by double-clicking its tab title
- Hold `Ctrl + Shift` while opening a folder to keep a separate window
- Tray-first runtime with a compact control panel
- Chinese and English UI switching
- Light and dark themes
- Startup registration and GitHub Release update checks

## Requirements

- Windows 11 22H2 or newer
- .NET 9 Desktop Runtime
- The installer detects a missing .NET 9 Desktop Runtime and prompts to download it during setup

## Download

Only one release artifact is provided:

- `WinTab_v1.0.0_Setup.exe`

Future releases also ship as installer `.exe` only. No zip packages, Chocolatey package, or WinGet package are maintained anymore.

## Local Build

### Prerequisites

- Visual Studio 2022 or newer
- Workload: `.NET desktop development`
- .NET 9 SDK
- Inno Setup 6

### Publish the App

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

### Build the Installer

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DPublishRoot="..\publish\net9.0-windows" `
  /DOutputDir="..\dist" `
  ".\installers\installer.iss"
```

Installer output:

```text
dist/WinTab_v1.0.0_Setup.exe
```

## GitHub Actions

The repository keeps a single release workflow:

- `.github/workflows/build-release.yml`

It builds the `x64`, `x86`, and `arm64` publish outputs, assembles one installer, and uploads only `WinTab_v*_Setup.exe` to the GitHub Release.

## Project Structure

```text
WinTab/
├── .github/                 # GitHub templates and release workflow
├── Assets/                  # Repository assets
├── installers/              # Inno Setup installer files
├── WinTab/                  # WPF application source
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

Local build output directories such as `publish/`, `dist/`, `bin/`, and `obj/` are intentionally excluded from the repository.

## Settings

Installed settings are stored at:

```text
%APPDATA%\WinTab\settings.json
```

## License

This project is licensed under the MIT License. See `LICENSE` for details.
