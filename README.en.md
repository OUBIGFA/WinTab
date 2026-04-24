<div align="center">
  <img src="WinTab/wintab-logo.png" width="96" alt="WinTab Logo">
  <h1>WinTab</h1>
  <p>File Explorer tab control for Windows 11</p>
  <p>
    English | <a href="README.md">简体中文</a>
  </p>
  <p>
    <img alt="Platform" src="https://img.shields.io/badge/platform-Windows%2011-2563eb">
    <img alt=".NET" src="https://img.shields.io/badge/.NET-9.0%20%7C%204.8.1-512bd4">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-111827">
  </p>
</div>

## Overview

WinTab is a focused Windows 11 utility for a cleaner File Explorer tab workflow. It merges newly opened folder windows into an existing File Explorer window and reuses an existing tab when the target path is already open.

WinTab does not take over the system folder association and does not replace Shell behavior through registry redirection. It uses Windows Shell COM and window observation, so folders opened by third-party applications still land on the intended target path.

## Features

- Merge newly opened File Explorer folder windows into existing tabs.
- Reuse an existing tab when the same path is already open.
- Preserve correct target paths when folders are opened by third-party apps.
- Close a File Explorer tab by double-clicking its tab title; blank title-bar areas keep native Windows behavior.
- Hold `Ctrl + Shift` while opening a folder to keep it as a separate File Explorer window.
- Tray-first runtime with a compact control panel.
- Chinese and English UI switching.
- Light and dark themes.
- Startup registration and GitHub Release update checks.

## System Requirements

- Windows 11 22H2 or newer.
- .NET 9 Desktop Runtime is recommended.
- The installer selects x64, x86, or ARM64 packages according to the system architecture.
- If .NET 9 Desktop Runtime is unavailable, the installer can fall back to the bundled .NET Framework 4.8.1 build or prompt for the runtime.

## Download and Usage

1. Download `WinTab_v*_Setup.exe` from GitHub Releases.
2. Run the installer and finish setup.
3. Start WinTab. It will stay in the system tray.
4. Open folders normally to let WinTab route them into the File Explorer tab flow.
5. Hold `Ctrl + Shift` while opening a folder when you need a separate File Explorer window.

Settings are stored at:

```text
%APPDATA%\WinTab\settings.json
```

## Local Build

### Requirements

- Visual Studio 2022 or newer.
- Workload: `.NET desktop development`.
- .NET 9 SDK.
- .NET Framework 4.8.1 Developer Pack when building the .NET Framework target.
- Inno Setup 6 when building the installer.

### Build the App

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1

& $msbuild ".\WinTab\WinTab.csproj" `
  /restore `
  /t:Publish `
  /p:Configuration=Release `
  /p:TargetFramework=net9.0-windows `
  /p:RuntimeIdentifier=win-x64 `
  /p:PublishDir="..\publish\net9.0-windows\x64"
```

### Build the Installer

The installer script is `installers/installer.iss`. It expects the matching framework and architecture ZIP packages to exist in `artifacts`.

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DSourceDir="..\artifacts" `
  ".\installers\installer.iss"
```

The generated installer is written to `artifacts`.

## GitHub Actions

The repository includes release workflows:

- `.github/workflows/build-release.yml`: builds x64, x86, and ARM64 packages for .NET 9 and .NET Framework 4.8.1, builds and signs the installer, then creates a Release.
- `.github/workflows/publish-winget.yml`: publishes to WinGet.
- `.github/workflows/publish-chocolatey.yml`: publishes to Chocolatey.

## Project Structure

```text
WinTab/
├── .github/              # Issue templates and release workflows
├── artifacts/            # Local build output and installers
├── installers/           # Inno Setup scripts and language files
├── tests/                # File Explorer behavior probes
├── WinTab/               # WPF application source
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

## FAQ

### Why did a folder not merge into a tab?

Make sure WinTab is running and the automatic merge option is enabled in the control panel. If `Ctrl + Shift` was held while opening the folder, WinTab keeps it as a separate File Explorer window.

### Will folders opened by third-party apps still work?

Yes. WinTab waits until the File Explorer window reaches the real target path before merging or reusing tabs, so third-party folder opens should keep their intended destination.

### Where should I double-click to close a tab?

Only the File Explorer tab title area closes a tab on double-click. Blank title-bar space, window borders, and other non-tab-title areas keep native Windows behavior.

### How do I fully exit WinTab?

Open the menu from the system tray icon and choose Exit. Closing the control panel window does not quit WinTab.

## Credits

WinTab's File Explorer tab engine evolves from the Shell/COM approach used by ExplorerTabUtility, with the product scope narrowed to File Explorer tab merging, reuse, and closing behavior.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
