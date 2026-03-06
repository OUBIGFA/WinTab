<div align="center">

<img src="src/WinTab.App/Assets/logo.png" width="80" />

# WinTab

**Open-source tab manager for Windows Explorer**

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)]()
[![.NET 9](https://img.shields.io/badge/.NET-9.0-512BD4.svg)]()

English | [简体中文](README.md)

</div>

---

## Overview

WinTab is an open-source tab manager for Windows Explorer. It intercepts or absorbs folder-open requests and routes them into tabs inside an existing Explorer window whenever possible, instead of leaving every navigation as a separate top-level window.

The project uses an event-driven Win32 + COM architecture and is designed for long-running background use without polling loops.

---

## Recent Improvements

Recent fixes and behavior refinements focused on these areas:

- **Better system-folder compatibility**: improved normalization and navigation for Shell namespace targets such as `shell:`, `::{GUID}`, and bare GUID values.
- **Fixed “inherit current tab path” behavior**: when enabled, a newly created tab now correctly aligns to the active tab path even if Explorer first opens that tab at `This PC`.
- **Reduced conversion flicker**: when direct interception is not possible and WinTab has to convert a newly opened Explorer window into a tab, the source window is hidden earlier and kept suppressed during conversion.
- **Expanded regression coverage**: additional tests now cover path normalization, Shell namespace navigation, and new-tab default correction logic.

---

## Features

### Core

- **Folder-open interception**: captures folder-open requests from the OS or other apps and opens them as tabs in an existing Explorer window whenever possible.
- **Inherit current tab path**: new tabs can default to the active tab path instead of `This PC`.
- **Open subfolders in new tabs**: optionally keep the current view and open child folders in separate tabs.
- **Double-click to close tab**: optionally close an Explorer tab by double-clicking its title.

### System Integration

- **Tray resident**: supports start minimized, tray residency, and restore from the notification area.
- **Run at startup**: optional startup integration.
- **Theme support**: light and dark modes.
- **Bilingual UI**: instant switch between Simplified Chinese and English.

### Reliability

- **Low idle overhead**: event-driven architecture with near-zero idle CPU usage.
- **Runtime and crash logs**: useful for diagnosing interception, navigation, and conversion behavior.
- **Registry self-check**: validates the Explorer open-verb takeover state on startup.
- **State restoration on exit/uninstall**: restores registry entries that WinTab has taken over.

---

## System Requirements

| Requirement | Detail |
|---|---|
| OS | Windows 10 / 11 (x64 only) |
| Runtime | [.NET 9 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/9.0) |
| Recommended | Windows 11 |

> Note: full Explorer tab interception depends on the Windows 11 Explorer tab environment. The app can still run on Windows 10, but some tab takeover features are unavailable there.

---

## How It Works

WinTab handles Explorer navigations with two layers:

### 1. Direct interception (preferred)

When the open request can be intercepted before Explorer creates a new window:

```text
System / app opens a folder
  -> RegistryOpenVerbInterceptor / ShellBridge intercepts the request
  -> Named Pipe forwards it to the main process
  -> WinTab locates an existing Explorer window
  -> The target opens directly as a new tab
```

This is the ideal path because it avoids the “open window, then close/merge it” sequence entirely.

### 2. Auto-convert fallback

If a new Explorer window cannot be intercepted in advance, WinTab falls back to:

- watching for new Explorer top-level windows
- hiding the new window as early as possible
- resolving its target location
- reopening that location in an existing Explorer window as a tab
- closing the source window

This fallback cannot eliminate window creation at the OS level, but the current implementation reduces visible flicker as much as possible.

---

## System Folder Compatibility

WinTab now has stronger handling for Shell namespace targets, including:

- `shell:`
- `::{GUID}`
- bare GUID values such as `{645FF040-5081-101B-9F08-00AA002F954E}`

These locations are normalized internally and are navigated via `Shell.Application.NameSpace(...)` + `Navigate2(object)` instead of being treated as plain filesystem paths.

This notably improves compatibility with:

- Recycle Bin
- This PC
- other Shell namespace folders

---

## Download & Install

Get the latest build from the [Releases](../../releases) page.

### Installer Build (Recommended)

Download `WinTab_Setup_<version>.exe` and run it.

- custom install path supported
- auto-detects missing .NET 9 Desktop Runtime
- initializes UI language from the system language
- can optionally create a Windows startup task

#### Reinstall Modes

If WinTab is already installed, the installer supports:

| Mode | Description |
|---|---|
| Uninstall then install | removes the old version first, then installs the new one; can also remove user data |
| Install directly | overwrites program files and keeps user settings |

Silent install example:

```powershell
WinTab_Setup_<version>.exe /SILENT /REINSTALLMODE=CLEAN /REMOVEUSERDATA=1
```

| Parameter | Description |
|---|---|
| `/REINSTALLMODE=CLEAN` | uninstall first, then install |
| `/REINSTALLMODE=DIRECT` | overwrite directly |
| `/REMOVEUSERDATA=1` | only meaningful in CLEAN mode; removes `%AppData%\WinTab` |

### Portable Build

Download `WinTab_<version>_portable.zip`, extract it, and run `WinTab.exe`.

The archive includes a `portable.txt` marker file, so the app automatically switches to portable mode:

- config and logs are stored in a local `data/` folder
- `%AppData%` is not used
- removal is just deleting the extracted folder

---

## Data Locations

### Installed

| Type | Path |
|---|---|
| Config | `%AppData%\WinTab\settings.json` |
| Runtime log | `%AppData%\WinTab\logs\wintab.log` |
| Crash log | `%AppData%\WinTab\logs\crash.log` |

### Portable

| Type | Path |
|---|---|
| Config | `<app folder>\data\settings.json` |
| Runtime log | `<app folder>\data\logs\wintab.log` |
| Crash log | `<app folder>\data\logs\crash.log` |

### Example `settings.json`

```json
{
  "StartMinimized": false,
  "ShowTrayIcon": true,
  "RunAtStartup": false,
  "Language": "English",
  "Theme": "Light",
  "EnableExplorerOpenVerbInterception": true,
  "PersistExplorerOpenVerbInterceptionAcrossExit": false,
  "OpenNewTabFromActiveTabPath": true,
  "OpenChildFolderInNewTabFromActiveTab": false,
  "CloseTabOnDoubleClick": false,
  "EnableAutoConvertExplorerWindows": true,
  "SchemaVersion": 2
}
```

### Key Settings

| Setting | Purpose |
|---|---|
| `EnableExplorerOpenVerbInterception` | enables Explorer open-verb interception |
| `PersistExplorerOpenVerbInterceptionAcrossExit` | keeps interception active after app exit |
| `OpenNewTabFromActiveTabPath` | inherits the active tab path for new tabs |
| `OpenChildFolderInNewTabFromActiveTab` | opens child folders in new tabs |
| `CloseTabOnDoubleClick` | closes a tab on double-click |
| `EnableAutoConvertExplorerWindows` | converts non-intercepted Explorer windows into tabs |

---

## Resource Usage

WinTab is designed for long-lived background execution and avoids active polling.

| Resource | Idle | Active |
|---|---|---|
| CPU | near 0% | short spikes during window events, navigation, or conversion |
| Memory | about 30–50 MB | no meaningful sustained growth |
| Disk | no background churn | settings/log writes only |
| Network | none | none |

Main internal mechanisms:

- `SetWinEventHook`
- `NamedPipeServerStream.WaitForConnectionAsync`
- `EventWaitHandle.WaitOne()`
- Explorer COM navigation and Shell namespace resolution

---

## Uninstall

From the in-app uninstall page you can choose:

- **Keep user data**: remove program files and keep config/logs
- **Full cleanup**: remove config/logs as well

The uninstall flow restores the Explorer open-verb registry entries that WinTab took over.

---

## Project Structure

```text
src/
  WinTab.App/             WPF host (DI, pages, services, tray, startup)
  WinTab.UI/              UI resources and localization
  WinTab.Core/            Core models and interfaces
  WinTab.Platform.Win32/  Win32 interop and window operations
  WinTab.Persistence/     Configuration, paths, persistence
  WinTab.Diagnostics/     Logging and crash handling
  WinTab.ShellBridge/     Shell open-verb / DelegateExecute bridge
  WinTab.Tests/           Unit tests
installers/
  WinTab.iss              Inno Setup installer script
.github/workflows/
  build-release.yml       CI / build / test / package
```

---

## Local Development

### Prerequisites

- Windows
- .NET SDK 9.x
- optional: [Inno Setup 6](https://jrsoftware.org/isinfo.php)

### Build & Test

```powershell
dotnet restore WinTab.slnx
dotnet build WinTab.slnx -c Release
dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release
dotnet run --project src/WinTab.App/WinTab.App.csproj
```

### Local Packaging

```powershell
# 1. Publish the app
dotnet publish src/WinTab.App/WinTab.App.csproj `
  -c Release -r win-x64 --self-contained false `
  -o publish/win-x64

# 2. Build the installer (requires Inno Setup 6)
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=1.0.0 installers/WinTab.iss

# 3. Create the portable ZIP
"WinTab Portable Mode" | Out-File publish/win-x64/portable.txt -Encoding utf8
Compress-Archive publish/win-x64\* publish/WinTab_1.0.0_portable.zip -Force
```

---

## Automated Releases

- push to `master`: build, test, package, upload GitHub Actions artifacts
- push a version tag such as `1.0.0`: create a GitHub Release with installer, portable ZIP, and checksums

```powershell
git push origin master
git tag 1.0.0
git push origin 1.0.0
```

---

## FAQ

**Q: The app says .NET 9 Desktop Runtime is missing.**  
A: The installer guides you to the download page, or you can install it manually from [dotnet 9 downloads](https://dotnet.microsoft.com/download/dotnet/9.0).

**Q: Does WinTab work on Windows 10?**  
A: The app runs on Windows 10, but some Explorer tab takeover features depend on the Windows 11 Explorer tab environment.

**Q: Why do I still sometimes see a new Explorer window flash briefly?**  
A: Some scenarios cannot be intercepted before Explorer creates the window, so WinTab must use the fallback “create window, then convert it into a tab” path. The current implementation suppresses visibility aggressively, but Explorer behavior still sets the lower bound.

**Q: Are settings preserved after uninstall?**  
A: Yes by default. They are removed only if you choose full cleanup.

---

## License

Released under the [MIT License](LICENSE).

---

## Credits

- Explorer tab handling logic is based on and adapted from [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)
- Installer packaging uses Inno Setup
