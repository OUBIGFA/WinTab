<div align="center">

<img src="src/WinTab.App/Assets/wintab.ico" width="80" />

# WinTab

**Open-source tab manager for Windows Explorer**

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)]()
[![.NET 9](https://img.shields.io/badge/.NET-9.0-512BD4.svg)]()

English | [简体中文](README.md)

</div>

---

## Introduction

WinTab is an open-source Windows window tab manager. It intercepts system folder-open events so you can manage File Explorer windows as **browser-style tabs** — all within a single window.

Thanks to its fully event-driven Win32 architecture, **CPU usage stays near 0% at idle**, with a memory footprint of roughly 30–80 MB.

---

## Features

### Core

- **Tab interception**: Catches folder-open requests from the OS or other apps and routes them into the current WinTab window as new tabs — no new windows
- **Clone current path**: New tabs default to the current directory instead of "This PC"
- **Open subfolders in new tabs**: Entering a subfolder spawns a fresh tab without overwriting your current view
- **Double-click title to close**: Double-click a tab's title bar to close it — equivalent to a middle-click

### System Integration

- **Tray resident**: Auto-start at boot + start hidden in tray; restore any time from the notification area
- **Tray icon toggle**: Show or hide the WinTab icon in the system notification area independently
- **Themes**: Light and dark modes
- **Bilingual UI**: Simplified Chinese / English, instant switch in Settings

### Reliability

- Automatic run log and crash log generation
- Registry changes are automatically restored on exit or uninstall — no leftover state
- Built-in registry state self-check on startup to prevent environment corruption

---

## System Requirements

| Requirement | Detail |
|---|---|
| OS | Windows 10 / 11 (x64 only) |
| Runtime | [.NET 9 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/9.0) |
| Recommended | Windows 11 (full tab interception support) |

> **Note**: The folder open-verb interception feature requires Windows 11 (`10.0.22000+`). The app runs on Windows 10 but this feature will be unavailable.

---

## Download & Install

Visit the [Releases](../../releases) page to get the latest version.

### Installed Version (Recommended)

Download `WinTab_Setup_<version>.exe` and run it.

- Custom installation path supported
- Automatically detects missing .NET 9 runtime and guides you to download it
- Auto-detects system language (Chinese / English)
- Optional **Run at Windows startup** task during setup

#### Reinstallation

Running the installer when a version is already installed presents two options:

| Mode | Description |
|---|---|
| **Uninstall then install (recommended)** | Fully uninstalls the old version first; optionally remove user data |
| **Install directly** | Overwrites program files, preserves all user settings |

Silent install parameters:

```
WinTab_Setup_<version>.exe /SILENT /REINSTALLMODE=CLEAN /REMOVEUSERDATA=1
```

| Parameter | Description |
|---|---|
| `/REINSTALLMODE=CLEAN` | Uninstall first, then install |
| `/REINSTALLMODE=DIRECT` | Overwrite directly (default in silent mode) |
| `/REMOVEUSERDATA=1` | CLEAN mode only — deletes `%AppData%\Roaming\WinTab` |

---

### Portable Version (No Install Required)

Download `WinTab_<version>_portable.zip` and extract it. Run `WinTab.exe` directly.

The archive contains a `portable.txt` marker file. The app automatically enters portable mode:

- Config and logs are written to a local `data/` folder next to the executable
- To remove, simply delete the extracted folder

---

## Data Locations

### Installed

| Type | Path |
|---|---|
| Config | `%AppData%\WinTab\settings.json` |
| Log | `%AppData%\WinTab\logs\wintab.log` |
| Crash log | `%AppData%\WinTab\logs\crash.log` |

### Portable

| Type | Path |
|---|---|
| Config | `<extracted folder>\data\settings.json` |
| Log | `<extracted folder>\data\logs\wintab.log` |
| Crash log | `<extracted folder>\data\logs\crash.log` |

### Example `settings.json`

```json
{
  "StartMinimized": false,
  "ShowTrayIcon": true,
  "RunAtStartup": false,
  "Language": "English",
  "Theme": "Light",
  "EnableExplorerOpenVerbInterception": true,
  "OpenChildFolderInNewTabFromActiveTab": false,
  "CloseTabOnDoubleClick": true,
  "SchemaVersion": 1
}
```

---

## Resource Usage

WinTab is built for long-term background residency. All monitoring is purely event-driven — **no polling loops**.

| Resource | Idle | Active use |
|---|---|---|
| CPU | ≈ 0% | Brief spikes on window events |
| Memory | 30–80 MB | No significant growth |
| Disk | No background writes | Log / settings writes only |
| Network | None | None |

**Internal mechanisms**:

- `SetWinEventHook` (Win32) — passive system window event callbacks, CPU-friendly
- `NamedPipeServerStream.WaitForConnectionAsync` — async `await` blocks until a client connects; no CPU consumed while idle
- `EventWaitHandle.WaitOne()` — OS-level blocking for single-instance activation signals
- `RegistryOpenVerbInterceptor` — one-time self-check at startup only; no periodic registry polling

---

## Uninstall

From the **Uninstall** page inside the app:

- **Keep user data** (default): Removes program files, leaves `%AppData%\WinTab` intact
- **Full cleanup**: Check "Remove user data" to also delete config and logs

The uninstall process automatically restores all `Folder/Directory/Drive` open-verb registry entries to their original values.

---

## How It Works

```
System / app opens a folder
        │
        ▼
RegistryOpenVerbInterceptor intercepts the open-verb
        │
        ▼
ExplorerOpenVerbHandler spawns (--wintab-open-folder)
        │
        ▼
NamedPipe → ExplorerOpenRequestServer (main process)
        │
        ▼
WindowManager finds / activates an existing Explorer window
        │
        ▼
Folder opens as a new tab in the existing window
```

---

## Project Structure

```
src/
  WinTab.App/             WPF host (DI, pages, services, tray, startup)
  WinTab.UI/              UI resources & localization (zh/en)
  WinTab.Core/            Core models and interfaces
  WinTab.Platform.Win32/  Win32 interop and window manipulation
  WinTab.Persistence/     Config and path management
  WinTab.Diagnostics/     Logging and crash handling
  WinTab.Tests/           Unit tests
installers/
  WinTab.iss              Inno Setup installer script
.github/workflows/
  build-release.yml       CI/CD (build → test → package → release)
```

---

## Local Development

### Prerequisites

- .NET SDK 9.x
- Windows (PowerShell recommended)
- Optional: [Inno Setup 6](https://jrsoftware.org/isinfo.php) (for building the installer)

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

## Automated Releases (GitHub Actions)

- **Push to `master`**: Builds, tests, packages and uploads as Actions Artifacts
- **Push a tag** (e.g., `1.0.0`): Creates a GitHub Release with the installer, portable zip, and their SHA256 files

```powershell
git push origin master
git tag 1.0.0
git push origin 1.0.0
```

---

## FAQ

**Q: I get a prompt about missing .NET 9 runtime.**
A: The installer will guide you to the download page. You can also manually visit https://dotnet.microsoft.com/download/dotnet/9.0 and download the Desktop Runtime.

**Q: Will my settings survive an uninstall?**
A: Yes, by default. Check "Remove user data" during uninstall to delete them.

**Q: Does it work on Windows 10?**
A: The app runs on Windows 10, but the folder interception feature requires the Windows 11 File Explorer tab interface and will not function on Windows 10.

**Q: The GitHub Release wasn't created automatically. Why?**
A: Verify that a `x.y.z` format tag was pushed, the workflow file exists in the tagged commit, Actions is enabled in the repository, and the workflow has `contents: write` permission.

---

## License

Released under the [MIT License](LICENSE).

---

## Acknowledgements

- Explorer tab handling logic adapted from [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility) (MIT)
- Inno Setup Simplified Chinese translation from the community, stored at `installers/Languages/ChineseSimplified.isl`
