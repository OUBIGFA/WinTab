<div align="center">

![WinTab Logo](Assets/128x128.png)

# WinTab

File Explorer tab utility for Windows 11

[简体中文](README.md) | English

![Platform](https://img.shields.io/badge/platform-Windows%2011-555555) ![.NET](https://img.shields.io/badge/.NET-9.0-555555) ![License](https://img.shields.io/badge/license-MIT-555555)

</div>

WinTab is a Windows 11 utility that brings a single-window, multi-tab experience by automatically merging newly opened File Explorer windows into the active window as tabs, with path deduplication, double-click to close, independent window retention via Ctrl+Shift, and system tray operation for more efficient file management.

---

![](Assets/UI.png)

## Features

- One window, multiple tabs: automatically merge new File Explorer windows into active tabs
- Reuse opened tabs instead of opening duplicates
- Continuously save tab groups for shutdown/crash recovery, with manual group and closed-tab recovery
- Open middle-clicked folders in the foreground, whether clicked in the navigation pane, the file list or Home
- Double-click tab titles to close tabs
- Hold `Ctrl + Shift` while opening folders to keep a separate window
- System tray background runtime

## Restore last tab group

WinTab records tab groups while running and restores only the last-used window: normal closure saves the last closed window, while shutdown or a crash restores the latest record of the last-used window. **Restore single-tab windows** is off by default: only windows with two or more tabs are recorded, and closing a single-tab window does not replace the saved group.

The tray's recovery section and the settings page provide:

- **Restore last tab group**: restore the group in a new window, keeping tab order and the saved active tab. Shortcut `Alt+E`, works globally.
- **Reopen last closed tab**: reopen the most recently closed tab, preferring the current window. Enabled by default; shortcut `Alt+W`, works only while Explorer is in the foreground. Up to 25 closures are kept until WinTab exits. When the most recently closed location can no longer be opened (a network location or an unplugged drive, for example), the newest tab that can be opened is reopened instead and the skip is reported.

Both can be disabled or re-bound in settings. Folder records live in `%APPDATA%\WinTab\session.json` and its backups; treat them as personal folder information.

### Automatic restoration (optional)

**Auto-restore last tab group** is off by default. When enabled, only a newly opened window can trigger it, and only when no other File Explorer windows are open:

- **Normal launch only (default)**: restore on launch to Home, This PC or the configured start folder.
- **Any folder**: keep the initial tab and append the saved tabs after it.

Torn-off windows, `Ctrl + Shift` windows and windows already open when the feature is enabled do not trigger restoration; missing or unsupported locations are skipped. Switching windows or tabs stops the restoration.

With Windows' **Restore previous folder windows at logon** turned on in Folder Options, Windows itself reopens the windows, tabs included, that were still open at shutdown or when Explorer crashed; automatic restoration then does not add the same group again. It can still be restored from the tray or with the shortcut.

## Diagnostic logs

WinTab records diagnostic logs by default; no setup is needed. Files are stored in `%LOCALAPPDATA%\WinTab\logs\` as `WinTab-yyyyMMdd.log`. Only today's and yesterday's logs are kept. Each file is limited to 8 MiB with one `.1.log` rollover file, bounding the two days to about 32 MiB in total.

If a middle-click does not activate its tab or restoration behaves unexpectedly, check the matching time for click outcomes, cancellation reasons and restoration stages. Review folder paths and other details before sharing logs. Writing happens in the background through a bounded queue; consecutive duplicate entries are collapsed into a count.

## Download

Pick the installer that matches your CPU architecture ([Releases](https://github.com/OUBIGFA/WinTab/releases)):

- `WinTab_<version>_x64_Setup.exe` — 64-bit Intel/AMD (most users)
- `WinTab_<version>_arm64_Setup.exe` — ARM64 Windows (Surface Pro X, Snapdragon laptops)
- `WinTab_<version>_x86_Setup.exe` — 32-bit Windows

In-app updates check the downloaded installer against the SHA-256 GitHub publishes for it and do not run an installer that does not match.

## Requirements

- Windows 11 22H2 or newer
- .NET 9 Desktop Runtime (the installer downloads it on first run if missing)

## License & Acknowledgements

Released under the MIT License.

Portions of the source code are derived from [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility) (Copyright (c) w4po, MIT License). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) for attribution and the full license text.
