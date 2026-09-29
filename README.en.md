<div align="center">

![WinTab Logo](Assets/128x128.png)

# WinTab

File Explorer tab utility for Windows 11

[简体中文](README.md) | English

![Platform](https://img.shields.io/badge/platform-Windows%2011-555555) ![.NET](https://img.shields.io/badge/.NET-9.0-555555) ![License](https://img.shields.io/badge/license-MIT-555555)

</div>

WinTab folds your File Explorer windows into one: newly opened windows become tabs automatically, the same folder reuses its existing tab, and your tab groups and recently closed tabs can be brought back at any time.

---

![](Assets/UI.png)

## Features

- One window, multiple tabs: newly opened File Explorer windows are merged into the active window as tabs
- Reuse the existing tab instead of opening the same folder twice
- Tab groups are saved continuously and can be recovered after shutdown or an Explorer crash; recently closed tabs can be reopened too
- Scroll the wheel over the tab row to switch to the previous or next tab, with Low / Medium / High sensitivity
- Double-click a tab title to close it
- Middle-click a folder in the navigation pane, file list, Home or the address bar to open it in the foreground
- When native focus such as "Open file location" lands on a background tab, that tab is brought to the front
- Hold `Ctrl + Shift` while opening folders to keep a separate window; drag a tab off the tab row to split it into a new window
- Runs from the system tray, starts with Windows, and switches between light / dark themes and Chinese / English

The last two days of logs are recorded in the background at `%LOCALAPPDATA%\WinTab\logs\`; the settings page has an "Open logs" button.

## Restore last tab group

WinTab records tab groups while running and restores only the last-used window: normal closure saves the last closed window, while shutdown or a crash restores the latest record of the last-used window. **Restore single-tab windows** is off by default: only windows with two or more tabs are recorded, and closing a single-tab window does not replace the saved group.

The tray's recovery section and the settings page provide:

- **Restore last tab group**: restore the group in a new window, keeping tab order and the saved active tab. Default shortcut `Alt + E`, works globally
- **Reopen last closed tab**: reopen the most recently closed tab, preferring the current window. Enabled by default; default shortcut `Alt + W`, works only while Explorer is in the foreground. Up to 25 closures are kept until WinTab exits. When the most recently closed location can no longer be opened (a network location or an unplugged drive, for example), the newest tab that can be opened is reopened instead and the skip is reported

Both can be disabled or re-bound in settings; shortcuts use `Ctrl` / `Alt` with a letter, top-row digit or `F1`–`F24`, and `Shift` is optional. Folder records live in `%APPDATA%\WinTab\session.json` and its backups; treat them as personal folder information.

### Automatic restoration (optional)

**Auto-restore last tab group** is off by default. When enabled, only a newly opened window can trigger it, and only when no other File Explorer windows are open:

- **Normal launch only (default)**: restore on launch to Home, This PC or the configured start folder
- **Any folder**: keep the initial tab and append the saved tabs after it

Torn-off windows, `Ctrl + Shift` windows and windows already open when the feature is enabled do not trigger restoration; missing or unsupported locations are skipped. With Windows' **Restore previous folder windows at logon** turned on in Folder Options, Windows itself reopens the windows, tabs included, that were still open at shutdown or when Explorer crashed; WinTab then does not add the same group again, and it can still be restored from the tray or with the shortcut.

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
