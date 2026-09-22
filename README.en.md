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
- Open middle-clicked folders in the foreground, whether clicked in the navigation pane, the file list or Home
- Double-click tab titles to close tabs
- Hold `Ctrl + Shift` while opening folders to keep a separate window
- System tray background runtime

## Download

Pick the installer that matches your CPU architecture ([Releases](https://github.com/OUBIGFA/WinTab/releases)):

- `WinTab_<version>_x64_Setup.exe` — 64-bit Intel/AMD (most users)
- `WinTab_<version>_arm64_Setup.exe` — ARM64 Windows (Surface Pro X, Snapdragon laptops)
- `WinTab_<version>_x86_Setup.exe` — 32-bit Windows

## Requirements

- Windows 11 22H2 or newer
- .NET 9 Desktop Runtime (the installer downloads it on first run if missing)

## License & Acknowledgements

Released under the MIT License.

Portions of the source code are derived from [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility) (Copyright (c) w4po, MIT License). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) for attribution and the full license text.
