<div align="center">

![WinTab Logo](Assets/128x128.png)

# WinTab

File Explorer tab utility for Windows 11

[简体中文](README.md) | English

![Platform](https://img.shields.io/badge/platform-Windows%2011-555555) ![.NET](https://img.shields.io/badge/.NET-9.0-555555) ![License](https://img.shields.io/badge/license-MIT-555555)

WinTab is a Windows 11 utility that automatically merges newly opened File Explorer windows into tabs, with path deduplication, double-click to close, independent window retention via Ctrl+Shift, and system tray operation for more efficient file management.

---

![](Assets/UI.png)

## Features

- Merge newly opened File Explorer windows
- Reuse tabs for paths that are already open
- Close the current tab by double-clicking the tab title
- Hold `Ctrl + Shift` for a separate Explorer window
- System tray runtime

## Download

Pick the installer that matches your CPU architecture ([Releases](https://github.com/OUBIGFA/WinTab/releases)):

- `WinTab_v1.0.0_x64_Setup.exe` — 64-bit Intel/AMD (most users)
- `WinTab_v1.0.0_arm64_Setup.exe` — ARM64 Windows (Surface Pro X, Snapdragon laptops)
- `WinTab_v1.0.0_x86_Setup.exe` — 32-bit Windows

## Requirements

- Windows 11 22H2 or newer
- .NET 9 Desktop Runtime (the installer downloads it on first run if missing)

## Acknowledgements

- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)

## License

MIT License

