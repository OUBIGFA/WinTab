<div align="center">

![WinTab Logo](Assets/128x128.png)

# WinTab

File Explorer tab utility for Windows 11

[简体中文](README.md) | English

![Platform](https://img.shields.io/badge/platform-Windows%2011-555555) ![.NET](https://img.shields.io/badge/.NET-9.0-555555) ![License](https://img.shields.io/badge/license-MIT-555555)

</div>

WinTab is a Windows 11 utility that automatically merges newly opened File Explorer windows into tabs, with path deduplication, double-click to close, independent window retention via Ctrl+Shift, and system tray operation for more efficient file management.

---

![](Assets/UI.png)

## Features

- Merge new File Explorer windows into tabs
- Reuse opened tabs instead of opening duplicates
- Open middle-clicked navigation pane tabs in the foreground
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

## Acknowledgements

- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)

## License

MIT License
