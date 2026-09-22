<div align="center">

![WinTab Logo](Assets/128x128.png)

# WinTab

Windows 11 文件资源管理器标签页工具

简体中文 | [English](README.en.md)

![Platform](https://img.shields.io/badge/platform-Windows%2011-555555) ![.NET](https://img.shields.io/badge/.NET-9.0-555555) ![License](https://img.shields.io/badge/license-MIT-555555)

</div>

WinTab 是一款 Windows 11 工具，实现「一窗多标签」：自动将新打开的资源管理器窗口合并为当前窗口的标签页，支持路径去重、双击关闭、独立窗口保留和托盘后台运行，让文件管理更整洁高效。

---

![](Assets/UI.png)

## 功能

- 一窗多标签：自动将新打开的资源管理器窗口合并为当前窗口标签页
- 复用同一标签页，避免重复打开相同文件夹
- 中键前台打开标签页（在侧边栏、文件列表或主页中中键点击文件夹时一律在前台打开）
- 双击标签页标题直接关闭
- 按住 `Ctrl + Shift` 打开文件夹可保留独立窗口
- 托盘后台静默运行与开机启动

## 下载

根据 CPU 架构选择对应安装器（[Releases](https://github.com/OUBIGFA/WinTab/releases)）：

- `WinTab_<版本>_x64_Setup.exe` — 64 位 Intel/AMD（大多数用户）
- `WinTab_<版本>_arm64_Setup.exe` — ARM64 Windows（Surface Pro X 等）
- `WinTab_<版本>_x86_Setup.exe` — 32 位 Windows

## 系统要求

- Windows 11 22H2 或更高版本
- .NET 9 Desktop Runtime（首次运行安装器会自动下载安装）

## 许可与致谢

本项目以 MIT License 发布。

部分源代码派生自 [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)（Copyright (c) w4po，MIT License）。相关归属与许可全文见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)。
