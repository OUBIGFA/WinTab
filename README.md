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
- 可选恢复最后关闭窗口的标签组，单标签恢复可独立开启（默认关闭）
- 中键前台打开标签页（在侧边栏、文件列表或主页中中键点击文件夹时一律在前台打开）
- 双击标签页标题直接关闭
- 按住 `Ctrl + Shift` 打开文件夹可保留独立窗口
- 托盘后台静默运行与开机启动

## 恢复上次标签组

在设置中开启「恢复上次标签组」（默认关闭）。开启后，WinTab 会保存最后关闭窗口的标签组；不是恢复所有历史窗口。「恢复单个标签」独立开关默认关闭，此时仅保存和恢复含两个及以上标签的窗口，关闭单标签窗口不会覆盖已保存的标签组；开启后单标签窗口也会保存并恢复。

仅当没有其他资源管理器窗口时，下次新打开的窗口才可能触发恢复，与「合并新窗口」独立：

- **仅普通启动（默认）**：从主页、此电脑或资源管理器选项中设定的启动位置普通启动时恢复；主动打开特定文件夹不恢复，避免打扰。
- **打开任意文件夹**：始终保留最初打开的标签并停留在它上面，在后方补回历史标签。从任务栏打开的主页、此电脑或自定义启动位置也会保留，不会被关闭。

拖出的窗口、按住 `Ctrl + Shift` 打开的独立窗口，以及开启此功能或 WinTab 启动时已存在的窗口不触发恢复。已不存在的文件夹、网络位置和不支持的位置会被跳过，回收站、图库等内置页面可以恢复。恢复过程中切换窗口、切换标签或浏览到别处会立即停止恢复；若系统不支持在后台创建标签，恢复也会停止，已打开的文件夹不受影响。

## 诊断日志

WinTab 默认记录运行与操作日志，无需手动开启。日志位于 `%LOCALAPPDATA%\WinTab\logs\`，按日期命名为 `WinTab-yyyyMMdd.log`，仅保留今天和昨天。单文件最多 8 MiB，超限滚动保留一个 `.1.log` 文件，两天合计最多约 32 MiB。

遇到中键未激活或恢复异常时，可按发生时间查看点击结果、取消原因和恢复阶段；分享日志前请检查其中的目录路径等信息。日志在后台写入，队列有上限，连续相同记录会合并计数。

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
