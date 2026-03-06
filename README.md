<div align="center">

<img src="src/WinTab.App/Assets/logo.png" width="80" />

# WinTab

**为 Windows 资源管理器提供浏览器式标签页管理的开源工具**

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)]()
[![.NET 9](https://img.shields.io/badge/.NET-9.0-512BD4.svg)]()

[English](README.en.md) | 简体中文

</div>

---

## 简介

WinTab 是一个面向 Windows Explorer 的标签页管理工具。它通过拦截或接管文件夹打开请求，把原本会打开成新窗口的 Explorer 导航，尽量合并到已有窗口中的新标签页里。

项目采用事件驱动的 Win32 + COM 架构，空闲时不依赖轮询，适合长期后台驻留。

---

## 最新状态

最近一轮修复和优化主要集中在以下几个方向：

- **系统文件夹兼容性增强**：改进了 `shell:`、`::{GUID}`、裸 GUID 等 Shell 命名空间路径的规范化与导航方式，减少“回收站 / 此电脑 / 其他系统文件夹”类目标的异常。
- **新建标签页继承当前路径修复**：开启“继承当前标签页路径”后，即使 Explorer 默认新标签页先落在“此电脑”，WinTab 也会继续把新标签页对齐到当前活动标签页路径。
- **打开/合并闪烁压制**：对于无法直接拦截、只能事后转换为标签页的 Explorer 新窗口，WinTab 会更早隐藏并持续压制源窗口可见性，降低“先弹窗再合并”的体感。
- **测试覆盖补强**：围绕路径规范化、Shell 命名空间导航、标签页默认位置修正等问题补充了回归测试。

---

## 功能特性

### 核心行为

- **文件夹打开接管**：拦截系统或外部应用发起的文件夹打开请求，优先直接在已有 Explorer 窗口中新建标签页。
- **继承当前标签页路径**：新建标签页默认跟随当前活动标签页路径，而不是停留在“此电脑”。
- **子文件夹新标签页打开**：可选将进入子文件夹的操作改为“在新标签页中打开”，保留当前浏览上下文。
- **双击标签页关闭**：可选支持双击 Explorer 标签页标题关闭标签页。

### 系统集成

- **托盘驻留**：支持启动时最小化、托盘常驻、从通知区域恢复主窗口。
- **开机自启**：可在设置中控制是否随 Windows 启动。
- **主题切换**：支持浅色 / 深色主题。
- **中英双语**：设置页可即时切换简体中文和 English。

### 稳定性与恢复

- **事件驱动低占用**：空闲时 CPU 占用接近 0。
- **运行日志与崩溃日志**：便于定位拦截、导航、窗口合并问题。
- **注册表状态自检**：启动时校验 Explorer open verb 相关接管状态。
- **退出/卸载恢复**：退出或卸载时恢复已接管的注册表项，避免残留系统状态。

---

## 系统要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（仅 x64） |
| 运行时 | [.NET 9 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/9.0) |
| 推荐系统 | Windows 11 |

> 说明：Explorer 标签页相关接管能力依赖 Windows 11 的 Explorer 标签页环境。Windows 10 可运行程序，但部分标签页接管能力不可用。

---

## 工作机制

WinTab 对 Explorer 打开行为采用两层策略：

### 1. 直接接管（首选）

当目标打开请求可以被 open verb / ShellBridge 捕获时：

```text
系统 / 应用请求打开文件夹
  -> RegistryOpenVerbInterceptor / ShellBridge 接管
  -> Named Pipe 发送到主进程
  -> 定位现有 Explorer 窗口
  -> 在现有窗口中直接打开为新标签页
```

这条路径通常没有“先新建窗口再关闭”的闪烁过程，也是 WinTab 的理想行为。

### 2. 自动转换（兜底）

对无法提前接管的 Explorer 新窗口，WinTab 会：

- 监听 Explorer 顶层窗口创建事件
- 尽量尽早隐藏新窗口
- 解析其目标位置
- 把该位置转移到现有窗口的新标签页
- 关闭源窗口

这条路径本质上是“事后折叠成标签页”，无法像直接接管那样彻底消除窗口创建，但当前版本已针对闪烁感做了额外压制。

---

## 对系统文件夹的兼容说明

WinTab 当前已针对以下 Shell 命名空间目标加强处理：

- `shell:`
- `::{GUID}`
- 裸 GUID（例如 `{645FF040-5081-101B-9F08-00AA002F954E}`）

这类位置在内部会被规范化，并优先通过 `Shell.Application.NameSpace(...)` + `Navigate2(object)` 导航，而不是简单按普通文件系统路径处理。

这能显著改善以下目标的兼容性：

- 回收站
- 此电脑
- 其他 Shell 命名空间文件夹

---

## 下载与安装

前往 [Releases](../../releases) 页面下载最新版本。

### 安装版（推荐）

下载 `WinTab_Setup_<version>.exe` 并运行。

- 支持自定义安装路径
- 自动检测是否缺少 .NET 9 Desktop Runtime
- 自动跟随系统语言初始化界面语言
- 安装时可选择是否创建开机自启任务

#### 重装模式

重复运行安装器时，支持两种方式：

| 模式 | 说明 |
|---|---|
| 卸载后重装（推荐） | 先卸载旧版本，再安装新版本；可选删除用户数据 |
| 直接覆盖安装 | 直接覆盖程序文件，保留原有配置 |

静默安装示例：

```powershell
WinTab_Setup_<version>.exe /SILENT /REINSTALLMODE=CLEAN /REMOVEUSERDATA=1
```

| 参数 | 说明 |
|---|---|
| `/REINSTALLMODE=CLEAN` | 先卸载再安装 |
| `/REINSTALLMODE=DIRECT` | 直接覆盖安装 |
| `/REMOVEUSERDATA=1` | 仅在 CLEAN 模式下删除 `%AppData%\WinTab` |

### 便携版

下载 `WinTab_<version>_portable.zip`，解压后直接运行 `WinTab.exe`。

便携包内包含 `portable.txt` 标记文件，程序会自动进入便携模式：

- 配置和日志写入程序目录下的 `data/`
- 不使用 `%AppData%`
- 删除整个解压目录即可完成移除

---

## 数据目录

### 安装版

| 类型 | 路径 |
|---|---|
| 配置文件 | `%AppData%\WinTab\settings.json` |
| 运行日志 | `%AppData%\WinTab\logs\wintab.log` |
| 崩溃日志 | `%AppData%\WinTab\logs\crash.log` |

### 便携版

| 类型 | 路径 |
|---|---|
| 配置文件 | `<程序目录>\data\settings.json` |
| 运行日志 | `<程序目录>\data\logs\wintab.log` |
| 崩溃日志 | `<程序目录>\data\logs\crash.log` |

### `settings.json` 示例

```json
{
  "StartMinimized": false,
  "ShowTrayIcon": true,
  "RunAtStartup": false,
  "Language": "Chinese",
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

### 关键配置说明

| 配置项 | 作用 |
|---|---|
| `EnableExplorerOpenVerbInterception` | 启用 Explorer open verb 接管 |
| `PersistExplorerOpenVerbInterceptionAcrossExit` | 退出后保留接管状态 |
| `OpenNewTabFromActiveTabPath` | 新建标签页继承当前活动标签页路径 |
| `OpenChildFolderInNewTabFromActiveTab` | 子文件夹改为在新标签页中打开 |
| `CloseTabOnDoubleClick` | 双击标签页标题关闭 |
| `EnableAutoConvertExplorerWindows` | 对未提前接管的新 Explorer 窗口做自动转换 |

---

## 资源占用

WinTab 面向长期后台运行设计，核心等待均基于系统事件或阻塞等待，不做主动轮询。

| 资源 | 空闲状态 | 活动状态 |
|---|---|---|
| CPU | 接近 0% | 仅在窗口事件、导航或转换时短暂波动 |
| 内存 | 约 30–50 MB | 无明显持续增长 |
| 磁盘 | 无后台刷写 | 仅设置和日志写入 |
| 网络 | 无 | 无 |

内部主要机制：

- `SetWinEventHook`
- `NamedPipeServerStream.WaitForConnectionAsync`
- `EventWaitHandle.WaitOne()`
- Explorer COM 导航与 Shell 命名空间解析

---

## 卸载

在应用内的卸载页面可以选择：

- **保留用户数据**：删除程序文件，保留配置和日志
- **完整清理**：连同配置、日志一起删除

卸载过程会恢复 WinTab 接管过的 Explorer open verb 注册表项。

---

## 项目结构

```text
src/
  WinTab.App/             WPF 主程序（DI、页面、服务、托盘、启动）
  WinTab.UI/              UI 资源与本地化
  WinTab.Core/            核心模型与接口
  WinTab.Platform.Win32/  Win32 互操作与窗口操作
  WinTab.Persistence/     配置、路径与持久化
  WinTab.Diagnostics/     日志与崩溃处理
  WinTab.ShellBridge/     Shell open verb / DelegateExecute 桥接
  WinTab.Tests/           单元测试
installers/
  WinTab.iss              Inno Setup 安装脚本
.github/workflows/
  build-release.yml       CI / build / test / package
```

---

## 本地开发

### 环境准备

- Windows
- .NET SDK 9.x
- 可选： [Inno Setup 6](https://jrsoftware.org/isinfo.php)

### 构建与测试

```powershell
dotnet restore WinTab.slnx
dotnet build WinTab.slnx -c Release
dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release
dotnet run --project src/WinTab.App/WinTab.App.csproj
```

### 本地打包

```powershell
# 1. 发布程序
dotnet publish src/WinTab.App/WinTab.App.csproj `
  -c Release -r win-x64 --self-contained false `
  -o publish/win-x64

# 2. 生成安装器（需要 Inno Setup 6）
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=1.0.0 installers/WinTab.iss

# 3. 生成便携版 ZIP
"WinTab Portable Mode" | Out-File publish/win-x64/portable.txt -Encoding utf8
Compress-Archive publish/win-x64\* publish/WinTab_1.0.0_portable.zip -Force
```

---

## 自动发布

- 推送到 `master`：执行构建、测试、打包，并上传 Actions Artifacts
- 推送版本标签（例如 `1.0.0`）：创建 GitHub Release，并附带安装版、便携版与校验文件

```powershell
git push origin master
git tag 1.0.0
git push origin 1.0.0
```

---

## FAQ

**Q：提示缺少 .NET 9 Desktop Runtime 怎么办？**  
A：安装器会引导下载；也可以手动访问 [dotnet 9 下载页](https://dotnet.microsoft.com/download/dotnet/9.0) 安装 Desktop Runtime。

**Q：Windows 10 能用吗？**  
A：可以运行，但部分依赖 Windows 11 Explorer 标签页环境的接管能力不可用。

**Q：为什么有时仍然会看到 Explorer 新窗口闪一下？**  
A：部分场景无法在请求发起前被直接接管，只能走“新窗口创建后自动转换为标签页”的兜底链路。当前版本已尽量把闪烁压低，但这类场景仍取决于 Explorer 自身行为。

**Q：卸载后配置会保留吗？**  
A：默认保留，只有在卸载时选择删除用户数据才会清空。

---

## 许可证

本项目基于 [MIT License](LICENSE) 发布。

---

## 致谢

- Explorer 标签页处理思路参考并移植自 [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)
- 安装器使用 Inno Setup
