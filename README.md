<div align="center">

<img src="src/WinTab.App/Assets/wintab.ico" width="80" />

# WinTab

**为 Windows 资源管理器添加标签页的开源工具**

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)]()
[![.NET 9](https://img.shields.io/badge/.NET-9.0-512BD4.svg)]()

[English](README.en.md) | 简体中文

</div>

---

## 简介

WinTab 是一款开源的 Windows 窗口标签管理工具。它通过拦截系统的文件夹打开事件，让你在一个窗口里以多个**标签页**的方式管理资源管理器——就像浏览器一样。

得益于完全基于 Win32 事件推送的架构设计，**长期后台驻留时 CPU 占用趋近于 0%**，内存常驻开销仅 30–80 MB。

---

## 功能特性

### 核心功能

- **标签页接管**：拦截由系统或其他应用发出的文件夹打开请求，自动以新标签页的方式并入当前 WinTab 窗口，无需新弹窗口
- **继承当前路径**：新建标签页时默认沿用当前目录，而非默认的"此电脑"
- **在新标签页中打开子文件夹**：点击目录时自动弹出新标签页，保留当前视图不被覆盖
- **双击标签页标题关闭**：不用瞄准小叉，对着标签头双击即可关闭，相当于鼠标中键

### 系统集成

- **托盘驻留**：开机自启动 + 启动时隐藏主窗口；支持随时从托盘恢复
- **显示托盘图标**：可独立控制是否在系统通知区域展示图标
- **主题**：支持浅色与深色两种界面主题
- **双语界面**：简体中文 / English，可在设置页随时切换，即时生效

### 稳定性

- 崩溃日志与运行日志自动记录
- 退出或卸载时自动还原注册表修改，不留残留
- 内置注册表状态自检机制，防止环境被污染

---

## 系统要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（仅 x64） |
| 运行时 | [.NET 9 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/9.0) |
| 推荐系统 | Windows 11（标签页拦截功能完整支持） |

> **注意**：资源管理器的文件夹标签接管功能需要 Windows 11（`10.0.22000+`）。Windows 10 下可正常运行，但此功能不可用。

---

## 下载与安装

前往 [Releases](../../releases) 页面下载最新版本。

### 安装版（推荐普通用户）

下载 `WinTab_Setup_<version>.exe` 双击安装。

- 支持自定义安装路径
- 安装向导会自动检测是否缺少 .NET 9 运行时并引导下载
- 默认跟随系统语言（中/英）
- 安装时可勾选**开机自启动**

#### 覆盖安装

再次运行安装包时，向导将提供两种模式：

| 模式 | 说明 |
|---|---|
| **卸载完再安装（推荐）** | 先完整卸载旧版本，可选是否同时删除用户数据 |
| **不卸载直接安装** | 直接覆盖程序文件，完整保留用户设置 |

静默安装参数：

```
WinTab_Setup_<version>.exe /SILENT /REINSTALLMODE=CLEAN /REMOVEUSERDATA=1
```

| 参数 | 说明 |
|---|---|
| `/REINSTALLMODE=CLEAN` | 先卸载再安装 |
| `/REINSTALLMODE=DIRECT` | 直接覆盖（静默模式默认值） |
| `/REMOVEUSERDATA=1` | 仅 CLEAN 模式有效，删除 `%AppData%\Roaming\WinTab` |

---

### 便携版（免安装）

下载 `WinTab_<version>_portable.zip`，解压后直接运行 `WinTab.exe`。

解压包中包含 `portable.txt` 标记文件，程序会自动进入便携模式：

- 配置与日志写入同级的 `data/` 目录，不使用 `%AppData%`
- 不再使用时直接删除整个解压文件夹即可

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
| 配置文件 | `<解压目录>\data\settings.json` |
| 运行日志 | `<解压目录>\data\logs\wintab.log` |
| 崩溃日志 | `<解压目录>\data\logs\crash.log` |

### 配置文件示例

```json
{
  "StartMinimized": false,
  "ShowTrayIcon": true,
  "RunAtStartup": false,
  "Language": "Chinese",
  "Theme": "Light",
  "EnableExplorerOpenVerbInterception": true,
  "OpenChildFolderInNewTabFromActiveTab": false,
  "CloseTabOnDoubleClick": true,
  "SchemaVersion": 1
}
```

---

## 资源占用说明

WinTab 专为长期后台驻留设计，所有监听机制均采用操作系统事件推送，**不进行主动轮询**。

| 资源 | 空闲时 | 活跃使用时 |
|---|---|---|
| CPU | ≈ 0% | 窗口事件处理时短暂唤醒 |
| 内存 | 30–80 MB | 无显著增长 |
| 磁盘 | 无后台读写 | 仅在产生日志或设置变更时写入 |
| 网络 | 无 | 无 |

**内部机制**：

- `SetWinEventHook`（Win32）——系统窗口事件推送，被动响应，CPU 友好
- `NamedPipeServerStream.WaitForConnectionAsync`——异步 `await` 阻塞等待，管道无连接时线程不消耗 CPU
- `EventWaitHandle.WaitOne()`——操作系统级阻塞，用于进程间单实例激活信号
- `RegistryOpenVerbInterceptor`——仅启动时执行一次自检，后续无任何周期性注册表操作

---

## 卸载

在应用内进入**卸载**页面，可选择：

- **保留用户数据**（默认）：仅删除程序文件，保留 `%AppData%\WinTab`
- **彻底清理**：勾选"删除用户数据"，同时清除配置与日志

卸载过程会自动还原注册表中的 `Folder/Directory/Drive` 打开命令至原始状态。

---

## 工作机制

```
系统/其他应用打开文件夹
        │
        ▼
RegistryOpenVerbInterceptor 拦截 open-verb
        │
        ▼
ExplorerOpenVerbHandler 进程启动（--wintab-open-folder）
        │
        ▼
NamedPipe → ExplorerOpenRequestServer（主进程）
        │
        ▼
WindowManager 查找/激活已有 Explorer 窗口
        │
        ▼
标签页在现有窗口中打开，无需新建窗口
```

---

## 项目结构

```
src/
  WinTab.App/             WPF 主程序（DI、页面、服务、托盘、启动流程）
  WinTab.UI/              UI 资源与本地化（中/英）
  WinTab.Core/            核心模型与接口
  WinTab.Platform.Win32/  Win32 互操作与窗口操作
  WinTab.Persistence/     配置与路径管理
  WinTab.Diagnostics/     日志与崩溃处理
  WinTab.Tests/           单元测试
installers/
  WinTab.iss              Inno Setup 安装包脚本
.github/workflows/
  build-release.yml       CI/CD（构建 → 测试 → 打包 → 发布）
```

---

## 本地开发

### 环境准备

- .NET SDK 9.x
- Windows（建议 PowerShell）
- 可选：[Inno Setup 6](https://jrsoftware.org/isinfo.php)（打安装包时需要）

### 构建与测试

```powershell
dotnet restore WinTab.slnx
dotnet build WinTab.slnx -c Release
dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release
dotnet run --project src/WinTab.App/WinTab.App.csproj
```

### 本地打包

```powershell
# 1. 发布应用
dotnet publish src/WinTab.App/WinTab.App.csproj `
  -c Release -r win-x64 --self-contained false `
  -o publish/win-x64

# 2. 生成安装包（需已安装 Inno Setup 6）
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=1.0.0 installers/WinTab.iss

# 3. 生成便携版 ZIP
"WinTab Portable Mode" | Out-File publish/win-x64/portable.txt -Encoding utf8
Compress-Archive publish/win-x64\* publish/WinTab_1.0.0_portable.zip -Force
```

---

## 自动发布（GitHub Actions）

- **推送到 `master`**：自动构建、测试、打包并上传 Actions Artifacts
- **推送 tag**（如 `1.0.0`）：自动创建 GitHub Release，上传安装包与便携版及对应的 SHA256 校验文件

```powershell
git push origin master
git tag 1.0.0
git push origin 1.0.0
```

---

## 常见问题

**Q：提示需要 .NET 9 运行时怎么办？**
A：安装包会自动引导下载页面。也可手动前往 https://dotnet.microsoft.com/download/dotnet/9.0 下载"Desktop Runtime"。

**Q：卸载后配置会保留吗？**
A：默认保留，除非卸载时勾选"删除用户数据"选项。

**Q：Windows 10 能用吗？**
A：可以运行，但"接管外部文件夹打开请求"功能依赖 Windows 11 的资源管理器标签页接口，在 Windows 10 上无效。


---

## 许可证

本项目基于 [MIT 许可证](LICENSE) 发布。

---

## 致谢

- Explorer 标签页处理逻辑参考并移植自 [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)（MIT）
- Inno Setup 简体中文语言文件来自其翻译生态，存放于 `installers/Languages/ChineseSimplified.isl`
