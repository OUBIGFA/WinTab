<div align="center">
  <img src="WinTab/wintab-logo.png" width="96" alt="WinTab Logo">
  <h1>WinTab</h1>
  <p>Windows 11 资源管理器标签管理工具</p>
  <p>
    <a href="README.en.md">English</a> | 简体中文
  </p>
  <p>
    <img alt="Platform" src="https://img.shields.io/badge/platform-Windows%2011-2563eb">
    <img alt=".NET" src="https://img.shields.io/badge/.NET-9.0%20%7C%204.8.1-512bd4">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-111827">
  </p>
</div>

## 简介

WinTab 是一个专注于 Windows 11 资源管理器标签页体验的小工具。它会把新打开的文件夹优先并入现有资源管理器窗口，并在目标路径已经打开时复用已有标签页，减少重复窗口和重复标签。

WinTab 不接管系统文件夹打开方式，不使用注册表替换 Shell 行为。它通过 Windows Shell COM 和窗口观察实现资源管理器标签管理，因此第三方软件打开文件夹时仍然会进入正确的目标路径。

## 功能特性

- 自动合并新打开的资源管理器文件夹窗口到现有标签页。
- 当相同路径已经打开时，优先复用已有标签页。
- 第三方软件或程序打开文件夹时保持目标路径正确。
- 双击资源管理器标签页标题关闭该标签页，空白标题栏区域保持系统原生行为。
- 打开文件夹时按住 `Ctrl + Shift` 可保留为独立资源管理器窗口。
- 托盘优先运行，提供精简控制面板。
- 支持中文和英文界面切换。
- 支持浅色和深色主题。
- 支持开机启动和 GitHub Release 更新检查。

## 系统要求

- Windows 11 22H2 或更高版本。
- 推荐安装 .NET 9 Desktop Runtime。
- 安装器会根据系统架构选择 x64、x86 或 ARM64 包。
- 如果系统没有 .NET 9 Desktop Runtime，安装器会尝试使用内置的 .NET Framework 4.8.1 版本，或提示安装运行时。

## 下载与使用

1. 从 GitHub Releases 下载 `WinTab_v*_Setup.exe`。
2. 运行安装器并完成安装。
3. 启动 WinTab 后，它会驻留在系统托盘。
4. 正常打开文件夹即可自动进入资源管理器标签页流程。
5. 需要临时打开独立资源管理器窗口时，按住 `Ctrl + Shift` 再打开文件夹。

安装后的配置文件位于：

```text
%APPDATA%\WinTab\settings.json
```

## 本地构建

### 必需组件

- Visual Studio 2022 或更新版本。
- 工作负载：`.NET 桌面开发`。
- .NET 9 SDK。
- 构建 .NET Framework 版本时需要 .NET Framework 4.8.1 Developer Pack。
- 打包安装器时需要 Inno Setup 6。

### 编译应用

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1

& $msbuild ".\WinTab\WinTab.csproj" `
  /restore `
  /t:Publish `
  /p:Configuration=Release `
  /p:TargetFramework=net9.0-windows `
  /p:RuntimeIdentifier=win-x64 `
  /p:PublishDir="..\publish\net9.0-windows\x64"
```

### 构建安装器

安装器脚本位于 `installers/installer.iss`。它要求 `artifacts` 目录中存在对应架构和框架的发布 ZIP 包。

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DSourceDir="..\artifacts" `
  ".\installers\installer.iss"
```

生成的安装器会输出到 `artifacts` 目录。

## GitHub Actions

仓库内置发布工作流：

- `.github/workflows/build-release.yml`：构建 x64、x86、ARM64 的 .NET 9 和 .NET Framework 4.8.1 包，生成并签名安装器，然后创建 Release。
- `.github/workflows/publish-winget.yml`：发布到 WinGet。
- `.github/workflows/publish-chocolatey.yml`：发布到 Chocolatey。

## 项目结构

```text
WinTab/
├── .github/              # Issue 模板和发布工作流
├── artifacts/            # 本地构建输出和安装器
├── installers/           # Inno Setup 安装器脚本与语言文件
├── tests/                # 资源管理器行为验证脚本
├── WinTab/               # WPF 应用源码
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

## 常见问题

### 为什么打开文件夹没有合并到标签页？

请确认 WinTab 正在运行，并且控制面板中的自动合并开关已开启。如果打开文件夹时按住了 `Ctrl + Shift`，WinTab 会保留独立资源管理器窗口。

### 第三方软件打开文件夹会不会被影响？

WinTab 会等待资源管理器窗口进入真实目标路径后再执行合并或复用，因此第三方软件打开文件夹时应保持目标路径正确。

### 双击哪里可以关闭标签页？

仅双击资源管理器标签页标题区域会关闭标签页。双击空白标题栏、窗口边框或其他非标签标题区域时，WinTab 会保留 Windows 原生窗口行为。

### 如何彻底退出 WinTab？

从系统托盘图标打开菜单并选择退出。关闭控制面板窗口不会退出 WinTab。

## 致谢

WinTab 的资源管理器标签管理核心基于 ExplorerTabUtility 的 Shell/COM 方案演进而来，并将产品范围收敛到资源管理器标签合并、复用和关闭体验。

## 许可证

本项目使用 MIT License。详见 `LICENSE`。
