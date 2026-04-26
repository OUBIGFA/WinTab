<div align="center">
  <img src="Assets/wintab-logo.png" width="96" alt="WinTab Logo">
  <h1>WinTab</h1>
  <p>Windows 11 资源管理器标签整理工具</p>
  <p>
    简体中文 | <a href="README.en.md">English</a>
  </p>
  <p>
    <img alt="Platform" src="https://img.shields.io/badge/platform-Windows%2011-2563eb">
    <img alt=".NET" src="https://img.shields.io/badge/.NET-9.0-512bd4">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-111827">
  </p>
</div>

## 简介

WinTab 是一个面向 Windows 11 的轻量工具，用来整理文件资源管理器的标签页体验。它会把新打开的文件夹窗口合并进已有的资源管理器窗口，并在目标路径已经打开时优先复用现有标签页。

WinTab 不会接管系统文件夹关联，也不会通过注册表重定向去替换 Shell 行为。它基于 Windows Shell COM 和窗口观察机制工作，因此第三方软件打开文件夹时仍然会落到正确的目标路径。

## 功能

- 自动将新打开的资源管理器窗口合并到现有标签页
- 目标路径已存在时优先复用已有标签页
- 保持第三方程序打开文件夹时的目标路径正确
- 双击标签标题即可关闭资源管理器标签页
- 打开文件夹时按住 `Ctrl + Shift` 可保留独立窗口
- 常驻系统托盘，提供轻量控制面板
- 支持中文和英文界面
- 支持浅色和深色主题
- 支持开机启动和 GitHub Release 更新检查

## 系统要求

- Windows 11 22H2 或更高版本
- .NET 9 Desktop Runtime
- 安装器会在运行时检测缺失的 .NET 9 Runtime，并提示下载安装

## 下载

只提供一种发布包：

- `WinTab_v1.0.0_Setup.exe`

后续版本也只发布 `.exe` 安装器，不再提供 zip 包、Chocolatey 包或 WinGet 包。

## 本地构建

### 环境要求

- Visual Studio 2022 或更高版本
- 工作负载：`.NET 桌面开发`
- .NET 9 SDK
- Inno Setup 6

### 生成发布文件

```powershell
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1

foreach ($arch in "x64", "x86", "arm64") {
  & $msbuild ".\WinTab\WinTab.csproj" `
    /restore `
    /t:Publish `
    /p:Configuration=Release `
    /p:TargetFramework=net9.0-windows `
    /p:RuntimeIdentifier="win-$arch" `
    /p:PublishDir="..\publish\net9.0-windows\$arch"
}
```

### 生成安装器

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DPublishRoot="..\publish\net9.0-windows" `
  /DOutputDir="..\dist" `
  ".\installers\installer.iss"
```

安装器输出目录：

```text
dist/WinTab_v1.0.0_Setup.exe
```

## GitHub Actions

仓库只保留一个发布工作流：

- `.github/workflows/build-release.yml`

该工作流会构建 `x64`、`x86` 和 `arm64` 的发布文件，打包成单一安装器，并只把 `WinTab_v*_Setup.exe` 上传到 GitHub Release。

## 项目结构

```text
WinTab/
├── .github/                 # GitHub 模板与发布工作流
├── Assets/                  # 仓库展示资源
├── installers/              # Inno Setup 安装器脚本
├── WinTab/                  # WPF 应用源码
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

本地构建输出目录 `publish/`、`dist/`、`bin/`、`obj/` 均不提交到仓库。

## 配置文件

安装后的配置文件位置：

```text
%APPDATA%\WinTab\settings.json
```

## 许可证

本项目使用 MIT License，详见 `LICENSE`。
