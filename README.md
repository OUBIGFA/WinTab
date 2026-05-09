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

## 功能

- 合并新打开的资源管理器窗口
- 复用已打开路径的标签页
- 双击标签标题关闭当前标签页
- `Ctrl + Shift` 保留独立资源管理器窗口
- 系统托盘常驻
- 中英文界面
- 浅色 / 深色主题
- 开机启动
- GitHub Release 更新检查

## 下载

- `WinTab_v1.0.0_Setup.exe`

## 系统要求

- Windows 11 22H2 或更高版本
- .NET 9 Desktop Runtime

## 本地构建

### 环境

- Visual Studio 2022 或更高版本
- `.NET desktop development`
- .NET 9 SDK
- Inno Setup 6

### 发布

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

### 安装器

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" `
  /DMyAppVersion="v1.0.0" `
  /DPublishRoot="..\publish\net9.0-windows" `
  /DOutputDir="..\dist" `
  ".\installers\installer.iss"
```

```text
dist/WinTab_v1.0.0_Setup.exe
```

## 项目结构

```text
WinTab/
├── .github/
├── Assets/
├── installers/
├── WinTab/
├── LICENSE
├── README.md
├── README.en.md
└── WinTab.sln
```

## 配置文件

```text
%APPDATA%\WinTab\settings.json
```

## 致谢

- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)

## 许可

MIT License
