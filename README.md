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

根据 CPU 架构选择对应安装器（[Releases](https://github.com/OUBIGFA/WinTab/releases)）：

- `WinTab_v1.0.0_x64_Setup.exe` — 64 位 Intel/AMD（大多数用户）
- `WinTab_v1.0.0_arm64_Setup.exe` — ARM64 Windows（Surface Pro X 等）
- `WinTab_v1.0.0_x86_Setup.exe` — 32 位 Windows

## 系统要求

- Windows 11 22H2 或更高版本
- .NET 9 Desktop Runtime（首次运行安装器会自动下载安装）

## 本地构建

### 环境

- Visual Studio 2022 或更高版本（含 `.NET desktop development`）
- .NET 9 SDK
- Inno Setup 6

### 一键构建（推荐）

```powershell
.\build.ps1
```

脚本会发布 x64 / x86 / arm64 三个架构并编译三个独立安装器，输出到 `dist/`：

```text
dist/
├── WinTab_v1.0.0_x64_Setup.exe
├── WinTab_v1.0.0_x86_Setup.exe
└── WinTab_v1.0.0_arm64_Setup.exe
```

可选参数：

```powershell
.\build.ps1 -Version v1.1.0          # 指定版本号
.\build.ps1 -Arch x64                # 只构建一个架构
.\build.ps1 -Combined                # 额外生成自动检测架构的组合安装器
.\build.ps1 -SkipPublish             # 复用已有的 publish 输出
```

### 手动构建

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

  & "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe" `
    /DMyAppVersion="v1.0.0" `
    /DPublishRoot="..\publish\net9.0-windows" `
    /DOutputDir="..\dist" `
    /DArch="$arch" `
    ".\installers\installer.iss"
}
```

### 发布到 GitHub Releases

向仓库推送形如 `v*` 的标签即可触发 CI 自动构建签名版本并创建草稿发布：

```powershell
git tag v1.0.0
git push origin v1.0.0
```

## 项目结构

```text
WinTab/
├── .github/
├── Assets/
├── installers/
├── WinTab/
├── build.ps1
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
