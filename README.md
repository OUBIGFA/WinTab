<div align="center">

  <img src="Assets/128x128.png" alt="WinTab Logo"/>

  <h1>WinTab</h1>

  <p>Windows 11 文件资源管理器标签页工具</p>

  <p>简体中文 | <a href="README.en.md">English</a></p>

  <p>
    <img src="https://img.shields.io/badge/platform-Windows%2011-555555" alt="Platform" />
    <img src="https://img.shields.io/badge/.NET-9.0-555555" alt=".NET" />
    <img src="https://img.shields.io/badge/license-MIT-555555" alt="License" />
  </p>
</div>

---

![](Assets/UI.png)

## 功能

- 合并新打开的资源管理器窗口
- 复用已打开路径的标签页
- 双击标签标题关闭当前标签页
- `Ctrl + Shift` 保留独立资源管理器窗口
- 托盘后台运行
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

## 开发与验证

此项目使用 `SHDocVw` / `Shell32` COM 引用，需要 Visual Studio MSBuild 构建；`dotnet build` 会因 `ResolveComReference` 限制失败。

```powershell
# 发布单个架构
.\build.ps1 -Arch x64 -SkipInstaller

# 构建解决方案和控制台自检项目
MSBuild.exe WinTab.sln /restore /t:Build
WinTab.Tests\bin\Debug\net9.0-windows\WinTab.Tests.exe
```

## 致谢

- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility)

## 许可

MIT License
