<div>![WinTab Logo](Assets/wintab-logo.png)   

# WinTab

Windows 11 资源管理器标签整理工具

     简体中文 | [English](README.en.md)   

     ![Platform](https://img.shields.io/badge/platform-Windows%2011-2563eb)     ![.NET](https://img.shields.io/badge/.NET-9.0-512bd4)     ![License](https://img.shields.io/badge/license-MIT-111827)   

</div>

![](Assets/UI.png)

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

