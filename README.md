# WinTab

WinTab 是一款开源的 Windows 窗口标签工具，让资源管理器拥有类似浏览器的标签页体验。项目基于 .NET 9（WPF）实现，支持安装版与便携版两种分发方式。

## 功能概览

- 资源管理器目录打开请求转发到 WinTab 进程，尽量在已有窗口中以标签页方式打开。
- 托盘运行：支持开机自启动、启动最小化到托盘、双击托盘图标恢复主界面。
- 双语界面：简体中文 / English，可在设置页切换。
- 主题切换：浅色 / 深色。
- 可选行为开关：
  - 接管资源管理器 `open` 行为（open-verb 拦截）。
  - 在当前活动标签页目录下打开子目录时，改为新建标签页。
  - 双击资源管理器标签页标题区域时关闭标签页（模拟鼠标中键单击，默认关闭）。
- 稳定性与清理：
  - 崩溃日志与运行日志。
  - 退出或卸载时尽量恢复注册表覆盖，减少残留。

## 系统要求

- Windows 10 / 11（x64）。
- `.NET 9 Desktop Runtime`（安装包会检测并提示）。
- 推荐 Windows 11 以获得完整的资源管理器标签页相关能力。

> 说明：open-verb 拦截在代码中要求 Windows 11（`10.0.22000+`）。

## 安装与使用

### 1) 安装版（推荐普通用户）

从 Releases 下载：`WinTab_Setup_<version>.exe`。

- 安装向导支持自定义安装路径。
- 默认语言会按系统语言自动检测（可在语言选择页手动改中/英）。
- 可勾选开机自启动。

#### 已安装后再次运行安装包

若检测到已安装版本，安装页会提供两种模式：

- `卸载完再安装（推荐）`：可额外勾选“删除用户数据”，默认不勾选。
- `不卸载直接安装`：直接覆盖安装程序文件，保留用户数据。

静默安装可通过参数控制重装模式：

- `/REINSTALLMODE=CLEAN`：先卸载再安装。
- `/REINSTALLMODE=DIRECT`：直接安装（静默模式默认）。
- `/REMOVEUSERDATA=1`：仅在 `CLEAN` 模式下生效，用于删除 `%AppData%\Roaming\WinTab`。

### 2) 便携版（免安装）

从 Releases 下载：`WinTab_<version>_portable.zip`，解压后直接运行 `WinTab.exe`。

便携版包含 `portable.txt` 标记文件，程序会自动进入便携模式：

- 配置、日志写入程序目录下的 `data/`。
- 不使用 `%AppData%\WinTab`。
- 不想用时直接删除整个解压目录即可。

## 数据目录与配置文件

### 安装版

- 配置：`%AppData%\WinTab\settings.json`
- 日志：`%AppData%\WinTab\logs\wintab.log`
- 崩溃日志：`%AppData%\WinTab\logs\crash.log`

### 便携版

- 配置：`<解压目录>\data\settings.json`
- 日志：`<解压目录>\data\logs\wintab.log`
- 崩溃日志：`<解压目录>\data\logs\crash.log`

### `settings.json` 示例

```json
{
  "StartMinimized": false,
  "RunAtStartup": false,
  "Language": "Chinese",
  "Theme": "Light",
  "EnableExplorerOpenVerbInterception": true,
  "OpenChildFolderInNewTabFromActiveTab": false,
  "CloseTabOnDoubleClick": false,
  "SchemaVersion": 1
}
```

## 卸载行为

卸载时默认策略：**保留用户配置与日志**。

- 在应用内“卸载”页面可勾选“删除用户数据”来决定是否删除 `%AppData%\Roaming\WinTab`。
- 复选框默认不勾选（保留数据）。
- 勾选后执行彻底清理（含用户数据目录）。

此外，卸载流程会尝试执行应用内清理命令以恢复 open-verb 相关注册表状态。

## 工作机制（简版）

WinTab 主要由以下能力组成：

- Win32 窗口事件监听（`SetWinEventHook`）。
- 资源管理器窗口/标签识别与切换（Win32 + COM 动态调用）。
- 命名管道 IPC（`WinTab_ExplorerOpenRequest`）用于将二次打开请求转发给主实例。
- 注册表拦截器：在 HKCU 下写入并恢复 `Folder/Directory/Drive` 的 `shell\\open\\command`。

> 高级开关：`WINTAB_AUTO_CONVERT_EXPLORER=1` 时会启用更积极的自动转换路径（默认关闭）。

## 内部命令行参数（调试/维护）

- `--wintab-open-folder <path>`：资源管理器 open-verb 入口。
- `--open-folder <path>`：历史兼容入口。
- `--wintab-companion <pid>`：伴生进程模式（监测主进程并兜底恢复）。
- `--wintab-cleanup`：卸载清理模式（用于恢复注册表/启动项）。

## 项目结构

```text
src/
  WinTab.App/             WPF 主程序（DI、页面、服务、托盘、启动流程）
  WinTab.UI/              UI 资源与本地化（中/英）
  WinTab.Core/            核心模型与接口
  WinTab.Platform.Win32/  Win32 互操作与窗口操作
  WinTab.Persistence/     配置与路径管理
  WinTab.Diagnostics/     日志与崩溃处理
  WinTab.Tests/           单元测试（当前主要覆盖 SettingsStore）
installers/
  WinTab.iss              Inno Setup 脚本
  Languages/ChineseSimplified.isl
.github/workflows/
  build-release.yml       CI/CD（构建、测试、打包、发布）
```

## 本地开发

### 环境准备

- .NET SDK 9.x
- Windows（建议 PowerShell）
- 可选：Inno Setup 6（需要打安装包时）

### 运行与测试

```powershell
dotnet restore WinTab.slnx
dotnet build WinTab.slnx -c Release
dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release
dotnet run --project src/WinTab.App/WinTab.App.csproj
```

### 本地打包（安装版 + 便携版）

```powershell
# 1) 发布程序
dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64

# 2) 去掉调试符号（发布产物更干净）
Get-ChildItem -Path publish/win-x64 -Filter *.pdb -Recurse | Remove-Item -Force

# 3) 生成安装包（需已安装 Inno Setup 6）
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=1.0.0 installers/WinTab.iss

# 4) 生成便携版 ZIP（含 portable 模式标记）
"WinTab Portable Mode - Data is stored in the 'data' folder next to this executable" |
  Out-File -FilePath publish/win-x64/portable.txt -Encoding utf8
Compress-Archive -Path publish/win-x64\* -DestinationPath publish/WinTab_1.0.0_portable.zip -Force
```

## GitHub Actions 自动打包 / 发布

工作流文件：`.github/workflows/build-release.yml`

- 推送到 `master`：自动构建、测试、打包并上传 Actions Artifacts。
- 推送 tag（如 `1.0.0`）：在上面基础上自动创建 GitHub Release 并上传：
  - `WinTab_Setup_<version>.exe`
  - `WinTab_Setup_<version>.exe.sha256`
  - `WinTab_<version>_portable.zip`
  - `WinTab_<version>_portable.zip.sha256`

发布建议流程：

```powershell
git push origin master
git tag 1.0.0
git push origin 1.0.0
```

> 注意：触发 Release 的 tag 对应提交里必须已经包含工作流文件。

## 常见问题

### Q1: 为什么没有自动发布到 Releases？

请依次检查：

- 仓库 `Actions` 是否启用。
- 是否推送了符合 `*.*.*` 规则的 tag（例如 `1.0.0`）。
- tag 指向的提交是否已包含 `.github/workflows/build-release.yml`。
- Workflow 是否有写入 Release 的权限（`contents: write`）。

### Q2: 卸载后还会残留配置吗？

默认会保留。卸载时选“是”可删除 `%AppData%\Roaming\WinTab`。

### Q3: Windows 10 能用吗？

可运行，但部分资源管理器标签页相关能力依赖 Windows 11。

## 许可证

本项目基于 `MIT` 许可证发布。详见 `LICENSE`。

## 致谢

- Explorer tab 处理逻辑参考并移植自 `ExplorerTabUtility`（MIT）。
- Inno Setup 简体中文语言文件来自其翻译生态（存放于 `installers/Languages/ChineseSimplified.isl`）。
