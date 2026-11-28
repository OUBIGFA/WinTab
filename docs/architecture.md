# WinTab 架构与修改边界

WinTab 由一个常驻 C# 宿主和一个按需启动的 Go 设置界面组成。设置、Explorer 操作和持久化由宿主负责。关闭设置界面不会停止标签页功能。x86 使用同一宿主内的 WPF 兼容界面。

## 模块所有权

| 范围 | 入口与职责 | 修改边界 |
| --- | --- | --- |
| 启动与退出 | `WinTab/App.xaml.cs`：单实例、启动请求、UI 生命周期 | 不在界面进程中启动第二组钩子 |
| 功能开关 | `Managers/HookManager.cs`：应用开关、合并与复用的联动、快捷键 | Go、WPF 和托盘通过同一入口修改功能状态 |
| 设置 | `AppSettings` 定义值与归一规则；`SettingsStore` 负责快照、异步保存和备份；`SettingsManager` 发布变化 | 新增设置时先确定默认值和旧配置行为，不在每个界面各写一套默认值 |
| Explorer 协调 | `Hooks/ExplorerWatcher.*`：注册、导航、合并、会话与资源回收 | partial 文件共享线程、窗口身份和生命周期，不是独立模块 |
| 交互规则 | `ExplorerTabDoubleClickCloseController`、`ExplorerTabCloseQueue`、`ExplorerTabWheelSwitchController`、`TabSelectionEngine` | 规则与输入观察、UIA 读取分开；钩子回调不等待磁盘或远端 UIA |
| 会话规则 | `ExplorerSessionTracker`、`ExplorerSessionRestorePlan`、`ExplorerSessionLocationPolicy`、`ExplorerClosedTabHistory`、`ExplorerSessionRestoreAttempt` | 接收快照或观察结果，不持有 Explorer COM 连接；用直接调用测试规则 |
| 原生边界 | `WinAPI/`、`Interop/`、`Helpers/WindowIdentity`、窗口发现和可见性辅助类 | 操作前校验窗口身份；窗口句柄本身不能证明窗口仍属于原操作 |
| 设置界面通信 | `UI/Desktop/DesktopCommand` 校验命令；`DesktopLineReader` 读取有界行；`DesktopUiProcess` 管理子进程和收发；`DesktopApplication` 执行业务命令 | 协议处理不直接实现 Explorer 规则；界面退出不取消宿主已接收的恢复任务 |
| Go 界面 | `main.go` 管理窗口和宿主状态；`bridge.go` 管理请求；`ui_model.go` 管理状态、草稿和交互；`ui_view.go` 管理页面框架和公共控件；`ui_pages.go` 管理功能页 | 渲染状态只在 MyGo 主线程修改；共享指针和宿主快照通过 `Desktop.mu` 发布 |
| 文案与主题 | Go 的 `copy.go`、`theme.go`；WPF/托盘的 `UiStrings`、`ThemeManager` | 两种界面采用相同设置值；中文和英文文案不加句末标点 |
| 安装与更新 | `build.ps1`、`WinTab.Desktop/build.ps1`、`installers/`、`UpdateReleaseParser` | 架构选择、运行时检测和校验必须与实际安装包一致 |

## 命令与快照

原生界面的修改顺序为 `uiModel → Desktop → bridge → DesktopUiProcess → DesktopCommand → DesktopApplication → HookManager/SettingsManager`。宿主生成完整 `DesktopState`，再由请求响应或推送返回界面。

- `ProtocolVersion` 标识协议；`Revision` 决定快照新旧。晚到的旧响应不能覆盖新推送，窗口初次创建也遵守这个规则
- C# 和 Go 的契约分别在 `DesktopState`、`state.go`；共享夹具为 `WinTab.Desktop/testdata/state.json`，两侧测试均读取它
- 命令按接收顺序在宿主 UI 线程开始执行。长时间的恢复和更新按请求编号独立回复，不阻塞后续设置
- C# 按块读取管道并保留下一行的数据；协议行和诊断行均有长度上限。协议不接受未结束的半行，诊断允许 EOF 前的最后一行
- Go 在发送前检查取消状态，包括等待上一条写入之后。已经送达宿主的命令不会因为界面取消等待而回滚
- 配置读取和更新时统一归一主题、语言、滚轮灵敏度及空白区域动作。快捷键原文保留，由快捷键校验器报告错误

## 线程与资源

| 执行位置 | 所有权与约束 |
| --- | --- |
| WPF Dispatcher | 设置命令、宿主 UI 和托盘；耗时工作返回任务，不同步等待外部服务 |
| `StaTaskScheduler` | Explorer COM 工作和需要消息循环的操作；重连与退出先停止旧工作，再释放连接 |
| 输入观察线程 | 采集输入和窗口事件；计算所需的本地事实后交给控制器或后台任务 |
| `BackgroundRefreshCache` | 两个 MTA 工作线程、有限队列；UIA 快照过期后重新计算，不按每次鼠标移动创建线程 |
| `SettingsStore` / `ExplorerSessionStore` | 单个保存工作者合并待写快照；失败可见，无法读取的原文件不会被默认值覆盖 |
| MyGo 主线程 | 创建窗口、修改 `uiModel` 和调用原生主题 API；后台回复通过窗口更新队列进入 |

会话恢复的取消规则由 `ExplorerSessionRestoreAttempt` 独立管理：只有目标窗口曾经获得前台后，离开前台才取消；初始导航允许完成，开始观察之后的用户导航才取消；退出后的迟到回调不再访问已释放资源。

## 并行修改约定

1. 按模块和测试文件分配任务，先声明负责的文件。不同功能页需要同时开发时，再按页面把 `ui_pages.go` 中相应函数拆出
2. `ExplorerWatcher.*` 的字段、锁、取消源和 COM 生命周期由一个任务统筹。不能只因为修改不同 partial 文件就认定互不影响
3. 变更设置或通信字段时，由一个任务同步 C#、Go、共享夹具和默认值测试；其他任务基于确定的契约实现
4. 新的规则优先放到现有控制器或策略类。需要独立所有权时再抽出类，不按文件行数建立转发层
5. 保留现有测试覆盖的行为：自动恢复默认关闭、记录与手动恢复独立、Ctrl+Shift 独立窗口、拖出标签页、控制面板排除、窗口恢复兜底
6. 临时工作放 `_temp/<任务>/`，安装包放 `dist/`；合并前检查差异和跟踪清单，不提交日志、截图、机器路径或工具缓存

## 验证入口

最终交付从仓库根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

该入口构建并运行 C# 测试，发布 x64/x86/arm64 宿主，检查并测试 Go 界面，构建 x64/arm64 界面，执行四种安装器运行时配置测试，最后覆盖 `dist/` 中的对应安装器。COM 引用需要 Visual Studio 的 MSBuild，不能用 `dotnet build` 替代。

迭代时只运行受影响的用例。C# 先重新构建测试项目，再使用 `WinTab.Tests.exe --filter '关键词一|关键词二'`；筛选无匹配时返回失败。测试入口在 `WinTab.Tests/Program.cs`，部分套件会继续枚举子套件，例如双击关闭测试包含关闭队列测试，避免重复注册。

Go 在 `WinTab.Desktop/` 中使用与发布相同的标签：

```powershell
go test -tags 'mygo_nowebview,mygo_noinspector,mygo_nodownload' -run 'TestBridge|TestDesktopInitialModel' .
```

优先使用纯规则、真实管道、消息专用窗口和 MyGo 无窗口测试。需要创建可显示窗口或运行安装器的验证放在私有桌面，不能操作用户的 Explorer。构建与这些自动测试不等于在每个架构上实际安装或完成交互验收。

## 后续演进

`ExplorerWatcher` 仍是主要复杂度集中点。下次修改某项功能时，沿“纯规则 → 快照 → 原生动作”继续分离，并用现有回归用例保护窗口身份、线程和恢复语义。当前不把所有 partial 文件重新分层，也不引入新的容器、消息总线或插件框架。

性能结论区分代码路径和实测体验：按块读取减少了每字符一次的异步读取；有界队列和行长限制约束积压。CPU、常驻内存、启动时间和 Explorer 操作延迟仍需独立基准，不能从文件变短或安装包变小直接推断。
