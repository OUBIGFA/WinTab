package main

type copyCatalog struct {
	Navigation         string
	Settings           string
	RecoveryNav        string
	Preferences        string
	ExplorerIntro      string
	RecoveryIntro      string
	PreferencesIntro   string
	AboutIntro         string
	WindowsAndTabs     string
	MouseActions       string
	Appearance         string
	ThemeLabel         string
	LightChoice        string
	DarkChoice         string
	SystemChoice       string
	LanguageLabel      string
	Background         string
	ManualRecovery     string
	Automatic          string
	RecoveryDetails    string
	Connected          string
	ShortcutsNote      string
	Retry              string
	LoadTitle          string
	LoadErrorTitle     string
	AboutDescription   string
	Support            string
	Project            string
	Version            string
	LicenseLabel       string
	Repository         string
	UpdatesDesc        string
	LogsDesc           string
	Connecting         string
	Explorer           string
	Merge              string
	MergeDesc          string
	Reuse              string
	ReuseDesc          string
	EmptyDoubleClick   string
	EmptyDoubleClickD  string
	EmptyAction        string
	Parent             string
	Back               string
	DoubleClick        string
	DoubleClickDesc    string
	ExplorerOnly       string
	IncludeNotepad     string
	Scope              string
	MiddleClick        string
	MiddleClickDesc    string
	Wheel              string
	WheelDesc          string
	Sensitivity        string
	Low                string
	Medium             string
	High               string
	LowHint            string
	MediumHint         string
	HighHint           string
	Bypass             string
	BypassEnd          string
	Startup            string
	StartupDesc        string
	Tray               string
	TrayDesc           string
	TrayHidden         string
	AutoUpdate         string
	AutoUpdateDesc     string
	Recovery           string
	Group              string
	Tab                string
	Restore            string
	Restoring          string
	Shortcut           string
	GroupShortcut      string
	TabShortcut        string
	EnableGroup        string
	EnableTab          string
	Global             string
	InExplorer         string
	Single             string
	SingleDesc         string
	Record             string
	RecordRequired     string
	RecordHint         string
	Apply              string
	Applied            string
	Draft              string
	CaptureHint        string
	CaptureHelp        string
	AutoRestore        string
	AutoRestoreDesc    string
	NormalLaunch       string
	AnyFolder          string
	RestoreMode        string
	NormalHint         string
	AnyHint            string
	AutoIndependent    string
	RecoveryNotes      string
	About              string
	OpenSource         string
	License            string
	Logs               string
	Updates            string
	Checking           string
	Saving             string
	Running            string
	Light              string
	Dark               string
	Language           string
	LoadError          string
	RequestError       string
}

var zhCopy = copyCatalog{
	Navigation:         "设置分区",
	Settings:           "设置",
	RecoveryNav:        "会话恢复",
	Preferences:        "外观与系统",
	ExplorerIntro:      "让文件夹各归其位，让标签操作更顺手",
	RecoveryIntro:      "找回刚刚关闭的标签，继续上次的工作",
	PreferencesIntro:   "调整界面和后台运行方式",
	AboutIntro:         "版本信息、开源许可与运行支持",
	WindowsAndTabs:     "窗口与标签",
	MouseActions:       "鼠标操作",
	Appearance:         "外观",
	ThemeLabel:         "主题",
	LightChoice:        "浅色",
	DarkChoice:         "深色",
	SystemChoice:       "跟随系统",
	LanguageLabel:      "界面语言",
	Background:         "后台运行",
	ManualRecovery:     "手动恢复与快捷键",
	Automatic:          "自动恢复",
	RecoveryDetails:    "恢复范围与例外",
	Connected:          "正在后台运行",
	ShortcutsNote:      "快捷键修改后需要应用",
	Retry:              "重新连接",
	LoadTitle:          "连接 WinTab",
	LoadErrorTitle:     "暂时无法连接",
	AboutDescription:   "合并文件夹窗口，切换标签，找回上次的工作",
	Support:            "更新与支持",
	Project:            "查看源代码",
	Version:            "当前版本",
	LicenseLabel:       "开源协议",
	Repository:         "项目仓库",
	UpdatesDesc:        "手动检查 GitHub 上的最新稳定版本",
	LogsDesc:           "查看运行记录，协助排查问题",
	Connecting:         "正在连接资源管理器",
	Explorer:           "资源管理器",
	Merge:              "合并新窗口",
	MergeDesc:          "将新打开的文件夹合并为当前窗口的标签",
	Reuse:              "复用已有标签",
	ReuseDesc:          "打开同一文件夹时，切换到已有标签",
	EmptyDoubleClick:   "双击空白处导航",
	EmptyDoubleClickD:  "在当前标签的文件列表空白处双击",
	EmptyAction:        "双击操作",
	Parent:             "返回上一层",
	Back:               "后退",
	DoubleClick:        "双击关闭标签",
	DoubleClickDesc:    "双击标签页标题直接关闭",
	ExplorerOnly:       "仅资源管理器",
	IncludeNotepad:     "包含记事本",
	Scope:              "作用范围",
	MiddleClick:        "中键打开并切换",
	MiddleClickDesc:    "中键点按文件夹后，切换到新标签",
	Wheel:              "滚轮切换标签",
	WheelDesc:          "在标签栏滚动，向上切换到左侧，向下切换到右侧",
	Sensitivity:        "灵敏度",
	Low:                "低",
	Medium:             "中",
	High:               "高",
	LowHint:            "每两格切换一次",
	MediumHint:         "每格切换一次，快速连滚防误跳",
	HighHint:           "每格立即切换",
	Bypass:             "按住",
	BypassEnd:          "打开文件夹可保留独立窗口",
	Startup:            "开机启动",
	StartupDesc:        "登录 Windows 后在后台运行",
	Tray:               "显示托盘图标",
	TrayDesc:           "从系统托盘打开设置或退出 WinTab",
	TrayHidden:         "再次启动 WinTab 可打开设置",
	AutoUpdate:         "自动检查更新",
	AutoUpdateDesc:     "发现新版本时提醒你",
	Recovery:           "会话与标签恢复",
	Group:              "恢复上次标签组",
	Tab:                "恢复上一个关闭的标签",
	Restore:            "立即恢复",
	Restoring:          "正在恢复",
	Shortcut:           "快捷键",
	GroupShortcut:      "标签组恢复快捷键",
	TabShortcut:        "单标签恢复快捷键",
	EnableGroup:        "启用标签组快捷键",
	EnableTab:          "启用单标签快捷键",
	Global:             "全局可用",
	InExplorer:         "仅资源管理器",
	Single:             "包含单标签窗口",
	SingleDesc:         "默认仅保存和恢复多标签窗口，不影响单标签恢复",
	Record:             "记录最近关闭的标签",
	RecordRequired:     "使用恢复快捷键时需保持开启",
	RecordHint:         "关闭后清空记录",
	Apply:              "应用快捷键",
	Applied:            "快捷键已应用",
	Draft:              "有未应用的快捷键",
	CaptureHint:        "点按输入框，按下组合键或直接输入",
	CaptureHelp:        "使用 Ctrl 或 Alt 加字母、主键盘数字或 F1–F24，可加 Shift",
	AutoRestore:        "自动恢复上次标签组",
	AutoRestoreDesc:    "没有其他资源管理器窗口时生效",
	NormalLaunch:       "仅普通启动",
	AnyFolder:          "打开任意文件夹",
	RestoreMode:        "自动恢复方式",
	NormalHint:         "从主页或此电脑启动时自动恢复",
	AnyHint:            "打开任意文件夹时同时恢复历史标签",
	AutoIndependent:    "关闭后仍可用按钮或快捷键恢复",
	RecoveryNotes:      "拖出的窗口和 Ctrl + Shift 独立窗口不自动恢复；已启用 Windows 登录还原时，关机或异常退出前仍打开的窗口由系统还原",
	About:              "关于",
	OpenSource:         "在浏览器中打开 GitHub 项目",
	License:            "MIT 许可",
	Logs:               "打开日志",
	Updates:            "检查更新",
	Checking:           "正在检查",
	Saving:             "正在应用设置",
	Running:            "关闭设置后，WinTab 继续在后台运行",
	Light:              "切换为浅色主题",
	Dark:               "切换为深色主题",
	Language:           "Switch to English",
	LoadError:          "无法读取设置，请重新连接，或关闭此窗口后再打开",
	RequestError:       "操作未完成，请重试或查看日志",
}

var enCopy = copyCatalog{
	Navigation:         "Settings sections",
	Settings:           "Settings",
	RecoveryNav:        "Session recovery",
	Preferences:        "Appearance & system",
	ExplorerIntro:      "Keep folders together and move between tabs with ease",
	RecoveryIntro:      "Reopen closed tabs and pick up where you left off",
	PreferencesIntro:   "Make WinTab feel at home on your desktop",
	AboutIntro:         "Version details, license and support",
	WindowsAndTabs:     "Windows & tabs",
	MouseActions:       "Mouse actions",
	Appearance:         "Appearance",
	ThemeLabel:         "Theme",
	LightChoice:        "Light",
	DarkChoice:         "Dark",
	SystemChoice:       "System",
	LanguageLabel:      "Language",
	Background:         "Background behavior",
	ManualRecovery:     "Manual recovery & shortcuts",
	Automatic:          "Automatic recovery",
	RecoveryDetails:    "Recovery scope & exceptions",
	Connected:          "Running in the background",
	ShortcutsNote:      "Apply shortcut changes",
	Retry:              "Reconnect",
	LoadTitle:          "Connecting to WinTab",
	LoadErrorTitle:     "Connection unavailable",
	AboutDescription:   "Bring folders together, move between tabs and return to your work",
	Support:            "Updates & support",
	Project:            "View source code",
	Version:            "Current version",
	LicenseLabel:       "License",
	Repository:         "Repository",
	UpdatesDesc:        "Check GitHub Releases for a newer version",
	LogsDesc:           "View activity logs to help troubleshoot a problem",
	Connecting:         "Connecting to File Explorer",
	Explorer:           "File Explorer",
	Merge:              "Merge new windows",
	MergeDesc:          "Open folders as tabs in the current window",
	Reuse:              "Reuse existing tabs",
	ReuseDesc:          "Switch to the existing tab when you open the same folder",
	EmptyDoubleClick:   "Double-click empty space",
	EmptyDoubleClickD:  "Navigate within the current tab from the file list background",
	EmptyAction:        "Double-click action",
	Parent:             "Up one level",
	Back:               "Back",
	DoubleClick:        "Double-click to close tabs",
	DoubleClickDesc:    "Double-click any tab title to close it",
	ExplorerOnly:       "Explorer only",
	IncludeNotepad:     "Include Notepad",
	Scope:              "Applies to",
	MiddleClick:        "Switch to middle-clicked folders",
	MiddleClickDesc:    "Middle-click a folder to open and select its new tab",
	Wheel:              "Scroll to switch tabs",
	WheelDesc:          "On the tab bar, scroll up to go left or down to go right",
	Sensitivity:        "Sensitivity",
	Low:                "Low",
	Medium:             "Medium",
	High:               "High",
	LowHint:            "Switch every two notches",
	MediumHint:         "Switch one notch with fast-scroll guard",
	HighHint:           "Switch immediately with each notch",
	Bypass:             "Hold",
	BypassEnd:          "when opening a folder to keep a separate window",
	Startup:            "Start with Windows",
	StartupDesc:        "Run in the background when you sign in",
	Tray:               "Show tray icon",
	TrayDesc:           "Open settings or quit WinTab from the system tray",
	TrayHidden:         "Open WinTab again to access settings",
	AutoUpdate:         "Check for updates automatically",
	AutoUpdateDesc:     "Get notified when a new version is available",
	Recovery:           "Session & tab recovery",
	Group:              "Restore last tab group",
	Tab:                "Reopen last closed tab",
	Restore:            "Restore now",
	Restoring:          "Restoring",
	Shortcut:           "Shortcut",
	GroupShortcut:      "Tab group recovery shortcut",
	TabShortcut:        "Closed tab recovery shortcut",
	EnableGroup:        "Enable group shortcut",
	EnableTab:          "Enable closed tab shortcut",
	Global:             "Global",
	InExplorer:         "Explorer only",
	Single:             "Include single-tab windows",
	SingleDesc:         "Off by default; saves and restores multi-tab windows only",
	Record:             "Remember recently closed tabs",
	RecordRequired:     "Required while tab shortcut is enabled",
	RecordHint:         "Turning this off clears the history",
	Apply:              "Apply shortcuts",
	Applied:            "Shortcuts applied",
	Draft:              "Unapplied shortcut changes",
	CaptureHint:        "Click a field, then press a shortcut or type its name",
	CaptureHelp:        "Use Ctrl or Alt with a letter, top-row digit or F1–F24, with optional Shift",
	AutoRestore:        "Auto-restore last tab group",
	AutoRestoreDesc:    "When no other Explorer windows are open",
	NormalLaunch:       "Normal launch only",
	AnyFolder:          "Any folder",
	RestoreMode:        "Automatic recovery mode",
	NormalHint:         "Restore from Home or This PC only",
	AnyHint:            "Restore tabs while keeping folder open",
	AutoIndependent:    "Buttons and shortcuts stay available",
	RecoveryNotes:      "Detached tabs and Ctrl + Shift windows do not restore automatically; if Windows restores folders at sign-in, it handles windows left open at shutdown or a crash",
	About:              "About",
	OpenSource:         "Open the GitHub project in your browser",
	License:            "MIT License",
	Logs:               "Open logs",
	Updates:            "Check for updates",
	Checking:           "Checking",
	Saving:             "Applying settings",
	Running:            "WinTab keeps running when you close settings",
	Light:              "Switch to light theme",
	Dark:               "Switch to dark theme",
	Language:           "切换到中文",
	LoadError:          "Could not load settings, reconnect or close this window and open it again",
	RequestError:       "The action did not complete, try again or check the logs",
}

func copyForLanguage(lang string) copyCatalog {
	if lang == "en-US" {
		return enCopy
	}
	return zhCopy
}
