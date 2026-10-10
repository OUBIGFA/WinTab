package main

// These wire types mirror the resident engine's versioned DesktopState contract.
// Preferences are never written by Go: the C# SettingsManager remains their only owner.
type Settings struct {
	WindowHook                     bool   `json:"windowHook"`
	ReuseTabs                      bool   `json:"reuseTabs"`
	RestoreTabs                    bool   `json:"restoreTabs"`
	ReopenClosedTab                bool   `json:"reopenClosedTab"`
	RestoreGroupShortcutEnabled    bool   `json:"restoreGroupShortcutEnabled"`
	RestoreGroupShortcut           string `json:"restoreGroupShortcut"`
	ReopenTabShortcutEnabled       bool   `json:"reopenTabShortcutEnabled"`
	ReopenTabShortcut              string `json:"reopenTabShortcut"`
	RestoreOnAnyFolder             bool   `json:"restoreOnAnyFolder"`
	RestoreSingleTab               bool   `json:"restoreSingleTab"`
	DoubleClickCloseTab            bool   `json:"doubleClickCloseTab"`
	DoubleClickCloseIncludeNotepad bool   `json:"doubleClickCloseIncludeNotepad"`
	MiddleClickForegroundTab       bool   `json:"middleClickForegroundTab"`
	WheelSwitchTab                 bool   `json:"wheelSwitchTab"`
	WheelSwitchSensitivity         string `json:"wheelSwitchSensitivity"`
	AutoUpdate                     bool   `json:"autoUpdate"`
	ShowTrayIcon                   bool   `json:"showTrayIcon"`
	Language                       string `json:"language"`
	Theme                          string `json:"theme"`
	FormSize                       *Size  `json:"formSize"`
}

type Size struct {
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
}

type State struct {
	ProtocolVersion  int      `json:"protocolVersion"`
	Revision         int64    `json:"revision"`
	Version          string   `json:"version"`
	Settings         Settings `json:"settings"`
	Startup          bool     `json:"startup"`
	ShellReady       bool     `json:"shellReady"`
	RecordClosedTabs bool     `json:"recordClosedTabs"`
	SessionBusy      bool     `json:"sessionBusy"`
	UpdateBusy       bool     `json:"updateBusy"`
	HookError        *string  `json:"hookError"`
	ShortcutError    *string  `json:"shortcutError"`
	SessionFeedback  *string  `json:"sessionFeedback"`
	UpdateFeedback   *string  `json:"updateFeedback"`
	StorageError     *string  `json:"storageError"`
}

type WindowInfo struct {
	Width        int  `json:"width"`
	MaxWidth     int  `json:"maxWidth"`
	MaxHeight    int  `json:"maxHeight"`
	HasSavedSize bool `json:"hasSavedSize"`
}

type Bootstrap struct {
	State  State      `json:"state"`
	Window WindowInfo `json:"window"`
}
