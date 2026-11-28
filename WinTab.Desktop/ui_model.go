package main

import (
	"github.com/egoist/mygo/ui"
	"strings"
)

// UI state and actions belong to the main thread; this file does not create a native window.
type uiActions interface {
	Set(key string, value any)
	Shortcuts(groupEnabled bool, group string, tabEnabled bool, tab string, notifyApplied bool)
	Restore(group bool)
	CheckUpdates()
	OpenLogs()
	OpenProject()
	Reconnect()
}

type navEntry struct {
	id    string
	label string
	intro string
	icon  *ui.SVG
}

type uiModel struct {
	state           *State
	section         string
	loadError       bool
	actionError     bool
	pending         map[string]bool
	groupDraft      string
	tabDraft        string
	syncedGroup     string
	syncedTab       string
	hasSynced       bool
	shortcutApplied bool
	detailsOpen     bool
	groupCapture    shortcutCapture
	tabCapture      shortcutCapture
	heading         ui.Handle
	focusHeading    bool
	workspaceScroll ui.ScrollState
	systemDark      bool // the desktop's appearance at the start of the current frame
	actions         uiActions
}

func newUIModel(initial *State, actions uiActions) *uiModel {
	m := &uiModel{
		section: "explorer",
		pending: make(map[string]bool),
		actions: actions,
	}
	if initial != nil {
		m.applyState(*initial)
	}
	return m
}

func (m *uiModel) applyState(next State) {
	if m.state != nil && next.Revision < m.state.Revision {
		return
	}
	copied := next
	m.state = &copied
	m.loadError = false
	if !m.hasSynced || next.Settings.RestoreGroupShortcut != m.syncedGroup {
		m.groupDraft = next.Settings.RestoreGroupShortcut
		m.syncedGroup = next.Settings.RestoreGroupShortcut
	}
	if !m.hasSynced || next.Settings.ReopenTabShortcut != m.syncedTab {
		m.tabDraft = next.Settings.ReopenTabShortcut
		m.syncedTab = next.Settings.ReopenTabShortcut
	}
	m.hasSynced = true
}

func (m *uiModel) language() string {
	if m.state != nil && m.state.Settings.Language != "" {
		return m.state.Settings.Language
	}
	return "zh-CN"
}

func (m *uiModel) dark() bool {
	if m.state == nil {
		return false
	}
	switch m.state.Settings.Theme {
	case "Dark":
		return true
	case "System":
		return m.systemDark
	}
	return false
}

func (m *uiModel) copy() copyCatalog {
	return copyForLanguage(m.language())
}

func (m *uiModel) sections(cp copyCatalog) []navEntry {
	return []navEntry{
		{id: "explorer", label: cp.Explorer, intro: cp.ExplorerIntro, icon: iconFolderOpen},
		{id: "recovery", label: cp.RecoveryNav, intro: cp.RecoveryIntro, icon: iconHistory},
		{id: "preferences", label: cp.Preferences, intro: cp.PreferencesIntro, icon: iconTuning2},
		{id: "about", label: cp.About, intro: cp.AboutIntro, icon: iconInfoCircle},
	}
}

func (m *uiModel) currentSection(cp copyCatalog) navEntry {
	list := m.sections(cp)
	for _, item := range list {
		if item.id == m.section {
			return item
		}
	}
	return list[0]
}

func (m *uiModel) navigate(next string) {
	if next == "" || next == m.section {
		return
	}
	m.section = next
	m.workspaceScroll.Y = 0
	m.focusHeading = true
}

func (m *uiModel) dirty() bool {
	if m.state == nil {
		return false
	}
	s := m.state.Settings
	return m.groupDraft != s.RestoreGroupShortcut || m.tabDraft != s.ReopenTabShortcut
}

func (m *uiModel) set(key string, value any) {
	if m.actions == nil || m.pending[key] {
		return
	}
	m.pending[key] = true
	m.actionError = false
	m.actions.Set(key, normalizeSettingValue(key, value))
}

func (m *uiModel) saveShortcuts(groupEnabled, tabEnabled, notifyApplied bool) {
	if m.actions == nil || m.pending["shortcuts"] {
		return
	}
	m.pending["shortcuts"] = true
	m.actionError = false
	m.shortcutApplied = false
	m.actions.Shortcuts(groupEnabled, m.groupDraft, tabEnabled, m.tabDraft, notifyApplied)
}

func (m *uiModel) runAction(key string, call func()) {
	if m.actions == nil || m.pending[key] {
		return
	}
	m.pending[key] = true
	m.actionError = false
	call()
}

// Note: normalizeSettingValue maps UI string keys back to wire values cleanly where bridge requires booleans.
func normalizeSettingValue(key string, value any) any {
	if key == "doubleClickCloseIncludeNotepad" {
		if str, ok := value.(string); ok {
			return strings.EqualFold(str, "include")
		}
	}
	if key == "restoreOnAnyFolder" {
		if str, ok := value.(string); ok {
			return strings.EqualFold(str, "any")
		}
	}
	return value
}
