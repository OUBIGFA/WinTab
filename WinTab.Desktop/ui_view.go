package main

import (
	"fmt"
	"strings"

	"github.com/egoist/mygo/ui"
)

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
	return m.state != nil && m.state.Settings.Theme == "Dark"
}

func (m *uiModel) copy() copyCatalog {
	return copyForLanguage(m.language())
}

func (m *uiModel) sections(cp copyCatalog) []navEntry {
	return []navEntry{
		{id: "explorer", label: cp.Explorer, intro: cp.ExplorerIntro, icon: iconFolderOpen},
		{id: "recovery", label: cp.RecoveryNav, intro: cp.RecoveryIntro, icon: iconRotateCCW},
		{id: "preferences", label: cp.Preferences, intro: cp.PreferencesIntro, icon: iconSettings2},
		{id: "about", label: cp.About, intro: cp.AboutIntro, icon: iconInfo},
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

func (m *uiModel) view(c *ui.Context) {
	dark := m.dark()
	t := appTheme(c, dark)
	c.SetTheme(t)
	if m.heading.Focused(c) && m.focusHeading {
		m.focusHeading = false
	}

	cp := m.copy()
	width, _ := c.Size()
	compact := width <= 720
	bar := c.TitleBar()
	current := m.currentSection(cp)

	if compact {
		ui.Column(c).Fill().Background(t.Background).Children(func() {
			m.renderTitleBar(c, bar, cp, current, dark, true)
			m.renderCompactNav(c, cp, dark)
			m.renderWorkspace(c, cp, current, dark, true)
		})
		return
	}

	ui.Row(c).Fill().AlignItems(ui.Stretch).Background(t.Background).Children(func() {
		m.renderSidebar(c, bar, cp, dark)
		ui.Column(c).Grow(1).MinWidth(0).FillHeight().Background(t.Background).Children(func() {
			m.renderTitleBar(c, bar, cp, current, dark, false)
			m.renderWorkspace(c, cp, current, dark, false)
		})
	})
}

func brandLogoPlate(c *ui.Context, dark bool, size, pad, radius float32) {
	if !dark {
		ui.Image(c, logoBitmap).Size(size, size).Fit(ui.Contain)
		return
	}
	inner := max(size-pad*2, 1)
	ui.Box(c).Size(size, size).Radius(radius).Background(ui.Hex("#ffffff")).Center().Children(func() {
		ui.Image(c, logoBitmap).Size(inner, inner).Fit(ui.Contain)
	})
}

func (m *uiModel) renderSidebar(c *ui.Context, bar ui.TitleBar, cp copyCatalog, dark bool) {
	t := c.Theme()
	ui.Column(c).Width(224).Shrink(0).FillHeight().Background(sidebarBackground(dark)).
		BorderWidth(0, 1, 0, 0).BorderColor(t.Border).Children(func() {
		ui.Row(c).Height(max(bar.Height, 44)).Padding(0, 16).Gap(9).AlignItems(ui.Center).
			DragWindow().Children(func() {
			brandLogoPlate(c, dark, 20, 2, 6)
			ui.Text(c, "WinTab").FontSize(14).FontWeight(600).SingleLine()
		})

		chosen := m.section
		if ui.Sidebar(c, &chosen, func() {
			for _, item := range m.sections(cp) {
				ui.SidebarItem(c, item.id, item.icon, item.label)
			}
		}).Grow(1).Background(ui.Transparent).Padding(8, 10).Label(cp.Navigation).Changed() {
			m.navigate(chosen)
		}

		m.renderSidebarFooter(c, cp, dark)
	})
}

func (m *uiModel) renderSidebarFooter(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()
	ui.Column(c).Padding(12, 16, 16).Gap(4).BorderWidth(1, 0, 0, 0).BorderColor(dividerColor(dark)).Children(func() {
		statusText := cp.Connecting
		statusIcon := (*ui.SVG)(nil)
		switch {
		case m.loadError:
			statusText = cp.LoadErrorTitle
			statusIcon = iconCircleAlert
		case len(m.pending) > 0:
			statusText = cp.Saving
		case m.state != nil && m.state.ShellReady:
			statusText = cp.Connected
			statusIcon = iconCheck
		}

		ui.Row(c).Gap(7).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
			if statusIcon != nil {
				ui.Icon(c, statusIcon).FontSize(13).TextColor(t.TextMuted)
			} else {
				ui.Spinner(c).FontSize(13).TextColor(t.TextMuted)
			}
			ui.Text(c, statusText).FontSize(12).TextColor(t.TextMuted).SingleLine()
		})
		if m.state != nil {
			ui.Textf(c, "WinTab v%s", m.state.Version).FontSize(12).TextColor(t.TextMuted).Padding(0, 0, 0, 20)
		}
	})
}

func (m *uiModel) renderCompactNav(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()
	ui.Row(c).FillWidth().Padding(8, 16).Gap(6).Background(sidebarBackground(dark)).
		BorderWidth(0, 0, 1, 0).BorderColor(t.Border).FocusGroup(ui.Horizontal).Role(ui.RoleTabList).Label(cp.Navigation).Children(func() {
		for _, item := range m.sections(cp) {
			active := item.id == m.section
			btn := ui.ButtonBase(c.Key("nav-"+item.id)).Padding(6, 10).Radius(t.Radius).Gap(6).AlignItems(ui.Center).
				Role(ui.RoleTab).Label(item.label)
			if active {
				btn.Background(t.Background).Border(1, t.Border).Current(ui.CurrentPage)
			} else if btn.Hovered() {
				btn.Background(t.SurfaceHover)
			}
			btn.Children(func() {
				col := t.TextMuted
				weight := 500
				if active {
					col = t.Text
					weight = 600
				}
				ui.Icon(c, item.icon).FontSize(14).TextColor(col)
				ui.Text(c, item.label).FontSize(13).FontWeight(weight).TextColor(col).SingleLine()
			})
			if btn.Clicked() {
				m.navigate(item.id)
			}
		}
	})
}

func (m *uiModel) renderTitleBar(c *ui.Context, bar ui.TitleBar, cp copyCatalog, current navEntry, dark, compact bool) {
	t := c.Theme()
	h := max(bar.Height, 44)
	header := ui.Box(c).FillWidth().Height(h).DragWindow().Background(t.Background).
		BorderWidth(0, 0, 1, 0).BorderColor(t.Border)

	header.Children(func() {
		padX := float32(36)
		if compact {
			padX = 16
		}
		track := ui.Row(c).FillHeight().FillWidth().Padding(0, padX).AlignItems(ui.Center).Gap(9)
		if !compact {
			track.MaxWidth(840).Margin(0, ui.Auto)
		}
		track.Children(func() {
			if compact {
				brandLogoPlate(c, dark, 20, 2, 6)
				ui.Text(c, "WinTab").FontSize(14).FontWeight(600).SingleLine()
			} else {
				ui.Row(c).Gap(6).AlignItems(ui.Center).Role(ui.RoleNone).Children(func() {
					ui.Text(c, cp.Settings).FontSize(13).TextColor(t.TextMuted).SingleLine()
					ui.Icon(c, iconChevronRight).FontSize(14).TextColor(t.TextMuted)
					ui.Text(c, current.label).FontSize(13).FontWeight(500).SingleLine()
				})
			}
		})

		rightInset := bar.Right + 12
		if compact {
			rightInset = bar.Right + 8
		}
		ui.Row(c).Attach(ui.AnchorRight, ui.AnchorRight).Right(rightInset).FillHeight().AlignItems(ui.Center).Gap(4).Children(func() {
			themeLabel := cp.Dark
			themeIcon := iconMoon
			nextTheme := "Dark"
			if dark {
				themeLabel = cp.Light
				themeIcon = iconSun
				nextTheme = "Light"
			}
			themeDisabled := m.state == nil || m.pending["theme"]
			if ghostIconButton(c.Key("theme-btn"), themeIcon, themeLabel, themeDisabled).Clicked() {
				m.set("theme", nextTheme)
			}

			nextLang := "en-US"
			if m.language() == "en-US" {
				nextLang = "zh-CN"
			}
			langDisabled := m.state == nil || m.pending["language"]
			if ghostIconButton(c.Key("lang-btn"), iconLanguages, cp.Language, langDisabled).Clicked() {
				m.set("language", nextLang)
			}
		})
	})
}

func ghostIconButton(c *ui.Context, ic *ui.SVG, label string, disabled bool) ui.Element {
	t := c.Theme()
	btn := ui.ButtonBase(c).Size(30, 30).Radius(t.Radius).Center().
		Disabled(disabled).Label(label).Tooltip(label)
	if !disabled && btn.Hovered() {
		btn.Background(t.SurfaceHover)
	}
	btn.Children(func() {
		col := t.TextMuted
		if !disabled && btn.Hovered() {
			col = t.Text
		}
		ui.Icon(c, ic).FontSize(15).TextColor(col)
	})
	return btn
}

func (m *uiModel) renderWorkspace(c *ui.Context, cp copyCatalog, current navEntry, dark, compact bool) {
	padTop, padX, padBottom := float32(24), float32(36), float32(22)
	if compact {
		padTop, padX, padBottom = 18, 16, 20
	}

	ui.Scroll(c.Key("scroll-" + m.section)).Grow(1).FillWidth().Children(func() {
		page := ui.Column(c).FillWidth().Padding(padTop, padX, padBottom)
		if !compact {
			page.MaxWidth(840).Margin(0, ui.Auto)
		}
		page.Children(func() {
			if m.state == nil || m.loadError {
				m.renderConnection(c, cp)
				return
			}

			// Page header
			ui.Column(c.Key("page-header")).Gap(4).Margin(0, 0, 20, 0).Children(func() {
				h1 := ui.Text(c, current.label).FontSize(24).FontWeight(600).
					Role(ui.RoleHeading).Level(1).Focusable().TabStop(false).Bind(&m.heading)
				if m.focusHeading {
					h1.Focus(ui.FocusOptions{PreventScroll: true})
					m.focusHeading = false
				}
				ui.Text(c, current.intro).FontSize(13.5).TextColor(c.Theme().TextMuted)
			})

			m.renderNotices(c, cp, dark)

			switch m.section {
			case "explorer":
				m.renderExplorer(c, cp, dark)
			case "recovery":
				m.renderRecovery(c, cp, dark)
			case "preferences":
				m.renderPreferences(c, cp, dark)
			case "about":
				m.renderAbout(c, cp, dark)
			}

			m.renderFooter(c, cp)
		})
	})
}

func (m *uiModel) renderConnection(c *ui.Context, cp copyCatalog) {
	t := c.Theme()
	ui.Column(c).Padding(40, 0).Gap(14).Role(ui.RoleStatus).Children(func() {
		if m.loadError {
			ui.Icon(c, iconCircleAlert).FontSize(26).TextColor(t.TextMuted)
			ui.Text(c, cp.LoadErrorTitle).FontSize(22).FontWeight(600).Role(ui.RoleHeading).Level(1)
			ui.Text(c, cp.LoadError).TextColor(t.TextMuted)
			if ui.PrimaryButton(c, "").Children(func() {
				ui.Icon(c, iconRefreshCW)
				ui.Text(c, cp.Retry).SingleLine()
			}).Clicked() && m.actions != nil {
				m.loadError = false
				m.actions.Reconnect()
			}
			return
		}
		ui.Spinner(c).FontSize(26).TextColor(t.TextMuted)
		ui.Text(c, cp.LoadTitle).FontSize(22).FontWeight(600).Role(ui.RoleHeading).Level(1)
		ui.Text(c, cp.Connecting).TextColor(t.TextMuted)
	})
}

func (m *uiModel) renderNotices(c *ui.Context, cp copyCatalog, dark bool) {
	var messages []string
	if m.actionError {
		messages = append(messages, cp.RequestError)
	}
	if m.state.HookError != nil && *m.state.HookError != "" {
		messages = append(messages, *m.state.HookError)
	}
	if m.state.StorageError != nil && *m.state.StorageError != "" {
		messages = append(messages, *m.state.StorageError)
	}
	if len(messages) == 0 {
		return
	}
	t := c.Theme()
	ui.Row(c).Padding(12, 16).Gap(10).Margin(0, 0, 16, 0).Radius(10).
		Background(cardSubtleBackground(dark)).Border(1, t.Border).AlignItems(ui.Start).
		Role(ui.RoleAlert).Label(strings.Join(messages, " ")).Children(func() {
		ui.Icon(c, iconCircleAlert).FontSize(16).Margin(2, 0, 0, 0)
		ui.Column(c).Grow(1).MinWidth(0).Gap(4).Children(func() {
			for _, msg := range messages {
				ui.Text(c, msg).FontSize(13.5)
			}
		})
	})
}

func (m *uiModel) renderFooter(c *ui.Context, cp copyCatalog) {
	t := c.Theme()
	show := m.section != "about" && !m.dirty() && !m.actionError &&
		(m.state.ShortcutError == nil || *m.state.ShortcutError == "") &&
		(m.state.StorageError == nil || *m.state.StorageError == "")

	ui.Row(c).MinHeight(36).Padding(18, 0, 0, 0).Gap(7).AlignItems(ui.Center).Children(func() {
		if !show {
			return
		}
		if len(m.pending) > 0 {
			ui.Spinner(c).FontSize(13).TextColor(t.TextMuted)
			ui.Text(c, cp.Saving).FontSize(12).TextColor(t.TextMuted)
		} else {
			ui.Icon(c, iconCheck).FontSize(13).TextColor(t.TextMuted)
			ui.Text(c, cp.Saved).FontSize(12).TextColor(t.TextMuted)
		}
	})
}

func settingGroup(c *ui.Context, title string, dark bool, body func()) {
	t := c.Theme()
	ui.Column(c).FillWidth().Gap(8).Margin(0, 0, 16, 0).Children(func() {
		if title != "" {
			ui.Text(c, title).FontSize(13).FontWeight(600).TextColor(t.TextMuted).Role(ui.RoleHeading).Level(2)
		}
		ui.Column(c).FillWidth().Radius(10).Clip().Background(cardBackground(dark)).Border(1, t.Border).
			Dividers(1, dividerColor(dark)).Children(body)
	})
}

func renderThemedSwitch(c *ui.Context, key string, val *bool, title string, disabled bool, titleEl, descEl ui.Element, dark bool) ui.Element {
	var sw ui.Element
	build := func() {
		sw = ui.Switch(c.Key("sw-"+key), val).Disabled(disabled).Label(title).LabelledBy(titleEl)
		if descEl.Valid() {
			sw.DescribedBy(descEl)
		}
	}
	if dark {
		ui.Themed(c, darkSwitchTheme(c.Theme())).Background(ui.Transparent).Children(build)
	} else {
		build()
	}
	return sw
}

func (m *uiModel) toggleRow(c *ui.Context, key, title, desc string, checked, disabled bool, extra func()) {
	t := c.Theme()
	ui.Column(c.Key(key)).FillWidth().Padding(12, 18).Children(func() {
		ui.Row(c).FillWidth().Gap(24).AlignItems(ui.Center).Children(func() {
			var titleEl, descEl ui.Element
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				titleEl = ui.Text(c, title).FontSize(14).FontWeight(600)
				if desc != "" {
					descEl = ui.Text(c, desc).FontSize(12.5).TextColor(t.TextMuted)
				}
			})
			val := checked
			if renderThemedSwitch(c, key, &val, title, disabled || m.pending[key], titleEl, descEl, m.dark()).Changed() {
				m.set(key, val)
			}
		})
		if extra != nil {
			ui.Column(c).FillWidth().Margin(10, 0, 0, 0).Padding(10, 0, 0, 0).
				BorderWidth(1, 0, 0, 0).BorderColor(t.Border).Children(extra)
		}
	})
}

func (m *uiModel) segmentedOptionRow(c *ui.Context, key, label, hint string, values, labels []string, current string, disabled bool) {
	t := c.Theme()
	selected := 0
	for i, val := range values {
		if val == current {
			selected = i
			break
		}
	}
	ui.Row(c.Key("opt-"+key)).FillWidth().Gap(16).AlignItems(ui.Center).Children(func() {
		ui.Column(c).Grow(1).MinWidth(0).Gap(2).Children(func() {
			ui.Text(c, label).FontSize(13).FontWeight(500)
			if hint != "" {
				ui.Text(c, hint).FontSize(12.5).TextColor(t.TextMuted)
			}
		})
		idx := selected
		if ui.Segmented(c.Key("seg-"+key), &idx, labels...).Disabled(disabled || m.pending[key]).Label(label).Changed() {
			if idx >= 0 && idx < len(values) {
				m.set(key, values[idx])
			}
		}
	})
}

func (m *uiModel) renderExplorer(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()
	s := m.state.Settings

	settingGroup(c, cp.WindowsAndTabs, dark, func() {
		m.toggleRow(c, "windowHook", cp.Merge, cp.MergeDesc, s.WindowHook, false, nil)
		m.toggleRow(c, "reuseTabs", cp.Reuse, cp.ReuseDesc, s.ReuseTabs, false, nil)
	})

	settingGroup(c, cp.MouseActions, dark, func() {
		m.toggleRow(c, "doubleClickCloseTab", cp.DoubleClick, cp.DoubleClickDesc, s.DoubleClickCloseTab, false, func() {
			cur := "explorer"
			if s.DoubleClickCloseIncludeNotepad {
				cur = "include"
			}
			m.segmentedOptionRow(c, "doubleClickCloseIncludeNotepad", cp.Scope, "",
				[]string{"explorer", "include"}, []string{cp.ExplorerOnly, cp.IncludeNotepad}, cur, !s.DoubleClickCloseTab)
			// Override callback semantics: doubleClickCloseIncludeNotepad sends bool to C# bridge
		})
		m.toggleRow(c, "doubleClickEmptyArea", cp.EmptyDoubleClick, cp.EmptyDoubleClickD, s.DoubleClickEmptyArea, false, func() {
			m.segmentedOptionRow(c, "doubleClickEmptyAreaAction", cp.EmptyAction, "",
				[]string{"Parent", "Back"}, []string{cp.Parent, cp.Back}, s.DoubleClickEmptyAreaAction, !s.DoubleClickEmptyArea)
		})
		m.toggleRow(c, "middleClickForegroundTab", cp.MiddleClick, cp.MiddleClickDesc, s.MiddleClickForegroundTab, false, nil)
		m.toggleRow(c, "wheelSwitchTab", cp.Wheel, cp.WheelDesc, s.WheelSwitchTab, false, func() {
			hint := cp.MediumHint
			switch s.WheelSwitchSensitivity {
			case "Low":
				hint = cp.LowHint
			case "High":
				hint = cp.HighHint
			}
			m.segmentedOptionRow(c, "wheelSwitchSensitivity", cp.Sensitivity, hint,
				[]string{"Low", "Medium", "High"}, []string{cp.Low, cp.Medium, cp.High}, s.WheelSwitchSensitivity, !s.WheelSwitchTab)
		})
	})

	// Hold Ctrl + Shift context note
	ui.Row(c).FillWidth().Padding(11, 16).Gap(8).Radius(t.Radius).Background(cardSubtleBackground(dark)).
		Border(1, t.Border).AlignItems(ui.Center).Children(func() {
		ui.Icon(c, iconInfo).FontSize(15).TextColor(t.TextMuted)
		ui.Text(c, cp.Bypass).FontSize(13).TextColor(t.TextMuted)
		keyCap(c, "Ctrl", dark)
		ui.Text(c, "+").FontSize(13).TextColor(t.TextMuted)
		keyCap(c, "Shift", dark)
		ui.Text(c, cp.BypassEnd).FontSize(13).TextColor(t.TextMuted)
	})
}

func keyCap(c *ui.Context, text string, dark bool) {
	t := c.Theme()
	ui.Box(c).Padding(1, 6).Radius(4).Background(cardBackground(dark)).Border(1, t.Border).Children(func() {
		ui.Text(c, text).FontSize(12).FontWeight(500)
	})
}

func (m *uiModel) renderRecovery(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()
	s := m.state.Settings

	settingGroup(c, cp.ManualRecovery, dark, func() {
		// Group restore block
		m.recoveryCommandBlock(c, cp, true, cp.Group, cp.Global, cp.EnableGroup, cp.GroupShortcut,
			s.RestoreGroupShortcutEnabled, &m.groupDraft, &m.groupCapture,
			m.state.ShellReady && !m.state.SessionBusy && !m.pending["restore"],
			cp.Single, cp.SingleDesc, "restoreSingleTab", s.RestoreSingleTab, false, dark)

		// Single-tab reopen block
		tabDesc := cp.RecordHint
		if s.ReopenTabShortcutEnabled {
			tabDesc = cp.RecordRequired
		}
		m.recoveryCommandBlock(c, cp, false, cp.Tab, cp.InExplorer, cp.EnableTab, cp.TabShortcut,
			s.ReopenTabShortcutEnabled, &m.tabDraft, &m.tabCapture,
			m.state.ShellReady && m.state.RecordClosedTabs && !m.state.SessionBusy && !m.pending["restore"],
			cp.Record, tabDesc, "reopenClosedTab", m.state.RecordClosedTabs, s.ReopenTabShortcutEnabled, dark)

		// Apply bar (fixed vertical slot height so zero bounce)
		ui.Row(c).FillWidth().Padding(13, 18).Gap(20).AlignItems(ui.Center).Background(cardSubtleBackground(dark)).Children(func() {
			dirty := m.dirty()
			hasErr := m.state.ShortcutError != nil && *m.state.ShortcutError != ""
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				ui.Text(c, cp.CaptureHint).FontSize(13).FontWeight(500)
				switch {
				case hasErr:
					ui.Row(c).Gap(6).AlignItems(ui.Center).Role(ui.RoleAlert).Children(func() {
						ui.Icon(c, iconCircleAlert).FontSize(14)
						ui.Text(c, *m.state.ShortcutError).FontSize(12.5).FontWeight(500)
					})
				case dirty:
					ui.Row(c).Gap(6).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
						ui.Icon(c, iconInfo).FontSize(14)
						ui.Text(c, cp.Draft).FontSize(12.5).FontWeight(500)
					})
				case m.shortcutApplied:
					ui.Row(c).Gap(6).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
						ui.Icon(c, iconCheck).FontSize(14)
						ui.Text(c, cp.Applied).FontSize(12.5).FontWeight(500)
					})
				default:
					ui.Text(c, cp.ShortcutsNote).FontSize(12.5).TextColor(t.TextMuted)
				}
			})

			applyBtn := func() ui.Element {
				if dirty {
					return ui.PrimaryButton(c.Key("apply-shortcuts"), "").Border(1, t.Accent)
				}
				return ui.Button(c.Key("apply-shortcuts"), "")
			}()
			applyBtn.MinHeight(34).Label(cp.Apply).Disabled(m.pending["shortcuts"] || !dirty).Children(func() {
				ui.Icon(c, iconCheck).FontSize(15)
				ui.Text(c, cp.Apply).SingleLine()
			})
			if applyBtn.Clicked() {
				m.saveShortcuts(s.RestoreGroupShortcutEnabled, s.ReopenTabShortcutEnabled, true)
			}
		})
	})

	if m.state.SessionFeedback != nil && *m.state.SessionFeedback != "" {
		ui.Row(c).FillWidth().Padding(10, 14).Gap(8).Margin(0, 0, 20, 0).Radius(t.Radius).
			Background(cardSubtleBackground(dark)).Border(1, t.Border).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
			ui.Icon(c, iconInfo).FontSize(14)
			ui.Text(c, *m.state.SessionFeedback).FontSize(13).FontWeight(500)
		})
	}

	settingGroup(c, cp.Automatic, dark, func() {
		m.toggleRow(c, "restoreTabs", cp.AutoRestore, cp.AutoRestoreDesc, s.RestoreTabs, false, func() {
			hint := cp.AutoIndependent
			if s.RestoreTabs {
				if s.RestoreOnAnyFolder {
					hint = cp.AnyHint
				} else {
					hint = cp.NormalHint
				}
			}
			cur := "normal"
			if s.RestoreOnAnyFolder {
				cur = "any"
			}
			m.segmentedOptionRow(c, "restoreOnAnyFolder", cp.RestoreMode, hint,
				[]string{"normal", "any"}, []string{cp.NormalLaunch, cp.AnyFolder}, cur, !s.RestoreTabs)
		})
	})

	// Recovery scope & exceptions collapsible
	ui.Box(c).FillWidth().Padding(12, 18).Radius(10).Background(cardBackground(dark)).Border(1, t.Border).Children(func() {
		ui.Collapsible(c.Key("recovery-details"), cp.RecoveryDetails, &m.detailsOpen, func() {
			ui.Text(c, cp.RecoveryNotes).FontSize(13).TextColor(t.TextMuted).Padding(6, 0, 0, 0)
		})
	})
}

func (m *uiModel) recoveryCommandBlock(
	c *ui.Context, cp copyCatalog, group bool,
	title, scope, enableLabel, inputLabel string,
	shortcutEnabled bool, draft *string, capture *shortcutCapture, canRestore bool,
	subTitle, subDesc, subKey string, subChecked, subDisabled, dark bool,
) {
	t := c.Theme()
	s := m.state.Settings
	blockKey := "cmd-tab"
	if group {
		blockKey = "cmd-group"
	}

	ui.Column(c.Key(blockKey)).FillWidth().Padding(14, 18).Gap(12).Children(func() {
		ui.Row(c).FillWidth().Gap(20).AlignItems(ui.Center).Children(func() {
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				ui.Text(c, title).FontSize(14).FontWeight(600)
				ui.Text(c, scope).FontSize(12.5).TextColor(t.TextMuted)
			})
			restoring := m.state.SessionBusy || m.pending["restore"]
			btnLabel := cp.Restore
			if restoring {
				btnLabel = cp.Restoring
			}
			btn := ui.Button(c.Key("btn-"+blockKey), "").Label(btnLabel).Disabled(!canRestore).Children(func() {
				if restoring {
					ui.Spinner(c).FontSize(14)
				} else {
					ui.Icon(c, iconRotateCCW).FontSize(14)
				}
				ui.Text(c, btnLabel).SingleLine()
			})
			if btn.Clicked() {
				m.runAction("restore", func() { m.actions.Restore(group) })
			}
		})

		ui.Row(c).FillWidth().Gap(12).AlignItems(ui.Center).Children(func() {
			checked := shortcutEnabled
			cb := ui.Checkbox(c.Key("cb-"+blockKey), &checked, cp.Shortcut).
				Label(enableLabel).Disabled(m.pending["shortcuts"])
			if cb.Changed() {
				ge, te := s.RestoreGroupShortcutEnabled, s.ReopenTabShortcutEnabled
				if group {
					ge = checked
				} else {
					te = checked
				}
				m.saveShortcuts(ge, te, false)
			}
			ui.Spacer(c)
			ui.Icon(c, iconKeyboard).FontSize(15).TextColor(t.TextMuted)
			hasErr := m.state.ShortcutError != nil && *m.state.ShortcutError != ""
			if shortcutField(c.Key("in-"+blockKey), draft, capture, inputLabel, cp.CaptureHelp, hasErr) {
				m.shortcutApplied = false
			}
		})

		ui.Row(c).FillWidth().Gap(24).Padding(10, 0, 0, 0).BorderWidth(1, 0, 0, 0).
			BorderColor(dividerColor(dark)).AlignItems(ui.Center).Children(func() {
			var subTitleEl, subDescEl ui.Element
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				subTitleEl = ui.Text(c, subTitle).FontSize(13.5).FontWeight(500)
				subDescEl = ui.Text(c, subDesc).FontSize(12.5).TextColor(t.TextMuted)
			})
			val := subChecked
			if renderThemedSwitch(c, subKey, &val, subTitle, subDisabled || m.pending[subKey], subTitleEl, subDescEl, dark).Changed() {
				m.set(subKey, val)
			}
		})
	})
}

func (m *uiModel) renderPreferences(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()
	s := m.state.Settings

	settingGroup(c, cp.Appearance, dark, func() {
		ui.Row(c).FillWidth().Padding(14, 18).Gap(20).AlignItems(ui.Center).Children(func() {
			ui.Text(c, cp.ThemeLabel).FontSize(14).FontWeight(600).Grow(1)
			idx := 0
			if s.Theme == "Dark" {
				idx = 1
			}
			if ui.Segmented(c.Key("seg-theme"), &idx, cp.LightChoice, cp.DarkChoice).
				Disabled(m.pending["theme"]).Label(cp.ThemeLabel).Changed() {
				if idx == 1 {
					m.set("theme", "Dark")
				} else {
					m.set("theme", "Light")
				}
			}
		})
		ui.Row(c).FillWidth().Padding(14, 18).Gap(20).AlignItems(ui.Center).Children(func() {
			ui.Text(c, cp.LanguageLabel).FontSize(14).FontWeight(600).Grow(1)
			idx := 0
			if m.language() == "en-US" {
				idx = 1
			}
			if ui.Segmented(c.Key("seg-language"), &idx, "简体中文", "English").
				Disabled(m.pending["language"]).Label(cp.LanguageLabel).Changed() {
				if idx == 1 {
					m.set("language", "en-US")
				} else {
					m.set("language", "zh-CN")
				}
			}
		})
	})

	settingGroup(c, cp.Background, dark, func() {
		m.toggleRow(c, "startup", cp.Startup, cp.StartupDesc, m.state.Startup, false, nil)
		trayDesc := cp.TrayHidden
		if s.ShowTrayIcon {
			trayDesc = cp.TrayDesc
		}
		m.toggleRow(c, "showTrayIcon", cp.Tray, trayDesc, s.ShowTrayIcon, false, nil)
		m.toggleRow(c, "autoUpdate", cp.AutoUpdate, cp.AutoUpdateDesc, s.AutoUpdate, false, nil)
	})

	ui.Row(c).FillWidth().Padding(11, 16).Gap(8).Radius(t.Radius).Background(cardSubtleBackground(dark)).
		Border(1, t.Border).AlignItems(ui.Center).Children(func() {
		ui.Icon(c, iconInfo).FontSize(15).TextColor(t.TextMuted)
		ui.Text(c, cp.Running).FontSize(13).TextColor(t.TextMuted)
	})
}

func (m *uiModel) renderAbout(c *ui.Context, cp copyCatalog, dark bool) {
	t := c.Theme()

	// Product identity card
	ui.Row(c).FillWidth().Padding(18, 20).Gap(18).Margin(0, 0, 20, 0).Radius(10).
		Background(cardBackground(dark)).Border(1, t.Border).AlignItems(ui.Center).Children(func() {
		brandLogoPlate(c, dark, 52, 6, 12)
		ui.Column(c).Grow(1).MinWidth(0).Gap(4).Children(func() {
			ui.Row(c).Gap(10).AlignItems(ui.Center).Children(func() {
				ui.Text(c, "WinTab").FontSize(20).FontWeight(600).Role(ui.RoleHeading).Level(2)
				ui.Badge(c, fmt.Sprintf("v%s", m.state.Version))
			})
			ui.Text(c, cp.AboutDescription).FontSize(13.5).TextColor(t.TextMuted)
		})
	})

	// Key facts card
	settingGroup(c, "", dark, func() {
		factRow(c, cp.Version, func() {
			ui.Textf(c, "v%s", m.state.Version).FontSize(13.5).FontWeight(500)
		})
		factRow(c, cp.LicenseLabel, func() {
			ui.Text(c, cp.License).FontSize(13.5).FontWeight(500)
		})
		factRow(c, cp.Repository, func() {
			lnk := ui.Link(c.Key("repo-link"), cp.Project, "").Tooltip(cp.OpenSource)
			if lnk.Clicked() && m.actions != nil {
				m.actions.OpenProject()
			}
		})
	})

	// Updates & support
	settingGroup(c, cp.Support, dark, func() {
		ui.Row(c).FillWidth().Padding(14, 18).Gap(24).AlignItems(ui.Center).Children(func() {
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				ui.Text(c, cp.Updates).FontSize(14).FontWeight(600)
				ui.Text(c, cp.UpdatesDesc).FontSize(12.5).TextColor(t.TextMuted)
				if m.state.UpdateFeedback != nil && *m.state.UpdateFeedback != "" {
					ui.Row(c).Margin(4, 0, 0, 0).Gap(6).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
						ui.Icon(c, iconCheck).FontSize(14)
						ui.Text(c, *m.state.UpdateFeedback).FontSize(12.5).FontWeight(500)
					})
				}
			})
			updating := m.state.UpdateBusy || m.pending["update"]
			updLabel := cp.Updates
			if updating {
				updLabel = cp.Checking
			}
			updBtn := ui.Button(c.Key("btn-update"), "").Label(updLabel).Disabled(updating).Children(func() {
				if updating {
					ui.Spinner(c).FontSize(14)
				} else {
					ui.Icon(c, iconRefreshCW).FontSize(14)
				}
				ui.Text(c, updLabel).SingleLine()
			})
			if updBtn.Clicked() {
				m.runAction("update", m.actions.CheckUpdates)
			}
		})

		ui.Row(c).FillWidth().Padding(14, 18).Gap(24).AlignItems(ui.Center).Children(func() {
			ui.Column(c).Grow(1).MinWidth(0).Gap(3).Children(func() {
				ui.Text(c, cp.Logs).FontSize(14).FontWeight(600)
				ui.Text(c, cp.LogsDesc).FontSize(12.5).TextColor(t.TextMuted)
			})
			logsBtn := ui.Button(c.Key("btn-logs"), "").Label(cp.Logs).Disabled(m.pending["logs"]).Children(func() {
				ui.Icon(c, iconFolderOpen).FontSize(14)
				ui.Text(c, cp.Logs).SingleLine()
			})
			if logsBtn.Clicked() {
				m.runAction("logs", m.actions.OpenLogs)
			}
		})
	})
}

func factRow(c *ui.Context, label string, right func()) {
	t := c.Theme()
	ui.Row(c).FillWidth().Padding(12, 18).Gap(24).Justify(ui.SpaceBetween).AlignItems(ui.Center).Children(func() {
		ui.Text(c, label).FontSize(13.5).TextColor(t.TextMuted)
		right()
	})
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
