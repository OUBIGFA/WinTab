package main

import (
	"fmt"
	"github.com/egoist/mygo/ui"
)

// Feature pages use the shared state and controls; the window shell lives in ui_view.go.
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
		ui.Icon(c, iconInfoCircle).FontSize(15).TextColor(t.TextMuted)
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
						ui.Icon(c, iconDangerCircle).FontSize(14)
						ui.Text(c, *m.state.ShortcutError).FontSize(12.5).FontWeight(500)
					})
				case dirty:
					ui.Row(c).Gap(6).AlignItems(ui.Center).Role(ui.RoleStatus).Children(func() {
						ui.Icon(c, iconInfoCircle).FontSize(14)
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
			ui.Icon(c, iconInfoCircle).FontSize(14)
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
					ui.Icon(c, iconHistory).FontSize(14)
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
			themes := []string{"Light", "Dark", "System"}
			idx := 0
			for i, value := range themes {
				if s.Theme == value {
					idx = i
				}
			}
			if ui.Segmented(c.Key("seg-theme"), &idx, cp.LightChoice, cp.DarkChoice, cp.SystemChoice).
				Disabled(m.pending["theme"]).Label(cp.ThemeLabel).Changed() && idx >= 0 && idx < len(themes) {
				m.set("theme", themes[idx])
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
		ui.Icon(c, iconInfoCircle).FontSize(15).TextColor(t.TextMuted)
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
					ui.Icon(c, iconRefresh).FontSize(14)
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
