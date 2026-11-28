package main

import (
	"github.com/egoist/mygo/ui"
	"strings"
)

// Window shell, navigation and shared controls. Feature pages live in ui_pages.go.
func (m *uiModel) view(c *ui.Context) {
	// Before SetTheme, the context carries the default theme, which follows the system appearance.
	m.systemDark = c.Theme().Dark
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
		ui.Themed(c, sidebarTheme(t)).Grow(1).Background(ui.Transparent).Children(func() {
			if ui.Sidebar(c, &chosen, func() {
				for _, item := range m.sections(cp) {
					row := ui.SidebarItem(c, item.id, item.icon, item.label).TextColor(t.Text)
					if item.id == chosen {
						// Override the focused selection fill without tinting the icons.
						row.Background(t.SurfacePressed)
					}
				}
			}).Grow(1).Background(ui.Transparent).Padding(8, 10).Label(cp.Navigation).Changed() {
				m.navigate(chosen)
			}
		})

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
			statusIcon = iconDangerCircle
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
					ui.Icon(c, iconAltArrowRight).FontSize(14).TextColor(t.TextMuted)
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
			if ghostIconButton(c.Key("lang-btn"), iconTranslate, cp.Language, langDisabled).Clicked() {
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

	ui.Scroll(c.Key("scroll-" + m.section)).TrackScroll(&m.workspaceScroll).Grow(1).FillWidth().Children(func() {
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
		})
	})
}

func (m *uiModel) renderConnection(c *ui.Context, cp copyCatalog) {
	t := c.Theme()
	ui.Column(c).Padding(40, 0).Gap(14).Role(ui.RoleStatus).Children(func() {
		if m.loadError {
			ui.Icon(c, iconDangerCircle).FontSize(26).TextColor(t.TextMuted)
			ui.Text(c, cp.LoadErrorTitle).FontSize(22).FontWeight(600).Role(ui.RoleHeading).Level(1)
			ui.Text(c, cp.LoadError).TextColor(t.TextMuted)
			if ui.PrimaryButton(c, "").Children(func() {
				ui.Icon(c, iconRefresh)
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
		ui.Icon(c, iconDangerCircle).FontSize(16).Margin(2, 0, 0, 0)
		ui.Column(c).Grow(1).MinWidth(0).Gap(4).Children(func() {
			for _, msg := range messages {
				ui.Text(c, msg).FontSize(13.5)
			}
		})
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
	ui.Row(c.Key("opt-" + key)).FillWidth().Gap(16).AlignItems(ui.Center).Children(func() {
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
