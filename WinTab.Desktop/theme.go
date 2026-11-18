package main

import "github.com/egoist/mygo/ui"

func appTheme(c *ui.Context, dark bool) *ui.Theme {
	prefs := c.Preferences()
	var base *ui.Theme
	if dark {
		base = ui.DarkTheme()
	} else {
		base = ui.LightTheme()
	}
	followed := base.WithPreferences(prefs)

	var accent ui.Color
	if dark {
		accent = ui.Hex("#f5f5f5")
	} else {
		accent = ui.Hex("#171717")
	}
	t := followed.WithAccent(accent)
	t.Radius = 8
	t.Spacing = 4
	t.FocusRingWidth = 2
	t.FocusRingOffset = 1

	if dark {
		t.Background = ui.Hex("#161616")
		t.Surface = ui.Hex("#262626")
		t.SurfaceHover = ui.Hex("#303030")
		t.SurfacePressed = ui.Hex("#3a3a3a")
		t.Border = ui.Hex("#323232")
		t.Text = ui.Hex("#f5f5f5")
		t.TextMuted = ui.Hex("#a3a3a3")
		t.AccentText = ui.Hex("#171717")
		t.InactiveAccent = ui.Hex("#a3a3a3")
		t.Danger = ui.Hex("#f5f5f5")
		t.Warning = ui.Hex("#d4d4d4")
		t.Success = ui.Hex("#f5f5f5")
		t.Scrollbar = ui.RGBA(255, 255, 255, 0.30)
		t.Inverse = ui.Hex("#f5f5f5")
		t.InverseText = ui.Hex("#171717")
	} else {
		t.Background = ui.Hex("#ffffff")
		t.Surface = ui.Hex("#f4f4f4")
		t.SurfaceHover = ui.Hex("#e8e8e8")
		t.SurfacePressed = ui.Hex("#dcdcdc")
		t.Border = ui.Hex("#d8d8d8")
		t.Text = ui.Hex("#171717")
		t.TextMuted = ui.Hex("#525252")
		t.InactiveAccent = ui.Hex("#525252")
		t.Danger = ui.Hex("#171717")
		t.Warning = ui.Hex("#525252")
		t.Success = ui.Hex("#171717")
		t.Scrollbar = ui.RGBA(23, 23, 23, 0.28)
		t.Inverse = ui.Hex("#171717")
		t.InverseText = ui.Hex("#fafafa")
	}

	if prefs.HighContrast {
		if dark {
			t.Border = ui.Hex("#737373")
			t.TextMuted = ui.Hex("#d4d4d4")
			t.Focus = ui.Hex("#ffffff")
		} else {
			t.Border = ui.Hex("#525252")
			t.TextMuted = ui.Hex("#262626")
			t.Focus = ui.Hex("#0a0a0a")
		}
	}
	return t
}

func sidebarBackground(dark bool) ui.Color {
	if dark {
		return ui.Hex("#111111")
	}
	return ui.Hex("#fafafa")
}

func cardBackground(dark bool) ui.Color {
	if dark {
		return ui.Hex("#1c1c1c")
	}
	return ui.Hex("#ffffff")
}

func cardSubtleBackground(dark bool) ui.Color {
	if dark {
		return ui.Hex("#181818")
	}
	return ui.Hex("#f7f7f7")
}

func dividerColor(dark bool) ui.Color {
	if dark {
		return ui.Hex("#282828")
	}
	return ui.Hex("#ebebef")
}

func sidebarTheme(t *ui.Theme) *ui.Theme {
	copy := *t
	// MyGo switches focused icons from Accent to AccentText. Keep both
	// foreground colors identical; item labels use Text independently.
	copy.Accent = t.TextMuted
	copy.InactiveAccent = t.TextMuted
	copy.AccentText = t.TextMuted
	return &copy
}

func darkSwitchTheme(t *ui.Theme) *ui.Theme {
	copy := *t
	copy.Accent = ui.Hex("#9e9e9e")
	copy.AccentHover = ui.Hex("#b0b0b0")
	copy.InactiveAccent = ui.Hex("#737373")
	return &copy
}
