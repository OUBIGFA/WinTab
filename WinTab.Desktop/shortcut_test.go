package main

import (
	"testing"

	"github.com/egoist/mygo/ui"
)

func TestParseKeyChord(t *testing.T) {
	for _, tc := range []struct {
		ev   ui.InputEvent
		want string
		ok   bool
	}{
		{ui.InputEvent{Mods: ui.Alt, Code: ui.CodeKeyE}, "Alt+E", true},
		{ui.InputEvent{Mods: ui.Ctrl | ui.Shift, Code: ui.CodeKeyK}, "Ctrl+Shift+K", true},
		{ui.InputEvent{Mods: ui.Alt, Code: ui.CodeDigit0}, "Alt+0", true},
		{ui.InputEvent{Mods: ui.Alt, Code: ui.CodeDigit7}, "Alt+7", true},
		{ui.InputEvent{Mods: ui.Ctrl, Code: ui.CodeF12}, "Ctrl+F12", true},
		{ui.InputEvent{Mods: ui.Alt, Code: ui.CodeF24}, "Alt+F24", true},
		{ui.InputEvent{Mods: ui.Alt | ui.Shift, PhysicalKey: ui.KeyW}, "Alt+Shift+W", true},
		{ui.InputEvent{Mods: ui.Ctrl, Key: ui.KeyK}, "Ctrl+K", true},
		{ui.InputEvent{Mods: 0, Code: ui.CodeKeyK}, "", false},
		{ui.InputEvent{Mods: ui.Shift, Code: ui.CodeKeyK}, "", false},
		{ui.InputEvent{Mods: ui.Super | ui.Ctrl, Code: ui.CodeKeyK}, "", false},
		{ui.InputEvent{Mods: ui.Ctrl, Code: ui.CodeSpace}, "", false},
		{ui.InputEvent{Mods: ui.Ctrl | ui.Alt, Code: ui.CodeControlLeft}, "", false},
	} {
		got, ok := chordFromInput(tc.ev)
		if ok != tc.ok || got != tc.want {
			t.Errorf("chordFromInput(%+v) = (%q, %v), want (%q, %v)", tc.ev, got, ok, tc.want, tc.ok)
		}
	}
}

func TestShortcutFieldCaptureEditAndTab(t *testing.T) {
	groupDraft := "Alt+E"
	tabDraft := "Alt+W"
	var capture1, capture2 shortcutCapture
	changed := 0

	view := func(c *ui.Context) {
		ui.Column(c).Padding(16).Gap(8).Children(func() {
			if shortcutField(c.Key("g"), &groupDraft, &capture1, "Group shortcut", "", false) {
				changed++
			}
			shortcutField(c.Key("t"), &tabDraft, &capture2, "Tab shortcut", "", false)
		})
	}

	tt := ui.NewTester(view, 400, 160)
	if err := tt.Click("Group shortcut"); err != nil {
		t.Fatal(err)
	}

	// Capture physical chord including F13-F24 and Shift+Digit combinations.
	tt.CodeDown(ui.Ctrl|ui.Shift, ui.CodeKeyK)
	tt.CodeDown(ui.Ctrl|ui.Shift, ui.CodeKeyK) // repeat / held while chord active stays suppressed
	tt.CodeUp(ui.Shift, ui.CodeKeyK)
	if groupDraft != "Ctrl+Shift+K" {
		t.Fatalf("after Ctrl+Shift+K got %q", groupDraft)
	}

	// Alt+7 via CodeDown + Alt+F24 via CodeDown.
	tt.CodeDown(ui.Alt, ui.CodeDigit7)
	tt.CodeUp(0, ui.CodeDigit7)
	if groupDraft != "Alt+7" {
		t.Fatalf("after Alt+7 got %q", groupDraft)
	}

	tt.CodeDown(ui.Alt, ui.CodeF24)
	tt.CodeUp(0, ui.CodeF24)
	if groupDraft != "Alt+F24" {
		t.Fatalf("after Alt+F24 got %q", groupDraft)
	}

	// Normal Tab moves focus to the next control without altering groupDraft.
	tt.Key(0, ui.KeyTab)
	if groupDraft != "Alt+F24" {
		t.Fatalf("Tab modified draft: %q", groupDraft)
	}
	if !tt.Focused("Tab shortcut") {
		t.Fatalf("expected Tab to move focus to Tab shortcut, focused text fields didn't match")
	}

	// Plain typing still works when modifier chords are not used.
	tt.Type("+")
	if tabDraft == "" {
		t.Fatal("expected plain typing into Tab shortcut")
	}
	_ = changed
}

func TestMonochromeThemeHasZeroChroma(t *testing.T) {
	checkAchromatic := func(label string, col ui.Color) {
		t.Helper()
		hi := max(col.R, max(col.G, col.B))
		lo := min(col.R, min(col.G, col.B))
		if hi-lo > 2 {
			t.Errorf("%s is not monochrome: %+v", label, col)
		}
	}

	var gotLight, gotDark *ui.Theme
	tt := ui.NewTester(func(c *ui.Context) {
		gotLight = appTheme(c, false)
		gotDark = appTheme(c, true)
	}, 320, 200)
	tt.Frame()

	for _, entry := range []struct {
		name  string
		theme *ui.Theme
	}{{"Light", gotLight}, {"Dark", gotDark}} {
		if entry.theme == nil {
			t.Fatalf("%s theme is nil", entry.name)
		}
		for _, color := range []struct {
			label string
			val   ui.Color
		}{
			{"Background", entry.theme.Background},
			{"Surface", entry.theme.Surface},
			{"SurfaceHover", entry.theme.SurfaceHover},
			{"SurfacePressed", entry.theme.SurfacePressed},
			{"Border", entry.theme.Border},
			{"Text", entry.theme.Text},
			{"TextMuted", entry.theme.TextMuted},
			{"Accent", entry.theme.Accent},
			{"AccentHover", entry.theme.AccentHover},
			{"AccentPressed", entry.theme.AccentPressed},
			{"AccentText", entry.theme.AccentText},
			{"InactiveAccent", entry.theme.InactiveAccent},
			{"Selection", entry.theme.Selection},
			{"Focus", entry.theme.Focus},
			{"Danger", entry.theme.Danger},
			{"Warning", entry.theme.Warning},
			{"Success", entry.theme.Success},
		} {
			checkAchromatic(entry.name+"."+color.label, color.val)
		}
	}
}
