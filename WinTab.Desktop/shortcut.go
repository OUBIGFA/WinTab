package main

import (
	"math"
	"strings"

	"github.com/egoist/mygo/ui"
)

type shortcutCapture struct {
	activeCode uint32
	released   bool
	selectAll  bool
	captured   bool
}

var codeKeyNames = map[ui.KeyCode]string{
	ui.CodeKeyA: "A", ui.CodeKeyB: "B", ui.CodeKeyC: "C", ui.CodeKeyD: "D", ui.CodeKeyE: "E",
	ui.CodeKeyF: "F", ui.CodeKeyG: "G", ui.CodeKeyH: "H", ui.CodeKeyI: "I", ui.CodeKeyJ: "J",
	ui.CodeKeyK: "K", ui.CodeKeyL: "L", ui.CodeKeyM: "M", ui.CodeKeyN: "N", ui.CodeKeyO: "O",
	ui.CodeKeyP: "P", ui.CodeKeyQ: "Q", ui.CodeKeyR: "R", ui.CodeKeyS: "S", ui.CodeKeyT: "T",
	ui.CodeKeyU: "U", ui.CodeKeyV: "V", ui.CodeKeyW: "W", ui.CodeKeyX: "X", ui.CodeKeyY: "Y",
	ui.CodeKeyZ: "Z",
	ui.CodeDigit0: "0", ui.CodeDigit1: "1", ui.CodeDigit2: "2", ui.CodeDigit3: "3", ui.CodeDigit4: "4",
	ui.CodeDigit5: "5", ui.CodeDigit6: "6", ui.CodeDigit7: "7", ui.CodeDigit8: "8", ui.CodeDigit9: "9",
	ui.CodeF1: "F1", ui.CodeF2: "F2", ui.CodeF3: "F3", ui.CodeF4: "F4",
	ui.CodeF5: "F5", ui.CodeF6: "F6", ui.CodeF7: "F7", ui.CodeF8: "F8",
	ui.CodeF9: "F9", ui.CodeF10: "F10", ui.CodeF11: "F11", ui.CodeF12: "F12",
	ui.CodeF13: "F13", ui.CodeF14: "F14", ui.CodeF15: "F15", ui.CodeF16: "F16",
	ui.CodeF17: "F17", ui.CodeF18: "F18", ui.CodeF19: "F19", ui.CodeF20: "F20",
	ui.CodeF21: "F21", ui.CodeF22: "F22", ui.CodeF23: "F23", ui.CodeF24: "F24",
}

var logicalKeyNames = map[ui.Key]string{
	ui.KeyA: "A", ui.KeyB: "B", ui.KeyC: "C", ui.KeyD: "D", ui.KeyE: "E",
	ui.KeyF: "F", ui.KeyG: "G", ui.KeyH: "H", ui.KeyI: "I", ui.KeyJ: "J",
	ui.KeyK: "K", ui.KeyL: "L", ui.KeyM: "M", ui.KeyN: "N", ui.KeyO: "O",
	ui.KeyP: "P", ui.KeyQ: "Q", ui.KeyR: "R", ui.KeyS: "S", ui.KeyT: "T",
	ui.KeyU: "U", ui.KeyV: "V", ui.KeyW: "W", ui.KeyX: "X", ui.KeyY: "Y",
	ui.KeyZ: "Z",
	ui.Key0: "0", ui.Key1: "1", ui.Key2: "2", ui.Key3: "3", ui.Key4: "4",
	ui.Key5: "5", ui.Key6: "6", ui.Key7: "7", ui.Key8: "8", ui.Key9: "9",
	ui.KeyF1: "F1", ui.KeyF2: "F2", ui.KeyF3: "F3", ui.KeyF4: "F4",
	ui.KeyF5: "F5", ui.KeyF6: "F6", ui.KeyF7: "F7", ui.KeyF8: "F8",
	ui.KeyF9: "F9", ui.KeyF10: "F10", ui.KeyF11: "F11", ui.KeyF12: "F12",
}

func keyNameFromInput(ev ui.InputEvent) (string, uint32) {
	if ev.Code != ui.CodeUnknown {
		name := codeKeyNames[ev.Code]
		return name, uint32(ev.Code) | 0x10000
	}
	if ev.PhysicalKey != ui.KeyUnknown {
		name := logicalKeyNames[ev.PhysicalKey]
		return name, uint32(ev.PhysicalKey) | 0x20000
	}
	if ev.Key != ui.KeyUnknown {
		name := logicalKeyNames[ev.Key]
		return name, uint32(ev.Key) | 0x40000
	}
	return "", 0
}

func eventToken(ev ui.InputEvent) uint32 {
	_, tok := keyNameFromInput(ev)
	if tok != 0 {
		return tok
	}
	if ev.Code != ui.CodeUnknown {
		return uint32(ev.Code) | 0x10000
	}
	if ev.PhysicalKey != ui.KeyUnknown {
		return uint32(ev.PhysicalKey) | 0x20000
	}
	if ev.Key != ui.KeyUnknown {
		return uint32(ev.Key) | 0x40000
	}
	return 0
}

func chordFromInput(ev ui.InputEvent) (string, bool) {
	if ev.Mods&ui.Super != 0 || ev.Mods&(ui.Ctrl|ui.Alt) == 0 {
		return "", false
	}
	name, _ := keyNameFromInput(ev)
	if name == "" {
		return "", false
	}
	var parts []string
	if ev.Mods&ui.Ctrl != 0 {
		parts = append(parts, "Ctrl")
	}
	if ev.Mods&ui.Alt != 0 {
		parts = append(parts, "Alt")
	}
	if ev.Mods&ui.Shift != 0 {
		parts = append(parts, "Shift")
	}
	parts = append(parts, name)
	return strings.Join(parts, "+"), true
}

func shortcutField(c *ui.Context, value *string, cap *shortcutCapture, label, tooltip string, invalid bool) bool {
	t := c.Theme()
	changed := cap.captured
	cap.captured = false

	in := ui.TextInput(c, value).
		Width(184).
		Padding(t.Space(1.5), t.Space(2.5)).
		FontSize(13.5).
		MaxLength(128).
		Label(label).
		Tooltip(tooltip).
		Invalid(invalid)

	if in.Changed() {
		changed = true
	}

	if !in.Focused() {
		cap.activeCode = 0
		cap.released = false
	} else if cap.selectAll {
		in.SetTextSelection(0, math.MaxInt)
		cap.selectAll = false
	}

	in.HandleInput(func(ev ui.InputEvent) bool {
		switch ev.Kind {
		case ui.InputKeyDown:
			if ev.Key == ui.KeyTab && ev.Mods&(ui.Ctrl|ui.Alt|ui.Super) == 0 {
				cap.activeCode = 0
				cap.released = false
				return false
			}
			if cap.activeCode != 0 {
				if cap.released && !ev.Repeat && eventToken(ev) != cap.activeCode {
					cap.activeCode = 0
					cap.released = false
				} else {
					if eventToken(ev) == cap.activeCode {
						cap.released = false
					}
					return true
				}
			}
			chord, ok := chordFromInput(ev)
			if !ok {
				return false
			}
			cap.activeCode = eventToken(ev)
			cap.released = false
			if !ev.Repeat {
				*value = chord
				cap.selectAll = true
				cap.captured = true
			}
			return true
		case ui.InputKeyUp:
			if cap.activeCode == 0 {
				return false
			}
			if eventToken(ev) == cap.activeCode {
				cap.released = true
			}
			if cap.released && ev.Mods&(ui.Ctrl|ui.Alt|ui.Shift|ui.Super) == 0 {
				cap.activeCode = 0
				cap.released = false
			}
			return true
		}
		return false
	})

	return changed
}
