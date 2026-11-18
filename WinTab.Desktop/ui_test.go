package main

import (
	"encoding/json"
	"fmt"
	"image/png"
	"math"
	"os"
	"path/filepath"
	"testing"

	"github.com/egoist/mygo/ui"
)

type fakeActions struct {
	model         *uiModel
	calls         []string
	rejectMethod  string
	onSet         func(st *State, key string, value any)
	applyManually bool
}

func loadFixtureState(t testing.TB) State {
	t.Helper()
	data, err := os.ReadFile("testdata/state.json")
	if err != nil {
		t.Fatal(err)
	}
	var st State
	if err := json.Unmarshal(data, &st); err != nil {
		t.Fatal(err)
	}
	return st
}

func newFakeHarness(t testing.TB, lang, theme string) (*uiModel, *fakeActions) {
	t.Helper()
	st := loadFixtureState(t)
	st.Version = "2.5.1"
	if lang != "" {
		st.Settings.Language = lang
	}
	if theme != "" {
		st.Settings.Theme = theme
	}
	fa := &fakeActions{}
	fa.model = newUIModel(&st, fa)
	return fa.model, fa
}

func (f *fakeActions) Set(key string, value any) {
	f.calls = append(f.calls, fmt.Sprintf("set:%s=%v", key, value))
	delete(f.model.pending, key)
	if f.rejectMethod == "set" {
		f.model.actionError = true
		return
	}
	next := *f.model.state
	next.Revision++
	switch key {
	case "windowHook":
		next.Settings.WindowHook = value.(bool)
		if !next.Settings.WindowHook {
			next.Settings.ReuseTabs = false
		}
	case "reuseTabs":
		next.Settings.ReuseTabs = value.(bool)
		if next.Settings.ReuseTabs {
			next.Settings.WindowHook = true
		}
	case "restoreTabs":
		next.Settings.RestoreTabs = value.(bool)
	case "restoreSingleTab":
		next.Settings.RestoreSingleTab = value.(bool)
	case "restoreOnAnyFolder":
		next.Settings.RestoreOnAnyFolder = value.(bool)
	case "reopenClosedTab":
		next.Settings.ReopenClosedTab = value.(bool)
	case "doubleClickCloseTab":
		next.Settings.DoubleClickCloseTab = value.(bool)
	case "doubleClickCloseIncludeNotepad":
		next.Settings.DoubleClickCloseIncludeNotepad = value.(bool)
	case "doubleClickEmptyArea":
		next.Settings.DoubleClickEmptyArea = value.(bool)
	case "doubleClickEmptyAreaAction":
		next.Settings.DoubleClickEmptyAreaAction = value.(string)
	case "middleClickForegroundTab":
		next.Settings.MiddleClickForegroundTab = value.(bool)
	case "wheelSwitchTab":
		next.Settings.WheelSwitchTab = value.(bool)
	case "wheelSwitchSensitivity":
		next.Settings.WheelSwitchSensitivity = value.(string)
	case "autoUpdate":
		next.Settings.AutoUpdate = value.(bool)
	case "showTrayIcon":
		next.Settings.ShowTrayIcon = value.(bool)
	case "startup":
		next.Startup = value.(bool)
	case "language":
		next.Settings.Language = value.(string)
	case "theme":
		next.Settings.Theme = value.(string)
	}
	if f.onSet != nil {
		f.onSet(&next, key, value)
	}
	next.RecordClosedTabs = next.Settings.ReopenClosedTab || next.Settings.ReopenTabShortcutEnabled
	f.model.applyState(next)
}

func (f *fakeActions) Shortcuts(ge bool, g string, te bool, tb string, notify bool) {
	f.calls = append(f.calls, fmt.Sprintf("shortcuts:%v/%s|%v/%s", ge, g, te, tb))
	delete(f.model.pending, "shortcuts")
	next := *f.model.state
	next.Revision++
	if g == tb {
		errMsg := "两个恢复功能不能使用相同快捷键"
		next.ShortcutError = &errMsg
	} else {
		next.ShortcutError = nil
		next.Settings.RestoreGroupShortcutEnabled = ge
		next.Settings.RestoreGroupShortcut = g
		next.Settings.ReopenTabShortcutEnabled = te
		next.Settings.ReopenTabShortcut = tb
		if notify {
			f.model.shortcutApplied = true
		}
	}
	next.RecordClosedTabs = next.Settings.ReopenClosedTab || next.Settings.ReopenTabShortcutEnabled
	f.model.applyState(next)
}

func (f *fakeActions) Restore(group bool) {
	f.calls = append(f.calls, fmt.Sprintf("restore:%v", group))
	delete(f.model.pending, "restore")
	next := *f.model.state
	next.Revision++
	msg := "恢复完成"
	next.SessionFeedback = &msg
	f.model.applyState(next)
}

func (f *fakeActions) CheckUpdates() {
	f.calls = append(f.calls, "update")
	delete(f.model.pending, "update")
	next := *f.model.state
	next.Revision++
	msg := "已是最新版本"
	next.UpdateFeedback = &msg
	f.model.applyState(next)
}

func (f *fakeActions) OpenLogs()    { f.calls = append(f.calls, "logs"); delete(f.model.pending, "logs") }
func (f *fakeActions) OpenProject() { f.calls = append(f.calls, "openProject") }
func (f *fakeActions) Reconnect() {
	st := loadFixtureState(new(testing.T))
	f.model.applyState(st)
}

func TestAllSectionsNavigationAccessibilityAndAlignment(t *testing.T) {
	for _, lang := range []string{"zh-CN", "en-US"} {
		for _, theme := range []string{"Light", "Dark"} {
			m, _ := newFakeHarness(t, lang, theme)
			cp := copyForLanguage(lang)
			tt := ui.NewTesterWith(m.view, 1064, 780, ui.TesterOptions{Dark: theme == "Dark"})

			for _, nav := range m.sections(cp) {
				if err := tt.Click(nav.label); err != nil {
					t.Fatalf("%s/%s click %q: %v", lang, theme, nav.label, err)
				}
				h1, ok := tt.AccessByRole(ui.RoleHeading, nav.label)
				if !ok || h1.Bounds.W == 0 {
					t.Fatalf("%s/%s missing H1 for %q", lang, theme, nav.label)
				}
				for _, acc := range tt.AccessAll() {
					if (acc.Role == ui.RoleButton || acc.Role == ui.RoleSwitch || acc.Role == ui.RoleCheckBox || acc.Role == ui.RoleTextField) && acc.Name == "" {
						t.Errorf("%s/%s section %s: interactive role %v at %+v has empty name", lang, theme, nav.id, acc.Role, acc.Bounds)
					}
				}
			}

			// Verify wide desktop centering symmetry at 1440x900
			tt.SetSize(1440, 900)
			if err := tt.Click(cp.Explorer); err != nil {
				t.Fatal(err)
			}
			h1, ok := tt.AccessByRole(ui.RoleHeading, cp.Explorer)
			if !ok {
				t.Fatal("missing heading at 1440x900")
			}
			// Sidebar is 224 DIP, workspace = 1216 DIP, MaxWidth = 840 DIP, inner padding = 36 DIP:
			// page box left inside workspace = (1216 - 840)/2 = 188; heading starts at 224 + 188 + 36 = 448.
			wantLeft := float32(224 + (1216-840)/2 + 36)
			if math.Abs(float64(h1.Bounds.X-wantLeft)) > 2 {
				t.Errorf("%s/%s 1440x900 heading X=%.1f, want centered near %.1f", lang, theme, h1.Bounds.X, wantLeft)
			}

			// Verify compact width switches cleanly to RoleTab strip without missing navigation
			tt.SetSize(640, 700)
			if _, ok := tt.AccessByRole(ui.RoleTab, cp.RecoveryNav); !ok {
				t.Errorf("%s/%s compact 640x700 missing compact tab %q", lang, theme, cp.RecoveryNav)
			}
		}
	}
}

func TestBackendAuthoritativeSettingsAndDraftPreservation(t *testing.T) {
	m, fa := newFakeHarness(t, "zh-CN", "Light")
	tt := ui.NewTester(m.view, 1064, 780)

	// Disabling windowHook clears reuseTabs on backend.
	if err := tt.Click("合并新窗口"); err != nil {
		t.Fatal(err)
	}
	if sw, ok := tt.AccessByRole(ui.RoleSwitch, "复用已有标签"); !ok || sw.Checked {
		t.Fatalf("expected reuseTabs unchecked after windowHook off: %+v", sw)
	}

	// Enabling reuseTabs enables windowHook on backend.
	if err := tt.Click("复用已有标签"); err != nil {
		t.Fatal(err)
	}
	if sw, ok := tt.AccessByRole(ui.RoleSwitch, "合并新窗口"); !ok || !sw.Checked {
		t.Fatalf("expected windowHook checked after reuseTabs on: %+v", sw)
	}

	// Backend rejection reverts visual switch state and shows RoleAlert.
	fa.rejectMethod = "set"
	if err := tt.Click("合并新窗口"); err != nil {
		t.Fatal(err)
	}
	if sw, ok := tt.AccessByRole(ui.RoleSwitch, "合并新窗口"); !ok || !sw.Checked {
		t.Fatalf("expected switch to stay checked after backend error: %+v", sw)
	}
	if alert, ok := tt.AccessByRole(ui.RoleAlert, ""); !ok || alert.Name == "" {
		t.Fatalf("expected RoleAlert after rejected set")
	}
	fa.rejectMethod = ""

	// Navigation & theme changes preserve unapplied shortcut drafts and duplicate detection.
	if err := tt.Click("会话恢复"); err != nil {
		t.Fatal(err)
	}
	if err := tt.Click("标签组恢复快捷键"); err != nil {
		t.Fatal(err)
	}
	tt.CodeDown(ui.Ctrl|ui.Shift, ui.CodeKeyK)
	tt.CodeUp(0, ui.CodeKeyK)
	if m.groupDraft != "Ctrl+Shift+K" {
		t.Fatalf("expected captured draft Ctrl+Shift+K, got %q", m.groupDraft)
	}
	if err := tt.Click("外观与系统"); err != nil {
		t.Fatal(err)
	}
	if err := tt.Click("切换为深色主题"); err != nil {
		t.Fatal(err)
	}
	if err := tt.Click("会话恢复"); err != nil {
		t.Fatal(err)
	}
	if m.groupDraft != "Ctrl+Shift+K" || !tt.HasText("有未应用的快捷键") {
		t.Fatalf("draft lost or indicator missing: draft=%q", m.groupDraft)
	}

	// Duplicate shortcut Alt+W triggers alert and preserves input value.
	m.groupDraft = "Alt+W"
	tt.Frame()
	if err := tt.Click("应用快捷键"); err != nil {
		t.Fatal(err)
	}
	if alert, ok := tt.AccessByRole(ui.RoleAlert, ""); !ok || !tt.HasText("不能使用相同快捷键") {
		t.Fatalf("expected duplicate shortcut alert, got %+v", alert)
	}
}

func TestSessionRecoveryZeroVerticalJump(t *testing.T) {
	for _, lang := range []string{"zh-CN", "en-US"} {
		cp := copyForLanguage(lang)
		m, _ := newFakeHarness(t, lang, "Light")
		tt := ui.NewTester(m.view, 900, 1050)
		if err := tt.Click(cp.RecoveryNav); err != nil {
			t.Fatal(err)
		}

		snapshot := func() map[string]float32 {
			applyBtn, ok1 := tt.AccessByRole(ui.RoleButton, cp.Apply)
			autoHeading, ok2 := tt.AccessByRole(ui.RoleHeading, cp.Automatic)
			detailsBtn, ok3 := tt.AccessByRole(ui.RoleDisclosure, cp.RecoveryDetails)
			if !ok1 || !ok2 || !ok3 {
				t.Fatalf("%s missing anchors: %v %v %v", lang, ok1, ok2, ok3)
			}
			return map[string]float32{
				"applyTop":   applyBtn.Bounds.Y,
				"autoTop":    autoHeading.Bounds.Y,
				"detailsTop": detailsBtn.Bounds.Y,
			}
		}

		baseline := snapshot()
		assertZeroShift := func(step string) {
			t.Helper()
			cur := snapshot()
			for k, base := range baseline {
				if math.Abs(float64(cur[k]-base)) > 0.5 {
					t.Fatalf("%s step %q shifted %s from %.1f to %.1f", lang, step, k, base, cur[k])
				}
			}
		}

		clickRole := func(role ui.Role, name string) {
			info, ok := tt.AccessByRole(role, name)
			if !ok {
				t.Fatalf("%s missing %v %q", lang, role, name)
			}
			tt.ClickAt(info.Bounds.X+info.Bounds.W/2, info.Bounds.Y+info.Bounds.H/2)
		}

		clickRole(ui.RoleCheckBox, cp.EnableGroup)
		assertZeroShift("uncheck enable-group")
		clickRole(ui.RoleCheckBox, cp.EnableGroup)
		assertZeroShift("recheck enable-group")
		clickRole(ui.RoleCheckBox, cp.EnableTab)
		assertZeroShift("uncheck enable-tab")
		clickRole(ui.RoleCheckBox, cp.EnableTab)
		assertZeroShift("recheck enable-tab")

		// Edit draft -> apply -> uncheck again: zero height shift across all states.
		clickRole(ui.RoleTextField, cp.GroupShortcut)
		tt.CodeDown(ui.Ctrl, ui.CodeKeyK)
		tt.CodeUp(0, ui.CodeKeyK)
		assertZeroShift("draft shortcut")
		clickRole(ui.RoleButton, cp.Apply)
		if !tt.HasText(cp.Applied) {
			t.Fatalf("%s expected applied text", lang)
		}
		assertZeroShift("applied shortcut")
		clickRole(ui.RoleCheckBox, cp.EnableGroup)
		assertZeroShift("checkbox after applied")

		// Auto-restore switch toggle: zero height shift.
		clickRole(ui.RoleSwitch, cp.AutoRestore)
		assertZeroShift("auto-restore on")
		clickRole(ui.RoleSwitch, cp.AutoRestore)
		assertZeroShift("auto-restore off")
	}
}

func TestDarkModeBrandLogoPlateIsPureWhite(t *testing.T) {
	m, _ := newFakeHarness(t, "zh-CN", "Dark")
	tt := ui.NewTesterWith(m.view, 1064, 780, ui.TesterOptions{Dark: true})
	if err := tt.Click("关于"); err != nil {
		t.Fatal(err)
	}
	img := tt.Image()
	if img == nil {
		t.Fatal("rendered image is nil")
	}
	// Sidebar logo box is at X=16..36, Y=12..32 (20x20 padded at x=16,y=12 inside 44px bar).
	// Center padding pixel (e.g. X=18, Y=14) must be pure white (#ffffff).
	c := img.RGBAAt(18, 14)
	if c.R != 255 || c.G != 255 || c.B != 255 {
		t.Fatalf("dark mode sidebar logo plate corner pixel at (18,14) = %+v, want pure white", c)
	}
}

func TestRenderPreviewScreenshots(t *testing.T) {
	if os.Getenv("WINTAB_CAPTURE_PREVIEWS") != "1" {
		return
	}
	outDir := filepath.Join("..", "_temp", "mygo-shots")
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		t.Fatal(err)
	}
	savePNG := func(path string, m *uiModel, dark bool, w, h int, scale float32, tab string) {
		tt := ui.NewTesterWith(m.view, w, h, ui.TesterOptions{Dark: dark})
		tt.SetScale(scale)
		if tab != "" {
			if err := tt.Click(tab); err != nil {
				t.Fatal(err)
			}
		}
		f, err := os.Create(path)
		if err != nil {
			t.Fatal(err)
		}
		defer f.Close()
		if err := png.Encode(f, tt.Image()); err != nil {
			t.Fatal(err)
		}
	}

	mLight, _ := newFakeHarness(t, "zh-CN", "Light")
	savePNG(filepath.Join("..", "Assets", "UI.png"), mLight, false, 1064, 780, 1.5, "")

	for _, entry := range []struct {
		theme string
		dark  bool
	}{{"Light", false}, {"Dark", true}} {
		for _, nav := range []struct{ id, label string }{
			{"explorer", zhCopy.Explorer},
			{"recovery", zhCopy.RecoveryNav},
			{"preferences", zhCopy.Preferences},
			{"about", zhCopy.About},
		} {
			m, _ := newFakeHarness(t, "zh-CN", entry.theme)
			savePNG(filepath.Join(outDir, fmt.Sprintf("%s-%s-1064.png", entry.theme, nav.id)), m, entry.dark, 1064, 780, 1.5, nav.label)
		}
		mWide, _ := newFakeHarness(t, "zh-CN", entry.theme)
		savePNG(filepath.Join(outDir, fmt.Sprintf("%s-wide-1440.png", entry.theme)), mWide, entry.dark, 1440, 900, 1.0, zhCopy.Explorer)
	}
}
