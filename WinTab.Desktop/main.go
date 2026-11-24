package main

import (
	"context"
	"encoding/json"
	"errors"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/egoist/mygo"
	"github.com/egoist/mygo/ui"
)

var version = "development"

// Desktop drives the native MyGo UI window while forwarding persisted settings and
// shell commands over anonymous stdio pipes to the resident WinTab process.
type Desktop struct {
	pipe  *bridge
	mu    sync.Mutex
	state State
	win   *mygo.Window // read and written only on MyGo's main thread
	ui    *uiModel     // read and written only on MyGo's main thread
	ready atomic.Bool
}

func (d *Desktop) accept(state State) {
	d.mu.Lock()
	if state.Revision < d.state.Revision {
		d.mu.Unlock()
		return
	}
	themeChanged := state.Settings.Theme != d.state.Settings.Theme
	d.state = state
	win := d.win
	uiModel := d.ui
	d.mu.Unlock()

	if win != nil && uiModel != nil {
		win.Update(func() {
			d.mu.Lock()
			latest := d.state
			d.mu.Unlock()
			uiModel.applyState(latest)
		})
	}
	if !themeChanged {
		return
	}
	// Replies and pushed events may overlap. Read the newest preference on the native
	// thread so an older queued update cannot repaint over a later theme selection.
	mygo.RunOnMain(func() {
		d.mu.Lock()
		theme := d.state.Settings.Theme
		d.mu.Unlock()
		if source := themeSource(theme); mygo.Theme.Source() != source {
			mygo.Theme.SetSource(source)
		}
	})
}

// themeSource maps the saved preference to the appearance MyGo applies to the window.
func themeSource(theme string) mygo.ThemeSource {
	switch theme {
	case "Dark":
		return mygo.ThemeDark
	case "System":
		return mygo.ThemeSystem
	}
	return mygo.ThemeLight
}

func (d *Desktop) request(ctx context.Context, method string, args any) (State, error) {
	ctx, cancel := context.WithTimeout(ctx, 90*time.Second)
	defer cancel()
	var state State
	if err := d.pipe.call(ctx, method, args, &state); err != nil {
		return State{}, err
	}
	if state.ProtocolVersion != 1 {
		return State{}, errors.New("incompatible WinTab bridge version")
	}
	d.accept(state)
	return state, nil
}

func (d *Desktop) runAction(key string, call func(context.Context) (State, error), onSuccess func()) {
	go func() {
		state, err := call(context.Background())
		d.mu.Lock()
		win := d.win
		m := d.ui
		d.mu.Unlock()
		if win == nil || m == nil {
			return
		}
		win.Update(func() {
			delete(m.pending, key)
			if err != nil {
				m.actionError = true
				return
			}
			m.applyState(state)
			if onSuccess != nil {
				onSuccess()
			}
		})
	}()
}

func (d *Desktop) Set(key string, value any) {
	d.runAction(key, func(ctx context.Context) (State, error) {
		return d.request(ctx, "set", map[string]any{"key": key, "value": value})
	}, nil)
}

func (d *Desktop) Shortcuts(groupEnabled bool, group string, tabEnabled bool, tab string, notifyApplied bool) {
	d.runAction("shortcuts", func(ctx context.Context) (State, error) {
		return d.request(ctx, "shortcuts", map[string]any{
			"groupEnabled": groupEnabled,
			"group":        group,
			"tabEnabled":   tabEnabled,
			"tab":          tab,
		})
	}, func() {
		if d.ui.state != nil && (d.ui.state.ShortcutError == nil || *d.ui.state.ShortcutError == "") && notifyApplied {
			d.ui.shortcutApplied = true
		}
	})
}

func (d *Desktop) Restore(group bool) {
	d.runAction("restore", func(ctx context.Context) (State, error) {
		return d.request(ctx, "restore", map[string]bool{"group": group})
	}, nil)
}

func (d *Desktop) CheckUpdates() {
	d.runAction("update", func(ctx context.Context) (State, error) {
		return d.request(ctx, "update", nil)
	}, nil)
}

func (d *Desktop) OpenLogs() {
	d.runAction("logs", func(ctx context.Context) (State, error) {
		return d.request(ctx, "logs", nil)
	}, nil)
}

func (d *Desktop) OpenProject() {
	if err := mygo.Shell.OpenExternal("https://github.com/OUBIGFA/WinTab"); err != nil && d.ui != nil {
		d.ui.actionError = true
	}
}

func (d *Desktop) Reconnect() {
	go func() {
		state, err := d.request(context.Background(), "state", nil)
		d.mu.Lock()
		win := d.win
		m := d.ui
		d.mu.Unlock()
		if win == nil || m == nil {
			return
		}
		win.Update(func() {
			if err != nil {
				m.loadError = true
				return
			}
			m.applyState(state)
		})
	}()
}

func (d *Desktop) open(state State) {
	d.ui = newUIModel(&state, d)
	d.accept(state)

	area := mygo.Screen.DisplayNearestPoint(mygo.Screen.CursorScreenPoint()).WorkArea
	width, height := initialSize(state.Settings.FormSize, area.Width, area.Height)
	bg := "#ffffff"
	// accept has already applied the preference, so the effective appearance includes the system choice.
	if mygo.Theme.IsDark() {
		bg = "#161616"
	}

	win := mygo.NewWindow(mygo.WindowOptions{
		Title:           "WinTab",
		Width:           width,
		Height:          height,
		X:               area.X + (area.Width-width)/2,
		Y:               area.Y + (area.Height-height)/2,
		MinWidth:        min(640, area.Width),
		MinHeight:       min(480, area.Height),
		TitleBarStyle:   mygo.TitleBarHidden,
		TitleBarHeight:  44,
		BackgroundColor: bg,
		Content:         ui.View(d.ui.view),
	})
	d.mu.Lock()
	d.win = win
	d.mu.Unlock()

	if err := win.SetIcon(logoPNG); err != nil {
		log.Printf("set window icon: %v", err)
	}
	d.ready.Store(true)
	trackUserResize(win, func(size Size) {
		if !d.ready.Load() {
			return
		}
		go func() {
			ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
			defer cancel()
			if _, err := d.request(ctx, "size", size); err != nil {
				log.Printf("save window size: %v", err)
			}
		}()
	})
}

func initialSize(saved *Size, width, height int) (int, int) {
	if saved != nil {
		return min(max(int(saved.Width), 400), width), min(max(int(saved.Height), 300), height)
	}
	return min(1064, width), min(780, height)
}

func main() {
	log.SetOutput(os.Stderr) // stdout is exclusively the protocol, never diagnostic text
	if len(os.Args) != 2 || os.Args[1] != "--bridge" {
		// A direct launch must go through the resident application's existing single-instance gate.
		exe, err := os.Executable()
		if err == nil {
			err = exec.Command(filepath.Join(filepath.Dir(exe), "WinTab.exe")).Start()
		}
		if err != nil {
			log.Print(err)
			os.Exit(1)
		}
		return
	}

	mygo.App.SetName("WinTab")
	mygo.App.SetVersion(version)
	desktop := &Desktop{pipe: newBridge(os.Stdout)}

	mygo.App.WhenReady(func() {
		go desktop.pipe.read(os.Stdin, func(event string, data json.RawMessage) {
			switch event {
			case "state":
				var state State
				if err := json.Unmarshal(data, &state); err != nil {
					desktop.pipe.finish(err)
					return
				}
				desktop.accept(state)
			case "show":
				mygo.RunOnMain(func() {
					if desktop.win == nil || !desktop.ready.Load() {
						return
					}
					if desktop.win.IsMinimized() {
						desktop.win.Restore()
					}
					desktop.win.Show()
					desktop.win.Focus()
				})
			}
		})
		go func() {
			<-desktop.pipe.done
			mygo.App.Quit()
		}()
		go func() {
			ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
			defer cancel()
			state, err := desktop.request(ctx, "state", nil)
			if err != nil {
				log.Printf("connect to WinTab: %v", err)
				mygo.App.Exit(1)
				return
			}
			mygo.RunOnMain(func() { desktop.open(state) })
		}()
	})

	if err := mygo.App.Run(); err != nil {
		log.Fatal(strings.TrimSpace(err.Error()))
	}
}
