package main

import (
	"context"
	"embed"
	"encoding/json"
	"errors"
	"io/fs"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/egoist/mygo"
)

//go:embed all:frontend/dist
var assets embed.FS

var version = "development"
var stateChanged = mygo.NewEvent[State]("state-changed")
var userResized = mygo.NewEvent[bool]("user-resized")

// Desktop exposes only application settings and fixed actions to the bundled page.
// External pages, arbitrary URLs, file paths and shell commands are never accepted.
type Desktop struct {
	pipe         *bridge
	mu           sync.Mutex
	state        State
	win          *mygo.Window // read and written only on MyGo's main thread
	ready        atomic.Bool
	hasSavedSize bool
}

func (d *Desktop) accept(state State) {
	d.mu.Lock()
	if state.Revision < d.state.Revision {
		d.mu.Unlock()
		return
	}
	themeChanged := state.Settings.Theme != d.state.Settings.Theme
	d.state = state
	d.mu.Unlock()
	if !themeChanged {
		return
	}
	// Replies and pushed events may overlap. Read the newest preference on the native
	// thread so an older queued update cannot repaint over a later theme selection.
	mygo.RunOnMain(func() {
		d.mu.Lock()
		theme := d.state.Settings.Theme
		d.mu.Unlock()
		source := mygo.ThemeLight
		if theme == "Dark" {
			source = mygo.ThemeDark
		}
		if mygo.Theme.Source() != source {
			mygo.Theme.SetSource(source)
		}
	})
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

// Load returns the current persisted state and the work area used for the first-launch fit.
func (d *Desktop) Load(ctx context.Context) (Bootstrap, error) {
	state, err := d.request(ctx, "state", nil)
	if err != nil {
		return Bootstrap{}, err
	}
	var info WindowInfo
	mygo.RunOnMain(func() {
		area := mygo.Screen.DisplayMatching(d.win.Bounds()).WorkArea
		size := d.win.ContentBounds()
		outer := d.win.Bounds()
		info = WindowInfo{size.Width, area.Width - (outer.Width - size.Width),
			area.Height - (outer.Height - size.Height), d.hasSavedSize}
	})
	return Bootstrap{state, info}, nil
}

func (d *Desktop) Set(ctx context.Context, key string, value any) (State, error) {
	return d.request(ctx, "set", map[string]any{"key": key, "value": value})
}

func (d *Desktop) Shortcuts(ctx context.Context, groupEnabled bool, group string, tabEnabled bool, tab string) (State, error) {
	return d.request(ctx, "shortcuts", map[string]any{
		"groupEnabled": groupEnabled, "group": group, "tabEnabled": tabEnabled, "tab": tab,
	})
}

func (d *Desktop) Restore(ctx context.Context, group bool) (State, error) {
	return d.request(ctx, "restore", map[string]bool{"group": group})
}

func (d *Desktop) Update(ctx context.Context) (State, error) { return d.request(ctx, "update", nil) }
func (d *Desktop) Logs(ctx context.Context) (State, error)   { return d.request(ctx, "logs", nil) }
func (*Desktop) OpenProject() error {
	return mygo.Shell.OpenExternal("https://github.com/OUBIGFA/WinTab")
}

// Ready is called after fonts and the real settings have been laid out, not at the first blank paint.
// Manual dimensions remain authoritative; initial content fitting never writes the user's size setting.
func (d *Desktop) Ready(width, height int) error {
	if width < 300 || height < 240 || width > 10000 || height > 10000 {
		return errors.New("invalid content size")
	}
	mygo.RunOnMain(func() {
		if d.ready.Load() {
			return
		}
		if !d.hasSavedSize {
			area := mygo.Screen.DisplayMatching(d.win.Bounds()).WorkArea
			outer, inner := d.win.Bounds(), d.win.ContentBounds()
			d.win.SetContentSize(min(width, area.Width-(outer.Width-inner.Width)), min(height, area.Height-(outer.Height-inner.Height)))
			d.win.Center()
		}
		d.ready.Store(true)
		d.win.Show()
		d.win.Focus()
	})
	return nil
}

func (d *Desktop) open(state State) {
	d.accept(state)
	area := mygo.Screen.DisplayNearestPoint(mygo.Screen.CursorScreenPoint()).WorkArea
	width, height := initialSize(state.Settings.FormSize, area.Width, area.Height)
	d.hasSavedSize = state.Settings.FormSize != nil
	d.win = mygo.NewWindow(mygo.WindowOptions{
		Title: "WinTab", URL: "/", Width: width, Height: height,
		X: area.X + (area.Width-width)/2, Y: area.Y + (area.Height-height)/2,
		MinWidth: min(640, area.Width), MinHeight: min(480, area.Height), Hidden: true,
		TitleBarStyle: mygo.TitleBarHidden, TitleBarHeight: 40,
		BackgroundColor: "light-dark(#fafafa, #141414)",
		Page:            mygo.PageOptions{DevTools: mygo.DevToolsDisabled},
	})
	web, _ := fs.Sub(assets, "frontend/dist")
	icons, _ := fs.Glob(web, "assets/wintab-logo-*.png")
	if len(icons) == 1 {
		if icon, err := fs.ReadFile(web, icons[0]); err == nil {
			d.win.SetIcon(icon)
		}
	}
	d.win.Page().OnWillNavigate(func(event *mygo.NavigateEvent) {
		// No browsing inside a privileged window, even when an accidental link reaches the page.
		event.PreventDefault()
	})
	trackUserResize(d.win, func(size Size) {
		if !d.ready.Load() {
			return
		}
		d.hasSavedSize = true
		_ = userResized.Broadcast(true)
		go func() {
			ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
			defer cancel()
			if _, err := d.request(ctx, "size", size); err != nil {
				log.Printf("save window size: %v", err)
			}
		}()
	})
	go func() {
		time.Sleep(25 * time.Second)
		if !d.ready.Load() {
			log.Print("settings page did not become ready")
			mygo.App.Exit(2)
		}
	}()
}

func initialSize(saved *Size, width, height int) (int, int) {
	if saved != nil {
		return min(max(int(saved.Width), 400), width), min(max(int(saved.Height), 300), height)
	}
	return min(1224, width), min(880, height)
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
	web, err := fs.Sub(assets, "frontend/dist")
	if err != nil {
		log.Fatal(err)
	}
	mygo.App.SetName("WinTab")
	mygo.App.SetVersion(version)
	if cache, err := os.UserCacheDir(); err == nil {
		mygo.App.SetPath(mygo.PathUserData, filepath.Join(cache, "WinTab", "WebView2"))
	}
	mygo.SetFrontend(web)
	desktop := &Desktop{pipe: newBridge(os.Stdout)}
	mygo.Bind(desktop)
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
				_ = stateChanged.Broadcast(state)
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
