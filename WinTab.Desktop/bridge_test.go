package main

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"io"
	"os"
	"sync"
	"testing"
	"time"
)

func TestWireState(t *testing.T) {
	data, err := os.ReadFile("frontend/tests/fixtures/state.json")
	if err != nil {
		t.Fatal(err)
	}
	var state State
	if err := json.Unmarshal(data, &state); err != nil {
		t.Fatal(err)
	}
	if state.ProtocolVersion != 1 || state.Settings.RestoreGroupShortcut != "Alt+E" || !state.RecordClosedTabs || state.Settings.RestoreTabs {
		t.Fatalf("wire defaults lost: %#v", state)
	}
}

func TestBridgeRoutesRepliesAndEvents(t *testing.T) {
	requests, writer := io.Pipe()
	replies, parent := io.Pipe()
	defer requests.Close()
	defer writer.Close()
	defer replies.Close()
	defer parent.Close()
	b := newBridge(writer)
	events := make(chan string, 1)
	go b.read(replies, func(name string, _ json.RawMessage) { events <- name })
	go func() {
		var request struct {
			ID     uint64 `json:"id"`
			Method string `json:"method"`
		}
		if err := json.NewDecoder(requests).Decode(&request); err != nil {
			return
		}
		_ = json.NewEncoder(parent).Encode(map[string]any{"event": "state", "data": map[string]int{"revision": 2}})
		_ = json.NewEncoder(parent).Encode(map[string]any{"id": request.ID, "result": map[string]string{"method": request.Method}})
	}()
	ctx, cancel := context.WithTimeout(context.Background(), time.Second)
	defer cancel()
	var result map[string]string
	if err := b.call(ctx, "state", nil, &result); err != nil {
		t.Fatal(err)
	}
	if result["method"] != "state" {
		t.Fatal(result)
	}
	select {
	case event := <-events:
		if event != "state" {
			t.Fatal(event)
		}
	case <-ctx.Done():
		t.Fatal("missing event")
	}
}

func TestBridgeDisconnectFailsPendingCalls(t *testing.T) {
	b := newBridge(io.Discard)
	result := make(chan error, 1)
	go func() { result <- b.call(context.Background(), "state", nil, nil) }()
	b.read(bytes.NewReader(nil), func(string, json.RawMessage) {})
	select {
	case err := <-result:
		if !errors.Is(err, io.EOF) {
			t.Fatal(err)
		}
	case <-time.After(time.Second):
		t.Fatal("pending call survived EOF")
	}
}

func TestBridgeCancellationAndOversize(t *testing.T) {
	b := newBridge(io.Discard)
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if err := b.call(ctx, "state", nil, nil); !errors.Is(err, context.Canceled) {
		t.Fatal(err)
	}
	b.mu.Lock()
	pending := len(b.pending)
	b.mu.Unlock()
	if pending != 0 {
		t.Fatal("canceled call leaked")
	}
	b.read(bytes.NewReader(bytes.Repeat([]byte("x"), maxMessage+1)), func(string, json.RawMessage) {})
	select {
	case <-b.done:
	default:
		t.Fatal("oversized input accepted")
	}
}

func TestBridgeConcurrentReplies(t *testing.T) {
	requests, writer := io.Pipe()
	replies, parent := io.Pipe()
	defer requests.Close()
	defer writer.Close()
	defer replies.Close()
	defer parent.Close()
	b := newBridge(writer)
	go b.read(replies, func(string, json.RawMessage) {})
	go func() {
		decoder := json.NewDecoder(requests)
		for i := 0; i < 20; i++ {
			var request struct {
				ID uint64 `json:"id"`
			}
			if decoder.Decode(&request) != nil {
				return
			}
			_ = json.NewEncoder(parent).Encode(map[string]any{"id": request.ID, "result": true})
		}
	}()
	var group sync.WaitGroup
	for i := 0; i < 20; i++ {
		group.Go(func() {
			ctx, cancel := context.WithTimeout(context.Background(), time.Second)
			defer cancel()
			var result bool
			if err := b.call(ctx, "state", nil, &result); err != nil || !result {
				t.Errorf("reply lost: %v", err)
			}
		})
	}
	group.Wait()
}

func TestInitialSizeRespectsSavedDimensionsAndScreen(t *testing.T) {
	for _, test := range []struct {
		saved                       *Size
		width, height, wantW, wantH int
	}{
		{nil, 1920, 1080, 1224, 880}, {nil, 800, 600, 800, 600},
		{&Size{1100, 1200}, 1920, 1400, 1100, 1200}, {&Size{2200, 1600}, 1280, 960, 1280, 960},
	} {
		width, height := initialSize(test.saved, test.width, test.height)
		if width != test.wantW || height != test.wantH {
			t.Fatalf("got %dx%d, want %dx%d", width, height, test.wantW, test.wantH)
		}
	}
}

func TestSettingsStateUpdateDoesNotWaitForNativeThemeWork(t *testing.T) {
	// A normal setting update must be accepted without dispatching native theme work.
	// No MyGo event loop or window is started, so an unnecessary dispatch would block here.
	d := &Desktop{state: State{Revision: 1, Settings: Settings{Theme: "Light"}}}
	done := make(chan struct{})
	go func() {
		d.accept(State{Revision: 2, Settings: Settings{Theme: "Light", WindowHook: true}})
		d.accept(State{Revision: 2, Settings: Settings{Theme: "Light", WindowHook: true}})
		close(done)
	}()
	select {
	case <-done:
		d.mu.Lock()
		defer d.mu.Unlock()
		if d.state.Revision != 2 || !d.state.Settings.WindowHook {
			t.Fatal("setting update was lost while suppressing redundant theme work")
		}
	case <-time.After(time.Second):
		t.Fatal("an unchanged theme caused a native dispatch instead of accepting the settings state")
	}
}
