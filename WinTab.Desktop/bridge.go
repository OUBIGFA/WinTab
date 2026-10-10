package main

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"sync"
)

const maxMessage = 65536

type reply struct {
	ID     uint64          `json:"id"`
	Result json.RawMessage `json:"result"`
	Error  string          `json:"error"`
	Event  string          `json:"event"`
	Data   json.RawMessage `json:"data"`
}

// bridge multiplexes calls over the parent's anonymous pipes. EOF fails every pending call
// and ends the UI, so a renderer never survives as a disconnected settings editor.
type bridge struct {
	mu      sync.Mutex
	writeMu sync.Mutex
	writer  io.Writer
	next    uint64
	pending map[uint64]chan reply
	done    chan struct{}
	err     error
}

func newBridge(writer io.Writer) *bridge {
	return &bridge{writer: writer, pending: make(map[uint64]chan reply), done: make(chan struct{})}
}

func (b *bridge) read(reader io.Reader, event func(string, json.RawMessage)) {
	scanner := bufio.NewScanner(reader)
	scanner.Buffer(make([]byte, 4096), maxMessage)
	for scanner.Scan() {
		var message reply
		if err := json.Unmarshal(scanner.Bytes(), &message); err != nil {
			b.finish(fmt.Errorf("invalid parent message: %w", err))
			return
		}
		if message.Event != "" {
			event(message.Event, message.Data)
			continue
		}
		b.mu.Lock()
		waiting := b.pending[message.ID]
		delete(b.pending, message.ID)
		b.mu.Unlock()
		if waiting != nil {
			waiting <- message
		}
	}
	err := scanner.Err()
	if err == nil {
		err = io.EOF
	}
	b.finish(err)
}

func (b *bridge) finish(err error) {
	b.mu.Lock()
	defer b.mu.Unlock()
	if b.err != nil {
		return
	}
	b.err = err
	close(b.done)
}

func (b *bridge) call(ctx context.Context, method string, params any, result any) error {
	b.mu.Lock()
	if b.err != nil {
		err := b.err
		b.mu.Unlock()
		return err
	}
	b.next++
	id := b.next
	waiting := make(chan reply, 1)
	b.pending[id] = waiting
	b.mu.Unlock()
	defer func() { b.mu.Lock(); delete(b.pending, id); b.mu.Unlock() }()
	request := struct {
		ID     uint64 `json:"id"`
		Method string `json:"method"`
		Params any    `json:"params"`
	}{id, method, params}
	data, err := json.Marshal(request)
	if err != nil {
		return err
	}
	if len(data) >= maxMessage {
		return errors.New("request is too large")
	}
	b.writeMu.Lock()
	_, err = b.writer.Write(append(data, '\n'))
	b.writeMu.Unlock()
	if err != nil {
		b.finish(err)
		return err
	}
	select {
	case message := <-waiting:
		if message.Error != "" {
			return errors.New(message.Error)
		}
		if result == nil {
			return nil
		}
		return json.Unmarshal(message.Result, result)
	case <-ctx.Done():
		return ctx.Err()
	case <-b.done:
		b.mu.Lock()
		err := b.err
		b.mu.Unlock()
		return err
	}
}
