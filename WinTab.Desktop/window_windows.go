//go:build windows

package main

import (
	"syscall"

	"github.com/egoist/mygo"
)

var (
	controls       = syscall.NewLazyDLL("comctl32.dll")
	setSubclass    = controls.NewProc("SetWindowSubclass")
	removeSubclass = controls.NewProc("RemoveWindowSubclass")
	defSubclass    = controls.NewProc("DefSubclassProc")
)

// Match the WPF compatibility window's rule: only a user's sizing gesture persists dimensions.
// Web layout, initial fitting, maximize, DPI changes and a plain move must not overwrite them.
func trackUserResize(win *mygo.Window, save func(Size)) {
	var start Size
	handle := win.NativeHandle()
	callback := syscall.NewCallback(func(hwnd uintptr, message uint32, wparam, lparam, id, data uintptr) uintptr {
		switch message {
		case 0x0231: // WM_ENTERSIZEMOVE
			size := win.Bounds()
			start = Size{float64(size.Width), float64(size.Height)}
		case 0x0232: // WM_EXITSIZEMOVE
			size := win.Bounds()
			current := Size{float64(size.Width), float64(size.Height)}
			if current != start && !win.IsMaximized() && !win.IsMinimized() {
				save(current)
			}
		}
		result, _, _ := defSubclass.Call(hwnd, uintptr(message), wparam, lparam)
		return result
	})
	if ok, _, err := setSubclass.Call(handle, callback, 1, 0); ok == 0 {
		panic(err)
	}
	win.OnClosed(func() { removeSubclass.Call(handle, callback, 1) })
}
