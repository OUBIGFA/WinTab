using System.Collections.Concurrent;
using WinTab.Helpers;

namespace WinTab.Models;

internal sealed class WindowTracker
{
    private readonly ConcurrentDictionary<nint, WindowLifecycleState> _states = new();
    private readonly ConcurrentDictionary<nint, long> _lastHideAttempt = new();
    private readonly ConcurrentDictionary<nint, long> _lastShowAttempt = new();

    public WindowLifecycleState GetState(nint hWnd)
    {
        return _states.TryGetValue(hWnd, out var state) ? state : WindowLifecycleState.Unknown;
    }

    public void SetState(nint hWnd, WindowLifecycleState state)
    {
        _states[hWnd] = state;
    }

    public bool ShouldDebounceHide(nint hWnd, int debounceMs)
    {
        var now = StopwatchHelper.GetTimestamp();
        if (_lastHideAttempt.TryGetValue(hWnd, out var last) && !StopwatchHelper.IsTimeUp(last, debounceMs))
            return true;

        _lastHideAttempt[hWnd] = now;
        return false;
    }

    public bool ShouldDebounceShow(nint hWnd, int debounceMs)
    {
        var now = StopwatchHelper.GetTimestamp();
        if (_lastShowAttempt.TryGetValue(hWnd, out var last) && !StopwatchHelper.IsTimeUp(last, debounceMs))
            return true;

        _lastShowAttempt[hWnd] = now;
        return false;
    }

    public void Clear(nint hWnd)
    {
        _states.TryRemove(hWnd, out _);
        _lastHideAttempt.TryRemove(hWnd, out _);
        _lastShowAttempt.TryRemove(hWnd, out _);
    }

    public void ClearAll()
    {
        _states.Clear();
        _lastHideAttempt.Clear();
        _lastShowAttempt.Clear();
    }
}


