using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WinTab.Helpers;
using WinTab.Models;
using WinTab.WinAPI;

namespace WinTab.Hooks;

internal sealed class GeneralWindowHost : IWindowHost
{
    private readonly ConcurrentDictionary<nint, WindowEntry> _windows = new();
    private nint _eventForegroundHook;
    private nint _eventCreateHook;
    private nint _eventDestroyHook;
    private WinEventDelegate? _foregroundCallback;
    private WinEventDelegate? _createCallback;
    private WinEventDelegate? _destroyCallback;
    private volatile bool _isReady;

    public WindowHostType HostType => WindowHostType.General;
    public bool IsReady => _isReady;

    public event Action<WindowEntry>? WindowCreated;
    public event Action<nint>? WindowDestroyed;
    public event Action<nint>? WindowActivated;

    public void Start()
    {
        if (_isReady) return;

        _foregroundCallback = OnForegroundChanged;
        _createCallback = OnWindowCreated;
        _destroyCallback = OnWindowDestroyed;

        _eventForegroundHook = WinApi.SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND, 0, _foregroundCallback, 0, 0, 0);
        _eventCreateHook = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_CREATE, WinApi.EVENT_OBJECT_CREATE, 0, _createCallback, 0, 0, 0);
        _eventDestroyHook = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_DESTROY, WinApi.EVENT_OBJECT_DESTROY, 0, _destroyCallback, 0, 0, 0);

        RefreshSnapshot();
        _isReady = true;
    }

    public void Stop()
    {
        if (!_isReady) return;

        WinApi.UnhookWinEvent(_eventForegroundHook);
        WinApi.UnhookWinEvent(_eventCreateHook);
        WinApi.UnhookWinEvent(_eventDestroyHook);
        _eventForegroundHook = 0;
        _eventCreateHook = 0;
        _eventDestroyHook = 0;
        _foregroundCallback = null;
        _createCallback = null;
        _destroyCallback = null;
        _windows.Clear();
        _isReady = false;
    }

    public IReadOnlyCollection<WindowEntry> GetWindows() => _windows.Values.ToList();

    public bool TryActivate(nint hWnd)
    {
        if (!IsCandidateWindow(hWnd)) return false;
        WinApi.RestoreWindowToForeground(hWnd);
        return true;
    }

    public void Dispose()
    {
        Stop();
    }

    private void RefreshSnapshot()
    {
        try
        {
            WinApi.EnumWindows(EnumWindowProc, 0);
        }
        catch (Exception ex)
        {
            Telemetry.ThrottledWarn("general_enum_fail", ex.GetType().Name);
        }
    }

    private bool EnumWindowProc(nint hWnd, nint lParam)
    {
        if (!IsCandidateWindow(hWnd))
            return true;

        var entry = BuildEntry(hWnd);
        if (entry == null) return true;

        _windows[hWnd] = entry;
        return true;
    }

    private void OnForegroundChanged(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (idObject != 0 || idChild != 0) return;
        if (!IsCandidateWindow(hWnd)) return;

        if (_windows.TryGetValue(hWnd, out var entry))
            entry.LastActivatedAt = DateTime.UtcNow;

        WindowActivated?.Invoke(hWnd);
    }

    private void OnWindowCreated(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (idObject != 0 || idChild != 0) return;
        if (!IsCandidateWindow(hWnd)) return;

        var entry = BuildEntry(hWnd);
        if (entry == null) return;

        _windows[hWnd] = entry;
        WindowCreated?.Invoke(entry);
    }

    private void OnWindowDestroyed(nint hWinEventHook, uint eventType, nint hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (idObject != 0 || idChild != 0) return;

        if (_windows.TryRemove(hWnd, out _))
            WindowDestroyed?.Invoke(hWnd);
    }

    private static bool IsCandidateWindow(nint hWnd)
    {
        if (hWnd == 0) return false;
        if (!WinApi.IsWindowVisible(hWnd)) return false;
        if (WinApi.GetWindow(hWnd, WinApi.GW_OWNER) != 0) return false;

        var exStyle = WinApi.GetWindowLong(hWnd, WinApi.GWL_EXSTYLE);
        if ((exStyle & WinApi.WS_EX_TOOLWINDOW) != 0) return false;

        return true;
    }

    private WindowEntry? BuildEntry(nint hWnd)
    {
        var title = GetWindowTitle(hWnd);
        if (string.IsNullOrWhiteSpace(title)) return null;

        var pid = 0u;
        WinApi.GetWindowThreadProcessId(hWnd, out pid);
        var path = WinApi.GetProcessPath((int)pid) ?? string.Empty;

        return new WindowEntry
        {
            Handle = hWnd,
            Title = title,
            ProcessPath = path,
            HostType = WindowHostType.General,
            State = WindowLifecycleState.Visible
        };
    }

    private static string GetWindowTitle(nint hWnd)
    {
        var length = WinApi.GetWindowTextLength(hWnd);
        if (length <= 0) return string.Empty;

        var sb = new StringBuilder(length + 1);
        WinApi.GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}


