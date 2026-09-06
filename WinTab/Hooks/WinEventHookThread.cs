using System;
using System.Threading;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>
/// Owns a dedicated STA message-loop thread that hosts the system-wide WinEvent hooks.
/// WinEvent callbacks are only delivered to the thread that installed the hook, so the
/// hook must live on a thread that keeps pumping messages.
/// </summary>
internal sealed class WinEventHookThread : IDisposable
{
    private readonly WinEventDelegate _showCallback;
    private readonly ManualResetEventSlim _started = new();
    private readonly ManualResetEventSlim _stopped = new();
    private Thread? _thread;
    private uint _threadId;
    private nint _foregroundHookId;
    private nint _showHookId;
    private bool _disposed;

    public WinEventHookThread(WinEventDelegate showCallback)
    {
        _showCallback = showCallback;
    }

    public void Start()
    {
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "WinTab Explorer WinEvent Hook"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();

        if (!_started.Wait(2_000))
            ExplorerDebugLog.Write("WinEvent hook thread did not start within 2000ms");
    }

    private void Run()
    {
        try
        {
            _threadId = WinApi.GetCurrentThreadId();
            _ = WinApi.PeekMessage(out _, 0, 0, 0, WinApi.PM_NOREMOVE);

            const uint hookFlags = WinApi.WINEVENT_OUTOFCONTEXT | WinApi.WINEVENT_SKIPOWNPROCESS;
            _foregroundHookId = WinApi.SetWinEventHook(WinApi.EVENT_SYSTEM_FOREGROUND, WinApi.EVENT_SYSTEM_FOREGROUND, 0, _showCallback, 0, 0, hookFlags);
            _showHookId = WinApi.SetWinEventHook(WinApi.EVENT_OBJECT_CREATE, WinApi.EVENT_OBJECT_SHOW, 0, _showCallback, 0, 0, hookFlags);
            ExplorerDebugLog.Write($"WinEvent hook thread started id={_threadId} foregroundHook={_foregroundHookId} showHook={_showHookId}");
            _started.Set();

            while (WinApi.GetMessage(out var message, 0, 0, 0) > 0)
            {
                WinApi.TranslateMessage(ref message);
                WinApi.DispatchMessage(ref message);
            }
        }
        finally
        {
            if (_foregroundHookId != 0)
                WinApi.UnhookWinEvent(_foregroundHookId);

            if (_showHookId != 0)
                WinApi.UnhookWinEvent(_showHookId);

            ExplorerDebugLog.Write("WinEvent hook thread stopped");
            _stopped.Set();
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;

        if (_threadId != 0)
            WinApi.PostThreadMessage(_threadId, WinApi.WM_QUIT, 0, 0);

        if (!_stopped.Wait(2_000))
        {
            ExplorerDebugLog.Write("WinEvent hook thread did not stop within 2000ms");
            return;
        }

        _thread?.Join();
        _started.Dispose();
        _stopped.Dispose();
    }
}
