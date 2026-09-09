using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using SHDocVw;
using WinTab.Helpers;
using WinTab.Models;

namespace WinTab.Hooks;

public partial class ExplorerWatcher
{
    private const int MergeTimeoutMs = 5_000;
    private const int MergeCompletionTimeoutMs = 1_500;
    private readonly Timer _mergeSafetyTimer;
    private readonly Timer _selectionTimer;
    private CancellationTokenSource _shellLifetime = new();
    private CancellationTokenSource _hookLifetime = new();
    private Task _shellTransitionTask = Task.CompletedTask;
    private int _shellGeneration;
    private int _hookGeneration;
    private int _shellTransitionScheduled;
    private int _recoveringMergeSources;
    private volatile bool _disposed;

    private sealed record ConcealedWindow(WindowIdentity Identity, int Generation, long StartedAt)
    {
        public long ExpiresAt = StartedAt + MergeTimeoutMs;
        public int RestoreAttempts;
        public bool Recovering;
    }

    public event Action<string>? StatusChanged;
    private CancellationToken CurrentCancellation => _currentMerge.Value?.Token ?? _shellLifetime.Token;

    private void ReportStatus(string message)
    {
        ExplorerDebugLog.Write(message);
        StatusChanged?.Invoke(message);
    }

    private WindowInfo CreateWindowInfo(InternetExplorer window) => new()
    {
        Identity = WindowIdentity.Capture(new IntPtr(window.HWND)),
        Generation = _shellGeneration
    };

    private bool IsRegisteredWindow(InternetExplorer window, WindowInfo info)
    {
        lock (_windowEntryDictLock)
            return info.Generation == _shellGeneration &&
                _windowEntryDict.TryGetValue(window, out WindowInfo? current) && ReferenceEquals(current, info);
    }

    private bool IsCurrentWindow(InternetExplorer window, WindowInfo info) =>
        !_disposed && !info.Closed && info.Identity.IsCurrent && IsRegisteredWindow(window, info);

    private int RemainingMergeTime(nint handle) => _mergeSourceHWnds.TryGetValue(handle, out var concealed)
        ? (int)Math.Max(1, Interlocked.Read(ref concealed.ExpiresAt) - Environment.TickCount64)
        : MergeTimeoutMs;

    private void EnsureCurrentMerge()
    {
        if (_disposed)
            throw new OperationCanceledException("WinTab is stopping.");
        _shellLifetime.Token.ThrowIfCancellationRequested();
        _currentMerge.Value?.ThrowIfInvalid();
    }

    private Task RunShellWorkAsync(Func<Task> action)
    {
        var generation = _shellGeneration;
        return Task.Factory.StartNew(async () =>
        {
            if (_disposed || generation != _shellGeneration || _shellWindows == null || _shellLifetime.IsCancellationRequested)
                return;
            try
            {
                await action();
            }
            catch (OperationCanceledException) when (_disposed || generation != _shellGeneration || _shellLifetime.IsCancellationRequested)
            {
                ExplorerDebugLog.Write("Retired Explorer work cancelled");
            }
            catch (Exception exception) when (IsDisconnectedShell(exception))
            {
                if (!_disposed && generation == _shellGeneration)
                {
                    RetireShellConnection("shell-work-disconnected");
                    ReportStatus($"Explorer connection disconnected ({exception.HResult:X8}); reconnecting.");
                }
            }
            catch (Exception exception)
            {
                ReportStatus($"Explorer background work failed: {exception.GetType().Name}:{exception.Message}");
                throw;
            }
        }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, _staTaskScheduler).Unwrap();
    }

    private void EnsureWindowIdentity(WindowIdentity identity)
    {
        EnsureCurrentMerge();
        if (!identity.IsCurrent)
            throw new OperationCanceledException("The window no longer belongs to this operation.");
    }

    private Task CacheActiveSelectionAsync()
    {
        if (!_isForcingTabs || !ExplorerWindowDiscovery.IsFileExplorerForeground(out var handle))
            return Task.CompletedTask;
        var tab = GetActiveTabHandle(handle);
        var window = GetWindowByTabHandle(tab, handle);
        if (window == null)
            return Task.CompletedTask;
        WindowInfo? info;
        lock (_windowEntryDictLock)
            _windowEntryDict.TryGetValue(window, out info);
        if (info == null || !IsCurrentWindow(window, info))
            return Task.CompletedTask;

        info.RefreshSelection(() => TryGetSelectedItems(window),
            () => IsCurrentWindow(window, info) && GetActiveTabHandle(handle) == tab);
        return Task.CompletedTask;
    }

    private void RecoverExpiredMergeSources(object? state)
    {
        if (Interlocked.Exchange(ref _recoveringMergeSources, 1) != 0)
            return;
        try
        {
            foreach (var pair in _mergeSourceHWnds)
            {
                var concealed = pair.Value;
                string? status = null;
                lock (concealed)
                {
                    if (!concealed.Recovering && _isForcingTabs && concealed.Generation == _hookGeneration &&
                        Environment.TickCount64 < concealed.ExpiresAt)
                        continue;
                    if (concealed.RestoreAttempts >= 8)
                        continue;

                    concealed.RestoreAttempts++;
                    if (RestoreConcealedWindow(concealed))
                        status = "A merge exceeded its time limit; the source window was restored.";
                    else if (concealed.RestoreAttempts == 8)
                        status = "Explorer did not accept window recovery; stop WinTab and check the source window.";
                }
                if (status != null)
                    ReportStatus(status);
            }
        }
        finally
        {
            Volatile.Write(ref _recoveringMergeSources, 0);
        }
    }

    private bool RestoreConcealedWindow(ConcealedWindow concealed)
    {
        lock (concealed)
        {
            concealed.Recovering = true;
            if (concealed.Identity.IsCurrent && !ExplorerWindowVisibility.Restore(concealed.Identity))
                return false;
            ExplorerWindowVisibility.Forget(concealed.Identity);
            _mergeSourceHWnds.TryRemove(new KeyValuePair<nint, ConcealedWindow>(concealed.Identity.Handle, concealed));
            RemoveClosingMergeSource(concealed.Identity);
            if (concealed.Identity.IsCurrent)
                PreventWindowHiding(concealed.Identity.Handle);
            return true;
        }
    }

    private MergeOperation? CreateMergeCompletionOperation(InternetExplorer window, MergeOperation source)
    {
        if (!TryGetTrackedEntry(window, out var entry) || entry.Value.Identity != source.Identity ||
            !_mergeSourceHWnds.TryGetValue(source.Identity.Handle, out var concealed))
            return null;

        lock (concealed)
        {
            if (concealed.Recovering || concealed.Identity != source.Identity || concealed.Generation != source.Generation)
                return null;
            var completion = new MergeOperation(source.Identity, source.Generation, _shellLifetime.Token,
                () => _isForcingTabs && source.Generation == _hookGeneration && IsCurrentWindow(window, entry.Value),
                MergeCompletionTimeoutMs, _hookLifetime.Token);
            if (!completion.IsCurrent)
            {
                completion.Dispose();
                return null;
            }
            concealed.ExpiresAt = Environment.TickCount64 + MergeCompletionTimeoutMs;
            return completion;
        }
    }

    private void RemoveClosingMergeSource(WindowIdentity identity)
    {
        if (_closingMergeSourceHWnds.TryGetValue(identity.Handle, out var operation) && operation.Identity == identity)
            _closingMergeSourceHWnds.TryRemove(new KeyValuePair<nint, MergeOperation>(identity.Handle, operation));
    }

    private async Task InitializeShellAsync(int processId)
    {
        try
        {
            await Task.WhenAll(_registrationWork.WhenIdle, _selectionWork.WhenIdle).ConfigureAwait(false);
            await Task.Factory.StartNew(async () =>
            {
                if (_disposed)
                    return;
                DisposeShellObjects();
                _shellLifetime.Dispose();
                _shellLifetime = new CancellationTokenSource();
                Interlocked.Increment(ref _shellGeneration);
                _mainExplorerProcessId = processId;
                await InitializeShellObjectsAsync();
                if (!_disposed && !_shellLifetime.IsCancellationRequested)
                {
                    _selectionTimer.Change(500, 500);
                    OnShellInitialized?.Invoke();
                }
            }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, _staTaskScheduler).Unwrap().ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            _mainExplorerProcessId = 0;
            ReportStatus($"Explorer connection failed: {exception.GetType().Name}");
        }
        finally
        {
            Volatile.Write(ref _shellTransitionScheduled, 0);
        }
    }

    private static bool IsDisconnectedShell(Exception exception) =>
        exception is InvalidComObjectException || exception is COMException && exception.HResult is
            unchecked((int)0x80010108) or unchecked((int)0x800401FD) or
            unchecked((int)0x800706BA) or unchecked((int)0x800706BE) or
            unchecked((int)0x80010007) or unchecked((int)0x80010012);

    private void RetireShellConnection(string reason)
    {
        _preExistingExplorerWindowsProtected = false;
        Interlocked.Increment(ref _shellGeneration);
        _shellLifetime.Cancel();
        StopMergeSourceConcealPulse();
        RecoverHiddenExplorerWindows(reason);
        _mainExplorerProcessId = 0;
    }

    private async Task FinishDisposalAsync()
    {
        try
        {
            await Task.WhenAll(_registrationWork.StopAsync(), _selectionWork.StopAsync(), _shellTransitionTask).ConfigureAwait(false);
            await Task.Factory.StartNew(DisposeShellObjects, CancellationToken.None,
                TaskCreationOptions.DenyChildAttach, _staTaskScheduler).ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            Trace.TraceError($"Explorer cleanup failed: {exception.GetType().Name}");
        }
        finally
        {
            _staTaskScheduler.Dispose();
            _shellLifetime.Dispose();
            _hookLifetime.Dispose();
            _instanceRunning = false;
        }
    }
}
