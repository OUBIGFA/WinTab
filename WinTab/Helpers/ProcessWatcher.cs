using System;
using System.Threading;
using System.Diagnostics;
using System.Collections.Concurrent;

namespace WinTab.Helpers;

public class ProcessEventArgs(int processId) : EventArgs
{
    public int ProcessId { get; } = processId;
}

/// <summary>
/// Raises <see cref="ProcessTerminated"/> for every process with the given name in the current
/// session. Explorer does not always deliver Exited for windows it restarts, so a periodic scan
/// backs up the event subscription.
/// </summary>
public sealed class ProcessWatcher : IDisposable
{
    private readonly string _processName;
    private readonly int _currentSessionId;
    private readonly ConcurrentDictionary<int, Process> _trackedProcesses = new();
    private readonly Timer _scanTimer;
    private readonly object _scanLock = new();
    private readonly SynchronizationContext? _syncContext;
    private bool _disposed;

    public event EventHandler<ProcessEventArgs>? ProcessTerminated;

    public ProcessWatcher(string processName, int scanIntervalMs = 5000)
    {
        _processName = processName ?? throw new ArgumentNullException(nameof(processName));
        _currentSessionId = Process.GetCurrentProcess().SessionId;
        _syncContext = SynchronizationContext.Current;

        var scanInterval = TimeSpan.FromMilliseconds(Math.Max(100, scanIntervalMs));
        _scanTimer = new Timer(ScanForProcesses, null, TimeSpan.Zero, scanInterval);
    }

    private void ScanForProcesses(object? state)
    {
        if (_disposed) return;
        if (!Monitor.TryEnter(_scanLock)) return;

        try
        {
            foreach (var (pid, process) in _trackedProcesses)
            {
                bool hasExited;
                try
                {
                    hasExited = process.HasExited;
                }
                catch (Exception)
                {
                    hasExited = true;
                }

                if (hasExited && _trackedProcesses.TryRemove(pid, out var removed))
                    NotifyTerminated(pid, removed);
            }

            foreach (var process in Process.GetProcessesByName(_processName))
            {
                try
                {
                    var processId = process.Id;
                    if (_trackedProcesses.ContainsKey(processId) || process.SessionId != _currentSessionId)
                    {
                        SafeDispose(process);
                        continue;
                    }

                    process.Exited += OnProcessExited;
                    process.EnableRaisingEvents = true;

                    if (!_trackedProcesses.TryAdd(processId, process))
                    {
                        process.Exited -= OnProcessExited;
                        SafeDispose(process);
                    }
                }
                catch (Exception)
                {
                    SafeDispose(process);
                }
            }
        }
        finally
        {
            Monitor.Exit(_scanLock);
        }
    }

    private void OnProcessExited(object? sender, EventArgs e)
    {
        if (_disposed || sender is not Process process) return;

        try
        {
            var processId = process.Id;
            if (_trackedProcesses.TryRemove(processId, out var tracked))
                NotifyTerminated(processId, tracked);
        }
        catch (Exception)
        {
            // Process object might be in an inconsistent state after exit
        }
    }

    private void NotifyTerminated(int processId, Process process)
    {
        try
        {
            process.Exited -= OnProcessExited;

            var handler = ProcessTerminated;
            if (handler == null) return;

            var args = new ProcessEventArgs(processId);
            if (_syncContext != null)
                _syncContext.Post(_ => handler(this, args), null);
            else
                handler(this, args);
        }
        catch (Exception)
        {
            // Ignore cleanup errors
        }
        finally
        {
            SafeDispose(process);
        }
    }

    private static void SafeDispose(Process process)
    {
        try
        {
            process.Dispose();
        }
        catch (Exception)
        {
            // Ignore disposal errors
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            _scanTimer.Change(Timeout.Infinite, Timeout.Infinite);
            _scanTimer.Dispose();
        }
        catch
        {
            // ignored
        }

        foreach (var process in _trackedProcesses.Values)
        {
            try
            {
                process.Exited -= OnProcessExited;
            }
            catch
            {
                // ignored
            }

            SafeDispose(process);
        }

        _trackedProcesses.Clear();
    }
}
