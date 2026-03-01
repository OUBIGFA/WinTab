using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using WinTab.Diagnostics;

namespace WinTab.App.Services;

public sealed class ExplorerOpenRequestServer : IDisposable
{
    public const string PipeName = "WinTab_ExplorerOpenRequest";

    private const int MaxRequestLineLength = 8192;
    private const int MaxInvalidLogBurst = 10;
    private const int InvalidLogWindowMilliseconds = 5000;

    private readonly Logger _logger;
    private readonly CancellationTokenSource _cts = new();
    private readonly object _invalidLogThrottleLock = new();
    private Task? _loop;
    private bool _disposed;
    private int _invalidLogWindowStartTick;
    private int _invalidLogCountInWindow;

    public ExplorerOpenRequestServer(Logger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Starts the server. The callback receives (path, clickTimeForegroundHwnd).
    /// clickTimeForegroundHwnd is IntPtr.Zero if the request came from a legacy client.
    /// </summary>
    public void Start(Func<string, IntPtr, Task> onOpenFolder)
    {
        ArgumentNullException.ThrowIfNull(onOpenFolder);
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_loop is not null)
            return;

        _loop = Task.Run(async () =>
        {
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    await using var server = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.In,
                        maxNumberOfServerInstances: 4,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await server.WaitForConnectionAsync(_cts.Token);

                    using var reader = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
                    string? line = await reader.ReadLineAsync(_cts.Token);
                    if (line is null)
                        continue;

                    if (line.Length > MaxRequestLineLength)
                    {
                        LogInvalidRequestThrottled($"Pipe: request rejected, line too long ({line.Length} chars).");
                        continue;
                    }

                    if (TryParseOpenExRequest(line, out string? openExPath, out IntPtr openExForeground, out string? openExInvalidReason))
                    {
                        _logger.Info($"Pipe: open-ex request (fg=0x{openExForeground.ToInt64():X}): {openExPath}");
                        _ = SafeInvokeCallback(onOpenFolder, openExPath!, openExForeground);
                        continue;
                    }

                    if (openExInvalidReason is not null)
                    {
                        LogInvalidRequestThrottled($"Pipe: invalid OPEN_EX request ({openExInvalidReason}).");
                        continue;
                    }

                    if (TryParseOpenRequest(line, out string? openPath))
                    {
                        _logger.Info($"Pipe: open request: {openPath}");
                        _ = SafeInvokeCallback(onOpenFolder, openPath!, IntPtr.Zero);
                        continue;
                    }

                    LogInvalidRequestThrottled("Pipe: invalid request protocol.");
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
                {
                    _logger.Error("Pipe server error.", ex);
                    try
                    {
                        await Task.Delay(250, _cts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }
            }
        });
    }

    private static bool TryParseOpenRequest(string line, out string? path)
    {
        path = null;

        if (!line.StartsWith("OPEN ", StringComparison.Ordinal))
            return false;

        string candidatePath = line[5..].Trim();
        if (string.IsNullOrWhiteSpace(candidatePath))
            return false;

        if (candidatePath.Length > 2048)
            return false;

        path = candidatePath;
        return true;
    }

    private static bool TryParseOpenExRequest(string line, out string? path, out IntPtr foreground, out string? invalidReason)
    {
        path = null;
        foreground = IntPtr.Zero;
        invalidReason = null;

        if (!line.StartsWith("OPEN_EX ", StringComparison.Ordinal))
            return false;

        string remainder = line[8..].Trim();
        if (string.IsNullOrWhiteSpace(remainder))
        {
            invalidReason = "missing hwnd and path";
            return false;
        }

        int spaceIdx = remainder.IndexOf(' ');
        if (spaceIdx <= 0)
        {
            invalidReason = "missing path";
            return false;
        }

        string hwndText = remainder[..spaceIdx];
        if (!long.TryParse(hwndText, out long hwndLong))
        {
            invalidReason = "invalid hwnd";
            return false;
        }

        if (IntPtr.Size == 4 && (hwndLong < int.MinValue || hwndLong > int.MaxValue))
        {
            invalidReason = "hwnd out of range";
            return false;
        }

        string candidatePath = remainder[(spaceIdx + 1)..].Trim();
        if (string.IsNullOrWhiteSpace(candidatePath))
        {
            invalidReason = "empty path";
            return false;
        }

        if (candidatePath.Length > 2048)
        {
            invalidReason = "path too long";
            return false;
        }

        foreground = new IntPtr(hwndLong);
        path = candidatePath;
        return true;
    }

    private void LogInvalidRequestThrottled(string message)
    {
        lock (_invalidLogThrottleLock)
        {
            int now = Environment.TickCount;
            if (_invalidLogWindowStartTick == 0 || now - _invalidLogWindowStartTick >= InvalidLogWindowMilliseconds)
            {
                _invalidLogWindowStartTick = now;
                _invalidLogCountInWindow = 0;
            }

            if (_invalidLogCountInWindow < MaxInvalidLogBurst)
            {
                _invalidLogCountInWindow++;
                _logger.Warn(message);
                return;
            }

            if (_invalidLogCountInWindow == MaxInvalidLogBurst)
            {
                _invalidLogCountInWindow++;
                _logger.Warn("Pipe: invalid request logging is temporarily throttled.");
            }
        }
    }

    private async Task SafeInvokeCallback(Func<string, IntPtr, Task> callback, string path, IntPtr foreground)
    {
        try
        {
            await callback(path, foreground);
        }
        catch (Exception ex)
        {
            _logger.Error($"Pipe callback error for '{path}'.", ex);
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        _cts.Cancel();
        try { _loop?.Wait(1000); } catch { /* ignore */ }
        _cts.Dispose();
    }
}
