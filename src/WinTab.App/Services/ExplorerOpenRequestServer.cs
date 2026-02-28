using System.IO;
using System.IO.Pipes;
using System.Text;
using WinTab.Diagnostics;

namespace WinTab.App.Services;

public sealed class ExplorerOpenRequestServer : IDisposable
{
    public const string PipeName = "WinTab_ExplorerOpenRequest";

    private readonly Logger _logger;
    private readonly CancellationTokenSource _cts = new();
    private Task? _loop;
    private bool _disposed;

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
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    // New protocol: OPEN_EX <foreground_hwnd_decimal> <path>
                    if (line.StartsWith("OPEN_EX ", StringComparison.Ordinal))
                    {
                        string remainder = line[8..].Trim();
                        int spaceIdx = remainder.IndexOf(' ');
                        if (spaceIdx > 0 &&
                            long.TryParse(remainder[..spaceIdx], out long hwndLong))
                        {
                            string path = remainder[(spaceIdx + 1)..].Trim();
                            _logger.Info($"Pipe: open-ex request (fg=0x{hwndLong:X}): {path}");
                            _ = SafeInvokeCallback(onOpenFolder, path, new IntPtr(hwndLong));
                            continue;
                        }
                    }

                    // Legacy protocol: OPEN <path>
                    if (line.StartsWith("OPEN ", StringComparison.Ordinal))
                    {
                        string path = line[5..].Trim();
                        _logger.Info($"Pipe: open request: {path}");
                        _ = SafeInvokeCallback(onOpenFolder, path, IntPtr.Zero);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException)
                {
                    _logger.Error("Pipe server error.", ex);
                    await Task.Delay(250);
                }
            }
        });
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
