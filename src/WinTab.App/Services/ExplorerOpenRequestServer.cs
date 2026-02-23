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

    public void Start(Func<string, Task> onOpenFolder)
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
                        maxNumberOfServerInstances: 1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await server.WaitForConnectionAsync(_cts.Token);

                    using var reader = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
                    string? line = await reader.ReadLineAsync(_cts.Token);
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    if (line.StartsWith("OPEN ", StringComparison.Ordinal))
                    {
                        string path = line[5..].Trim();
                        _logger.Info($"Pipe: open request: {path}");
                        await onOpenFolder(path);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.Error("Pipe server error.", ex);
                    await Task.Delay(250);
                }
            }
        });
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
