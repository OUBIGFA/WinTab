using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using WinTab.Hooks;
using WinTab.UI.Localization;

namespace WinTab.UI.Desktop;

/// <summary>
/// Starts the settings renderer only on demand. Anonymous stdio pipes belong to this child alone:
/// there is no listening port, shared secret, file polling, or second owner of Explorer hooks.
/// </summary>
internal sealed class DesktopUiProcess(Dispatcher dispatcher, Func<DesktopCommand, Task<object>> execute,
    Action<uint> processChanged) : IDisposable
{
    private Process? _process;
    private Channel<string>? _outbox;
    private bool _disposed;

    public void Show()
    {
        if (_disposed) return;
        if (_process is { HasExited: false })
        {
            Send(new { @event = "show" });
            return;
        }

        var executable = Path.Combine(AppContext.BaseDirectory, "WinTab.UI.exe");
        if (!File.Exists(executable))
        {
            MessageBox.Show(UiStrings.DesktopUiMissing, "WinTab", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8, WorkingDirectory = AppContext.BaseDirectory
        };
        start.ArgumentList.Add("--bridge");
        try
        {
            var process = Process.Start(start) ?? throw new IOException("UI process did not start");
            var outbox = Channel.CreateBounded<string>(new BoundedChannelOptions(128)
            {
                SingleReader = true, FullMode = BoundedChannelFullMode.Wait
            });
            _process = process;
            _outbox = outbox;
            processChanged((uint)process.Id);
            // A launch from the tray or the single-instance signal grants the child foreground rights.
            AllowSetForegroundWindow((uint)process.Id);
            _ = RunAsync(process, outbox);
        }
        catch (Exception error)
        {
            ExplorerDebugLog.Write($"Could not start desktop UI: {error}");
            MessageBox.Show(UiStrings.DesktopUiFailed, "WinTab", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void Publish(object state) => Send(new { @event = "state", data = state });

    private void Send(object message)
    {
        if (_disposed || _outbox is not { } queue) return;
        // Never let a stalled renderer accumulate an unbounded queue inside the resident engine.
        if (!queue.Writer.TryWrite(JsonSerializer.Serialize(message, DesktopCommand.Json)))
        {
            ExplorerDebugLog.Write("Desktop UI stopped consuming its pipe; closing only the renderer");
            if (_process is { HasExited: false } process) process.Kill(entireProcessTree: true);
        }
    }

    private async Task RunAsync(Process process, Channel<string> outbox)
    {
        var writer = WriteAsync(process, outbox);
        var diagnostics = ReadDiagnosticsAsync(process);
        var pending = new List<Task>();
        try
        {
            while (await ReadMessageAsync(process.StandardOutput).ConfigureAwait(false) is { } line)
            {
                pending.RemoveAll(task => task.IsCompleted);
                // Normal settings finish immediately; a damaged child must not retain unbounded
                // commands while a recovery or update is awaiting external work.
                if (pending.Count >= 128) throw new IOException("Too many pending UI requests");
                long id = 0;
                Task<object> operation;
                try
                {
                    using var document = JsonDocument.Parse(line);
                    if (document.RootElement.TryGetProperty("id", out var value)) value.TryGetInt64(out id);
                    var command = DesktopCommand.Parse(line);
                    // Await dispatch, not command completion: mutations still begin in wire order
                    // on the one UI thread, while long operations reply independently by request id.
                    operation = await dispatcher.InvokeAsync(() => execute(command)).Task.ConfigureAwait(false);
                }
                catch (Exception error)
                {
                    operation = Task.FromException<object>(error);
                }
                pending.Add(ReplyAsync(id, operation, outbox.Writer));
            }
            await process.WaitForExitAsync().ConfigureAwait(false);
            if (process.ExitCode != 0 && !_disposed)
            {
                ExplorerDebugLog.Write($"Desktop UI exited with code {process.ExitCode}");
                await dispatcher.InvokeAsync(() => MessageBox.Show(UiStrings.DesktopUiFailed, "WinTab",
                    MessageBoxButton.OK, MessageBoxImage.Error)).Task.ConfigureAwait(false);
            }
        }
        catch (Exception error) when (error is IOException or ObjectDisposedException or InvalidOperationException)
        {
            if (!_disposed) ExplorerDebugLog.Write($"Desktop UI pipe ended: {error}");
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        }
        finally
        {
            outbox.Writer.TryComplete();
            await writer.ConfigureAwait(false);
            await diagnostics.ConfigureAwait(false);
            try
            {
                // Await the actual Task explicitly; ConfigureAwait.Fody cannot safely rewrite
                // a direct DispatcherOperation await, and shutdown must still release the handle.
                await dispatcher.InvokeAsync(() =>
                {
                    if (!ReferenceEquals(_process, process)) return;
                    _process = null;
                    _outbox = null;
                    processChanged(0);
                }).Task.ConfigureAwait(false);
            }
            finally { process.Dispose(); }
        }
    }

    private static async Task ReplyAsync(long id, Task<object> operation, ChannelWriter<string> replies)
    {
        string response;
        try
        {
            var result = await operation.ConfigureAwait(false);
            response = JsonSerializer.Serialize(new { id, result }, DesktopCommand.Json);
        }
        catch (Exception error)
        {
            // Keep observing the operation even after its renderer exits; the core owns its lifetime.
            ExplorerDebugLog.Write($"Desktop UI request rejected: {error}");
            response = JsonSerializer.Serialize(new { id, error = UiStrings.DesktopCommandFailed }, DesktopCommand.Json);
        }
        try { await replies.WriteAsync(response).ConfigureAwait(false); }
        catch (ChannelClosedException) { /* A disconnected renderer no longer accepts its pending reply. */ }
    }

    /// <summary>Bound a line before allocating it, even if a damaged child never emits a newline.</summary>
    internal static async Task<string?> ReadMessageAsync(TextReader reader)
    {
        var buffer = new char[1];
        var line = new StringBuilder();
        // ReadLineAsync has no size limit; a small char buffer also handles partial UTF-8 reads safely.
        while (true)
        {
            var count = await reader.ReadAsync(buffer, 0, 1).ConfigureAwait(false);
            if (count == 0) return line.Length == 0 ? null : throw new IOException("Incomplete UI message");
            if (buffer[0] == '\n') return line.ToString().TrimEnd('\r');
            if (line.Length >= DesktopCommand.MaximumMessageLength) throw new IOException("UI message too large");
            line.Append(buffer[0]);
        }
    }

    private static async Task WriteAsync(Process process, Channel<string> outbox)
    {
        try
        {
            await foreach (var line in outbox.Reader.ReadAllAsync().ConfigureAwait(false))
            {
                await process.StandardInput.WriteLineAsync(line).ConfigureAwait(false);
                await process.StandardInput.FlushAsync().ConfigureAwait(false);
            }
        }
        catch (Exception error) when (error is IOException or ObjectDisposedException or InvalidOperationException)
        {
            ExplorerDebugLog.Write($"Desktop UI input closed: {error.Message}");
        }
    }

    private static async Task ReadDiagnosticsAsync(Process process)
    {
        try
        {
            while (await process.StandardError.ReadLineAsync().ConfigureAwait(false) is { } line)
                ExplorerDebugLog.Write("Desktop UI: " + line);
        }
        catch (Exception error) when (error is IOException or ObjectDisposedException or InvalidOperationException)
        {
            ExplorerDebugLog.Write($"Desktop UI diagnostics closed: {error.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        processChanged(0);
        _outbox?.Writer.TryComplete();
        if (_process is not { } process) return;
        // EOF makes MyGo quit even after a parent crash. Normal exit gives it time to release WebView2.
        try
        {
            process.StandardInput.Close();
            if (!process.WaitForExit(1200)) process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException) { /* The captured child has already exited. */ }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool AllowSetForegroundWindow(uint processId);
}
