using System;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace WinTab.Helpers;

internal sealed class BufferedDiagnosticLog
{
    private readonly Channel<string> _messages;
    private readonly Task _writer;
    private readonly int _maximumMessageCharacters;
    private long _droppedMessages;
    private Exception? _lastError;

    public BufferedDiagnosticLog(Func<Stream> openStream, int capacity = 256, int maximumFileBytes = 1_048_576)
    {
        ArgumentNullException.ThrowIfNull(openStream);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(capacity);
        ArgumentOutOfRangeException.ThrowIfLessThan(maximumFileBytes, 128);
        _maximumMessageCharacters = Math.Min(2_048, (maximumFileBytes - 64) / 3);
        _messages = Channel.CreateBounded<string>(new BoundedChannelOptions(capacity)
        {
            FullMode = BoundedChannelFullMode.Wait,
            SingleReader = true,
            AllowSynchronousContinuations = false
        });
        _writer = Task.Run(() => WritePendingAsync(openStream, maximumFileBytes));
    }

    public long DroppedMessages => Interlocked.Read(ref _droppedMessages);
    public Exception? LastError => Volatile.Read(ref _lastError);

    public bool TryWrite(string message)
    {
        if (message.Length > _maximumMessageCharacters)
        {
            var length = _maximumMessageCharacters;
            if (char.IsHighSurrogate(message[length - 1]))
                length--;
            message = message[..length] + "…";
        }

        if (_messages.Writer.TryWrite($"{DateTime.Now:HH:mm:ss.fff} {message}{Environment.NewLine}"))
            return true;

        Interlocked.Increment(ref _droppedMessages);
        return false;
    }

    public Task CompleteAsync()
    {
        _messages.Writer.TryComplete();
        return _writer;
    }

    private async Task WritePendingAsync(Func<Stream> openStream, int maximumFileBytes)
    {
        try
        {
            using var stream = openStream();
            await foreach (var message in _messages.Reader.ReadAllAsync().ConfigureAwait(false))
            {
                var bytes = Encoding.UTF8.GetBytes(message);
                if (stream.Length + bytes.Length > maximumFileBytes)
                    stream.SetLength(0);
                stream.Position = stream.Length;
                stream.Write(bytes);
                stream.Flush();
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or SecurityException)
        {
            Volatile.Write(ref _lastError, exception);
            _messages.Writer.TryComplete();
            Debug.WriteLine($"Diagnostic logging stopped: {exception.Message}");
        }
    }
}
