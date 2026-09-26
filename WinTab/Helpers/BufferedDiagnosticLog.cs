using System;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace WinTab.Helpers;

/// <summary>Where formatted diagnostic lines end up. Only the log's single writer task calls it.</summary>
internal interface IDiagnosticLogSink : IDisposable
{
    void Append(DateTime timestamp, byte[] line);
}

/// <summary>
/// Callers on input hooks and shell threads only queue a line; a single background task writes it, so
/// logging never makes them wait for the disk. A full queue drops lines and counts them instead. A line that
/// repeats the previous one is counted rather than written again, so a failure that recurs on a timer cannot
/// bury everything else in the file.
/// </summary>
internal sealed class BufferedDiagnosticLog
{
    internal const int RepeatSummaryLimit = 1_000;
    private readonly Channel<(DateTime Timestamp, string Message)> _messages;
    private readonly Task _writer;
    private readonly int _maximumMessageCharacters;
    private long _droppedMessages;
    private Exception? _lastError;

    /// <summary>Writes to one stream, which is emptied once it would outgrow <paramref name="maximumFileBytes"/>.</summary>
    public BufferedDiagnosticLog(Func<Stream> openStream, int capacity = 256, int maximumFileBytes = 1_048_576)
        : this(new StreamLogSink(openStream, maximumFileBytes), capacity, Math.Min(2_048, (maximumFileBytes - 64) / 3))
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(maximumFileBytes, 128);
    }

    /// <param name="sink">Used only by the writer task; it should not touch the disk before its first line.</param>
    public BufferedDiagnosticLog(IDiagnosticLogSink sink, int capacity = 1_024, int maximumMessageCharacters = 2_048)
    {
        ArgumentNullException.ThrowIfNull(sink);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(capacity);
        ArgumentOutOfRangeException.ThrowIfLessThan(maximumMessageCharacters, 2);
        _maximumMessageCharacters = maximumMessageCharacters;
        _messages = Channel.CreateBounded<(DateTime, string)>(new BoundedChannelOptions(capacity)
        {
            FullMode = BoundedChannelFullMode.Wait,
            SingleReader = true,
            AllowSynchronousContinuations = false
        });
        _writer = Task.Run(() => WritePendingAsync(sink));
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

        if (_messages.Writer.TryWrite((DateTime.Now, message)))
            return true;

        Interlocked.Increment(ref _droppedMessages);
        return false;
    }

    public Task CompleteAsync()
    {
        _messages.Writer.TryComplete();
        return _writer;
    }

    private async Task WritePendingAsync(IDiagnosticLogSink sink)
    {
        try
        {
            using (sink)
            {
                string? previous = null;
                var repeats = 0;
                var lastRepeat = default(DateTime);
                var reportedDrops = 0L;
                void WriteRepeats()
                {
                    if (repeats > 0)
                        sink.Append(lastRepeat, Format(lastRepeat, $"(previous line repeated {repeats} more time{(repeats == 1 ? "" : "s")})"));
                    repeats = 0;
                }
                void WriteDropped(DateTime timestamp)
                {
                    var dropped = DroppedMessages;
                    if (dropped == reportedDrops)
                        return;
                    WriteRepeats();
                    sink.Append(timestamp, Format(timestamp, $"(diagnostic queue dropped {dropped - reportedDrops} messages)"));
                    reportedDrops = dropped;
                    previous = null;
                }
                await foreach (var (timestamp, message) in _messages.Reader.ReadAllAsync().ConfigureAwait(false))
                {
                    WriteDropped(timestamp);
                    if (message == previous && timestamp.Date == lastRepeat.Date && repeats < RepeatSummaryLimit)
                    {
                        repeats++;
                        lastRepeat = timestamp;
                        continue;
                    }
                    WriteRepeats();
                    sink.Append(timestamp, Format(timestamp, message));
                    previous = message;
                    lastRepeat = timestamp;
                }
                WriteDropped(DateTime.Now);
                WriteRepeats();
            }
        }
        catch (Exception exception) when (IsStorageFailure(exception))
        {
            Volatile.Write(ref _lastError, exception);
            _messages.Writer.TryComplete();
            Debug.WriteLine($"Diagnostic logging stopped: {exception.Message}");
        }
    }

    private static byte[] Format(DateTime timestamp, string message) =>
        Encoding.UTF8.GetBytes($"{timestamp:HH:mm:ss.fff} {message}{Environment.NewLine}");

    internal static bool IsStorageFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or SecurityException;

    private sealed class StreamLogSink(Func<Stream> openStream, int maximumFileBytes) : IDiagnosticLogSink
    {
        private Stream? _stream;

        public void Append(DateTime timestamp, byte[] line)
        {
            var stream = _stream ??= openStream();
            if (stream.Length + line.Length > maximumFileBytes)
                stream.SetLength(0);
            stream.Position = stream.Length;
            stream.Write(line);
            stream.Flush();
        }

        public void Dispose() => _stream?.Dispose();
    }
}
