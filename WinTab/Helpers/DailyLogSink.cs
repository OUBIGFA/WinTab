using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace WinTab.Helpers;

/// <summary>
/// One log file per day (<c>WinTab-yyyyMMdd.log</c>) in a folder of its own. Only today's and yesterday's
/// files are kept: older ones are deleted when the log opens and whenever the date changes while WinTab runs.
/// A day that outgrows its size limit moves its file aside once (<c>.1.log</c>) and continues in a new one,
/// bounding a day to twice the limit. A failed write is retried after half a minute; the next successful write
/// reports the dropped lines and failure so that a logging gap is not mistaken for a missing input event.
/// </summary>
internal sealed class DailyLogSink : IDiagnosticLogSink
{
    internal const string FilePrefix = "WinTab-";
    internal const string FileExtension = ".log";
    internal const int RetainedDays = 2;
    private const string DateFormat = "yyyyMMdd";
    private const int RetryAfterFailureMs = 30_000;
    private readonly Func<long> _getTickCount;
    private readonly string _folder;
    private readonly long _maximumFileBytes;
    private FileStream? _stream;
    private DateTime _day;
    private long _retryAt;
    private long _reportedDrops;
    private DateTime _firstDropAt;

    public DailyLogSink(string folder, long maximumFileBytes = 8 * 1024 * 1024, Func<long>? getTickCount = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(folder);
        ArgumentOutOfRangeException.ThrowIfLessThan(maximumFileBytes, 1_024);
        _getTickCount = getTickCount ?? (() => Environment.TickCount64);
        _folder = Path.GetFullPath(folder);
        _maximumFileBytes = maximumFileBytes;
    }

    public static string FileName(DateTime day) => FilePrefix + day.ToString(DateFormat, CultureInfo.InvariantCulture) + FileExtension;

    public long DroppedLines { get; private set; }
    public Exception? LastError { get; private set; }

    public void Append(DateTime timestamp, byte[] line)
    {
        if (_retryAt != 0 && _getTickCount() < _retryAt)
        {
            DropLine(timestamp);
            return;
        }
        try
        {
            if (_stream == null || timestamp.Date != _day)
                OpenDay(timestamp.Date);
            if (DroppedLines > _reportedDrops)
            {
                WriteLine(Encoding.UTF8.GetBytes($"{timestamp:HH:mm:ss.fff} Diagnostic logging resumed; dropped {DroppedLines - _reportedDrops} lines " +
                    $"since {_firstDropAt:yyyy-MM-dd HH:mm:ss.fff}: {LastError?.GetType().Name}: {LastError?.Message}{Environment.NewLine}"));
                _reportedDrops = DroppedLines;
            }
            WriteLine(line);
            _retryAt = 0;
        }
        catch (Exception exception) when (BufferedDiagnosticLog.IsStorageFailure(exception))
        {
            _stream?.Dispose();
            _stream = null;
            LastError = exception;
            DropLine(timestamp);
            _retryAt = _getTickCount() + RetryAfterFailureMs;
        }
    }

    private void DropLine(DateTime timestamp)
    {
        if (DroppedLines == _reportedDrops)
            _firstDropAt = timestamp;
        DroppedLines++;
    }

    private void WriteLine(byte[] line)
    {
        if (_stream!.Length > 0 && _stream.Length + line.Length > _maximumFileBytes)
            StartOverflowFile();
        _stream!.Write(line);
        _stream.Flush();
    }

    private void OpenDay(DateTime day)
    {
        _stream?.Dispose();
        _stream = null;
        Directory.CreateDirectory(_folder);
        DeleteExpiredFiles(_folder, day);
        _day = day;
        _stream = Open(Path.Combine(_folder, FileName(day)));
    }

    private void StartOverflowFile()
    {
        var path = Path.Combine(_folder, FileName(_day));
        _stream!.Dispose();
        _stream = null;
        File.Move(path, Path.ChangeExtension(path, ".1" + FileExtension), overwrite: true);
        _stream = Open(path);
    }

    private static FileStream Open(string path) =>
        new(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);

    /// <summary>
    /// Deletes this log's files older than the retained days. Only names this log writes are considered, and a
    /// file another program holds open is left for the next pass instead of stopping the log.
    /// </summary>
    internal static void DeleteExpiredFiles(string folder, DateTime today)
    {
        var oldestKept = today.Date.AddDays(1 - RetainedDays);
        foreach (var path in Directory.EnumerateFiles(folder, FilePrefix + "*" + FileExtension))
        {
            if (TryReadDay(Path.GetFileName(path), out var day) && day < oldestKept)
            {
                try { File.Delete(path); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            }
        }
    }

    internal static bool TryReadDay(string fileName, out DateTime day)
    {
        day = default;
        if (!fileName.StartsWith(FilePrefix, StringComparison.OrdinalIgnoreCase) ||
            !fileName.EndsWith(FileExtension, StringComparison.OrdinalIgnoreCase))
            return false;
        var stem = fileName[FilePrefix.Length..^FileExtension.Length];
        if (stem.EndsWith(".1", StringComparison.Ordinal))
            stem = stem[..^2];
        return stem.Length == DateFormat.Length &&
            DateTime.TryParseExact(stem, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out day);
    }

    public void Dispose() => _stream?.Dispose();
}
