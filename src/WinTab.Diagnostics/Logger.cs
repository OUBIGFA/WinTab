using System.Globalization;
using System.Text;

namespace WinTab.Diagnostics;

/// <summary>
/// Severity levels for log messages, ordered from most verbose to most severe.
/// </summary>
public enum LogLevel
{
    Debug = 0,
    Info = 1,
    Warn = 2,
    Error = 3,
}

/// <summary>
/// Thread-safe file logger with automatic log rotation.
/// Rotates when a file exceeds <see cref="MaxFileSizeBytes"/> and retains
/// up to <see cref="MaxRetainedFiles"/> historical files.
/// </summary>
public sealed class Logger : IDisposable
{
    /// <summary>Maximum size of a single log file before rotation occurs.</summary>
    public const long MaxFileSizeBytes = 5 * 1024 * 1024; // 5 MB

    /// <summary>Maximum number of rotated (historical) log files to keep.</summary>
    public const int MaxRetainedFiles = 5;

    private const string TimestampFormat = "yyyy-MM-dd HH:mm:ss.fff zzz";

    private readonly string _logFilePath;
    private readonly object _lock = new();
    private StreamWriter? _writer;
    private bool _disposed;

    /// <summary>
    /// Messages with a level below <see cref="MinLevel"/> are silently discarded.
    /// This property is safe to change at any time; reads and writes of enum-sized
    /// values are atomic on all .NET runtimes.
    /// </summary>
    public LogLevel MinLevel { get; set; } = LogLevel.Debug;

    /// <summary>
    /// Creates a new <see cref="Logger"/> that writes to <paramref name="logFilePath"/>.
    /// The parent directory is created automatically if it does not exist.
    /// </summary>
    /// <param name="logFilePath">Absolute path to the primary log file.</param>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="logFilePath"/> is null, empty, or whitespace.
    /// </exception>
    public Logger(string logFilePath)
    {
        if (string.IsNullOrWhiteSpace(logFilePath))
            throw new ArgumentException("Log file path must not be null or empty.", nameof(logFilePath));

        _logFilePath = Path.GetFullPath(logFilePath);

        string? directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        OpenWriter(append: true);
    }

    // ── Public convenience methods ──────────────────────────────────────

    public void Debug(string message) => Write(LogLevel.Debug, message);
    public void Info(string message) => Write(LogLevel.Info, message);
    public void Warn(string message) => Write(LogLevel.Warn, message);
    public void Error(string message) => Write(LogLevel.Error, message);

    public void Error(string message, Exception exception)
    {
        Write(LogLevel.Error, $"{message}{Environment.NewLine}{exception}");
    }

    // ── Core write path ─────────────────────────────────────────────────

    /// <summary>
    /// Writes a single log entry if <paramref name="level"/> meets the
    /// <see cref="MinLevel"/> threshold. The call is thread-safe.
    /// </summary>
    public void Write(LogLevel level, string message)
    {
        if (level < MinLevel)
            return;

        string timestamp = DateTimeOffset.Now.ToString(TimestampFormat, CultureInfo.InvariantCulture);
        string levelTag = level switch
        {
            LogLevel.Debug => "DEBUG",
            LogLevel.Info  => "INFO",
            LogLevel.Warn  => "WARN",
            LogLevel.Error => "ERROR",
            _              => level.ToString().ToUpperInvariant(),
        };

        string line = $"{timestamp} [{levelTag}] {message}";

        lock (_lock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            EnsureWriter();
            _writer!.WriteLine(line);
            _writer.Flush();

            RotateIfNeeded();
        }
    }

    // ── Rotation logic ──────────────────────────────────────────────────

    private void RotateIfNeeded()
    {
        // Check current file size via the underlying stream.
        if (_writer?.BaseStream is not { Length: >= MaxFileSizeBytes })
            return;

        CloseWriter();

        string directory = Path.GetDirectoryName(_logFilePath)!;
        string baseName = Path.GetFileNameWithoutExtension(_logFilePath);
        string extension = Path.GetExtension(_logFilePath);

        // Delete the oldest file if it would exceed our retention count.
        string oldest = Path.Combine(directory, $"{baseName}.{MaxRetainedFiles}{extension}");
        if (File.Exists(oldest))
            File.Delete(oldest);

        // Shift existing rotated files: .4 -> .5, .3 -> .4, ... , .1 -> .2
        for (int i = MaxRetainedFiles - 1; i >= 1; i--)
        {
            string src = Path.Combine(directory, $"{baseName}.{i}{extension}");
            string dst = Path.Combine(directory, $"{baseName}.{i + 1}{extension}");
            if (File.Exists(src))
                File.Move(src, dst, overwrite: true);
        }

        // Rotate the current file to .1
        string rotatedPath = Path.Combine(directory, $"{baseName}.1{extension}");
        if (File.Exists(_logFilePath))
            File.Move(_logFilePath, rotatedPath, overwrite: true);

        OpenWriter(append: false);
    }

    // ── Stream management ───────────────────────────────────────────────

    private void OpenWriter(bool append)
    {
        var stream = new FileStream(
            _logFilePath,
            append ? FileMode.Append : FileMode.Create,
            FileAccess.Write,
            FileShare.Read);

        _writer = new StreamWriter(stream, Encoding.UTF8)
        {
            AutoFlush = false,
        };
    }

    private void EnsureWriter()
    {
        if (_writer is null)
            OpenWriter(append: true);
    }

    private void CloseWriter()
    {
        _writer?.Dispose();
        _writer = null;
    }

    // ── IDisposable ─────────────────────────────────────────────────────

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed)
                return;

            _disposed = true;
            CloseWriter();
        }
    }
}
