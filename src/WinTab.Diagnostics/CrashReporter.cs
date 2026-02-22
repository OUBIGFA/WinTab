using System.Globalization;
using System.Text;

namespace WinTab.Diagnostics;

/// <summary>
/// Event arguments carrying details about an unhandled crash.
/// </summary>
public sealed class CrashEventArgs : EventArgs
{
    /// <summary>The exception that caused the crash, if available.</summary>
    public Exception? Exception { get; }

    /// <summary>Absolute path to the crash log file that was written.</summary>
    public string CrashLogPath { get; }

    /// <summary>UTC timestamp at which the crash was captured.</summary>
    public DateTimeOffset Timestamp { get; }

    public CrashEventArgs(Exception? exception, string crashLogPath, DateTimeOffset timestamp)
    {
        Exception = exception;
        CrashLogPath = crashLogPath;
        Timestamp = timestamp;
    }
}

/// <summary>
/// Registers global handlers for <see cref="AppDomain.UnhandledException"/> and
/// <see cref="TaskScheduler.UnobservedTaskException"/>, writes crash details to a
/// dedicated crash log file, and raises <see cref="CrashDetected"/> so that the
/// UI layer can present a notification.
/// </summary>
public static class CrashReporter
{
    private static readonly object Lock = new();
    private static string? _crashLogPath;
    private static bool _installed;

    private const string TimestampFormat = "yyyy-MM-dd HH:mm:ss.fff zzz";
    private const string Separator = "════════════════════════════════════════════════════════════════";

    /// <summary>
    /// Raised on the thread that caught the exception, immediately after the
    /// crash information has been persisted to disk. Subscribers should avoid
    /// long-running work because the process may be terminating.
    /// </summary>
    public static event EventHandler<CrashEventArgs>? CrashDetected;

    /// <summary>
    /// Registers global exception handlers. Safe to call multiple times;
    /// only the first invocation has any effect.
    /// </summary>
    /// <param name="crashLogPath">
    /// Absolute path to the file where crash reports are appended.
    /// The parent directory is created automatically if it does not exist.
    /// </param>
    /// <exception cref="ArgumentException">
    /// Thrown when <paramref name="crashLogPath"/> is null, empty, or whitespace.
    /// </exception>
    public static void Install(string crashLogPath)
    {
        if (string.IsNullOrWhiteSpace(crashLogPath))
            throw new ArgumentException("Crash log path must not be null or empty.", nameof(crashLogPath));

        lock (Lock)
        {
            if (_installed)
                return;

            _crashLogPath = Path.GetFullPath(crashLogPath);

            string? directory = Path.GetDirectoryName(_crashLogPath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

            _installed = true;
        }
    }

    // ── Event handlers ──────────────────────────────────────────────────

    private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        Exception? exception = e.ExceptionObject as Exception;
        RecordCrash(exception, isTerminating: e.IsTerminating);
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        // Flatten AggregateException to get the root causes.
        Exception exception = e.Exception.InnerExceptions.Count == 1
            ? e.Exception.InnerExceptions[0]
            : e.Exception;

        RecordCrash(exception, isTerminating: false);
    }

    // ── Core recording logic ────────────────────────────────────────────

    private static void RecordCrash(Exception? exception, bool isTerminating)
    {
        string path;
        lock (Lock)
        {
            if (_crashLogPath is null)
                return;
            path = _crashLogPath;
        }

        DateTimeOffset now = DateTimeOffset.Now;

        string report = BuildReport(exception, now, isTerminating);
        WriteToDisk(path, report);

        try
        {
            CrashDetected?.Invoke(null, new CrashEventArgs(exception, path, now));
        }
        catch
        {
            // Swallow subscriber exceptions so that crash reporting itself
            // never becomes a source of secondary failures.
        }
    }

    private static string BuildReport(Exception? exception, DateTimeOffset timestamp, bool isTerminating)
    {
        var sb = new StringBuilder();

        sb.AppendLine(Separator);
        sb.AppendLine($"  CRASH REPORT  —  {timestamp.ToString(TimestampFormat, CultureInfo.InvariantCulture)}");
        sb.AppendLine(Separator);
        sb.AppendLine();

        sb.AppendLine($"Terminating:    {isTerminating}");
        sb.AppendLine($"CLR Version:    {Environment.Version}");
        sb.AppendLine($"OS:             {Environment.OSVersion}");
        sb.AppendLine($"64-bit Process: {Environment.Is64BitProcess}");
        sb.AppendLine();

        if (exception is not null)
        {
            AppendException(sb, exception, depth: 0);
        }
        else
        {
            sb.AppendLine("Exception:      (not available — ExceptionObject was not a System.Exception)");
        }

        sb.AppendLine();
        return sb.ToString();
    }

    private static void AppendException(StringBuilder sb, Exception ex, int depth)
    {
        string indent = depth > 0 ? new string(' ', depth * 2) : string.Empty;
        string label = depth == 0 ? "Exception" : "Inner Exception";

        sb.AppendLine($"{indent}{label} Type:    {ex.GetType().FullName}");
        sb.AppendLine($"{indent}Message:         {ex.Message}");

        if (ex.Source is not null)
            sb.AppendLine($"{indent}Source:          {ex.Source}");

        if (ex.TargetSite is not null)
            sb.AppendLine($"{indent}Target Site:     {ex.TargetSite}");

        sb.AppendLine($"{indent}Stack Trace:");

        if (!string.IsNullOrWhiteSpace(ex.StackTrace))
        {
            foreach (string traceLine in ex.StackTrace.Split(Environment.NewLine))
                sb.AppendLine($"{indent}  {traceLine.TrimStart()}");
        }
        else
        {
            sb.AppendLine($"{indent}  (no stack trace available)");
        }

        sb.AppendLine();

        if (ex is AggregateException aggregate)
        {
            foreach (Exception inner in aggregate.InnerExceptions)
                AppendException(sb, inner, depth + 1);
        }
        else if (ex.InnerException is not null)
        {
            AppendException(sb, ex.InnerException, depth + 1);
        }
    }

    // ── Disk I/O ────────────────────────────────────────────────────────

    private static void WriteToDisk(string path, string content)
    {
        try
        {
            // Append so that multiple crashes within a session are captured
            // in a single file, making post-mortem analysis easier.
            File.AppendAllText(path, content, Encoding.UTF8);
        }
        catch
        {
            // Last-resort: if we cannot write to the intended path, try a
            // fallback next to the executable so we never silently lose data.
            try
            {
                string fallback = Path.Combine(
                    AppContext.BaseDirectory,
                    $"crash_{DateTimeOffset.UtcNow:yyyyMMdd_HHmmss}.log");

                File.WriteAllText(fallback, content, Encoding.UTF8);
            }
            catch
            {
                // Nothing more we can do — the process may be in a critically
                // broken state. Swallow to avoid masking the original crash.
            }
        }
    }
}
