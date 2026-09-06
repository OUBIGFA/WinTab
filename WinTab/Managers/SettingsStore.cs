using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace WinTab.Managers;

internal sealed class SettingsStore : IDisposable
{
    private static readonly JsonSerializerOptions SerializerOptions = new() { WriteIndented = true };
    private readonly object _gate = new();
    private readonly string _path;
    private readonly Action<AppSettings> _write;
    private readonly Timer _timer;
    private AppSettings _settings;
    private Exception? _lastError;
    private TaskCompletionSource<bool>? _completion;
    private long _revision;
    private long _requestedRevision = -1;
    private long _writtenRevision = -1;
    private bool _saving;
    private bool _preserveBackup;
    private bool _disposed;

    public SettingsStore(string path, Action<AppSettings>? write = null)
    {
        _path = path;
        _write = write ?? WriteAtomically;
        _settings = Load();
        _timer = new Timer(state => FlushAsync(), null, Timeout.Infinite, Timeout.Infinite);
    }

    public AppSettings Snapshot => Volatile.Read(ref _settings);
    public Exception? LastError => Volatile.Read(ref _lastError);
    public event Action? ErrorChanged;

    public bool Update(Func<AppSettings, AppSettings> update, bool deferred = false)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            var next = update(_settings);
            if (next == _settings)
                return false;

            Volatile.Write(ref _settings, next);
            _revision++;
            _timer.Change(deferred ? 500 : 0, Timeout.Infinite);
            return true;
        }
    }

    public Task<bool> FlushAsync()
    {
        lock (_gate)
        {
            if (_disposed)
                return Task.FromResult(_writtenRevision == _revision && LastError == null);

            _timer.Change(Timeout.Infinite, Timeout.Infinite);
            _requestedRevision = _revision;
            if (_saving)
                return _completion!.Task;
            if (_writtenRevision == _revision && LastError == null)
                return Task.FromResult(true);

            _saving = true;
            _completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _ = Task.Run(SavePending);
            return _completion.Task;
        }
    }

    private void SavePending()
    {
        while (true)
        {
            AppSettings snapshot;
            long revision;
            lock (_gate)
            {
                snapshot = _settings;
                revision = _revision;
            }

            Exception? error = null;
            try
            {
                _write(snapshot);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
            {
                error = exception;
            }

            var previousError = Interlocked.Exchange(ref _lastError, error);
            if (error != null || previousError != null)
                ErrorChanged?.Invoke();

            lock (_gate)
            {
                if (error == null)
                    _writtenRevision = revision;
                if (_requestedRevision > revision)
                    continue;

                _saving = false;
                _completion!.TrySetResult(error == null);
                return;
            }
        }
    }

    private AppSettings Load()
    {
        try
        {
            return Read(_path);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
        {
            if (exception is not FileNotFoundException and not DirectoryNotFoundException)
                _lastError = exception;
        }

        try
        {
            var backup = Read(_path + ".bak");
            _preserveBackup = true;
            _lastError ??= new IOException("The settings file was missing; the last valid backup was restored.");
            return backup;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
        {
            if (exception is not FileNotFoundException and not DirectoryNotFoundException)
                _lastError ??= exception;
            return new AppSettings();
        }
    }

    private static AppSettings Read(string path) =>
        JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path, Encoding.UTF8))
        ?? throw new JsonException("The settings file does not contain an object.");

    private void WriteAtomically(AppSettings snapshot)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        var temporaryPath = _path + ".tmp";
        using (var stream = new FileStream(temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            JsonSerializer.Serialize(stream, snapshot, SerializerOptions);
            stream.Flush(flushToDisk: true);
        }

        if (File.Exists(_path))
            File.Replace(temporaryPath, _path, _preserveBackup ? null : _path + ".bak");
        else
            File.Move(temporaryPath, _path);
        _preserveBackup = false;
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
                return;
            _disposed = true;
            _timer.Dispose();
        }
    }
}
