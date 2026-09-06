using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;

namespace WinTab.Helpers;

internal sealed class BackgroundRefreshCache<TKey, TValue> : IDisposable where TKey : notnull where TValue : class
{
    private sealed class Entry
    {
        public TValue? Value;
        public int Version;
        public int Scheduled;
    }

    private readonly ConcurrentDictionary<TKey, Entry> _entries = new();
    private readonly BlockingCollection<(TKey Key, Entry Entry)> _queue = new(64);
    private readonly object _queueGate = new();
    private readonly Func<TKey, TValue?> _compute;
    private bool _disposed;

    public BackgroundRefreshCache(Func<TKey, TValue?> compute)
    {
        _compute = compute;
        var thread = new Thread(Run) { IsBackground = true, Name = "WinTab tab bounds" };
        thread.SetApartmentState(ApartmentState.MTA);
        thread.Start();
    }

    public bool TryGet(TKey key, out TValue? value)
    {
        value = _entries.TryGetValue(key, out var entry) ? Volatile.Read(ref entry.Value) : null;
        return value != null;
    }

    public void Request(TKey key)
    {
        if (!_entries.ContainsKey(key) && _entries.Count >= 256)
            return;
        Schedule(key, _entries.GetOrAdd(key, static key => new Entry()));
    }

    public void Invalidate(TKey key)
    {
        if (_entries.TryGetValue(key, out var entry))
        {
            lock (entry)
            {
                entry.Version++;
                Volatile.Write(ref entry.Value, null);
            }
            Schedule(key, entry);
        }
        else
        {
            Request(key);
        }
    }

    private void Schedule(TKey key, Entry entry)
    {
        lock (_queueGate)
        {
            if (_disposed || Interlocked.CompareExchange(ref entry.Scheduled, 1, 0) != 0)
                return;
            if (!_queue.TryAdd((key, entry)))
                Volatile.Write(ref entry.Scheduled, 0);
        }
    }

    private bool IsCurrent(TKey key, Entry entry) =>
        _entries.TryGetValue(key, out var current) && ReferenceEquals(entry, current);

    private void Run()
    {
        try
        {
            foreach (var (key, entry) in _queue.GetConsumingEnumerable())
            {
                var version = Volatile.Read(ref entry.Version);
                try
                {
                    if (!IsCurrent(key, entry))
                        continue;
                    var value = _compute(key);
                    lock (entry)
                    {
                        if (version == entry.Version && IsCurrent(key, entry))
                            Volatile.Write(ref entry.Value, value);
                    }
                }
                catch (Exception exception)
                {
                    Trace.TraceError($"Tab bounds refresh failed: {exception.GetType().Name}");
                }
                finally
                {
                    Volatile.Write(ref entry.Scheduled, 0);
                    if (version != Volatile.Read(ref entry.Version) && IsCurrent(key, entry))
                        Schedule(key, entry);
                }
            }
        }
        finally
        {
            _queue.Dispose();
        }
    }

    public void Forget(TKey key) => _entries.TryRemove(key, out _);
    public void Clear() => _entries.Clear();

    public void Dispose()
    {
        lock (_queueGate)
        {
            if (_disposed)
                return;
            _disposed = true;
            _entries.Clear();
            _queue.CompleteAdding();
        }
    }
}
