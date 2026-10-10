using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Windows.Threading;
using WinTab.Helpers;
using WinTab.UI.Desktop;

internal static class DesktopUiProcessTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("desktop pipe clean disconnect releases its process without a woven await failure", CleanDisconnect);
        yield return ("desktop pipe handles settings while a restore request remains pending", () => LongOperation("restore"));
        yield return ("desktop pipe handles settings while an update request remains pending", () => LongOperation("update"));
    }

    // An isolated console peer exercises real anonymous pipes; it never creates an application window,
    // installs hooks, changes user preferences, or sends native input.
    internal static async Task<int> RunPeerAsync(string resultPath, int replyCount)
    {
        Console.Write(Environment.GetEnvironmentVariable("WINTAB_PIPE_TEST_REQUESTS"));
        Console.Out.Flush();
        for (var i = 0; i < replyCount; i++)
        {
            var line = await Console.In.ReadLineAsync();
            if (line is null) return 3;
            await File.AppendAllTextAsync(resultPath, line + "\n", new UTF8Encoding(false));
        }
        return 0;
    }

    private static async Task CleanDisconnect()
    {
        using var pipe = new PipeFixture(Request(1, "state"), 1, command => Task.FromResult<object>(new { revision = command.Id }));
        await pipe.Completion.WaitAsync(TimeSpan.FromSeconds(5));
        var reply = pipe.Replies()[0];
        Check.Equal(1L, reply.GetProperty("id").GetInt64(), "A clean exit must deliver the final response before releasing the child");
        Check.That(reply.TryGetProperty("result", out _), "The completed command must retain its result");
        Check.That(pipe.ProcessReleased, "The renderer must no longer be registered after disconnect");
        try { _ = pipe.Child.Handle; }
        catch (InvalidOperationException) { return; }
        throw new InvalidOperationException("The renderer's Process handle was not disposed after disconnect");
    }

    private static async Task LongOperation(string method)
    {
        var release = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var settingHandled = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var order = new List<string>();
        Task<object> Execute(DesktopCommand command)
        {
            order.Add(command.Method);
            if (command.Method == method) { started.TrySetResult(); return release.Task; }
            settingHandled.TrySetResult();
            return Task.FromResult<object>(new { applied = command.Key });
        }
        using var pipe = new PipeFixture(Request(1, method, method == "restore" ? new { group = true } : null) +
            Request(2, "set", new { key = "theme", value = "Dark" }), 2, Execute);
        try
        {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            try { await settingHandled.Task.WaitAsync(TimeSpan.FromSeconds(1)); }
            catch (TimeoutException) { throw new InvalidOperationException("The setting request was blocked by the pending " + method); }
            await pipe.WaitForReplyAsync();
            var first = pipe.Replies()[0];
            Check.Equal(2L, first.GetProperty("id").GetInt64(), "A fast setting must receive its own reply before the long operation completes");
            Check.Equal("theme", first.GetProperty("result").GetProperty("applied").GetString());
            Check.Equal(method + ",set", string.Join(",", order), "Commands must still begin on the dispatcher in wire order");
        }
        finally { release.TrySetResult(new { completed = true }); }
        await pipe.Completion.WaitAsync(TimeSpan.FromSeconds(5));
        Check.Equal(1L, pipe.Replies()[1].GetProperty("id").GetInt64(), "The original operation must keep its id and deliver its result later");
    }

    private static string Request(long id, string method, object? parameters = null) =>
        JsonSerializer.Serialize(new { id, method, @params = parameters }) + "\n";

    private sealed class PipeFixture : IDisposable
    {
        private readonly string _directory = Path.Combine(Directory.GetCurrentDirectory(), "_temp", "desktop-pipe-" + Guid.NewGuid().ToString("N"));
        private readonly StaTaskScheduler _scheduler = new();
        private readonly string _results;
        public Process Child { get; }
        public Task Completion { get; }
        public bool ProcessReleased { get; private set; }

        public PipeFixture(string requests, int replies, Func<DesktopCommand, Task<object>> execute)
        {
            Directory.CreateDirectory(_directory);
            _results = Path.Combine(_directory, "replies.jsonl");
            var start = new ProcessStartInfo(Path.ChangeExtension(typeof(DesktopUiProcessTests).Assembly.Location, ".exe"))
            {
                UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            start.ArgumentList.Add("--desktop-pipe-peer");
            start.ArgumentList.Add(_results);
            start.ArgumentList.Add(replies.ToString(System.Globalization.CultureInfo.InvariantCulture));
            start.Environment["WINTAB_PIPE_TEST_REQUESTS"] = requests;
            Child = Process.Start(start) ?? throw new IOException("The background pipe peer did not start");
            Completion = Task.Factory.StartNew(() =>
            {
                var ui = new DesktopUiProcess(Dispatcher.CurrentDispatcher, execute, pid => ProcessReleased = pid == 0);
                // Attach the owned peer to the actual lifecycle, without Show() granting foreground rights.
                typeof(DesktopUiProcess).GetField("_process", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(ui, Child);
                return (Task)typeof(DesktopUiProcess).GetMethod("RunAsync", BindingFlags.Instance | BindingFlags.NonPublic)!
                    .Invoke(ui, new object[] { Child, Channel.CreateBounded<string>(128) })!;
            }, CancellationToken.None, TaskCreationOptions.None, _scheduler).Unwrap();
        }

        public JsonElement[] Replies()
        {
            if (!File.Exists(_results)) return [];
            var replies = new List<JsonElement>();
            foreach (var line in File.ReadAllLines(_results, Encoding.UTF8))
                if (line.Length != 0) { using var document = JsonDocument.Parse(line); replies.Add(document.RootElement.Clone()); }
            return replies.ToArray();
        }

        public async Task WaitForReplyAsync()
        {
            var deadline = Stopwatch.StartNew();
            while (Replies().Length == 0 && deadline.Elapsed < TimeSpan.FromSeconds(5))
                await Task.Delay(20);
            Check.That(Replies().Length > 0, "The pipe peer did not receive the expected response");
        }

        public void Dispose()
        {
            try { if (!Child.HasExited) Child.Kill(entireProcessTree: true); }
            catch (InvalidOperationException) { /* The production lifecycle has already disposed this peer. */ }
            try { Completion.Wait(TimeSpan.FromSeconds(5)); }
            catch (AggregateException) { /* The test reports the original bridge failure through its awaited task. */ }
            Child.Dispose();
            _scheduler.Dispose();
            TestCleanup.DeleteDirectory(_directory);
        }
    }
}
