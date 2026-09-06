using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;
using WinTab.WinAPI;

internal static class WindowSafetyTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("window identities reject an old registration on the same thread and handle", RejectsOldIdentity);
        yield return ("window commands time out without waiting for an unresponsive window", WindowCommandsAreBounded);
        yield return ("completed merges cannot perform late window operations", DisposedMergeCannotAct);
        yield return ("merge generations and deadlines invalidate pending work", MergeLifetimeIsBounded);
        yield return ("window recovery restores only the opacity it changed", RecoveryPreservesOpacity);
        yield return ("window recovery rejects a replacement with the same handle", RecoveryRejectsReplacement);
    }

    private static async Task RejectsOldIdentity()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            var handle = CreateTestWindow();
            try
            {
                var original = WindowIdentity.Capture(handle);
                Check.That(original.IsCurrent, "The original registration must be valid.");
                original.Release();
                var replacement = WindowIdentity.Capture(handle);
                Check.That(replacement.IsCurrent && !original.IsCurrent,
                    "A new registration must invalidate the old token, even with unchanged process and thread IDs.");
                replacement.Release();
            }
            finally
            {
                DestroyWindow(handle);
            }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static async Task WindowCommandsAreBounded()
    {
        using var scheduler = new StaTaskScheduler();
        var handle = await Task.Factory.StartNew(CreateTestWindow, CancellationToken.None, TaskCreationOptions.None, scheduler);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var blocked = Task.Factory.StartNew(() =>
        {
            started.Set();
            WaitForSingleObject(release.WaitHandle.SafeWaitHandle.DangerousGetHandle(), 5_000);
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
        try
        {
            Check.That(started.Wait(2_000), "The test window thread must be blocked.");
            var timer = Stopwatch.StartNew();
            Check.That(!WinApi.TrySendMessage(handle, WinApi.WM_COMMAND, 0, 0),
                "An unresponsive window must be reported as a failed command.");
            Check.That(timer.ElapsedMilliseconds < 800, "The native command must honor its 200 ms timeout.");
        }
        finally
        {
            release.Set();
            await blocked;
            await Task.Factory.StartNew(() => DestroyWindow(handle), CancellationToken.None, TaskCreationOptions.None, scheduler);
        }
    }

    private static Task RecoveryPreservesOpacity() => WithVisibilityWindow(handle =>
    {
        ExplorerWindowVisibility.UpdateLayeredStyle(handle, remove: false);
        Check.That(WinApi.SetLayeredWindowAttributes(handle, 0, 128, WinApi.LWA_ALPHA), "Set up the original opacity.");
        var identity = WindowIdentity.Capture(handle);
        ExplorerWindowVisibility.Hide(identity);
        ExplorerWindowVisibility.Hide(identity);
        Check.That(WinApi.GetLayeredWindowAttributes(handle, out _, out var hiddenAlpha, out _) && hiddenAlpha == 0,
            "A merge must conceal its own source window.");
        Check.That(ExplorerWindowVisibility.Restore(identity), "The owned source must be restored.");
        Check.That(WinApi.GetLayeredWindowAttributes(handle, out _, out var restoredAlpha, out _) && restoredAlpha == 128,
            "Recovery must preserve the original opacity rather than remove unrelated window styling.");
        Check.That(!WinApi.IsWindowVisible(handle), "Recovery must not show a window that was already hidden.");
        Check.That(!ExplorerWindowVisibility.Contains(handle), "Successful recovery must release its tracking entry.");
    });

    private static Task RecoveryRejectsReplacement() => WithVisibilityWindow(handle =>
    {
        var original = WindowIdentity.Capture(handle);
        ExplorerWindowVisibility.Hide(original);
        original.Release();
        var replacement = WindowIdentity.Capture(handle);
        Check.That(WinApi.SetLayeredWindowAttributes(handle, 0, 192, WinApi.LWA_ALPHA), "Set up replacement opacity.");
        Check.That(!ExplorerWindowVisibility.Restore(original), "A stale registration must not restore another window.");
        Check.That(WinApi.GetLayeredWindowAttributes(handle, out _, out var alpha, out _) && alpha == 192,
            "A replacement window must keep its own opacity.");
        replacement.Release();
    });

    private static async Task WithVisibilityWindow(Action<nint> action)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            var handle = CreateWindowEx(0, "STATIC", "WinTab isolated opacity test", 0, 0, 0, 20, 20, 0, 0, 0, 0);
            Check.That(handle != 0, "The test must create its own hidden window.");
            try { action(handle); }
            finally
            {
                ExplorerWindowVisibility.Forget(handle);
                DestroyWindow(handle);
            }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static async Task DisposedMergeCannotAct()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(() =>
        {
            var handle = CreateTestWindow();
            try
            {
                using var operation = new MergeOperation(WindowIdentity.Capture(handle), 1,
                    CancellationToken.None, () => true, 5_000);
                Check.That(operation.IsCurrent, "An active merge must own its source window.");
                operation.Dispose();
                Check.That(!operation.IsCurrent, "A completed merge must invalidate every late callback.");
                Check.Throws<OperationCanceledException>(operation.ThrowIfInvalid,
                    "A late operation must be rejected after disposal.");
            }
            finally { DestroyWindow(handle); }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler);
    }

    private static async Task MergeLifetimeIsBounded()
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            var handle = CreateTestWindow();
            try
            {
                var currentGeneration = 1;
                using var lifetime = new CancellationTokenSource();
                using var operation = new MergeOperation(WindowIdentity.Capture(handle), 1, lifetime.Token,
                    () => currentGeneration == 1, 40);
                currentGeneration++;
                Check.That(!operation.IsCurrent, "A generation change must reject old work.");
                Check.Throws<OperationCanceledException>(operation.ThrowIfInvalid,
                    "Old callbacks must not act on a new generation.");
                currentGeneration = 1;
                try { await Task.Delay(Timeout.Infinite, operation.Token).WaitAsync(TimeSpan.FromSeconds(1)); }
                catch (OperationCanceledException) { }
                Check.That(!operation.IsCurrent, "The absolute deadline must invalidate the merge.");
            }
            finally { DestroyWindow(handle); }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    private static nint CreateTestWindow()
    {
        var handle = CreateWindowEx(0, "STATIC", "WinTab isolated test", 0, 0, 0, 10, 10, (nint)(-3), 0, 0, 0);
        Check.That(handle != 0, "The isolated message-only window must be created.");
        return handle;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string name, uint style,
        int left, int top, int width, int height, nint parent, nint menu, nint instance, nint parameter);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(nint handle);

    [DllImport("kernel32.dll")]
    private static extern uint WaitForSingleObject(nint handle, uint milliseconds);
}
