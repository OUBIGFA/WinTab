using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Interop;
using WinTab.Helpers;
using WinTab.WinAPI;

/// <summary>
/// Exercises taskbar operation results and recovery ordering with supplied operations and message-only handles.
/// No test creates a taskbar button or changes the desktop foreground.
/// </summary>
internal static class TaskbarButtonTests
{
    private static readonly string AppId = "WinTab.Tests.TaskbarButton." + Guid.NewGuid().ToString("N");

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("taskbar operations work when HrInit is not implemented", () => InitializationResult(unchecked((int)0x80004001), 0, true));
        yield return ("taskbar operations follow successful initialization", () => InitializationResult(0, 0, true));
        yield return ("taskbar operations follow nonzero successful initialization", () => InitializationResult(1, 0, true));
        yield return ("taskbar initialization failure prevents the operation", () => InitializationResult(unchecked((int)0x80070005), 0, false));
        yield return ("taskbar operation failure is not hidden by HrInit compatibility", () => InitializationResult(unchecked((int)0x80004001), unchecked((int)0x80004005), true));
        yield return ("taskbar removal failure is retried on the next conceal pulse", RemovalFailureIsRetried);
        yield return ("busy taskbar removal cannot block native concealment or its next pulse", BusyRemovalDoesNotBlockConcealment);
        yield return ("taskbar recovery follows in-flight removal without blocking native concealment", RecoveryFollowsPendingRemoval);
        yield return ("taskbar restoration failure retains the recovery record until retry succeeds", RestorationFailureIsRetried);
        yield return ("taskbar restoration failure survives shell identity retirement", RestorationFailureSurvivesShellRetirement);
        yield return ("shell identity retirement releases windows without recovery records", ShellRetirementReleasesUntrackedIdentity);
        yield return ("busy taskbar recovery cannot hold up recovery of another window", BusyRecoveryDoesNotBlockOtherWindows);
    }

    private static Task InitializationResult(int initialization, int operation, bool expectedCall)
    {
        var called = false;
        var result = TaskbarButton.InitializeAndInvoke(() => initialization, () =>
        {
            called = true;
            return operation;
        });
        Check.Equal(expectedCall, called, "Only success or E_NOTIMPL from HrInit permits the operation.");
        Check.Equal(expectedCall ? operation : initialization, result, "The actual operation's failure must remain observable.");
        return Task.CompletedTask;
    }

    private static async Task BusyRecoveryDoesNotBlockOtherWindows()
    {
        using var first = new RemoteExplorerFrame(visible: false);
        using var second = new RemoteExplorerFrame(visible: false);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var firstIdentity = WindowIdentity.Capture(first.Handle);
        var secondIdentity = WindowIdentity.Capture(second.Handle);
        var removal = BackgroundWindowVisibility.Hide(firstIdentity, _ =>
        {
            entered.Set();
            return release.Wait(5_000);
        });
        await BackgroundWindowVisibility.Hide(secondIdentity, _ => true);
        Task? recovery = null;
        try
        {
            Check.That(entered.Wait(2_000), "The first window must have a blocked taskbar request.");
            recovery = Task.Run(() =>
            {
                BackgroundWindowVisibility.Restore(firstIdentity, true, _ => true);
                BackgroundWindowVisibility.Restore(secondIdentity, true, _ => true);
            });
            await Task.WhenAny(recovery, Task.Delay(1_000));
            Check.That(recovery.IsCompleted, "A taskbar request must not hold the recovery caller indefinitely.");
            Check.That((TestWindowOpacity.Instance.ReadStyle(second.Handle) & WinApi.WS_EX_LAYERED) == 0,
                "The next source must regain its opacity before the first taskbar request finishes.");
            Check.That(WindowVisibilitySnapshot.Read(first.Handle, TestWindowOpacity.Instance) != null,
                "A bounded wait must retain the first window's unfinished recovery record.");
        }
        finally
        {
            release.Set();
            await removal;
            if (recovery != null) await recovery;
            BackgroundWindowVisibility.Restore(firstIdentity, true, _ => true);
            BackgroundWindowVisibility.Restore(secondIdentity, true, _ => true);
        }
    }

    private static Task RemovalFailureIsRetried() => WithTaskbarWindow(async (window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        var attempts = 0;
        bool Remove(nint _) => ++attempts > 1;
        await BackgroundWindowVisibility.Hide(identity, Remove);
        await BackgroundWindowVisibility.Hide(identity, Remove);
        await BackgroundWindowVisibility.Hide(identity, Remove);
        Check.Equal(2, attempts, "Removal must retry failures but stop repeating after success.");
        Check.That(BackgroundWindowVisibility.Restore(identity, true, _ => true), "Test concealment must be restored.");
    });

    private static async Task BusyRemovalDoesNotBlockConcealment()
    {
        using var window = new RemoteExplorerFrame(visible: false);
        var identity = WindowIdentity.Capture(window.Handle);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var calls = 0;
        var removal = BackgroundWindowVisibility.Hide(identity, _ =>
        {
            Interlocked.Increment(ref calls);
            entered.Set();
            return release.Wait(5_000);
        });
        try
        {
            Check.That(entered.Wait(2_000), "The removal worker must enter the simulated busy taskbar.");
            Check.That(!removal.IsCompleted, "Taskbar COM must still be busy when Hide returns.");
            Check.That(TestWindowOpacity.Instance.TryRead(window.Handle, out _, out var firstAlpha, out _) && firstAlpha == 0,
                "Native opacity must already be zero before the taskbar answers.");
            TestWindowOpacity.Instance.TryWrite(window.Handle, 0, 255, WinApi.LWA_ALPHA);
            var pulse = Task.Run(() =>
            {
                BackgroundWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("Duplicate taskbar call"));
            });
            await pulse.WaitAsync(TimeSpan.FromSeconds(1));
            Check.That(TestWindowOpacity.Instance.TryRead(window.Handle, out _, out var alpha, out _) && alpha == 0,
                "An in-flight taskbar operation must not hold the opacity lock against another conceal pulse.");
            Check.Equal(1, calls, "Repeated pulses must not accumulate pending COM requests.");
        }
        finally
        {
            release.Set();
            await removal;
            BackgroundWindowVisibility.Restore(identity, true, _ => true);
        }
    }

    private static async Task RecoveryFollowsPendingRemoval()
    {
        using var window = new RemoteExplorerFrame(visible: false);
        var identity = WindowIdentity.Capture(window.Handle);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var sequence = new List<string>();
        var removal = BackgroundWindowVisibility.Hide(identity, _ =>
        {
            entered.Set();
            release.Wait(5_000);
            lock (sequence) sequence.Add("delete");
            return true;
        });
        Task<bool>? recovery = null;
        try
        {
            Check.That(entered.Wait(2_000), "The simulated DeleteTab must be in flight.");
            recovery = Task.Run(() => BackgroundWindowVisibility.Restore(identity, true, _ =>
            {
                lock (sequence) sequence.Add("add");
                return true;
            }));
            var deadline = Environment.TickCount64 + 1_000;
            while ((TestWindowOpacity.Instance.ReadStyle(window.Handle) & WinApi.WS_EX_LAYERED) != 0 &&
                   Environment.TickCount64 < deadline)
                await Task.Delay(10);
            Check.That((TestWindowOpacity.Instance.ReadStyle(window.Handle) & WinApi.WS_EX_LAYERED) == 0,
                "Recovery must restore native opacity before waiting for busy taskbar COM.");
            await BackgroundWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("Late delete after recovery"))
                .WaitAsync(TimeSpan.FromSeconds(1));
            Check.That(!recovery.IsCompleted, "AddTab must wait until the in-flight DeleteTab has finished.");
        }
        finally
        {
            release.Set();
            await removal;
            if (recovery != null)
                Check.That(await recovery, "Recovery must finish once the taskbar answers.");
        }
        Check.Equal("delete,add", string.Join(',', sequence), "A late delete must never remove the restored taskbar button.");
        Check.That(!ExplorerWindowVisibility.Contains(window.Handle), "Ordered recovery must release its state.");
    }

    private static Task RestorationFailureIsRetried() => WithTaskbarWindow((window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        BackgroundWindowVisibility.Hide(identity, _ => true);
        Check.That(!BackgroundWindowVisibility.Restore(identity, true, _ => false), "A missing taskbar entry is not successful recovery.");
        Check.That(ExplorerWindowVisibility.Contains(window.Handle), "Recovery must retain the in-memory retry state.");
        Check.That(WindowVisibilitySnapshot.Read(window.Handle, TestWindowOpacity.Instance) != null, "Recovery must retain the persisted restart recovery record.");
        BackgroundWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("A recovering window must not be hidden again."));
        Check.That((TestWindowOpacity.Instance.ReadStyle(window.Handle) & WinApi.WS_EX_LAYERED) == 0,
            "Taskbar recovery failure must not make the already restored window transparent again.");
        Check.That(BackgroundWindowVisibility.Restore(identity, true, _ => true), "A successful retry must finish recovery.");
        Check.That(!ExplorerWindowVisibility.Contains(window.Handle) && WindowVisibilitySnapshot.Read(window.Handle, TestWindowOpacity.Instance) == null,
            "Only successful recovery may discard both recovery records.");
        return Task.CompletedTask;
    });

    private static Task RestorationFailureSurvivesShellRetirement() => WithTaskbarWindow((window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        BackgroundWindowVisibility.Hide(identity, _ => true);
        var snapshot = WindowVisibilitySnapshot.Read(window.Handle, TestWindowOpacity.Instance);
        Check.That(snapshot != null, "Concealment must persist a recovery record.");
        Check.That(!BackgroundWindowVisibility.Restore(identity, true, _ => false), "Taskbar recovery must fail before shell retirement.");

        // Shell cleanup retires the COM registration, not the still-live native window. Repeated
        // cleanup passes must not invalidate the identity shared by the pending recovery record.
        for (var attempt = 0; attempt < 2; attempt++)
        {
            ExplorerWindowVisibility.ReleaseIdentityIfUntracked(identity);
            Check.That(identity.IsCurrent && ExplorerWindowVisibility.Contains(window.Handle),
                "Shell retirement must keep the live window's pending recovery identity.");
            var addAttempted = false;
            Check.That(!BackgroundWindowVisibility.Restore(identity, true, _ => { addAttempted = true; return false; }),
                "An unsuccessful taskbar retry must remain unsuccessful.");
            Check.That(addAttempted, "Recovery must retry AddTab instead of forgetting a retired identity.");
            Check.Equal(snapshot, WindowVisibilitySnapshot.Read(window.Handle, TestWindowOpacity.Instance),
                "Failed recovery after shell retirement must retain the original persisted snapshot.");
        }

        Check.That(BackgroundWindowVisibility.Restore(identity, true, _ => true), "Taskbar recovery must succeed once AddTab succeeds.");
        Check.That(!ExplorerWindowVisibility.Contains(window.Handle) && WindowVisibilitySnapshot.Read(window.Handle, TestWindowOpacity.Instance) == null,
            "Only successful recovery may discard both recovery records.");
        ExplorerWindowVisibility.ReleaseIdentityIfUntracked(identity);
        Check.That(!identity.IsCurrent, "Shell cleanup may release an identity once recovery no longer owns it.");
        return Task.CompletedTask;
    });

    private static Task ShellRetirementReleasesUntrackedIdentity() => WithTaskbarWindow((window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        Check.That(identity.IsCurrent, "The window must start with a current identity.");
        ExplorerWindowVisibility.ReleaseIdentityIfUntracked(identity);
        Check.That(!identity.IsCurrent, "A registration without pending recovery must still release its identity.");
        return Task.CompletedTask;
    });

    private static async Task WithTaskbarWindow(Func<TaskbarWindow, string, Task> body)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            var title = "WinTab taskbar test " + Guid.NewGuid().ToString("N")[..8];
            using var window = new TaskbarWindow(title);
            try
            {
                await body(window, title);
            }
            finally
            {
                ExplorerWindowVisibility.Forget(window.Handle);
            }
        }, CancellationToken.None, TaskCreationOptions.None, scheduler).Unwrap();
    }

    /// <summary>A property-bearing message endpoint that cannot appear on the desktop or taskbar.</summary>
    private sealed class TaskbarWindow : IDisposable
    {
        private readonly HwndSource _host;

        public TaskbarWindow(string title)
        {
            _host = new HwndSource(new HwndSourceParameters(title)
            {
                ParentWindow = (nint)(-3),
                WindowStyle = 0,
                Width = 120,
                Height = 80
            });
        }

        public nint Handle => _host.Handle;

        public void Dispose() => _host.Dispose();
    }

    [DllImport("shell32.dll")]
    private static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string appId);
}
