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
/// A concealed merge source is still a shown window, so the taskbar would count it as a second Explorer
/// window until Explorer has closed it. These tests watch the real taskbar: while a window is concealed it
/// has no button, and the button is back once the window is restored.
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
        yield return ("taskbar recovery retains minimized placement and does not activate the window", RecoveryPreservesPlacement);
        yield return ("a concealed window has no taskbar button until it is restored", ConcealedWindowHasNoButton);
        yield return ("a frame concealed before it is shown gets no taskbar button when Explorer shows it", ConcealedFrameShownLaterHasNoButton);
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
        var removal = ExplorerWindowVisibility.Hide(firstIdentity, _ =>
        {
            entered.Set();
            return release.Wait(5_000);
        });
        await ExplorerWindowVisibility.Hide(secondIdentity, _ => true);
        Task? recovery = null;
        try
        {
            Check.That(entered.Wait(2_000), "The first window must have a blocked taskbar request.");
            recovery = Task.Run(() =>
            {
                ExplorerWindowVisibility.Restore(firstIdentity, true, _ => true);
                ExplorerWindowVisibility.Restore(secondIdentity, true, _ => true);
            });
            await Task.WhenAny(recovery, Task.Delay(1_000));
            Check.That(recovery.IsCompleted, "A taskbar request must not hold the recovery caller indefinitely.");
            Check.That((WinApi.GetWindowLong(second.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
                "The next source must regain its opacity before the first taskbar request finishes.");
            Check.That(WindowVisibilitySnapshot.Read(first.Handle) != null,
                "A bounded wait must retain the first window's unfinished recovery record.");
        }
        finally
        {
            release.Set();
            await removal;
            if (recovery != null) await recovery;
            ExplorerWindowVisibility.Restore(firstIdentity, true, _ => true);
            ExplorerWindowVisibility.Restore(secondIdentity, true, _ => true);
        }
    }

    private static Task RemovalFailureIsRetried() => WithTaskbarWindow(async (window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        var attempts = 0;
        bool Remove(nint _) => ++attempts > 1;
        await ExplorerWindowVisibility.Hide(identity, Remove);
        await ExplorerWindowVisibility.Hide(identity, Remove);
        await ExplorerWindowVisibility.Hide(identity, Remove);
        Check.Equal(2, attempts, "Removal must retry failures but stop repeating after success.");
        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Test concealment must be restored.");
    });

    private static async Task BusyRemovalDoesNotBlockConcealment()
    {
        using var window = new RemoteExplorerFrame(visible: false);
        var identity = WindowIdentity.Capture(window.Handle);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var calls = 0;
        var removal = ExplorerWindowVisibility.Hide(identity, _ =>
        {
            Interlocked.Increment(ref calls);
            entered.Set();
            return release.Wait(5_000);
        });
        try
        {
            Check.That(entered.Wait(2_000), "The removal worker must enter the simulated busy taskbar.");
            Check.That(!removal.IsCompleted, "Taskbar COM must still be busy when Hide returns.");
            Check.That(WinApi.GetLayeredWindowAttributes(window.Handle, out _, out var firstAlpha, out _) && firstAlpha == 0,
                "Native opacity must already be zero before the taskbar answers.");
            WinApi.SetLayeredWindowAttributes(window.Handle, 0, 255, WinApi.LWA_ALPHA);
            var pulse = Task.Run(() =>
            {
                ExplorerWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("Duplicate taskbar call"));
            });
            await pulse.WaitAsync(TimeSpan.FromSeconds(1));
            Check.That(WinApi.GetLayeredWindowAttributes(window.Handle, out _, out var alpha, out _) && alpha == 0,
                "An in-flight taskbar operation must not hold the opacity lock against another conceal pulse.");
            Check.Equal(1, calls, "Repeated pulses must not accumulate pending COM requests.");
        }
        finally
        {
            release.Set();
            await removal;
            ExplorerWindowVisibility.Restore(identity, true, _ => true);
        }
    }

    private static async Task RecoveryFollowsPendingRemoval()
    {
        using var window = new RemoteExplorerFrame(visible: false);
        var identity = WindowIdentity.Capture(window.Handle);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var sequence = new List<string>();
        var removal = ExplorerWindowVisibility.Hide(identity, _ =>
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
            recovery = Task.Run(() => ExplorerWindowVisibility.Restore(identity, true, _ =>
            {
                lock (sequence) sequence.Add("add");
                return true;
            }));
            var deadline = Environment.TickCount64 + 1_000;
            while ((WinApi.GetWindowLong(window.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) != 0 &&
                   Environment.TickCount64 < deadline)
                await Task.Delay(10);
            Check.That((WinApi.GetWindowLong(window.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
                "Recovery must restore native opacity before waiting for busy taskbar COM.");
            await ExplorerWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("Late delete after recovery"))
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
        ExplorerWindowVisibility.Hide(identity, _ => true);
        Check.That(!ExplorerWindowVisibility.Restore(identity, true, _ => false), "A missing taskbar entry is not successful recovery.");
        Check.That(ExplorerWindowVisibility.Contains(window.Handle), "Recovery must retain the in-memory retry state.");
        Check.That(WindowVisibilitySnapshot.Read(window.Handle) != null, "Recovery must retain the persisted restart recovery record.");
        ExplorerWindowVisibility.Hide(identity, _ => throw new InvalidOperationException("A recovering window must not be hidden again."));
        Check.That((WinApi.GetWindowLong(window.Handle, WinApi.GWL_EXSTYLE) & WinApi.WS_EX_LAYERED) == 0,
            "Taskbar recovery failure must not make the already restored window transparent again.");
        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "A successful retry must finish recovery.");
        Check.That(!ExplorerWindowVisibility.Contains(window.Handle) && WindowVisibilitySnapshot.Read(window.Handle) == null,
            "Only successful recovery may discard both recovery records.");
        return Task.CompletedTask;
    });

    private static Task RestorationFailureSurvivesShellRetirement() => WithTaskbarWindow((window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        ExplorerWindowVisibility.Hide(identity, _ => true);
        var snapshot = WindowVisibilitySnapshot.Read(window.Handle);
        Check.That(snapshot != null, "Concealment must persist a recovery record.");
        Check.That(!ExplorerWindowVisibility.Restore(identity, true, _ => false), "Taskbar recovery must fail before shell retirement.");

        // Shell cleanup retires the COM registration, not the still-live native window. Repeated
        // cleanup passes must not invalidate the identity shared by the pending recovery record.
        for (var attempt = 0; attempt < 2; attempt++)
        {
            ExplorerWindowVisibility.ReleaseIdentityIfUntracked(identity);
            Check.That(identity.IsCurrent && ExplorerWindowVisibility.Contains(window.Handle),
                "Shell retirement must keep the live window's pending recovery identity.");
            var addAttempted = false;
            Check.That(!ExplorerWindowVisibility.Restore(identity, true, _ => { addAttempted = true; return false; }),
                "An unsuccessful taskbar retry must remain unsuccessful.");
            Check.That(addAttempted, "Recovery must retry AddTab instead of forgetting a retired identity.");
            Check.Equal(snapshot, WindowVisibilitySnapshot.Read(window.Handle),
                "Failed recovery after shell retirement must retain the original persisted snapshot.");
        }

        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Taskbar recovery must succeed once AddTab succeeds.");
        Check.That(!ExplorerWindowVisibility.Contains(window.Handle) && WindowVisibilitySnapshot.Read(window.Handle) == null,
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

    private static Task RecoveryPreservesPlacement() => WithTaskbarWindow(async (window, _) =>
    {
        WinApi.ShowWindow(window.Handle, 7); // SW_SHOWMINNOACTIVE
        Check.That(WinApi.IsIconic(window.Handle), "The owned test window must start minimized.");
        var foreground = WinApi.GetForegroundWindow();
        var identity = WindowIdentity.Capture(window.Handle);
        await ExplorerWindowVisibility.Hide(identity, _ => true);
        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Recover the taskbar style.");
        Check.That(WinApi.IsIconic(window.Handle), "Re-registration must not unminimize the window.");
        Check.Equal(foreground, WinApi.GetForegroundWindow(), "Recovery must not steal foreground.");
    });

    private static Task ConcealedWindowHasNoButton() => WithTaskbarWindow(async (window, title) =>
    {
        if (!TaskbarPresent())
            throw new TestSkippedException("No taskbar is available in this desktop session");
        window.Show();
        Check.That(await WaitForButtonAsync(title, present: true), "The owned test window must first appear on the available taskbar.");

        ExplorerWindowVisibility.Hide(window.Handle);
        Check.That(await WaitForButtonAsync(title, present: false), "A concealed window must not keep a taskbar button.");
        ExplorerWindowVisibility.Hide(window.Handle);
        await Task.Delay(300);
        Check.That(!HasButton(title), "Concealing again must not bring the button back.");

        Check.That(ExplorerWindowVisibility.Restore(window.Handle), "The owned window must be restored.");
        Check.That(await WaitForButtonAsync(title, present: true), "A restored window must get its taskbar button back.");
    });

    private static Task ConcealedFrameShownLaterHasNoButton() => WithTaskbarWindow(async (window, title) =>
    {
        if (!TaskbarPresent())
            throw new TestSkippedException("No taskbar is available in this desktop session");

        // Windows 11 preloads a hidden frame; WinTab conceals it before Explorer shows it for a folder.
        ExplorerWindowVisibility.Hide(window.Handle);
        window.Show();
        await Task.Delay(1_000);
        Check.That(!HasButton(title), "Showing a concealed frame must not add a taskbar button.");

        Check.That(ExplorerWindowVisibility.Restore(window.Handle), "The owned window must be restored.");
        Check.That(await WaitForButtonAsync(title, present: true),
            "A frame first shown while concealed must gain its taskbar button when restored.");
    });

    private static async Task WithTaskbarWindow(Func<TaskbarWindow, string, Task> body)
    {
        using var scheduler = new StaTaskScheduler();
        await Task.Factory.StartNew(async () =>
        {
            SetCurrentProcessExplicitAppUserModelID(AppId);
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

    private static async Task<bool> WaitForButtonAsync(string title, bool present, int timeoutMs = 3_000)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Environment.TickCount64 < deadline)
        {
            if (HasButton(title) == present)
                return true;
            await Task.Delay(50);
        }
        return HasButton(title) == present;
    }

    private static AutomationElement? Tray() =>
        AutomationElement.RootElement.FindFirst(TreeScope.Children,
            new PropertyCondition(AutomationElement.ClassNameProperty, "Shell_TrayWnd"));

    private static bool TaskbarPresent() => Tray() != null;

    private static bool HasButton(string title)
    {
        var tray = Tray();
        if (tray == null)
            return false;
        var buttons = tray.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
        foreach (AutomationElement button in buttons)
        {
            var current = button.Current;
            if (current.AutomationId == "Appid: " + AppId ||
                current.Name.Contains(title, StringComparison.Ordinal))
                return true;
        }
        return false;
    }

    /// <summary>An ordinary top-level window kept off screen; unlike the tab test windows it is not a tool window, so the taskbar lists it.</summary>
    private sealed class TaskbarWindow : IDisposable
    {
        private readonly HwndSource _host;

        public TaskbarWindow(string title)
        {
            _host = new HwndSource(new HwndSourceParameters(title)
            {
                WindowStyle = 0x00CF0000,
                ExtendedWindowStyle = 0,
                PositionX = -32000,
                PositionY = -32000,
                Width = 120,
                Height = 80
            });
        }

        public nint Handle => _host.Handle;

        public void Show() => WinApi.ShowWindow(Handle, WinApi.SW_SHOWNOACTIVATE);

        public void Dispose() => _host.Dispose();
    }

    [DllImport("shell32.dll")]
    private static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string appId);
}
