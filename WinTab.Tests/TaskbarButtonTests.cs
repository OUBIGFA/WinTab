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
    private const string AppId = "WinTab.Tests.TaskbarButton";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("taskbar removal failure is retried on the next conceal pulse", RemovalFailureIsRetried);
        yield return ("taskbar restoration failure retains the recovery record until retry succeeds", RestorationFailureIsRetried);
        yield return ("taskbar restoration failure survives shell identity retirement", RestorationFailureSurvivesShellRetirement);
        yield return ("shell identity retirement releases windows without recovery records", ShellRetirementReleasesUntrackedIdentity);
        yield return ("a concealed window has no taskbar button until it is restored", ConcealedWindowHasNoButton);
        yield return ("a frame concealed before it is shown gets no taskbar button when Explorer shows it", ConcealedFrameShownLaterHasNoButton);
    }

    private static Task RemovalFailureIsRetried() => WithTaskbarWindow((window, _) =>
    {
        var identity = WindowIdentity.Capture(window.Handle);
        var attempts = 0;
        bool Remove(nint _) => ++attempts > 1;
        ExplorerWindowVisibility.Hide(identity, Remove);
        ExplorerWindowVisibility.Hide(identity, Remove);
        ExplorerWindowVisibility.Hide(identity, Remove);
        Check.Equal(2, attempts, "Removal must retry failures but stop repeating after success.");
        Check.That(ExplorerWindowVisibility.Restore(identity, true, _ => true), "Test concealment must be restored.");
        return Task.CompletedTask;
    });

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

    private static Task ConcealedWindowHasNoButton() => WithTaskbarWindow(async (window, title) =>
    {
        window.Show();
        if (!await WaitForButtonAsync(title, present: true))
        {
            Console.WriteLine("    note: the taskbar does not expose this window's button here; the check cannot run.");
            return;
        }

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
        {
            Console.WriteLine("    note: no taskbar in this session; the check cannot run.");
            return;
        }

        // Windows 11 preloads a hidden frame; WinTab conceals it before Explorer shows it for a folder.
        ExplorerWindowVisibility.Hide(window.Handle);
        window.Show();
        await Task.Delay(1_000);
        Check.That(!HasButton(title), "Showing a concealed frame must not add a taskbar button.");

        Check.That(ExplorerWindowVisibility.Restore(window.Handle), "The owned window must be restored.");
        if (!await WaitForButtonAsync(title, present: true))
        {
            Console.WriteLine("    note: the taskbar does not expose this window's button here; the restore check cannot run.");
        }
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
