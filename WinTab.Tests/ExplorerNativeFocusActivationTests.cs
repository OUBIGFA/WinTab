using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WinTab.Helpers;
using WinTab.Hooks;

internal static class ExplorerNativeFocusActivationTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("activation focus follows a fresh folder request through the XAML bridge", () => AcceptsToolbarFocus("Microsoft.UI.Content.DesktopChildSiteBridge"));
        yield return ("activation focus follows a fresh folder request through the input site", () => AcceptsToolbarFocus("InputSiteWindowClass"));
        yield return ("activation focus rejects new user input", () => RejectsTiming(1_010, 1_030, 1_020));
        yield return ("activation focus rejects input that overtook the foreground callback", RejectsInputBeforeCallback);
        yield return ("activation focus rejects an event from before foreground activation", () => RejectsTiming(990, 1_030, 900));
        yield return ("activation focus rejects a late folder event", () => RejectsTiming(1_400, 1_410, 900));
        yield return ("activation focus rejects a request delayed in the shell queue", () => RejectsTiming(1_010, 2_000, 900));
        yield return ("activation focus handles the Windows tick counter wrapping", TickCounterWraps);
        yield return ("activation focus rejects another tab's file view", RejectsAnotherView);
        yield return ("activation focus rejects another frame's input site", RejectsAnotherFrame);
        yield return ("activation focus rejects an expired window identity", RejectsRetiredWindow);
        yield return ("activation focus does not guess when native focus is unavailable", RejectsMissingFocus);
    }

    private static Task AcceptsToolbarFocus(string className) => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        // Recorded failure: the folder view emitted focus 15 ms after activation, but GetGUIThreadInfo
        // already pointed at the top-level XAML bridge when the out-of-context callback arrived.
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 900);
        var focus = fixture.Window.CreateFrameInputSite(className);
        Check.That(activation.CanFollowFocus(fixture.Window.Handle, focus, 1_015, 1_030, 900),
            "The native toolbar handoff must not discard the freshly requested folder tab.");
        return Task.CompletedTask;
    });

    private static Task RejectsTiming(uint eventTime, uint now, uint inputTime) => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 900);
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, fixture.Window.CreateFrameInputSite(), eventTime, now, inputTime),
            "Old events, expired activation and newer user input must not authorize a tab switch.");
        return Task.CompletedTask;
    });

    private static Task TickCounterWraps() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), uint.MaxValue - 10, uint.MaxValue - 20);
        Check.That(activation.CanFollowFocus(fixture.Window.Handle, fixture.Window.CreateFrameInputSite(), 4, 20, uint.MaxValue - 20),
            "A recent focus request remains recent across the 32-bit system tick rollover.");
        return Task.CompletedTask;
    });

    private static Task RejectsInputBeforeCallback() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        // Both notifications can queue behind the hook thread. Capturing input only when their callback
        // runs must not make a later user click look like the input that opened this folder.
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 1_020);
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, fixture.Window.CreateFrameInputSite(), 1_015, 1_030, 1_020),
            "Input newer than the folder event must retire it even if the foreground callback observed that input too.");
        return Task.CompletedTask;
    });

    private static Task RejectsAnotherView() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 900);
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, fixture.Window.CreateFolderView(1), 1_015, 1_030, 900),
            "Focus in another file view is not the frame's temporary toolbar focus.");
        return Task.CompletedTask;
    });

    private static Task RejectsAnotherFrame() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        using var other = new ExplorerTabActivationTests.ActivationWindow();
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 900);
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, other.CreateFrameInputSite(), 1_015, 1_030, 900),
            "A toolbar in another window must not authorize reuse in this frame.");
        return Task.CompletedTask;
    });

    private static Task RejectsRetiredWindow() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var identity = WindowIdentity.Capture(fixture.Window.Handle);
        var activation = new ExplorerNativeFocusActivation(identity, 1_000, 900);
        var focus = fixture.Window.CreateFrameInputSite();
        identity.Release();
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, focus, 1_015, 1_030, 900),
            "A retired frame must not receive a queued activation request.");
        return Task.CompletedTask;
    });

    private static Task RejectsMissingFocus() => ExplorerTabLifetimeTests.WithFixture(fixture =>
    {
        var activation = new ExplorerNativeFocusActivation(WindowIdentity.Capture(fixture.Window.Handle), 1_000, 900);
        Check.That(!activation.CanFollowFocus(fixture.Window.Handle, 0, 1_015, 1_030, 900),
            "Missing focus provides no evidence of a toolbar handoff.");
        return Task.CompletedTask;
    });
}
