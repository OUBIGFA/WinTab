using WinTab.Helpers;
using WinTab.WinAPI;

namespace WinTab.Hooks;

/// <summary>A native foreground activation that may briefly hand file-view focus to Explorer's toolbar.</summary>
internal sealed record ExplorerNativeFocusActivation(WindowIdentity Window, uint EventTime, uint LastInputTime)
{
    internal bool CanFollowFocus(nint parent, nint focus, uint focusEventTime, uint now, uint lastInputTime)
    {
        // Out-of-context focus events can arrive after Explorer has restored keyboard focus to its
        // XAML toolbar. Accept only that short activation handoff, with no intervening user input.
        const uint activationHandoffMs = 250;
        if (Window.Handle != parent || !Window.IsCurrent || LastInputTime != lastInputTime ||
            unchecked((int)(focusEventTime - lastInputTime)) < 0 ||
            unchecked(focusEventTime - EventTime) > activationHandoffMs ||
            unchecked(now - EventTime) > activationHandoffMs ||
            unchecked(now - focusEventTime) > activationHandoffMs ||
            focus == 0 || WinApi.GetAncestor(focus, WinApi.GA_ROOT) != parent)
            return false;

        if (!WinApi.IsWindowHasClassName(focus, "InputSiteWindowClass") &&
            !WinApi.IsWindowHasClassName(focus, "Microsoft.UI.Content.DesktopChildSiteBridge"))
            return false;

        // A similarly named control inside a tab belongs to that tab, not the frame toolbar.
        for (var ancestor = WinApi.GetParent(focus); ancestor != 0 && ancestor != parent; ancestor = WinApi.GetParent(ancestor))
        {
            if (WinApi.IsWindowHasClassName(ancestor, "ShellTabWindowClass"))
                return false;
        }
        return true;
    }
}
