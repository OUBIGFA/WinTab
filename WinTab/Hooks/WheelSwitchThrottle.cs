using System;

namespace WinTab.Hooks;

public enum WheelSwitchSensitivity
{
    Low,
    Medium,
    High
}

/// <summary>
/// Turns raw wheel deltas into tab steps. Each sensitivity level sets how many notches make one switch and
/// how long a switch holds before the next one may start. Notches that arrive during that hold are dropped
/// rather than queued, so a quick flick never lands several tabs further than the user meant.
/// </summary>
internal sealed class WheelSwitchThrottle
{
    private const int WheelDelta = 120;
    private const int GestureGapMs = 400;
    // Far enough in the past to never fall inside a window, yet safe to subtract from without overflow.
    private const long Never = long.MinValue / 2;

    private long _lastWheelAt = Never;
    private long _lastSwitchAt = Never;
    private int _accumulated;

    internal static (int Notches, int HoldMs) Profile(WheelSwitchSensitivity sensitivity) => sensitivity switch
    {
        WheelSwitchSensitivity.High => (1, 60),
        WheelSwitchSensitivity.Low => (2, 350),
        _ => (1, 200)
    };

    public void Reset()
    {
        _accumulated = 0;
        _lastWheelAt = Never;
        _lastSwitchAt = Never;
    }

    /// <summary>Returns -1 for the previous tab, +1 for the next, or 0 when this delta does not switch.</summary>
    public int Accept(int delta, long now, WheelSwitchSensitivity sensitivity)
    {
        if (delta == 0)
            return 0;

        var (notches, holdMs) = Profile(sensitivity);

        // A pause ends the gesture, and a reversal starts a new one: leftovers must not carry over.
        if (now - _lastWheelAt > GestureGapMs || Math.Sign(delta) != Math.Sign(_accumulated))
            _accumulated = 0;
        _lastWheelAt = now;

        if (now - _lastSwitchAt < holdMs)
        {
            _accumulated = 0;
            return 0;
        }

        _accumulated += delta;
        if (Math.Abs(_accumulated) < notches * WheelDelta)
            return 0;

        _accumulated = 0;
        _lastSwitchAt = now;
        // Wheel up (positive delta) goes to the tab on the left.
        return delta > 0 ? -1 : 1;
    }
}
