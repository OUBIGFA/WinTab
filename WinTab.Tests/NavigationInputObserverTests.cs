using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using WinTab.Hooks;
using WinTab.WinAPI;

/// <summary>Raw Input registrations are process-local; these tests never synthesize user input.</summary>
internal static class NavigationInputObserverTests
{
    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("navigation input observer registers both devices and releases its own registrations", RegistrationLifetime);
        yield return ("navigation input observer repairs a missing device registration", RepairsRegistration);
        yield return ("retiring a navigation input observer cannot unregister its replacement", ReplacementKeepsRegistration);
    }

    private static NavigationInputObserver Observe() => new(_ => { }, () => { }, _ => { });
    private static nint Window(NavigationInputObserver observer) => (nint)typeof(NavigationInputObserver)
        .GetField("_window", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(observer)!;

    private static Task RegistrationLifetime()
    {
        nint target;
        using (var observer = Observe())
        {
            target = Window(observer);
            Check.That(target != 0, "The observer must own a native message window.");
            AssertOwnsBoth(target);
        }
        Check.That(ReadDevices().All(device => device.Target != target),
            "Stopping the observer must remove the raw input registrations of its message window.");
        return Task.CompletedTask;
    }

    private static Task RepairsRegistration()
    {
        using var observer = Observe();
        var target = Window(observer);
        Check.That(RegisterRawInputDevices([new RawDevice { UsagePage = 1, Usage = 2, Flags = 1 }], 1,
            (uint)Marshal.SizeOf<RawDevice>()), "The test removes only its own process's mouse registration.");
        Check.That(!ReadDevices().Any(device => device.UsagePage == 1 && device.Usage == 2 && device.Target == target),
            "The mouse registration must really be missing before exercising recovery.");
        Check.That(WinApi.TrySendMessage(target, 0x0113, 1, 0, 1_000), "The observer must process its registration timer.");
        AssertOwnsBoth(target);
        return Task.CompletedTask;
    }

    private static Task ReplacementKeepsRegistration()
    {
        using var retiring = Observe();
        using var replacement = Observe();
        var target = Window(replacement);
        AssertOwnsBoth(target);
        retiring.Dispose();
        AssertOwnsBoth(target);
        return Task.CompletedTask;
    }

    private static void AssertOwnsBoth(nint target)
    {
        var devices = ReadDevices();
        foreach (ushort usage in new ushort[] { 2, 6 })
            Check.That(devices.Any(device => device.UsagePage == 1 && device.Usage == usage && device.Target == target &&
                    (device.Flags & 0x100) != 0),
                $"Device usage {usage} must still deliver background input to the current observer's window.");
    }

    private static RawDevice[] ReadDevices()
    {
        uint count = 0;
        var size = (uint)Marshal.SizeOf<RawDevice>();
        Check.That(GetRegisteredRawInputDevices(null, ref count, size) != uint.MaxValue, "The registered device count must be readable.");
        if (count == 0) return [];
        var devices = new RawDevice[count];
        var read = GetRegisteredRawInputDevices(devices, ref count, size);
        Check.That(read != uint.MaxValue, "The current process's device registrations must be readable.");
        return devices.Take((int)read).ToArray();
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawDevice { public ushort UsagePage, Usage; public uint Flags; public nint Target; }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(RawDevice[] devices, uint count, uint size);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRegisteredRawInputDevices([Out] RawDevice[]? devices, ref uint count, uint size);
}
