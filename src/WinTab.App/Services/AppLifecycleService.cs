using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;
using WinTab.Diagnostics;
using WinTab.Persistence;
using WinTab.Platform.Win32;

namespace WinTab.App.Services;

/// <summary>
/// Orchestrates the startup and shutdown of all background services
/// (window event watcher, drag detector, hotkey manager, etc.)
/// in the correct order.
/// </summary>
public sealed class AppLifecycleService
{
    private readonly Logger _logger;
    private readonly AppSettings _settings;
    private readonly SettingsStore _settingsStore;
    private readonly SessionStore _sessionStore;
    private readonly IWindowEventSource _windowEventSource;
    private readonly IHotKeyManager _hotKeyManager;
    private readonly IWindowManager _windowManager;
    private readonly IGroupManager? _groupManager;

    private bool _started;

    public AppLifecycleService(
        Logger logger,
        AppSettings settings,
        SettingsStore settingsStore,
        SessionStore sessionStore,
        IWindowEventSource windowEventSource,
        IHotKeyManager hotKeyManager,
        IWindowManager windowManager,
        IGroupManager? groupManager = null)
    {
        _logger = logger;
        _settings = settings;
        _settingsStore = settingsStore;
        _sessionStore = sessionStore;
        _windowEventSource = windowEventSource;
        _hotKeyManager = hotKeyManager;
        _windowManager = windowManager;
        _groupManager = groupManager;
    }

    /// <summary>
    /// Starts all background services based on current settings.
    /// </summary>
    public void Start(AppSettings settings)
    {
        if (_started) return;
        _started = true;

        _logger.Info("AppLifecycleService starting...");

        // Restore session if enabled
        if (settings.RestoreSessionOnStartup)
        {
            RestoreSession();
        }

        _logger.Info("AppLifecycleService started.");
    }

    /// <summary>
    /// Stops all background services and saves session state.
    /// </summary>
    public void Stop()
    {
        if (!_started) return;

        _logger.Info("AppLifecycleService stopping...");

        // Save current session
        SaveSession();

        // Dispose event source
        _windowEventSource.Dispose();

        // Save settings
        try
        {
            _settingsStore.Save(_settings);
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to save settings during shutdown.", ex);
        }

        _started = false;
        _logger.Info("AppLifecycleService stopped.");
    }

    private void RegisterHotKeys(AppSettings settings)
    {
        foreach (HotKeyBinding binding in settings.HotKeys)
        {
            if (!binding.Enabled || binding.Key == 0)
                continue;

            int id = (int)binding.Action;
            bool registered = _hotKeyManager.Register(id, binding.Modifiers, binding.Key);

            if (registered)
                _logger.Info($"Hotkey registered: {binding.Action} -> {binding.DisplayString}");
            else
                _logger.Warn($"Failed to register hotkey: {binding.Action} -> {binding.DisplayString}");
        }
    }

    private void OnHotKeyPressed(object? sender, HotKeyAction action)
    {
        try
        {
            IntPtr foreground = NativeMethods.GetForegroundWindow();
            if (foreground == IntPtr.Zero)
                return;

            if (action == HotKeyAction.NewInstance)
            {
                LaunchNewInstanceForWindow(foreground);
                return;
            }

            if (_groupManager is null)
                return;

            TabGroup? group = _groupManager.GetGroupForWindow(foreground);
            if (group is null)
                return;

            switch (action)
            {
                case HotKeyAction.NextTab:
                    SwitchTabRelative(group, +1);
                    break;

                case HotKeyAction.PreviousTab:
                    SwitchTabRelative(group, -1);
                    break;

                case HotKeyAction.CloseTab:
                    _groupManager.CloseTab(group.Id, foreground);
                    break;

                case HotKeyAction.DetachTab:
                    _groupManager.RemoveFromGroup(foreground);
                    break;

                case HotKeyAction.MoveTabLeft:
                    _groupManager.MoveTab(group.Id, foreground, -1);
                    break;

                case HotKeyAction.MoveTabRight:
                    _groupManager.MoveTab(group.Id, foreground, 1);
                    break;

                case HotKeyAction.ToggleGrouping:
                    _groupManager.DisbandGroup(group.Id);
                    break;

                default:
                    _logger.Info($"Hotkey action is not implemented yet: {action}");
                    break;
            }
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to process hotkey action: {action}", ex);
        }
    }

    private void SwitchTabRelative(TabGroup group, int delta)
    {
        if (group.Tabs.Count <= 1)
            return;

        int currentIndex = group.Tabs.FindIndex(t => t.Handle == group.ActiveHandle);
        if (currentIndex < 0)
            currentIndex = 0;

        int nextIndex = (currentIndex + delta + group.Tabs.Count) % group.Tabs.Count;
        _groupManager!.SwitchTab(group.Id, nextIndex);
    }

    private void LaunchNewInstanceForWindow(IntPtr hwnd)
    {
        var info = _windowManager.GetWindowInfo(hwnd);
        string? processPath = info?.ProcessPath;

        if (string.IsNullOrWhiteSpace(processPath) || !File.Exists(processPath))
        {
            _logger.Warn("Unable to launch a new instance: process path was not available.");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = processPath,
            UseShellExecute = true
        });
    }

    private void RestoreSession()
    {
        if (_groupManager is null)
        {
            _logger.Info("No group manager available; skipping session restore.");
            return;
        }

        if (!_sessionStore.HasSession())
        {
            _logger.Info("No previous session to restore.");
            return;
        }

        try
        {
            List<GroupWindowState> savedGroups = _sessionStore.LoadSession();
            _settings.SavedGroupStates = savedGroups;

            if (savedGroups.Count == 0)
            {
                _logger.Info("Session file was empty.");
                return;
            }

            List<WindowInfo> availableWindows = _windowManager
                .EnumerateTopLevelWindows(includeInvisible: true)
                .Where(CanUseForRestore)
                .ToList();

            int restoredCount = 0;
            foreach (GroupWindowState savedGroup in savedGroups)
            {
                if (TryRestoreGroup(savedGroup, availableWindows))
                    restoredCount++;
            }

            _logger.Info($"Session restore completed: {restoredCount}/{savedGroups.Count} group(s) restored.");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to restore session.", ex);
        }
    }

    private void SaveSession()
    {
        if (_groupManager is null)
        {
            _logger.Info("No group manager available; skipping session save.");
            return;
        }

        try
        {
            var groups = _groupManager.GetAllGroups();
            var states = groups
                .Select(BuildGroupState)
                .Where(state => state.Tabs.Count >= 2)
                .ToList();

            _sessionStore.SaveSession(states);
            _logger.Info($"Session saved with {states.Count} group(s).");
        }
        catch (Exception ex)
        {
            _logger.Error("Failed to save session.", ex);
        }
    }

    private GroupWindowState BuildGroupState(TabGroup group)
    {
        int activeTabIndex = group.Tabs.FindIndex(t => t.Handle == group.ActiveHandle);

        return new GroupWindowState
        {
            GroupName = group.Name,
            Left = group.Left,
            Top = group.Top,
            Width = group.Width,
            Height = group.Height,
            ActiveTabIndex = activeTabIndex < 0 ? 0 : activeTabIndex,
            State = GetWindowStateMode(group.ActiveHandle),
            Tabs = group.Tabs
                .Select((tab, index) => BuildTabState(tab, index))
                .ToList()
        };
    }

    private GroupWindowTabState BuildTabState(TabItem tab, int order)
    {
        WindowInfo? liveInfo = _windowManager.GetWindowInfo(tab.Handle);

        return new GroupWindowTabState
        {
            Order = order,
            ProcessName = liveInfo?.ProcessName ?? tab.ProcessName,
            WindowTitle = liveInfo?.Title ?? tab.Title,
            ClassName = liveInfo?.ClassName ?? string.Empty,
            ProcessPath = liveInfo?.ProcessPath
        };
    }

    private static GroupWindowStateMode GetWindowStateMode(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return GroupWindowStateMode.Normal;

        var placement = new NativeStructs.WINDOWPLACEMENT
        {
            length = (uint)Marshal.SizeOf<NativeStructs.WINDOWPLACEMENT>()
        };

        if (!NativeMethods.GetWindowPlacement(hwnd, ref placement))
            return GroupWindowStateMode.Normal;

        return placement.showCmd switch
        {
            (uint)NativeConstants.SW_SHOWMINIMIZED => GroupWindowStateMode.Minimized,
            (uint)NativeConstants.SW_SHOWMAXIMIZED => GroupWindowStateMode.Maximized,
            _ => GroupWindowStateMode.Normal
        };
    }

    private bool CanUseForRestore(WindowInfo window)
    {
        if (window.Handle == IntPtr.Zero)
            return false;

        string currentProcessName = Process.GetCurrentProcess().ProcessName;
        if (string.Equals(
                NormalizeProcessName(window.ProcessName),
                NormalizeProcessName(currentProcessName),
                StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string? currentPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(currentPath) &&
            !string.IsNullOrWhiteSpace(window.ProcessPath) &&
            string.Equals(window.ProcessPath, currentPath, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private bool TryRestoreGroup(GroupWindowState savedGroup, List<WindowInfo> availableWindows)
    {
        List<GroupWindowTabState> tabs = savedGroup.Tabs
            .OrderBy(tab => tab.Order)
            .ToList();

        if (tabs.Count < 2)
        {
            _logger.Warn($"Skipping session group '{savedGroup.GroupName}': not enough tab descriptors.");
            return false;
        }

        var usedHandles = new HashSet<IntPtr>();
        var matchedHandles = new List<IntPtr>(tabs.Count);

        foreach (GroupWindowTabState tab in tabs)
        {
            WindowInfo? match = FindBestWindowMatch(tab, availableWindows, usedHandles);
            if (match is null)
            {
                _logger.Warn($"Skipping session group '{savedGroup.GroupName}': could not match all windows.");
                return false;
            }

            usedHandles.Add(match.Handle);
            matchedHandles.Add(match.Handle);
        }

        try
        {
            TabGroup group = _groupManager!.CreateGroup(matchedHandles[0], matchedHandles[1]);
            for (int i = 2; i < matchedHandles.Count; i++)
                _groupManager.AddToGroup(group.Id, matchedHandles[i]);

            if (!string.IsNullOrWhiteSpace(savedGroup.GroupName))
                _groupManager.RenameGroup(group.Id, savedGroup.GroupName);

            ApplySavedLayout(group.Id, matchedHandles, savedGroup);
            availableWindows.RemoveAll(window => usedHandles.Contains(window.Handle));

            return true;
        }
        catch (Exception ex)
        {
            _logger.Error($"Failed to restore session group '{savedGroup.GroupName}'.", ex);
            return false;
        }
    }

    private static WindowInfo? FindBestWindowMatch(
        GroupWindowTabState savedTab,
        IReadOnlyList<WindowInfo> candidates,
        HashSet<IntPtr> usedHandles)
    {
        WindowInfo? bestMatch = null;
        int bestScore = int.MinValue;

        foreach (WindowInfo candidate in candidates)
        {
            if (usedHandles.Contains(candidate.Handle))
                continue;

            int score = ScoreWindowMatch(savedTab, candidate);
            if (score > bestScore)
            {
                bestScore = score;
                bestMatch = candidate;
            }
        }

        return bestScore >= 30 ? bestMatch : null;
    }

    private static int ScoreWindowMatch(GroupWindowTabState savedTab, WindowInfo candidate)
    {
        int score = 0;

        string expectedProcess = NormalizeProcessName(savedTab.ProcessName);
        string candidateProcess = NormalizeProcessName(candidate.ProcessName);
        if (!string.IsNullOrWhiteSpace(expectedProcess))
        {
            if (!string.Equals(expectedProcess, candidateProcess, StringComparison.OrdinalIgnoreCase))
                return int.MinValue;

            score += 100;
        }

        if (!string.IsNullOrWhiteSpace(savedTab.ProcessPath) && !string.IsNullOrWhiteSpace(candidate.ProcessPath))
        {
            if (!string.Equals(savedTab.ProcessPath, candidate.ProcessPath, StringComparison.OrdinalIgnoreCase))
                return int.MinValue;

            score += 60;
        }

        if (!string.IsNullOrWhiteSpace(savedTab.ClassName))
        {
            if (!string.Equals(savedTab.ClassName, candidate.ClassName, StringComparison.OrdinalIgnoreCase))
                return int.MinValue;

            score += 20;
        }

        string expectedTitle = savedTab.WindowTitle.Trim();
        string candidateTitle = candidate.Title.Trim();
        if (!string.IsNullOrWhiteSpace(expectedTitle))
        {
            if (string.Equals(expectedTitle, candidateTitle, StringComparison.OrdinalIgnoreCase))
            {
                score += 40;
            }
            else if (candidateTitle.Contains(expectedTitle, StringComparison.OrdinalIgnoreCase) ||
                     expectedTitle.Contains(candidateTitle, StringComparison.OrdinalIgnoreCase))
            {
                score += 20;
            }
        }

        return score;
    }

    private static string NormalizeProcessName(string? processName)
    {
        if (string.IsNullOrWhiteSpace(processName))
            return string.Empty;

        string normalized = processName.Trim();
        if (normalized.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            normalized = normalized[..^4];

        return normalized;
    }

    private void ApplySavedLayout(Guid groupId, IReadOnlyList<IntPtr> handles, GroupWindowState savedGroup)
    {
        if (handles.Count == 0)
            return;

        int width = (int)Math.Round(savedGroup.Width);
        int height = (int)Math.Round(savedGroup.Height);
        if (width > 0 && height > 0)
        {
            _windowManager.SetBounds(
                handles[0],
                (int)Math.Round(savedGroup.Left),
                (int)Math.Round(savedGroup.Top),
                width,
                height);
        }

        int activeIndex = savedGroup.ActiveTabIndex;
        if (activeIndex < 0 || activeIndex >= handles.Count)
            activeIndex = 0;

        if (activeIndex != 0)
            _groupManager!.SwitchTab(groupId, activeIndex);

        ApplyWindowState(handles[activeIndex], savedGroup.State);
    }

    private void ApplyWindowState(IntPtr hwnd, GroupWindowStateMode state)
    {
        switch (state)
        {
            case GroupWindowStateMode.Minimized:
                _windowManager.Minimize(hwnd);
                break;

            case GroupWindowStateMode.Maximized:
                NativeMethods.ShowWindow(hwnd, NativeConstants.SW_SHOWMAXIMIZED);
                break;

            default:
                _windowManager.Restore(hwnd);
                break;
        }
    }
}
