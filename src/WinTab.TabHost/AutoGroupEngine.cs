using System.Text.RegularExpressions;
using WinTab.Core.Enums;
using WinTab.Core.Interfaces;
using WinTab.Core.Models;

namespace WinTab.TabHost;

/// <summary>
/// Automatically groups windows based on user-defined <see cref="AutoGroupRule"/> entries.
/// When a new window appears, the engine waits a short stabilization delay, then checks
/// the window against each enabled rule (sorted by priority). If a match is found, the
/// window is added to an existing group with the matching group name, or a new group
/// is created if none exists yet.
/// </summary>
public sealed class AutoGroupEngine : IDisposable
{
    private readonly IGroupManager _groupManager;
    private readonly IWindowEventSource _windowEventSource;
    private readonly IWindowManager _windowManager;
    private readonly Func<AppSettings> _settingsProvider;

    /// <summary>
    /// Delay in milliseconds to wait after a window appears before evaluating rules.
    /// This ensures the window is fully initialized (title, class name, etc.).
    /// </summary>
    private const int StabilizationDelayMs = 500;

    private bool _running;
    private bool _disposed;

    /// <summary>
    /// Tracks pending evaluations so we can cancel them when stopping.
    /// </summary>
    private CancellationTokenSource? _cts;

    /// <summary>
    /// Maps group names to their corresponding group IDs for existing auto-created groups.
    /// </summary>
    private readonly Dictionary<string, Guid> _namedGroups = new(StringComparer.OrdinalIgnoreCase);

    // ─── Constructor ────────────────────────────────────────────────────

    /// <summary>
    /// Creates a new <see cref="AutoGroupEngine"/>.
    /// </summary>
    /// <param name="groupManager">The group manager to create/update groups through.</param>
    /// <param name="windowEventSource">Source of system-wide window events.</param>
    /// <param name="windowManager">Window information/manipulation service.</param>
    /// <param name="settingsProvider">Provides the current application settings (including rules and exclusions).</param>
    public AutoGroupEngine(
        IGroupManager groupManager,
        IWindowEventSource windowEventSource,
        IWindowManager windowManager,
        Func<AppSettings> settingsProvider)
    {
        _groupManager = groupManager ?? throw new ArgumentNullException(nameof(groupManager));
        _windowEventSource = windowEventSource ?? throw new ArgumentNullException(nameof(windowEventSource));
        _windowManager = windowManager ?? throw new ArgumentNullException(nameof(windowManager));
        _settingsProvider = settingsProvider ?? throw new ArgumentNullException(nameof(settingsProvider));

        // Keep track of groups that are disbanded so we can remove stale named-group entries.
        _groupManager.GroupDisbanded += OnGroupDisbanded;
    }

    // ─── Public API ─────────────────────────────────────────────────────

    /// <summary>
    /// Starts monitoring for new windows and applying auto-group rules.
    /// </summary>
    public void Start()
    {
        if (_running || _disposed) return;

        _cts = new CancellationTokenSource();
        _windowEventSource.WindowShown += OnWindowShown;
        _running = true;
    }

    /// <summary>
    /// Stops monitoring. Any pending evaluations are cancelled.
    /// </summary>
    public void Stop()
    {
        if (!_running) return;

        _windowEventSource.WindowShown -= OnWindowShown;
        _cts?.Cancel();
        _cts?.Dispose();
        _cts = null;
        _running = false;
    }

    /// <summary>
    /// Whether the engine is currently running.
    /// </summary>
    public bool IsRunning => _running;

    // ─── Event Handlers ─────────────────────────────────────────────────

    private void OnWindowShown(object? sender, IntPtr hwnd)
    {
        if (!_running || hwnd == IntPtr.Zero) return;

        var settings = _settingsProvider();
        if (!settings.AutoApplyRules) return;

        // Fire-and-forget: evaluate after a stabilization delay.
        var token = _cts?.Token ?? CancellationToken.None;
        _ = EvaluateWindowAsync(hwnd, token);
    }

    private void OnGroupDisbanded(object? sender, TabGroup group)
    {
        // Remove any named-group mapping that pointed to this disbanded group.
        var keysToRemove = _namedGroups
            .Where(kvp => kvp.Value == group.Id)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var key in keysToRemove)
            _namedGroups.Remove(key);
    }

    // ─── Evaluation Logic ───────────────────────────────────────────────

    private async Task EvaluateWindowAsync(IntPtr hwnd, CancellationToken ct)
    {
        try
        {
            // Wait for the window to stabilize.
            await Task.Delay(StabilizationDelayMs, ct).ConfigureAwait(false);

            if (ct.IsCancellationRequested) return;

            // Ensure the window is still alive and visible.
            if (!_windowManager.IsAlive(hwnd) || !_windowManager.IsVisible(hwnd))
                return;

            // Check if the window is already in a group.
            if (_groupManager.GetGroupForWindow(hwnd) is not null)
                return;

            var windowInfo = _windowManager.GetWindowInfo(hwnd);
            if (windowInfo is null) return;

            var settings = _settingsProvider();

            // Check exclusion list.
            if (IsExcluded(windowInfo, settings.ExcludedProcesses))
                return;

            // Evaluate rules sorted by priority (lower number = higher priority).
            var rules = settings.AutoGroupRules
                .Where(r => r.Enabled)
                .OrderBy(r => r.Priority)
                .ToList();

            foreach (var rule in rules)
            {
                if (ct.IsCancellationRequested) return;

                if (!IsMatch(rule, windowInfo)) continue;

                // We have a match. Dispatch to UI thread for WPF group operations.
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    ApplyRule(hwnd, rule);
                });

                return; // Only apply the first matching rule.
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when stopping.
        }
        catch (Exception)
        {
            // Swallow exceptions from individual evaluations to prevent
            // one bad window from breaking the entire engine.
        }
    }

    private void ApplyRule(IntPtr hwnd, AutoGroupRule rule)
    {
        // Re-check: window might have been grouped while we were waiting.
        if (_groupManager.GetGroupForWindow(hwnd) is not null)
            return;

        string groupName = rule.GroupName;

        // Check if a group with this name already exists.
        if (_namedGroups.TryGetValue(groupName, out var existingGroupId))
        {
            // Verify the group still exists.
            var allGroups = _groupManager.GetAllGroups();
            var existingGroup = allGroups.FirstOrDefault(g => g.Id == existingGroupId);

            if (existingGroup is not null)
            {
                // Add to existing group.
                _groupManager.AddToGroup(existingGroupId, hwnd);
                return;
            }
            else
            {
                // Group was disbanded; remove stale entry.
                _namedGroups.Remove(groupName);
            }
        }

        // No existing group with this name. We need a second window to create a group.
        // Look for another visible, ungrouped window that matches the same rule.
        var candidates = _windowManager.EnumerateTopLevelWindows(includeInvisible: false);

        foreach (var candidate in candidates)
        {
            if (candidate.Handle == hwnd) continue;
            if (_groupManager.GetGroupForWindow(candidate.Handle) is not null) continue;
            if (!IsMatch(rule, candidate)) continue;

            // Found a matching partner. Create a new group.
            var newGroup = _groupManager.CreateGroup(candidate.Handle, hwnd);
            _namedGroups[groupName] = newGroup.Id;
            return;
        }

        // No matching partner found yet. The window will remain ungrouped for now.
        // It will be picked up if another matching window appears later.
    }

    // ─── Matching Logic ─────────────────────────────────────────────────

    /// <summary>
    /// Determines whether a window matches the given auto-group rule.
    /// </summary>
    /// <param name="rule">The rule to test.</param>
    /// <param name="window">The window information to match against.</param>
    /// <returns>True if the window matches the rule criteria.</returns>
    public static bool IsMatch(AutoGroupRule rule, WindowInfo window)
    {
        if (string.IsNullOrWhiteSpace(rule.MatchValue))
            return false;

        return rule.MatchType switch
        {
            AutoGroupMatchType.ProcessName =>
                string.Equals(window.ProcessName, rule.MatchValue, StringComparison.OrdinalIgnoreCase),

            AutoGroupMatchType.WindowTitleContains =>
                window.Title.Contains(rule.MatchValue, StringComparison.OrdinalIgnoreCase),

            AutoGroupMatchType.WindowTitleRegex =>
                TryRegexMatch(window.Title, rule.MatchValue),

            AutoGroupMatchType.ClassName =>
                string.Equals(window.ClassName, rule.MatchValue, StringComparison.OrdinalIgnoreCase),

            AutoGroupMatchType.ProcessPath =>
                !string.IsNullOrEmpty(window.ProcessPath) &&
                window.ProcessPath.Contains(rule.MatchValue, StringComparison.OrdinalIgnoreCase),

            _ => false
        };
    }

    /// <summary>
    /// Attempts a regex match, returning false on invalid patterns instead of throwing.
    /// </summary>
    private static bool TryRegexMatch(string input, string pattern)
    {
        try
        {
            return Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(100));
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// Checks whether a window should be excluded based on the process exclusion list.
    /// </summary>
    private static bool IsExcluded(WindowInfo window, IReadOnlyList<string> excludedProcesses)
    {
        if (excludedProcesses.Count == 0) return false;

        return excludedProcesses.Any(excluded =>
            string.Equals(window.ProcessName, excluded, StringComparison.OrdinalIgnoreCase));
    }

    // ─── IDisposable ────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        Stop();
        _groupManager.GroupDisbanded -= OnGroupDisbanded;
        _namedGroups.Clear();
    }
}
