using System.Text.Json;
using WinTab.Core;

namespace WinTab.Persistence;

public sealed class SettingsStore
{
    private readonly string _settingsPath;
    private readonly string _attachmentsPath;
    private readonly JsonSerializerOptions _options = new() { WriteIndented = true };

    public SettingsStore(string settingsPath)
    {
        _settingsPath = settingsPath;
        var directory = Path.GetDirectoryName(settingsPath);
        _attachmentsPath = Path.Combine(directory ?? string.Empty, "window_attachments.json");
    }

    public AppSettings Load()
    {
        if (!File.Exists(_settingsPath))
        {
            return new AppSettings();
        }

        try
        {
            var json = File.ReadAllText(_settingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            settings.AutoGroupRules ??= new List<AutoGroupRule>();
            settings.GroupWindowStates ??= new List<GroupWindowState>();
            settings.WindowAttachments ??= new List<WindowAttachment>();

            foreach (var rule in settings.AutoGroupRules)
            {
                if (string.IsNullOrWhiteSpace(rule.GroupName))
                {
                    rule.GroupName = "Default";
                }

                if (string.IsNullOrWhiteSpace(rule.MatchValue) && !string.IsNullOrWhiteSpace(rule.ProcessName))
                {
                    rule.MatchType = AutoGroupMatchType.ProcessName;
                    rule.MatchValue = rule.ProcessName;
                }

                if (rule.MatchType == AutoGroupMatchType.ProcessName &&
                    string.IsNullOrWhiteSpace(rule.ProcessName) &&
                    !string.IsNullOrWhiteSpace(rule.MatchValue))
                {
                    rule.ProcessName = rule.MatchValue;
                }
            }

            return settings;
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        var directory = Path.GetDirectoryName(_settingsPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(settings, _options);
        File.WriteAllText(_settingsPath, json);
    }

    public void SaveGroupWindowStates(List<GroupWindowState> states)
    {
        var settings = Load();
        settings.GroupWindowStates = states ?? new List<GroupWindowState>();
        Save(settings);
    }

    public void SaveWindowAttachments(List<WindowAttachment> attachments)
    {
        var json = JsonSerializer.Serialize(attachments, _options);
        File.WriteAllText(_attachmentsPath, json);
    }

    public List<WindowAttachment> LoadWindowAttachments()
    {
        if (!File.Exists(_attachmentsPath))
        {
            return new List<WindowAttachment>();
        }

        try
        {
            var json = File.ReadAllText(_attachmentsPath);
            return JsonSerializer.Deserialize<List<WindowAttachment>>(json, _options)
                   ?? new List<WindowAttachment>();
        }
        catch
        {
            return new List<WindowAttachment>();
        }
    }

    public void BackupSession()
    {
        try
        {
            var settings = Load();
            var attachments = settings.WindowAttachments ?? new List<WindowAttachment>();
            SaveWindowAttachments(attachments);
        }
        catch
        {
        }
    }

    public void ClearBackup()
    {
        try
        {
            if (File.Exists(_attachmentsPath))
            {
                File.Delete(_attachmentsPath);
            }
        }
        catch
        {
        }
    }
}
