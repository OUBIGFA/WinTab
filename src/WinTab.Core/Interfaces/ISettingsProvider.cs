using WinTab.Core.Models;

namespace WinTab.Core.Interfaces;

public interface ISettingsProvider
{
    AppSettings Settings { get; }
    void Save();
    void Reload();
    event EventHandler? SettingsChanged;
}
