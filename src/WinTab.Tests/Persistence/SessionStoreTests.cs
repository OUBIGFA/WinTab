using System.IO;
using WinTab.Core.Models;
using WinTab.Persistence;
using Xunit;

namespace WinTab.Tests.Persistence;

public sealed class SessionStoreTests
{
    [Fact]
    public void SaveAndLoadSession_PreservesTabDescriptors()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"WinTab.Tests.{Guid.NewGuid():N}");
        string sessionPath = Path.Combine(baseDir, "session.json");

        try
        {
            var store = new SessionStore(sessionPath);
            var source = new List<GroupWindowState>
            {
                new()
                {
                    GroupName = "Browsers",
                    Left = 10,
                    Top = 20,
                    Width = 1400,
                    Height = 900,
                    ActiveTabIndex = 1,
                    Tabs =
                    [
                        new GroupWindowTabState
                        {
                            Order = 0,
                            ProcessName = "chrome",
                            WindowTitle = "Docs - Chrome",
                            ClassName = "Chrome_WidgetWin_1",
                            ProcessPath = @"C:\Program Files\Google\Chrome\Application\chrome.exe"
                        },
                        new GroupWindowTabState
                        {
                            Order = 1,
                            ProcessName = "msedge",
                            WindowTitle = "Search - Microsoft Edge",
                            ClassName = "Chrome_WidgetWin_1",
                            ProcessPath = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                        }
                    ]
                }
            };

            store.SaveSession(source);
            List<GroupWindowState> restored = store.LoadSession();

            Assert.Single(restored);
            Assert.Equal("Browsers", restored[0].GroupName);
            Assert.Equal(2, restored[0].Tabs.Count);
            Assert.Equal("chrome", restored[0].Tabs[0].ProcessName);
            Assert.Equal("msedge", restored[0].Tabs[1].ProcessName);
            Assert.Equal(1, restored[0].ActiveTabIndex);
        }
        finally
        {
            if (Directory.Exists(baseDir))
                Directory.Delete(baseDir, recursive: true);
        }
    }

    [Fact]
    public void LoadSession_OldSchemaWithoutTabs_UsesEmptyTabList()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), $"WinTab.Tests.{Guid.NewGuid():N}");
        string sessionPath = Path.Combine(baseDir, "session.json");

        try
        {
            Directory.CreateDirectory(baseDir);

            string legacyJson = """
[
  {
    "GroupName": "Legacy",
    "Left": 0,
    "Top": 0,
    "Width": 1000,
    "Height": 700,
    "State": "Normal",
    "ActiveTabIndex": 0
  }
]
""";
            File.WriteAllText(sessionPath, legacyJson);

            var store = new SessionStore(sessionPath);
            List<GroupWindowState> restored = store.LoadSession();

            Assert.Single(restored);
            Assert.NotNull(restored[0].Tabs);
            Assert.Empty(restored[0].Tabs);
        }
        finally
        {
            if (Directory.Exists(baseDir))
                Directory.Delete(baseDir, recursive: true);
        }
    }
}
