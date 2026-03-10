using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace WinTab.Tests.App;

public sealed class ChineseCopyRegressionTests
{
    [Theory]
    [InlineData("General_Lead", "设置 WinTab 的启动方式、驻留方式，以及界面语言和主题。")]
    [InlineData("Behavior_Lead", "设置 WinTab 接管资源管理器窗口和标签页的方式。建议按使用场景逐项启用。")]
    [InlineData("About_Lead", "WinTab 是用于资源管理器标签化管理的开源工具，尽量在不改变原有操作习惯的前提下减少窗口切换。")]
    [InlineData("Uninstall_Lead", "可在此卸载 WinTab、决定是否删除本地数据，或仅恢复 WinTab 写入系统的配置。")]
    public void ChineseStrings_ShouldUseNeutralProfessionalCopy(string key, string expectedValue)
    {
        string stringsPath = GetProjectFilePath("WinTab.UI", "Localization", "Strings.zh-CN.xaml");
        XDocument strings = XDocument.Load(stringsPath);

        string? actualValue = strings
            .Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "String" && (string?)e.Attribute(XName.Get("Key", "http://schemas.microsoft.com/winfx/2006/xaml")) == key)
            ?.Value;

        actualValue.Should().Be(expectedValue);
    }

    private static string GetProjectFilePath(string projectFolder, params string[] parts)
    {
        string current = AppContext.BaseDirectory;

        for (int i = 0; i < 8; i++)
        {
            string candidate = Path.Combine(current, "src", projectFolder);
            if (Directory.Exists(candidate))
            {
                return Path.Combine(candidate, Path.Combine(parts));
            }

            string? parent = Directory.GetParent(current)?.FullName;
            if (string.IsNullOrWhiteSpace(parent))
            {
                break;
            }

            current = parent;
        }

        return Path.Combine(AppContext.BaseDirectory, Path.Combine(parts));
    }
}
