using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Win32;
using WinTab.Managers;

internal static class RecycleBinOpenRegistrationTests
{
    private const string Page = @"Software\Classes\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}";
    private const string Shell = Page + @"\shell";
    private const string Open = Shell + @"\open";
    private const string Command = Open + @"\command";
    private const string Executable = @"C:\Program Files\WinTab\WinTab.exe";

    public static IEnumerable<(string Name, Func<Task> Body)> All()
    {
        yield return ("recycle registration routes opening to the running app and restores a missing override", () => WithRoot(root =>
        {
            Check.That(RecycleBinOpenRegistration.Update(root, true, Executable), "Registration must succeed.");
            using (var command = root.OpenSubKey(Command))
                Check.Equal("\"C:\\Program Files\\WinTab\\WinTab.exe\" --open-recycle-bin", command?.GetValue("") as string,
                    "The command must send the request directly, not launch an intermediate Explorer window.");
            using (var shell = root.OpenSubKey(Shell))
                Check.Equal("open", shell?.GetValue("") as string, "The custom verb must be the default open action.");
            Check.That(RecycleBinOpenRegistration.Update(root, false, Executable), "Removal must succeed.");
            using var page = root.OpenSubKey(Page);
            Check.That(page == null, "A previously absent page override must be removed completely.");
        }));
        yield return ("recycle registration restores the original default after repeated enable and process restart", () => WithRoot(root =>
        {
            using (var shell = root.CreateSubKey(Shell)) shell.SetValue("", "%OriginalVerb%", RegistryValueKind.ExpandString);
            using (var icon = root.CreateSubKey(Page + @"\DefaultIcon")) icon.SetValue("", "my-icon");
            Check.That(RecycleBinOpenRegistration.Update(root, true, Executable), "Registration must succeed.");
            Check.That(RecycleBinOpenRegistration.Update(root, true, Executable), "Repeated enabling must not replace the original backup.");
            Check.That(RecycleBinOpenRegistration.Update(root, false, Executable), "A fresh caller must restore persistent ownership metadata.");
            using var restored = root.OpenSubKey(Shell);
            Check.Equal("%OriginalVerb%", restored?.GetValue("", null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string,
                "The original string must be restored without expansion.");
            Check.Equal(RegistryValueKind.ExpandString, restored!.GetValueKind(""));
            using var preservedIcon = root.OpenSubKey(Page + @"\DefaultIcon");
            Check.Equal("my-icon", preservedIcon?.GetValue("") as string, "Unrelated customization must survive.");
        }));
        yield return ("recycle registration leaves another application's open command untouched", () => WithRoot(root =>
        {
            using (var command = root.CreateSubKey(Command)) command.SetValue("", "other-app");
            Check.That(!RecycleBinOpenRegistration.Update(root, true, Executable), "An existing custom handler must be reported as a conflict.");
            Check.That(RecycleBinOpenRegistration.Update(root, false, Executable), "Disabling must not claim a foreign handler.");
            using var preserved = root.OpenSubKey(Command);
            Check.Equal("other-app", preserved?.GetValue("") as string);
        }));
        yield return ("recycle registration preserves a command changed while WinTab was running", () => WithRoot(root =>
        {
            Check.That(RecycleBinOpenRegistration.Update(root, true, Executable), "Registration must succeed.");
            using (var command = root.OpenSubKey(Command, true)) command!.SetValue("", "new-custom-command");
            Check.That(!RecycleBinOpenRegistration.Update(root, false, Executable), "A changed handler must not be removed as ours.");
            using var preserved = root.OpenSubKey(Command);
            Check.Equal("new-custom-command", preserved?.GetValue("") as string);
        }));
        yield return ("recycle registration preserves a newer default verb when removing its own handler", () => WithRoot(root =>
        {
            Check.That(RecycleBinOpenRegistration.Update(root, true, Executable), "Registration must succeed.");
            using (var shell = root.OpenSubKey(Shell, true)) shell!.SetValue("", "new-verb");
            Check.That(RecycleBinOpenRegistration.Update(root, false, Executable), "Only our command should be removed.");
            using var preserved = root.OpenSubKey(Shell);
            Check.Equal("new-verb", preserved?.GetValue("") as string);
            using var removed = root.OpenSubKey(Open);
            Check.That(removed == null, "Our empty verb must be removed.");
        }));
        yield return ("disabled recycle registration creates no registry keys", () => WithRoot(root =>
        {
            Check.That(RecycleBinOpenRegistration.Update(root, false, Executable), "An absent registration is already disabled.");
            Check.Equal(0, root.SubKeyCount);
        }));
    }

    private static Task WithRoot(Action<RegistryKey> test)
    {
        // Every registry path in these tests is nested under a unique test-owned key.
        var path = @"Software\WinTab.Tests\" + Guid.NewGuid().ToString("N");
        try { using var root = Registry.CurrentUser.CreateSubKey(path); test(root); }
        finally { Registry.CurrentUser.DeleteSubKeyTree(path, false); }
        return Task.CompletedTask;
    }
}
