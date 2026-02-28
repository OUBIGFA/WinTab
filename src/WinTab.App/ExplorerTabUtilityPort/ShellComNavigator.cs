using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WinTab.App.ExplorerTabUtilityPort.Interop;
using WinTab.Diagnostics;
using WinTab.Platform.Win32;

namespace WinTab.App.ExplorerTabUtilityPort;

/// <summary>
/// Encapsulates COM interactions with IShellWindows and Shell.Application.
/// Provides methods to snapshot, enumerate, and navigate Explorer tabs via COM.
/// </summary>
public sealed class ShellComNavigator
{
    private static class NativeConstants
    {
        public const uint SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000;
    }

    private static readonly Guid ShellBrowserGuid = typeof(IShellBrowser).GUID;
    private static readonly Guid ShellWindowsClsid = new("9BA05972-F6A8-11CF-A442-00A0C90A8F39");
    private readonly Logger _logger;
    private static readonly object ShellWindowsInitLock = new();
    private static object? _shellWindows;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SHGetPathFromIDListW(IntPtr pidl, [Out] out string pszPath);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int SHGetNameFromIDList(IntPtr pidl, uint sigdnName, [Out] out string ppszName);


    public ShellComNavigator(Logger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Gets a thread-safe snapshot of currently open Explorer COM windows.
    /// Caller is responsible for releasing the COM objects returned in the list.
    /// </summary>
    public List<object> GetShellWindowsSnapshotUi()
    {
        var snapshot = new List<object>();
        object? windows = null;

        lock (ShellWindowsInitLock)
        {
            if (_shellWindows is null)
            {
                Type? shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType is not null)
                {
                    try
                    {
                        _shellWindows = Activator.CreateInstance(shellType);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("ShellComNavigator: Failed to initialize Shell.Application", ex);
                    }
                }
            }
            windows = _shellWindows;
        }

        if (windows is null)
            return snapshot;

        try
        {
            int count = (int)windows.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windows, null)!;
            for (int i = 0; i < count; i++)
            {
                try
                {
                    object? window = windows.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windows, new object[] { i });
                    if (window is not null)
                        snapshot.Add(window);
                }
                catch
                {
                    // ignore individual item failures
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Error("ShellComNavigator: Failed to enumerate ShellWindows", ex);
        }

        return snapshot;
    }

    public string? TryGetComLocation(object comTab)
    {
        try
        {
            dynamic win = comTab;
            string url = (string?)win.LocationURL ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(url) && Uri.TryCreate(url, UriKind.Absolute, out Uri? uri) && uri.IsFile)
                return Uri.UnescapeDataString(uri.LocalPath);
        }
        catch (Exception ex) when (IsComOrRpcException(ex))
        {
            // Ignore RPC disconnected
        }
        catch
        {
            // ignore
        }

        object? document = null;
        object? folder = null;
        object? self = null;

        try
        {
            dynamic win = comTab;
            document = win.Document;
            if (document is not null)
            {
                folder = ((dynamic)document).Folder;
                if (folder is not null)
                {
                    self = ((dynamic)folder).Self;
                    if (self is not null)
                    {
                        string path = (string?)((dynamic)self).Path ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(path))
                            return path;
                    }
                }
            }
        }
        catch (Exception ex) when (IsComOrRpcException(ex))
        {
            // Ignore RPC disconnected
        }
        catch
        {
            // ignore
        }
        finally
        {
            if (self is not null) Marshal.FinalReleaseComObject(self);
            if (folder is not null) Marshal.FinalReleaseComObject(folder);
            if (document is not null) Marshal.FinalReleaseComObject(document);
        }

        return null;
    }

    public bool TryNavigateComTab(object comTab, string location)
    {
        try
        {
            dynamic win = comTab;

            if (location.Contains('#'))
            {
                Type? shellType = Type.GetTypeFromProgID("Shell.Application");
                if (shellType is null)
                    return false;

                object? shell = null;
                try
                {
                    shell = Activator.CreateInstance(shellType);
                    if (shell is null) return false;

                    object? navigationFolder = null;
                    try
                    {
                        navigationFolder = ((dynamic)shell).NameSpace(location);
                        if (navigationFolder is null)
                            return false;
                        win.Navigate2(navigationFolder);
                        return true;
                    }
                    catch (Exception ex) when (IsComOrRpcException(ex))
                    {
                        return false;
                    }
                    finally
                    {
                        if (navigationFolder is not null)
                            Marshal.FinalReleaseComObject(navigationFolder);
                    }
                }
                finally
                {
                    if (shell is not null)
                        Marshal.FinalReleaseComObject(shell);
                }
            }

            win.Navigate2(location);
            return true;
        }
        catch (Exception ex) when (IsComOrRpcException(ex))
        {
            return false;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsComOrRpcException(Exception ex)
    {
        if (ex is COMException) return true;
        if (ex.InnerException is COMException) return true;
        if (ex.GetType().Name == "RuntimeBinderException") return true;
        return false;
    }
}
