using System;
using System.IO;
using Microsoft.VisualBasic.FileIO;

/// <summary>
/// Removes directories owned by the test run.
///
/// Cleanup only uses the recycle bin. If the shell cannot recycle a directory, retain it and report
/// its path without hiding the original test result or silently switching to permanent deletion.
/// </summary>
internal static class TestCleanup
{
    /// <summary>Recycles a test-owned directory, reporting any files left behind.</summary>
    public static void DeleteDirectory(string directory)
    {
        if (!Directory.Exists(directory))
            return;

        try
        {
            FileSystem.DeleteDirectory(directory, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                      or System.Runtime.InteropServices.COMException)
        {
            // SHFileOperation may report an error after recycling everything successfully.
            if (Directory.Exists(directory))
                Console.Error.WriteLine($"WARN Test files could not be recycled and were retained at {directory}: {ex.Message}");
        }
    }
}
