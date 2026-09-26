using System;
using System.IO;
using Microsoft.VisualBasic.FileIO;

/// <summary>
/// Removes directories owned by the test run.
///
/// Cleanup prefers the recycle bin so a mistaken path stays recoverable, but falls back to a
/// permanent delete when the shell cannot recycle it. <c>SHFileOperation</c> fails in restricted or
/// non-interactive sessions (services, CI runners) and when the recycle bin is unavailable, and a
/// cleanup that throws would fail an otherwise passing test — as the finally blocks used to.
/// </summary>
internal static class TestCleanup
{
    /// <summary>Deletes a test-owned directory, falling back to a permanent delete.</summary>
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
            // The shell can report a failure after it has already removed the directory — this build
            // environment does exactly that — so the fallback must tolerate a missing path. Remove
            // whatever is left and never let a leftover test directory fail an otherwise passing run.
            try
            {
                Directory.Delete(directory, true);
            }
            catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException)
            {
                // Leftover test directories are harmless; a failed test result would not be.
            }
        }
    }
}
