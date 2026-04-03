# Explorer Fallback Recursion Fix

## Context

After WinTab installs and enables Explorer open-verb interception, opening folders can fail even when WinTab is not running. The failure path matters more than the steady-state path: once the delegate-execute bridge cannot forward to the app, Explorer must still open normally.

## Root Cause

Physical-folder fallback used:

- `explorer.exe`
- target path as argument
- `UseShellExecute = true`

That fallback re-entered the same shell open-verb resolution path that WinTab had already intercepted. When the pipe to the running app was unavailable, the delegate-execute handler could trigger another shell-mediated open, which could recurse back into WinTab instead of directly launching Explorer.

## Chosen Fix

Build a dedicated `ProcessStartInfo` for Explorer fallback and force:

- `FileName = "explorer.exe"`
- `Arguments = "\"<path>\""`
- `UseShellExecute = false`

This bypasses shell verb resolution and directly creates the Explorer process.

## Test Strategy

1. Add a failing regression test for the physical-folder fallback start info.
2. Verify the helper uses direct process creation instead of shell execute.
3. Run the focused `AppEnvironmentTests`.
4. Run the full `WinTab.Tests` suite.
5. Rebuild publish output and regenerate the installer.

## Expected Outcome

When the delegate-execute bridge cannot reach WinTab, physical folders fall back to native Explorer launch without re-entering the intercepted open verb, so Explorer remains usable whether WinTab is running or not.
