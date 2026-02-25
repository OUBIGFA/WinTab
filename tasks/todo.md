# Task Plan

- [x] Diagnose startup theme initialization regression after reinstall
- [x] Implement startup theme fix so first render matches saved/default theme
- [x] Verify fix by building and running tests
- [x] Repackage installer from latest code
- [x] Update checksum and record artifact details

# Review

- Root cause: theme was applied before `MainWindow` existed, and `MainWindow` had a local `Background="Transparent"` override, causing inconsistent first paint until manual theme toggle
- Fix: removed `Background="Transparent"` from `MainWindow.xaml` and re-applied theme once after window creation in startup path
- `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors)
- `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2)
- `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded
- `"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installers/WinTab.iss` succeeded
- Installer: `publish/installer/WinTab_Setup_1.0.0.exe`
- SHA256: `130904400ee6d4dca4a834be94b0c855f2497351039cac10db1e3fe98ce1d12e`

## 2026-02-24 Duplicate Process Fix Plan

- [x] Confirm why two `WinTab.exe` processes are present with different PIDs
- [x] Remove the root cause while preserving Explorer open-folder fallback behavior
- [x] Build and run tests to verify no regression
- [x] Record review notes and verification evidence

## 2026-02-24 Duplicate Process Fix Review

- Root cause: `App.OnStartup` always called `StartCompanionWatcher()` when Explorer open-verb interception was enabled, which launched a second long-lived `WinTab.exe --wintab-companion <pid>` process.
- Fix: removed companion watcher startup and deleted the `ExplorerOpenVerbCompanion` implementation; kept `--wintab-companion` as a no-op fast exit for legacy compatibility.
- Safety: Explorer open-folder fallback/repair path remains in `TryHandleOpenFolderInvocation` and startup self-check still repairs inconsistent registry state.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2).

## 2026-02-25 Tab Double-Click Misfire Fix Plan

- [x] Analyze double-click close hook path and identify why toolbar rapid clicks can close tabs
- [x] Tighten tab target resolution and hit-test gating to tab-title-only interactions
- [x] Build and run tests to verify no regression
- [x] Record implementation review and verification evidence

## 2026-02-25 Tab Double-Click Misfire Fix Review

- Root cause: when MSAA hit-testing returned non-tab/unknown roles, code fell back to a broad top-header Y-range, so rapid clicks on Explorer toolbar controls (back/forward/refresh) could be treated as tab-title double-clicks.
- Fix: removed geometric header fallback and now only allow close-on-double-click when accessibility classification resolves to `Tab`; other roles are rejected.
- Hardening: tab handle resolution now follows the clicked window's parent chain to `ShellTabWindowClass`, instead of always selecting the first tab under the Explorer top-level window.
- Files changed: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabMouseHookService.cs`.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2).

## 2026-02-25 Uninstall UX/Robustness Plan

- [x] Analyze uninstall flow and identify missing running-instance handling plus data-deletion UX gaps
- [x] Replace uninstall-time data deletion prompt with uninstall-page checkbox (default unchecked)
- [x] Wire checkbox state into uninstaller command line and script-side parsing
- [x] Harden installer/uninstaller running-process handling via setup options
- [x] Build and run tests to verify no regression
- [x] Record implementation review and verification evidence

## 2026-02-25 Uninstall UX/Robustness Review

- UX: removed uninstall-time delete-data popup and moved the choice into app uninstall page as a checkbox (`Remove user data`), default unchecked.
- Behavior: checkbox state is forwarded to uninstaller as `/REMOVEUSERDATA=1`; uninstaller script now reads command tail and only removes `%AppData%\Roaming\WinTab` when flag is present.
- Compatibility: system-started uninstall (without app UI) keeps previous safe default (do not remove user data).
- Robustness: installer script now declares `AppMutex=WinTab_SingleInstance` and enables `CloseApplications` with `CloseApplicationsFilter=WinTab.exe` to improve running-process handling during uninstall.
- Docs/localization updated for new behavior in uninstall page text and README.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2).

## 2026-02-25 Reinstall Choice Flow Plan

- [x] Analyze current installer/uninstaller flow and identify insertion points for reinstall choice UI
- [x] Implement installer page for "uninstall then install" vs "install directly" with remove-user-data checkbox default unchecked
- [x] Wire selected mode into pre-install execution path and silent-mode command line parameters
- [x] Update README with reinstall behavior and silent parameter documentation
- [x] Compile installer script to verify syntax and packaging path

## 2026-02-25 Reinstall Choice Flow Review

- Added custom reinstall page (shown only when existing install is detected) with default mode `uninstall then install`, and optional `remove user data` checkbox default unchecked.
- Added pre-install execution logic in `PrepareToInstall`: launches existing uninstaller for clean mode, appends `/REMOVEUSERDATA=1` only when selected, and aborts on uninstall failure/non-zero exit.
- Added silent-mode controls: `/REINSTALLMODE=CLEAN|DIRECT` and `/REMOVEUSERDATA=1`; default for silent mode is `DIRECT`.
- Added bilingual installer messages in `installers/WinTab.iss` for reinstall page and failure paths.
- Updated README to document reinstall options and silent parameters.
- Verification: `"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installers/WinTab.iss` succeeded; output `publish/installer/WinTab_Setup_1.0.0.exe`.
