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

## 2026-02-25 Double-Click Close Regression Plan

- [x] Reproduce and pinpoint why double-clicking Explorer tab title does not close tabs
- [x] Implement targeted fix in Explorer mouse hook without broad behavior changes
- [x] Build and run tests to verify no regression
- [x] Package a latest installer for user validation
- [x] Record review notes with root cause and verification evidence

## 2026-02-25 Double-Click Close Regression Review

- Root cause: double-click recognition depended on tab handle stability across both clicks, but Explorer can switch active tab after the first click (especially when clicking an inactive tab), causing second-click mismatch and no close.
- Root cause: title-area hit testing relied only on one geometric window relation in some builds; when that relation is unstable, valid title clicks can be filtered out.
- Fix: relaxed double-click matching to use top-level Explorer window + system double-click time/position thresholds, and improved title-area classification with layered checks (MSAA tab positive, navigation-control negative, tab-content negative, geometry fallback with metrics-based header fallback).
- Hook behavior: when close triggers, the low-level hook now consumes the click so Explorer native double-click title actions do not race with close.
- File changed: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabMouseHookService.cs`.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (3/3).
- Packaging: `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded.
- Packaging: Inno Setup compile succeeded, installer output `publish/installer/WinTab_Setup_1.0.0.exe`.
- SHA256: `7ad2251901c354fea11e0cda50fa45226412701fc6fa45e935ac38b709e2cd9d`.

## 2026-02-25 Double-Click Close Regression Plan (Round 2)

- [x] Re-check failure path against user logs and identify over-filtering conditions
- [x] Remove brittle hit-test rejection and adjust close command target strategy
- [x] Build and run tests to verify no regression
- [x] Package latest installer for user verification
- [x] Record review notes and updated checksum

## 2026-02-25 Double-Click Close Regression Review (Round 2)

- User signal: feature was enabled in logs, but no close behavior observed in Explorer.
- Root cause candidate 1: `IsPointWithinTabWindow(...)` hard rejection could classify valid tab-title clicks as invalid on some Explorer builds where `WindowFromPoint` maps title regions under tab window ancestry.
- Root cause candidate 2: posting close command to top-level window can be less reliable than posting to resolved tab handle in this codebase's existing path.
- Fix: removed `IsPointWithinTabWindow(...)` rejection from title-area predicate and switched close command target preference to resolved `tabHandle` (fallback to top-level only when tab handle is unavailable).
- Observability: added info log `Explorer tab double-click detected; sending close command.` to confirm recognition path in user logs.
- File changed: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabMouseHookService.cs`.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (3/3).
- Packaging: `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded.
- Packaging: Inno Setup compile succeeded, installer output `publish/installer/WinTab_Setup_1.0.0.exe`.
- SHA256: `25d7e73527d79324fc70e832b697302e7b759748b08f305c549c710b3a20fa22`.

## 2026-02-25 Double-Click Scope Tightening Plan

- [x] Analyze current hit-test path and confirm why address/toolbar area can still trigger close
- [x] Tighten predicate to tab-title-only behavior with strict fallback gating
- [x] Build and run tests to verify no regression
- [x] Package latest installer for user verification
- [x] Record review notes and checksum

## 2026-02-25 Double-Click Scope Tightening Review

- Requirement update: only tab title area should trigger close; address bar and navigation/refresh toolbar region must not trigger.
- Fix: in `IsPointInTabTitleArea`, `Unknown/Other` fallback now requires point ancestry to be under `ShellTabWindowClass` and uses a capped top strip (`maxTabStripHeight`) instead of broad `ExplorerTop->TabTop` range.
- Existing hard block for accessibility `NavigationControl` roles remains, reducing accidental toolbar/button closes.
- File changed: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabMouseHookService.cs`.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (3/3).
- Packaging: `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded.
- Packaging: Inno Setup compile succeeded, installer output `publish/installer/WinTab_Setup_1.0.0.exe`.
- SHA256: `bfd32ad0407b21e89abf8700e67745310dc6428a7f0ad635df98740506535d56`.

## 2026-02-25 Double-Click Scope Tightening Plan (Round 2)

- [x] Re-check why strict title-only build disabled closing entirely on user machine
- [x] Remove over-strict ancestry gate and keep narrow top-strip + navigation-role exclusion
- [x] Build and run tests to verify no regression
- [x] Package latest installer for user verification
- [x] Record review notes and checksum

## 2026-02-25 Double-Click Scope Tightening Review (Round 2)

- User feedback indicated strict ancestry gate made close-on-double-click non-functional.
- Root cause: requiring point ancestry under `ShellTabWindowClass` can reject valid tab-title clicks on some Explorer builds.
- Fix: removed ancestry hard gate and kept two constraints only: accessibility navigation-role rejection + narrow top-strip cap geometry for fallback.
- Tuning: top-strip cap uses `caption + frame + 8` clamped to `[30,56]`, reducing address/toolbar inclusion while keeping tab-title clicks viable.
- File changed: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabMouseHookService.cs`.
- Verification: `dotnet build WinTab.slnx -c Release` succeeded (0 warnings, 0 errors).
- Verification: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (3/3).
- Packaging: `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded.
- Packaging: Inno Setup compile succeeded, installer output `publish/installer/WinTab_Setup_1.0.0.exe`.
- SHA256: `81bd15eb0b79a0b8c534d21907dea629df6984d0128c0ccae54048dcefe326bc`.
