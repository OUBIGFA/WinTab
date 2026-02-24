# Task Plan

- [x] Refine close-on-double-click hit testing so it only triggers on tab title area
- [x] Verify behavior by running build and tests
- [x] Publish app binaries for installer packaging
- [x] Build latest installer executable
- [x] Record artifact path(s) and command results

# Review

- [x] Refine close-on-double-click hit testing so it only triggers on tab title area
- [x] Verify behavior by running build and tests
- [x] Publish app binaries for installer packaging
- [x] Build latest installer executable
- [x] Record artifact path(s) and command results

- Hotfix: hit-testing now resolves Explorer top-level window from cursor point first, then closes the active `ShellTabWindowClass` tab for that window
- Added tab-title-area filtering in mouse hook: first prefer accessibility role (`ROLE_SYSTEM_PAGETAB` / `ROLE_SYSTEM_PAGETABLIST`), then fallback to a top header band heuristic
- Updated behavior toggle copy to clarify it applies to tab title area
- `dotnet build WinTab.slnx -c Release` succeeded
- `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2)
- `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded
- `"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installers/WinTab.iss` succeeded
- Installer artifact: `publish/installer/WinTab_Setup_1.0.0.exe`
- SHA256: `17baf7b7b37744f84a51942be1863a0941f5565cbde486fb893e76ead3f5ffdd`
