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
