# Task Plan

- [x] Verify packaging prerequisites (dotnet publish + Inno Setup compiler)
- [x] Publish WinTab app binaries for installer input directory
- [x] Compile Inno Setup installer executable
- [x] Generate artifact checksum for GitHub release upload
- [x] Document packaging results and artifact paths

# Review

- `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64` succeeded
- `"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installers/WinTab.iss` succeeded and generated `publish/installer/WinTab_Setup_3.0.0.exe`
- SHA256: `83c3fcb193e30af41d75a95453b4d7992d51aa2dca3faa2ed90cc80bcfa6b522`
- `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release` passed (2/2)
