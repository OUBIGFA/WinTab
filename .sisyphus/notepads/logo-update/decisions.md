2026-03-02: Kept all existing reference paths unchanged and replaced only logo assets (`logo.png` and `wintab.ico`) to minimize scope.
2026-03-02: Regenerated `wintab.ico` directly from the new `logo.png` as a PNG-in-ICO single-image icon, ensuring installer/app icon source consistency.
2026-03-02: Validation gate for completion is `dotnet build WinTab.slnx -c Release` success plus hash match proof.
