# Recycle Bin Native Routing Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Preserve existing folder tab-merging behavior while forcing Recycle Bin opens to use the native Windows shell path with no extra dependencies.

**Architecture:** Introduce a shared open-target classification layer that distinguishes physical file-system targets from native-only shell namespace targets. Route Recycle Bin requests away from the pipe/tab flow at ingress, unify native shell launching for namespace fallback, and keep the existing Explorer tab executor for normal folders.

**Tech Stack:** .NET 9, WPF, xUnit, FluentAssertions, Moq, Inno Setup 6

---

### Task 1: Lock Reproducible Failures in Tests

**Files:**
- Modify: `src/WinTab.Tests/App/ExplorerOpenVerbHandlerTests.cs`
- Modify: `src/WinTab.Tests/App/ExplorerOpenRequestServerTests.cs`
- Modify: `src/WinTab.Tests/App/ShellComNavigatorTests.cs`
- Create: `src/WinTab.Tests/App/OpenTargetRoutingTests.cs`

**Step 1: Write failing tests for Recycle Bin routing**

Cover:
- handler path must bypass pipe for Recycle Bin and use native shell launch
- ShellBridge delegate-execute path must bypass pipe for Recycle Bin and use native shell launch
- fallback for shell namespace targets must not use raw `explorer.exe "::{GUID}"`
- normal physical folders must still use existing pipe/tab path

**Step 2: Run targeted tests to verify they fail**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "OpenTargetRoutingTests|ExplorerOpenVerbHandlerTests|ExplorerOpenRequestServerTests|ShellComNavigatorTests"`

Expected: FAIL on new routing expectations because the current implementation still forwards Recycle Bin through the pipe or raw fallback path.

**Step 3: Keep failing test output for reference**

Capture which assertions fail so the refactor addresses the exact behavior mismatch.

### Task 2: Introduce Shared Target Classification

**Files:**
- Create: `src/WinTab.Platform.Win32/OpenTargetKind.cs`
- Create: `src/WinTab.Platform.Win32/OpenTargetInfo.cs`
- Create: `src/WinTab.Platform.Win32/OpenTargetClassifier.cs`
- Modify: `src/WinTab.Platform.Win32/WinTab.Platform.Win32.csproj`
- Test: `src/WinTab.Tests/App/OpenTargetRoutingTests.cs`

**Step 1: Write the next failing classifier tests**

Cover:
- physical folder path => physical file-system target
- `::{645FF040-...}`, `shell:::{645FF040-...}`, `shell:RecycleBinFolder` => native-only shell target
- non-Recycle-Bin shell namespaces remain distinguishable from physical folders

**Step 2: Run classifier tests to verify they fail**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "OpenTargetRoutingTests"`

Expected: FAIL because classifier types do not exist yet.

**Step 3: Implement minimal classifier**

Keep it dependency-free and use the existing `ShellLocationIdentityService` normalization/equivalence behavior.

**Step 4: Re-run classifier tests**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "OpenTargetRoutingTests"`

Expected: PASS

### Task 3: Add Native Shell Launcher Abstraction

**Files:**
- Create: `src/WinTab.Platform.Win32/INativeShellLauncher.cs`
- Create: `src/WinTab.Platform.Win32/NativeShellLauncher.cs`
- Modify: `src/WinTab.App/App.xaml.cs`
- Modify: `src/WinTab.ShellBridge/WinTabOpenFolderDelegateExecute.cs`
- Test: `src/WinTab.Tests/App/OpenTargetRoutingTests.cs`

**Step 1: Write failing tests for native launcher usage**

Cover:
- Recycle Bin requests invoke the native shell launcher
- pipe send is skipped for native-only targets
- no extra dependencies or external binaries are introduced

**Step 2: Run tests to verify failure**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "OpenTargetRoutingTests|ExplorerOpenVerbHandlerTests"`

Expected: FAIL because no launcher abstraction exists yet.

**Step 3: Implement launcher abstraction**

Use Windows shell COM / existing platform APIs only. Do not add NuGet packages.

**Step 4: Re-run targeted tests**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "OpenTargetRoutingTests|ExplorerOpenVerbHandlerTests"`

Expected: PASS

### Task 4: Refactor Ingress Routing

**Files:**
- Modify: `src/WinTab.App/Services/ExplorerOpenVerbHandler.cs`
- Modify: `src/WinTab.ShellBridge/WinTabOpenFolderDelegateExecute.cs`
- Modify: `src/WinTab.App/Services/AppEnvironment.cs`
- Modify: `src/WinTab.ShellBridge/PathNormalization.cs`
- Test: `src/WinTab.Tests/App/ExplorerOpenVerbHandlerTests.cs`
- Test: `src/WinTab.Tests/App/OpenTargetRoutingTests.cs`

**Step 1: Write failing tests for ingress split**

Cover:
- Recycle Bin is opened natively before pipe forwarding
- folder paths still use the existing forwarding path
- invalid targets still fail safely

**Step 2: Run tests to verify failure**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "ExplorerOpenVerbHandlerTests|OpenTargetRoutingTests"`

Expected: FAIL on the new Recycle Bin native-bypass assertions.

**Step 3: Implement ingress routing split**

Keep handler signatures stable where possible to preserve external behavior.

**Step 4: Re-run targeted tests**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "ExplorerOpenVerbHandlerTests|OpenTargetRoutingTests"`

Expected: PASS

### Task 5: Refactor Main Explorer Routing and Fallback

**Files:**
- Modify: `src/WinTab.App/ExplorerTabUtilityPort/ExplorerTabHookService.cs`
- Modify: `src/WinTab.App/ExplorerTabUtilityPort/NativeBrowseFallbackBypassStore.cs`
- Modify: `src/WinTab.App/ExplorerTabUtilityPort/ShellComNavigator.cs`
- Test: `src/WinTab.Tests/App/ExplorerTabHookServiceHelpersTests.cs`
- Test: `src/WinTab.Tests/App/ShellComNavigatorTests.cs`
- Test: `src/WinTab.Tests/App/OpenTargetRoutingTests.cs`

**Step 1: Write failing tests for main-process routing**

Cover:
- native-only shell targets never enter tab executor logic
- namespace fallback never uses raw explorer command-line launch
- existing file-system path behavior remains unchanged

**Step 2: Run tests to verify failure**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "ExplorerTabHookServiceHelpersTests|ShellComNavigatorTests|OpenTargetRoutingTests"`

Expected: FAIL until the routing split is implemented.

**Step 3: Implement executor split and semantic cleanup**

Prefer explicit target-kind checks over string-prefix heuristics. Keep public behavior for physical folders unchanged.

**Step 4: Re-run targeted tests**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release --filter "ExplorerTabHookServiceHelpersTests|ShellComNavigatorTests|OpenTargetRoutingTests"`

Expected: PASS

### Task 6: Full Verification and Packaging

**Files:**
- Modify if needed: `installers/WinTab.iss`
- Overwrite output: `installers/WinTab_Setup_1.0.0.exe`

**Step 1: Run full test suite**

Run: `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release`

Expected: PASS

**Step 2: Publish release build**

Run: `dotnet publish src/WinTab.App/WinTab.App.csproj -c Release -r win-x64 --self-contained false -o publish/win-x64`

Expected: PASS and refreshed published app payload.

**Step 3: Build installer**

Run: `& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DAppVersion=1.0.0 installers/WinTab.iss`

Expected: PASS and overwrite `installers/WinTab_Setup_1.0.0.exe`.

**Step 4: Final review**

Run targeted diff review plus a final code-review pass before reporting completion.
