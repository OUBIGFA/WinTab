# WinTab Design Polish Phase 1 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Upgrade WinTab's desktop UI toward a more polished, native-feeling Windows 11 experience without changing product scope or core behaviors.

**Architecture:** Keep the existing WPF + WPF-UI stack and refactor presentation at the resource, shell, page-layout, and localization layers. Centralize visual rhythm in `App.xaml`, then recompose the existing pages so hierarchy, danger states, and helper content are clearer while preserving all current bindings and commands.

**Tech Stack:** .NET 9, WPF, WPF-UI, CommunityToolkit.Mvvm, localized XAML resource dictionaries, Inno Setup

---

### Task 1: Establish shared visual tokens and layout primitives

**Files:**
- Modify: `src/WinTab.App/App.xaml`
- Verify: `src/WinTab.App/Views/MainWindow.xaml`

**Steps:**
1. Add shared brushes for subtle surfaces, emphasized surfaces, and destructive surfaces.
2. Add shared typography, page-header, section-label, helper-text, and card spacing styles.
3. Add reusable layout values so pages stop repeating ad-hoc font sizes and margins.
4. Keep all changes compatible with existing WPF-UI dynamic theme resources.

### Task 2: Upgrade the application shell

**Files:**
- Modify: `src/WinTab.App/Views/MainWindow.xaml`
- Verify: `src/WinTab.App/Views/MainWindow.xaml.cs`

**Steps:**
1. Improve shell spacing and create a stronger content container.
2. Add a compact top summary area that reinforces app identity and navigation context.
3. Strengthen navigation clarity without changing the current route model.
4. Preserve tray/minimize and default navigation behavior.

### Task 3: Recompose the settings pages for clearer hierarchy

**Files:**
- Modify: `src/WinTab.App/Views/Pages/GeneralPage.xaml`
- Modify: `src/WinTab.App/Views/Pages/BehaviorPage.xaml`

**Steps:**
1. Add page headers with concise titles and support text.
2. Group settings into clearer sections with stronger visual separation.
3. Distinguish system-impacting options from convenience options.
4. Preserve all existing bindings and behavior.

### Task 4: Improve About and Uninstall clarity

**Files:**
- Modify: `src/WinTab.App/Views/Pages/AboutPage.xaml`
- Modify: `src/WinTab.App/Views/Pages/UninstallPage.xaml`

**Steps:**
1. Improve About page hierarchy and make the branding area feel intentional, not decorative.
2. Fix narrow-layout copy risks and make diagnostics/actions easier to scan.
3. Separate destructive uninstall actions from safe helper actions.
4. Preserve every existing command and visibility rule.

### Task 5: Refine bilingual UX copy

**Files:**
- Modify: `src/WinTab.UI/Localization/Strings.zh-CN.xaml`
- Modify: `src/WinTab.UI/Localization/Strings.en-US.xaml`

**Steps:**
1. Tighten titles, labels, and descriptions to be more direct and product-like.
2. Keep terminology aligned between navigation, settings, uninstall, and tray interactions.
3. Preserve all existing keys unless a new UI surface requires a new key.

### Task 6: Verify and package

**Files:**
- Verify: `WinTab.slnx`
- Verify: `src/WinTab.Tests/WinTab.Tests.csproj`
- Package: `installers/WinTab.iss`

**Steps:**
1. Run diagnostics on modified code files where supported.
2. Run `dotnet build WinTab.slnx -c Release`.
3. Run `dotnet test src/WinTab.Tests/WinTab.Tests.csproj -c Release`.
4. Run Inno Setup compile to overwrite the current installer output.
5. Report exact verification commands and resulting installer path.
