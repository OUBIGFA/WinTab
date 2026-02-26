# Lessons

- When the user clarifies default behavior (for example, uninstall should preserve settings by default), treat that preference as a hard requirement and implement an explicit opt-in switch for the destructive alternative.
- For new interaction features, default to off unless the user explicitly asks to enable by default, and always expose the toggle in the user-facing Behavior settings page when requested.
- For Explorer tab interactions, do not assume cursor-hit child windows are `ShellTabWindowClass`; resolve Explorer top-level from the cursor first, then target the active tab handle under that top-level to avoid no-op click handling.
- For theme regressions that self-heal after toggling, always verify startup order and avoid local visual overrides (for example `Background="Transparent"` on the main window) that can bypass initial theme resources.
- If a reliability helper runs as the same executable (for example `WinTab.exe --wintab-companion`), treat it as a user-visible duplicate instance risk; prefer in-process recovery paths or one-shot handlers over long-lived sidecar processes.
- For Explorer tab-bar hit-testing, do not assume `ShellTabWindowClass` always sits below the title strip on every Windows build; avoid hard rejections based on point-within-tab-window unless validated by live telemetry.
- When users require tab-title-only behavior, enforce strict fallback gates (tab-window ancestry + narrow top-strip cap) so address/navigation/refresh regions cannot trigger close actions.
- Do not require tab-window ancestry as a hard gate for tab-title clicks; on some Explorer builds it can reject all valid tab-title interactions. Prefer role-based navigation exclusion plus a narrowly capped top-strip geometry fallback.
