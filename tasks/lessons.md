# Lessons

- When the user clarifies default behavior (for example, uninstall should preserve settings by default), treat that preference as a hard requirement and implement an explicit opt-in switch for the destructive alternative.
- For new interaction features, default to off unless the user explicitly asks to enable by default, and always expose the toggle in the user-facing Behavior settings page when requested.
- For Explorer tab interactions, do not assume cursor-hit child windows are `ShellTabWindowClass`; resolve Explorer top-level from the cursor first, then target the active tab handle under that top-level to avoid no-op click handling.
