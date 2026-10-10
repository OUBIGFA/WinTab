import { call, on } from 'mygo-runtime'

export interface Settings {
  windowHook: boolean
  reuseTabs: boolean
  restoreTabs: boolean
  reopenClosedTab: boolean
  restoreGroupShortcutEnabled: boolean
  restoreGroupShortcut: string
  reopenTabShortcutEnabled: boolean
  reopenTabShortcut: string
  restoreOnAnyFolder: boolean
  restoreSingleTab: boolean
  doubleClickCloseTab: boolean
  doubleClickCloseIncludeNotepad: boolean
  middleClickForegroundTab: boolean
  wheelSwitchTab: boolean
  wheelSwitchSensitivity: 'Low' | 'Medium' | 'High'
  autoUpdate: boolean
  showTrayIcon: boolean
  language: 'zh-CN' | 'en-US'
  theme: 'Light' | 'Dark'
  formSize: { width: number; height: number } | null
}

export interface State {
  protocolVersion: number
  revision: number
  version: string
  settings: Settings
  startup: boolean
  shellReady: boolean
  recordClosedTabs: boolean
  sessionBusy: boolean
  updateBusy: boolean
  hookError: string | null
  shortcutError: string | null
  sessionFeedback: string | null
  updateFeedback: string | null
  storageError: string | null
}

export interface WindowInfo { width: number; maxWidth: number; maxHeight: number; hasSavedSize: boolean }
export interface Bootstrap { state: State; window: WindowInfo }
export type BooleanSetting = { [K in keyof Settings]: Settings[K] extends boolean ? K : never }[keyof Settings]
export type SettableKey = Exclude<BooleanSetting, 'restoreGroupShortcutEnabled' | 'reopenTabShortcutEnabled'>
  | 'startup' | 'language' | 'theme' | 'wheelSwitchSensitivity'

// The official runtime is the only transport. Browser tests inject its documented API;
// a packaged page never substitutes mock settings or a localStorage-only success path.
export const desktop = {
  load: () => call<Bootstrap>('Desktop.Load'),
  set: (key: SettableKey, value: boolean | string) => call<State>('Desktop.Set', key, value),
  shortcuts: (groupEnabled: boolean, group: string, tabEnabled: boolean, tab: string) =>
    call<State>('Desktop.Shortcuts', groupEnabled, group, tabEnabled, tab),
  restore: (group: boolean) => call<State>('Desktop.Restore', group),
  update: () => call<State>('Desktop.Update'),
  logs: () => call<State>('Desktop.Logs'),
  openProject: () => call<void>('Desktop.OpenProject'),
  ready: (width: number, height: number) => call<void>('Desktop.Ready', width, height),
  onState: (listener: (state: State) => void) => on<State>('state-changed', listener),
  onUserResize: (listener: () => void) => on<boolean>('user-resized', listener),
}
