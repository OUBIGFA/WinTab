import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { Check, CircleAlert, FolderOpen, Info, Languages, LoaderCircle, Moon, RefreshCw, RotateCcw, Settings2, Sun } from 'lucide-react'
import logo from '../../../WinTab/wintab-logo.png'
import { desktop, type SettableKey, type State, type WindowInfo } from './desktop'
import { copies } from './i18n'
import { Checkbox, SectionHeading, Segments, Setting, ShortcutInput, Toggle } from './components'

export function App() {
  const [state, setState] = useState<State | null>(null)
  const [windowInfo, setWindowInfo] = useState<WindowInfo | null>(null)
  const [error, setError] = useState(false)
  const [loadError, setLoadError] = useState(false)
  const [pending, setPending] = useState<Set<string>>(new Set())
  const pendingRef = useRef(new Set<string>())
  const [groupDraft, setGroupDraft] = useState('')
  const [tabDraft, setTabDraft] = useState('')
  const [shortcutApplied, setShortcutApplied] = useState(false)
  const pageRef = useRef<HTMLDivElement>(null)
  const accept = useCallback((next: State) => {
    setState(previous => !previous || next.revision > previous.revision ? next : previous)
  }, [])
  const language = state?.settings.language ?? (navigator.language.startsWith('zh') ? 'zh-CN' : 'en-US')
  const t = copies[language]

  useEffect(() => {
    let live = true
    const unsubscribe = desktop.onState(accept)
    desktop.load().then(result => {
      if (!live) return
      if (result.state.protocolVersion !== 1) throw new Error('Unsupported protocol')
      accept(result.state)
      setWindowInfo(result.window)
    }).catch(() => { if (live) setLoadError(true) })
    return () => { live = false; unsubscribe() }
  }, [accept])

  useLayoutEffect(() => {
    document.documentElement.lang = language
    document.documentElement.dataset.theme = state?.settings.theme ?? 'Light'
  }, [language, state?.settings.theme])

  // Only changes to saved values replace drafts. Tray updates and theme/language changes must not
  // erase an unapplied shortcut, including a draft whose field no longer has keyboard focus.
  useEffect(() => { setGroupDraft(state?.settings.restoreGroupShortcut ?? '') }, [state?.settings.restoreGroupShortcut])
  useEffect(() => { setTabDraft(state?.settings.reopenTabShortcut ?? '') }, [state?.settings.reopenTabShortcut])
  useEffect(() => desktop.onUserResize(() => {
    setWindowInfo(info => info ? { ...info, hasSavedSize: true } : info)
  }), [])

  useEffect(() => {
    if (!windowInfo) return
    let live = true
    const fit = async () => {
      await document.fonts.ready
      await new Promise<void>(resolve => requestAnimationFrame(() => resolve()))
      if (!live || !pageRef.current) return
      const page = pageRef.current
      page.style.zoom = '1'
      page.style.width = ''
      if (windowInfo.hasSavedSize) {
        await desktop.ready(windowInfo.width, innerHeight)
        return
      }
      const naturalWidth = Math.min(windowInfo.maxWidth, document.documentElement.clientWidth)
      const titlebarHeight = document.querySelector('.titlebar')!.getBoundingClientRect().height
      const availableHeight = windowInfo.maxHeight - titlebarHeight
      // Freeze the logical width while measuring. Zoom rounds font/line metrics, so scaling the
      // original height is not enough; measure the resulting layout before showing the window.
      page.style.width = `${naturalWidth}px`
      let height = page.getBoundingClientRect().height
      let scale = Math.min(1, (availableHeight - 1) / height)
      page.style.zoom = String(scale)
      height = page.getBoundingClientRect().height
      while (Math.ceil(height) > availableHeight) {
        scale *= (availableHeight - 1) / height
        page.style.zoom = String(scale)
        height = page.getBoundingClientRect().height
      }
      await desktop.ready(Math.ceil(naturalWidth * scale), Math.ceil(height + titlebarHeight))
      await new Promise<void>(resolve => requestAnimationFrame(() => resolve()))
      page.style.width = ''
    }
    void fit().catch(() => setLoadError(true))
    return () => { live = false }
  }, [windowInfo])

  const run = async (key: string, action: () => Promise<State>): Promise<State | null> => {
    if (pendingRef.current.has(key)) return null
    pendingRef.current.add(key)
    setPending(new Set(pendingRef.current))
    setError(false)
    try {
      const next = await action()
      accept(next)
      return next
    } catch {
      setError(true)
      return null
    } finally {
      pendingRef.current.delete(key)
      setPending(new Set(pendingRef.current))
    }
  }

  if (!state || loadError) return <main className="connection-view" role={loadError ? 'alert' : 'status'}>
    {loadError ? <CircleAlert size={28} /> : <LoaderCircle className="spinning" size={28} />}
    <p>{loadError ? t.loadError : t.connecting}</p>
  </main>

  const s = state.settings
  const dirty = groupDraft !== s.restoreGroupShortcut || tabDraft !== s.reopenTabShortcut
  const restoring = state.sessionBusy || pending.has('restore')
  const updating = state.updateBusy || pending.has('update')
  const set = (key: SettableKey, value: string | boolean) => { void run(key, () => desktop.set(key, value)) }
  const toggle = (key: SettableKey, title: string, checked: boolean, disabled = false, described = true) =>
    <Toggle id={key} label={title} description={described ? `${key}-description` : undefined} checked={checked}
      disabled={disabled || pending.has(key)} onChange={value => set(key, value)} />
  const saveShortcuts = async (groupEnabled = s.restoreGroupShortcutEnabled, tabEnabled = s.reopenTabShortcutEnabled) => {
    setShortcutApplied(false)
    const next = await run('shortcuts', () => desktop.shortcuts(groupEnabled, groupDraft, tabEnabled, tabDraft))
    if (next && !next.shortcutError) setShortcutApplied(true)
  }
  const restoreButton = (group: boolean) => <button className="button button-outline restore-button"
    disabled={restoring || !state.shellReady || (!group && !state.recordClosedTabs)}
    onClick={() => { void run('restore', () => desktop.restore(group)) }}>
    {restoring ? <LoaderCircle size={15} className="spinning" /> : <RotateCcw size={15} />}
    {restoring ? t.restoring : t.restore}
  </button>
  const shortcutLine = (group: boolean) => <div className="shortcut-line">
    <Checkbox id={group ? 'enable-group' : 'enable-tab'} label={group ? t.enableGroup : t.enableTab}
      checked={group ? s.restoreGroupShortcutEnabled : s.reopenTabShortcutEnabled} disabled={pending.has('shortcuts')}
      onChange={enabled => { void saveShortcuts(group ? enabled : s.restoreGroupShortcutEnabled, group ? s.reopenTabShortcutEnabled : enabled) }}>
      {t.shortcut}
    </Checkbox>
    <ShortcutInput id={group ? 'group-shortcut' : 'tab-shortcut'} label={group ? t.groupShortcut : t.tabShortcut}
      value={group ? groupDraft : tabDraft} invalid={!!state.shortcutError} hint={t.captureHelp}
      onChange={value => { (group ? setGroupDraft : setTabDraft)(value); setShortcutApplied(false) }} />
    <span className="scope-label">{group ? t.global : t.inExplorer}</span>
  </div>

  return <>
    <div className="titlebar">
      <div className="titlebar-brand"><img src={logo} alt="" /><span>WinTab</span><span className="titlebar-divider" /><span className="titlebar-section">{t.preferences}</span></div>
      <div className="titlebar-actions">
        <button className="icon-button" aria-label={s.theme === 'Dark' ? t.light : t.dark} title={s.theme === 'Dark' ? t.light : t.dark}
          disabled={pending.has('theme')} onClick={() => set('theme', s.theme === 'Dark' ? 'Light' : 'Dark')}>
          {s.theme === 'Dark' ? <Sun size={17} /> : <Moon size={17} />}
        </button>
        <button className="icon-button" aria-label={t.language} title={t.language} disabled={pending.has('language')}
          onClick={() => set('language', language === 'zh-CN' ? 'en-US' : 'zh-CN')}><Languages size={18} /></button>
      </div>
    </div>

    <div className="page" ref={pageRef}>
      <header className="page-header">
        <h1>{t.preferences}</h1>
        <p className="background-note" role="status">{pending.size ? t.saving : state.shellReady ? t.running : t.connecting}</p>
      </header>

      <main className="settings-grid">
        <div className="settings-column">
          <section aria-labelledby="explorer-heading">
            <SectionHeading id="explorer-heading" icon={FolderOpen} title={t.explorer} />
            <div className="settings-surface">
              <Setting id="windowHook" title={t.merge} description={t.mergeDesc} control={toggle('windowHook', t.merge, s.windowHook)} />
              <Setting id="reuseTabs" title={t.reuse} description={t.reuseDesc} control={toggle('reuseTabs', t.reuse, s.reuseTabs)} />
              <Setting id="doubleClickCloseTab" title={t.doubleClick} control={toggle('doubleClickCloseTab', t.doubleClick, s.doubleClickCloseTab, false, false)}>
                <div className="option-line"><span className="option-label">{t.scope}</span>
                  <Segments label={t.scope} value={s.doubleClickCloseIncludeNotepad ? 'include' : 'explorer'}
                    options={[{ value: 'explorer', label: t.explorerOnly }, { value: 'include', label: t.includeNotepad }]}
                    disabled={!s.doubleClickCloseTab || pending.has('doubleClickCloseIncludeNotepad')}
                    onChange={value => set('doubleClickCloseIncludeNotepad', value === 'include')} />
                </div>
              </Setting>
              <Setting id="middleClickForegroundTab" title={t.middleClick} description={t.middleClickDesc} control={toggle('middleClickForegroundTab', t.middleClick, s.middleClickForegroundTab)} />
              <Setting id="wheelSwitchTab" title={t.wheel} description={t.wheelDesc} control={toggle('wheelSwitchTab', t.wheel, s.wheelSwitchTab)}>
                <div className="option-line"><span className="option-label">{t.sensitivity}</span>
                  <Segments label={t.sensitivity} value={s.wheelSwitchSensitivity}
                    options={[{ value: 'Low', label: t.low }, { value: 'Medium', label: t.medium }, { value: 'High', label: t.high }]}
                    disabled={!s.wheelSwitchTab || pending.has('wheelSwitchSensitivity')}
                    onChange={value => set('wheelSwitchSensitivity', value)} />
                </div>
                <p className="fine-print">{s.wheelSwitchSensitivity === 'Low' ? t.lowHint : s.wheelSwitchSensitivity === 'High' ? t.highHint : t.mediumHint}</p>
              </Setting>
            </div>
            <p className="bypass-hint"><Info size={14} /><span>{t.bypass} <kbd>Ctrl</kbd> + <kbd>Shift</kbd> {t.bypassEnd}</span></p>
          </section>

          <section aria-labelledby="system-heading">
            <SectionHeading id="system-heading" icon={Settings2} title={t.system} />
            <div className="settings-surface">
              <Setting compact id="startup" title={t.startup} description={t.startupDesc} control={toggle('startup', t.startup, state.startup)} />
              <Setting compact id="showTrayIcon" title={t.tray} description={s.showTrayIcon ? t.trayDesc : t.trayHidden} control={toggle('showTrayIcon', t.tray, s.showTrayIcon)} />
              <Setting compact id="autoUpdate" title={t.autoUpdate} description={t.autoUpdateDesc} control={toggle('autoUpdate', t.autoUpdate, s.autoUpdate)} />
            </div>
          </section>
        </div>

        <section className="recovery-section" aria-labelledby="recovery-heading">
          <SectionHeading id="recovery-heading" icon={RotateCcw} title={t.recovery} />
          <div className="settings-surface recovery-surface">
            <p className="recovery-intro">{t.recoveryDesc}</p>
            <div className="recovery-action">
              <div className="recovery-action-heading"><h3>{t.group}</h3>{restoreButton(true)}</div>
              {shortcutLine(true)}
              <Setting compact id="restoreSingleTab" title={t.single}
                control={toggle('restoreSingleTab', t.single, s.restoreSingleTab, false, false)} />
            </div>
            <div className="recovery-action">
              <div className="recovery-action-heading"><h3>{t.tab}</h3>{restoreButton(false)}</div>
              {shortcutLine(false)}
              <Setting compact id="reopenClosedTab" title={t.record} description={s.reopenTabShortcutEnabled ? t.recordRequired : t.recordHint}
                control={toggle('reopenClosedTab', t.record, state.recordClosedTabs, s.reopenTabShortcutEnabled)} />
            </div>
            <div className="shortcut-save">
              <div><p className="description">{t.captureHint}</p><p className="sr-only" id="shortcut-help">{t.captureHelp}</p></div>
              <button className={`button ${dirty ? 'button-primary' : 'button-outline'}`} disabled={pending.has('shortcuts') || !dirty}
                onClick={() => { void saveShortcuts() }}><Check size={15} />{t.apply}</button>
            </div>
            {(dirty || shortcutApplied || state.shortcutError) && <p className="inline-feedback" role={state.shortcutError ? 'alert' : 'status'}>
              {state.shortcutError ? <CircleAlert size={14} /> : <Check size={14} />}
              {state.shortcutError || (dirty ? t.draft : t.applied)}
            </p>}
            <div className="automatic-recovery">
              <Setting id="restoreTabs" title={t.autoRestore} description={t.autoRestoreDesc} control={toggle('restoreTabs', t.autoRestore, s.restoreTabs)}>
                <Segments label={t.restoreMode} value={s.restoreOnAnyFolder ? 'any' : 'normal'}
                  options={[{ value: 'normal', label: t.normalLaunch }, { value: 'any', label: t.anyFolder }]}
                  disabled={!s.restoreTabs || pending.has('restoreOnAnyFolder')} onChange={value => set('restoreOnAnyFolder', value === 'any')} />
                <p className="fine-print">{s.restoreTabs ? (s.restoreOnAnyFolder ? t.anyHint : t.normalHint) : t.autoIndependent}</p>
              </Setting>
              <p className="recovery-notes">{t.recoveryNotes}</p>
            </div>
          </div>
          {state.sessionFeedback && <p className="section-feedback" role="status"><Info size={15} />{state.sessionFeedback}</p>}
        </section>
      </main>

      {(error || state.hookError || state.storageError) && <div className="notice" role="alert"><CircleAlert size={17} /><div>
        {error && <p>{t.requestError}</p>}{state.hookError && <p>{state.hookError}</p>}{state.storageError && <p>{state.storageError}</p>}
      </div></div>}

      <section className="about-section" aria-labelledby="about-heading">
        <SectionHeading id="about-heading" icon={Info} title={t.about} />
        <footer className="about">
          <div className="about-details">
            <h3 className="about-version">WinTab v{state.version}</h3>
            <p className="about-links"><span>{t.license}</span><span aria-hidden="true">·</span>
              <a href="https://github.com/OUBIGFA/WinTab" title={t.openSource} onClick={event => {
                event.preventDefault()
                void desktop.openProject().catch(() => setError(true))
              }}>GitHub</a>
            </p>
            {state.updateFeedback && <p className="update-feedback" role="status">{state.updateFeedback}</p>}
          </div>
          <div className="about-actions">
            <button className="button button-outline" disabled={pending.has('logs')} onClick={() => { void run('logs', desktop.logs) }}><FolderOpen size={16} />{t.logs}</button>
            <button className="button button-outline" disabled={updating} onClick={() => { void run('update', desktop.update) }}>
              {updating ? <LoaderCircle size={16} className="spinning" /> : <RefreshCw size={16} />}{updating ? t.checking : t.updates}
            </button>
          </div>
        </footer>
      </section>
    </div>
  </>
}
