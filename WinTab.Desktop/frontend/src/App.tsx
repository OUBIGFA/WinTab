import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { ArrowUpRight, Check, ChevronRight, CircleAlert, FolderOpen, Info, Languages, LoaderCircle, Moon, RefreshCw, RotateCcw, Settings2, Sun } from 'lucide-react'
import logo from '../../../WinTab/wintab-logo.png'
import { desktop, type SettableKey, type State } from './desktop'
import { copies } from './i18n'
import { Checkbox, Segments, Setting, ShortcutInput, Toggle } from './components'

type Section = 'explorer' | 'recovery' | 'preferences' | 'about'

export function App() {
  const [state, setState] = useState<State | null>(null)
  const [section, setSection] = useState<Section>('explorer')
  const [loadAttempt, setLoadAttempt] = useState(0)
  const [loadError, setLoadError] = useState(false)
  const [error, setError] = useState(false)
  const [pending, setPending] = useState<Set<string>>(new Set())
  const pendingRef = useRef(new Set<string>())
  const [groupDraft, setGroupDraft] = useState('')
  const [tabDraft, setTabDraft] = useState('')
  const [shortcutApplied, setShortcutApplied] = useState(false)
  const contentRef = useRef<HTMLElement>(null)
  const headingRef = useRef<HTMLHeadingElement>(null)
  const navRef = useRef<HTMLElement>(null)
  const accept = useCallback((next: State) => {
    setState(previous => !previous || next.revision > previous.revision ? next : previous)
  }, [])
  const language = state?.settings.language ?? (navigator.language.startsWith('zh') ? 'zh-CN' : 'en-US')
  const t = copies[language]

  useEffect(() => desktop.onState(accept), [accept])
  useEffect(() => {
    let live = true
    setLoadError(false)
    const load = async () => {
      try {
        const result = await desktop.load()
        if (!live) return
        if (result.state.protocolVersion !== 1) throw new Error('Unsupported protocol')
        accept(result.state)
        await document.fonts.ready
        if (!live) return
        // Pages scroll at their natural text size. A first launch has a comfortable fixed size;
        // persisted dimensions belong to the native window and must never be overwritten here.
        await desktop.ready(Math.min(result.window.width, result.window.maxWidth), Math.min(780, result.window.maxHeight))
      } catch {
        if (!live) return
        setLoadError(true)
        // Reveal the retry surface even when bootstrap fails, instead of hiding the window until timeout.
        try { await desktop.ready(innerWidth, innerHeight) } catch (revealError) { console.error('Could not reveal the connection screen', revealError) }
      }
    }
    void load()
    return () => { live = false }
  }, [accept, loadAttempt])

  useLayoutEffect(() => {
    document.documentElement.lang = language
    document.documentElement.dataset.theme = state?.settings.theme ?? 'Light'
  }, [language, state?.settings.theme])

  // Drafts live above the pages, so navigation and unrelated backend updates cannot discard input.
  useEffect(() => { setGroupDraft(state?.settings.restoreGroupShortcut ?? '') }, [state?.settings.restoreGroupShortcut])
  useEffect(() => { setTabDraft(state?.settings.reopenTabShortcut ?? '') }, [state?.settings.reopenTabShortcut])

  // Keep the active pill visible when the navigation wraps into a horizontal strip at narrow widths.
  useEffect(() => {
    const active = navRef.current?.querySelector<HTMLElement>('[aria-current="page"]')
    active?.scrollIntoView({ block: 'nearest', inline: 'nearest' })
  }, [section, language])

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
  const set = (key: SettableKey, value: string | boolean) => { void run(key, () => desktop.set(key, value)) }
  const navigate = (next: Section) => {
    setSection(next)
    requestAnimationFrame(() => {
      contentRef.current?.scrollTo({ top: 0 })
      headingRef.current?.focus({ preventScroll: true })
    })
  }
  const sections = [
    { id: 'explorer', label: t.explorer, description: t.explorerIntro, icon: FolderOpen },
    { id: 'recovery', label: t.recoveryNav, description: t.recoveryIntro, icon: RotateCcw },
    { id: 'preferences', label: t.preferences, description: t.preferencesIntro, icon: Settings2 },
    { id: 'about', label: t.about, description: t.aboutIntro, icon: Info },
  ] as const
  const current = sections.find(item => item.id === section)!
  const s = state?.settings
  const dirty = !!s && (groupDraft !== s.restoreGroupShortcut || tabDraft !== s.reopenTabShortcut)
  const restoring = state?.sessionBusy || pending.has('restore')
  const updating = state?.updateBusy || pending.has('update')
  const toggle = (key: SettableKey, title: string, checked: boolean, disabled = false, described = true) =>
    <Toggle id={key} label={title} description={described ? `${key}-description` : undefined} checked={checked}
      disabled={disabled || pending.has(key)} onChange={value => set(key, value)} />
  const saveShortcuts = async (groupEnabled = s!.restoreGroupShortcutEnabled, tabEnabled = s!.reopenTabShortcutEnabled) => {
    setShortcutApplied(false)
    const next = await run('shortcuts', () => desktop.shortcuts(groupEnabled, groupDraft, tabEnabled, tabDraft))
    if (next && !next.shortcutError) setShortcutApplied(true)
  }
  const restoreButton = (group: boolean) => <button type="button" className="button button-outline"
    disabled={restoring || !state?.shellReady || (!group && !state.recordClosedTabs)}
    onClick={() => { void run('restore', () => desktop.restore(group)) }}>
    {restoring ? <LoaderCircle className="spinning" aria-hidden="true" /> : <RotateCcw aria-hidden="true" />}
    <span>{restoring ? t.restoring : t.restore}</span>
  </button>
  const shortcutLine = (group: boolean) => <div className="shortcut-line">
    <Checkbox id={group ? 'enable-group' : 'enable-tab'} label={group ? t.enableGroup : t.enableTab}
      checked={group ? s!.restoreGroupShortcutEnabled : s!.reopenTabShortcutEnabled} disabled={pending.has('shortcuts')}
      onChange={enabled => { void saveShortcuts(group ? enabled : s!.restoreGroupShortcutEnabled, group ? s!.reopenTabShortcutEnabled : enabled) }}>{t.shortcut}</Checkbox>
    <ShortcutInput id={group ? 'group-shortcut' : 'tab-shortcut'} label={group ? t.groupShortcut : t.tabShortcut}
      value={group ? groupDraft : tabDraft} invalid={!!state?.shortcutError} hint={t.captureHelp}
      onChange={value => { (group ? setGroupDraft : setTabDraft)(value); setShortcutApplied(false) }} />
  </div>

  return <div className="app-shell">
    <div className="sidebar-brand">
      <img src={logo} alt="" />
      <span>WinTab</span>
    </div>
    <header className="titlebar">
      <div className="titlebar-brand">
        <img src={logo} alt="" />
        <span>WinTab</span>
      </div>
      <div className="titlebar-breadcrumb" aria-hidden="true">
        <span>{t.settings}</span>
        <ChevronRight />
        <span>{current.label}</span>
      </div>
      <div className="titlebar-actions">
        <button type="button" className="icon-button" aria-label={s?.theme === 'Dark' ? t.light : t.dark} title={s?.theme === 'Dark' ? t.light : t.dark}
          disabled={!s || pending.has('theme')} onClick={() => set('theme', s?.theme === 'Dark' ? 'Light' : 'Dark')}>
          {s?.theme === 'Dark' ? <Sun /> : <Moon />}
        </button>
        <button type="button" className="icon-button" aria-label={t.language} title={t.language} disabled={!s || pending.has('language')}
          onClick={() => set('language', language === 'zh-CN' ? 'en-US' : 'zh-CN')}><Languages /></button>
        <span className="titlebar-divider" aria-hidden="true" />
      </div>
    </header>
    <aside className="sidebar">
      <nav ref={navRef} aria-label={t.navigation}>
        {sections.map(({ id, label, icon: Icon }) => <button type="button" key={id} className="nav-item" aria-current={section === id ? 'page' : undefined}
          onClick={() => navigate(id)}><Icon aria-hidden="true" /><span>{label}</span></button>)}
      </nav>
      <div className="sidebar-footer">
        <p className="connection-status" role="status">
          {loadError ? <CircleAlert aria-hidden="true" />
            : pending.size || !state?.shellReady ? <LoaderCircle className="spinning" aria-hidden="true" />
            : <Check aria-hidden="true" />}
          <span>{loadError ? t.loadErrorTitle : pending.size ? t.saving : state?.shellReady ? t.connected : t.connecting}</span>
        </p>
        {state && <span className="sidebar-version">WinTab v{state.version}</span>}
      </div>
    </aside>
    <main className="workspace" ref={contentRef} id="main-content">
      {!state || loadError ? <div className="connection-view" role={loadError ? 'alert' : 'status'}>
        {loadError ? <CircleAlert aria-hidden="true" /> : <LoaderCircle className="spinning" aria-hidden="true" />}
        <h1>{loadError ? t.loadErrorTitle : t.loadTitle}</h1>
        <p>{loadError ? t.loadError : t.connecting}</p>
        {loadError && <button type="button" className="button button-primary" onClick={() => setLoadAttempt(value => value + 1)}><RefreshCw aria-hidden="true" /><span>{t.retry}</span></button>}
      </div> : <div className="page">
        <header className="page-header">
          <h1 ref={headingRef} tabIndex={-1}>{current.label}</h1>
          <p>{current.description}</p>
        </header>
        {(error || state.hookError || state.storageError) && <div className="notice" role="alert"><CircleAlert aria-hidden="true" /><div>
          {error && <p>{t.requestError}</p>}{state.hookError && <p>{state.hookError}</p>}{state.storageError && <p>{state.storageError}</p>}
        </div></div>}

        <section hidden={section !== 'explorer'} aria-label={t.explorer}>
          <div className="settings-group">
            <h2>{t.windowsAndTabs}</h2>
            <div className="settings-card">
              <Setting id="windowHook" title={t.merge} description={t.mergeDesc} control={toggle('windowHook', t.merge, s!.windowHook)} />
              <Setting id="reuseTabs" title={t.reuse} description={t.reuseDesc} control={toggle('reuseTabs', t.reuse, s!.reuseTabs)} />
            </div>
          </div>
          <div className="settings-group">
            <h2>{t.mouseActions}</h2>
            <div className="settings-card">
              <Setting id="doubleClickCloseTab" title={t.doubleClick} description={t.doubleClickDesc}
                control={toggle('doubleClickCloseTab', t.doubleClick, s!.doubleClickCloseTab)}>
                <div className="option-line">
                  <span className="option-label">{t.scope}</span>
                  <Segments label={t.scope} value={s!.doubleClickCloseIncludeNotepad ? 'include' : 'explorer'}
                    options={[{ value: 'explorer', label: t.explorerOnly }, { value: 'include', label: t.includeNotepad }]}
                    disabled={!s!.doubleClickCloseTab || pending.has('doubleClickCloseIncludeNotepad')}
                    onChange={value => set('doubleClickCloseIncludeNotepad', value === 'include')} />
                </div>
              </Setting>
              <Setting id="doubleClickEmptyArea" title={t.emptyDoubleClick} description={t.emptyDoubleClickDesc}
                control={toggle('doubleClickEmptyArea', t.emptyDoubleClick, s!.doubleClickEmptyArea)}>
                <div className="option-line">
                  <span className="option-label">{t.emptyAction}</span>
                  <Segments label={t.emptyAction} value={s!.doubleClickEmptyAreaAction}
                    options={[{ value: 'Parent', label: t.parent }, { value: 'Back', label: t.back }]}
                    disabled={!s!.doubleClickEmptyArea || pending.has('doubleClickEmptyAreaAction')}
                    onChange={value => set('doubleClickEmptyAreaAction', value)} />
                </div>
              </Setting>
              <Setting id="middleClickForegroundTab" title={t.middleClick} description={t.middleClickDesc} control={toggle('middleClickForegroundTab', t.middleClick, s!.middleClickForegroundTab)} />
              <Setting id="wheelSwitchTab" title={t.wheel} description={t.wheelDesc} control={toggle('wheelSwitchTab', t.wheel, s!.wheelSwitchTab)}>
                <div className="option-line">
                  <div className="option-copy">
                    <span className="option-label">{t.sensitivity}</span>
                    <p className="fine-print">{s!.wheelSwitchSensitivity === 'Low' ? t.lowHint : s!.wheelSwitchSensitivity === 'High' ? t.highHint : t.mediumHint}</p>
                  </div>
                  <Segments label={t.sensitivity} value={s!.wheelSwitchSensitivity}
                    options={[{ value: 'Low', label: t.low }, { value: 'Medium', label: t.medium }, { value: 'High', label: t.high }]}
                    disabled={!s!.wheelSwitchTab || pending.has('wheelSwitchSensitivity')}
                    onChange={value => set('wheelSwitchSensitivity', value)} />
                </div>
              </Setting>
            </div>
          </div>
          <p className="context-note"><Info aria-hidden="true" /><span>{t.bypass} <kbd>Ctrl</kbd> + <kbd>Shift</kbd> {t.bypassEnd}</span></p>
        </section>

        <section hidden={section !== 'recovery'} aria-label={t.recovery}>
          <div className="settings-group">
            <h2>{t.manualRecovery}</h2>
            <div className="settings-card">
              <div className="setting recovery-command">
                <div className="setting-main">
                  <div className="setting-copy">
                    <h3 className="setting-title">{t.group}</h3>
                    <p className="description">{t.global}</p>
                  </div>
                  {restoreButton(true)}
                </div>
                <div className="recovery-shortcut">
                  {shortcutLine(true)}
                </div>
                <div className="setting-subrow">
                  <div className="setting-copy">
                    <label htmlFor="restoreSingleTab" className="setting-subtitle">{t.single}</label>
                    <p id="restoreSingleTab-description" className="description">{t.singleDesc}</p>
                  </div>
                  {toggle('restoreSingleTab', t.single, s!.restoreSingleTab)}
                </div>
              </div>
              <div className="setting recovery-command">
                <div className="setting-main">
                  <div className="setting-copy">
                    <h3 className="setting-title">{t.tab}</h3>
                    <p className="description">{t.inExplorer}</p>
                  </div>
                  {restoreButton(false)}
                </div>
                <div className="recovery-shortcut">
                  {shortcutLine(false)}
                </div>
                <div className="setting-subrow">
                  <div className="setting-copy">
                    <label htmlFor="reopenClosedTab" className="setting-subtitle">{t.record}</label>
                    <p id="reopenClosedTab-description" className="description">{s!.reopenTabShortcutEnabled ? t.recordRequired : t.recordHint}</p>
                  </div>
                  {toggle('reopenClosedTab', t.record, state.recordClosedTabs, s!.reopenTabShortcutEnabled)}
                </div>
              </div>
              <div className="shortcut-save">
                <div className="shortcut-save-copy">
                  <p className="shortcut-save-title">{t.captureHint}</p>
                  <p className="shortcut-save-note">{t.shortcutsNote}</p>
                  <p className="sr-only" id="shortcut-help">{t.captureHelp}</p>
                  {(dirty || shortcutApplied || state.shortcutError) && <p className="inline-feedback" role={state.shortcutError ? 'alert' : 'status'}>
                    {state.shortcutError ? <CircleAlert aria-hidden="true" /> : <Check aria-hidden="true" />}
                    <span>{state.shortcutError || (dirty ? t.draft : t.applied)}</span>
                  </p>}
                </div>
                <button type="button" className={`button ${dirty ? 'button-primary' : 'button-outline'}`} disabled={pending.has('shortcuts') || !dirty}
                  onClick={() => { void saveShortcuts() }}><Check aria-hidden="true" /><span>{t.apply}</span></button>
              </div>
            </div>
          </div>
          {state.sessionFeedback && <p className="section-feedback" role="status"><Info aria-hidden="true" /><span>{state.sessionFeedback}</span></p>}
          <div className="settings-group">
            <h2>{t.automatic}</h2>
            <div className="settings-card">
              <Setting id="restoreTabs" title={t.autoRestore} description={t.autoRestoreDesc} control={toggle('restoreTabs', t.autoRestore, s!.restoreTabs)}>
                <div className="option-line">
                  <div className="option-copy">
                    <span className="option-label">{t.restoreMode}</span>
                    <p className="fine-print">{s!.restoreTabs ? (s!.restoreOnAnyFolder ? t.anyHint : t.normalHint) : t.autoIndependent}</p>
                  </div>
                  <Segments label={t.restoreMode} value={s!.restoreOnAnyFolder ? 'any' : 'normal'}
                    options={[{ value: 'normal', label: t.normalLaunch }, { value: 'any', label: t.anyFolder }]}
                    disabled={!s!.restoreTabs || pending.has('restoreOnAnyFolder')} onChange={value => set('restoreOnAnyFolder', value === 'any')} />
                </div>
              </Setting>
            </div>
          </div>
          <details className="help-details"><summary><span>{t.recoveryDetails}</span><ChevronRight aria-hidden="true" /></summary><p>{t.recoveryNotes}</p></details>
        </section>

        <section hidden={section !== 'preferences'} aria-label={t.preferences}>
          <div className="settings-group">
            <h2>{t.appearance}</h2>
            <div className="settings-card">
              <div className="setting appearance-row">
                <span className="setting-title">{t.themeLabel}</span>
                <Segments label={t.themeLabel} value={s!.theme} options={[{value: 'Light', label: t.lightChoice}, {value: 'Dark', label: t.darkChoice}]}
                  disabled={pending.has('theme')} onChange={value => set('theme', value)} />
              </div>
              <div className="setting appearance-row">
                <span className="setting-title">{t.languageLabel}</span>
                <Segments label={t.languageLabel} value={language} options={[{value: 'zh-CN', label: '简体中文'}, {value: 'en-US', label: 'English'}]}
                  disabled={pending.has('language')} onChange={value => set('language', value)} />
              </div>
            </div>
          </div>
          <div className="settings-group">
            <h2>{t.background}</h2>
            <div className="settings-card">
              <Setting id="startup" title={t.startup} description={t.startupDesc} control={toggle('startup', t.startup, state.startup)} />
              <Setting id="showTrayIcon" title={t.tray} description={s!.showTrayIcon ? t.trayDesc : t.trayHidden} control={toggle('showTrayIcon', t.tray, s!.showTrayIcon)} />
              <Setting id="autoUpdate" title={t.autoUpdate} description={t.autoUpdateDesc} control={toggle('autoUpdate', t.autoUpdate, s!.autoUpdate)} />
            </div>
          </div>
          <p className="context-note"><Info aria-hidden="true" /><span>{t.running}</span></p>
        </section>

        <section hidden={section !== 'about'} aria-label={t.about}>
          <div className="settings-card product-identity">
            <img src={logo} alt="" />
            <div className="product-identity-copy">
              <div className="product-title-row">
                <h2>WinTab</h2>
                <span className="version-badge">v{state.version}</span>
              </div>
              <p>{t.aboutDescription}</p>
            </div>
          </div>
          <div className="settings-group">
            <dl className="settings-card product-facts">
              <div><dt>{t.version}</dt><dd>v{state.version}</dd></div>
              <div><dt>{t.licenseLabel}</dt><dd>{t.license}</dd></div>
              <div><dt>{t.repository}</dt><dd><a href="https://github.com/OUBIGFA/WinTab" title={t.openSource} onClick={event => {
                event.preventDefault(); void desktop.openProject().catch(() => setError(true))
              }}><span>{t.project}</span><ArrowUpRight aria-hidden="true" /></a></dd></div>
            </dl>
          </div>
          <div className="settings-group">
            <h2>{t.support}</h2>
            <div className="settings-card">
              <div className="setting support-row">
                <div className="setting-copy">
                  <h3 className="setting-title">{t.updates}</h3>
                  <p className="description">{t.updatesDesc}</p>
                  {state.updateFeedback && <p className="inline-feedback" role="status"><Check aria-hidden="true" /><span>{state.updateFeedback}</span></p>}
                </div>
                <button type="button" className="button button-outline" disabled={updating} onClick={() => { void run('update', desktop.update) }}>
                  {updating ? <LoaderCircle className="spinning" aria-hidden="true" /> : <RefreshCw aria-hidden="true" />}<span>{updating ? t.checking : t.updates}</span>
                </button>
              </div>
              <div className="setting support-row">
                <div className="setting-copy">
                  <h3 className="setting-title">{t.logs}</h3>
                  <p className="description">{t.logsDesc}</p>
                </div>
                <button type="button" className="button button-outline" disabled={pending.has('logs')} onClick={() => { void run('logs', desktop.logs) }}>
                  <FolderOpen aria-hidden="true" /><span>{t.logs}</span>
                </button>
              </div>
            </div>
          </div>
        </section>

        <footer className="page-footer">{section !== 'about' && !dirty && !error && !state.shortcutError && !state.storageError && <>
          {pending.size ? <LoaderCircle className="spinning" aria-hidden="true" /> : <Check aria-hidden="true" />}
          <span>{pending.size ? t.saving : t.saved}</span>
        </>}</footer>
      </div>}
    </main>
  </div>
}
