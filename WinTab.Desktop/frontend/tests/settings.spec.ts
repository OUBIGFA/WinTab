import { expect, test, type Page } from '@playwright/test'
import AxeBuilder from '@axe-core/playwright'
import fixture from './fixtures/state.json' with { type: 'json' }
import { copies } from '../src/i18n'
import type { State } from '../src/desktop'

async function installBridge(page: Page, overrides: Partial<State['settings']> = {}, maxHeight = 1100) {
  await page.addInitScript(({ initial, overrides, maxHeight }) => {
    const state = structuredClone(initial)
    Object.assign(state.settings, overrides)
    const listeners = new Map<string, ((value: unknown) => void)[]>()
    const host = window as unknown as Record<string, any>
    host.__calls = []
    host.__state = state
    host.__emit = (patch: Record<string, unknown>, revision = state.revision + 1) => {
      Object.assign(state.settings, patch)
      state.revision = revision
      for (const listener of listeners.get('state-changed') ?? []) listener(structuredClone(state))
    }
    host.mygo = {
      on: (name: string, listener: (value: unknown) => void) => {
        listeners.set(name, [...(listeners.get(name) ?? []), listener])
        return () => listeners.set(name, (listeners.get(name) ?? []).filter(item => item !== listener))
      },
      call: async (method: string, ...args: any[]) => {
        host.__calls.push({ method, args })
        if (host.__reject === method) throw new Error('simulated pipe error')
        if (method === 'Desktop.Load') return { state: structuredClone(state), window: { width: innerWidth, maxWidth: 1920, maxHeight, hasSavedSize: !!state.settings.formSize } }
        if (method === 'Desktop.Ready') { host.__ready = { width: args[0], height: args[1] }; return }
        if (method === 'Desktop.OpenProject') return
        if (method === 'Desktop.Set') {
          if (args[0] === 'startup') state.startup = args[1]
          else Object.assign(state.settings, { [args[0]]: args[1] })
          if (args[0] === 'windowHook' && !args[1]) state.settings.reuseTabs = false
          if (args[0] === 'reuseTabs' && args[1]) state.settings.windowHook = true
        }
        if (method === 'Desktop.Shortcuts') {
          if (args[1] === args[3]) state.shortcutError = '两个恢复功能不能使用相同快捷键'
          else {
            state.shortcutError = null
            Object.assign(state.settings, { restoreGroupShortcutEnabled: args[0], restoreGroupShortcut: args[1], reopenTabShortcutEnabled: args[2], reopenTabShortcut: args[3] })
          }
        }
        state.recordClosedTabs = state.settings.reopenClosedTab || state.settings.reopenTabShortcutEnabled
        if (method === 'Desktop.Restore') state.sessionFeedback = '恢复完成'
        if (method === 'Desktop.Update') state.updateFeedback = '已是最新版本'
        state.revision++
        return structuredClone(state)
      },
    }
  }, { initial: fixture as State, overrides, maxHeight })
}

async function open(page: Page, overrides: Partial<State['settings']> = {}, maxHeight = 1100) {
  await installBridge(page, overrides, maxHeight)
  await page.goto('/')
  await expect(page.getByRole('switch', { name: copies[overrides.language ?? 'zh-CN'].merge, exact: true })).toBeVisible()
  await page.waitForFunction(() => !!(window as any).__ready)
}

async function navigate(page: Page, name: string) {
  await page.getByRole('navigation').getByRole('button', { name, exact: true }).click()
  await expect(page.getByRole('navigation').getByRole('button', { name, exact: true })).toHaveAttribute('aria-current', 'page')
}

async function fitNativeViewport(page: Page) {
  const size = await page.evaluate(() => (window as any).__ready as { width: number; height: number })
  await page.setViewportSize(size)
}

test('section navigation keeps readable controls and preserves shortcut drafts', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 650 })
  await open(page, { formSize: { width: 900, height: 650 } })
  const navigation = page.getByRole('navigation', { name: '设置分区' })
  await expect(navigation).toBeVisible()
  await navigation.getByRole('button', { name: '会话恢复', exact: true }).click()
  const shortcut = page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true })
  await shortcut.fill('Ctrl+Shift+K')
  await navigation.getByRole('button', { name: '外观与系统', exact: true }).click()
  await expect(shortcut).not.toBeVisible()
  await navigation.getByRole('button', { name: '会话恢复', exact: true }).click()
  await expect(shortcut).toHaveValue('Ctrl+Shift+K')
  await expect(page.getByText('有未应用的快捷键', { exact: true })).toBeVisible()
  await page.setViewportSize({ width: 640, height: 480 })
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
  expect(await shortcut.evaluate(element => parseFloat(getComputedStyle(element).fontSize))).toBeGreaterThanOrEqual(14)
})

for (const language of ['zh-CN', 'en-US'] as const) {
  for (const theme of ['Light', 'Dark'] as const) {
    test(`all sections stay usable in ${language} ${theme}`, async ({ page }, testInfo) => {
      await page.setViewportSize({ width: 1064, height: 780 })
      await open(page, { language, theme })
      await fitNativeViewport(page)
      const t = copies[language]
      for (const [id, label] of [['explorer', t.explorer], ['recovery', t.recoveryNav], ['preferences', t.preferences], ['about', t.about]]) {
        await navigate(page, label)
        await expect(page.getByRole('heading', { name: label, exact: true, level: 1 })).toBeVisible()
        expect(await page.evaluate(() => document.documentElement.scrollHeight <= innerHeight + 1)).toBe(true)
        expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
        const accessibility = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21aa']).analyze()
        expect(accessibility.violations, `${language} ${theme} ${id}`).toEqual([])
        await page.screenshot({ path: testInfo.outputPath(`${id}.png`), animations: 'disabled' })
      }
      await expect(page.getByRole('button', { name: t.updates, exact: true })).toBeInViewport()
    })
  }
}

test('short screen scrolls content without shrinking text or navigation', async ({ page }) => {
  await open(page, {}, 740)
  await fitNativeViewport(page)
  const dimensions = await page.evaluate(() => ({
    height: document.documentElement.scrollHeight, viewport: innerHeight, width: innerWidth,
    page: document.querySelector('.page')!.getBoundingClientRect().toJSON(),
    zoom: getComputedStyle(document.querySelector('.page')!).zoom,
  }))
  expect(dimensions.height, JSON.stringify(dimensions)).toBeLessThanOrEqual(dimensions.viewport + 1)
  expect(dimensions.zoom).toBe('1')
  await page.getByRole('radio', { name: '高', exact: true }).scrollIntoViewIfNeeded()
  await expect(page.getByRole('navigation')).toBeInViewport()
  await navigate(page, '关于')
  await expect(page.getByRole('button', { name: '检查更新', exact: true })).toBeVisible()
})

test('saved small size stays manual with all content reachable by scrolling', async ({ page }) => {
  await page.setViewportSize({ width: 760, height: 620 })
  await open(page, { formSize: { width: 760, height: 620 } })
  expect(await page.locator('.page').evaluate(element => getComputedStyle(element).zoom)).toBe('1')
  expect(await page.locator('.workspace').evaluate(element => element.scrollHeight > element.clientHeight)).toBe(true)
  await navigate(page, '关于')
  await expect(page.getByRole('button', { name: '检查更新', exact: true })).toBeInViewport()
})

test('settings reflect dependent backend state, not optimistic local toggles', async ({ page }) => {
  await open(page)
  await page.getByRole('switch', { name: '合并新窗口', exact: true }).click()
  await expect(page.getByRole('switch', { name: '合并新窗口', exact: true })).not.toBeChecked()
  await expect(page.getByRole('switch', { name: '复用已有标签', exact: true })).not.toBeChecked()
  await page.getByRole('switch', { name: '复用已有标签', exact: true }).click()
  await expect(page.getByRole('switch', { name: '合并新窗口', exact: true })).toBeChecked()
  await page.evaluate(() => { (window as any).__reject = 'Desktop.Set' })
  await page.getByRole('switch', { name: '合并新窗口', exact: true }).click()
  await expect(page.getByRole('alert')).toContainText('操作未完成')
  await expect(page.getByRole('switch', { name: '合并新窗口', exact: true })).toBeChecked()
})

test('physical capture, manual text and drafts survive unrelated updates', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  const group = page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true })
  await group.focus()
  await page.keyboard.press('Control+Shift+K')
  await expect(group).toHaveValue('Ctrl+Shift+K')
  await page.getByRole('button', { name: '切换为深色主题' }).click()
  await expect(group).toHaveValue('Ctrl+Shift+K')
  await navigate(page, '外观与系统')
  await page.getByRole('switch', { name: '显示托盘图标', exact: true }).click()
  await navigate(page, '会话恢复')
  await expect(group).toHaveValue('Ctrl+Shift+K')
  await group.fill('Alt+F12')
  await page.getByRole('button', { name: '应用快捷键', exact: true }).click()
  await expect(page.getByText('快捷键已应用', { exact: true })).toBeVisible()
  await expect(group).toHaveValue('Alt+F12')
  const calls = await page.evaluate(() => (window as any).__calls)
  expect(calls.some((call: any) => call.method === 'Desktop.Shortcuts' && call.args[1] === 'Alt+F12')).toBe(true)
})

test('repeat ownership, top-row digits, function keys and normal Tab navigation', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  const field = page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true })
  await field.focus()
  await page.keyboard.down('Control'); await page.keyboard.down('k'); await page.keyboard.up('Control')
  await page.keyboard.down('k'); await page.keyboard.up('k')
  await expect(field).toHaveValue('Ctrl+K')
  await page.keyboard.press('Alt+Digit7')
  await expect(field).toHaveValue('Alt+7')
  await page.keyboard.press('Control+F12')
  await expect(field).toHaveValue('Ctrl+F12')
  await page.keyboard.press('Tab')
  await expect(field).not.toBeFocused()
})

test('deferred shortcut selection preserves focus after Tab navigation', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  const field = page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true })
  await field.focus()
  // Replay a slow frame after keyboard navigation, keeping the input and browser focus behavior real
  await page.evaluate(() => {
    const host = window as typeof window & { releaseShortcutFrame?: () => void }
    const scheduleFrame = window.requestAnimationFrame.bind(window)
    window.requestAnimationFrame = callback => {
      window.requestAnimationFrame = scheduleFrame
      host.releaseShortcutFrame = () => callback(performance.now())
      return 0
    }
  })
  await page.keyboard.press('Control+K')
  await page.keyboard.press('Tab')
  await expect(field).not.toBeFocused()
  await page.evaluate(() => {
    const host = window as typeof window & { releaseShortcutFrame?: () => void }
    host.releaseShortcutFrame!()
  })
  await expect(field).not.toBeFocused()
  await expect(field).toHaveValue('Ctrl+K')
})

test('manual recovery, shortcut prerequisites and optional automatic restore stay independent', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  await expect(page.getByRole('switch', { name: '自动恢复上次标签组', exact: true })).not.toBeChecked()
  await expect(page.getByRole('radio', { name: '仅普通启动', exact: true })).toBeDisabled()
  await expect(page.getByRole('button', { name: '立即恢复', exact: true }).first()).toBeEnabled()
  await expect(page.getByRole('switch', { name: '记录最近关闭的标签', exact: true })).toBeChecked()
  await expect(page.getByRole('switch', { name: '记录最近关闭的标签', exact: true })).toBeDisabled()
  await page.getByRole('checkbox', { name: '启用单标签快捷键', exact: true }).click()
  await expect(page.getByRole('switch', { name: '记录最近关闭的标签', exact: true })).toBeEnabled()
  await page.getByRole('switch', { name: '记录最近关闭的标签', exact: true }).click()
  await expect(page.getByRole('button', { name: '立即恢复', exact: true }).nth(1)).toBeDisabled()
  await page.getByRole('button', { name: '立即恢复', exact: true }).first().click()
  await expect(page.getByText('恢复完成', { exact: true })).toBeVisible()
})

test('Radix options support keyboard navigation and bilingual theme changes', async ({ page }) => {
  await open(page)
  await page.getByRole('radio', { name: '中', exact: true }).focus()
  await page.keyboard.press('ArrowRight')
  await expect(page.getByRole('radio', { name: '高', exact: true })).toBeChecked()
  await page.getByRole('button', { name: 'Switch to English' }).click()
  await expect(page.getByRole('heading', { name: 'File Explorer', exact: true })).toBeVisible()
  await expect(page.getByText('Preferences', { exact: true })).toHaveCount(0)
  await page.getByRole('button', { name: 'Switch to dark theme' }).click()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'Dark')
  await expect(page.getByRole('radio', { name: 'High', exact: true })).toBeChecked()
})

test('empty-space navigation is opt-in with independent persistent action tabs', async ({ page }) => {
  await open(page)
  const toggle = page.getByRole('switch', { name: '双击空白处导航', exact: true })
  const parent = page.getByRole('radio', { name: '返回上一层', exact: true })
  const back = page.getByRole('radio', { name: '后退', exact: true })
  await expect(toggle).not.toBeChecked()
  await expect(parent).toBeChecked()
  await expect(back).toBeDisabled()
  await expect(page.getByText('偏好设置', { exact: true })).toHaveCount(0)
  await toggle.click()
  await back.click()
  await expect(back).toBeChecked()
  await toggle.click()
  await expect(back).toBeChecked()
  await expect(back).toBeDisabled()
  await toggle.click()
  await expect(back).toBeEnabled()
  await back.focus()
  await page.keyboard.press('ArrowLeft')
  await expect(parent).toBeChecked()
  await expect(page.getByRole('switch', { name: '双击关闭标签', exact: true })).toBeChecked()
  await expect(page.getByRole('switch', { name: '中键打开并切换', exact: true })).toBeChecked()
  const calls = await page.evaluate(() => (window as any).__calls)
  expect(calls.some((call: any) => call.method === 'Desktop.Set' && call.args[0] === 'doubleClickEmptyAreaAction' && call.args[1] === 'Back')).toBe(true)
  await page.getByRole('button', { name: 'Switch to English' }).click()
  await expect(page.getByRole('radio', { name: 'Up one level', exact: true })).toBeChecked()
  await expect(page.getByRole('radio', { name: 'Back', exact: true })).toBeEnabled()
})

test('both languages have consistent punctuation-free UI copy', () => {
  for (const copy of Object.values(copies)) {
    for (const value of Object.values(copy)) expect(value).not.toMatch(/[。.…]$/)
  }
})

test('duplicate shortcuts show validation while retaining the draft', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  const field = page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true })
  await field.fill('Alt+W')
  await page.getByRole('button', { name: '应用快捷键', exact: true }).click()
  await expect(page.getByRole('alert')).toContainText('不能使用相同快捷键')
  await expect(field).toHaveValue('Alt+W')
  await expect(field).toHaveAttribute('aria-invalid', 'true')
})

test('save guidance does not contradict a draft or failed request', async ({ page }) => {
  await open(page)
  await navigate(page, '会话恢复')
  await page.getByRole('textbox', { name: '标签组恢复快捷键', exact: true }).fill('Alt+K')
  await expect(page.locator('.page-footer')).not.toContainText(copies['zh-CN'].saved)
  await page.getByRole('button', { name: '应用快捷键', exact: true }).click()
  await expect(page.locator('.page-footer')).toContainText(copies['zh-CN'].saved)
  await page.evaluate(() => { (window as any).__reject = 'Desktop.Set' })
  await navigate(page, '资源管理器')
  await page.getByRole('switch', { name: '合并新窗口', exact: true }).click()
  await expect(page.getByRole('alert')).toContainText(copies['zh-CN'].requestError)
  await expect(page.locator('.page-footer')).not.toContainText(copies['zh-CN'].saved)
  await expect(page.getByRole('switch', { name: '合并新窗口', exact: true })).toBeChecked()
})

test('failed bootstrap reveals a retry surface and reconnects to real state', async ({ page }) => {
  await installBridge(page)
  await page.addInitScript(() => { (window as any).__reject = 'Desktop.Load' })
  await page.goto('/')
  await expect(page.getByRole('heading', { name: 'Connection unavailable', exact: true })).toBeVisible()
  await expect(page.getByRole('button', { name: 'Reconnect', exact: true })).toBeVisible()
  await page.waitForFunction(() => !!(window as any).__ready)
  await page.evaluate(() => { (window as any).__reject = null })
  await page.getByRole('button', { name: 'Reconnect', exact: true }).click()
  await expect(page.getByRole('switch', { name: '合并新窗口', exact: true })).toBeChecked()
  await expect(page.getByRole('heading', { name: '资源管理器', exact: true })).toBeVisible()
})

test('appearance, support and navigation work by keyboard without changing unrelated settings', async ({ page }) => {
  await open(page)
  const nav = page.getByRole('navigation')
  await nav.getByRole('button', { name: '外观与系统', exact: true }).focus()
  await page.keyboard.press('Enter')
  await expect(page.getByRole('heading', { level: 1 })).toBeFocused()
  await page.getByRole('radio', { name: '深色', exact: true }).click()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'Dark')
  await page.getByRole('radio', { name: 'English', exact: true }).click()
  await expect(page.locator('html')).toHaveAttribute('lang', 'en-US')
  await expect(nav.getByRole('button', { name: 'Appearance & system', exact: true })).toHaveAttribute('aria-current', 'page')
  await navigate(page, 'About')
  await page.getByRole('button', { name: 'Check for updates', exact: true }).click()
  await expect(page.getByText('已是最新版本', { exact: true })).toBeVisible()
  await page.getByRole('button', { name: 'Open logs', exact: true }).click()
  await page.getByRole('link', { name: 'View source code', exact: true }).click()
  const calls = await page.evaluate(() => (window as any).__calls)
  expect(calls.filter((call: any) => call.method === 'Desktop.Set').map((call: any) => call.args[0])).toEqual(['theme', 'language'])
  expect(calls.map((call: any) => call.method)).toEqual(expect.arrayContaining(['Desktop.Update', 'Desktop.Logs', 'Desktop.OpenProject']))
})

for (const width of [390, 640, 900]) {
  test(`all sections remain reachable at ${width}px without horizontal overflow`, async ({ page }, testInfo) => {
    await page.setViewportSize({ width, height: 620 })
    await open(page, { language: 'en-US', theme: 'Dark', formSize: { width, height: 620 } })
    for (const label of ['File Explorer', 'Session recovery', 'Appearance & system', 'About']) {
      await navigate(page, label)
      const overflow = await page.locator('.workspace').evaluate(element => element.scrollWidth > element.clientWidth)
      expect(overflow, label).toBe(false)
      await expect(page.getByRole('heading', { level: 1 })).toBeInViewport()
      await page.locator('.page-footer').scrollIntoViewIfNeeded()
      await expect(page.locator('.page-footer')).toBeInViewport()
    }
    await page.screenshot({ path: testInfo.outputPath('narrow.png'), animations: 'disabled' })
  })
}
