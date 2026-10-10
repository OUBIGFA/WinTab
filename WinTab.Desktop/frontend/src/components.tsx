import { useRef, type ReactNode, type KeyboardEvent } from 'react'
import * as SwitchPrimitive from '@radix-ui/react-switch'
import * as CheckboxPrimitive from '@radix-ui/react-checkbox'
import * as RadioGroup from '@radix-ui/react-radio-group'
import { Check, Keyboard, type LucideIcon } from 'lucide-react'

export function Toggle({ id, label, description, checked, disabled, onChange }: {
  id: string; label: string; description?: string; checked: boolean; disabled?: boolean; onChange: (value: boolean) => void
}) {
  return <SwitchPrimitive.Root id={id} className="switch" checked={checked} disabled={disabled}
    onCheckedChange={onChange} aria-label={label} aria-describedby={description}>
    <SwitchPrimitive.Thumb className="switch-thumb" />
  </SwitchPrimitive.Root>
}

export function Checkbox({ id, label, children, checked, disabled, onChange }: {
  id: string; label: string; children: ReactNode; checked: boolean; disabled?: boolean; onChange: (value: boolean) => void
}) {
  return <div className="checkbox-label">
    <CheckboxPrimitive.Root id={id} className="checkbox" checked={checked} disabled={disabled}
      aria-label={label} onCheckedChange={value => onChange(value === true)}>
      <CheckboxPrimitive.Indicator><Check size={12} strokeWidth={2.5} /></CheckboxPrimitive.Indicator>
    </CheckboxPrimitive.Root>
    <label htmlFor={id}>{children}</label>
  </div>
}

export function Segments<T extends string>({ label, value, options, onChange, disabled }: {
  label: string; value: T; options: { value: T; label: string }[]; onChange: (value: T) => void; disabled?: boolean
}) {
  return <RadioGroup.Root className="segments" value={value} disabled={disabled} aria-label={label}
    orientation="horizontal" onValueChange={value => onChange(value as T)}
    onKeyDownCapture={event => {
      const direction = event.key === 'ArrowRight' || event.key === 'ArrowDown' ? 1
        : event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 0
      if (!direction || disabled || event.altKey || event.ctrlKey || event.metaKey) return
      event.preventDefault()
      // Radix moves arrow focus on a timer. A quick key release can arrive before that timer
      // and leave the focused option unchecked, so commit this choice in the keydown itself.
      const items = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>('[role="radio"]'))
      const current = items.indexOf(document.activeElement as HTMLButtonElement)
      if (current < 0) return
      const next = (current + direction + items.length) % items.length
      items[next].focus()
      onChange(options[next].value)
    }}>
    {options.map(item => <RadioGroup.Item className="segment" key={item.value} value={item.value}>{item.label}</RadioGroup.Item>)}
  </RadioGroup.Root>
}

export function SectionHeading({ icon: Icon, title, id }: { icon: LucideIcon; title: string; id: string }) {
  return <div className="section-heading">
    <Icon className="section-icon" size={19} strokeWidth={1.7} aria-hidden="true" />
    <h2 id={id}>{title}</h2>
  </div>
}

export function Setting({ id, title, description, control, children, compact = false }: {
  id: string; title: string; description?: string; control: ReactNode; children?: ReactNode; compact?: boolean
}) {
  return <div className={`setting${compact ? ' setting-compact' : ''}`}>
    <div className="setting-main">
      <div className="setting-copy">
        <label htmlFor={id} className="setting-title">{title}</label>
        {description && <p id={`${id}-description`} className="description">{description}</p>}
      </div>
      {control}
    </div>
    {children && <div className="setting-extra">{children}</div>}
  </div>
}

export function chordFromEvent(event: Pick<KeyboardEvent, 'code' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey' | 'getModifierState'>): string | null {
  if (event.metaKey || event.getModifierState('AltGraph') || (!event.ctrlKey && !event.altKey)) return null
  let key: string
  if (/^Key[A-Z]$/.test(event.code)) key = event.code.slice(3)
  else if (/^Digit[0-9]$/.test(event.code)) key = event.code.slice(5)
  else if (/^F([1-9]|1[0-9]|2[0-4])$/.test(event.code)) key = event.code
  else return null
  return [...(event.ctrlKey ? ['Ctrl'] : []), ...(event.altKey ? ['Alt'] : []), ...(event.shiftKey ? ['Shift'] : []), key].join('+')
}

// Physical codes avoid keyboard-layout/IME ambiguity. Unmodified typing and context-menu paste
// stay available; a captured chord owns repeats until every key has been released, as in WPF.
export function ShortcutInput({ id, label, value, onChange, invalid, hint }: {
  id: string; label: string; value: string; onChange: (value: string) => void; invalid: boolean; hint: string
}) {
  const input = useRef<HTMLInputElement>(null)
  const captured = useRef<{ code: string; released: boolean } | null>(null)
  return <div className="shortcut-input">
    <Keyboard size={15} aria-hidden="true" />
    <input ref={input} id={id} aria-label={label} aria-invalid={invalid || undefined} title={hint} aria-describedby="shortcut-help"
      value={value} spellCheck={false} autoComplete="off" autoCapitalize="off" maxLength={128}
      onChange={event => onChange(event.target.value)} onBlur={() => { captured.current = null }}
      onKeyDown={event => {
        if (event.code === 'Tab' && !event.ctrlKey && !event.altKey && !event.metaKey) { captured.current = null; return }
        if (captured.current) {
          event.preventDefault()
          if (event.code === captured.current.code) captured.current.released = false
          return
        }
        const chord = chordFromEvent(event)
        if (!chord) return
        event.preventDefault()
        captured.current = { code: event.code, released: false }
        if (!event.repeat) onChange(chord)
        requestAnimationFrame(() => {
          // Selecting also focuses the input, so a delayed frame must respect later Tab/click navigation
          if (document.activeElement === input.current) input.current?.select()
        })
      }}
      onKeyUp={event => {
        if (!captured.current) return
        event.preventDefault()
        if (event.code === captured.current.code) captured.current.released = true
        if (captured.current.released && !event.ctrlKey && !event.altKey && !event.shiftKey && !event.metaKey) captured.current = null
      }} />
  </div>
}
