/**
 * NEWDOC-DRAG-1 — everything interactive in the Electron toolbar must opt out of
 * the OS drag region.
 *
 * THE BUG THIS EXISTS FOR. `.app--electron .toolbar--electron` sets
 * `-webkit-app-region: drag` so the toolbar can move the frameless window.
 * `-webkit-app-region` is an INHERITED property, so that reaches every descendant
 * unless it opts back out. The dropdown trigger buttons carry `.toolbar__nodrag`;
 * the `<ul class="dropdown__menu">` each one opens did not. Windows therefore
 * treated a click anywhere on an open menu as a title-bar drag, the event never
 * reached the DOM, `onClick` never fired, and the menu just sat there. Reported
 * from the real app as "the New document buttons simply don't work" — and it
 * affected all four toolbar menus, not only New.
 *
 * WHY THIS IS A CSS TEST AND NOT A CLICK TEST. Nothing that drives the renderer
 * can catch it. Playwright dispatches input over CDP directly into the renderer,
 * which bypasses the OS drag-region hit-testing completely, so
 * `tests/e2e/new-document.spec.ts` clicks these very items successfully on a build
 * where a real mouse cannot. jsdom has no concept of app regions at all. The only
 * honest automated guard is to assert the rule exists — the behaviour itself needs
 * a human with a real mouse, which is worth remembering for anything else added to
 * this toolbar.
 *
 * There IS a way to check the behaviour, just not from a test runner:
 * `scripts/verify-drag-regions.mjs` synthesises a genuine OS mouse event. It was
 * used to prove this fix both ways — with the rule removed and the app rebuilt it
 * reports the menu still open and no dialog, which is the bug exactly as reported.
 */
import fs from 'node:fs'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

const CSS = fs.readFileSync(path.resolve(process.cwd(), 'src/index.css'), 'utf8')

/** The `-webkit-app-region` value a selector resolves to, or null if it sets none. */
function appRegionFor(selector: string): string | null {
  // Deliberately literal: the point is that a rule for this exact selector exists,
  // not that some equivalent one might.
  const pattern = new RegExp(
    `${selector.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}\\s*\\{[^}]*?-webkit-app-region:\\s*([a-z-]+)`,
    's',
  )
  return CSS.match(pattern)?.[1] ?? null
}

describe('Electron toolbar drag regions (NEWDOC-DRAG-1)', () => {
  it('makes the toolbar itself draggable, which is the whole reason for the hazard', () => {
    expect(appRegionFor('.app--electron .toolbar--electron')).toBe('drag')
  })

  it('opts dropdown PANELS out of the drag region, not just their trigger buttons', () => {
    // Without this, every item in the New/Export/Theme/Language menus is dead to
    // a real mouse while looking and testing perfectly fine.
    expect(appRegionFor('.app--electron .dropdown__menu')).toBe('no-drag')
  })

  it('still opts the trigger buttons out', () => {
    expect(appRegionFor('.app--electron .toolbar__nodrag')).toBe('no-drag')
  })

  it('every toolbar dropdown uses the class that rule targets', () => {
    // The rule is only as good as its coverage: a new menu that styles itself
    // differently would silently reintroduce the bug.
    const components = ['ExportMenu', 'NewDocumentMenu', 'ThemeMenu', 'LanguageMenu']
    for (const name of components) {
      const source = fs.readFileSync(path.resolve(process.cwd(), `src/components/${name}.tsx`), 'utf8')
      expect(source, `${name} should render its panel with the dropdown__menu class`).toContain('dropdown__menu')
    }
  })
})
