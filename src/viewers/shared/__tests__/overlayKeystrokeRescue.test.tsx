/**
 * MATRIX-FLAKE-2 — keystrokes that land in the wrong place while a freshly
 * mounted cell editor is still getting focus.
 *
 * The chain this covers, which took four full e2e runs to pin down:
 *   1. Typing into a cell opens an overlay. Opening it makes the grid
 *      re-render, and that re-render can leave DOM focus on a `<td>` of
 *      glide-data-grid's own accessibility table instead of the new textarea.
 *   2. `OVERLAY_FOCUS_GUARD_MS`'s guard notices and moves focus to the overlay
 *      — but it cannot un-deliver the keys already dispatched to that `<td>`.
 *      Those characters are gone.
 *   3. So the overlay ends up focused and holding only its seed character, the
 *      Enter is lost too, the edit never commits, and a Ctrl+S queued behind it
 *      saves nothing.
 *
 * The reported failure was exactly that: `overlayValue="a"` after `"after"` was
 * typed, with `focus=TEXTAREA` by the time it was reported — the guard had done
 * its job, too late for the keys.
 *
 * `KeyHold` already covers the window BEFORE the overlay mounts. These tests
 * cover the window after it mounts and before it has focus, which nothing did.
 */
import { act, render } from '@testing-library/react'
import { useState } from 'react'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { SpreadsheetDataEditor } from '../SpreadsheetDataEditor'

/**
 * Stands in for glide-data-grid, rendering the real `TextCellEditor` the way
 * glide does for a type-triggered open.
 *
 * The editor is reached through `provideEditor`, which is internal to the
 * module — so the mock captures it from the props the component passes down and
 * renders it, which is what glide itself does.
 */
type CapturedEditor = {
  render: () => React.ReactElement | null
  finished: Array<{ data: string; movement: readonly [number, number] }>
  changes: string[]
}

const captured: CapturedEditor = { render: () => null, finished: [], changes: [] }

vi.mock('@glideapps/glide-data-grid', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@glideapps/glide-data-grid')>()
  return {
    ...actual,
    DataEditor: (props: {
      provideEditor?: (cell: unknown) => React.ComponentType<Record<string, unknown>> | undefined
    }) => {
      const Editor = props.provideEditor?.({ kind: actual.GridCellKind.Text })
      captured.render = () => (Editor ? <EditorHost Editor={Editor} /> : null)
      return null
    },
  }
})

/**
 * Stands in for glide-data-grid's own ownership of the cell value.
 *
 * STATEFUL on purpose. The editor renders a React-CONTROLLED textarea, so its
 * DOM value is whatever the `value` prop last said. A harness that passed a
 * fixed literal would re-render with the old text and snap the textarea back,
 * which is exactly what happened on the first attempt here and looked like the
 * rescue failing — it was the double being wrong, not the code. glide keeps the
 * value and feeds it back through `onChange`; so does this.
 */
function EditorHost({ Editor }: { readonly Editor: React.ComponentType<Record<string, unknown>> }) {
  const [data, setData] = useState('a')
  return (
    <Editor
      value={{ kind: 'text', data, displayData: data, allowOverlay: true }}
      initialValue="a"
      forceEditMode
      isHighlighted={false}
      onChange={(next: { data: string }) => {
        captured.changes.push(next.data)
        setData(next.data)
      }}
      onFinishedEditing={(next: { data: string } | undefined, movement: readonly [number, number]) => {
        captured.finished.push({ data: next?.data ?? '', movement })
      }}
    />
  )
}

/** The textarea the overlay renders, found the way the production code finds it. */
function overlay(): HTMLTextAreaElement {
  const el = document.querySelector('#portal textarea') ?? document.querySelector('textarea')
  if (!(el instanceof HTMLTextAreaElement)) throw new Error('no overlay textarea rendered')
  return el
}

/** Puts focus where the bug puts it: a `<td>` of the accessibility table. */
function focusAccessibilityCell(): HTMLTableCellElement {
  const table = document.createElement('table')
  const row = table.insertRow()
  const cell = row.insertCell()
  cell.tabIndex = 0
  document.body.appendChild(table)
  cell.focus()
  return cell
}

/**
 * The other places focus lands under CPU throttling, which the first version of
 * this fix did not cover.
 *
 * It listed the wrong places to be — nothing, `<body>`, a table cell — and the
 * e2e flake survived it: at 8x throttling focus also ends up on the grid's own
 * `<canvas>` and on its scroller `<div>`. The condition is now stated the other
 * way round (is this a text target the user chose?), so these are covered by
 * construction rather than by having been thought of.
 */
function focusElement(tag: 'canvas' | 'div'): HTMLElement {
  const el = document.createElement(tag)
  el.tabIndex = 0
  document.body.appendChild(el)
  el.focus()
  return el
}

function dispatchKey(key: string, options: KeyboardEventInit = {}): void {
  // `act` so the state update the rescue triggers through React's own
  // `onChange` is flushed before the assertion reads the DOM.
  act(() => {
    document.activeElement?.dispatchEvent(
      new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...options }),
    )
  })
}

beforeEach(() => {
  captured.finished = []
  captured.changes = []
  // The guard polls with rAF; jsdom has one, but a deterministic clock makes
  // "before focus lands" an explicit state rather than a race.
  vi.stubGlobal(
    'requestAnimationFrame',
    (() => 1) as unknown as typeof requestAnimationFrame,
  )
  vi.stubGlobal('cancelAnimationFrame', (() => {}) as unknown as typeof cancelAnimationFrame)
})

afterEach(() => {
  vi.unstubAllGlobals()
  document.body.innerHTML = ''
})

describe('the window before the editor component\'s effect runs (MATRIX-FLAKE-2)', () => {
  /**
   * glide-data-grid creates the overlay DOM during render; React runs effects
   * after paint. So there is always at least a frame where the textarea exists
   * and anything installed from the EDITOR's own effect does not — which is
   * where two earlier versions of this fix let keys through.
   *
   * Reproduced here by mounting only the long-lived parent and putting a
   * textarea into `#portal` by hand: the editor component never mounts, so
   * nothing it would have installed exists.
   */
  function overlayWithoutEditor(initial: string): HTMLTextAreaElement {
    render(<SpreadsheetDataEditor rows={1} columns={[]} getCellContent={() => ({}) as never} />)
    const portal = document.createElement('div')
    portal.id = 'portal'
    const textarea = document.createElement('textarea')
    textarea.value = initial
    portal.appendChild(textarea)
    document.body.appendChild(portal)
    return textarea
  }

  it('rescues a key although the editor component has not mounted', () => {
    const textarea = overlayWithoutEditor('t')
    focusElement('canvas')

    dispatchKey('w')

    // Without the parent-level listener this character is lost: nothing is
    // listening yet, which is the structural gap the move fixes.
    expect(textarea.value).toBe('tw')
  })

  it('rescues a whole word typed in that window', () => {
    const textarea = overlayWithoutEditor('t')
    focusElement('canvas')

    for (const key of ['w', 'o']) dispatchKey(key)

    expect(textarea.value).toBe('two')
  })

  it('still leaves another focused control alone in that window', () => {
    const textarea = overlayWithoutEditor('t')
    const other = document.createElement('input')
    document.body.appendChild(other)
    other.focus()

    dispatchKey('w')

    expect(textarea.value).toBe('t')
  })

  it('is inert when no overlay exists at all', () => {
    // Almost always the case: the listener lives for the grid's whole life, so
    // it must do nothing on ordinary typing outside an edit.
    render(<SpreadsheetDataEditor rows={1} columns={[]} getCellContent={() => ({}) as never} />)
    const canvas = focusElement('canvas')
    const before = document.body.innerHTML

    dispatchKey('w')

    expect(document.body.innerHTML).toBe(before)
    expect(document.activeElement).toBe(canvas)
  })
})

describe('overlay keystroke rescue (MATRIX-FLAKE-2)', () => {
  function mountEditor(): void {
    render(<SpreadsheetDataEditor rows={1} columns={[]} getCellContent={() => ({}) as never} />)
    render(<div id="portal">{captured.render()}</div>)
  }

  it('applies a character typed while focus is still on the accessibility cell', () => {
    mountEditor()
    const textarea = overlay()
    expect(textarea.value).toBe('a')

    focusAccessibilityCell()
    dispatchKey('f')

    // Without the rescue this character is delivered to the `<td>` and lost.
    expect(textarea.value).toBe('af')
    expect(captured.changes.at(-1)).toBe('af')
  })

  it('accumulates several misplaced characters in order', () => {
    mountEditor()
    const textarea = overlay()
    focusAccessibilityCell()

    for (const key of ['f', 't', 'e', 'r']) dispatchKey(key)

    // The whole word, which is what the failing e2e was asserting.
    expect(textarea.value).toBe('after')
  })

  it('commits on an Enter that lands in the wrong place', () => {
    // The half that made the failure look like "the save did nothing": the
    // edit never committed, so the Ctrl+S behind it had nothing to write.
    mountEditor()
    focusAccessibilityCell()
    dispatchKey('f')
    dispatchKey('Enter')

    expect(captured.finished).toHaveLength(1)
    expect(captured.finished[0].data).toBe('af')
    expect(captured.finished[0].movement).toEqual([0, 1])
  })

  it('commits sideways on a misplaced Tab, and backwards on Shift+Tab', () => {
    mountEditor()
    focusAccessibilityCell()
    dispatchKey('Tab')
    expect(captured.finished.at(-1)?.movement).toEqual([1, 0])

    captured.finished = []
    focusAccessibilityCell()
    dispatchKey('Tab', { shiftKey: true })
    expect(captured.finished.at(-1)?.movement).toEqual([-1, 0])
  })

  it.each(['canvas', 'div'] as const)('rescues a key when focus is on the grid %s', (tag) => {
    // The regression the first fix missed. Enumerating wrong places left these
    // out; asking "is this a text target the user chose" covers them.
    mountEditor()
    focusElement(tag)

    dispatchKey('f')

    expect(overlay().value).toBe('af')
  })

  it('rescues a key when focus has been lost entirely', () => {
    mountEditor()
    const stray = focusElement('div')
    stray.blur()

    dispatchKey('f')

    expect(overlay().value).toBe('af')
  })

  it('does NOT steal keystrokes from a control the user deliberately focused', () => {
    // The safety property. If the user clicked into a search box while an
    // overlay happened to be open, their typing belongs to that box.
    mountEditor()
    const other = document.createElement('input')
    document.body.appendChild(other)
    other.focus()

    dispatchKey('f')

    expect(overlay().value).toBe('a')
    expect(captured.changes).toEqual([])
  })

  it.each(['input', 'select'] as const)('leaves a focused %s alone', (tag) => {
    // The safety property, for each kind of control the condition names.
    mountEditor()
    const el = document.createElement(tag)
    document.body.appendChild(el)
    el.focus()

    dispatchKey('f')

    expect(overlay().value).toBe('a')
  })

  it('leaves a focused contenteditable alone', () => {
    mountEditor()
    const el = document.createElement('div')
    el.setAttribute('contenteditable', 'true')
    // jsdom does not derive `isContentEditable` from the attribute.
    Object.defineProperty(el, 'isContentEditable', { value: true })
    el.tabIndex = 0
    document.body.appendChild(el)
    el.focus()

    dispatchKey('f')

    expect(overlay().value).toBe('a')
  })

  it('does nothing once the overlay has focus itself', () => {
    // The normal path: the textarea handles its own keys, and the rescue must
    // not double-apply them.
    mountEditor()
    const textarea = overlay()
    textarea.focus()

    dispatchKey('f')

    expect(textarea.value).toBe('a')
    expect(captured.changes).toEqual([])
  })

  it('ignores a shortcut rather than typing it into the cell', () => {
    // Ctrl+S must stay a save, not become the letter "s" in a cell.
    mountEditor()
    focusAccessibilityCell()

    dispatchKey('s', { ctrlKey: true })
    dispatchKey('c', { metaKey: true })
    dispatchKey('F2')
    dispatchKey('ArrowDown')

    expect(overlay().value).toBe('a')
    expect(captured.changes).toEqual([])
  })
})
