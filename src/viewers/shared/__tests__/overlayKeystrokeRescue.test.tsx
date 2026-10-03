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
import { render } from '@testing-library/react'
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
      captured.render = () => {
        if (!Editor) return null
        return (
          <Editor
            value={{ kind: actual.GridCellKind.Text, data: 'a', displayData: 'a', allowOverlay: true }}
            initialValue="a"
            forceEditMode
            isHighlighted={false}
            onChange={(next: { data: string }) => captured.changes.push(next.data)}
            onFinishedEditing={(next: { data: string } | undefined, movement: readonly [number, number]) => {
              captured.finished.push({ data: next?.data ?? '', movement })
            }}
          />
        )
      }
      return null
    },
  }
})

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

function dispatchKey(key: string, options: KeyboardEventInit = {}): void {
  document.activeElement?.dispatchEvent(
    new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...options }),
  )
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
