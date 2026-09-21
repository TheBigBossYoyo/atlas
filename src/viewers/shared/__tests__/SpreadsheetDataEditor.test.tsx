/**
 * F6 — the editing lifecycle must seed a freshly-opened cell editor with
 * EVERY keystroke that raced ahead of glide-data-grid's own overlay-mount
 * (not just the very first one), not silently drop the ones that arrived
 * before the textarea existed. See SpreadsheetDataEditor.tsx's module header
 * for the full root-cause writeup — this locks in the fix at the unit level,
 * independent of the e2e reproduction against real saved bytes.
 *
 * `DataEditor` itself is mocked (as in SpreadsheetViewer.editing.test.tsx):
 * real glide-data-grid renders to a `<canvas>` and needs a live DOM/portal
 * timing environment this file doesn't try to reproduce. What's under test
 * is `SpreadsheetDataEditor`'s OWN wrapping — the `onKeyDown` interception
 * and the `TextCellEditor` it hands to `provideEditor` — driven by calling
 * the props the (mocked) `DataEditor` captured, exactly the shape/sequence
 * the real library would call them with per the traced root cause:
 * `onKeyDown({key, location, cancel, ...})` for every keydown, then
 * `provideEditor(cell)` mounting the returned component with
 * `forceEditMode`/`initialValue` seeded from the first one. The mock renders
 * that returned component AS ITS OWN CHILD (toggled via `openEditor`) rather
 * than in a separate `render()` call, so it stays nested under
 * `SpreadsheetDataEditor`'s own providers exactly like the real portal-mounted
 * overlay does.
 */
import { render } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'
import type { ComponentType } from 'react'
import type { DataEditorProps, EditableGridCell, GridCell, GridKeyEventArgs, Item } from '@glideapps/glide-data-grid'

let lastDataEditorProps: DataEditorProps | null = null
let openEditor: {
  cell: GridCell
  forceEditMode: boolean
  initialValue?: string
  onChange: (v: EditableGridCell) => void
  onFinishedEditing?: (v: EditableGridCell | undefined, movement: readonly [number, number]) => void
} | null = null

vi.mock('@glideapps/glide-data-grid', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@glideapps/glide-data-grid')>()
  return {
    ...actual,
    DataEditor: (props: DataEditorProps) => {
      lastDataEditorProps = props
      if (openEditor === null) return null
      const editorProvider = props.provideEditor?.(openEditor.cell)
      const Editor = (typeof editorProvider === 'function' ? editorProvider : editorProvider?.editor) as
        | ComponentType<Record<string, unknown>>
        | undefined
      if (!Editor) return null
      return (
        <Editor
          value={openEditor.cell}
          onChange={openEditor.onChange}
          onFinishedEditing={openEditor.onFinishedEditing ?? vi.fn()}
          isHighlighted={false}
          forceEditMode={openEditor.forceEditMode}
          initialValue={openEditor.initialValue}
          target={{ x: 0, y: 0, width: 100, height: 30 }}
          theme={{}}
        />
      )
    },
  }
})

// Imported AFTER the mock so it wraps the mocked DataEditor (matches the
// existing SpreadsheetViewer.editing.test.tsx pattern).
const { SpreadsheetDataEditor } = await import('../SpreadsheetDataEditor')

function textCell(data: string): GridCell {
  return { kind: 'text', data, displayData: data, allowOverlay: true } as GridCell
}

function keyEvent(overrides: Partial<GridKeyEventArgs> & { key: string; location: Item }): GridKeyEventArgs {
  return {
    bounds: { x: 0, y: 0, width: 100, height: 30 },
    keyCode: 0,
    altKey: false,
    shiftKey: false,
    ctrlKey: false,
    metaKey: false,
    cancel: vi.fn(),
    stopPropagation: vi.fn(),
    preventDefault: vi.fn(),
    rawEvent: undefined,
    ...overrides,
  }
}

describe('SpreadsheetDataEditor — first-keystroke seeding race (F6)', () => {
  it('buffers keystrokes that arrive for the same cell before the overlay mounts, then applies them all, in order, once it does', () => {
    openEditor = null
    const { rerender } = render(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)
    const props = lastDataEditorProps!
    expect(props.onKeyDown).toBeDefined()

    // First qualifying keystroke ("H") for cell (0, 0): glide-data-grid's own
    // editOnType would open the overlay from this — we must NOT cancel it.
    const first = keyEvent({ key: 'H', location: [0, 0] })
    props.onKeyDown!(first)
    expect(first.cancel).not.toHaveBeenCalled()

    // Second and third keystrokes for the SAME cell, arriving before the
    // overlay has mounted (glide-data-grid's own accessibility element still
    // has focus at this point — see module header). These must be
    // intercepted so they can't re-trigger `reselect()` and wipe out "H".
    const second = keyEvent({ key: 'E', location: [0, 0] })
    props.onKeyDown!(second)
    expect(second.cancel).toHaveBeenCalled()
    expect(second.preventDefault).toHaveBeenCalled()

    const third = keyEvent({ key: 'L', location: [0, 0] })
    props.onKeyDown!(third)
    expect(third.cancel).toHaveBeenCalled()

    // The overlay finally mounts (type-triggered: forceEditMode + initialValue,
    // exactly like glide-data-grid's real `reselect(bounds, true, 'H')`),
    // seeded with just "H" per glide-data-grid's own reselect() — the buffered
    // "E" and "L" must be appended once, in order, on mount.
    const onChange = vi.fn()
    openEditor = { cell: textCell('H'), forceEditMode: true, initialValue: 'H', onChange }
    rerender(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)

    expect(onChange).toHaveBeenCalledWith(expect.objectContaining({ data: 'HEL' }))
  })

  it('commits immediately on mount when Enter arrived before the overlay did (typing a whole word faster than the mount)', () => {
    // The measured failure mode this covers: at ordinary typing speed, a
    // short word plus Enter can finish before the overlay has mounted at
    // all (see module header — ~200-300ms observed). Enter then still
    // targets glide-data-grid's own accessibility element, which treats it
    // as plain navigation and would otherwise discard the whole edit.
    openEditor = null
    const { rerender } = render(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)
    const props = lastDataEditorProps!

    const first = keyEvent({ key: 'H', location: [0, 0] })
    props.onKeyDown!(first)
    expect(first.cancel).not.toHaveBeenCalled()

    for (const key of ['E', 'L', 'L', 'O']) {
      const ev = keyEvent({ key, location: [0, 0] })
      props.onKeyDown!(ev)
      expect(ev.cancel).toHaveBeenCalled()
    }

    const enter = keyEvent({ key: 'Enter', location: [0, 0] })
    props.onKeyDown!(enter)
    expect(enter.cancel).toHaveBeenCalled()
    expect(enter.preventDefault).toHaveBeenCalled()

    const onChange = vi.fn()
    const onFinishedEditing = vi.fn()
    openEditor = { cell: textCell('H'), forceEditMode: true, initialValue: 'H', onChange, onFinishedEditing }
    rerender(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)

    expect(onFinishedEditing).toHaveBeenCalledWith(expect.objectContaining({ data: 'HELLO' }), [0, 1])
    // The commit path skips the intermediate onChange entirely — it hands
    // the full seeded text straight to onFinishedEditing.
    expect(onChange).not.toHaveBeenCalled()
  })

  it('does not touch the buffer for a double-click/F2 open (forceEditMode false, no initialValue)', () => {
    openEditor = null
    const { rerender } = render(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)
    const props = lastDataEditorProps!

    // A keystroke buffered for a DIFFERENT interaction (e.g. a previous,
    // abandoned type-to-edit) must never leak into an unrelated double-click
    // open — forceEditMode/initialValue is the guard (see module header).
    props.onKeyDown!(keyEvent({ key: 'H', location: [0, 0] }))
    props.onKeyDown!(keyEvent({ key: 'E', location: [0, 0] }))

    const onChange = vi.fn()
    openEditor = { cell: textCell('Existing'), forceEditMode: false, onChange }
    rerender(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)

    expect(onChange).not.toHaveBeenCalled()
  })

  it('does not intercept a keystroke for a DIFFERENT cell than the one currently opening', () => {
    openEditor = null
    render(<SpreadsheetDataEditor getCellContent={() => textCell('')} columns={[]} rows={10} />)
    const props = lastDataEditorProps!

    props.onKeyDown!(keyEvent({ key: 'H', location: [0, 0] }))
    // A keystroke that lands on a different cell (e.g. the user re-selected
    // elsewhere) is a fresh edit-start, not a race — glide-data-grid must
    // handle it normally.
    const elsewhere = keyEvent({ key: 'X', location: [2, 5] })
    props.onKeyDown!(elsewhere)
    expect(elsewhere.cancel).not.toHaveBeenCalled()
  })
})
