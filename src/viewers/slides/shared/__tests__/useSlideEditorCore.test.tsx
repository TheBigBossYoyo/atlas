/**
 * USR-16 — the editing session shared by the PPTX and ODP editors.
 *
 * Written to close a coverage gap found by an audit of the tree: this hook is
 * the ONE core both slide editors delegate to (package load, undo/redo, the
 * edit queue, the dirty/save contract, the editor's own Ctrl+Z/Ctrl+Y), and it
 * sat at 49% statements / 12.5% branches with no test file of its own — every
 * line here was only ever exercised incidentally, through whichever PPTX/ODP
 * viewer test happened to mount it. A regression in it lands in both formats
 * at once, which is exactly the case worth testing directly.
 *
 * The format-specific halves (`parseDeck`, `saveFilter`) are supplied as spies,
 * since the point is the shared machinery and not either format's parser. The
 * package itself is a real zip through `loadOfficePackage`/`writeOfficePackage`
 * rather than a mock, so the round trip these tests assert on is the real one.
 */
import { act, render, screen, waitFor } from '@testing-library/react'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { ShortcutManagerProvider } from '../../../../hooks/ShortcutManagerProvider'
import { withParts, type OfficePackage } from '../../../../office/officePackage'
import { ViewerProvider } from '../../../shared/ViewerContext'
import { useViewerIsDirty, useViewerSaveAs } from '../../../shared/useViewerContext'
import type { SlideData } from '../../../shared/SlideDeck.types'
import { useSlideEditorCore, type DeckState, type SlideEditorCore } from '../useSlideEditorCore'

const FILE_PATH = 'C:\\decks\\quarterly.pptx'
const SAVE_FILTER = { name: 'PowerPoint Presentation', extensions: ['pptx'] }
const SLIDE1_TEXT = '<p:sld><t>one</t></p:sld>'

/** A real (tiny) OPC-shaped zip — `loadOfficePackage` unzips this for real. */
async function deckBuffer(slideXml = SLIDE1_TEXT): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<Types/>')
  zip.file('ppt/presentation.xml', '<p:presentation/>')
  zip.file('ppt/slides/slide1.xml', slideXml)
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

function slide(title: string): SlideData {
  return { title, blocks: [], notes: null } as unknown as SlideData
}

/** Parses "however many slide parts there are" into one SlideData each, named after the part's text. */
function countingParseDeck(pkg: OfficePackage): ReadonlyArray<SlideData> {
  const slides: SlideData[] = []
  for (const [path, content] of pkg.parts) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(path)) continue
    slides.push(slide(typeof content === 'string' ? content : path))
  }
  return slides
}

/** The rendered deck as a list, for assertions that must distinguish 'two' from 'two-prime'. */
function slideTitles(): string[] {
  const text = screen.getByTestId('slides').textContent ?? ''
  return text === '' ? [] : text.split('|')
}

function DirtyProbe() {
  return <span data-testid="dirty">{String(useViewerIsDirty())}</span>
}

/** Proves `registerSaveAs` actually reached the context (the Ctrl+Shift+S path). */
function ContextSaveAsProbe() {
  const saveAs = useViewerSaveAs()
  return (
    <button type="button" onClick={() => void saveAs()}>
      Context Save As
    </button>
  )
}

type HarnessOptions = {
  readonly buffer: ArrayBuffer | null
  readonly parseDeck: SlideEditorCoreOptionsParse
  readonly onSavedPath?: (startedFrom: string, saved: string) => void
}
type SlideEditorCoreOptionsParse = (pkg: OfficePackage, previous: DeckState | null) => Promise<ReadonlyArray<SlideData>>

function Harness({ buffer, parseDeck, onReady }: HarnessOptions & { onReady: (core: SlideEditorCore) => void }) {
  const core = useSlideEditorCore({ buffer, filePath: FILE_PATH, parseDeck, saveFilter: SAVE_FILTER })
  onReady(core)
  return (
    <>
      <span data-testid="status">{core.status}</span>
      <span data-testid="slides">{core.slides.map((s) => (s as { title: string }).title).join('|')}</span>
      <span data-testid="can-undo">{String(core.canUndo)}</span>
      <span data-testid="can-redo">{String(core.canRedo)}</span>
      {core.error !== null && <span data-testid="error">{core.error}</span>}
      {core.saveError !== null && <span data-testid="save-error">{core.saveError}</span>}
      <DirtyProbe />
      <ContextSaveAsProbe />
    </>
  )
}

function renderCore(options: HarnessOptions) {
  let latest: SlideEditorCore | null = null
  const utils = render(
    <ShortcutManagerProvider>
      <ViewerProvider filePath={FILE_PATH} onSavedPath={options.onSavedPath}>
        <Harness {...options} onReady={(core) => { latest = core }} />
      </ViewerProvider>
    </ShortcutManagerProvider>,
  )
  return { ...utils, core: () => latest as unknown as SlideEditorCore }
}

/** Adds a second slide part — a real edit, so `apply` produces a new package. */
function addSlide(text: string) {
  return (pkg: OfficePackage) => withParts(pkg, { 'ppt/slides/slide2.xml': text })
}

let saveBinaryFile: ReturnType<typeof vi.fn>

beforeEach(() => {
  saveBinaryFile = vi.fn().mockResolvedValue({ saved: true, path: FILE_PATH, name: 'quarterly.pptx' })
  window.electronAPI = { saveBinaryFile } as unknown as typeof window.electronAPI
})

afterEach(() => {
  vi.restoreAllMocks()
})

describe('useSlideEditorCore — loading the package', () => {
  it('parses the package and reports ready with the parsed slides', async () => {
    const parseDeck = vi.fn<SlideEditorCoreOptionsParse>(async (pkg) => countingParseDeck(pkg))
    renderCore({ buffer: await deckBuffer(), parseDeck })

    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    expect(screen.getByTestId('slides')).toHaveTextContent('<p:sld><t>one</t></p:sld>')
    expect(screen.queryByTestId('error')).toBeNull()
    // First parse gets `previous: null` — there is no earlier deck to reuse.
    expect(parseDeck).toHaveBeenCalledTimes(1)
    expect(parseDeck.mock.calls[0][1]).toBeNull()
  })

  it('stays loading while there are no bytes yet', async () => {
    const parseDeck = vi.fn(async () => [])
    renderCore({ buffer: null, parseDeck })

    expect(screen.getByTestId('status')).toHaveTextContent('loading')
    expect(parseDeck).not.toHaveBeenCalled()
  })

  it('reports an error for a package that parses to zero slides', async () => {
    renderCore({ buffer: await deckBuffer(), parseDeck: async () => [] })

    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('error'))
    expect(screen.getByTestId('error')).toHaveTextContent('No slides found in this presentation.')
  })

  it("surfaces a parser's own failure message", async () => {
    renderCore({
      buffer: await deckBuffer(),
      parseDeck: async () => {
        throw new Error('slide3.xml is not valid XML')
      },
    })

    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('error'))
    expect(screen.getByTestId('error')).toHaveTextContent('slide3.xml is not valid XML')
  })

  it('surfaces an unzip failure for bytes that are not a package at all', async () => {
    const notAZip = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]).buffer
    renderCore({ buffer: notAZip, parseDeck: async (pkg) => countingParseDeck(pkg) })

    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('error'))
    expect(screen.getByTestId('error')).not.toHaveTextContent('No slides found')
  })
})

describe('useSlideEditorCore — applying edits', () => {
  it('re-parses and records history for an edit that changes the package', async () => {
    const parseDeck = vi.fn<SlideEditorCoreOptionsParse>(async (pkg) => countingParseDeck(pkg))
    const { core } = renderCore({ buffer: await deckBuffer(), parseDeck })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      await core().apply(addSlide('two'))
    })

    expect(screen.getByTestId('slides')).toHaveTextContent('two')
    expect(screen.getByTestId('can-undo')).toHaveTextContent('true')
    // The re-parse is handed the PREVIOUS deck, so a format can reuse slides
    // it knows are untouched instead of re-rendering all of them.
    expect(parseDeck.mock.calls[1][1]).not.toBeNull()
  })

  it('is a no-op when the edit returns the same package (no re-parse, no history entry)', async () => {
    const parseDeck = vi.fn<SlideEditorCoreOptionsParse>(async (pkg) => countingParseDeck(pkg))
    const { core } = renderCore({ buffer: await deckBuffer(), parseDeck })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    parseDeck.mockClear()

    await act(async () => {
      await core().apply((pkg) => pkg)
    })

    expect(parseDeck).not.toHaveBeenCalled()
    expect(screen.getByTestId('can-undo')).toHaveTextContent('false')
    expect(screen.getByTestId('dirty')).toHaveTextContent('false')
  })

  it('serializes concurrent edits so each one sees the previous one applied', async () => {
    // The whole reason `apply` owns a queue: both edits below read the package
    // they are handed. Without serialization the second would start from the
    // pre-first package and silently drop slide2.
    const seen: number[] = []
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => {
        // A slow parse widens the window an unqueued second edit would race in.
        await new Promise((resolve) => setTimeout(resolve, 5))
        return countingParseDeck(pkg)
      },
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      const first = core().apply((pkg) => {
        seen.push(pkg.parts.size)
        return withParts(pkg, { 'ppt/slides/slide2.xml': 'two' })
      })
      const second = core().apply((pkg) => {
        seen.push(pkg.parts.size)
        return withParts(pkg, { 'ppt/slides/slide3.xml': 'three' })
      })
      await Promise.all([first, second])
    })

    expect(seen).toEqual([3, 4])
    expect(screen.getByTestId('slides')).toHaveTextContent('two|three')
  })
})

describe('useSlideEditorCore — undo/redo', () => {
  it('undo restores the previous deck and redo reapplies it', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    await act(async () => {
      core().undo()
    })
    await waitFor(() => expect(screen.getByTestId('slides')).not.toHaveTextContent('two'))
    expect(screen.getByTestId('can-redo')).toHaveTextContent('true')

    await act(async () => {
      core().redo()
    })
    await waitFor(() => expect(screen.getByTestId('slides')).toHaveTextContent('two'))
  })

  it('an edit queued behind an undo starts from the undone deck, not the one before it', async () => {
    // Why undo/redo join the same queue as `apply`.
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    await act(async () => {
      core().undo()
      await core().apply(addSlide('two-prime'))
    })

    // Exact list, not a substring check: 'two' is a prefix of 'two-prime', so
    // "does it still contain two" cannot tell the two decks apart.
    expect(slideTitles()).toEqual([SLIDE1_TEXT, 'two-prime'])
  })

  it('Ctrl+Z undoes and Ctrl+Y / Ctrl+Shift+Z redo', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }))
    })
    await waitFor(() => expect(screen.getByTestId('slides')).not.toHaveTextContent('two'))

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'y', ctrlKey: true, bubbles: true }))
    })
    await waitFor(() => expect(screen.getByTestId('slides')).toHaveTextContent('two'))

    // Ctrl+Shift+Z is the OTHER redo binding, not an undo - so undo once more
    // first, then prove Shift+Z redoes the same way Ctrl+Y just did.
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }))
    })
    await waitFor(() => expect(slideTitles()).toEqual([SLIDE1_TEXT]))

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, shiftKey: true, bubbles: true }))
    })
    await waitFor(() => expect(slideTitles()).toEqual([SLIDE1_TEXT, 'two']))
  })

  it('leaves Ctrl+Z alone while a text field or contentEditable has focus', async () => {
    // The field's own native undo must keep working — the deck must not move.
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    const input = document.createElement('input')
    document.body.appendChild(input)
    const textarea = document.createElement('textarea')
    document.body.appendChild(textarea)
    // `<select>` is in this list because the shared `isPlainFieldTarget` does
    // NOT cover it — the handler checks it separately, and this is what keeps
    // that from being dropped as redundant.
    const select = document.createElement('select')
    document.body.appendChild(select)
    const editable = document.createElement('div')
    // `setAttribute`, not the `contentEditable` property: jsdom implements
    // neither `isContentEditable` nor the property's attribute reflection, so
    // the property alone leaves nothing for the dispatcher's documented
    // attribute-lookup fallback to find. React emits the attribute for
    // `contentEditable={true}`, so this is also what the real DOM looks like.
    editable.setAttribute('contenteditable', 'true')
    document.body.appendChild(editable)
    const insideEditable = document.createElement('span')
    editable.appendChild(insideEditable)

    for (const target of [input, textarea, select, editable, insideEditable]) {
      await act(async () => {
        target.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true }))
      })
      expect(slideTitles()).toEqual([SLIDE1_TEXT, 'two'])
    }

    input.remove()
    textarea.remove()
    select.remove()
    editable.remove()
  })

  it('ignores an unmodified z (plain typing is not an undo)', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', bubbles: true }))
    })

    expect(screen.getByTestId('slides')).toHaveTextContent('two')
  })
})

describe('useSlideEditorCore — dirty tracking and save', () => {
  it('starts clean, goes dirty on an edit, and is clean again after a save', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    // Waits on `status`, NOT on `dirty` being false. `dirty` is false before the
    // package has loaded as well as after (`savedPkg` is still null, so the
    // dirty check is false by definition), so waiting on it would let the edit
    // below run against the empty pre-load deck — and the load's own
    // `history.reset` would then wipe that edit, leaving the deck unchanged and
    // clean. That is a real race this test hit intermittently, and only under
    // parallel load, because the load resolving first is what usually hides it.
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    expect(screen.getByTestId('dirty')).toHaveTextContent('false')

    await act(async () => {
      await core().apply(addSlide('two'))
    })
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('true'))

    await act(async () => {
      expect(await core().save()).toBe(true)
    })
    await waitFor(() => expect(screen.getByTestId('dirty')).toHaveTextContent('false'))
  })

  it('saves in place: the opened path is passed as existingPath, with the format filter', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      await core().save()
    })

    const req = saveBinaryFile.mock.calls[0][0] as {
      content: Uint8Array
      suggestedName: string
      existingPath?: string
      filters: Array<{ name: string; extensions: string[] }>
    }
    expect(req.existingPath).toBe(FILE_PATH)
    expect(req.suggestedName).toBe('quarterly.pptx')
    expect(req.filters).toEqual([SAVE_FILTER])
    // A real re-zip of the live package, not a passthrough of the input bytes.
    expect(req.content.byteLength).toBeGreaterThan(0)
  })

  it('saveAs omits existingPath so the main process always shows the dialog', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      await core().saveAs()
    })

    expect(saveBinaryFile.mock.calls[0][0]).not.toHaveProperty('existingPath')
  })

  it('reports the new path to the shell when Save As writes somewhere else (SAVE-1)', async () => {
    const onSavedPath = vi.fn()
    saveBinaryFile.mockResolvedValue({ saved: true, path: 'C:\\decks\\copy.pptx', name: 'copy.pptx' })
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
      onSavedPath,
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      await core().saveAs()
    })

    // The path the save STARTED from is what identifies the tab, not whatever
    // is showing when the write resolves.
    expect(onSavedPath).toHaveBeenCalledWith(FILE_PATH, 'C:\\decks\\copy.pptx')
  })

  it('does not report a path when the save landed on the same file', async () => {
    const onSavedPath = vi.fn()
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
      onSavedPath,
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      await core().save()
    })

    expect(onSavedPath).not.toHaveBeenCalled()
  })

  it('surfaces the main process\'s error and stays dirty when the save fails', async () => {
    saveBinaryFile.mockResolvedValue({ saved: false, error: 'This file is open in PowerPoint.' })
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))
    await act(async () => {
      await core().apply(addSlide('two'))
    })

    await act(async () => {
      expect(await core().save()).toBe(false)
    })

    expect(screen.getByTestId('save-error')).toHaveTextContent('This file is open in PowerPoint.')
    expect(screen.getByTestId('dirty')).toHaveTextContent('true')
  })

  it('falls back to a generic message when a cancelled save reports no error', async () => {
    saveBinaryFile.mockResolvedValue({ saved: false })
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      expect(await core().save()).toBe(false)
    })

    expect(screen.getByTestId('save-error')).toHaveTextContent('Save was cancelled or unavailable.')
  })

  it('reports a thrown write error instead of rejecting', async () => {
    saveBinaryFile.mockRejectedValue(new Error('EACCES: permission denied'))
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      expect(await core().save()).toBe(false)
    })

    expect(screen.getByTestId('save-error')).toHaveTextContent('EACCES: permission denied')
  })

  it('waits for a queued edit before writing, so a save never misses the last keystroke', async () => {
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => {
        await new Promise((resolve) => setTimeout(resolve, 5))
        return countingParseDeck(pkg)
      },
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      void core().apply(addSlide('two'))
      await core().save() // queued behind the still-running re-parse above
    })

    const written = (saveBinaryFile.mock.calls[0][0] as { content: Uint8Array }).content
    const reread = await JSZip.loadAsync(written)
    expect(reread.files['ppt/slides/slide2.xml']).toBeDefined()
  })

  it('registers Save As on the context, so the global Ctrl+Shift+S reaches it', async () => {
    renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      screen.getByRole('button', { name: 'Context Save As' }).click()
    })

    await waitFor(() => expect(saveBinaryFile).toHaveBeenCalledTimes(1))
    expect(saveBinaryFile.mock.calls[0][0]).not.toHaveProperty('existingPath')
  })

  it('reports saved:false rather than throwing when the desktop bridge is absent', async () => {
    window.electronAPI = undefined as unknown as typeof window.electronAPI
    const { core } = renderCore({
      buffer: await deckBuffer(),
      parseDeck: async (pkg) => countingParseDeck(pkg),
    })
    await waitFor(() => expect(screen.getByTestId('status')).toHaveTextContent('ready'))

    await act(async () => {
      expect(await core().save()).toBe(false)
    })

    expect(screen.getByTestId('save-error')).toHaveTextContent('Save was cancelled or unavailable.')
  })
})
