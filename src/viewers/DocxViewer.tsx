import {
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type FormEvent,
  type ClipboardEvent as ReactClipboardEvent,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'

import { Printer, Save, X } from 'lucide-react'

import type { NavItem, ViewerProps } from '../formats/types'
import { loadDocx, saveDocx, type DocxBundle } from '../docx'
import { paginate, PaginationCancelledError } from '../docx/layout'
import type { Page, PaginationProgress } from '../docx/layout'
import type { FontResolver } from '../docx/layout/types'
import { FONT_FAMILIES, loadFontMetrics, resolveFontFamily, wrapWithCanvasAdvance, type FontMetrics, type FontVariant } from '../docx/fonts'
import { type Document as DocxDocument, type Paragraph, type ParagraphChild, type Run, type RunChild } from '../docx/model'
import { resolveParaProps } from '../docx/parser/cascade'
import { MediaContext, PageStack } from '../docx/render'
import { useArchiveMediaResolver } from '../docx/render/archiveMedia'
import {
  CommentsPane,
  FindReplace,
  History,
  SpellCheckMenu,
  applyCommand,
  buildReplaceAllCommands,
  buildReplaceCommands,
  buildSpellCheckReplacement,
  decodeImageNaturalSizePt,
  ensureListNumbering,
  findAll,
  findNext,
  findEnclosingTable,
  findParagraph,
  findPrev,
  handleBeforeInput,
  handleKeyDown,
  insertHyperlinkIntoBundle,
  insertImageIntoBundle,
  buildPasteCommands,
  friendlyDocxErrorMessage,
  htmlToParagraphs,
  textToParagraphs,
  toolbarToCommand,
  useComposition,
  useSpellCheck,
  type Command,
  type EnclosingTable,
  type FindOptions,
  type ImageMimeType,
  type Range,
  type TrackChangesContext,
} from '../docx/editor'
import { addCommentToDocument, deleteCommentFromDocument, replyToComment } from '../docx/editor/commentMutations'
import { comparePositions } from '../docx/editor/Selection'
import { domPointToPosition, positionToDomRange } from '../docx/editor/Cursor'
import { Toolbar } from '../docx/editor/toolbar/Toolbar'
import type { ToolbarCommand, ToolbarState } from '../docx/editor/toolbar/toolbarTypes'
import './__styles__/viewer-docx.css'
import { useRegisterViewerSave, useSetNavItems, useSetViewerDirty, useSetViewerStats } from './shared/useViewerContext'
import { useViewerShortcuts } from '../hooks/useShortcutManager'

type HeadingNavSeed = {
  id: string
  label: string
  level: number
  paragraphIndex: number
}

const DEFAULT_FIND_OPTIONS: FindOptions = {
  caseSensitive: false,
  wholeWord: false,
  useRegex: false,
}

const DOCX_SAVE_FILTERS = [{ name: 'Word Documents', extensions: ['docx'] }]

/** D12/DXE-03 — keys whose native contentEditable behavior always mutates
 * content; see handleKeyDownEvent's preventDefault-safety comment. */
const MUTATING_KEYS: ReadonlySet<string> = new Set(['Backspace', 'Delete', 'Enter', 'Tab'])

const DEFAULT_FONT_METRICS: FontMetrics = Object.freeze({
  unitsPerEm: 1000,
  ascender: 800,
  descender: -200,
  lineGap: 0,
  xHeight: 500,
  capHeight: 700,
  advanceWidth: () => 500,
  hasGlyph: () => true,
})

/**
 * Register every bundled substitute font via the FontFace API using URLs
 * emitted by Vite (`?url` imports in families.ts). This replaces the previous
 * CSS @font-face approach which used absolute `/fonts/...` URLs that resolve
 * to `file:///C:/fonts/...` in the packaged Electron app and 404, leaving
 * canvas measurements returning 0/NaN and producing blank pages (3.0.7 bug
 * on LaBoetie_Dossier_v2_Bac2026.docx). Registering both wordName and
 * substituteName guarantees the paginator's measured typeface is the one
 * the browser paints.
 */
const FONT_VARIANT_DESCRIPTORS: ReadonlyArray<{
  readonly variant: FontVariant
  readonly weight: string
  readonly style: string
}> = [
  { variant: 'regular', weight: '400', style: 'normal' },
  { variant: 'bold', weight: '700', style: 'normal' },
  { variant: 'italic', weight: '400', style: 'italic' },
  { variant: 'boldItalic', weight: '700', style: 'italic' },
]

let fontsRegistered = false
let fontsRegisteringPromise: Promise<void> | null = null

async function preloadDocxFonts(): Promise<void> {
  if (typeof document === 'undefined' || document.fonts === undefined) {
    return
  }
  if (fontsRegistered) {
    return
  }
  if (fontsRegisteringPromise !== null) {
    return fontsRegisteringPromise
  }

  fontsRegisteringPromise = (async () => {
    const loads: Promise<FontFace>[] = []
    for (const family of FONT_FAMILIES) {
      for (const descriptor of FONT_VARIANT_DESCRIPTORS) {
        const url = family.files[descriptor.variant]
        for (const name of [family.wordName, family.substituteName]) {
          const face = new FontFace(name, `url(${url})`, {
            weight: descriptor.weight,
            style: descriptor.style,
            display: 'block',
          })
          loads.push(
            face.load().then((loaded) => {
              document.fonts.add(loaded)
              return loaded
            }),
          )
        }
      }
    }
    await Promise.all(loads)
    await document.fonts.ready
    fontsRegistered = true
  })()

  try {
    await fontsRegisteringPromise
  } finally {
    fontsRegisteringPromise = null
  }
}

function createFontResolver(): FontResolver {
  const cache = new Map<string, Promise<FontMetrics>>()

  return async (family: string, variant: FontVariant): Promise<FontMetrics> => {
    const resolved = resolveFontFamily(family)
    if (resolved === null) {
      // Unknown family: still measure via canvas under the requested name so the
      // browser's OS-fallback width is what the paginator sees. We use a
      // synthetic 1000-em base with neutral vertical metrics; vertical metrics
      // for unknown fonts aren't critical (lines won't clip vertically).
      const syntheticBase: FontMetrics = {
        unitsPerEm: 1000,
        ascender: 800,
        descender: -200,
        lineGap: 0,
        xHeight: 500,
        capHeight: 700,
        advanceWidth: () => 500,
        hasGlyph: () => true,
      }
      return wrapWithCanvasAdvance(syntheticBase, family, variant)
    }

    const cacheKey = `${resolved.substituteName}@${variant}`
    const existing = cache.get(cacheKey)
    if (existing !== undefined) {
      return existing
    }

    // Load TTF for accurate vertical metrics, then override per-glyph advances
    // with canvas measurements taken under the requested Word font name. The
    // @font-face stack registered by fontFaces.css maps the Word name to the
    // bundled substitute, so canvas measures what the DOM will actually paint.
    const promise = loadFontMetrics(resolved.wordName, variant)
      .then((base) => wrapWithCanvasAdvance(base, resolved.wordName, variant))
      .catch(() => DEFAULT_FONT_METRICS)
    cache.set(cacheKey, promise)
    return promise
  }
}

function normalizeDocxBuffer(content: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (content instanceof ArrayBuffer) {
    return content
  }

  return Uint8Array.from(content).buffer
}

function getSuggestedFileName(path: string): string {
  const normalized = path.replace(/\\/g, '/')
  const parts = normalized.split('/')
  return parts[parts.length - 1] || 'document.docx'
}

function extractRunText(run: Run): string {
  return run.children.map(extractRunChildText).join('')
}

function extractRunChildText(child: RunChild): string {
  if (child.kind === 'text') {
    return child.value
  }

  if (child.kind === 'tab') {
    return ' '
  }

  if (child.kind === 'break') {
    return '\n'
  }

  return ''
}

function extractParagraphText(children: ReadonlyArray<ParagraphChild>): string {
  return children
    .map(child => {
      if (child.kind === 'run') {
        return extractRunText(child)
      }

      if (child.kind === 'hyperlink') {
        return child.children
          .map(grandchild => (grandchild.kind === 'run' ? extractRunText(grandchild) : ''))
          .join('')
      }

      return ''
    })
    .join('')
}

function getEditableRuns(paragraph: Paragraph): ReadonlyArray<Run> {
  return paragraph.children.flatMap(child => {
    if (child.kind === 'run') {
      return [child]
    }

    if (child.kind === 'hyperlink') {
      return child.children.flatMap(grandchild => (grandchild.kind === 'run' ? [grandchild] : []))
    }

    return []
  })
}

/**
 * DXE-14 — seed values for the table properties dialog, read straight off
 * the enclosing table's own `props` so opening the dialog shows what the
 * table actually has rather than a fixed default. `null` fields mean "not
 * inside a table" (the dialog stays disabled) or "no explicit value set on
 * this table" (the dialog falls back to its own placeholder default).
 * `tableBordersOn` only inspects the top border as a representative sample —
 * a table with a genuinely mixed on/off border set is a finer distinction
 * this basic on/off toggle doesn't attempt to preserve, matching Word's own
 * "Borders" quick-toggle rather than its full per-edge borders dialog.
 */
function tableToolbarFields(
  enclosing: EnclosingTable | null,
): Pick<ToolbarState, 'insideTable' | 'tableWidthTwips' | 'tableAlignment' | 'tableBordersOn'> {
  const props = enclosing?.table.props

  return {
    insideTable: enclosing !== null,
    tableWidthTwips: props?.tblW?.type === 'dxa' && typeof props.tblW.value === 'number' ? props.tblW.value : null,
    tableAlignment: props?.jc === 'center' ? 'center' : props?.jc === 'end' ? 'right' : props?.jc === 'start' ? 'left' : null,
    tableBordersOn: props?.tblBorders?.top?.style !== 'none' && props?.tblBorders?.top?.style !== 'nil',
  }
}

function createToolbarState(
  document: DocxDocument,
  range: Range | null,
  liveState: { readonly spellCheck: boolean; readonly trackChanges: boolean },
): ToolbarState {
  const paragraph = range ? findParagraph(document, range.focus.paragraphPath) : null
  const activeFormats = new Set<'bold' | 'italic' | 'underline' | 'strike' | 'subscript' | 'superscript'>()
  // DXE-14 — gates the table-editing button group and seeds the table
  // properties dialog.
  const enclosingTable = range !== null ? findEnclosingTable(document, range.focus.paragraphPath) : null
  const tableFields = tableToolbarFields(enclosingTable)

  if (paragraph !== null && range !== null) {
    const run = getEditableRuns(paragraph)[range.focus.runIndex]
    const props = run?.props

    if (props?.bold) {
      activeFormats.add('bold')
    }
    if (props?.italic) {
      activeFormats.add('italic')
    }
    if (props?.underline) {
      activeFormats.add('underline')
    }
    if (props?.strike || props?.dstrike) {
      activeFormats.add('strike')
    }
    if (props?.vertAlign === 'subscript') {
      activeFormats.add('subscript')
    }
    if (props?.vertAlign === 'superscript') {
      activeFormats.add('superscript')
    }

    return {
      activeFormats,
      alignment:
        paragraph.props?.jc === 'center'
          ? 'center'
          : paragraph.props?.jc === 'end'
            ? 'right'
            : paragraph.props?.jc === 'both'
              ? 'justify'
              : 'left',
      fontFamily: props?.rFonts?.ascii ?? props?.rFonts?.hAnsi ?? null,
      fontSizePt: typeof props?.sz === 'number' ? props.sz / 2 : null,
      styleId: paragraph.props?.pStyle ?? null,
      spellCheck: liveState.spellCheck,
      trackChanges: liveState.trackChanges,
      ...tableFields,
    }
  }

  return {
    activeFormats,
    alignment: null,
    fontFamily: null,
    fontSizePt: null,
    styleId: null,
    spellCheck: liveState.spellCheck,
    trackChanges: liveState.trackChanges,
    ...tableFields,
  }
}

function rangeEquals(left: Range | null, right: Range | null): boolean {
  if (left === null || right === null) {
    return left === right
  }

  return (
    comparePositions(left.anchor, right.anchor) === 0 &&
    comparePositions(left.focus, right.focus) === 0
  )
}

function getCurrentMatchIndex(matches: ReadonlyArray<{ range: Range }>, range: Range | null): number | null {
  if (range === null) {
    return null
  }

  const index = matches.findIndex(match => rangeEquals(match.range, range))
  return index >= 0 ? index : null
}

function visitBlock(
  block: DocxDocument['sections'][number]['blocks'][number],
  visitor: (paragraph: Paragraph, paragraphIndex: number) => void,
  paragraphIndexRef: { current: number },
): void {
  if (block.kind === 'paragraph') {
    visitor(block, paragraphIndexRef.current)
    paragraphIndexRef.current += 1
    return
  }

  if (block.kind === 'table') {
    block.rows.forEach(row => {
      if (row.kind !== 'table-row') return
      row.cells.forEach(cell => {
        if (cell.kind !== 'table-cell') return
        cell.blocks.forEach(cellBlock => {
          visitBlock(cellBlock, visitor, paragraphIndexRef)
        })
      })
    })
  }
}

function collectDocxMetrics(document: DocxDocument): {
  words: number
  navSeeds: ReadonlyArray<HeadingNavSeed>
} {
  const paragraphIndexRef = { current: 0 }
  const navSeeds: HeadingNavSeed[] = []
  let words = 0

  for (const section of document.sections) {
    for (const block of section.blocks) {
      visitBlock(
        block,
        (paragraph, paragraphIndex) => {
          const text = extractParagraphText(paragraph.children).trim()

          if (text.length > 0) {
            words += text.split(/\s+/).filter(Boolean).length
          }

          const effectiveProps = resolveParaProps(
            paragraph.props,
            paragraph.props?.pStyle,
            document.styles,
            { pPr: document.defaults?.paragraph },
          )

          const styleMatch = effectiveProps.pStyle?.match(/^Heading([1-6])$/)
          const outlineLevel =
            effectiveProps.outlineLvl !== undefined
              ? Math.min(Math.max(effectiveProps.outlineLvl + 1, 1), 6)
              : undefined
          const level = styleMatch ? Number(styleMatch[1]) : outlineLevel

          if (level !== undefined && text.length > 0) {
            navSeeds.push({
              id: `docx-heading-${navSeeds.length}`,
              label: text,
              level,
              paragraphIndex,
            })
          }
        },
        paragraphIndexRef,
      )
    }
  }

  return { words, navSeeds }
}

function getSelectionFromDom(root: HTMLElement): Range | null {
  const domSelection = root.ownerDocument.getSelection()
  if (domSelection === null || domSelection.rangeCount === 0) {
    return null
  }

  const anchorNode = domSelection.anchorNode
  const focusNode = domSelection.focusNode
  if (anchorNode === null || focusNode === null) {
    return null
  }

  if (!root.contains(anchorNode) || !root.contains(focusNode)) {
    return null
  }

  const anchor = domPointToPosition(anchorNode, domSelection.anchorOffset, root.ownerDocument)
  const focus = domPointToPosition(focusNode, domSelection.focusOffset, root.ownerDocument)

  return anchor && focus ? { anchor, focus } : null
}

function syncSelectionToDom(root: HTMLElement, range: Range | null): void {
  if (range === null) {
    return
  }

  const anchor = positionToDomRange(range.anchor, root)
  const focus = positionToDomRange(range.focus, root)
  if (anchor === null || focus === null) {
    return
  }

  const selection = root.ownerDocument.getSelection()
  if (selection === null) {
    return
  }

  const domRange = root.ownerDocument.createRange()
  domRange.setStart(anchor.node, anchor.offset)
  domRange.setEnd(focus.node, focus.offset)
  selection.removeAllRanges()
  selection.addRange(domRange)
}

function getReplaceValue(root: HTMLElement | null): string {
  if (root === null) {
    return ''
  }

  const input = root.querySelector<HTMLInputElement>('.docx-find__input[aria-label="Replace"]')
  return input?.value ?? ''
}

function DocxEditor({
  bundle,
  file,
  containerRef,
  onBundleChange,
}: {
  bundle: DocxBundle
  file: Extract<ViewerProps['file'], { kind: 'binary' }>
  containerRef: React.RefObject<HTMLDivElement | null>
  onBundleChange: (next: DocxBundle) => void
}) {
  const editorRootRef = useRef<HTMLDivElement | null>(null)
  const historyRef = useRef(new History())
  const fontResolver = useMemo(() => createFontResolver(), [])
  const [documentModel, setDocumentModel] = useState(bundle.document)
  const [range, setRange] = useState<Range | null>(null)
  const [pages, setPages] = useState<ReadonlyArray<Page> | null>(null)
  const [paginationProgress, setPaginationProgress] = useState<PaginationProgress | null>(null)
  const [paginationError, setPaginationError] = useState<string | null>(null)
  const [savePath, setSavePath] = useState(file.path)
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const [findOptions, setFindOptions] = useState<FindOptions>(DEFAULT_FIND_OPTIONS)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [commentsPaneOpen, setCommentsPaneOpen] = useState(bundle.document.comments.size > 0)
  const [resolvedCommentIds, setResolvedCommentIds] = useState<ReadonlySet<string>>(new Set())
  // DXE-11/D17 — seeded from the source document's real `word/settings.xml`
  // `<w:trackChanges/>` setting (parsed by `docx/parser/settings.ts`) rather
  // than always starting `false`, and every toggle writes back into
  // `bundle.settings` (see the `toggle-track-changes` handler below) so
  // `saveDocx` persists it via `writeSettingsXml` instead of losing it to
  // the passthrough copy of the original part.
  const [trackChangesEnabled, setTrackChangesEnabled] = useState(bundle.settings?.trackChanges ?? false)
  // DXE-11 — the context threaded into Input.ts so a genuine typed
  // insertion/deletion (never History's own undo/redo replay — see
  // `InputContext`'s own doc comment) is recorded as `w:ins`/`w:del`.
  // `author` is a static "Atlas" rather than the OS user's real name: Word
  // itself falls back to a generic label when it can't resolve one, and
  // wiring a real OS-username IPC channel would mean adding one in
  // electron/main.cjs + preload.cjs, both outside this task's editor-only
  // ownership boundary (see the wave plan) and shared with other in-flight
  // branches. Recomputed fresh on every keystroke (not memoized) so `date`
  // is always "now," matching what Word itself stamps per edit.
  const getTrackChanges = useCallback((): TrackChangesContext | undefined => {
    return trackChangesEnabled
      ? { enabled: true, author: 'Atlas', date: new Date().toISOString() }
      : undefined
  }, [trackChangesEnabled])
  const [spellCheckEnabled, setSpellCheckEnabled] = useState(true)
  const composition = useComposition()
  const spellCheck = useSpellCheck()

  // P1.1/SHELL-03/DXE-09 — document-session capability contract: report dirty
  // whenever documentModel diverges (by reference — every edit replaces it
  // immutably) from the last-saved-or-loaded snapshot, and register our own
  // save() so App.tsx's global Ctrl+S/Save can reach it when DOCX is active.
  //
  // `lastSavedDocument` is state, not a ref: dirty is derived from it and the
  // *current* `documentModel` together in one effect below, so that if the
  // user keeps typing while an async save() is still in flight, the edits
  // made after the save started are correctly still reported dirty once the
  // save resolves — the effect re-reads `documentModel` fresh on every
  // render rather than trusting whatever value save()'s own closure captured
  // when it started.
  const [lastSavedDocument, setLastSavedDocument] = useState(bundle.document)
  const setDirty = useSetViewerDirty()
  const registerSave = useRegisterViewerSave()

  const matches = useMemo(() => {
    if (findQuery.length === 0) {
      return []
    }

    return findAll(documentModel, findQuery, findOptions)
  }, [documentModel, findOptions, findQuery])

  const currentMatchIndex = useMemo(() => getCurrentMatchIndex(matches, range), [matches, range])

  const toolbarState = useMemo(
    () => createToolbarState(documentModel, range, { spellCheck: spellCheckEnabled, trackChanges: trackChangesEnabled }),
    [documentModel, range, spellCheckEnabled, trackChangesEnabled],
  )

  const availableStyles = useMemo(() => {
    return Array.from(documentModel.styles.values())
      .filter(style => style.type === 'paragraph')
      .map(style => ({ id: style.id, name: style.name ?? style.id }))
  }, [documentModel.styles])

  // DXE-25 — save-failure feedback used to be cleared by commitState, which
  // runs on every single edit, so the banner vanished the instant the user
  // typed the next character rather than staying up until the user dismissed
  // it or a save actually succeeded. commitState no longer touches saveError
  // at all; handleSave (below) is the only place that sets or clears it.
  const commitState = useCallback((nextDocument: DocxDocument, nextRange: Range | null) => {
    setDocumentModel(nextDocument)
    setRange(nextRange)
  }, [])

  useEffect(() => {
    setCommentsPaneOpen(bundle.document.comments.size > 0)
    setResolvedCommentIds(new Set())
  }, [bundle.document.comments.size])

  const applyResult = useCallback(
    (result: { document: DocxDocument; range: Range | null } | null): boolean => {
      if (result === null) {
        return false
      }

      commitState(result.document, result.range)
      return true
    },
    [commitState],
  )

  const applyEditorCommand = useCallback(
    (command: Command): boolean => {
      try {
        const result = applyCommand(documentModel, command)
        historyRef.current.push(result.inverse)
        commitState(result.document, result.range ?? range)
        return true
      } catch {
        return false
      }
    },
    [commitState, documentModel, range],
  )

  // DXE-02/DXE-17 — a whole command batch (paste, Replace All) is wrapped in
  // one `composite` command instead of applied+pushed one at a time: if any
  // sub-command throws, `applyCommand`'s own composite handling means NOTHING
  // in this batch was applied and nothing was pushed to history (previously,
  // commands that succeeded before a later failure had already been pushed
  // onto the undo stack even though commitState — and so the visible document
  // — never picked up the change, corrupting the undo stack). A successful
  // batch pushes exactly one inverse, so Ctrl+Z undoes the whole batch in one
  // step.
  const applyEditorCommands = useCallback(
    (commands: ReadonlyArray<Command>, nextRange: Range | null): boolean => {
      if (commands.length === 0) {
        return false
      }

      try {
        const batch: Command = commands.length === 1 ? commands[0] : { kind: 'composite', commands }
        const result = applyCommand(documentModel, batch)
        historyRef.current.push(result.inverse)
        // Prefer the caller's own cursor computation (e.g. buildPasteCommands'
        // finalCursor) over the composite's own `range` (the LAST
        // sub-command's natural result, e.g. an apply-run-format's selection)
        // — callers here already know the semantically-correct end position.
        commitState(result.document, nextRange ?? result.range ?? null)
        return true
      } catch {
        return false
      }
    },
    [commitState, documentModel],
  )

  const syncRangeFromDom = useCallback(() => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    const nextRange = getSelectionFromDom(root)
    setRange(current => (rangeEquals(current, nextRange) ? current : nextRange))
  }, [])

  const mediaResolver = useArchiveMediaResolver(bundle.rawArchive, bundle.relationships)

  const handleInsertImage = useCallback(async () => {
    const electronApi = window.electronAPI
    const picker = electronApi?.image?.pick
    if (picker === undefined) {
      setSaveError('Image picker is unavailable in this environment.')
      return
    }

    const focus = range?.focus ?? null
    if (focus === null) {
      setSaveError('Place the cursor in the document before inserting an image.')
      return
    }

    try {
      const result = await picker()
      if (result.cancelled) {
        return
      }

      const bytes = result.bytes instanceof Uint8Array ? result.bytes : new Uint8Array(result.bytes)
      const mime: ImageMimeType = result.mime
      // DXE-18 — decode the image's own pixel dimensions (falls back to the
      // old fixed 200x150pt box when unavailable) instead of always forcing
      // a hardcoded size, so a portrait photo isn't squashed into a 4:3 box.
      const { widthPt, heightPt } = await decodeImageNaturalSizePt(bytes, mime)

      // P1.6/DXE-01 — operate on the *live* documentModel, not the stale
      // `bundle` prop (which still holds whatever was loaded from disk):
      // otherwise inserting an image silently reverts every edit made since
      // file-load, since insertImageIntoBundle appends to `bundle.document`.
      const insertResult = insertImageIntoBundle({ ...bundle, document: documentModel }, focus, {
        bytes,
        mime,
        suggestedName: result.suggestedName,
        widthPt,
        heightPt,
      })

      // DXE-18 — route through History like every other edit so Ctrl+Z
      // removes the inserted image again.
      historyRef.current.push(insertResult.inverse)
      onBundleChange(insertResult.bundle)
      commitState(insertResult.document, insertResult.range)
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error))
    }
  }, [bundle, commitState, documentModel, onBundleChange, range])

  const handleInsertHyperlink = useCallback(() => {
    const selection = range
    if (selection === null) {
      setSaveError('Select text or place the cursor before inserting a hyperlink.')
      return
    }

    const url = window.prompt('Enter a URL', 'https://')
    if (url === null) {
      return
    }
    const trimmedUrl = url.trim()
    if (trimmedUrl.length === 0) {
      return
    }

    try {
      const insertResult = insertHyperlinkIntoBundle({ ...bundle, document: documentModel }, selection, trimmedUrl)
      historyRef.current.push(insertResult.inverse)
      onBundleChange(insertResult.bundle)
      commitState(insertResult.document, insertResult.range)
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error))
    }
  }, [bundle, commitState, documentModel, onBundleChange, range])

  const handleToggleList = useCallback(
    (kind: 'bullet' | 'number') => {
      const command = toolbarToCommand(
        { kind: kind === 'bullet' ? 'toggle-bullet-list' : 'toggle-numbered-list' },
        range,
        documentModel,
      )
      if (command === null || command.kind !== 'insert-list') {
        setSaveError('Select a paragraph before toggling a list.')
        return
      }

      try {
        const nextBundle = ensureListNumbering({ ...bundle, document: documentModel }, command.numId, kind)
        const result = applyCommand(nextBundle.document, command)
        historyRef.current.push(result.inverse)
        onBundleChange({ ...nextBundle, document: result.document })
        commitState(result.document, result.range ?? range)
      } catch (error) {
        setSaveError(error instanceof Error ? error.message : String(error))
      }
    },
    [bundle, commitState, documentModel, onBundleChange, range],
  )

  const handlePaste = useCallback(
    (event: ReactClipboardEvent<HTMLDivElement>) => {
      const clipboard = event.clipboardData
      if (clipboard === null) {
        return
      }

      const focus = range?.focus ?? null
      if (focus === null) {
        return
      }

      const html = clipboard.getData('text/html')
      const text = clipboard.getData('text/plain')

      const paragraphs =
        html.length > 0 ? htmlToParagraphs(html) : textToParagraphs(text)

      if (paragraphs.length === 0) {
        return
      }

      event.preventDefault()

      const { commands, finalCursor } = buildPasteCommands(paragraphs, focus, range)
      const finalRange: Range = { anchor: finalCursor, focus: finalCursor }
      applyEditorCommands(commands, finalRange)
    },
    [applyEditorCommands, range],
  )

  const handleBeforeInputEvent = useCallback(
    (event: FormEvent<HTMLDivElement>) => {
      const nativeEvent = event.nativeEvent as InputEvent

      // D12/DXE-03 — this contentEditable's DOM is entirely React-rendered
      // from `documentModel`, so a native mutation must never be allowed to
      // touch it directly: previously preventDefault() only ran when our own
      // command *succeeded*, so an edge case our model doesn't support yet
      // (typing inside a hyperlink used to be one) would fail silently while
      // the browser's native contentEditable behavior still mutated the DOM
      // out from under the model. Always prevent the native mutation first,
      // then apply our own model-level command if we have one.
      event.preventDefault()

      if (composition.shouldSwallow(nativeEvent)) {
        return
      }

      const applied = applyResult(
        handleBeforeInput(nativeEvent, {
          document: documentModel,
          range,
          history: historyRef.current,
          trackChanges: getTrackChanges(),
        }),
      )

      // DXE-20 — record that this text already landed in the model so a
      // trailing compositionend (some browsers fire beforeinput(insertText)
      // with the final composed text just before it) doesn't insert it again.
      if (applied && nativeEvent.inputType === 'insertText' && composition.state.active && typeof nativeEvent.data === 'string') {
        composition.markApplied(nativeEvent.data)
      }
    },
    [applyResult, composition, documentModel, getTrackChanges, range],
  )

  // Wave F.4 — refs for handlers defined later in this component, so the
  // keyboard-shortcut callback below can reference them without TDZ errors.
  const handlePrintRef = useRef<() => void>(() => {})
  const handleToolbarCommandRef = useRef<(cmd: ToolbarCommand) => void>(() => {})
  const handleSaveRef = useRef<() => Promise<boolean>>(async () => false)

  const handleKeyDownEvent = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      const ctrl = event.ctrlKey || event.metaKey
      const shift = event.shiftKey
      const lowerKey = event.key.toLowerCase()

      // Wave F.4 — MS Word keyboard parity.  Intercept the parity-only
      // shortcuts here BEFORE falling through to the editor's handleKeyDown
      // (which already covers Ctrl+B/I/U/Z/Y/A and motion keys).

      // Ctrl+F — Find
      if (ctrl && !shift && lowerKey === 'f') {
        event.preventDefault()
        setFindOpen(true)
        return
      }

      // Ctrl+H — Replace (same panel as Find; FindReplace component shows both)
      if (ctrl && !shift && lowerKey === 'h') {
        event.preventDefault()
        setFindOpen(true)
        return
      }

      // Ctrl+S — save directly (DXE-07/RUN-03/SHELL-10: this used to swallow
      // the key and do nothing, since Ctrl+S has no browser default to
      // prevent here anyway — Save is otherwise only reachable via the
      // button or App.tsx's global shortcut routed through the shared
      // document-session `save()` contract).
      if (ctrl && !shift && lowerKey === 's') {
        event.preventDefault()
        void handleSaveRef.current()
        return
      }

      // Ctrl+P — Print
      if (ctrl && !shift && lowerKey === 'p') {
        event.preventDefault()
        handlePrintRef.current()
        return
      }

      // F7 — toggle spell check
      if (event.key === 'F7' && !ctrl) {
        event.preventDefault()
        handleToolbarCommandRef.current({ kind: 'toggle-spell-check' })
        return
      }

      // Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J — alignment
      if (ctrl && !shift && (lowerKey === 'l' || lowerKey === 'e' || lowerKey === 'r' || lowerKey === 'j')) {
        const align =
          lowerKey === 'l' ? 'left' : lowerKey === 'e' ? 'center' : lowerKey === 'r' ? 'right' : 'justify'
        event.preventDefault()
        handleToolbarCommandRef.current({ kind: 'set-alignment', align })
        return
      }

      // Ctrl+Shift+L — bullet list
      if (ctrl && shift && lowerKey === 'l') {
        event.preventDefault()
        handleToolbarCommandRef.current({ kind: 'toggle-bullet-list' })
        return
      }

      // Ctrl+1 / Ctrl+2 / Ctrl+5 — line spacing (1, 2, 1.5)
      if (ctrl && !shift && (event.key === '1' || event.key === '2' || event.key === '5')) {
        const spacing = event.key === '1' ? 1 : event.key === '2' ? 2 : 1.5
        event.preventDefault()
        handleToolbarCommandRef.current({ kind: 'set-line-spacing', spacing })
        return
      }

      // Ctrl+K — insert hyperlink
      if (ctrl && !shift && lowerKey === 'k') {
        event.preventDefault()
        handleToolbarCommandRef.current({ kind: 'insert-hyperlink' })
        return
      }

      const applied = applyResult(
        handleKeyDown(event.nativeEvent, {
          document: documentModel,
          range,
          history: historyRef.current,
          trackChanges: getTrackChanges(),
        }),
      )

      // D12/DXE-03 — Backspace/Delete/Enter/Tab always mutate content when
      // the browser handles them natively. Input.ts's handleKeyDown already
      // routes all four through the same try/catch'd command pipeline as
      // beforeinput, so a failure here (e.g. an edit our model doesn't
      // support yet) must still block the native keydown default — otherwise
      // the browser performs its own contentEditable edit on a DOM the model
      // never sees, leaving the two silently out of sync. Every other key
      // (navigation, undo/redo, format toggles) keeps the old behavior of
      // only preventing default when we actually handled it.
      if (applied || MUTATING_KEYS.has(event.key)) {
        event.preventDefault()
      }
    },
    [applyResult, documentModel, getTrackChanges, range],
  )

  const handleCompositionStart = useCallback((event: FormEvent<HTMLDivElement>) => {
    composition.handlers.onCompositionStart(event.nativeEvent as CompositionEvent)
  }, [composition.handlers])

  const handleCompositionUpdate = useCallback((event: FormEvent<HTMLDivElement>) => {
    composition.handlers.onCompositionUpdate(event.nativeEvent as CompositionEvent)
  }, [composition.handlers])

  const handleCompositionEnd = useCallback(
    (event: FormEvent<HTMLDivElement>) => {
      const committedText = composition.handlers.onCompositionEnd(event.nativeEvent as CompositionEvent)
      if (committedText === null) {
        return
      }

      applyResult(
        handleBeforeInput(new InputEvent('beforeinput', { inputType: 'insertText', data: committedText }), {
          document: documentModel,
          range,
          history: historyRef.current,
          trackChanges: getTrackChanges(),
        }),
      )
    },
    [applyResult, composition.handlers, documentModel, getTrackChanges, range],
  )

  const handleFind = useCallback((query: string, options: FindOptions) => {
    setFindQuery(query)
    setFindOptions(options)

    if (query.length === 0) {
      return
    }

    const nextMatch = findNext(documentModel, query, options, null)
    if (nextMatch !== null) {
      setRange(nextMatch.range)
    }
  }, [documentModel])

  const handleFindNext = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const nextMatch = findNext(documentModel, findQuery, findOptions, range?.focus ?? null)
    if (nextMatch !== null) {
      setRange(nextMatch.range)
    }
  }, [documentModel, findOptions, findQuery, range])

  const handleFindPrev = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const prevMatch = findPrev(documentModel, findQuery, findOptions, range?.anchor ?? null)
    if (prevMatch !== null) {
      setRange(prevMatch.range)
    }
  }, [documentModel, findOptions, findQuery, range])

  const handleReplace = useCallback(() => {
    if (findQuery.length === 0 || currentMatchIndex === null) {
      return
    }

    const replacement = getReplaceValue(containerRef.current)
    const match = matches[currentMatchIndex]
    const commands = buildReplaceCommands(match, replacement)
    const nextPos = {
      paragraphPath: match.range.anchor.paragraphPath,
      runIndex: match.range.anchor.runIndex,
      charOffset: match.range.anchor.charOffset + replacement.length,
    }

    applyEditorCommands(commands, { anchor: nextPos, focus: nextPos })
  }, [applyEditorCommands, containerRef, currentMatchIndex, findQuery, matches])

  const handleReplaceAll = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const replacement = getReplaceValue(containerRef.current)
    const commands = buildReplaceAllCommands(documentModel, findQuery, findOptions, replacement)
    applyEditorCommands(commands, range)
  }, [applyEditorCommands, containerRef, documentModel, findOptions, findQuery, range])

  const handleScrollToParagraph = useCallback((paragraphIndex: number) => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    const lines = root.querySelectorAll<HTMLElement>('.docx-page__line[data-paragraph-path]')
    const target = lines.item(paragraphIndex)
    target?.scrollIntoView({ behavior: 'smooth', block: 'center' })
  }, [])

  // DXE-08 — accepting a suggestion used to only call Electron's native
  // replaceMisspelling bridge, which mutates the contentEditable DOM
  // directly without touching documentModel at all, so the fix rendered
  // correctly until the next re-render and was silently lost on save. This
  // now builds a real delete+insert command pair and applies it through the
  // normal pipeline (undoable, and part of what gets saved), using the
  // current selection as a hint for which occurrence to fix when the word
  // appears more than once.
  const handleSpellReplace = useCallback(
    (misspelled: string, replacement: string) => {
      if (misspelled.length === 0 || replacement === misspelled) {
        return
      }

      const hint = range?.focus ?? null
      const replacementCommands = buildSpellCheckReplacement(documentModel, misspelled, replacement, hint)
      if (replacementCommands === null) {
        return
      }

      applyEditorCommands(replacementCommands.commands, replacementCommands.range)
      spellCheck.dismiss()
    },
    [applyEditorCommands, documentModel, range, spellCheck],
  )

  const handleAddToDictionary = useCallback(
    (word: string) => {
      void spellCheck.addToDictionary(word)
    },
    [spellCheck],
  )

  const handleAddComment = useCallback(() => {
    const selection = range
    if (selection === null) {
      setSaveError('Select text before adding a comment.')
      return
    }

    const text = window.prompt('Add comment', '')
    if (text === null) {
      return
    }

    try {
      const result = addCommentToDocument(documentModel, selection, text, 'Atlas')
      commitState(result.document, selection)
      setCommentsPaneOpen(true)
      setResolvedCommentIds((current) => {
        const next = new Set(current)
        next.delete(result.commentId)
        return next
      })
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error))
    }
  }, [commitState, documentModel, range])

  const handleReplyToComment = useCallback(
    (commentId: string) => {
      const text = window.prompt('Reply', '')
      if (text === null) {
        return
      }

      commitState(replyToComment(documentModel, commentId, text, 'Atlas'), range)
      setCommentsPaneOpen(true)
      setResolvedCommentIds((current) => {
        const next = new Set(current)
        next.delete(commentId)
        return next
      })
    },
    [commitState, documentModel, range],
  )

  const handleResolveComment = useCallback((commentId: string) => {
    setResolvedCommentIds((current) => new Set(current).add(commentId))
  }, [])

  const handleDeleteComment = useCallback(
    (commentId: string) => {
      commitState(deleteCommentFromDocument(documentModel, commentId), range)
      setResolvedCommentIds((current) => {
        const next = new Set(current)
        next.delete(commentId)
        return next
      })
    },
    [commitState, documentModel, range],
  )

  const commentsDocument = useMemo<DocxDocument>(() => {
    if (resolvedCommentIds.size === 0) {
      return documentModel
    }

    const filtered = new Map(
      [...documentModel.comments.entries()].filter(([id, comment]) => {
        if (resolvedCommentIds.has(id)) {
          return false
        }
        return comment.parentId === undefined || !resolvedCommentIds.has(comment.parentId)
      }),
    )

    return {
      ...documentModel,
      comments: filtered,
    }
  }, [documentModel, resolvedCommentIds])

  const handleToolbarCommand = useCallback(
    (toolbarCommand: ToolbarCommand) => {
      if (toolbarCommand.kind === 'open-find-replace') {
        setFindOpen(true)
        return
      }

      if (toolbarCommand.kind === 'undo') {
        applyResult(
          handleKeyDown(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true }), {
            document: documentModel,
            range,
            history: historyRef.current,
          }),
        )
        return
      }

      if (toolbarCommand.kind === 'redo') {
        applyResult(
          handleKeyDown(new KeyboardEvent('keydown', { key: 'y', ctrlKey: true }), {
            document: documentModel,
            range,
            history: historyRef.current,
          }),
        )
        return
      }

      // DXE-12 — spell check and track changes are real, local toggles now
      // (rather than always-on / never-wired): both flip a piece of
      // component state the render below reads (contentEditable's own
      // spellCheck attribute; the ToolbarState.trackChanges shown in the
      // Review tab).
      if (toolbarCommand.kind === 'toggle-spell-check') {
        setSpellCheckEnabled((enabled) => !enabled)
        return
      }

      if (toolbarCommand.kind === 'toggle-track-changes') {
        setTrackChangesEnabled((enabled) => {
          const next = !enabled
          // DXE-11/D17 — persist into the bundle so saveDocx (which reads
          // `bundle.settings`, not this component's local state) writes the
          // change into word/settings.xml instead of it only ever affecting
          // this session's toolbar display.
          onBundleChange({ ...bundle, settings: { trackChanges: next } })
          return next
        })
        return
      }

      if (toolbarCommand.kind === 'open-comments-pane') {
        setCommentsPaneOpen(true)
        return
      }

      if (toolbarCommand.kind === 'insert-comment') {
        handleAddComment()
        return
      }

      if (toolbarCommand.kind === 'insert-image') {
        void handleInsertImage()
        return
      }

      if (toolbarCommand.kind === 'insert-hyperlink') {
        handleInsertHyperlink()
        return
      }

      if (toolbarCommand.kind === 'toggle-bullet-list') {
        handleToggleList('bullet')
        return
      }

      if (toolbarCommand.kind === 'toggle-numbered-list') {
        handleToggleList('number')
        return
      }

      const command = toolbarToCommand(toolbarCommand, range, documentModel)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [
      applyEditorCommand,
      applyResult,
      bundle,
      documentModel,
      handleAddComment,
      handleInsertHyperlink,
      handleInsertImage,
      handleToggleList,
      onBundleChange,
      range,
    ],
  )

  // P1.1 — returns whether the save actually succeeded, so both the shared
  // document-session `save()` contract (App.tsx's global Ctrl+S/Save and the
  // unsaved-changes confirmation dialog) and the Save button here can tell.
  // DXE-25 — a save failure sets saveError, which now only ever gets cleared
  // by a later save attempt (right here, at the start) or a successful save
  // — never by commitState on the next keystroke (see commitState above), so
  // the banner stays visible until the user either fixes the problem and
  // saves again or the save actually succeeds.
  const handleSaveInternal = useCallback(
    async (options?: { readonly forceDialog?: boolean }): Promise<boolean> => {
      setSaveError(null)

      try {
        const nextBytes = await saveDocx({
          ...bundle,
          document: documentModel,
        })

        const result = await window.electronAPI?.saveBinaryFile?.({
          content: nextBytes,
          suggestedName: getSuggestedFileName(savePath),
          // DXE-24 — omitting `existingPath` makes the main process show the
          // save dialog instead of silently overwriting `savePath`, which is
          // exactly "Save As".
          ...(options?.forceDialog ? {} : { existingPath: savePath }),
          filters: DOCX_SAVE_FILTERS,
        })

        if (result?.saved && result.path) {
          setSavePath(result.path)
        }

        if (!result?.saved) {
          setSaveError(result?.error ?? 'Save was cancelled or unavailable.')
          return false
        }

        // Record *what was actually written* (this closure's own `documentModel`
        // snapshot, taken when the save started) as the new saved baseline. We
        // deliberately do NOT also call `setDirty(false)` here: if the user kept
        // editing while the `await`s above were in flight, `documentModel` may
        // already have moved on since this snapshot, and the effect below
        // recomputes dirty from whatever the *current* render's `documentModel`
        // is once this state update lands — never from this stale closure.
        setLastSavedDocument(documentModel)
        return true
      } catch (error) {
        // RUN-14 — wrap JSZip/fast-xml-parser/DocxSaveError internals in a
        // friendly, actionable message instead of showing the raw exception.
        setSaveError(friendlyDocxErrorMessage(error, 'save'))
        return false
      }
    },
    [bundle, documentModel, savePath],
  )

  const handleSave = useCallback((): Promise<boolean> => handleSaveInternal(), [handleSaveInternal])
  const handleSaveAs = useCallback(
    (): Promise<boolean> => handleSaveInternal({ forceDialog: true }),
    [handleSaveInternal],
  )

  useEffect(() => {
    handleSaveRef.current = handleSave
  }, [handleSave])

  // P1.1/SHELL-03/DXE-09 — publish dirty state and this viewer's save
  // implementation to the shared document-session contract. Every commitState
  // call replaces documentModel with a new object, so a plain reference
  // compare against the last-saved-or-loaded snapshot is exactly "has this
  // document changed since load/save". Deriving this from both pieces of
  // state together (rather than reaching into a ref from inside handleSave)
  // is what makes it correct even when an edit lands while a save is still
  // in flight: this effect always compares against the render's *current*
  // documentModel, never a snapshot captured before the save resolved.
  useEffect(() => {
    setDirty(documentModel !== lastSavedDocument)
  }, [documentModel, lastSavedDocument, setDirty])

  useEffect(() => {
    registerSave(handleSave)
    return () => registerSave(null)
  }, [registerSave, handleSave])

  // P2.1/SHELL-08/SHELL-09/UX-03/RUN-02 — claim every combo this viewer's own
  // `handleKeyDownEvent`/commands.ts already recognize at the *active-viewer*
  // precedence tier, so shell-global shortcuts (sidebar toggle, Export menu,
  // print/export) never see them — independent of exactly where focus is
  // within the viewer, not just whether the contentEditable itself is
  // focused (e.g. after clicking a Toolbar button, which isn't itself the
  // contentEditable). Detection-only: it never performs the actual
  // edit/print/save itself (that stays solely in
  // `handleKeyDownEvent`/commands.ts, reached through React's own bubble
  // phase, which always runs first when the contentEditable has focus) —
  // this only prevents the *shell* from also reacting to the same keystroke.
  //
  // Deliberately bails out for a real `<input>`/`<textarea>` target (the
  // Find/Replace panel's own search boxes, rendered as siblings of the
  // contentEditable, not inside it) so their native text-editing shortcuts
  // — Ctrl+A select-all, Ctrl+Z/Y undo/redo the typed text, etc. — keep
  // working. Only the document's contentEditable surface (a `<div>`, never
  // an `<input>`/`<textarea>`) and non-field focus targets (e.g. a Toolbar
  // button) should have these combos reserved for document commands.
  useViewerShortcuts(
    useCallback((event) => {
      const ctrl = event.ctrlKey || event.metaKey
      if (!ctrl) return false
      if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
        return false
      }
      const key = event.key.toLowerCase()
      const isReservedCombo =
        key === 'f' || key === 'h' || key === 's' || key === 'p' || key === 'k' ||
        key === 'l' || key === 'e' || key === 'r' || key === 'j' ||
        key === '1' || key === '2' || key === '5' ||
        key === 'z' || key === 'y' || key === 'a' || key === 'b' || key === 'i' || key === 'u'
      if (!isReservedCombo) return false
      event.preventDefault()
      return true
    }, []),
  )

  const handlePrint = useCallback(() => {
    if (typeof window === 'undefined') {
      return
    }
    // Print uses native browser print dialog. The print stylesheet in
    // viewer-docx.css hides toolbar/comments/save UI and forces a page
    // break after each .docx-page so output matches on-screen pagination.
    document.body.classList.add('atlas-printing')
    try {
      window.print()
    } finally {
      document.body.classList.remove('atlas-printing')
    }
  }, [])

  useEffect(() => {
    handlePrintRef.current = handlePrint
  }, [handlePrint])

  useEffect(() => {
    handleToolbarCommandRef.current = handleToolbarCommand
  }, [handleToolbarCommand])

  useEffect(() => {
    let cancelled = false

    setPaginationError(null)
    setPaginationProgress(null)

    void (async () => {
      // Ensure the browser has actually downloaded and registered the bundled
      // substitute fonts BEFORE we measure-and-paint. Otherwise pagination
      // computes widths from real TTF metrics while the DOM still renders with
      // a fallback font, producing accumulated drift -> mid-line gaps and
      // right-edge clipping.
      await preloadDocxFonts()

      if (cancelled) {
        return
      }

      try {
        const nextPages = await paginate({
          document: documentModel,
          fontResolver,
          theme: bundle.theme,
          onProgress: progress => {
            if (!cancelled) {
              setPaginationProgress(progress)
            }
          },
          shouldCancel: () => cancelled,
        })

        if (!cancelled) {
          setPages(nextPages)
          setPaginationProgress(null)
        }
      } catch (error) {
        if (error instanceof PaginationCancelledError) {
          return
        }
        if (!cancelled) {
          // DXE-27 — the failure is already surfaced to the user via
          // paginationError (rendered below); no need to also log it.
          setPaginationError(error instanceof Error ? error.message : String(error))
          setPaginationProgress(null)
        }
      }
    })()

    return () => {
      cancelled = true
    }
  }, [documentModel, fontResolver, bundle.theme])

  useLayoutEffect(() => {
    const root = editorRootRef.current
    if (root !== null) {
      syncSelectionToDom(root, range)
    }
  }, [pages, range])

  const pageCount = useMemo(
    () => pages?.length ?? (documentModel.sections.length || 1),
    [documentModel.sections.length, pages],
  )

  return (
    <div className="docx-viewer__editor-shell">
      <div className="docx-viewer__topbar">
        <Toolbar state={toolbarState} onCommand={handleToolbarCommand} availableStyles={availableStyles} />
        <button className="docx-viewer__save-button" type="button" onClick={() => void handleSave()}>
          <Save size={16} aria-hidden="true" />
          <span>Save</span>
        </button>
        <button
          className="docx-viewer__save-as-button"
          type="button"
          onClick={() => void handleSaveAs()}
          aria-label="Save As"
        >
          <span>Save As…</span>
        </button>
        <button
          className="docx-viewer__print-button"
          type="button"
          onClick={handlePrint}
          aria-label="Print document"
        >
          <Printer size={16} aria-hidden="true" />
          <span>Print</span>
        </button>
      </div>
      <FindReplace
        open={findOpen}
        onClose={() => setFindOpen(false)}
        onFind={handleFind}
        onFindNext={handleFindNext}
        onFindPrev={handleFindPrev}
        onReplace={handleReplace}
        onReplaceAll={handleReplaceAll}
        matchCount={matches.length}
        currentMatchIndex={currentMatchIndex}
      />
      <div className="docx-viewer__meta">Pages: {pageCount}</div>
      <div className="docx-viewer__workspace">
        <div
          ref={editorRootRef}
          className="docx-viewer__surface"
          role="textbox"
          aria-multiline="true"
          aria-label="Document editor"
          contentEditable
          suppressContentEditableWarning
          spellCheck={spellCheckEnabled}
          onBeforeInput={handleBeforeInputEvent}
          onKeyDown={handleKeyDownEvent}
          onMouseUp={() => syncRangeFromDom()}
          onKeyUp={() => syncRangeFromDom()}
          onFocus={() => syncRangeFromDom()}
          onCompositionStart={handleCompositionStart}
          onCompositionUpdate={handleCompositionUpdate}
          onCompositionEnd={handleCompositionEnd}
          onPaste={handlePaste}
        >
          {pages !== null ? (
            <MediaContext.Provider value={mediaResolver}>
              <PageStack
                pages={pages}
                zoom={1}
                document={documentModel}
                theme={bundle.theme}
                relationships={bundle.relationships}
              />
            </MediaContext.Provider>
          ) : paginationError !== null ? (
            <div className="docx-viewer__loading docx-viewer__loading--error" role="alert">
              <div className="docx-viewer__loading-title">Failed to lay out document</div>
              <div className="docx-viewer__loading-detail">{paginationError}</div>
            </div>
          ) : (
            <div className="docx-viewer__loading" role="status" aria-live="polite">
              <div className="docx-viewer__loading-spinner" aria-hidden="true" />
              <div className="docx-viewer__loading-title">Laying out document…</div>
              {paginationProgress !== null && paginationProgress.totalBlocks > 0 ? (
                <>
                  <div
                    className="docx-viewer__loading-bar"
                    role="progressbar"
                    aria-valuemin={0}
                    aria-valuemax={paginationProgress.totalBlocks}
                    aria-valuenow={paginationProgress.completedBlocks}
                  >
                    <div
                      className="docx-viewer__loading-bar-fill"
                      style={{
                        width: `${Math.min(
                          100,
                          Math.round(
                            (paginationProgress.completedBlocks / paginationProgress.totalBlocks) * 100,
                          ),
                        )}%`,
                      }}
                    />
                  </div>
                  <div className="docx-viewer__loading-detail">
                    {paginationProgress.completedBlocks} / {paginationProgress.totalBlocks} blocks
                  </div>
                </>
              ) : null}
            </div>
          )}
        </div>
        {commentsPaneOpen || documentModel.comments.size > 0 ? (
          <CommentsPane
            document={commentsDocument}
            onScrollToParagraph={handleScrollToParagraph}
            onAddComment={handleAddComment}
            onReply={handleReplyToComment}
            onResolve={handleResolveComment}
            onDelete={handleDeleteComment}
          />
        ) : null}
      </div>
      {saveError !== null ? (
        <div className="docx-viewer__error" role="alert">
          <span>{saveError}</span>
          <button
            type="button"
            className="docx-viewer__error-dismiss"
            onClick={() => setSaveError(null)}
            aria-label="Dismiss error"
          >
            <X size={14} aria-hidden="true" />
          </button>
        </div>
      ) : null}
      {spellCheck.state !== null ? (
        <SpellCheckMenu
          word={spellCheck.state.word}
          suggestions={spellCheck.state.suggestions}
          x={spellCheck.state.x}
          y={spellCheck.state.y}
          onReplace={handleSpellReplace}
          onAddToDictionary={handleAddToDictionary}
          onDismiss={spellCheck.dismiss}
        />
      ) : null}
    </div>
  )
}

function DocxViewerBase({ file }: ViewerProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [bundle, setBundle] = useState<DocxBundle | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const setNavItems = useSetNavItems()
  const setStats = useSetViewerStats()

  useEffect(() => {
    let cancelled = false

    // Runs synchronously up to the first `await`, so this reset lands in the
    // same tick as the effect itself — identical timing to setting state
    // directly in the effect body, but nested in a callback so it synchronizes
    // with the external `loadDocx` call rather than reading as derivable state.
    void (async () => {
      setBundle(null)
      setNavItems([])
      setStats(null)
      setErrorMessage(null)

      if (file.kind === 'text') {
        return
      }

      try {
        const nextBundle = await loadDocx(normalizeDocxBuffer(file.content))

        if (!cancelled) {
          setBundle(nextBundle)
        }
      } catch (error) {
        if (!cancelled) {
          setBundle(null)
          setNavItems([])
          setStats(null)
          // RUN-14 — see friendlyDocxErrorMessage's own doc comment.
          setErrorMessage(friendlyDocxErrorMessage(error, 'open'))
        }
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file, setNavItems, setStats])

  const documentModel = bundle?.document ?? null

  const metrics = useMemo(() => {
    return documentModel ? collectDocxMetrics(documentModel) : null
  }, [documentModel])

  const navItems = useMemo<ReadonlyArray<NavItem>>(
    () =>
      metrics === null
        ? []
        : metrics.navSeeds.map(seed => ({
            id: seed.id,
            label: seed.label,
            level: seed.level,
            onSelect: () => {
              containerRef.current
                ?.querySelectorAll<HTMLElement>('.docx-page__line[data-paragraph-path]')
                .item(seed.paragraphIndex)
                ?.scrollIntoView({ behavior: 'smooth', block: 'start' })
            },
          })),
    [metrics],
  )

  useEffect(() => {
    setNavItems(navItems)
  }, [navItems, setNavItems])

  useEffect(() => {
    if (documentModel === null || metrics === null) {
      return
    }

    setStats({
      kind: 'document',
      words: metrics.words,
      pages: documentModel.sections.length || 1,
    })
  }, [documentModel, metrics, setStats])

  if (file.kind === 'text') {
    return <div className="docx-viewer docx-viewer--error">Unexpected text file routed to DocxViewer.</div>
  }

  if (errorMessage !== null) {
    // RUN-14 — errorMessage is already a complete, friendly sentence (see
    // friendlyDocxErrorMessage), so it's shown as-is rather than appended to
    // a generic "Failed to render DOCX:" prefix.
    return <div className="docx-viewer docx-viewer--error">{errorMessage}</div>
  }

  return (
    <div ref={containerRef} className="docx-viewer">
      {bundle !== null ? <DocxEditor bundle={bundle} file={file} containerRef={containerRef} onBundleChange={setBundle} /> : null}
    </div>
  )
}

export const DocxViewer = memo(DocxViewerBase)
