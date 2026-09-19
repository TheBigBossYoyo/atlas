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
  type MouseEvent as ReactMouseEvent,
} from 'react'
import { flushSync } from 'react-dom'

import { ListTree, PanelTop, Printer, RefreshCw, Save, SaveAll, X, ZoomIn, ZoomOut } from 'lucide-react'

import type { NavItem, ViewerProps } from '../formats/types'
import { loadDocx, saveDocx, type DocxBundle } from '../docx'
import { paginate, PaginationCancelledError } from '../docx/layout'
import type { Page, PaginationProgress } from '../docx/layout'
import type { FontResolver } from '../docx/layout/types'
import {
  buildMetrics,
  collectReferencedFontFamilies,
  FONT_FAMILIES,
  loadEmbeddedFonts,
  loadFontMetrics,
  parseTtf,
  resolveFontFamily,
  wrapWithCanvasAdvance,
  type EmbeddedFontFamily,
  type FontMetrics,
  type FontVariant,
} from '../docx/fonts'
import {
  collectBookmarkMaps,
  parseCoreProps,
  updateFields,
  updateTableOfContents,
  type FieldEvaluationContext,
} from '../docx/fields'
import { type Document as DocxDocument, type Paragraph, type ParagraphChild, type Run, type RunChild } from '../docx/model'
import { resolveParaProps } from '../docx/parser/cascade'
import { MediaContext, PageStack } from '../docx/render'
// D23-PERF-2 — imported from the concrete module, not the `docx/render`
// barrel: `DocxViewer.editor.test.tsx`/`.fields.test.tsx`/`.load.test.tsx`
// each `vi.mock('../../docx/render', ...)` with a bare-bones fake `PageStack`
// (they don't exercise real pagination/rendering), and going through the
// barrel here would make this a required export of that mock too.
// `findPagesForParagraphPath` is a pure function with no rendering
// dependency, so importing it directly sidesteps that entirely.
import { findPagesForParagraphPath } from '../docx/render/pageVirtualization'
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
  buildRichPasteCommands,
  bundleContextFor,
  friendlyDocxErrorMessage,
  htmlToParagraphs,
  htmlToPasteBlocks,
  textToParagraphs,
  toolbarToCommand,
  useComposition,
  useSpellCheck,
  type Command,
  type EnclosingTable,
  type FindOptions,
  type ImageMimeType,
  type PasteBlock,
  type Position,
  type Range,
  type TrackChangesContext,
  listHeaderFooterParts,
  buildHeaderFooterTextEdit,
  buildInsertHeaderFooterParagraph,
  buildRemoveHeaderFooterParagraph,
  type HeaderFooterKind,
} from '../docx/editor'
import { addCommentToDocument, deleteCommentFromDocument, replyToComment } from '../docx/editor/commentMutations'
import { comparePositions } from '../docx/editor/Selection'
import {
  domPointToPosition,
  paragraphRangeFromClientPoint,
  positionFromClientPoint,
  positionToDomRange,
  wordRangeFromClientPoint,
} from '../docx/editor/Cursor'
import { Toolbar } from '../docx/editor/toolbar/Toolbar'
import { TableEditMenuItems } from '../docx/editor/toolbar/TableEditMenuItems'
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

/** D23-PERF-2 — stable empty-Set reference for `pinnedPageIndices` when
 * there's nothing to pin, so an unchanged render doesn't hand `PageStack` a
 * fresh `Set` object every time (see `PageStack`'s own `memo` doc comment). */
const EMPTY_PINNED_PAGE_INDICES: ReadonlySet<number> = new Set()

const DEFAULT_FIND_OPTIONS: FindOptions = {
  caseSensitive: false,
  wholeWord: false,
  useRegex: false,
}

const DOCX_SAVE_FILTERS = [{ name: 'Word Documents', extensions: ['docx'] }]
// D24/DXL-19
const MIN_ZOOM = 0.25
const MAX_ZOOM = 3
const ZOOM_STEP = 0.1

function clampZoom(zoom: number): number {
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom))
}

/** D12/DXE-03 — keys whose native contentEditable behavior always mutates
 * content; see handleKeyDownEvent's preventDefault-safety comment. */
const MUTATING_KEYS: ReadonlySet<string> = new Set(['Backspace', 'Delete', 'Enter', 'Tab'])

/** USR-11 — always offered in the font picker, even before (or without)
 * system font enumeration. */
const COMMON_OFFICE_FONTS: ReadonlyArray<string> = [
  'Aptos', 'Arial', 'Calibri', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia',
  'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Impact', 'Lucida Console',
  'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana',
]

/** USR-07/USR-09 — navigation keys the model does not handle yet (vertical
 * movement needs layout), so the browser moves the caret and the DOM
 * selection is read back on keyup. Every other key keeps the model range. */
const NATIVE_NAVIGATION_KEYS: ReadonlySet<string> = new Set(['ArrowUp', 'ArrowDown', 'PageUp', 'PageDown'])

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

/**
 * DEFER-4 / DXP-13 — registers a document's own embedded fonts (already
 * de-obfuscated by `loadEmbeddedFonts`) via the same `FontFace` API used for
 * the bundled substitutes, under the font's real Word name so runs that
 * reference it paint with the actual embedded typeface instead of falling
 * back to a metric-substitute or the OS default. Unlike `preloadDocxFonts`
 * (a one-time, app-lifetime registration of the 5 bundled families), this
 * runs per document — callers are responsible for un-registering the
 * returned faces (via `unregisterEmbeddedFonts`) when the document changes,
 * since two different documents can embed two different fonts under the
 * same family name.
 *
 * A face that fails to parse/load (corrupt data, unsupported table format)
 * is skipped individually rather than failing the whole document — the
 * family's other faces, or the bundled-substitute fallback, still work.
 */
async function registerEmbeddedFonts(
  families: ReadonlyArray<EmbeddedFontFamily>,
): Promise<ReadonlyArray<FontFace>> {
  if (typeof document === 'undefined' || document.fonts === undefined || families.length === 0) {
    return []
  }

  const registered: FontFace[] = []
  for (const family of families) {
    for (const descriptor of FONT_VARIANT_DESCRIPTORS) {
      const data = family.faces[descriptor.variant]
      if (data === undefined) {
        continue
      }

      try {
        const face = new FontFace(family.name, toArrayBuffer(data), {
          weight: descriptor.weight,
          style: descriptor.style,
          display: 'block',
        })
        const loaded = await face.load()
        document.fonts.add(loaded)
        registered.push(loaded)
      } catch {
        // Corrupt/unsupported embedded font data: leave this face
        // unregistered so it falls back to the bundled substitute (or OS
        // default), same as an unresolvable font family today.
      }
    }
  }

  return registered
}

function unregisterEmbeddedFonts(faces: ReadonlyArray<FontFace>): void {
  if (typeof document === 'undefined' || document.fonts === undefined) {
    return
  }
  for (const face of faces) {
    document.fonts.delete(face)
  }
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  // `.slice()` always allocates a fresh, non-shared ArrayBuffer (unlike
  // `.buffer.slice()`, whose return type widens to ArrayBufferLike because the
  // source buffer could in principle be a SharedArrayBuffer).
  return bytes.slice().buffer
}

function createFontResolver(embeddedFonts: ReadonlyArray<EmbeddedFontFamily> = []): FontResolver {
  const embeddedByName = new Map<string, EmbeddedFontFamily>()
  for (const family of embeddedFonts) {
    embeddedByName.set(family.name.trim().toLowerCase(), family)
  }

  const cache = new Map<string, Promise<FontMetrics>>()

  return async (family: string, variant: FontVariant): Promise<FontMetrics> => {
    const embeddedFace = embeddedByName.get(family.trim().toLowerCase())?.faces[variant]
    if (embeddedFace !== undefined) {
      const cacheKey = `embedded:${family.trim().toLowerCase()}@${variant}`
      const existing = cache.get(cacheKey)
      if (existing !== undefined) {
        return existing
      }

      const promise = (async () => {
        try {
          const tables = parseTtf(toArrayBuffer(embeddedFace))
          return wrapWithCanvasAdvance(buildMetrics(tables), family, variant)
        } catch {
          // Corrupt embedded font data: fall back to a neutral synthetic
          // base measured under the requested name, same treatment an
          // unresolvable bundled family gets below.
          return wrapWithCanvasAdvance(DEFAULT_FONT_METRICS, family, variant)
        }
      })()
      cache.set(cacheKey, promise)
      return promise
    }

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

const SELECTION_HIGHLIGHT_NAME = 'docx-selection'

type HighlightRegistry = { set: (name: string, highlight: unknown) => void; delete: (name: string) => void }
type HighlightConstructor = new (...ranges: globalThis.Range[]) => unknown

/**
 * USR-06 — paints a model range without touching the DOM selection (and so
 * without stealing focus), using the CSS Custom Highlight API. `null` clears
 * it. A no-op where the API is unavailable (e.g. jsdom).
 */
function paintSelectionHighlight(root: HTMLElement | null, range: Range | null): void {
  const registry = (globalThis.CSS as unknown as { highlights?: HighlightRegistry } | undefined)?.highlights
  const HighlightCtor = (globalThis as unknown as { Highlight?: HighlightConstructor }).Highlight
  if (registry === undefined || HighlightCtor === undefined) {
    return
  }
  if (root === null || range === null) {
    registry.delete(SELECTION_HIGHLIGHT_NAME)
    return
  }
  const anchor = positionToDomRange(range.anchor, root)
  const focus = positionToDomRange(range.focus, root)
  if (anchor === null || focus === null) {
    registry.delete(SELECTION_HIGHLIGHT_NAME)
    return
  }
  const domRange = root.ownerDocument.createRange()
  domRange.setStart(anchor.node, anchor.offset)
  domRange.setEnd(focus.node, focus.offset)
  if (domRange.collapsed) {
    domRange.setStart(focus.node, focus.offset)
    domRange.setEnd(anchor.node, anchor.offset)
  }
  registry.set(SELECTION_HIGHLIGHT_NAME, new HighlightCtor(domRange))
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

/**
 * DEFER-5 / DXS-20 — builds a `(sectionIndex, blockIndex) -> 1-based page`
 * lookup from the already-computed pagination result, for PAGE/NUMPAGES
 * field evaluation and TOC page numbers. Reads only `Page`'s already-public
 * shape (`PageLineRef.paragraphPath`'s first element is the paragraph's
 * index among its section's direct blocks, matching the addressing
 * `updateFields`/`collectTocEntries` use) — this stays a read-only consumer
 * of pagination's output, not a change to pagination itself (paginate.ts/
 * breakLines.ts are out of this branch's scope). A paragraph spanning
 * several pages resolves to the EARLIEST page it appears on.
 */
function buildPageOfParagraph(
  pages: ReadonlyArray<Page>,
): (sectionIndex: number, blockIndex: number) => number | undefined {
  const pageByKey = new Map<string, number>()

  for (const page of pages) {
    const pageNumber = page.pageIndex + 1
    for (const column of page.columns) {
      for (const line of column.lines) {
        const blockIndex = line.paragraphPath[0]
        if (blockIndex === undefined) {
          continue
        }
        const key = `${page.sectionIndex}:${blockIndex}`
        const existing = pageByKey.get(key)
        if (existing === undefined || pageNumber < existing) {
          pageByKey.set(key, pageNumber)
        }
      }
    }
  }

  return (sectionIndex, blockIndex) => pageByKey.get(`${sectionIndex}:${blockIndex}`)
}

function DocxEditor({
  bundle,
  file,
  containerRef,
  onBundleChange,
  onPageCountChange,
}: {
  bundle: DocxBundle
  file: Extract<ViewerProps['file'], { kind: 'binary' }>
  containerRef: React.RefObject<HTMLDivElement | null>
  onBundleChange: (next: DocxBundle) => void
  onPageCountChange?: (pageCount: number) => void
}) {
  const editorRootRef = useRef<HTMLDivElement | null>(null)
  // USR-05 — pointer selection is driven from the model (see
  // handleSurfaceMouseDown): the anchor of an in-progress drag selection.
  const dragAnchorRef = useRef<Range['anchor'] | null>(null)
  // USR-06 — set by Find next/prev so the post-render selection sync scrolls
  // the match into view.
  const revealSelectionRef = useRef(false)
  // USR-07/USR-09 — whether the last keydown was handled by the model; only
  // natively-handled navigation keys re-read the DOM selection on keyup, so a
  // model selection (Ctrl+A, formatting) is never overwritten by a stale DOM.
  const nativeNavigationKeyRef = useRef(false)
  const historyRef = useRef(new History())
  // DEFER-4 / DXP-13 — this document's own embedded fonts (already
  // de-obfuscated), keyed off the raw archive so switching to a different
  // document (a new `bundle.rawArchive`) re-derives them.
  const embeddedFonts = useMemo(() => loadEmbeddedFonts(bundle.rawArchive), [bundle.rawArchive])
  const fontResolver = useMemo(() => createFontResolver(embeddedFonts), [embeddedFonts])
  // Resolves once the current document's embedded fonts have finished
  // registering via FontFace — the pagination effect awaits this so canvas
  // measurement (which reads what the browser has actually registered)
  // never races the registration it depends on.
  const embeddedFontsReadyRef = useRef<Promise<void>>(Promise.resolve())
  // DEFER-5 / DXS-20 — feedback for "Update Fields"/"Update TOC" below.
  // Not the shared app-wide toast system (`useToast`): that requires a
  // `<ToastProvider>` ancestor the isolated viewer-level tests that mount
  // `DocxEditor`/`DocxViewer` directly don't set up, and this file already
  // has its own established pattern (see `saveError`) for a dismissible
  // inline status message instead.
  const [fieldUpdateMessage, setFieldUpdateMessage] = useState<string | null>(null)
  const [documentModel, setDocumentModel] = useState(bundle.document)
  const [headerFooterOpen, setHeaderFooterOpen] = useState(false)
  // UX — the header/footer panel had no keyboard affordances at all: no
  // Escape-to-close, no focus moved into it on open, no focus restored to
  // the toolbar toggle on close (every other dialog/panel in the shell
  // does all three — see ShortcutsModal/UnsavedChangesDialog). Mirrors
  // FindReplace's own local (not the shared shortcut-dispatcher) Escape
  // handling just above, since this is likewise a plain inline panel, not
  // a modal.
  const headerFooterPanelRef = useRef<HTMLDivElement>(null)
  const headerFooterToggleRef = useRef<HTMLButtonElement>(null)
  const [range, setRange] = useState<Range | null>(null)
  // DXE-14 — right-click table-editing context menu; `null` when closed.
  // Screen coordinates only (not which table/cell), since by the time a
  // menu item fires, `range` (synced onto the click below) is already the
  // source of truth `handleToolbarCommand`/`toolbarToCommand` resolve against.
  const [tableContextMenuAt, setTableContextMenuAt] = useState<{ x: number; y: number } | null>(null)
  const [pages, setPages] = useState<ReadonlyArray<Page> | null>(null)
  // D23-PERF — the exact `documentModel` a completed `pages` array was laid
  // out against, updated in the SAME setState batch as `pages` (see the
  // pagination effect below). `PageStack` reads this instead of the live
  // `documentModel` directly: `documentModel` changes on every keystroke, and
  // since it flows into every `PageView`'s per-page `document`-derived memos
  // (run/hyperlink metadata, bookmark names — each walking the WHOLE
  // document), passing it straight through made every keystroke force a full
  // re-render of every page TWICE — once the instant `documentModel` changed
  // (still showing the OLD, not-yet-repaginated `pages`), and again once
  // pagination actually finished and `pages` itself updated. Keeping
  // `document`/`pages` pinned to the same pagination pass also fixes a latent
  // correctness gap: `pages`' `paragraphPath`s are block indices into
  // whichever document produced them, which can be stale (pointing at the
  // wrong paragraph) against a newer `documentModel` for the brief window
  // between an edit and its repagination — e.g. an insert/delete shifting
  // later block indices.
  const [pagesDocument, setPagesDocument] = useState(bundle.document)
  // D23-PERF-2 — bypasses PageStack's page virtualization while true, so
  // print (`handlePrint`) and PDF export (the `atlas:docx-*-full-render`
  // window events below, dispatched by `exportDocxPdf`) always see every
  // page in the live DOM instead of whatever happened to be scrolled into
  // view. See `PageStack.tsx`'s `forceRenderAll` prop doc comment.
  const [forceRenderAll, setForceRenderAll] = useState(false)
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
  // D24/DXL-19 — PageView already fully supports a CSS-transform zoom (its
  // `zoom` prop); this was simply never wired to anything, hardcoded to 1
  // below. Percent steps match the common Word/most-viewers convention.
  const [zoom, setZoom] = useState(1)
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

  // D23-PERF-2 — pages that must stay mounted no matter where the viewport
  // currently is: whichever page(s) hold the caret/selection. Computed
  // straight from `range`/`pages` during render (not in a `useEffect`) so
  // that the SAME commit that moves `range` (a click, arrow-key motion, or
  // Find revealing a match via `revealSelectionRef`) also mounts that page's
  // real DOM — the selection-sync `useLayoutEffect` further below runs
  // immediately after that commit and needs `positionToDomRange` to find a
  // real node, not a placeholder.
  const pinnedPageIndices = useMemo(() => {
    if (pages === null || range === null) {
      return EMPTY_PINNED_PAGE_INDICES
    }
    const anchorPages = findPagesForParagraphPath(pages, range.anchor.paragraphPath)
    const focusPages = findPagesForParagraphPath(pages, range.focus.paragraphPath)
    if (anchorPages.length === 0 && focusPages.length === 0) {
      return EMPTY_PINNED_PAGE_INDICES
    }
    return new Set([...anchorPages, ...focusPages])
  }, [pages, range])

  const toolbarState = useMemo(
    () => createToolbarState(documentModel, range, { spellCheck: spellCheckEnabled, trackChanges: trackChangesEnabled }),
    [documentModel, range, spellCheckEnabled, trackChangesEnabled],
  )

  const availableStyles = useMemo(() => {
    return Array.from(documentModel.styles.values())
      .filter(style => style.type === 'paragraph')
      .map(style => ({ id: style.id, name: style.name ?? style.id }))
  }, [documentModel.styles])

  // USR-11 — fonts the document uses first, then every installed system font
  // (via the main process), the bundled metric-compatible substitutes and
  // common Office fonts.
  const documentFonts = useMemo(() => collectReferencedFontFamilies(documentModel), [documentModel])
  const [systemFonts, setSystemFonts] = useState<ReadonlyArray<string>>([])
  useEffect(() => {
    let cancelled = false
    const listFonts = window.electronAPI?.fonts?.list
    if (listFonts === undefined) {
      return undefined
    }
    listFonts()
      .then((families) => {
        if (!cancelled) setSystemFonts(families)
      })
      .catch(() => undefined)
    return () => {
      cancelled = true
    }
  }, [])
  const availableFonts = useMemo(() => {
    const merged = new Set<string>([...COMMON_OFFICE_FONTS, ...FONT_FAMILIES.map((family) => family.wordName), ...systemFonts])
    return Array.from(merged).sort((a, b) => a.localeCompare(b))
  }, [systemFonts])

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

  // D23-PERF — `applyEditorCommand` gets a new identity every keystroke (it
  // closes over `documentModel`/`range`), so a callback built directly on top
  // of it would too — and that callback is `PageStack`'s `onResizeTableColumn`
  // prop, which would defeat `PageStack`/`PageView`'s memoization (see
  // `pagesDocument`'s doc comment above for the matching `document` prop
  // fix) on every single keystroke even though a column-resize drag is rare.
  // Routing the call through a ref keeps `handleResizeTableColumn` itself
  // referentially stable across renders while still always invoking the
  // latest `applyEditorCommand`.
  const applyEditorCommandRef = useRef(applyEditorCommand)
  useEffect(() => {
    applyEditorCommandRef.current = applyEditorCommand
  }, [applyEditorCommand])

  // DXE-14 — column resize by dragging a table's column border
  // (PageView.tsx's TableColumnResizeHandle, threaded here through
  // PageStack). A drag fires this exactly once, on release, with the final
  // width — never a stream of intermediate commands per pixel moved.
  const handleResizeTableColumn = useCallback(
    (tablePath: ReadonlyArray<number>, columnIndex: number, widthTwips: number) => {
      applyEditorCommandRef.current({ kind: 'resize-table-column', tablePath, columnIndex, widthTwips })
    },
    [],
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
    (commands: ReadonlyArray<Command>, nextRange: Range | null, trackChanges?: TrackChangesContext): boolean => {
      if (commands.length === 0) {
        return false
      }

      try {
        const batch: Command = commands.length === 1 ? commands[0] : { kind: 'composite', commands }
        const result = applyCommand(documentModel, batch, trackChanges)
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

  // D29 follow-up — a header/footer field commits its edit as ONE undo step
  // on blur, not per keystroke, so its `onChange` (further below) only
  // remembers the latest uncommitted value here rather than applying a
  // command immediately. This is also what fixes the Ctrl+S-while-focused
  // bug: the field previously only committed on blur, so saving (which reads
  // `documentModel`) while the field still had focus wrote the pre-edit
  // text. `flushHeaderFooterEdits` (used by the save path below) commits
  // whatever is still pending here — normally at most the one field the user
  // is still typing in — before the document is serialized. Declared this
  // early (rather than beside the other header/footer handlers, further
  // down) so `handleSaveInternal`'s dependency array below can reference it
  // — a `useCallback` dependency array is evaluated immediately, unlike the
  // callback body, so a later `const` would still be in its temporal dead
  // zone at that point.
  const pendingHeaderFooterEditsRef = useRef(
    new Map<string, { readonly kind: HeaderFooterKind; readonly id: string; readonly blockIndex: number; readonly text: string }>(),
  )

  // Commits every still-pending header/footer field edit as one command each
  // (a no-op, returning `documentModel` unchanged, when nothing is pending)
  // and returns the resulting document SYNCHRONOUSLY — `commitState`'s own
  // `setDocumentModel` is async, and the save path below needs the
  // just-committed text in hand immediately, not on the next render.
  const flushHeaderFooterEdits = useCallback((): DocxDocument => {
    const pending = pendingHeaderFooterEditsRef.current
    if (pending.size === 0) {
      return documentModel
    }

    const commands: Command[] = []
    for (const [key, edit] of pending) {
      const command = buildHeaderFooterTextEdit(documentModel, edit.kind, edit.id, edit.blockIndex, edit.text)
      if (command !== null) {
        commands.push(command)
      }
      pending.delete(key)
    }
    if (commands.length === 0) {
      return documentModel
    }

    try {
      const batch: Command = commands.length === 1 ? commands[0] : { kind: 'composite', commands }
      const result = applyCommand(documentModel, batch)
      historyRef.current.push(result.inverse)
      commitState(result.document, result.range ?? range)
      return result.document
    } catch {
      return documentModel
    }
  }, [commitState, documentModel, range])

  const syncRangeFromDom = useCallback(() => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    const nextRange = getSelectionFromDom(root)
    setRange(current => (rangeEquals(current, nextRange) ? current : nextRange))
  }, [])

  /**
   * USR-04/USR-05 — pointer selection computed from the model layout instead
   * of the browser's caret placement, which is unreliable on this
   * absolutely-positioned page layout (a click beside a line used to jump to
   * another paragraph; a double-click selected one character; a triple-click
   * spilled into the next paragraph). Single click places the caret (Shift
   * extends), double-click selects the word, triple-click the paragraph, and
   * dragging extends from the mousedown anchor.
   */
  const handleSurfaceMouseDown = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const root = editorRootRef.current
      if (event.button !== 0 || root === null) {
        return
      }
      const target = event.target as HTMLElement
      if (target.closest('button, input, select, textarea, [role="separator"], [data-resize-handle], .docx-table-resize-handle') !== null) {
        return
      }

      event.preventDefault()
      root.focus({ preventScroll: true })

      if (event.detail >= 3) {
        const paragraph = paragraphRangeFromClientPoint(event.clientX, event.clientY, root)
        dragAnchorRef.current = null
        if (paragraph !== null) {
          setRange(paragraph)
        }
        return
      }

      if (event.detail === 2) {
        const word = wordRangeFromClientPoint(event.clientX, event.clientY, root)
        dragAnchorRef.current = null
        if (word !== null) {
          setRange(word)
        }
        return
      }

      const position = positionFromClientPoint(event.clientX, event.clientY, root)
      if (position === null) {
        return
      }
      const anchor = event.shiftKey && range !== null ? range.anchor : position
      dragAnchorRef.current = anchor
      setRange({ anchor, focus: position })
    },
    [range],
  )

  useEffect(() => {
    const handleMove = (event: MouseEvent): void => {
      const root = editorRootRef.current
      const anchor = dragAnchorRef.current
      if (root === null || anchor === null || (event.buttons & 1) === 0) {
        return
      }
      const focus = positionFromClientPoint(event.clientX, event.clientY, root)
      if (focus !== null) {
        setRange((current) => {
          const next = { anchor, focus }
          return rangeEquals(current, next) ? current : next
        })
      }
    }
    const handleUp = (): void => {
      dragAnchorRef.current = null
    }
    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [])

  /**
   * DXE-14 — right-click inside a table cell opens Atlas's own table-editing
   * menu instead of the native context menu; right-clicking anywhere else
   * leaves the native menu alone (this viewer doesn't yet have a general
   * replacement for it — cut/copy/paste, spell-check suggestions, etc.).
   * Reads the DOM selection directly (`getSelectionFromDom`) rather than the
   * `range` state, which a right-click's `mousedown` has already moved in
   * the real DOM by the time this fires but this component's `onMouseUp`
   * sync hasn't caught up with yet — see the `onContextMenu`/`onMouseUp`
   * ordering note on the surface element below.
   */
  const handleTableContextMenu = useCallback(
    (event: ReactMouseEvent<HTMLDivElement>) => {
      const root = editorRootRef.current
      if (root === null) {
        return
      }

      const domRange = getSelectionFromDom(root)
      const enclosing = domRange === null ? null : findEnclosingTable(documentModel, domRange.focus.paragraphPath)
      if (enclosing === null) {
        return
      }

      event.preventDefault()
      setRange(domRange)
      setTableContextMenuAt({ x: event.clientX, y: event.clientY })
    },
    [documentModel],
  )

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

  // DXE-19 — rich HTML paste (tables/hyperlinks/images/colors/lists) is
  // async (image natural-size decoding needs it), unlike every other
  // editor command, so it can't reuse `applyEditorCommands` directly: it
  // applies its own composite command once `buildRichPasteCommands`
  // resolves, folding in a bundle patch (new relationships/media/numbering)
  // in the same `onBundleChange` as the document update, matching
  // `handleInsertImage`/`handleInsertHyperlink`'s own bundle+document
  // pairing above. `documentModel`/`bundle` are captured at call time (the
  // same closure-capture hazard those two helpers already accept for their
  // own async picker/decode work), so a paste started just before another
  // edit lands could in principle race it — acceptable for a user-initiated,
  // effectively-instantaneous paste.
  //
  // DXE-11 — `getTrackChanges()` is passed the same way Input.ts's keyboard
  // handlers pass it, so pasted text lands as `w:ins` (not a silent,
  // untracked insertion) whenever Track Changes is on; `applyComposite`
  // forwards it to each streamed `insert-text` sub-command exactly as a
  // single one would get it. Sub-commands that don't consult `trackChanges`
  // (`insert-table`, `apply-run-format`, `insert-hyperlink`) are unaffected —
  // see `applyComposite`'s own doc comment.
  const handleRichPaste = useCallback(
    (blocks: ReadonlyArray<PasteBlock>, focus: Position, selection: Range | null) => {
      void (async () => {
        try {
          const result = await buildRichPasteCommands(
            documentModel,
            bundleContextFor(bundle),
            blocks,
            focus,
            selection,
          )
          if (result === null) {
            return
          }

          const batch: Command =
            result.commands.length === 1 ? result.commands[0] : { kind: 'composite', commands: result.commands }
          const applied = applyCommand(documentModel, batch, getTrackChanges())
          historyRef.current.push(applied.inverse)
          const finalRange: Range = { anchor: result.finalCursor, focus: result.finalCursor }
          if (result.bundlePatch !== null) {
            onBundleChange({ ...bundle, ...result.bundlePatch, document: applied.document })
          }
          commitState(applied.document, finalRange)
        } catch (error) {
          setSaveError(error instanceof Error ? error.message : String(error))
        }
      })()
    },
    [bundle, commitState, documentModel, getTrackChanges, onBundleChange],
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

      if (html.length > 0) {
        const blocks = htmlToPasteBlocks(html)
        if (blocks.length > 0) {
          event.preventDefault()
          handleRichPaste(blocks, focus, range)
          return
        }
      }

      const text = clipboard.getData('text/plain')
      const paragraphs =
        html.length > 0 ? htmlToParagraphs(html) : textToParagraphs(text)

      if (paragraphs.length === 0) {
        return
      }

      event.preventDefault()

      const { commands, finalCursor } = buildPasteCommands(paragraphs, focus, range)
      const finalRange: Range = { anchor: finalCursor, focus: finalCursor }
      // DXE-11 — plain-text/plain-paragraph paste (no parseable HTML, or
      // `DOMParser` unavailable) is still an insertion; track it the same
      // way the rich-paste path above does.
      applyEditorCommands(commands, finalRange, getTrackChanges())
    },
    [applyEditorCommands, getTrackChanges, handleRichPaste, range],
  )

  // USR-01 — this handler must receive the NATIVE `beforeinput` InputEvent.
  // React's synthetic `onBeforeInput` is a polyfill built on the legacy
  // `textInput`/`keypress` events in Chromium/Electron: its nativeEvent has
  // no `inputType`, so handleBeforeInput matched no case and returned null
  // while preventDefault() still blocked the browser's own insertion — every
  // keystroke was silently dropped. It is attached with addEventListener in
  // the effect right below the surface ref instead of via the React prop.
  const handleBeforeInputEvent = useCallback(
    (event: InputEvent) => {
      const nativeEvent = event

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

  const beforeInputHandlerRef = useRef(handleBeforeInputEvent)
  useEffect(() => {
    beforeInputHandlerRef.current = handleBeforeInputEvent
  }, [handleBeforeInputEvent])

  useEffect(() => {
    const surface = editorRootRef.current
    if (surface === null) {
      return undefined
    }

    const listener = (event: Event): void => {
      beforeInputHandlerRef.current(event as InputEvent)
    }
    surface.addEventListener('beforeinput', listener)
    return () => surface.removeEventListener('beforeinput', listener)
  }, [])

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
      nativeNavigationKeyRef.current = NATIVE_NAVIGATION_KEYS.has(event.key)

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
      revealSelectionRef.current = true
      setRange(nextMatch.range)
    }
  }, [documentModel])

  const handleFindNext = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const nextMatch = findNext(documentModel, findQuery, findOptions, range?.focus ?? null)
    if (nextMatch !== null) {
      revealSelectionRef.current = true
      setRange(nextMatch.range)
    }
  }, [documentModel, findOptions, findQuery, range])

  const handleFindPrev = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const prevMatch = findPrev(documentModel, findQuery, findOptions, range?.anchor ?? null)
    if (prevMatch !== null) {
      revealSelectionRef.current = true
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

  // D23-PERF-2 — outline/heading navigation addresses a paragraph by its
  // running position among every `.docx-page__line[data-paragraph-path]` in
  // the document (matching how `collectDocxMetrics` numbered `navSeeds`),
  // not by the model's own `paragraphPath`, so it can't reuse
  // `pinnedPageIndices`' path-based lookup — that positional line simply
  // doesn't exist in the DOM yet when its page has been virtualized out.
  // Try the fast path first (works whenever the target is already mounted —
  // on-screen, or incidentally pinned); if it isn't there, force every page
  // to mount synchronously (`flushSync`, so the DOM is actually updated
  // before the retry runs, not just scheduled), retry, then hand back to
  // the normal virtualized window on the next frame. Jumping via the
  // outline is a deliberate, infrequent action, not a hot path — unlike
  // typing, a one-time full mount here is an acceptable cost.
  const handleScrollToParagraph = useCallback((paragraphIndex: number) => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    const tryScroll = (): boolean => {
      const lines = root.querySelectorAll<HTMLElement>('.docx-page__line[data-paragraph-path]')
      const target = lines.item(paragraphIndex)
      if (target === null) {
        return false
      }
      target.scrollIntoView({ behavior: 'smooth', block: 'center' })
      return true
    }

    if (tryScroll()) {
      return
    }

    flushSync(() => setForceRenderAll(true))
    tryScroll()
    requestAnimationFrame(() => setForceRenderAll(false))
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

      // D29 follow-up — commit any header/footer field the user is still
      // typing in (Ctrl+S never blurs it) before reading the document to
      // save; see `flushHeaderFooterEdits`'s own doc comment above.
      const documentToSave = flushHeaderFooterEdits()

      try {
        const nextBytes = await saveDocx({
          ...bundle,
          document: documentToSave,
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

        // Record *what was actually written* (`documentToSave`, including any
        // header/footer edit `flushHeaderFooterEdits` just committed above)
        // as the new saved baseline. We deliberately do NOT also call
        // `setDirty(false)` here: if the user kept editing while the
        // `await`s above were in flight, `documentModel` may already have
        // moved on since this snapshot, and the effect below recomputes
        // dirty from whatever the *current* render's `documentModel` is once
        // this state update lands — never from this stale closure.
        setLastSavedDocument(documentToSave)
        return true
      } catch (error) {
        // RUN-14 — wrap JSZip/fast-xml-parser/DocxSaveError internals in a
        // friendly, actionable message instead of showing the raw exception.
        setSaveError(friendlyDocxErrorMessage(error, 'save'))
        return false
      }
    },
    [bundle, flushHeaderFooterEdits, savePath],
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
    // D23-PERF-2 — `window.print()`'s `@media print` pass reads whatever is
    // ACTUALLY in the DOM at the moment it's called; a virtualized page
    // that's currently a placeholder would print blank. `flushSync` forces
    // the `forceRenderAll` state update (and PageStack's resulting mount of
    // every page) to commit synchronously, so every page is real before
    // `window.print()` ever runs — a plain `setForceRenderAll(true)` would
    // only schedule that render, racing the synchronous print call right
    // after it.
    flushSync(() => setForceRenderAll(true))
    // Print uses native browser print dialog. The print stylesheet in
    // viewer-docx.css hides toolbar/comments/save UI and forces a page
    // break after each .docx-page so output matches on-screen pagination.
    document.body.classList.add('atlas-printing')
    try {
      window.print()
    } finally {
      document.body.classList.remove('atlas-printing')
      setForceRenderAll(false)
    }
  }, [])

  // D23-PERF-2 — `src/utils/export/pdf.ts`'s `exportDocxPdf` clones the live
  // `.docx-page-stack` from OUTSIDE this component (it's invoked generically
  // by element id from `App.tsx`'s export menu, with no reference into this
  // viewer's own state), so it can't call `setForceRenderAll` directly. It
  // instead dispatches these two `window` events immediately before/after
  // reading the DOM; `dispatchEvent` invokes listeners synchronously, and
  // `flushSync` below forces the resulting full-mount render to actually
  // commit before `dispatchEvent` returns — so by the time `exportDocxPdf`
  // resumes and queries `.docx-page-stack`, every page is real, not a
  // placeholder. See that file's own doc comment.
  useEffect(() => {
    const handleForceFullRender = (): void => {
      flushSync(() => setForceRenderAll(true))
    }
    const handleReleaseFullRender = (): void => {
      setForceRenderAll(false)
    }
    window.addEventListener('atlas:docx-force-full-render', handleForceFullRender)
    window.addEventListener('atlas:docx-release-full-render', handleReleaseFullRender)
    return () => {
      window.removeEventListener('atlas:docx-force-full-render', handleForceFullRender)
      window.removeEventListener('atlas:docx-release-full-render', handleReleaseFullRender)
    }
  }, [])

  useEffect(() => {
    handlePrintRef.current = handlePrint
  }, [handlePrint])

  /**
   * DEFER-5 / DXS-20 — "Update field(s)": recalculates every resolvable
   * field's cached display text (DATE/TIME/AUTHOR/TITLE/REF/PAGEREF/SEQ/
   * PAGE/NUMPAGES; HYPERLINK/TOC/unknown fields are never touched here —
   * see `updateFields`'s doc comment). Author/title come from a light
   * regex read of `docProps/core.xml` (`parseCoreProps`) — Atlas has no
   * broader docProps model to draw on. Bookmark text/page maps come from
   * `collectBookmarkMaps` walking the current `documentModel` (see that
   * module's doc comment on scope — a bookmark inside a table cell or a
   * header/footer/footnote/endnote gets its text but no page). Applied
   * directly to `documentModel`, bypassing History/undo: a follow-up could
   * route it through a `replace-blocks`-shaped command instead, but wiring
   * into the shared editor Command/History system
   * (`docx/editor/commandTypes.ts`) is left to wave3/docx-editing's scope.
   */
  const handleUpdateFields = useCallback(() => {
    const coreXmlBytes = bundle.rawArchive?.get('docProps/core.xml')
    const coreXml = coreXmlBytes !== undefined ? new TextDecoder().decode(coreXmlBytes) : undefined
    const { author, title } = parseCoreProps(coreXml)
    const pageOfParagraph = pages !== null ? buildPageOfParagraph(pages) : undefined
    const { bookmarkText, bookmarkPage } = collectBookmarkMaps(documentModel, pageOfParagraph)

    const context: FieldEvaluationContext = {
      ...(author !== undefined ? { author } : {}),
      ...(title !== undefined ? { title } : {}),
      bookmarkText,
      bookmarkPage,
      ...(pages !== null ? { pageCount: pages.length } : {}),
      ...(pageOfParagraph !== undefined
        ? { currentPageOf: (path: ReadonlyArray<number>) => (path[1] !== undefined ? pageOfParagraph(path[0], path[1]) : undefined) }
        : {}),
      sequenceCounters: new Map(),
    }

    const { document: updated, updatedCount } = updateFields(documentModel, context)
    if (updatedCount === 0) {
      setFieldUpdateMessage('No fields needed updating.')
      return
    }

    setDocumentModel(updated)
    setFieldUpdateMessage(`Updated ${updatedCount} field${updatedCount === 1 ? '' : 's'}.`)
  }, [bundle.rawArchive, documentModel, pages])

  /**
   * DEFER-5 / DXS-20 — "Update table of contents": regenerates a
   * single-paragraph TOC field's entries from the document's current
   * headings (see `toc.ts`'s doc comment on that scope). No-ops with a
   * status message when the document has no such field, matching Word's
   * own behavior of the command doing nothing without a TOC.
   */
  const handleUpdateTableOfContents = useCallback(() => {
    const pageOfParagraph = pages !== null ? buildPageOfParagraph(pages) : undefined
    const { document: updated, updated: didUpdate } = updateTableOfContents(documentModel, pageOfParagraph)

    if (!didUpdate) {
      setFieldUpdateMessage('No table of contents found to update.')
      return
    }

    setDocumentModel(updated)
    setFieldUpdateMessage('Table of contents updated.')
  }, [documentModel, pages])

  // D29 — headers and footers are edited one plain-text paragraph at a time
  // (see docx/editor/headerFooter.ts for why: rewriting a whole part used to
  // flatten away images/fields/tables it held), through the normal command
  // path so each edit is its own undo step and the save path writes the part
  // back out.
  const headerFooterParts = useMemo(() => listHeaderFooterParts(documentModel), [documentModel])

  // UX — move focus into the panel (its close button) when it opens, and
  // back to the toolbar toggle that opened it once it closes, matching the
  // focus-management every other shell dialog/panel already has.
  useEffect(() => {
    if (!headerFooterOpen) return
    const panel = headerFooterPanelRef.current
    const toggle = headerFooterToggleRef.current
    panel?.querySelector<HTMLElement>('button, input')?.focus()
    return () => {
      toggle?.focus()
    }
  }, [headerFooterOpen])

  const handleHeaderFooterFieldChange = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number, text: string) => {
      pendingHeaderFooterEditsRef.current.set(`${kind}:${id}:${blockIndex}`, { kind, id, blockIndex, text })
    },
    [],
  )

  const handleHeaderFooterFieldBlur = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number, text: string) => {
      pendingHeaderFooterEditsRef.current.delete(`${kind}:${id}:${blockIndex}`)
      const command = buildHeaderFooterTextEdit(documentModel, kind, id, blockIndex, text)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  const handleAddHeaderFooterParagraph = useCallback(
    (kind: HeaderFooterKind, id: string) => {
      const command = buildInsertHeaderFooterParagraph(documentModel, kind, id)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  const handleRemoveHeaderFooterParagraph = useCallback(
    (kind: HeaderFooterKind, id: string, blockIndex: number) => {
      pendingHeaderFooterEditsRef.current.delete(`${kind}:${id}:${blockIndex}`)
      const command = buildRemoveHeaderFooterParagraph(documentModel, kind, id, blockIndex)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, documentModel],
  )

  // D24/DXL-19
  const handleZoomIn = useCallback(() => {
    setZoom((current) => clampZoom(current + ZOOM_STEP))
  }, [])
  const handleZoomOut = useCallback(() => {
    setZoom((current) => clampZoom(current - ZOOM_STEP))
  }, [])
  const handleZoomReset = useCallback(() => {
    setZoom(1)
  }, [])

  useEffect(() => {
    handleToolbarCommandRef.current = handleToolbarCommand
  }, [handleToolbarCommand])

  // DEFER-4 / DXP-13 — register this document's embedded fonts (if any) via
  // FontFace whenever it changes, and un-register the previous document's
  // faces on cleanup so a family name embedded differently by two different
  // documents never bleeds from one into the other.
  useEffect(() => {
    let cancelled = false
    let registeredFaces: ReadonlyArray<FontFace> = []

    const readyPromise = registerEmbeddedFonts(embeddedFonts).then((loaded) => {
      if (cancelled) {
        unregisterEmbeddedFonts(loaded)
        return
      }
      registeredFaces = loaded
    })
    embeddedFontsReadyRef.current = readyPromise

    return () => {
      cancelled = true
      unregisterEmbeddedFonts(registeredFaces)
    }
  }, [embeddedFonts])

  useEffect(() => {
    let cancelled = false

    setPaginationError(null)
    setPaginationProgress(null)

    void (async () => {
      // Ensure the browser has actually downloaded and registered the bundled
      // substitute fonts BEFORE we measure-and-paint. Otherwise pagination
      // computes widths from real TTF metrics while the DOM still renders with
      // a fallback font, producing accumulated drift -> mid-line gaps and
      // right-edge clipping. This document's own embedded fonts (DEFER-4) must
      // finish registering for the same reason.
      await preloadDocxFonts()
      await embeddedFontsReadyRef.current

      if (cancelled) {
        return
      }

      try {
        const nextPages = await paginate({
          document: documentModel,
          fontResolver,
          theme: bundle.theme,
          // D11 milestone 1/5 — these were parsed from word/settings.xml
          // (see docx/parser/settings.ts) but never threaded through to the
          // paginator in production; only paginate.test.ts's direct calls
          // exercised them. Without this, a real document that turns on
          // w:evenAndOddHeaders, or sets a non-default footnote/endnote
          // numbering restart/format, silently fell back to "off"/
          // "continuous" in the actual app.
          evenAndOddHeaders: bundle.settings?.evenAndOddHeaders,
          footnoteNumbering: bundle.settings?.footnotePr,
          endnoteNumbering: bundle.settings?.endnotePr,
          onProgress: progress => {
            if (!cancelled) {
              setPaginationProgress(progress)
            }
          },
          shouldCancel: () => cancelled,
        })

        if (!cancelled) {
          // D23-PERF — batched together so PageStack only ever sees a
          // `document`/`pages` pair produced by the SAME pagination pass;
          // see `pagesDocument`'s own doc comment.
          setPages(nextPages)
          setPagesDocument(documentModel)
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
  }, [documentModel, fontResolver, bundle.theme, bundle.settings])

  useLayoutEffect(() => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    // USR-06 — placing the DOM selection inside a contentEditable moves
    // keyboard focus into it in Chromium, which stole focus from the Find
    // box after the first match (Enter then edited the document). While a
    // form field outside the editor has focus, paint the model selection with
    // the CSS Custom Highlight API instead of moving the DOM selection.
    const activeElement = root.ownerDocument.activeElement
    const fieldHasFocus =
      activeElement !== null &&
      !root.contains(activeElement) &&
      activeElement.matches('input, textarea, select, [contenteditable="true"]')
    if (fieldHasFocus) {
      paintSelectionHighlight(root, range)
    } else {
      paintSelectionHighlight(root, null)
      syncSelectionToDom(root, range)
    }

    if (revealSelectionRef.current && range !== null) {
      revealSelectionRef.current = false
      const point = positionToDomRange(range.focus, root)
      const node = point?.node ?? null
      const element = node === null ? null : node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement
      if (element !== null && typeof element.scrollIntoView === 'function') {
        element.scrollIntoView({ block: 'center', inline: 'nearest' })
      }
    }
  }, [pages, range])

  useEffect(() => () => paintSelectionHighlight(null, null), [])

  useEffect(() => {
    onPageCountChange?.(pages?.length ?? 1)
  }, [onPageCountChange, pages])

  const pageCount = useMemo(
    () => pages?.length ?? (documentModel.sections.length || 1),
    [documentModel.sections.length, pages],
  )

  return (
    <div className="docx-viewer__editor-shell">
      <div className="docx-viewer__topbar">
        <Toolbar
          state={toolbarState}
          onCommand={handleToolbarCommand}
          availableStyles={availableStyles}
          availableFonts={availableFonts}
          documentFonts={documentFonts}
          trailing={
            <>
              <span className="docx-toolbar__page-count docx-viewer__meta">Pages: {pageCount}</span>
              <span className="docx-toolbar__action-divider" aria-hidden="true" />
              <div className="docx-viewer__zoom-controls" role="group" aria-label="Zoom">
                <button className="docx-toolbar__action" type="button" onClick={handleZoomOut} disabled={zoom <= MIN_ZOOM} aria-label="Zoom out" title="Zoom out">
                  <ZoomOut aria-hidden="true" />
                </button>
                <button className="docx-toolbar__action docx-toolbar__zoom-level" type="button" onClick={handleZoomReset} aria-label="Reset zoom to 100%" title="Reset zoom">
                  {Math.round(zoom * 100)}%
                </button>
                <button className="docx-toolbar__action" type="button" onClick={handleZoomIn} disabled={zoom >= MAX_ZOOM} aria-label="Zoom in" title="Zoom in">
                  <ZoomIn aria-hidden="true" />
                </button>
              </div>
              <span className="docx-toolbar__action-divider" aria-hidden="true" />
              <button
                className="docx-toolbar__action"
                type="button"
                onClick={handleUpdateFields}
                aria-label="Update fields"
                title="Update fields — recalculate DATE/TIME/AUTHOR/TITLE/REF/PAGEREF/SEQ/PAGE/NUMPAGES"
              >
                <RefreshCw aria-hidden="true" />
              </button>
              <button
                className="docx-toolbar__action"
                type="button"
                onClick={handleUpdateTableOfContents}
                aria-label="Update table of contents"
                title="Update table of contents"
              >
                <ListTree aria-hidden="true" />
              </button>
              {headerFooterParts.length > 0 && (
                <button
                  ref={headerFooterToggleRef}
                  className="docx-toolbar__action"
                  type="button"
                  onClick={() => setHeaderFooterOpen((open) => !open)}
                  aria-label="Header and footer"
                  aria-pressed={headerFooterOpen}
                  title="Edit header and footer"
                >
                  <PanelTop aria-hidden="true" />
                </button>
              )}
              <button className="docx-toolbar__action" type="button" onClick={handlePrint} aria-label="Print document" title="Print (Ctrl+P)">
                <Printer aria-hidden="true" />
              </button>
              <button className="docx-toolbar__action" type="button" onClick={() => void handleSaveAs()} aria-label="Save As" title="Save as…">
                <SaveAll aria-hidden="true" />
              </button>
              <button className="docx-toolbar__action docx-toolbar__action--primary" type="button" onClick={() => void handleSave()} aria-label="Save" title="Save (Ctrl+S)">
                <Save aria-hidden="true" />
                <span>Save</span>
              </button>
            </>
          }
        />
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
      {headerFooterOpen && headerFooterParts.length > 0 ? (
        <div
          ref={headerFooterPanelRef}
          className="docx-viewer__header-footer"
          role="group"
          aria-label="Header and footer"
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.stopPropagation()
              setHeaderFooterOpen(false)
            }
          }}
        >
          <div className="docx-viewer__header-footer-title">
            <span>Header and footer</span>
            <button
              type="button"
              className="docx-viewer__error-dismiss"
              onClick={() => setHeaderFooterOpen(false)}
              aria-label="Close header and footer"
            >
              <X aria-hidden="true" />
            </button>
          </div>
          {headerFooterParts.map((part) => {
            const base = `${part.kind === 'header' ? 'Header' : 'Footer'}${part.type === 'default' ? '' : ` (${part.type} page)`}`
            // A single plain-text paragraph and nothing else keeps the plain
            // "Header"/"Footer" label; anything with more than one row (a
            // second paragraph, or a placeholder alongside the text) numbers
            // each row by its position so it's clear which line is which.
            const numbered = part.rows.length > 1
            // Matches `buildRemoveHeaderFooterParagraph`'s own guard: never
            // offer to empty a part down to zero blocks.
            const canRemove = part.rows.length > 1
            return (
              <div key={`${part.kind}-${part.id}`} className="docx-viewer__header-footer-part">
                <span className="docx-viewer__header-footer-part-title">{base}</span>
                {part.rows.map((row, rowIndex) => {
                  if (row.kind === 'placeholder') {
                    return (
                      <div key={`${part.kind}-${part.id}-${row.blockIndex}`} className="docx-viewer__header-footer-placeholder">
                        {numbered ? `Line ${rowIndex + 1}: ` : ''}
                        {row.label}
                        <span className="docx-viewer__meta">
                          {' '}
                          Not plain text — left exactly as it is.
                        </span>
                      </div>
                    )
                  }

                  const label = numbered ? `${base} line ${rowIndex + 1}` : base
                  return (
                    <label
                      key={`${part.kind}-${part.id}-${row.blockIndex}-${row.text}`}
                      className="docx-viewer__header-footer-field"
                    >
                      <span>{label}</span>
                      <span className="docx-viewer__header-footer-row">
                        <input
                          type="text"
                          defaultValue={row.text}
                          aria-label={label}
                          onChange={(event) =>
                            handleHeaderFooterFieldChange(part.kind, part.id, row.blockIndex, event.currentTarget.value)
                          }
                          onBlur={(event) =>
                            handleHeaderFooterFieldBlur(part.kind, part.id, row.blockIndex, event.currentTarget.value)
                          }
                        />
                        {canRemove && (
                          <button
                            type="button"
                            className="docx-viewer__error-dismiss"
                            onClick={() => handleRemoveHeaderFooterParagraph(part.kind, part.id, row.blockIndex)}
                            aria-label={`Remove ${label}`}
                            title="Remove this line"
                          >
                            <X aria-hidden="true" />
                          </button>
                        )}
                      </span>
                    </label>
                  )
                })}
                <button
                  type="button"
                  className="docx-viewer__header-footer-add"
                  onClick={() => handleAddHeaderFooterParagraph(part.kind, part.id)}
                >
                  + Add line
                </button>
              </div>
            )
          })}
        </div>
      ) : null}
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
          onKeyDown={handleKeyDownEvent}
          onMouseDown={handleSurfaceMouseDown}
          onKeyUp={() => {
            if (nativeNavigationKeyRef.current) {
              syncRangeFromDom()
            }
          }}
          onFocus={() => {
            if (range === null) {
              syncRangeFromDom()
            }
          }}
          onCompositionStart={handleCompositionStart}
          onCompositionUpdate={handleCompositionUpdate}
          onCompositionEnd={handleCompositionEnd}
          onPaste={handlePaste}
          onContextMenu={handleTableContextMenu}
        >
          {pages !== null ? (
            <MediaContext.Provider value={mediaResolver}>
              <PageStack
                pages={pages}
                zoom={zoom}
                document={pagesDocument}
                theme={bundle.theme}
                relationships={bundle.relationships}
                onResizeTableColumn={handleResizeTableColumn}
                // D23-PERF-2 — `containerRef` (the outer `.docx-viewer`,
                // owned by the parent `DocxViewer` component, threaded down
                // as a prop) is the element that actually scrolls: it's the
                // one with a real `height`/`overflow: auto` in the cascade
                // (`index.css`'s `.docx-viewer` rule). `editorRootRef` (the
                // inner `.docx-viewer__surface`, this contentEditable) sits
                // one level in and never itself overflows — its own
                // `overflow: auto` never engages because nothing constrains
                // its height below its content's — so measuring scroll
                // position against it would see a `clientHeight` that
                // always equals the full document height and "virtualize"
                // nothing.
                scrollContainerRef={containerRef}
                pinnedPageIndices={pinnedPageIndices}
                forceRenderAll={forceRenderAll}
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
        {tableContextMenuAt !== null ? (
          <>
            {/* DXE-14 — a full-viewport transparent layer that closes the menu on
                any outside click/right-click, mirroring Toolbar.tsx's popovers
                (whose own useClickOutside hook isn't exported/reusable here). */}
            <div
              className="docx-viewer__context-menu-overlay"
              onClick={() => setTableContextMenuAt(null)}
              onContextMenu={(event) => {
                event.preventDefault()
                setTableContextMenuAt(null)
              }}
            />
            <div
              className="docx-toolbar__popover docx-viewer__context-menu"
              style={{ position: 'fixed', top: tableContextMenuAt.y, left: tableContextMenuAt.x }}
              role="menu"
              aria-label="Table editing"
            >
              <TableEditMenuItems
                onCommand={handleToolbarCommand}
                onAfterCommand={() => setTableContextMenuAt(null)}
              />
            </div>
          </>
        ) : null}
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
      {fieldUpdateMessage !== null ? (
        <div className="docx-viewer__field-status" role="status">
          <span>{fieldUpdateMessage}</span>
          <button
            type="button"
            className="docx-viewer__error-dismiss"
            onClick={() => setFieldUpdateMessage(null)}
            aria-label="Dismiss message"
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
  // USR-13 — the status bar used the section count as the page count; the
  // editor reports the real paginated page count instead.
  const [layoutPageCount, setLayoutPageCount] = useState<number | null>(null)
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
      pages: layoutPageCount ?? (documentModel.sections.length || 1),
    })
  }, [documentModel, layoutPageCount, metrics, setStats])

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
      {bundle !== null ? <DocxEditor bundle={bundle} file={file} containerRef={containerRef} onBundleChange={setBundle} onPageCountChange={setLayoutPageCount} /> : null}
    </div>
  )
}

export const DocxViewer = memo(DocxViewerBase)
