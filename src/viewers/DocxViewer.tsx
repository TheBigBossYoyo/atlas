import {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type FormEvent,
  type ClipboardEvent as ReactClipboardEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from 'react'

import { ListTree, PanelTop, Printer, RefreshCw, Save, SaveAll, X, ZoomIn, ZoomOut } from 'lucide-react'

import type { NavItem, ViewerProps } from '../formats/types'
import { useTranslate, type TranslateFn } from '../i18n'
import {
  loadDocx,
  type DocxBundle,
} from '../docx'
import { useDocxPagination } from '../docx/render/useDocxPagination'
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
  buildSpellCheckReplacement,
  decodeImageNaturalSizePt,
  ensureListNumbering,
  findEnclosingTable,
  findParagraph,
  extendOrCollapse,
  handleBeforeInput,
  handleKeyDown,
  moveCursorToLineEnd,
  moveCursorToLineStart,
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
  type ImageMimeType,
  type PasteBlock,
  type Position,
  type Range,
  type TrackChangesContext,
} from '../docx/editor'
import { getSelectionFromDom } from '../docx/render/selectionDom'
import { useDocxSelectionPainting } from '../docx/render/useDocxSelectionPainting'
import { useDocxComments } from '../docx/render/useDocxComments'
import { useDocxHeaderFooter } from '../docx/render/useDocxHeaderFooter'
import { useDocxSave } from '../docx/render/useDocxSave'
import { useDocxPrompt } from '../docx/render/useDocxPrompt'
import { useDocxHyperlink } from '../docx/render/useDocxHyperlink'
import { useDocxFields } from '../docx/render/useDocxFields'
import { useDocxFullRender } from '../docx/render/useDocxFullRender'
import { useDocxVerticalCaret } from '../docx/render/useDocxVerticalCaret'
import { useDocxFind } from '../docx/render/useDocxFind'
import { useDocxFontChoices } from '../docx/render/useDocxFontChoices'
import { MAX_ZOOM, MIN_ZOOM, useDocxZoom } from '../docx/render/useDocxZoom'
import { isRangeCollapsed, normalizeRange, rangeEquals } from '../docx/editor/Selection'
import {
  caretClientRect,
  paragraphRangeFromClientPoint,
  positionFromClientPoint,
  positionOnAdjacentLine,
  positionOnePageVertically,
  wordRangeFromClientPoint,
  type VerticalDirection,
} from '../docx/editor/Cursor'
import { Toolbar } from '../docx/editor/toolbar/Toolbar'
import { TableEditMenuItems } from '../docx/editor/toolbar/TableEditMenuItems'
import type { ToolbarCommand, ToolbarState } from '../docx/editor/toolbar/toolbarTypes'
import { DocxPromptDialog } from './DocxPromptDialog'
import './__styles__/viewer-docx.css'
import {
  useSetNavItems,
  useSetViewerStats,
} from './shared/useViewerContext'
import { useViewerShortcuts } from '../hooks/useShortcutManager'
import { scrollIntoViewRespectingMotionPreference } from '../utils/motionPreference'

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

// D24/DXL-19
/** D12/DXE-03 — keys whose native contentEditable behavior always mutates
 * content; see handleKeyDownEvent's preventDefault-safety comment. */
const MUTATING_KEYS: ReadonlySet<string> = new Set(['Backspace', 'Delete', 'Enter', 'Tab'])

/** DOCX-1 — vertical-caret keys, resolved entirely by `handleKeyDownEvent`
 * itself (via `Cursor.ts`'s line-geometry helpers) before `Input.ts` ever
 * sees them; always `preventDefault()`ed so the browser's own broken native
 * vertical-caret handling on this `inline-block`-run layout never runs. */
const VERTICAL_MOVE_KEYS: ReadonlySet<string> = new Set(['ArrowUp', 'ArrowDown', 'PageUp', 'PageDown'])

function normalizeDocxBuffer(content: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (content instanceof ArrayBuffer) {
    return content
  }

  return Uint8Array.from(content).buffer
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

// i18n — `docx/editor/headerFooter.ts` labels read-only header/footer
// content with a handful of fixed `[Bracketed]` English tags (`[Image]`,
// `[Table]`, `[Link: text]`, …). That module stays English-only and
// untouched here (it's covered by exact-string tests shared with the DOCX
// round-trip/serializer work) — this purely presentational helper swaps
// each known tag for its translation right before render, so a reviewer
// diffing headerFooter.ts sees no change while the panel still reads in
// the active locale.
function translateHeaderFooterLabel(label: string, t: TranslateFn): string {
  const linked = label.replace(/\[Link: ([^\]]*)\]/g, (_match, text: string) => `[${t('docx.headerFooter.placeholderLinkPrefix')}: ${text}]`)
  const TAG_TRANSLATIONS: ReadonlyArray<readonly [string, string]> = [
    ['[Page number]', `[${t('docx.headerFooter.placeholderPageNumber')}]`],
    ['[Total pages]', `[${t('docx.headerFooter.placeholderTotalPages')}]`],
    ['[Date]', `[${t('docx.headerFooter.placeholderDate')}]`],
    ['[Field]', `[${t('docx.headerFooter.placeholderField')}]`],
    ['[Comment]', `[${t('docx.headerFooter.placeholderComment')}]`],
    ['[Footnote]', `[${t('docx.headerFooter.placeholderFootnote')}]`],
    ['[Endnote]', `[${t('docx.headerFooter.placeholderEndnote')}]`],
    ['[Other content]', `[${t('docx.headerFooter.placeholderOtherContent')}]`],
    ['[Table]', `[${t('docx.headerFooter.placeholderTable')}]`],
    ['[Unknown content]', `[${t('docx.headerFooter.placeholderUnknownContent')}]`],
    ['[Empty paragraph]', `[${t('docx.headerFooter.placeholderEmptyParagraph')}]`],
    ['[Tracked change]', `[${t('docx.headerFooter.placeholderTrackedChange')}]`],
    ['[Link]', `[${t('docx.headerFooter.placeholderLink')}]`],
  ]
  return TAG_TRANSLATIONS.reduce((current, [tag, translated]) => current.split(tag).join(translated), linked)
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
  const t = useTranslate()
  const editorRootRef = useRef<HTMLDivElement | null>(null)
  // USR-05 — pointer selection is driven from the model (see
  // handleSurfaceMouseDown): the anchor of an in-progress drag selection.
  const dragAnchorRef = useRef<Range['anchor'] | null>(null)
  // DOCX-1 — the client-x "goal column" ArrowUp/ArrowDown/PageUp/PageDown
  // keep across consecutive presses, so moving down through a run of short
  // lines stays under the same visual column instead of snapping to each
  // line's own length (matches every real text editor). `null` means "no
  // goal yet — derive one from the current caret" and is restored after
  // every move that isn't itself a vertical one (typing, clicking,
  // Left/Right/Home/End, …see handleKeyDownEvent/handleSurfaceMouseDown).
  const verticalGoalXRef = useRef<number | null>(null)
  // DOCX-SMALL-WINDOW-2 — a vertical move (ArrowUp/Down/PageUp/Down) that
  // needs to land on a page adjacent to the current one has to wait for that
  // page's real DOM to exist first: `pinnedPageIndices` below only pins the
  // page(s) holding the CURRENT caret, and `PageStack`'s scroll-driven
  // virtualization (`src/docx/render/pageVirtualization.ts`) only mounts a
  // page's content once it (or a viewport-height buffer around it) has
  // actually scrolled into view — which plain keyboard movement never
  // triggers on its own. At a small enough window (or a document long
  // enough relative to it), the very next page can be a bare placeholder
  // with zero `.docx-page__line` elements, so `positionOnAdjacentLine`'s
  // geometry search finds nothing and the caret gets stuck at the last
  // MOUNTED line, silently indistinguishable from a genuine document
  // boundary. Set when a vertical move's target page isn't mounted yet;
  // pinning it (via `pinnedPageIndices` below) forces `PageStack` to mount
  // it, and the `useLayoutEffect` further below retries the exact same move
  // once that DOM exists.
   const historyRef = useRef(new History())
  // DEFER-5 / DXS-20 — feedback for "Update Fields"/"Update TOC" below.
  // Not the shared app-wide toast system (`useToast`): that requires a
  // `<ToastProvider>` ancestor the isolated viewer-level tests that mount
  // `DocxEditor`/`DocxViewer` directly don't set up, and this file already
  // has its own established pattern (see `saveError`) for a dismissible
  // inline status message instead.
  // Round-trip fidelity audit, DXS round 2 follow-up — the same dismissible
  // inline-status pattern as `fieldUpdateMessage`, for
  // `detectLossySaveWarnings`' findings after a save. `warnedAboutFidelityRef`
  // (not state: it must never itself trigger a re-render) makes this a
  // once-per-document notice rather than one on every save — see
  // `handleSaveInternal`'s doc comment for why. Both reset naturally on a
  // genuinely new document: `DocxEditor` remounts fresh whenever
  // `DocxViewerBase` swaps in a different `bundle` (see its own render,
  // `bundle !== null ? <DocxEditor .../> : null`), which is also why this
  // doesn't need to reset when `savePath` changes via "Save As" — that's
  // still the same in-memory document, not a newly opened one.
  const [documentModel, setDocumentModel] = useState(bundle.document)
  // DIRTY-1 — `historyRef.current`'s revision counter (see History.ts's own
  // doc comment on `getRevision`/`bumpRevision`) at the moment `documentModel`
  // last changed. Mirrored into state alongside `documentModel` itself
  // (rather than read straight off the ref where dirty is computed) for the
  // same reason `lastSavedRevision` below is state and not a ref: the dirty
  // effect must react to it changing on its own render, not just be able to
  // read it. Starts at 0 to match `History`'s own initial `currentRevision`.
  const [documentRevision, setDocumentRevision] = useState(0)
  // DOCX-17 — mirrors `documentModel` for the handful of callbacks below
  // (`syncRangeFromDom`, the drag-select `mousemove` listener) that are
  // deliberately kept identity-stable across renders (an empty `useCallback`/
  // `useEffect` dependency array — a `mousemove` listener in particular gets
  // re-subscribed to `window` on every dependency change otherwise) but still
  // need this turn's document model to resolve a table cell's paragraph path
  // (`Cursor.ts`'s DOM<->Position mapping needs it for table content, which
  // renders with no `data-paragraph-path` of its own — see Cursor.ts's
  // "Table-cell paragraph-path resolution" section).
  const documentModelRef = useRef(documentModel)
  useEffect(() => {
    documentModelRef.current = documentModel
  }, [documentModel])
  // UX — the header/footer panel had no keyboard affordances at all: no
  // Escape-to-close, no focus moved into it on open, no focus restored to
  // the toolbar toggle on close (every other dialog/panel in the shell
  // does all three — see ShortcutsModal/UnsavedChangesDialog). Mirrors
  // FindReplace's own local (not the shared shortcut-dispatcher) Escape
  // handling just above, since this is likewise a plain inline panel, not
  // a modal.
  const [range, setRange] = useState<Range | null>(null)
  // DXE-14 — right-click table-editing context menu; `null` when closed.
  // Screen coordinates only (not which table/cell), since by the time a
  // menu item fires, `range` (synced onto the click below) is already the
  // source of truth `handleToolbarCommand`/`toolbarToCommand` resolve against.
  const [tableContextMenuAt, setTableContextMenuAt] = useState<{ x: number; y: number } | null>(null)
  // Pagination + the font registration it depends on (see
  // `docx/render/useDocxPagination.ts`). `pagesDocument` is the exact model
  // `pages` was laid out against, which is what PageStack must render with —
  // see that hook for why passing the live `documentModel` instead re-rendered
  // every page twice per keystroke.
  const { pages, pagesDocument, paginationProgress, paginationError } = useDocxPagination(documentModel, bundle)

  // Paints `range` onto the DOM after every commit that could have moved the
  // selection or replaced the pages under it, and owns the "scroll the next
  // painted selection into view" flag that Find/outline navigation set.
  const { revealSelectionOnNextPaint } = useDocxSelectionPainting(
    editorRootRef,
    range,
    documentModelRef,
    pages,
  )
  const { forceRenderAll, handlePrint, forceAllPagesNow, releaseAllPages } = useDocxFullRender()
  const [saveError, setSaveError] = useState<string | null>(null)
  // F1 — the single-line text prompt backing Insert Hyperlink/Add Comment/
  // Reply (see DocxPromptDialog.tsx and handlePromptConfirm below). `id` is
  // bumped on every open so the dialog remounts instead of reusing stale
  // input state from a previous prompt.
  const { promptRequest, requestPrompt, closePrompt } = useDocxPrompt()
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
  // whenever documentModel diverges from the last-saved-or-loaded snapshot,
  // and register our own save() so App.tsx's global Ctrl+S/Save can reach it
  // when DOCX is active.
  //
  // DIRTY-1 — this used to compare `documentModel` against
  // `lastSavedDocument` *by reference*, on the theory that every edit
  // replaces `documentModel` immutably so reference-inequality is exactly
  // "changed since load/save". That held for ordinary edits, but not for
  // undo/redo: `History.undo`/`redo` (History.ts) rebuild the document by
  // applying the stored inverse command, which produces a brand-new object
  // graph every time — even when undoing lands back on content that is
  // byte-for-byte what was last saved, it is never `===` to that saved
  // snapshot, so the document stayed dirty forever. Comparing revision
  // numbers instead of document references fixes that without ever walking
  // the document: see History.ts's `getRevision`/`bumpRevision` doc comment
  // for why a counter that undo/redo can wind backward and forward again is
  // the right identity here, and for the alternatives (deep-equal every
  // keystroke; content hashing) that were considered and rejected.
  //
  // `lastSavedRevision` is state, not a ref: dirty is derived from it and
  // the *current* `documentRevision` together in one effect below, so that
  // if the user keeps typing while an async save() is still in flight, the
  // edits made after the save started are correctly still reported dirty
  // once the save resolves — the effect re-reads `documentRevision` fresh on
  // every render rather than trusting whatever value save()'s own closure
  // captured when it started.


  // D23-PERF-2 — pages that must stay mounted no matter where the viewport
  // currently is: whichever page(s) hold the caret/selection. Computed
  // straight from `range`/`pages` during render (not in a `useEffect`) so
  // that the SAME commit that moves `range` (a click, arrow-key motion, or
  // Find revealing a match via `revealSelectionRef`) also mounts that page's
  // real DOM — the selection-sync `useLayoutEffect` further below runs
  // immediately after that commit and needs `positionToDomRange` to find a
  // real node, not a placeholder.
  // DOCX-SMALL-WINDOW-2 — a vertical caret move that had to wait for its target
  // page to mount (see `docx/render/useDocxVerticalCaret.ts`).
  const { pendingMountPageIndex, requestVerticalMove } = useDocxVerticalCaret(
    documentModel,
    editorRootRef,
    setRange,
  )

  const pinnedPageIndices = useMemo(() => {
    if (pages === null || range === null) {
      return pendingMountPageIndex === null ? EMPTY_PINNED_PAGE_INDICES : new Set([pendingMountPageIndex])
    }
    const anchorPages = findPagesForParagraphPath(pages, range.anchor.paragraphPath)
    const focusPages = findPagesForParagraphPath(pages, range.focus.paragraphPath)
    if (anchorPages.length === 0 && focusPages.length === 0 && pendingMountPageIndex === null) {
      return EMPTY_PINNED_PAGE_INDICES
    }
    const indices = new Set([...anchorPages, ...focusPages])
    // DOCX-SMALL-WINDOW-2 — see `pendingVerticalMove`'s own doc comment: force
    // the target page to mount so the retry below has real DOM to search.
    if (pendingMountPageIndex !== null) {
      indices.add(pendingMountPageIndex)
    }
    return indices
  }, [pages, range, pendingMountPageIndex])

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
  const { documentFonts, availableFonts } = useDocxFontChoices(documentModel)
  const { zoom, handleZoomIn, handleZoomOut, handleZoomReset } = useDocxZoom()

  // DXE-25 — save-failure feedback used to be cleared by commitState, which
  // runs on every single edit, so the banner vanished the instant the user
  // typed the next character rather than staying up until the user dismissed
  // it or a save actually succeeded. commitState no longer touches saveError
  // at all; handleSave (below) is the only place that sets or clears it.
  // DIRTY-1 — reads `historyRef.current.getRevision()` synchronously rather
  // than accepting it as a parameter: every caller below already mutated
  // `historyRef.current` (push/coalesceWithLast, or Input.ts's own
  // history.undo/redo) *before* calling commitState, so the ref already
  // reflects this exact edit by the time we read it here — one call site to
  // keep in sync instead of threading the revision through every caller.
  const commitState = useCallback((nextDocument: DocxDocument, nextRange: Range | null) => {
    setDocumentModel(nextDocument)
    setDocumentRevision(historyRef.current.getRevision())
    setRange(nextRange)
  }, [])

  // Comments pane (see `docx/render/useDocxComments.ts`). `bumpRevision` is the
  // one slice of the undo history it needs: a comment edit is not undoable
  // through the command history, so without a fresh revision the dirty check
  // silently misses it (DIRTY-1).
  const bumpCommentRevision = useCallback(() => {
    historyRef.current.bumpRevision()
  }, [])
  const {
    commentsPaneOpen,
    setCommentsPaneOpen,
    commentsDocument,
    handleResolveComment,
    handleDeleteComment,
    confirmComment,
    confirmReply,
  } = useDocxComments(
    documentModel,
    range,
    commitState,
    bumpCommentRevision,
    bundle.document.comments.size,
    setSaveError,
  )

  // "Update field(s)" / "Update table of contents" (see
  // `docx/render/useDocxFields.ts`). `applyFieldUpdate` rather than
  // `commitState`: both actions bypass History and must not move the selection.
  const applyFieldUpdate = useCallback((next: DocxDocument) => {
    historyRef.current.bumpRevision()
    setDocumentModel(next)
    setDocumentRevision(historyRef.current.getRevision())
  }, [])
  const {
    fieldUpdateMessage,
    dismissFieldUpdateMessage,
    handleUpdateFields,
    handleUpdateTableOfContents,
  } = useDocxFields(bundle, documentModel, pages, applyFieldUpdate)

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

  // Find & replace (see `docx/render/useDocxFind.ts`). It needs the selection
  // to start from and to move, the command applier for the two replace actions,
  // and the reveal request so a match on a not-yet-mounted virtualized page is
  // scrolled to once it renders.
  const {
    findOpen,
    setFindOpen,
    matches,
    currentMatchIndex,
    handleFind,
    handleFindNext,
    handleFindPrev,
    handleReplace,
    handleReplaceAll,
  } = useDocxFind(documentModel, range, setRange, containerRef, applyEditorCommands, revealSelectionOnNextPaint)

  // Header/footer editing (see `docx/render/useDocxHeaderFooter.ts`). Declared
  // here, before `handleSaveInternal` below, because that callback's dependency
  // array references `flushHeaderFooterEdits` — a dependency array is evaluated
  // immediately, unlike the callback body, so a later `const` would still be in
  // its temporal dead zone at that point.
  const pushUndo = useCallback((inverse: Command) => {
    historyRef.current.push(inverse)
  }, [])
  const {
    headerFooterOpen,
    setHeaderFooterOpen,
    headerFooterPanelRef,
    headerFooterToggleRef,
    headerFooterParts,
    flushHeaderFooterEdits,
    handleHeaderFooterFieldChange,
    handleHeaderFooterFieldBlur,
    handleAddHeaderFooterParagraph,
    handleRemoveHeaderFooterParagraph,
  } = useDocxHeaderFooter(documentModel, range, commitState, applyEditorCommand, pushUndo)

  // Saving (see `docx/render/useDocxSave.ts`). Placed after the header/footer
  // hook because a save must flush whichever field the user is still typing in
  // before it reads the document to write.
  const getRevision = useCallback(() => historyRef.current.getRevision(), [])
  const { handleSave, handleSaveAs, fidelityWarningMessage, dismissFidelityWarning } = useDocxSave({
    bundle,
    documentRevision,
    range,
    filePath: file.path,
    flushHeaderFooterEdits,
    getRevision,
    editorRootRef,
    documentModelRef,
    setSaveError,
  })

  const syncRangeFromDom = useCallback(() => {
    const root = editorRootRef.current
    if (root === null) {
      return
    }

    const nextRange = getSelectionFromDom(root, documentModelRef.current)
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
  /**
   * DOCX-2 — the actual mapping from a click's client point to a model
   * position (below) is pure geometry (`getBoundingClientRect()` over
   * `.docx-page__line`/run elements), which is unaffected by an ancestor's
   * CSS clipping/scroll — unlike the browser's own native hit-testing
   * (`elementFromPoint`), which is. When the page renders wider than the
   * `.docx-viewer__surface` scrollport (e.g. once the Comments panel narrows
   * the available width), part of a line can fall into a clipped dead zone:
   * a click there never reaches `editorRootRef`'s subtree at all (its native
   * event target is an ancestor, e.g. `.docx-viewer__workspace`, not a
   * descendant), so `onMouseDown` bound only to the surface never fires,
   * `range` is never updated, and every following keystroke silently no-ops
   * behind `Input.ts`'s `range === null` guard — with focus/selection
   * already sitting wherever they were before the click (often the
   * document's very start from initial mount), giving zero visual
   * indication anything is wrong. Splitting the resolution logic out lets
   * `handleWorkspaceMouseDown` (below) run the exact same geometry-based
   * lookup for that dead-zone case — it doesn't need the click to have
   * physically landed inside `root`'s clipped box, only `root` itself as the
   * element to search and focus.
   */
  const resolveClickPosition = useCallback(
    (root: HTMLElement, clientX: number, clientY: number, detail: number, shiftKey: boolean): void => {
      root.focus({ preventScroll: true })
      verticalGoalXRef.current = null

      if (detail >= 3) {
        const paragraph = paragraphRangeFromClientPoint(clientX, clientY, root, documentModel)
        dragAnchorRef.current = null
        if (paragraph !== null) {
          setRange(paragraph)
        }
        return
      }

      if (detail === 2) {
        const word = wordRangeFromClientPoint(clientX, clientY, root, documentModel)
        dragAnchorRef.current = null
        if (word !== null) {
          setRange(word)
        }
        return
      }

      const position = positionFromClientPoint(clientX, clientY, root, documentModel)
      if (position === null) {
        return
      }
      const anchor = shiftKey && range !== null ? range.anchor : position
      dragAnchorRef.current = anchor
      setRange({ anchor, focus: position })
    },
    [documentModel, range],
  )

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
      resolveClickPosition(root, event.clientX, event.clientY, event.detail, event.shiftKey)
    },
    [resolveClickPosition],
  )

  /**
   * DOCX-2 fallback — a NATIVE listener on `containerRef.current` (the outer
   * `.docx-viewer`), not a React `onMouseDown` prop. `containerRef`'s own
   * `<div>` is owned by the parent `DocxViewerBase` component, one or more
   * DOM levels OUTSIDE everything this component renders — the actual
   * clickable/scrollable box of `.docx-viewer__surface` (and its child
   * `.docx-viewer__workspace`) can start well inside the page's own
   * left/right edges once the page renders wider than the available column,
   * e.g. with the Comments panel open (see `resolveClickPosition`'s doc
   * comment), and the resulting click's native event TARGET is
   * `.docx-viewer` itself, an ANCESTOR of everything this component owns —
   * never a DESCENDANT of it. React's synthetic event system only invokes a
   * handler on an element that is the target or one of the target's
   * ancestors *in the React tree*; attaching a JSX `onMouseDown` to any
   * `<div>` this component renders (all of them are DESCENDANTS of
   * `.docx-viewer`) can never fire for that click. A real
   * `addEventListener` on the actual DOM node has no such restriction.
   *
   * Only acts when the native click target is NOT inside the surface at all
   * (the dead-zone case) AND is not inside any of the other, non-editor UI
   * `.docx-viewer` also contains (toolbar, header/footer panel, Comments
   * panel, dialogs, menus) — so a normal click on rendered text is handled
   * exactly once, by `handleSurfaceMouseDown`, and a click on the
   * toolbar/Comments panel/a dialog is left entirely alone. Word's own
   * behavior: a click anywhere on the page places the caret at the nearest
   * real position; this is that, plus it guarantees the editor is never left
   * "looks focused but range is null" after a click.
   */
  const nonEditorUiSelector = [
    'button',
    'input',
    'select',
    'textarea',
    '[role="separator"]',
    '[role="dialog"]',
    '[data-resize-handle]',
    '.docx-table-resize-handle',
    '.docx-viewer__topbar',
    '.docx-viewer__header-footer',
    '.comments-pane',
    '.docx-toolbar__popover',
    '.docx-viewer__context-menu',
    '.docx-viewer__context-menu-overlay',
    '.modal-backdrop',
    '.docx-viewer__error',
    '.docx-viewer__field-status',
  ].join(', ')

  useEffect(() => {
    const container = containerRef.current
    if (container === null) {
      return undefined
    }

    const handleContainerMouseDown = (event: MouseEvent): void => {
      const root = editorRootRef.current
      const target = event.target as HTMLElement | null
      if (event.button !== 0 || root === null || target === null || root.contains(target)) {
        return
      }
      if (target.closest(nonEditorUiSelector) !== null) {
        return
      }

      event.preventDefault()
      resolveClickPosition(root, event.clientX, event.clientY, event.detail, event.shiftKey)
    }

    container.addEventListener('mousedown', handleContainerMouseDown)
    return () => container.removeEventListener('mousedown', handleContainerMouseDown)
  }, [containerRef, nonEditorUiSelector, resolveClickPosition])

  useEffect(() => {
    const handleMove = (event: MouseEvent): void => {
      const root = editorRootRef.current
      const anchor = dragAnchorRef.current
      if (root === null || anchor === null || (event.buttons & 1) === 0) {
        return
      }
      const focus = positionFromClientPoint(event.clientX, event.clientY, root, documentModelRef.current)
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

      const domRange = getSelectionFromDom(root, documentModel)
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
      setSaveError(t('docx.viewer.imagePickerUnavailable'))
      return
    }

    const focus = range?.focus ?? null
    if (focus === null) {
      setSaveError(t('docx.viewer.placeCursorBeforeImage'))
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
  }, [bundle, commitState, documentModel, onBundleChange, range, t])

  // Hyperlink insert (see `docx/render/useDocxHyperlink.ts`) — the one editor
  // action that mutates the BUNDLE as well as the document, since a link needs a
  // relationship id.
  const { handleInsertHyperlink, confirmHyperlink } = useDocxHyperlink(
    bundle,
    documentModel,
    range,
    commitState,
    onBundleChange,
    pushUndo,
    requestPrompt,
    setSaveError,
    t,
  )

  const handleToggleList = useCallback(
    (kind: 'bullet' | 'number') => {
      const command = toolbarToCommand(
        { kind: kind === 'bullet' ? 'toggle-bullet-list' : 'toggle-numbered-list' },
        range,
        documentModel,
      )
      if (command === null || command.kind !== 'insert-list') {
        setSaveError(t('docx.viewer.selectParagraphBeforeList'))
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
    [bundle, commitState, documentModel, onBundleChange, range, t],
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

      // DOCX-2 — the editor must never silently swallow a keystroke just
      // because `range` state hasn't caught up with a real, valid native
      // selection (e.g. a click resolved outside this component's own
      // pointer handlers, or any other path that left state stale). Before
      // giving up, re-read the DOM directly one time as a last resort —
      // cheap, and far better than a keystroke vanishing with no feedback.
      const effectiveRange = range ?? (() => {
        const root = editorRootRef.current
        const resynced = root === null ? null : getSelectionFromDom(root, documentModel)
        if (resynced !== null) {
          setRange(resynced)
        }
        return resynced
      })()

      const applied = applyResult(
        handleBeforeInput(nativeEvent, {
          document: documentModel,
          range: effectiveRange,
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

      // DOCX-1 — resolved entirely here, before Input.ts's `handleKeyDown`
      // (which has no DOM access and always no-ops on these four keys — see
      // its own comment). Always `preventDefault()`ed so the browser's own
      // broken native vertical-caret handling on this `inline-block`-run
      // layout never runs, even when there's nothing for us to move to yet
      // (e.g. no root mounted, or no selection).
      if (VERTICAL_MOVE_KEYS.has(event.key)) {
        event.preventDefault()
        const root = editorRootRef.current
        if (root === null || range === null) {
          return
        }

        const direction: VerticalDirection = event.key === 'ArrowUp' || event.key === 'PageUp' ? 'up' : 'down'
        const isPageMove = event.key === 'PageUp' || event.key === 'PageDown'

        // Matches ArrowLeft/ArrowRight's own convention just below: the
        // first non-Shift vertical press on an existing selection only
        // collapses it (to the edge that direction already reads towards),
        // it doesn't also move past that same edge on the same keystroke.
        if (!shift && !isRangeCollapsed(range)) {
          const { start, end } = normalizeRange(range)
          const collapsed = direction === 'up' ? start : end
          verticalGoalXRef.current = null
          setRange({ anchor: collapsed, focus: collapsed })
          return
        }

        const fromRect = caretClientRect(range.focus, root, documentModel)
        if (fromRect === null) {
          // Can't locate the caret in the live DOM (e.g. mid-render) — leave
          // the selection alone rather than guess; default is still
          // prevented above.
          return
        }

        const goalX = verticalGoalXRef.current ?? fromRect.left
        verticalGoalXRef.current = goalX
        const viewportHeight = containerRef.current?.clientHeight ?? root.getBoundingClientRect().height

        const target = isPageMove
          ? positionOnePageVertically(
              fromRect,
              goalX,
              direction,
              // USR-04's own doc comment on `scrollContainerRef` explains why
              // `containerRef` (the outer `.docx-viewer`), not `root`, is the
              // element that actually has a real scrollable viewport height.
              viewportHeight,
              root,
              documentModel,
            )
          : positionOnAdjacentLine(range.focus, goalX, direction, root, documentModel)

        if (target === null && pages !== null) {
          // DOCX-SMALL-WINDOW-2 — before concluding this is a genuine
          // document boundary, check whether there's a NEXT page in this
          // direction that simply isn't mounted yet (see
          // `pendingVerticalMove`'s own doc comment). Only the page(s) the
          // CURRENT range sits on are guaranteed mounted, so an adjacent page
          // that hasn't scrolled into view can be a bare placeholder with no
          // lines at all, indistinguishable from "no more document" to
          // `positionOnAdjacentLine`'s geometry search.
          const focusPageIndices = findPagesForParagraphPath(pages, range.focus.paragraphPath)
          if (focusPageIndices.length > 0) {
            const currentPageIndex = direction === 'down' ? Math.max(...focusPageIndices) : Math.min(...focusPageIndices)
            const mountPageIndex = direction === 'down' ? currentPageIndex + 1 : currentPageIndex - 1
            if (mountPageIndex >= 0 && mountPageIndex < pages.length && !pinnedPageIndices.has(mountPageIndex)) {
              requestVerticalMove({ direction, goalX, isPageMove, fromRect, viewportHeight, shift, baseRange: range, mountPageIndex })
              return
            }
          }
        }

        // No line to move to (top/bottom of the document, or of what's
        // currently mounted, with no further page left to try mounting) —
        // land on this line's own start/end instead of doing nothing,
        // matching Home/End's behavior for the current line.
        const newFocus =
          target ??
          (direction === 'up'
            ? moveCursorToLineStart(range.focus, documentModel)
            : moveCursorToLineEnd(range.focus, documentModel))

        setRange(extendOrCollapse(range, newFocus, shift))
        return
      }

      verticalGoalXRef.current = null

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
    [applyResult, containerRef, documentModel, getTrackChanges, pages, pinnedPageIndices, range, requestVerticalMove, setFindOpen],
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
      scrollIntoViewRespectingMotionPreference(target, { behavior: 'smooth', block: 'center' })
      return true
    }

    if (tryScroll()) {
      return
    }

    forceAllPagesNow()
    tryScroll()
    // Released next frame, not now: the smooth scroll above needs its target to
    // stay mounted while it animates.
    requestAnimationFrame(releaseAllPages)
  }, [forceAllPagesNow, releaseAllPages])

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

  // F1 — see handleInsertHyperlink's comment above: opens the shared
  // DocxPromptDialog instead of the (Electron-unsupported) `window.prompt`;
  // the actual mutation happens in `handlePromptConfirm`.
  const handleAddComment = useCallback(() => {
    const selection = range
    if (selection === null) {
      setSaveError(t('docx.viewer.selectBeforeComment'))
      return
    }

    requestPrompt({ kind: 'comment', selection })
  }, [range, requestPrompt, t])

  const handleReplyToComment = useCallback(
    (commentId: string) => {
      requestPrompt({ kind: 'reply', commentId, range })
    },
    [range, requestPrompt],
  )

  // F1 — applies whichever DocxPromptDialog request is pending once the user
  // submits it. Each branch is the exact body the old synchronous
  // `window.prompt(...)` callers ran right after reading a non-null result;
  // only the "how do we get the string" part changed (a real dialog instead
  // of a call Electron doesn't support), not the mutation logic itself.
  // Each branch's work lives with the feature that owns it — see
  // `useDocxHyperlink` and `useDocxComments`. This is only the dispatch, and it
  // closes the dialog first, as the previous single handler did.
  const handlePromptConfirm = useCallback(
    (value: string) => {
      const request = promptRequest
      closePrompt()
      if (request === null) return

      if (request.kind === 'hyperlink') {
        confirmHyperlink(request.selection, value)
      } else if (request.kind === 'comment') {
        confirmComment(request.selection, value)
      } else {
        confirmReply(request.commentId, request.range, value)
      }
    },
    [closePrompt, confirmComment, confirmHyperlink, confirmReply, promptRequest],
  )

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
      setCommentsPaneOpen,
      setFindOpen,
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

  useEffect(() => {
    handleSaveRef.current = handleSave
  }, [handleSave])


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
        key === 'f' || key === 'h' || key === 'p' || key === 'k' ||
        key === 'l' || key === 'e' || key === 'r' || key === 'j' ||
        key === '1' || key === '2' || key === '5' ||
        key === 'z' || key === 'y' || key === 'a' || key === 'b' || key === 'i' || key === 'u' ||
        // Ctrl+S is only ever handled here unshifted (see handleKeyDownEvent
        // above) — Ctrl+Shift+S must fall through to the shell tier's global
        // Save As (App.tsx's saveFileAs(), routed via registerSaveAs above),
        // not get silently eaten here as though this viewer already handled
        // it (it doesn't: it has no Ctrl+Shift+S branch of its own).
        (key === 's' && !event.shiftKey)
      if (!isReservedCombo) return false
      event.preventDefault()
      return true
    }, []),
  )


  useEffect(() => {
    handlePrintRef.current = handlePrint
  }, [handlePrint])


  // D29 — headers and footers are edited one plain-text paragraph at a time
  // (see docx/editor/headerFooter.ts for why: rewriting a whole part used to
  // flatten away images/fields/tables it held), through the normal command
  // path so each edit is its own undo step and the save path writes the part
  // back out.
  useEffect(() => {
    handleToolbarCommandRef.current = handleToolbarCommand
  }, [handleToolbarCommand])


  useEffect(() => {
    onPageCountChange?.(pages?.length ?? 1)
  }, [onPageCountChange, pages])

  const pageCount = useMemo(
    () => pages?.length ?? (documentModel.sections.length || 1),
    [documentModel.sections.length, pages],
  )

  // F1 — per-request copy (and, for hyperlink, the `https://` prefill the
  // old `window.prompt(..., 'https://')` call used) for the shared
  // DocxPromptDialog rendered below. Computed unconditionally (not just
  // while a request is pending) so the dialog always gets a defined
  // `initialValue`/labels even while `isOpen` is false and it renders null.
  const promptDialogCopy =
    promptRequest === null || promptRequest.kind === 'hyperlink'
      ? {
          titleText: t('docx.promptDialog.hyperlinkTitle'),
          label: t('docx.promptDialog.hyperlinkLabel'),
          confirmLabel: t('docx.promptDialog.hyperlinkConfirm'),
          initialValue: 'https://',
        }
      : promptRequest.kind === 'comment'
        ? {
            titleText: t('docx.promptDialog.commentTitle'),
            label: t('docx.promptDialog.commentLabel'),
            confirmLabel: t('docx.promptDialog.commentConfirm'),
            initialValue: '',
          }
        : {
            titleText: t('docx.promptDialog.replyTitle'),
            label: t('docx.promptDialog.replyLabel'),
            confirmLabel: t('docx.promptDialog.replyConfirm'),
            initialValue: '',
          }

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
              <span className="docx-toolbar__page-count docx-viewer__meta">{t('docx.viewer.pagesCount', { count: pageCount })}</span>
              <span className="docx-toolbar__action-divider" aria-hidden="true" />
              <div className="docx-viewer__zoom-controls" role="group" aria-label={t('docx.viewer.zoomGroupAria')}>
                <button className="docx-toolbar__action" type="button" onClick={handleZoomOut} disabled={zoom <= MIN_ZOOM} aria-label={t('docx.viewer.zoomOut')} title={t('docx.viewer.zoomOut')}>
                  <ZoomOut aria-hidden="true" />
                </button>
                <button className="docx-toolbar__action docx-toolbar__zoom-level" type="button" onClick={handleZoomReset} aria-label={t('docx.viewer.zoomReset')} title={t('docx.viewer.zoomResetTitle')}>
                  {Math.round(zoom * 100)}%
                </button>
                <button className="docx-toolbar__action" type="button" onClick={handleZoomIn} disabled={zoom >= MAX_ZOOM} aria-label={t('docx.viewer.zoomIn')} title={t('docx.viewer.zoomIn')}>
                  <ZoomIn aria-hidden="true" />
                </button>
              </div>
              <span className="docx-toolbar__action-divider" aria-hidden="true" />
              <button
                className="docx-toolbar__action"
                type="button"
                onClick={handleUpdateFields}
                aria-label={t('docx.viewer.updateFields')}
                title={t('docx.viewer.updateFieldsTitle')}
              >
                <RefreshCw aria-hidden="true" />
              </button>
              <button
                className="docx-toolbar__action"
                type="button"
                onClick={handleUpdateTableOfContents}
                aria-label={t('docx.viewer.updateToc')}
                title={t('docx.viewer.updateToc')}
              >
                <ListTree aria-hidden="true" />
              </button>
              {headerFooterParts.length > 0 && (
                <button
                  ref={headerFooterToggleRef}
                  className="docx-toolbar__action"
                  type="button"
                  onClick={() => setHeaderFooterOpen((open) => !open)}
                  aria-label={t('docx.headerFooter.title')}
                  aria-pressed={headerFooterOpen}
                  title={t('docx.headerFooter.editTitle')}
                >
                  <PanelTop aria-hidden="true" />
                </button>
              )}
              <button className="docx-toolbar__action" type="button" onClick={handlePrint} aria-label={t('docx.viewer.printDocument')} title={t('docx.viewer.printTitle')}>
                <Printer aria-hidden="true" />
              </button>
              <button className="docx-toolbar__action" type="button" onClick={() => void handleSaveAs()} aria-label={t('docx.viewer.saveAs')} title={t('docx.viewer.saveAsTitle')}>
                <SaveAll aria-hidden="true" />
              </button>
              <button className="docx-toolbar__action docx-toolbar__action--primary" type="button" onClick={() => void handleSave()} aria-label={t('docx.viewer.save')} title={t('docx.viewer.saveTitle')}>
                <Save aria-hidden="true" />
                <span>{t('docx.viewer.save')}</span>
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
          aria-label={t('docx.headerFooter.title')}
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.stopPropagation()
              setHeaderFooterOpen(false)
            }
          }}
        >
          <div className="docx-viewer__header-footer-title">
            <span>{t('docx.headerFooter.title')}</span>
            <button
              type="button"
              className="docx-viewer__error-dismiss"
              onClick={() => setHeaderFooterOpen(false)}
              aria-label={t('docx.headerFooter.close')}
            >
              <X aria-hidden="true" />
            </button>
          </div>
          {headerFooterParts.map((part) => {
            const partLabel = part.kind === 'header' ? t('docx.headerFooter.header') : t('docx.headerFooter.footer')
            const pageType =
              part.type === 'first'
                ? t('docx.headerFooter.pageTypeFirst')
                : part.type === 'even'
                  ? t('docx.headerFooter.pageTypeEven')
                  : null
            const base = pageType === null ? partLabel : t('docx.headerFooter.partWithType', { part: partLabel, type: pageType })
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
                        {numbered ? t('docx.headerFooter.lineNumberPrefix', { n: rowIndex + 1 }) : ''}
                        {translateHeaderFooterLabel(row.label, t)}
                        <span className="docx-viewer__meta">
                          {' '}
                          {t('docx.headerFooter.notPlainText')}
                        </span>
                      </div>
                    )
                  }

                  const label = numbered ? t('docx.headerFooter.lineLabel', { base, n: rowIndex + 1 }) : base

                  if (row.kind === 'mixed') {
                    // A paragraph mixing plain text with a drawing/field/
                    // hyperlink/etc (D29 follow-up 2): one editable input per
                    // text segment, interleaved with a read-only chip per
                    // atom, all inside the same visual line.
                    return (
                      <div key={`${part.kind}-${part.id}-${row.blockIndex}`} className="docx-viewer__header-footer-field">
                        <span>{label}</span>
                        <span className="docx-viewer__header-footer-row docx-viewer__header-footer-mixed">
                          {row.segments.map((segment, segmentPos) =>
                            segment.kind === 'atom' ? (
                              <span key={`atom-${segmentPos}`} className="docx-viewer__header-footer-atom">
                                {translateHeaderFooterLabel(segment.label, t)}
                              </span>
                            ) : (
                              <input
                                key={`text-${segment.segmentIndex}-${segment.text}`}
                                type="text"
                                defaultValue={segment.text}
                                aria-label={t('docx.headerFooter.editableTextAria', { label })}
                                className="docx-viewer__header-footer-segment-input"
                                onChange={(event) =>
                                  handleHeaderFooterFieldChange(
                                    part.kind,
                                    part.id,
                                    row.blockIndex,
                                    event.currentTarget.value,
                                    segment.segmentIndex,
                                  )
                                }
                                onBlur={(event) =>
                                  handleHeaderFooterFieldBlur(
                                    part.kind,
                                    part.id,
                                    row.blockIndex,
                                    event.currentTarget.value,
                                    segment.segmentIndex,
                                  )
                                }
                              />
                            ),
                          )}
                          {canRemove && (
                            <button
                              type="button"
                              className="docx-viewer__error-dismiss"
                              onClick={() => handleRemoveHeaderFooterParagraph(part.kind, part.id, row.blockIndex)}
                              aria-label={t('docx.headerFooter.removeLabel', { label })}
                              title={t('docx.headerFooter.removeLine')}
                            >
                              <X aria-hidden="true" />
                            </button>
                          )}
                        </span>
                        <span className="docx-viewer__meta">{t('docx.headerFooter.nonTextParts')}</span>
                      </div>
                    )
                  }

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
                            aria-label={t('docx.headerFooter.removeLabel', { label })}
                            title={t('docx.headerFooter.removeLine')}
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
                  {t('docx.headerFooter.addLine')}
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
          aria-label={t('docx.viewer.editorAria')}
          contentEditable
          suppressContentEditableWarning
          spellCheck={spellCheckEnabled}
          onKeyDown={handleKeyDownEvent}
          onMouseDown={handleSurfaceMouseDown}
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
              <div className="docx-viewer__loading-title">{t('docx.viewer.layoutFailedTitle')}</div>
              <div className="docx-viewer__loading-detail">{paginationError}</div>
            </div>
          ) : (
            <div className="docx-viewer__loading" role="status" aria-live="polite">
              <div className="docx-viewer__loading-spinner" aria-hidden="true" />
              <div className="docx-viewer__loading-title">{t('docx.viewer.layingOut')}</div>
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
                    {t('docx.viewer.blocksProgress', { completed: paginationProgress.completedBlocks, total: paginationProgress.totalBlocks })}
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
              aria-label={t('docx.viewer.tableEditingAria')}
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
        <DocxPromptDialog
          key={promptRequest?.id ?? 'docx-prompt-closed'}
          isOpen={promptRequest !== null}
          titleText={promptDialogCopy.titleText}
          label={promptDialogCopy.label}
          initialValue={promptDialogCopy.initialValue}
          confirmLabel={promptDialogCopy.confirmLabel}
          onConfirm={handlePromptConfirm}
          onCancel={closePrompt}
        />
      </div>
      {saveError !== null ? (
        <div className="docx-viewer__error" role="alert">
          <span>{saveError}</span>
          <button
            type="button"
            className="docx-viewer__error-dismiss"
            onClick={() => setSaveError(null)}
            aria-label={t('docx.viewer.dismissError')}
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
            onClick={dismissFieldUpdateMessage}
            aria-label={t('docx.viewer.dismissMessage')}
          >
            <X size={14} aria-hidden="true" />
          </button>
        </div>
      ) : null}
      {fidelityWarningMessage !== null ? (
        <div className="docx-viewer__field-status" role="status">
          <span>{fidelityWarningMessage}</span>
          <button
            type="button"
            className="docx-viewer__error-dismiss"
            onClick={dismissFidelityWarning}
            aria-label={t('docx.viewer.dismissMessage')}
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
  const t = useTranslate()
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
              const target = containerRef.current
                ?.querySelectorAll<HTMLElement>('.docx-page__line[data-paragraph-path]')
                .item(seed.paragraphIndex)
              if (target) scrollIntoViewRespectingMotionPreference(target, { behavior: 'smooth', block: 'start' })
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
    return <div className="docx-viewer docx-viewer--error">{t('docx.viewer.unexpectedTextFile')}</div>
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
