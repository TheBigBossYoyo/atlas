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

import { Printer, Save } from 'lucide-react'

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
  findAll,
  findNext,
  findParagraph,
  findPrev,
  handleBeforeInput,
  handleKeyDown,
  insertImageIntoBundle,
  buildPasteCommands,
  htmlToParagraphs,
  textToParagraphs,
  toolbarToCommand,
  useComposition,
  useEditor,
  useSpellCheck,
  type Command,
  type FindOptions,
  type ImageMimeType,
  type Range,
} from '../docx/editor'
import { addCommentToDocument, deleteCommentFromDocument, replyToComment } from '../docx/editor/commentMutations'
import { comparePositions, normalizeRange } from '../docx/editor/Selection'
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

function collectParagraphPaths(document: DocxDocument): ReadonlyArray<ReadonlyArray<number>> {
  const paths: Array<ReadonlyArray<number>> = []

  document.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind === 'paragraph') {
        paths.push(Object.freeze([sectionIndex, blockIndex]))
      }
    })
  })

  return paths
}

function getParagraphPathsForRange(
  document: DocxDocument,
  range: Range | null,
): ReadonlyArray<ReadonlyArray<number>> {
  if (range === null) {
    return []
  }

  const normalized = normalizeRange(range)
  const allPaths = collectParagraphPaths(document)

  return allPaths.filter(path => {
    const atStart = comparePositions({ ...normalized.start, runIndex: 0, charOffset: 0 }, { ...normalized.start, paragraphPath: path })
    const atEnd = comparePositions({ ...normalized.end, paragraphPath: path }, { ...normalized.end, runIndex: Number.MAX_SAFE_INTEGER, charOffset: Number.MAX_SAFE_INTEGER })
    return atStart <= 0 && atEnd <= 0
  })
}
void getParagraphPathsForRange

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

function createToolbarState(document: DocxDocument, range: Range | null): ToolbarState {
  const paragraph = range ? findParagraph(document, range.focus.paragraphPath) : null
  const activeFormats = new Set<'bold' | 'italic' | 'underline' | 'strike' | 'subscript' | 'superscript'>()

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
      spellCheck: true,
      trackChanges: false,
    }
  }

  return {
    activeFormats,
    alignment: null,
    fontFamily: null,
    fontSizePt: null,
    styleId: null,
    spellCheck: true,
    trackChanges: false,
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
  const composition = useComposition()
  const shadowEditor = useEditor(bundle.document)
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

  const toolbarState = useMemo(() => createToolbarState(documentModel, range), [documentModel, range])

  const availableStyles = useMemo(() => {
    return Array.from(documentModel.styles.values())
      .filter(style => style.type === 'paragraph')
      .map(style => ({ id: style.id, name: style.name ?? style.id }))
  }, [documentModel.styles])

  const commitState = useCallback((nextDocument: DocxDocument, nextRange: Range | null) => {
    setDocumentModel(nextDocument)
    setRange(nextRange)
    setSaveError(null)
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

        let nextRange = range
        if (command.kind === 'insert-text') {
          const nextPos = {
            paragraphPath: command.at.paragraphPath,
            runIndex: command.at.runIndex,
            charOffset: command.at.charOffset + command.text.length,
          }
          nextRange = { anchor: nextPos, focus: nextPos }
        } else if (command.kind === 'delete-range') {
          const normalized = normalizeRange(command.range)
          nextRange = { anchor: normalized.start, focus: normalized.start }
        } else if (command.kind === 'insert-paragraph-break') {
          const nextPath = [...command.at.paragraphPath]
          nextPath[nextPath.length - 1] = (nextPath[nextPath.length - 1] ?? 0) + 1
          const nextPos = { paragraphPath: Object.freeze(nextPath), runIndex: 0, charOffset: 0 }
          nextRange = { anchor: nextPos, focus: nextPos }
        }

        commitState(result.document, nextRange)
        return true
      } catch {
        return false
      }
    },
    [commitState, documentModel, range],
  )

  const applyEditorCommands = useCallback(
    (commands: ReadonlyArray<Command>, nextRange: Range | null): boolean => {
      let workingDocument = documentModel

      try {
        for (const command of commands) {
          const result = applyCommand(workingDocument, command)
          workingDocument = result.document
          historyRef.current.push(result.inverse)
        }

        commitState(workingDocument, nextRange)
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

    const focus = range?.focus ?? shadowEditor.range?.focus ?? null
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
      // Default sizing: 200pt wide, preserve a 4:3 ratio fallback until image
      // intrinsic dimensions are decoded (handled by future enhancement).
      const widthPt = 200
      const heightPt = 150

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

      onBundleChange(insertResult.bundle)
      commitState(insertResult.document, insertResult.range)
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error))
    }
  }, [bundle, commitState, documentModel, onBundleChange, range, shadowEditor.range])

  const handlePaste = useCallback(
    (event: ReactClipboardEvent<HTMLDivElement>) => {
      const clipboard = event.clipboardData
      if (clipboard === null) {
        return
      }

      const focus = range?.focus ?? shadowEditor.range?.focus ?? null
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
    [applyEditorCommands, range, shadowEditor.range],
  )

  const handleBeforeInputEvent = useCallback(
    (event: FormEvent<HTMLDivElement>) => {
      const nativeEvent = event.nativeEvent as InputEvent
      if (composition.shouldSwallow(nativeEvent)) {
        event.preventDefault()
        return
      }

      if (
        applyResult(
          handleBeforeInput(nativeEvent, {
            document: documentModel,
            range,
            history: historyRef.current,
          }),
        )
      ) {
        event.preventDefault()
      }
    },
    [applyResult, composition, documentModel, range],
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

      if (
        applyResult(
          handleKeyDown(event.nativeEvent, {
            document: documentModel,
            range,
            history: historyRef.current,
          }),
        )
      ) {
        event.preventDefault()
      }
    },
    [applyResult, documentModel, range],
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
        }),
      )
    },
    [applyResult, composition.handlers, documentModel, range],
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

  const handleSpellReplace = useCallback(
    (misspelled: string, replacement: string) => {
      if (misspelled.length === 0 || replacement === misspelled) {
        return
      }

      void spellCheck.replaceMisspelling(replacement)
    },
    [spellCheck],
  )

  const handleAddToDictionary = useCallback(
    (word: string) => {
      void spellCheck.addToDictionary(word)
    },
    [spellCheck],
  )

  const handleAddComment = useCallback(() => {
    const selection = range ?? shadowEditor.range
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
  }, [commitState, documentModel, range, shadowEditor.range])

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

      if (toolbarCommand.kind === 'toggle-spell-check') {
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

      const selection = range ?? shadowEditor.range
      const command = toolbarToCommand(toolbarCommand, selection)
      if (command !== null) {
        applyEditorCommand(command)
      }
    },
    [applyEditorCommand, applyResult, documentModel, handleAddComment, handleInsertImage, range, shadowEditor.range],
  )

  // P1.1 — returns whether the save actually succeeded, so both the shared
  // document-session `save()` contract (App.tsx's global Ctrl+S/Save and the
  // unsaved-changes confirmation dialog) and the Save button here can tell.
  const handleSave = useCallback(async (): Promise<boolean> => {
    try {
      const nextBytes = await saveDocx({
        ...bundle,
        document: documentModel,
      })

      const result = await window.electronAPI?.saveBinaryFile?.({
        content: nextBytes,
        suggestedName: getSuggestedFileName(savePath),
        existingPath: savePath,
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
      setSaveError(error instanceof Error ? error.message : String(error))
      return false
    }
  }, [bundle, documentModel, savePath])

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
  // focused. Detection-only: it never performs the actual edit/print/save
  // itself (that stays solely in `handleKeyDownEvent`/commands.ts, reached
  // through React's own bubble phase, which always runs first when the
  // contentEditable has focus) — this only prevents the *shell* from also
  // reacting to the same keystroke.
  useViewerShortcuts(
    useCallback((event) => {
      const ctrl = event.ctrlKey || event.metaKey
      if (!ctrl) return false
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
    shadowEditor.setRange(range)
  }, [range, shadowEditor])

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
          console.info(
            `[Atlas/DocxViewer] paginate returned ${nextPages.length} page(s) for document with ${documentModel.sections.length} section(s)`,
          )
          setPages(nextPages)
          setPaginationProgress(null)
        }
      } catch (error) {
        if (error instanceof PaginationCancelledError) {
          return
        }
        if (!cancelled) {
          console.error('[Atlas/DocxViewer] pagination failed', error)
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
          spellCheck={true}
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
              <PageStack pages={pages} zoom={1} document={documentModel} theme={bundle.theme} />
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
      {saveError !== null ? <div className="docx-viewer__error">{saveError}</div> : null}
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
          setErrorMessage(error instanceof Error ? error.message : String(error))
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
    return <div className="docx-viewer docx-viewer--error">Failed to render DOCX: {errorMessage}</div>
  }

  return (
    <div ref={containerRef} className="docx-viewer">
      {bundle !== null ? <DocxEditor bundle={bundle} file={file} containerRef={containerRef} onBundleChange={setBundle} /> : null}
    </div>
  )
}

export const DocxViewer = memo(DocxViewerBase)
