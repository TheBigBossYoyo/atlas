/**
 * Find & replace for the DOCX editor: the query/options the Find bar holds, the
 * matches they resolve to, and the six actions the bar can fire.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that file's
 * one very large component. The search itself lives in `docx/editor/find.ts` and
 * the replacements are ordinary editor commands — what was tangled into the
 * component was only the state around them, which turns out to touch just three
 * things it does not own: the current `range` (every action either reads it as a
 * starting point or replaces it), the editor's command applier, and the request
 * to scroll a newly-selected match into view.
 *
 * Two deliberate details carried over unchanged:
 *   - `findNext`/`findPrev` start from `range.focus`/`range.anchor` respectively,
 *     not from a stored match index, so Find stays correct when the user clicks
 *     elsewhere or types between presses.
 *   - the replacement text is read out of the DOM at the moment of the action
 *     rather than mirrored into React state, because the Find bar's replace
 *     field is uncontrolled. See `getReplaceValue` in `./selectionDom`.
 */
import { useCallback, useMemo, useState, type RefObject } from 'react'

import {
  buildReplaceAllCommands,
  buildReplaceCommands,
  findAll,
  findNext,
  findPrev,
  type Command,
  type FindOptions,
} from '../editor'
import { rangeEquals } from '../editor/Selection'
import type { Range } from '../editor'
import type { Document as DocxDocument } from '../model'
import { getReplaceValue } from './selectionDom'

const DEFAULT_FIND_OPTIONS: FindOptions = {
  caseSensitive: false,
  wholeWord: false,
  useRegex: false,
}

/** Which match in `matches` the current selection is sitting on, or null. */
function getCurrentMatchIndex(matches: ReadonlyArray<{ range: Range }>, range: Range | null): number | null {
  if (range === null) {
    return null
  }

  const index = matches.findIndex((match) => rangeEquals(match.range, range))
  return index >= 0 ? index : null
}

export type DocxFind = {
  readonly findOpen: boolean
  readonly setFindOpen: (open: boolean) => void
  readonly matches: ReturnType<typeof findAll>
  readonly currentMatchIndex: number | null
  readonly handleFind: (query: string, options: FindOptions) => void
  readonly handleFindNext: () => void
  readonly handleFindPrev: () => void
  readonly handleReplace: () => void
  readonly handleReplaceAll: () => void
}

/**
 * @param documentModel  - the live model to search.
 * @param range          - the current selection; the starting point for
 *                         next/previous and the target for replace.
 * @param setRange       - moves the selection onto a found match.
 * @param containerRef   - the viewer root, for reading the uncontrolled replace
 *                         field out of the DOM.
 * @param applyEditorCommands - the editor's command applier, for the two
 *                         replace actions.
 * @param revealSelectionOnNextPaint - from `useDocxSelectionPainting`; called
 *                         before moving the selection so the match is scrolled
 *                         into view once it has actually rendered.
 */
export function useDocxFind(
  documentModel: DocxDocument,
  range: Range | null,
  setRange: (range: Range | null) => void,
  containerRef: RefObject<HTMLDivElement | null>,
  applyEditorCommands: (commands: ReadonlyArray<Command>, nextRange: Range | null) => void,
  revealSelectionOnNextPaint: () => void,
): DocxFind {
  const [findOpen, setFindOpen] = useState(false)
  const [findQuery, setFindQuery] = useState('')
  const [findOptions, setFindOptions] = useState<FindOptions>(DEFAULT_FIND_OPTIONS)

  const matches = useMemo(() => {
    if (findQuery.length === 0) {
      return []
    }

    return findAll(documentModel, findQuery, findOptions)
  }, [documentModel, findOptions, findQuery])

  const currentMatchIndex = useMemo(() => getCurrentMatchIndex(matches, range), [matches, range])

  const handleFind = useCallback(
    (query: string, options: FindOptions) => {
      setFindQuery(query)
      setFindOptions(options)

      if (query.length === 0) {
        return
      }

      const nextMatch = findNext(documentModel, query, options, null)
      if (nextMatch !== null) {
        revealSelectionOnNextPaint()
        setRange(nextMatch.range)
      }
    },
    [documentModel, revealSelectionOnNextPaint, setRange],
  )

  const handleFindNext = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const nextMatch = findNext(documentModel, findQuery, findOptions, range?.focus ?? null)
    if (nextMatch !== null) {
      revealSelectionOnNextPaint()
      setRange(nextMatch.range)
    }
  }, [documentModel, findOptions, findQuery, range, revealSelectionOnNextPaint, setRange])

  const handleFindPrev = useCallback(() => {
    if (findQuery.length === 0) {
      return
    }

    const prevMatch = findPrev(documentModel, findQuery, findOptions, range?.anchor ?? null)
    if (prevMatch !== null) {
      revealSelectionOnNextPaint()
      setRange(prevMatch.range)
    }
  }, [documentModel, findOptions, findQuery, range, revealSelectionOnNextPaint, setRange])

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

  return {
    findOpen,
    setFindOpen,
    matches,
    currentMatchIndex,
    handleFind,
    handleFindNext,
    handleFindPrev,
    handleReplace,
    handleReplaceAll,
  }
}
