/**
 * "Update field(s)" and "Update table of contents" for the DOCX viewer, plus the
 * inline status line both report through.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`, the last cluster to come out of
 * that component. Both actions apply their result straight to the document,
 * deliberately bypassing the editor's Command/History system — which is why they
 * take `applyFieldUpdate` rather than `commitState`: they must bump the revision
 * counter by hand (otherwise the dirty check never sees the change, DIRTY-1) and
 * must NOT move the selection, which `commitState` would.
 *
 * Routing these through a real `replace-blocks`-shaped command so they became
 * undoable is a known follow-up, left where it was: out of scope here, and
 * unchanged by this extraction.
 */
import { useCallback, useState } from 'react'

import {
  collectBookmarkMaps,
  parseCoreProps,
  updateFields,
  updateTableOfContents,
  type FieldEvaluationContext,
} from '../fields'
import type { DocxBundle } from '../index'
import type { Page } from '../layout'
import type { Document as DocxDocument } from '../model'
import { useTranslate } from '../../i18n'

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

export type DocxFields = {
  /** Inline status for the last field action, or null. */
  readonly fieldUpdateMessage: string | null
  readonly dismissFieldUpdateMessage: () => void
  readonly handleUpdateFields: () => void
  readonly handleUpdateTableOfContents: () => void
}

/**
 * @param bundle           - for `docProps/core.xml` (AUTHOR/TITLE fields).
 * @param documentModel    - the live model.
 * @param pages            - the current pagination, for PAGE/NUMPAGES and TOC
 *                           page numbers. `null` before the first pass, which
 *                           simply leaves those fields unresolved.
 * @param applyFieldUpdate - replaces the document and bumps the revision,
 *                           without touching the selection.
 */
export function useDocxFields(
  bundle: DocxBundle,
  documentModel: DocxDocument,
  pages: ReadonlyArray<Page> | null,
  applyFieldUpdate: (next: DocxDocument) => void,
): DocxFields {
  const t = useTranslate()
  const [fieldUpdateMessage, setFieldUpdateMessage] = useState<string | null>(null)

  const dismissFieldUpdateMessage = useCallback(() => {
    setFieldUpdateMessage(null)
  }, [])

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
      setFieldUpdateMessage(t('docx.viewer.noFieldsToUpdate'))
      return
    }

    // DIRTY-1 — a field update bypasses History entirely (see this module's
    // header), so `applyFieldUpdate` has to bump the revision by hand or the
    // dirty check never notices the change.
    applyFieldUpdate(updated)
    setFieldUpdateMessage(t('docx.viewer.fieldsUpdated', { count: updatedCount }))
  }, [applyFieldUpdate, bundle.rawArchive, documentModel, pages, t])

  /**
   * DEFER-5 / DXS-20 — "Update table of contents": regenerates a
   * single-paragraph TOC field's entries from the document's current
   * headings (see `toc.ts`'s doc comment on that scope). No-ops with a
   * status message when the document has no such field, matching Word's
   * own behavior of the command doing nothing without a TOC.
   */
  const handleUpdateTableOfContents = useCallback(() => {
    const pageOfParagraph = pages !== null ? buildPageOfParagraph(pages) : undefined
    const { document: updated, updated: didUpdate, reason } = updateTableOfContents(
      documentModel,
      pageOfParagraph,
    )

    if (!didUpdate) {
      // DEFER-5 — each of these used to report "No table of contents found",
      // which on a real Word document is untrue: the document HAS one, Atlas
      // just cannot rewrite a field that spans paragraphs. Telling the user
      // that is worth more than a message that makes them doubt their file.
      setFieldUpdateMessage(
        t(
          reason === 'unsupported-span'
            ? 'docx.viewer.tocSpanUnsupported'
            : reason === 'locked'
              ? 'docx.viewer.tocLocked'
              : 'docx.viewer.noTocFound',
        ),
      )
      return
    }

    // DIRTY-1 — see handleUpdateFields above: bypasses History, so
    // `applyFieldUpdate` bumps the revision by hand.
    applyFieldUpdate(updated)
    setFieldUpdateMessage(t('docx.viewer.tocUpdated'))
  }, [applyFieldUpdate, documentModel, pages, t])

  return { fieldUpdateMessage, dismissFieldUpdateMessage, handleUpdateFields, handleUpdateTableOfContents }
}
