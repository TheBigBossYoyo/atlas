/**
 * Inserting a hyperlink over the current selection.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` alongside `useDocxPrompt`. It is
 * its own module, small as it is, because it is the one editor action that
 * mutates the BUNDLE and not just the document: a link needs a relationship id
 * in `word/_rels/document.xml.rels`, so `insertHyperlinkIntoBundle` hands back a
 * new bundle as well as a new document, and both have to land together.
 *
 * Two halves, matching the prompt flow: `handleInsertHyperlink` opens the dialog
 * (F1 — Electron does not implement `window.prompt`, so this used to do nothing
 * at all in the packaged app), and `confirmHyperlink` applies the URL the dialog
 * collected. The selection comes from the request the dialog was opened with,
 * captured at that moment, while the bundle and document are read here at confirm
 * time — which is exactly why this branch belongs in a hook with current
 * dependencies rather than in a closure stored when the dialog opened.
 */
import { useCallback } from 'react'

import { insertHyperlinkIntoBundle, type Command, type Range } from '../editor'
import type { DocxBundle } from '../index'
import type { Document as DocxDocument } from '../model'
import type { TranslateFn } from '../../i18n'
import type { DocxPromptRequestSpec } from './useDocxPrompt'

export type DocxHyperlink = {
  /** Toolbar/shortcut entry point: opens the prompt, or explains why it cannot. */
  readonly handleInsertHyperlink: () => void
  /** Applies the URL the prompt collected over the selection it was opened with. */
  readonly confirmHyperlink: (selection: Range, url: string) => void
}

export function useDocxHyperlink(
  bundle: DocxBundle,
  documentModel: DocxDocument,
  range: Range | null,
  commitState: (nextDocument: DocxDocument, nextRange: Range | null) => void,
  onBundleChange: (next: DocxBundle) => void,
  pushUndo: (inverse: Command) => void,
  requestPrompt: (spec: DocxPromptRequestSpec) => void,
  setSaveError: (message: string | null) => void,
  t: TranslateFn,
): DocxHyperlink {
  const handleInsertHyperlink = useCallback(() => {
    const selection = range
    if (selection === null) {
      setSaveError(t('docx.viewer.selectBeforeHyperlink'))
      return
    }

    requestPrompt({ kind: 'hyperlink', selection })
  }, [range, requestPrompt, setSaveError, t])

  const confirmHyperlink = useCallback(
    (selection: Range, url: string) => {
      // Matches the old `url === null` (Cancel) vs. empty-after-trim handling:
      // both are silent no-ops, never an insertion.
      const trimmedUrl = url.trim()
      if (trimmedUrl.length === 0) {
        return
      }

      try {
        const insertResult = insertHyperlinkIntoBundle(
          { ...bundle, document: documentModel },
          selection,
          trimmedUrl,
        )
        pushUndo(insertResult.inverse)
        onBundleChange(insertResult.bundle)
        commitState(insertResult.document, insertResult.range)
      } catch (error) {
        setSaveError(error instanceof Error ? error.message : String(error))
      }
    },
    [bundle, commitState, documentModel, onBundleChange, pushUndo, setSaveError],
  )

  return { handleInsertHyperlink, confirmHyperlink }
}
