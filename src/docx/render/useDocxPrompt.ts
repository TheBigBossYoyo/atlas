/**
 * The DOCX viewer's single text-prompt dialog, and the pending request it shows.
 *
 * F1 — Electron does not implement `window.prompt` (it throws), so the three
 * actions that used to call it silently did nothing in the packaged app. They now
 * open one shared `DocxPromptDialog`, and this hook is the state machine behind
 * it: which request is pending, and the monotonic token that gives the dialog a
 * fresh React `key` per open so it never reuses the previous prompt's leftover
 * input value.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`. Only the machinery lives here —
 * what to DO with a submitted value belongs to whichever feature opened the
 * prompt (`useDocxComments` for comment/reply, `useDocxHyperlink` for a link),
 * so the viewer is left with a short dispatch on `kind` rather than a
 * three-branch handler that reached into three different collaborators.
 *
 * **The selection is captured at open time, not read at confirm time.** That
 * matches the old synchronous `window.prompt` timing, and it has to: the dialog
 * takes focus, so the document selection is not kept current while it is open.
 * Everything else (the document, the bundle) IS read at confirm time, by the
 * feature hooks' own callbacks — which is why those branches moved to the
 * features instead of into this module.
 */
import { useCallback, useRef, useState } from 'react'

import type { Range } from '../editor'

/** A pending prompt. `id` is the dialog's remount key — see this module's header. */
export type DocxPromptRequest =
  | { readonly kind: 'hyperlink'; readonly id: number; readonly selection: Range }
  | { readonly kind: 'comment'; readonly id: number; readonly selection: Range }
  | { readonly kind: 'reply'; readonly id: number; readonly commentId: string; readonly range: Range | null }

/** A request without its `id`; the hook allocates that. */
export type DocxPromptRequestSpec =
  | { readonly kind: 'hyperlink'; readonly selection: Range }
  | { readonly kind: 'comment'; readonly selection: Range }
  | { readonly kind: 'reply'; readonly commentId: string; readonly range: Range | null }

export type DocxPrompt = {
  /** The pending request, or null when the dialog is closed. */
  readonly promptRequest: DocxPromptRequest | null
  /** Opens the dialog for `spec`, allocating a fresh remount token. */
  readonly requestPrompt: (spec: DocxPromptRequestSpec) => void
  /**
   * Closes the dialog. The caller reads `promptRequest` itself before calling
   * this and then dispatches on it — closing before acting is deliberate, and
   * matches the previous ordering where the dialog was dismissed before any
   * mutation ran.
   */
  readonly closePrompt: () => void
}

export function useDocxPrompt(): DocxPrompt {
  const [promptRequest, setPromptRequest] = useState<DocxPromptRequest | null>(null)
  const promptIdRef = useRef(0)

  const requestPrompt = useCallback((spec: DocxPromptRequestSpec) => {
    promptIdRef.current += 1
    setPromptRequest({ ...spec, id: promptIdRef.current })
  }, [])

  const closePrompt = useCallback(() => {
    setPromptRequest(null)
  }, [])

  return { promptRequest, requestPrompt, closePrompt }
}
