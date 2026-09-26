/**
 * The DOCX viewer's pagination session: laid-out pages, and the fonts they were
 * measured against.
 *
 * Extracted from `src/viewers/DocxViewer.tsx`, where this was two `useEffect`s
 * and four `useState`s sitting in a ~2,600-line component among twenty other
 * concerns. It comes out cleanly because nothing else in that component ever
 * writes any of this state, and because the font halves and the pagination half
 * are only coupled to each other: `embeddedFonts`/`fontResolver`/
 * `embeddedFontsReadyRef` had no other reader in the file. Keeping the coupling
 * inside one hook is the point — the ordering constraint between registering a
 * font and measuring with it is a real invariant, and it is easier to keep true
 * when both halves are ten lines apart.
 *
 * Behaviour is unchanged from the inline version, deliberately: same effects,
 * same dependency arrays, same batching of `pages`/`pagesDocument`.
 */
import { useEffect, useMemo, useRef, useState } from 'react'

import { createFontResolver, preloadDocxFonts, registerEmbeddedFonts, unregisterEmbeddedFonts } from '../fonts/register'
import { loadEmbeddedFonts } from '../fonts'
import { paginate, PaginationCancelledError } from '../layout'
import type { Page, PaginationProgress } from '../layout'
import type { Document as DocxDocument } from '../model'
import type { DocxBundle } from '../index'

export type DocxPagination = {
  /** `null` until the first pagination pass completes. */
  readonly pages: ReadonlyArray<Page> | null
  /** The exact document `pages` was laid out against — see its own note below. */
  readonly pagesDocument: DocxDocument
  readonly paginationProgress: PaginationProgress | null
  readonly paginationError: string | null
}

/**
 * @param documentModel - the live model; every change re-paginates (the
 *                        per-paragraph line cache in `layout/lineCache.ts` is
 *                        what keeps that affordable on a keystroke).
 * @param bundle        - the parsed file, for `rawArchive` (embedded fonts),
 *                        `theme` and `settings`.
 */
export function useDocxPagination(documentModel: DocxDocument, bundle: DocxBundle): DocxPagination {
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
  const [paginationProgress, setPaginationProgress] = useState<PaginationProgress | null>(null)
  const [paginationError, setPaginationError] = useState<string | null>(null)

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

    void (async () => {
      // Clearing the previous pass's error/progress belongs to THIS pass, and it
      // still happens in the same synchronous tick the effect runs in — an async
      // IIFE body runs synchronously up to its first `await`, which is below.
      // Nested here rather than at the effect's top level for the reason
      // `components/markdown/useMarkdownHastTree.ts` documents for its own
      // identical reset: at the top level it reads as a lint-flagged "setState
      // directly in an effect body" (react-hooks/set-state-in-effect) instead of
      // as part of the pagination this effect exists to start. Moving it changes
      // when nothing; leaving it out would leave a stale error from the previous
      // document on screen while the new one paginates.
      setPaginationError(null)
      setPaginationProgress(null)

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

  return { pages, pagesDocument, paginationProgress, paginationError }
}
