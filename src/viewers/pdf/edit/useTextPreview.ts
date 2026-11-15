/**
 * The live preview of text edits (PDF-EDIT-2): the real edited file, rebuilt
 * in the background whenever the text edits change and opened as a second
 * pdf.js document, so an edited page shows exactly what saving will write —
 * the document's own font, the re-wrapped lines, the old words gone —
 * instead of an approximation drawn on top.
 *
 * Rebuilding takes a moment (it rewrites the whole file), so the preview
 * records which edits it contains. A page is drawn from it only while that
 * page's edits are exactly the current ones; in between, the edit layer draws
 * a quick stand-in for the edits the preview doesn't have yet.
 */

import { useEffect, useRef, useState } from 'react'

import type { PdfDocument } from '../types'
import type { TextEdit } from './types'

/** The part of pdf.js's module this needs. */
export type PreviewOpener = (bytes: Uint8Array) => { readonly promise: Promise<PdfDocument>; destroy(): Promise<void> }

export type TextPreview = {
  readonly doc: PdfDocument
  readonly edits: ReadonlyArray<TextEdit>
}

const REBUILD_DELAY_MS = 250

export function useTextPreview(
  open: PreviewOpener | null,
  textEdits: ReadonlyArray<TextEdit>,
  build: (edits: ReadonlyArray<TextEdit>) => Promise<Uint8Array | null>,
  onError: (error: unknown) => void,
): TextPreview | null {
  const [preview, setPreview] = useState<TextPreview | null>(null)
  const destroyRef = useRef<(() => Promise<void>) | null>(null)
  const buildRef = useRef(build)
  const onErrorRef = useRef(onError)
  useEffect(() => {
    buildRef.current = build
    onErrorRef.current = onError
  }, [build, onError])

  useEffect(() => {
    if (!open || textEdits.length === 0) return undefined
    let cancelled = false
    const timer = setTimeout(() => {
      void (async () => {
        try {
          const bytes = await buildRef.current(textEdits)
          if (cancelled || !bytes) return
          const task = open(bytes)
          const doc = await task.promise
          if (cancelled) {
            void task.destroy()
            return
          }
          const previous = destroyRef.current
          destroyRef.current = () => task.destroy()
          setPreview({ doc, edits: textEdits })
          // The old preview may still be on screen for a frame; let React swap first.
          setTimeout(() => void previous?.(), 0)
        } catch (error) {
          if (!cancelled) onErrorRef.current(error)
        }
      })()
    }, REBUILD_DELAY_MS)
    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [open, textEdits])

  useEffect(
    () => () => {
      void destroyRef.current?.()
      destroyRef.current = null
    },
    [],
  )

  // With no text edits left (all undone), there is nothing to preview; the
  // stale document is simply not offered, and is destroyed when replaced.
  return textEdits.length === 0 ? null : preview
}

/**
 * Whether `sourcePage` should be drawn from the preview: it has text edits,
 * and the preview has exactly those. Unedited pages always stay on the
 * original document, so a rebuilt preview never re-renders them.
 */
export function previewCovers(
  preview: TextPreview | null,
  current: ReadonlyArray<TextEdit>,
  sourcePage: number,
): boolean {
  if (!preview) return false
  const mine = current.filter((edit) => edit.page === sourcePage)
  if (mine.length === 0) return false
  const theirs = preview.edits.filter((edit) => edit.page === sourcePage)
  return mine.length === theirs.length && mine.every((edit) => theirs.includes(edit))
}
