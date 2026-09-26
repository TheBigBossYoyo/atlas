/**
 * What the DOCX toolbar's font picker offers (USR-11), and which of those the
 * open document actually uses.
 *
 * Extracted from `src/viewers/DocxViewer.tsx` as part of breaking up that file's
 * one very large component. Two lists, kept apart on purpose:
 *   - `documentFonts` — families this document references, so the picker can
 *     show them first rather than burying them in an alphabetical list of
 *     several hundred.
 *   - `availableFonts` — everything selectable: a fixed set of common Office
 *     families, the bundled substitutes, and whatever the OS reports.
 *
 * The OS enumeration is optional by design. It arrives over IPC, so it is absent
 * in a plain browser tab and can simply fail; either way the two fixed lists
 * still give a usable picker, which is why the failure path is a silent
 * `catch` rather than an error surfaced to the user.
 */
import { useEffect, useMemo, useState } from 'react'

import { collectReferencedFontFamilies, FONT_FAMILIES } from '../fonts'
import type { Document as DocxDocument } from '../model'

/** USR-11 — always offered in the font picker, even before (or without) system font enumeration. */
const COMMON_OFFICE_FONTS: ReadonlyArray<string> = [
  'Aptos', 'Arial', 'Calibri', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia',
  'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Impact', 'Lucida Console',
  'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana',
]

export type DocxFontChoices = {
  /** Families this document references. */
  readonly documentFonts: ReturnType<typeof collectReferencedFontFamilies>
  /** Every family the picker offers, de-duplicated and locale-sorted. */
  readonly availableFonts: ReadonlyArray<string>
}

export function useDocxFontChoices(documentModel: DocxDocument): DocxFontChoices {
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

  return { documentFonts, availableFonts }
}
