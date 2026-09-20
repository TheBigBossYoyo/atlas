/**
 * USR-16 — PPTX editing session: the shared slide-editor core (undo/redo,
 * edit queueing, save/dirty) plus the PPTX-specific edits, and a re-parse
 * that only touches the slides whose parts actually changed.
 */
import { useCallback } from 'react'

import type { SlideData } from '../../../shared/SlideDeck.types'
import { packageArchive, readPart, type OfficePackage } from '../../../../office/officePackage'
import { useSlideEditorCore, type DeckState } from '../../shared/useSlideEditorCore'
import { parseOneSlide, resolveSlidePaths } from '../parser'
import { REL_TYPE, readRels, relsPathFor } from './opcXml'
import * as edits from './pptxEdits'
import * as slideOps from './pptxSlideOps'

// Empty, not a filler word: insertTextBox() now opens the new shape for
// editing immediately (see SlideDeck's pendingTextBoxSourceId), and its text
// editor places the caret at the END of any existing text — a filler word
// here would sit in front of whatever the user types next instead of being
// replaced by it.
const NEW_TEXT_BOX_TEXT = ''

/** The parts a parsed slide depends on; unchanged references mean the cached parse is still valid. */
function slideInputs(pkg: OfficePackage, path: string): ReadonlyArray<unknown> {
  const relsPath = relsPathFor(path)
  const notesPath = readRels(readPart(pkg, relsPath), path).find((rel) => rel.type === REL_TYPE.notesSlide)?.target
  return [pkg.parts.get(path), pkg.parts.get(relsPath), notesPath ? pkg.parts.get(notesPath) : undefined]
}

async function parsePptxDeck(pkg: OfficePackage, previous: DeckState | null): Promise<ReadonlyArray<SlideData>> {
  const archive = packageArchive(pkg)
  const signal = { cancelled: false }
  const { slidePaths, size } = await resolveSlidePaths(archive, signal)
  return Promise.all(
    slidePaths.map(async (path, index) => {
      const cached = previous?.slides.find((slide) => slide.partPath === path)
      const unchanged =
        cached !== undefined &&
        previous !== null &&
        slideInputs(pkg, path).every((part, i) => part === slideInputs(previous.pkg, path)[i])
      if (unchanged && cached) return { ...cached, index, id: `slide-${index + 1}` }
      return parseOneSlide(archive, path, index, size, signal)
    }),
  )
}

export type SlideBox = edits.ShapeBox

export function usePptxEditor(buffer: ArrayBuffer | null, filePath: string) {
  const core = useSlideEditorCore({
    buffer,
    filePath,
    parseDeck: parsePptxDeck,
    saveFilter: { name: 'PowerPoint Presentation', extensions: ['pptx'] },
  })

  const onSlide = useCallback(
    (index: number, edit: (pkg: OfficePackage, path: string) => OfficePackage) =>
      core.apply((pkg) => {
        const path = core.current().slides[index]?.partPath
        return path ? edit(pkg, path) : pkg
      }),
    [core],
  )

  return {
    status: core.status,
    error: core.error,
    saveError: core.saveError,
    slides: core.slides,
    canUndo: core.canUndo,
    canRedo: core.canRedo,
    undo: core.undo,
    redo: core.redo,
    save: core.save,
    saveAs: core.saveAs,
    canEditNotes: (index: number): boolean => {
      const { pkg, slides } = core.current()
      const path = slides[index]?.partPath
      return path ? edits.canEditNotes(pkg, path) : false
    },
    setShapeText: (index: number, sourceId: string, text: string) =>
      onSlide(index, (pkg, path) => edits.setShapeText(pkg, path, sourceId, text)),
    setShapeBox: (index: number, sourceId: string, box: SlideBox) =>
      onSlide(index, (pkg, path) => edits.setShapeBox(pkg, path, sourceId, box)),
    deleteShape: (index: number, sourceId: string) =>
      onSlide(index, (pkg, path) => edits.deleteShape(pkg, path, sourceId)),
    setNotes: (index: number, text: string) => onSlide(index, (pkg, path) => edits.setSlideNotes(pkg, path, text)),
    /** Resolves the new text box's shape id once it is on the slide. */
    insertTextBox: async (index: number, box: SlideBox): Promise<string | null> => {
      let sourceId: string | null = null
      await onSlide(index, (pkg, path) => {
        const result = edits.insertTextBox(pkg, path, box, NEW_TEXT_BOX_TEXT)
        sourceId = result.sourceId
        return result.pkg
      })
      return sourceId
    },
    /** Resolves the new slide's index. */
    addSlide: async (afterIndex: number): Promise<number> => {
      let index = afterIndex
      await core.apply((pkg) => {
        const result = slideOps.addSlide(pkg, afterIndex)
        index = result.index
        return result.pkg
      })
      return index
    },
    duplicateSlide: (index: number) => core.apply((pkg) => slideOps.duplicateSlide(pkg, index)),
    deleteSlide: (index: number) => core.apply((pkg) => slideOps.deleteSlide(pkg, index)),
    moveSlide: (from: number, to: number) => core.apply((pkg) => slideOps.moveSlide(pkg, from, to)),
  }
}

export type PptxEditor = ReturnType<typeof usePptxEditor>
