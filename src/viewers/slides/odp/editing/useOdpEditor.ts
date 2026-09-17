/**
 * USR-16 (ODP) — editing session for OpenDocument presentations: the shared
 * slide-editor core plus the ODF edits. Everything lives in one part
 * (`content.xml`), so any edit re-parses the whole deck — ODP decks are small
 * enough that this is cheaper than tracking which slide changed.
 */
import type { SlideData } from '../../../shared/SlideDeck.types'
import { packageArchive, type OfficePackage } from '../../../../office/officePackage'
import { useSlideEditorCore } from '../../shared/useSlideEditorCore'
import { parseOdpSlides } from '../parser'
import * as edits from './odpEdits'

const NEW_TEXT_BOX_TEXT = 'Text'

function parseOdpDeck(pkg: OfficePackage): Promise<ReadonlyArray<SlideData>> {
  return parseOdpSlides(packageArchive(pkg), { cancelled: false })
}

export type SlideBox = edits.ShapeBox

export function useOdpEditor(buffer: ArrayBuffer | null, filePath: string) {
  const core = useSlideEditorCore({
    buffer,
    filePath,
    parseDeck: parseOdpDeck,
    saveFilter: { name: 'OpenDocument Presentation', extensions: ['odp'] },
  })

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
    /** ODF keeps notes inside the slide itself, so every slide can take them. */
    canEditNotes: (): boolean => true,
    setShapeText: (index: number, sourceId: string, text: string) =>
      core.apply((pkg) => edits.setShapeText(pkg, index, sourceId, text)),
    setShapeBox: (index: number, sourceId: string, box: SlideBox) =>
      core.apply((pkg) => edits.setShapeBox(pkg, index, sourceId, box)),
    deleteShape: (index: number, sourceId: string) => core.apply((pkg) => edits.deleteShape(pkg, index, sourceId)),
    setNotes: (index: number, text: string) => core.apply((pkg) => edits.setSlideNotes(pkg, index, text)),
    insertTextBox: async (index: number, box: SlideBox): Promise<string | null> => {
      let sourceId: string | null = null
      await core.apply((pkg) => {
        const result = edits.insertTextBox(pkg, index, box, NEW_TEXT_BOX_TEXT)
        sourceId = result.sourceId
        return result.pkg
      })
      return sourceId
    },
    addSlide: async (afterIndex: number): Promise<number> => {
      let index = afterIndex
      await core.apply((pkg) => {
        const result = edits.addSlide(pkg, afterIndex)
        index = result.index
        return result.pkg
      })
      return index
    },
    duplicateSlide: (index: number) => core.apply((pkg) => edits.duplicateSlide(pkg, index)),
    deleteSlide: (index: number) => core.apply((pkg) => edits.deleteSlide(pkg, index)),
    moveSlide: (from: number, to: number) => core.apply((pkg) => edits.moveSlide(pkg, from, to)),
  }
}

export type OdpEditor = ReturnType<typeof useOdpEditor>
