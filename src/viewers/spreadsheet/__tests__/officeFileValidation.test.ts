/**
 * Wires `scripts/lib/officeValidator.mjs` (the spec-level OPC/OOXML/ODF
 * validation harness) into CI against Atlas's real spreadsheet save paths:
 * `writeWorkbookThroughOriginal` (save-through-original, USR-17) after a mix
 * of structural edits, and `writeWorkbookBytesWithTables` (the fresh-
 * workbook writer) for every bookType it's asked to produce. This is what
 * caught `.ods` saves shipping with `mimetype` neither first nor stored
 * (fixed in `spreadsheetWrite.ts`'s `fixOdsPackaging`) — a defect none of
 * this file's existing round-trip-fidelity tests could have noticed, since
 * they only ever inspect cell values through `XLSX.read`, never the raw
 * zip's physical layout.
 */
import { describe, expect, it } from 'vitest'

import { attachSheetSources, attachTables, parseWorkbookBuffer } from '../../shared/spreadsheetGrid'
import {
  addSheet,
  createDocument,
  deleteColumnAt,
  deleteRowAt,
  deleteSheet,
  insertRowAt,
  renameSheet,
  setCellValue,
  type SpreadsheetDocument,
} from '../spreadsheetDocument'
import { readSheetPartPaths, readSheetTables } from '../spreadsheetTables'
import { writeWorkbookBytesWithTables } from '../spreadsheetWrite'
import { writeWorkbookThroughOriginal } from '../xlsxPassthrough'
import {
  buildCrossSheetFormulaWorkbook,
  buildMultiSheetWorkbook,
  buildSharedFormulaWorkbook,
  buildSharedStringsWorkbook,
  buildStyledWorkbook,
  buildWorkbookWithOwnedParts,
  buildWorkbookWithSharedImage,
} from './styledWorkbook'

import { validateOfficeFile } from '../../../../scripts/lib/officeValidator.mjs'

async function load(buffer: ArrayBuffer): Promise<SpreadsheetDocument> {
  const [tables, partPaths] = await Promise.all([readSheetTables(buffer), readSheetPartPaths(buffer)])
  return createDocument(attachSheetSources(attachTables(parseWorkbookBuffer(buffer), tables), partPaths))
}

function errorsOnly(issues: ReadonlyArray<{ readonly severity: string }>) {
  return issues.filter((i) => i.severity === 'error')
}

describe('spreadsheet saves pass spec-level OPC/ODF validation', () => {
  it('writeWorkbookThroughOriginal after an edited cell', async () => {
    const original = await buildStyledWorkbook()
    const doc = setCellValue(await load(original), 0, 1, 0, 'Recycled paper')
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(result.format).toEqual({ family: 'opc', kind: 'xlsx' })
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after a row insert + column delete', async () => {
    const original = await buildStyledWorkbook()
    let doc = await load(original)
    doc = insertRowAt(doc, 0, 1)
    doc = deleteColumnAt(doc, 0, 2)
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after adding, renaming and deleting sheets', async () => {
    const original = await buildMultiSheetWorkbook()
    let doc = await load(original)
    doc = addSheet(doc, 'Fresh Sheet')
    doc = setCellValue(doc, doc.sheets.length - 1, 0, 0, 'New content')
    doc = renameSheet(doc, 0, 'Renamed First')
    doc = deleteSheet(doc, 1)
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after a cross-sheet formula rename, delete and row insert', async () => {
    const original = await buildCrossSheetFormulaWorkbook()
    let doc = await load(original)
    doc = renameSheet(doc, 1, 'Info') // Notes -> Info: Budget!D2's "=Notes!A1" must follow
    doc = insertRowAt(doc, 0, 1) // Budget's own row insert: Notes/Info!A2's "=Budget!C2" must follow
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after deleting a sheet a formula referenced (#REF!)', async () => {
    const original = await buildCrossSheetFormulaWorkbook()
    const doc = deleteSheet(await load(original), 1) // delete Notes — Budget!D2's formula becomes #REF!
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after deleting a sheet with its own table/comments/drawing/chart/image', async () => {
    const original = await buildWorkbookWithOwnedParts()
    const doc = deleteSheet(await load(original), 1) // delete "Extra" — sweeps its owned parts, including its now-unreferenced image
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal after deleting a sheet whose image is still referenced by a surviving sheet\'s drawing', async () => {
    const original = await buildWorkbookWithSharedImage()
    const doc = deleteSheet(await load(original), 1) // delete "Extra" — "Main" keeps its own drawing on the same image
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal re-anchors a shared-formula group\'s ref span through a row insert', async () => {
    const original = await buildSharedFormulaWorkbook()
    let doc = await load(original)
    doc = insertRowAt(doc, 0, 0)
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal re-anchors a shared-formula group\'s ref span through a row delete inside the group', async () => {
    const original = await buildSharedFormulaWorkbook()
    let doc = await load(original)
    doc = deleteRowAt(doc, 0, 2)
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it('writeWorkbookThroughOriginal compacts sharedStrings.xml without corrupting the package', async () => {
    const original = await buildSharedStringsWorkbook()
    const doc = await load(original)
    const bytes = await writeWorkbookThroughOriginal(original, doc)
    expect(bytes).not.toBeNull()

    const result = validateOfficeFile(Buffer.from(bytes!))
    expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])
  })

  it.for(['xlsx', 'xlsm', 'xlsb', 'ods'] as const)(
    'writeWorkbookBytesWithTables — fresh %s workbook',
    async (bookType) => {
      const original = await buildMultiSheetWorkbook()
      let doc = await load(original)
      doc = setCellValue(doc, 0, 0, 0, 'Fresh-written value')
      doc = addSheet(doc, 'Extra')

      const bytes = await writeWorkbookBytesWithTables(doc, bookType)
      const result = validateOfficeFile(Buffer.from(bytes))
      expect(errorsOnly(result.issues), JSON.stringify(result.issues, null, 2)).toEqual([])

      if (bookType === 'ods') {
        expect(result.format).toMatchObject({ family: 'odf', kind: 'ods' })
        // Regression for the exact defect this harness found: every `.ods`
        // Atlas saves goes through this path (no OOXML original to patch
        // through for a non-xlsx/xlsm target), and SheetJS's own
        // `bookType: 'ods'` writer does not put `mimetype` first-and-stored
        // — see `fixOdsPackaging`'s header in `spreadsheetWrite.ts`.
        expect(result.issues.filter((i) => i.code.startsWith('odf-mimetype'))).toEqual([])
      } else {
        expect(result.format.family).toBe('opc')
      }
    },
  )
})
