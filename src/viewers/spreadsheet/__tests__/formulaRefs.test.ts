/**
 * USR-17 follow-up — formula-reference-syntax rewriting through a renamed,
 * reordered or deleted sheet, and through a row/column insert or delete.
 */
import { describe, expect, it } from 'vitest'

import { rewriteFormulaReferences, type FormulaRewriteOptions, type SheetChange } from '../formulaRefs'

function change(overrides: Partial<SheetChange> = {}): SheetChange {
  return { newName: undefined, deleted: false, rowSources: undefined, colSources: undefined, ...overrides }
}

/** `sources[newIndex] = originalIndex | null` — same convention as `spreadsheetRangeShift.test.ts`. */
function sourcesAfter(originalCount: number, deletedOriginal: ReadonlyArray<number>, insertedAt: ReadonlyArray<number>): (number | null)[] {
  const survivors = Array.from({ length: originalCount }, (_, i) => i).filter((i) => !deletedOriginal.includes(i))
  const result: (number | null)[] = [...survivors]
  for (const at of [...insertedAt].sort((a, b) => b - a)) result.splice(at, 0, null)
  return result
}

function rewrite(formula: string, changes: ReadonlyMap<string, SheetChange>, overrides: Partial<FormulaRewriteOptions> = {}): string {
  const options: FormulaRewriteOptions = { changesByOriginalName: changes, remapCoordinates: true, ...overrides }
  return rewriteFormulaReferences(formula, options)
}

describe('rewriteFormulaReferences — sheet rename/delete', () => {
  it('rewrites a plain sheet-qualified single-cell reference on rename', () => {
    const changes = new Map([['Sheet2', change({ newName: 'Data' })]])
    expect(rewrite('Sheet2!A1', changes)).toBe('Data!A1')
  })

  it('rewrites a sheet-qualified range on rename', () => {
    const changes = new Map([['Sheet2', change({ newName: 'Data' })]])
    expect(rewrite('Sheet2!A1:B2', changes)).toBe('Data!A1:B2')
  })

  it('quotes the new name when it needs quoting (contains a space)', () => {
    const changes = new Map([['Sheet2', change({ newName: 'My Sheet' })]])
    expect(rewrite('Sheet2!A1', changes)).toBe("'My Sheet'!A1")
  })

  it('drops unnecessary quoting when the new name no longer needs it', () => {
    const changes = new Map([['My Sheet', change({ newName: 'Data' })]])
    expect(rewrite("'My Sheet'!A1", changes)).toBe('Data!A1')
  })

  it('decodes an escaped quote in the original quoted sheet name and re-encodes it in the new one', () => {
    const changes = new Map([["It's Mine", change({ newName: "Even Weirder's" })]])
    expect(rewrite("'It''s Mine'!A1", changes)).toBe("'Even Weirder''s'!A1")
  })

  it('replaces the whole reference with #REF! when the sheet was deleted', () => {
    const changes = new Map([['Sheet2', change({ deleted: true })]])
    expect(rewrite('Sheet2!A1:B2', changes)).toBe('#REF!')
  })

  it('leaves a reference to a sheet not in the change map untouched (already current, or external)', () => {
    const changes = new Map([['Sheet2', change({ newName: 'Data' })]])
    expect(rewrite('Sheet3!A1', changes)).toBe('Sheet3!A1')
  })

  it('leaves an unrenamed, undeleted sheet reference byte-for-byte untouched', () => {
    const changes = new Map([['Sheet1', change()]])
    expect(rewrite('Sheet1!A1', changes)).toBe('Sheet1!A1')
  })

  it('rewrites two independent references in one formula', () => {
    const changes = new Map([
      ['Sheet2', change({ newName: 'Data' })],
      ['Sheet3', change({ deleted: true })],
    ])
    expect(rewrite('Sheet2!A1+Sheet3!B2', changes)).toBe('Data!A1+#REF!')
  })
})

describe('rewriteFormulaReferences — 3-D references', () => {
  it('rewrites both endpoints of a 3-D reference independently on rename', () => {
    const changes = new Map([
      ['Sheet1', change({ newName: 'First' })],
      ['Sheet3', change({ newName: 'Last' })],
    ])
    expect(rewrite('SUM(Sheet1:Sheet3!A1)', changes)).toBe('SUM(First:Last!A1)')
  })

  it('quotes the whole "Name1:Name2" span together when either side needs it', () => {
    const changes = new Map([
      ['Sheet1', change({ newName: 'First One' })],
      ['Sheet3', change()],
    ])
    expect(rewrite('Sheet1:Sheet3!A1', changes)).toBe("'First One:Sheet3'!A1")
  })

  it('replaces the whole 3-D reference with #REF! when either endpoint sheet was deleted', () => {
    const changes = new Map([
      ['Sheet1', change()],
      ['Sheet3', change({ deleted: true })],
    ])
    expect(rewrite('Sheet1:Sheet3!A1', changes)).toBe('#REF!')
  })

  it('re-anchors the range of a 3-D reference using the first sheet\'s own row/column history', () => {
    const rowSources = sourcesAfter(3, [], [1]) // a row inserted at index 1 in the first sheet
    const changes = new Map([
      ['Sheet1', change({ rowSources })],
      ['Sheet3', change()],
    ])
    expect(rewrite('Sheet1:Sheet3!A2', changes)).toBe('Sheet1:Sheet3!A3')
  })
})

describe('rewriteFormulaReferences — coordinate re-anchoring (remapCoordinates: true)', () => {
  it('shifts a sheet-qualified single cell through a row insert on the TARGET sheet', () => {
    const rowSources = sourcesAfter(3, [], [1]) // insert a row at index 1
    const changes = new Map([['Sheet2', change({ rowSources })]])
    expect(rewrite('Sheet2!A2', changes)).toBe('Sheet2!A3')
  })

  it('shifts a sheet-qualified rectangular range through a row delete', () => {
    const rowSources = sourcesAfter(4, [1], []) // delete row index 1
    const changes = new Map([['Sheet2', change({ rowSources })]])
    expect(rewrite('Sheet2!A1:A3', changes)).toBe('Sheet2!A1:A2')
  })

  it('replaces the reference with #REF! when the range was fully consumed by a delete', () => {
    const rowSources = sourcesAfter(3, [1], [])
    const changes = new Map([['Sheet2', change({ rowSources })]])
    expect(rewrite('Sheet2!A2', changes)).toBe('#REF!')
  })

  it('preserves each corner\'s own absolute ($) markers independently through a shift', () => {
    const rowSources = sourcesAfter(4, [], [0]) // insert a row before row 0
    const changes = new Map([['Sheet2', change({ rowSources })]])
    expect(rewrite('Sheet2!$A$1:B$2', changes)).toBe('Sheet2!$A$2:B$3')
  })

  it('re-anchors a whole-column range through a column insert', () => {
    const colSources = sourcesAfter(4, [], [0]) // insert a column before A
    const changes = new Map([['Sheet2', change({ colSources })]])
    expect(rewrite('Sheet2!A:C', changes)).toBe('Sheet2!B:D')
  })

  it('re-anchors a whole-row range through a row delete', () => {
    const rowSources = sourcesAfter(5, [0], []) // delete row 0
    const changes = new Map([['Sheet2', change({ rowSources })]])
    expect(rewrite('Sheet2!$2:$3', changes)).toBe('Sheet2!$1:$2')
  })

  it('shifts an unqualified reference using the caller-supplied own-sheet sources', () => {
    const rowSources = sourcesAfter(3, [], [0]) // insert a row before row 0, on THIS sheet
    expect(rewrite('A1+B2', new Map(), { ownRowSources: rowSources })).toBe('A2+B3')
  })

  it('does not touch an unqualified reference when no own-sheet sources are supplied', () => {
    expect(rewrite('A1+B2', new Map())).toBe('A1+B2')
  })
})

describe('rewriteFormulaReferences — remapCoordinates: false (a formula typed in Atlas)', () => {
  it('still fixes a stale sheet name (rename)', () => {
    const changes = new Map([['Sheet2', change({ newName: 'Data' })]])
    expect(rewrite('Sheet2!A1', changes, { remapCoordinates: false })).toBe('Data!A1')
  })

  it('still turns a reference to a deleted sheet into #REF!', () => {
    const changes = new Map([['Sheet2', change({ deleted: true })]])
    expect(rewrite('Sheet2!A1', changes, { remapCoordinates: false })).toBe('#REF!')
  })

  it('does NOT re-anchor coordinates even when the target sheet has recorded row/column shifts', () => {
    const rowSources = sourcesAfter(3, [], [1])
    const changes = new Map([['Sheet2', change({ newName: 'Data', rowSources })]])
    // The rename still applies; "A2" is left exactly as typed (already current-layout coordinates).
    expect(rewrite('Sheet2!A2', changes, { remapCoordinates: false })).toBe('Data!A2')
  })

  it('never re-anchors an unqualified reference, even with own-sheet sources supplied', () => {
    const rowSources = sourcesAfter(3, [], [0])
    expect(rewrite('A1+B2', new Map(), { remapCoordinates: false, ownRowSources: rowSources })).toBe('A1+B2')
  })
})

describe('rewriteFormulaReferences — leaves everything else alone', () => {
  const changes = new Map([
    ['Sheet1', change({ newName: 'Renamed' })],
    ['Sheet2', change({ deleted: true })],
  ])

  it('leaves a string literal untouched, including one shaped like a reference', () => {
    expect(rewrite('"Sheet1!A1"&Sheet1!B2', changes)).toBe('"Sheet1!A1"&Renamed!B2')
  })

  it('leaves a doubled quote inside a string literal untouched', () => {
    expect(rewrite('"Say ""hi""; see Sheet1!A1"&Sheet1!A1', changes)).toBe('"Say ""hi""; see Sheet1!A1"&Renamed!A1')
  })

  it('leaves a structured table reference untouched', () => {
    expect(rewrite('Table1[Column]+Sheet1!A1', changes)).toBe('Table1[Column]+Renamed!A1')
  })

  it('leaves a nested structured reference untouched', () => {
    expect(rewrite('SUM(Table1[[#Headers],[Column1]])', changes)).toBe('SUM(Table1[[#Headers],[Column1]])')
  })

  it('leaves a bare in-table calculated-column reference untouched', () => {
    expect(rewrite('[@Column1]*[@Column2]', changes)).toBe('[@Column1]*[@Column2]')
  })

  it('leaves a bare defined name untouched', () => {
    expect(rewrite('MyDefinedRange*2', changes)).toBe('MyDefinedRange*2')
  })

  it('leaves a function name before "(" untouched', () => {
    expect(rewrite('SUM(A1:A10)', changes)).toBe('SUM(A1:A10)')
  })

  it('leaves an unquoted external-workbook reference untouched even when the sheet name matches a local one', () => {
    expect(rewrite('[1]Sheet1!A1', changes)).toBe('[1]Sheet1!A1')
  })

  it('leaves a quoted external-workbook reference untouched', () => {
    expect(rewrite("'[1]Sheet1'!A1", changes)).toBe("'[1]Sheet1'!A1")
  })

  it('returns the exact same formula text when nothing needs rewriting', () => {
    const formula = 'IF(A1>0,"yes","no")'
    expect(rewrite(formula, changes)).toBe(formula)
  })
})

describe('rewriteFormulaReferences — realistic composite formulas', () => {
  it('rewrites a multi-area union of sheet-qualified references, preserving the comma', () => {
    const changes = new Map([['Sheet1', change({ newName: 'Plan' })]])
    expect(rewrite('Sheet1!$A$1,Sheet1!$C$3', changes)).toBe('Plan!$A$1,Plan!$C$3')
  })

  it('rewrites a reference nested inside a function call', () => {
    const rowSources = sourcesAfter(10, [], [0])
    const changes = new Map([['Sheet1', change({ newName: 'Plan', rowSources })]])
    expect(rewrite('SUM(Sheet1!A1:A5)', changes)).toBe('SUM(Plan!A2:A6)')
  })

  it('rewrites multiple sheet references across nested function calls', () => {
    const changes = new Map([['Data', change({ newName: 'Facts' })]])
    expect(rewrite('OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)', changes)).toBe('OFFSET(Facts!$A$1,0,0,COUNTA(Facts!$A:$A),1)')
  })
})
