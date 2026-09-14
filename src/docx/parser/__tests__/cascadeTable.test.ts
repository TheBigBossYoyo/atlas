/**
 * Tests for src/docx/parser/cascadeTable.ts (D7 / DXP-06, DXL-08, DXS-05)
 */

import { describe, expect, it } from 'vitest'

import { eighthPoint, hexColor } from '../../model'
import type { Style } from '../../model'
import { resolveTableCellStyle, resolveTableStyle } from '../cascadeTable'
import type { TableCellStyleContext } from '../cascadeTable'
import { DocxParseError } from '../unzip'

function makeStyles(styles: ReadonlyArray<Style>): ReadonlyMap<string, Style> {
  const map = new Map<string, Style>()
  for (const style of styles) {
    map.set(style.id, style)
  }
  return map
}

// Mirrors scripts/generate-docx-corpus.mjs's `table-styled-banded` fixture:
// a header row (firstRow), then alternating band1Horz/band2Horz data rows,
// banding-only-horizontal (noVBand true), 4 rows x 3 columns.
const BANDED_STYLES = makeStyles([
  {
    id: 'AtlasBandedTable',
    type: 'table',
    basedOn: 'TableNormal',
    table: { rowBandSize: 1, colBandSize: 1 },
    conditionalFormats: new Map([
      [
        'firstRow',
        {
          run: { bold: true, color: hexColor('FFFFFF') },
          cell: { shd: { pattern: 'clear', color: 'auto', fill: hexColor('4472C4') } },
        },
      ],
      [
        'band1Horz',
        { cell: { shd: { pattern: 'clear', color: 'auto', fill: hexColor('D9E2F3') } } },
      ],
    ]),
  },
  {
    id: 'TableNormal',
    type: 'table',
    table: { borders: { top: { style: 'single', size: eighthPoint(4) } } },
  },
])

const BANDED_LOOK = { firstRow: true, lastRow: false, firstColumn: false, lastColumn: false, noHBand: false, noVBand: true }

function contextFor(rowIndex: number, columnStart: number = 0): TableCellStyleContext {
  return { rowIndex, rowCount: 4, columnStart, gridSpan: 1, columnCount: 3 }
}

describe('resolveTableStyle', () => {
  it('returns undefined when no styles map is supplied', () => {
    expect(resolveTableStyle('AtlasBandedTable', undefined)).toBeUndefined()
  })

  it('returns undefined for an unresolvable tblStyle id', () => {
    expect(resolveTableStyle('DoesNotExist', BANDED_STYLES)).toBeUndefined()
  })

  it('merges the basedOn chain, most-derived style winning', () => {
    const resolved = resolveTableStyle('AtlasBandedTable', BANDED_STYLES)
    expect(resolved?.rowBandSize).toBe(1)
    expect(resolved?.colBandSize).toBe(1)
    // Inherited from TableNormal (the basedOn ancestor) since AtlasBandedTable
    // doesn't set its own borders.
    expect(resolved?.borders?.top?.style).toBe('single')
  })
})

describe('resolveTableCellStyle', () => {
  it('returns {} when no styles map is supplied', () => {
    expect(resolveTableCellStyle('AtlasBandedTable', undefined, BANDED_LOOK, contextFor(0))).toEqual({})
  })

  it('returns {} when tblStyle is unresolvable', () => {
    expect(resolveTableCellStyle(undefined, BANDED_STYLES, BANDED_LOOK, contextFor(0))).toEqual({})
  })

  it('applies the firstRow conditional format to the header row when tblLook.firstRow is set', () => {
    const resolved = resolveTableCellStyle('AtlasBandedTable', BANDED_STYLES, BANDED_LOOK, contextFor(0))
    expect(resolved.cell?.shd?.fill).toBe('4472C4')
    expect(resolved.run?.bold).toBe(true)
    expect(resolved.run?.color).toBe('FFFFFF')
  })

  it('applies band1Horz to odd stripe rows and leaves band2Horz rows unformatted (style defines no band2Horz block)', () => {
    // Row 1 (first data row, index 1): stripeIndex 1 -> band2Horz (no block defined) -> no shading.
    const row1 = resolveTableCellStyle('AtlasBandedTable', BANDED_STYLES, BANDED_LOOK, contextFor(1))
    expect(row1.cell?.shd).toBeUndefined()

    // Row 2: stripeIndex 2 -> band1Horz -> shaded.
    const row2 = resolveTableCellStyle('AtlasBandedTable', BANDED_STYLES, BANDED_LOOK, contextFor(2))
    expect(row2.cell?.shd?.fill).toBe('D9E2F3')

    // Row 3: stripeIndex 3 -> band2Horz -> no shading.
    const row3 = resolveTableCellStyle('AtlasBandedTable', BANDED_STYLES, BANDED_LOOK, contextFor(3))
    expect(row3.cell?.shd).toBeUndefined()
  })

  it('does not apply row banding when noHBand is true', () => {
    const noBanding = { ...BANDED_LOOK, noHBand: true }
    const row2 = resolveTableCellStyle('AtlasBandedTable', BANDED_STYLES, noBanding, contextFor(2))
    expect(row2.cell?.shd).toBeUndefined()
  })

  it('lets a corner-cell format (nwCell) override firstRow/firstCol when all three are defined', () => {
    const styles = makeStyles([
      {
        id: 'Cornered',
        type: 'table',
        conditionalFormats: new Map([
          ['firstRow', { cell: { shd: { fill: hexColor('111111') } } }],
          ['firstCol', { cell: { shd: { fill: hexColor('222222') } } }],
          ['nwCell', { cell: { shd: { fill: hexColor('333333') } } }],
        ]),
      },
    ])
    const look = { firstRow: true, firstColumn: true, noHBand: true, noVBand: true }

    const resolved = resolveTableCellStyle('Cornered', styles, look, contextFor(0, 0))
    expect(resolved.cell?.shd?.fill).toBe('333333')
  })

  it('throws DocxParseError on a circular basedOn chain instead of recursing forever', () => {
    const circular = makeStyles([
      { id: 'A', type: 'table', basedOn: 'B' },
      { id: 'B', type: 'table', basedOn: 'A' },
    ])

    expect(() => resolveTableStyle('A', circular)).toThrow(DocxParseError)
  })
})
