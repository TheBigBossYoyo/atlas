/**
 * SHEETFMT-3 — the cell decorations the grid's own renderer cannot draw.
 *
 * `themeOverride` covers background, text colour and the font shorthand, which
 * is why bold, italic, size, colour, fill and alignment worked through it.
 * Underline, strike-through and cell borders are not font properties: they are
 * LINES, and the grid has no concept of them. Atlas read and saved all three
 * and drew none of them, so a cell could be underlined in the file while the
 * grid showed it plain.
 *
 * The drawing happens on a canvas, so these tests hand the callback a recording
 * context and assert on the path it builds. That makes the geometry decisions —
 * which edge, how thick, where the underline sits under right-aligned text —
 * checkable without a real canvas, which jsdom does not provide.
 */
import { renderHook } from '@testing-library/react'
import { describe, expect, it } from 'vitest'

import { useSpreadsheetGrid } from '../useSpreadsheetGrid'
import { DEFAULT_CELL_FORMAT, NO_BORDER, type ResolvedCellFormat } from '../../spreadsheet/xlsxCellStyles'

type Op =
  | { kind: 'move'; x: number; y: number }
  | { kind: 'line'; x: number; y: number }
  | { kind: 'stroke'; style: string; width: number }

/** A canvas 2D context that records the path operations instead of painting. */
function recordingContext() {
  const ops: Op[] = []
  const ctx = {
    ops,
    strokeStyle: '#000000',
    lineWidth: 1,
    measureText: (text: string) => ({ width: text.length * 7 }),
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    moveTo: (x: number, y: number) => ops.push({ kind: 'move', x, y }),
    lineTo: (x: number, y: number) => ops.push({ kind: 'line', x, y }),
    stroke: () => ops.push({ kind: 'stroke', style: String(ctx.strokeStyle), width: ctx.lineWidth }),
  }
  return ctx
}

const RECT = { x: 100, y: 200, width: 120, height: 32 }

function format(overrides: Partial<ResolvedCellFormat>): ResolvedCellFormat {
  return { ...DEFAULT_CELL_FORMAT, ...overrides }
}

/**
 * Runs the hook's `drawCell` for one cell and returns what it drew.
 *
 * `drawContent` is the grid's own rendering, which the callback must still
 * invoke — a decorator that forgot to would blank every cell.
 */
function draw(
  cellFormat: ResolvedCellFormat,
  options: { readonly text?: string; readonly align?: 'left' | 'center' | 'right' } = {},
) {
  const { result } = renderHook(() =>
    useSpreadsheetGrid({
      rows: [[options.text ?? 'hello']],
      colCount: 1,
      cellFormat: () => cellFormat,
    }),
  )

  const drawCell = result.current.drawCell
  expect(drawCell, 'drawCell should be provided when cell formatting is').toBeTypeOf('function')

  const ctx = recordingContext()
  let contentDrawn = 0
  drawCell!(
    {
      ctx: ctx as unknown as CanvasRenderingContext2D,
      cell: {
        kind: 'text',
        data: options.text ?? 'hello',
        displayData: options.text ?? 'hello',
        allowOverlay: false,
        ...(options.align ? { contentAlign: options.align } : {}),
      },
      theme: { cellHorizontalPadding: 8, textDark: '#111111' },
      rect: RECT,
      col: 0,
      row: 0,
      hoverAmount: 0,
      hoverX: undefined,
      hoverY: undefined,
      highlighted: false,
      imageLoader: undefined,
    } as never,
    () => {
      contentDrawn += 1
    },
  )

  return { ops: ctx.ops, contentDrawn }
}

describe('cell decorations (SHEETFMT-3)', () => {
  it('is not provided at all when the grid has no formatting', () => {
    // A `.csv`, or a workbook whose styles could not be read, must keep the
    // grid's original rendering path untouched.
    const { result } = renderHook(() => useSpreadsheetGrid({ rows: [['a']], colCount: 1 }))
    expect(result.current.drawCell).toBeUndefined()
  })

  it('always draws the cell content, then decorates', () => {
    const { contentDrawn, ops } = draw(format({ underline: true }))
    expect(contentDrawn, 'the grid\'s own rendering must still happen').toBe(1)
    expect(ops.length).toBeGreaterThan(0)
  })

  it('draws nothing extra for an unformatted cell', () => {
    const { ops, contentDrawn } = draw(DEFAULT_CELL_FORMAT)
    expect(contentDrawn).toBe(1)
    expect(ops, 'an unstyled cell must be drawn exactly as before').toEqual([])
  })

  it('underlines across the width of the text, not the cell', () => {
    // "hello" at the mock's 7px/char is 35px wide, from the left padding.
    const { ops } = draw(format({ underline: true }))
    const move = ops.find((op) => op.kind === 'move')!
    const line = ops.find((op) => op.kind === 'line')!
    expect(move).toMatchObject({ x: RECT.x + 8 })
    expect(line.kind === 'line' ? line.x - move.x : 0).toBe(35)
  })

  it('puts the underline under right-aligned text', () => {
    // The whole reason the text is measured: a line drawn from the left padding
    // would sit nowhere near right-aligned glyphs.
    const { ops } = draw(format({ underline: true }), { align: 'right' })
    const move = ops.find((op) => op.kind === 'move')!
    expect(move.x).toBe(RECT.x + RECT.width - 8 - 35)
  })

  it('centres the underline under centred text', () => {
    const { ops } = draw(format({ underline: true }), { align: 'center' })
    const move = ops.find((op) => op.kind === 'move')!
    const available = RECT.width - 16
    expect(move.x).toBe(RECT.x + 8 + (available - 35) / 2)
  })

  it('draws a strike-through through the middle, above the underline', () => {
    const { ops } = draw(format({ underline: true, strike: true }))
    const ys = ops.filter((op) => op.kind === 'move').map((op) => op.y)
    expect(ys).toHaveLength(2)
    // Strike sits at the cell's middle; the underline below it.
    expect(Math.min(...ys)).toBeLessThan(Math.max(...ys))
  })

  it('uses the cell\'s own text colour for the line, not the theme default', () => {
    const { ops } = draw(format({ underline: true, color: '#FF0000' }))
    expect(ops.find((op) => op.kind === 'stroke')).toMatchObject({ style: '#FF0000' })
  })

  it('falls back to the theme text colour when the cell sets none', () => {
    const { ops } = draw(format({ underline: true }))
    expect(ops.find((op) => op.kind === 'stroke')).toMatchObject({ style: '#111111' })
  })

  it('draws nothing for text with no width', () => {
    const { ops } = draw(format({ underline: true }), { text: '' })
    expect(ops).toEqual([])
  })

  it('draws only the border edges the cell actually has', () => {
    const { ops } = draw(
      format({ border: { ...NO_BORDER, bottom: { weight: 'thin', color: '#0000FF' } } }),
    )
    const strokes = ops.filter((op) => op.kind === 'stroke')
    expect(strokes).toHaveLength(1)
    expect(strokes[0]).toMatchObject({ style: '#0000FF', width: 1 })

    // On the bottom edge, as a horizontal line.
    const move = ops.find((op) => op.kind === 'move')!
    const line = ops.find((op) => op.kind === 'line')!
    expect(move.y).toBeCloseTo(RECT.y + RECT.height - 0.5)
    expect(line.y).toBe(move.y)
    expect(move.x).toBe(RECT.x)
  })

  it('draws all four edges of a boxed cell, each at its own weight', () => {
    const { ops } = draw(
      format({
        border: {
          top: { weight: 'thin', color: '#111111' },
          right: { weight: 'medium', color: '#111111' },
          bottom: { weight: 'thick', color: '#111111' },
          left: { weight: 'thin', color: '#111111' },
        },
      }),
    )
    const widths = ops.filter((op) => op.kind === 'stroke').map((op) => (op.kind === 'stroke' ? op.width : 0))
    expect(widths).toEqual([1, 2, 3, 1])
  })

  it('uses a mid grey for an automatic-coloured border, readable in either theme', () => {
    // OOXML's `indexed="64"` automatic colour resolves to `undefined`. Black
    // would vanish in Atlas's dark themes.
    const { ops } = draw(format({ border: { ...NO_BORDER, top: { weight: 'thin', color: undefined } } }))
    expect(ops.find((op) => op.kind === 'stroke')).toMatchObject({ style: '#808080' })
  })

  it('draws both the text decoration and the borders when a cell has both', () => {
    const { ops } = draw(
      format({ underline: true, border: { ...NO_BORDER, left: { weight: 'thin', color: '#111111' } } }),
    )
    expect(ops.filter((op) => op.kind === 'stroke')).toHaveLength(2)
  })
})
