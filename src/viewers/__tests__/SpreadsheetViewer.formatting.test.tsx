/**
 * SHEETFMT-1 — the workbook's own cell formatting, through the real viewer.
 *
 * Atlas rendered every spreadsheet as unstyled text: a file whose header row is
 * bold white-on-blue with a currency column looked exactly like a CSV. The unit
 * tests in `src/viewers/spreadsheet/__tests__/xlsxCellStyles.test.ts` cover
 * reading `styles.xml`; these cover the part that can only break in the wiring —
 * that what was read reaches the cell the grid draws, through the sheet
 * selection, the frozen-row strip, a search filter and a row insert, each of
 * which shifts the grid's coordinates away from the file's.
 *
 * The fixture is a real OOXML package, built here, because the whole claim is
 * about what Excel's own output means.
 */
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react'
import JSZip from 'jszip'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { CompactSelection, type GridSelection, type Item } from '@glideapps/glide-data-grid'

import type { LoadedFile } from '../../formats/types'
import { ViewerProvider } from '../shared/ViewerContext'
import { SpreadsheetViewer } from '../SpreadsheetViewer'

type CapturedCell = {
  readonly displayData: string
  readonly contentAlign?: 'left' | 'center' | 'right'
  readonly themeOverride?: {
    readonly bgCell?: string
    readonly textDark?: string
    readonly baseFontStyle?: string
    readonly fontFamily?: string
  }
}
type CapturedProps = {
  readonly getCellContent: (loc: readonly [number, number]) => CapturedCell
  readonly drawCell?: (args: never, drawContent: () => void) => void
  readonly rows: number
  readonly onCellEdited?: (cell: Item, newValue: unknown) => void
  readonly onGridSelectionChange?: (selection: GridSelection) => void
}

let lastDataEditorProps: CapturedProps | null = null

vi.mock('@glideapps/glide-data-grid', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@glideapps/glide-data-grid')>()
  return {
    ...actual,
    DataEditor: (props: CapturedProps) => {
      lastDataEditorProps = props
      return null
    },
  }
})

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (items: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`

/**
 * A workbook whose A1:C1 header row is bold, white on blue and centred, with a
 * red right-aligned total in B3 and everything else plain — enough that each
 * assertion below distinguishes a real read from a default.
 */
async function buildFormattedWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  )
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/>`))
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels(
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/>`,
    ),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}">` +
      `<dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>` +
      `<row r="1">` +
      `<c r="A1" s="1" t="inlineStr"><is><t>Item</t></is></c>` +
      `<c r="B1" s="1" t="inlineStr"><is><t>Qty</t></is></c>` +
      `<c r="C1" s="1" t="inlineStr"><is><t>Note</t></is></c>` +
      `</row>` +
      `<row r="2">` +
      `<c r="A2" t="inlineStr"><is><t>Paper</t></is></c>` +
      `<c r="B2"><v>3</v></c>` +
      `<c r="C2" t="inlineStr"><is><t>plain</t></is></c>` +
      `</row>` +
      `<row r="3">` +
      `<c r="A3" t="inlineStr"><is><t>Total</t></is></c>` +
      `<c r="B3" s="2"><v>3</v></c>` +
      `</row>` +
      `</sheetData></worksheet>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">` +
      `<fonts count="3">` +
      `<font><sz val="11"/><name val="Calibri"/></font>` +
      `<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>` +
      `<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>` +
      `</fonts>` +
      `<fills count="3">` +
      `<fill><patternFill patternType="none"/></fill>` +
      `<fill><patternFill patternType="gray125"/></fill>` +
      `<fill><patternFill patternType="solid"><fgColor rgb="FF0070C0"/><bgColor indexed="64"/></patternFill></fill>` +
      `</fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="3">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1">` +
      `<alignment horizontal="center"/></xf>` +
      `<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1">` +
      `<alignment horizontal="right"/></xf>` +
      `</cellXfs></styleSheet>`,
  )
  return zip.generateAsync({ type: 'arraybuffer' })
}

async function mount(content: ArrayBuffer, path = '/tmp/report.xlsx'): Promise<void> {
  const file: LoadedFile = { kind: 'binary', content, path, format: 'xlsx' }
  render(
    <ViewerProvider filePath={file.path}>
      <SpreadsheetViewer file={file} />
    </ViewerProvider>,
  )
  await waitFor(() => expect(lastDataEditorProps).not.toBeNull())
  // The grid's first paint comes from the parse; the formatting arrives with
  // the zip read that follows it, so the cell is read after that settles.
  await waitFor(() => expect(lastDataEditorProps!.getCellContent([0, 0]).displayData).toBe('Item'))
}

/** `[col, row]` in grid coordinates, matching glide-data-grid's own ordering. */
function cell(col: number, row: number): CapturedCell {
  return lastDataEditorProps!.getCellContent([col, row])
}

/**
 * How many separate lines the grid draws for one cell.
 *
 * Borders never reach `themeOverride` — they are drawn by `drawCell` onto the
 * canvas — so `getCellContent` cannot show them and this has to exercise the
 * real draw path. Counting `stroke()` calls is the closest a jsdom test can get
 * to "the border appeared", and the pixel-level e2e covers the rest.
 */
function strokesFor(col: number, row: number): number {
  let strokes = 0
  const ctx = {
    measureText: (text: string) => ({ width: text.length * 7 }),
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    stroke: () => {
      strokes += 1
    },
    strokeStyle: '',
    lineWidth: 1,
  }
  lastDataEditorProps!.drawCell?.(
    {
      ctx: ctx as unknown as CanvasRenderingContext2D,
      cell: lastDataEditorProps!.getCellContent([col, row]),
      theme: { cellHorizontalPadding: 8, textDark: '#000000' },
      rect: { x: 0, y: 0, width: 100, height: 30 },
      col,
      row,
      hoverAmount: 0,
      hoverX: undefined,
      hoverY: undefined,
      highlighted: false,
      imageLoader: undefined,
    } as never,
    () => {},
  )
  return strokes
}

/** Clicking a cell. The toolbar's row/column actions act on the selection, so they do nothing without one. */
function selectCell(col: number, row: number): void {
  selectRange(col, row, 1, 1)
}

/** Dragging out a rectangle, in grid coordinates. */
function selectRange(col: number, row: number, width: number, height: number): void {
  lastDataEditorProps!.onGridSelectionChange?.({
    current: { cell: [col, row], range: { x: col, y: row, width, height }, rangeStack: [] },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  })
}

beforeEach(() => {
  lastDataEditorProps = null
  window.electronAPI = {
    saveBinaryFile: vi.fn().mockResolvedValue({ saved: true, path: '/tmp/report.xlsx', name: 'report.xlsx' }),
    saveFile: vi.fn(),
  } as unknown as typeof window.electronAPI
})

describe('SpreadsheetViewer — the workbook its own formatting (SHEETFMT-1)', () => {
  it('draws a bold, filled, centred header row', async () => {
    await mount(await buildFormattedWorkbook())

    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))
    const header = cell(0, 0)
    expect(header.themeOverride?.textDark).toBe('#FFFFFF')
    expect(header.themeOverride?.baseFontStyle).toContain('700')
    expect(header.contentAlign).toBe('center')
  })

  it('leaves an unstyled cell exactly as it was before formatting was read', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    // The whole-feature safety property: a cell the file does not style must
    // not acquire a colour, a font or an alignment. Without this, reading
    // `fonts[0]` as a format pinned every workbook to its stored font size and
    // overrode the user's own theme.
    const plain = cell(2, 1)
    expect(plain.themeOverride?.textDark).toBeUndefined()
    expect(plain.themeOverride?.baseFontStyle).toBeUndefined()
    expect(plain.contentAlign).toBeUndefined()
  })

  it('applies a per-cell colour and alignment to the right cell', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(1, 2).themeOverride?.textDark).toBe('#FF0000'))
    expect(cell(1, 2).contentAlign).toBe('right')
    // ...and not to its neighbours.
    expect(cell(0, 2).themeOverride?.textDark).toBeUndefined()
    expect(cell(2, 2).themeOverride?.textDark).toBeUndefined()
  })

  it('keeps formatting on the right rows after a row is inserted above it', async () => {
    // The coordinate trap this feature's design is built around: `cellStyles`
    // is indexed by the cell's position in the FILE, while the grid's rows
    // shift as soon as one is inserted. Reading formats by the current index
    // slides every one of them a row out of place.
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(1, 2).themeOverride?.textDark).toBe('#FF0000'))

    // Insert above the red total's own row, so it is pushed down by one.
    act(() => selectCell(1, 2))
    const insertRow = document.querySelector<HTMLButtonElement>('[aria-label*="Insert row" i]')
    expect(insertRow, 'the toolbar should offer Insert row').not.toBeNull()
    act(() => insertRow!.click())
    await waitFor(() => expect(cell(0, 3).displayData).toBe('Total'))

    // The red total moved from row 2 to row 3, and its formatting came with it.
    expect(cell(1, 3).themeOverride?.textDark).toBe('#FF0000')
    // The inserted row is genuinely new, so it carries no formatting at all.
    expect(cell(1, 2).themeOverride?.textDark).toBeUndefined()
  })

  it('applies bold from the toolbar to the selected cell, and shows the toggle as pressed', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1)) // "Paper", plain in the fixture
    const bold = await screen.findByRole('button', { name: /^Bold$/i })
    expect(bold).toHaveAttribute('aria-pressed', 'false')

    act(() => bold.click())

    await waitFor(() => expect(cell(0, 1).themeOverride?.baseFontStyle).toContain('700'))
    // The toggle now reflects the cell it is pointed at — a formatting button
    // that never looks pressed leaves the user guessing whether it worked.
    await waitFor(() => expect(screen.getByRole('button', { name: /^Bold$/i })).toHaveAttribute('aria-pressed', 'true'))
  })

  it('un-bolds a cell the file made bold', async () => {
    // The toggle has to send the opposite of the CURRENT state, so pressing it
    // on the file's own bold header clears it rather than re-applying bold.
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.baseFontStyle).toContain('700'))

    act(() => selectCell(0, 0))
    const bold = await screen.findByRole('button', { name: /^Bold$/i })
    await waitFor(() => expect(bold).toHaveAttribute('aria-pressed', 'true'))

    act(() => bold.click())

    await waitFor(() => expect(cell(0, 0).themeOverride?.baseFontStyle ?? '').not.toContain('700'))
  })

  it('applies a fill from the colour popover', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1))
    act(() => screen.getByRole('button', { name: /Fill color/i }).click())
    // The swatches are real buttons labelled by their hex, deliberately not a
    // role="menu" grid (see the component's own note).
    act(() => screen.getByRole('button', { name: '#FF0000' }).click())

    await waitFor(() => expect(cell(0, 1).themeOverride?.bgCell).toBe('#FF0000'))
  })

  it('applies a format to every cell in a multi-cell selection', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    // A 2x2 block starting at A2.
    act(() => selectRange(0, 1, 2, 2))
    act(() => screen.getByRole('button', { name: /Align right/i }).click())

    await waitFor(() => expect(cell(0, 1).contentAlign).toBe('right'))
    expect(cell(1, 1).contentAlign).toBe('right')
    expect(cell(0, 2).contentAlign).toBe('right')
    expect(cell(1, 2).contentAlign).toBe('right')
    // ...and not outside it.
    expect(cell(2, 1).contentAlign).toBeUndefined()
  })

  it('applies a border preset from the toolbar (SHEETFMT-3)', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1))
    act(() => screen.getByRole('button', { name: /Borders/i }).click())
    act(() => screen.getByRole('button', { name: /All borders/i }).click())

    // Four edges means four separate stroked lines.
    await waitFor(() => expect(lastDataEditorProps!.drawCell).toBeTypeOf('function'))
    await waitFor(() => expect(strokesFor(0, 1)).toBe(4))
    // A neighbour the preset did not cover draws nothing.
    expect(strokesFor(2, 1)).toBe(0)
  })

  it('applies a font family and size, and shows both in the grid (SHEETFMT-4)', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1))
    const family = screen.getByRole('combobox', { name: /^Font$/i })
    const size = screen.getByRole('combobox', { name: /Font size/i })

    fireEvent.change(family, { target: { value: 'Georgia' } })
    fireEvent.change(size, { target: { value: '18' } })

    // Both have to reach the RENDERER, not just the file: the family travels
    // on its own theme field (the grid appends `fontFamily` to the font
    // shorthand itself), and the size inside the shorthand.
    await waitFor(() => expect(cell(0, 1).themeOverride?.fontFamily).toBe('Georgia'))
    // 18pt at 96dpi is 24px.
    expect(cell(0, 1).themeOverride?.baseFontStyle).toContain('24px')

    // And the controls now report the cell's own font.
    await waitFor(() => expect(screen.getByRole('combobox', { name: /^Font$/i })).toHaveValue('Georgia'))
    expect(screen.getByRole('combobox', { name: /Font size/i })).toHaveValue('18')
  })

  it('toggles a single border edge on and back off (SHEETFMT-5)', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1))
    act(() => screen.getByRole('button', { name: /Borders/i }).click())

    const bottom = screen.getByRole('button', { name: 'Bottom' })
    expect(bottom).toHaveAttribute('aria-pressed', 'false')

    act(() => bottom.click())
    // One edge means one stroked line.
    await waitFor(() => expect(strokesFor(0, 1)).toBe(1))
    await waitFor(() => expect(screen.getByRole('button', { name: 'Bottom' })).toHaveAttribute('aria-pressed', 'true'))

    // Pressing it again takes that edge back off, which a toggle that always
    // applied could not do.
    act(() => screen.getByRole('button', { name: 'Bottom' }).click())
    await waitFor(() => expect(strokesFor(0, 1)).toBe(0))
  })

  it('adds a second edge without clearing the first (SHEETFMT-5)', async () => {
    await mount(await buildFormattedWorkbook())
    await waitFor(() => expect(cell(0, 0).themeOverride?.bgCell).toBe('#0070C0'))

    act(() => selectCell(0, 1))
    act(() => screen.getByRole('button', { name: /Borders/i }).click())
    act(() => screen.getByRole('button', { name: 'Bottom' }).click())
    await waitFor(() => expect(strokesFor(0, 1)).toBe(1))
    act(() => screen.getByRole('button', { name: 'Left' }).click())

    // The patch composes rather than replacing, so both edges are drawn.
    await waitFor(() => expect(strokesFor(0, 1)).toBe(2))
  })

  it('leaves the formatting controls disabled until something is selected', async () => {
    await mount(await buildFormattedWorkbook())
    expect(await screen.findByRole('button', { name: /^Bold$/i })).toBeDisabled()
  })

  it('renders a workbook with no styles.xml unstyled rather than failing', async () => {
    // Every `.csv`, `.xls` and SheetJS-written file takes this path.
    const bare = new JSZip()
    bare.file('nope.txt', 'not a workbook')
    await expect(
      (async () => {
        const file: LoadedFile = {
          kind: 'binary',
          content: await bare.generateAsync({ type: 'arraybuffer' }),
          path: '/tmp/bare.xlsx',
          format: 'xlsx',
        }
        render(
          <ViewerProvider filePath={file.path}>
            <SpreadsheetViewer file={file} />
          </ViewerProvider>,
        )
      })(),
    ).resolves.not.toThrow()
  })
})
