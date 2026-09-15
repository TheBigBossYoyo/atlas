import { describe, expect, it } from 'vitest'

import type {
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  NumberingDef,
  Paragraph,
  ParagraphChild,
  Run,
  Section,
  Style,
  Table,
  TableRow,
} from '../../model'
import { twip } from '../../model'
import type { Position, Range } from '../commandTypes'
import { toolbarToCommand } from '../toolbarAdapter'

function pos(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

function collapsed(p: Position): Range {
  return { anchor: p, focus: p }
}

function createRun(text: string): Run {
  return Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) as Run
}

function createRevisionParagraph(kind: 'ins-revision' | 'del-revision'): Paragraph {
  const revision: ParagraphChild = Object.freeze({
    kind,
    id: 'r1',
    children: Object.freeze([createRun('tracked')]),
  }) as ParagraphChild

  return Object.freeze({ kind: 'paragraph', children: Object.freeze([revision]) }) as Paragraph
}

function createDocument(paragraphs: ReadonlyArray<Paragraph>): Document {
  const section = Object.freeze({
    kind: 'section',
    props: {},
    blocks: Object.freeze([...paragraphs]),
  }) satisfies Section

  return Object.freeze({
    kind: 'document',
    sections: Object.freeze([section]),
    styles: new Map<string, Style>(),
    numbering: new Map<string, NumberingDef>(),
    comments: new Map<string, Comment>(),
    footnotes: new Map<string, Footnote>(),
    endnotes: new Map<string, Endnote>(),
    headers: new Map<string, Header>(),
    footers: new Map<string, Footer>(),
  }) satisfies Document
}

/** DXE-14 — a `rows` x `cols` grid table as the document's only top-level
 * block, matching `commands.ts`'s `tableCommands.test.ts` fixture shape. */
function createTableDocument(rows: number, cols: number): Document {
  const table: Table = Object.freeze({
    kind: 'table',
    tblGrid: Object.freeze(Array.from({ length: cols }, () => twip(1440))),
    rows: Object.freeze(
      Array.from({ length: rows }, () =>
        Object.freeze({
          kind: 'table-row',
          cells: Object.freeze(
            Array.from({ length: cols }, () =>
              Object.freeze({
                kind: 'table-cell',
                blocks: Object.freeze([
                  Object.freeze({ kind: 'paragraph', children: Object.freeze([]) }) as Paragraph,
                ]),
              }),
            ),
          ),
        } satisfies TableRow),
      ),
    ),
  }) as Table

  const section = Object.freeze({ kind: 'section', props: {}, blocks: Object.freeze([table]) }) satisfies Section

  return Object.freeze({
    kind: 'document',
    sections: Object.freeze([section]),
    styles: new Map<string, Style>(),
    numbering: new Map<string, NumberingDef>(),
    comments: new Map<string, Comment>(),
    footnotes: new Map<string, Footnote>(),
    endnotes: new Map<string, Endnote>(),
    headers: new Map<string, Header>(),
    footers: new Map<string, Footer>(),
  }) satisfies Document
}

/** Selection sitting in the first paragraph of the table's `(rowIndex,
 * cellIndex)` cell — table is the document's only top-level block (index 0)
 * of section 0, so the full path is `[0, 0, rowIndex, cellIndex, 0]`. */
function cellPos(rowIndex: number, cellIndex: number): Range {
  return collapsed(pos([0, 0, rowIndex, cellIndex, 0], 0, 0))
}

describe('toolbarToCommand — DXE-14 table structural editing', () => {
  it('returns null for every table command when the selection is outside a table', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('plain')]) }) as Paragraph,
    ])
    const selection = collapsed(pos([0], 0, 0))
    const kinds = [
      'insert-table-row-above',
      'insert-table-row-below',
      'insert-table-column-left',
      'insert-table-column-right',
      'delete-table-row',
      'delete-table-column',
      'delete-table',
      'merge-table-cell-right',
      'split-table-cell',
    ] as const

    for (const kind of kinds) {
      expect(toolbarToCommand({ kind }, selection, document)).toBeNull()
    }
  })

  it('resolves insert-table-row-above/below against the cursor\'s row', () => {
    const document = createTableDocument(2, 2)

    expect(toolbarToCommand({ kind: 'insert-table-row-above' }, cellPos(1, 0), document)).toEqual({
      kind: 'insert-table-row',
      tablePath: [0, 0],
      at: 1,
    })
    expect(toolbarToCommand({ kind: 'insert-table-row-below' }, cellPos(1, 0), document)).toEqual({
      kind: 'insert-table-row',
      tablePath: [0, 0],
      at: 2,
    })
  })

  it('resolves insert-table-column-left/right against the cursor\'s column', () => {
    const document = createTableDocument(2, 3)

    expect(toolbarToCommand({ kind: 'insert-table-column-left' }, cellPos(0, 1), document)).toEqual({
      kind: 'insert-table-column',
      tablePath: [0, 0],
      at: 1,
    })
    expect(toolbarToCommand({ kind: 'insert-table-column-right' }, cellPos(0, 1), document)).toEqual({
      kind: 'insert-table-column',
      tablePath: [0, 0],
      at: 2,
    })
  })

  // ---------------------------------------------------------------------------
  // Regression — `EnclosingTable.cellIndex` is an array index into
  // `row.cells`, which stops matching the table's *grid* column index (what
  // insert/delete-column actually address) as soon as an earlier cell in the
  // row has been merged wider. Resolving straight from `cellIndex` used to
  // pick the wrong column once that happened.
  // ---------------------------------------------------------------------------

  /** A single row spanning 3 grid columns but only 2 cells — `AB` (gridSpan
   * 2, as `merge-table-cells` would produce) followed by plain cell `C`, so
   * `C`'s array index (1) diverges from its grid column (2). */
  function createMergedTableDocument(): Document {
    const emptyCellBlocks = Object.freeze([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([]) }) as Paragraph,
    ])
    const table: Table = Object.freeze({
      kind: 'table',
      tblGrid: Object.freeze([twip(1440), twip(1440), twip(1440)]),
      rows: Object.freeze([
        Object.freeze({
          kind: 'table-row',
          cells: Object.freeze([
            Object.freeze({ kind: 'table-cell', props: { gridSpan: 2 }, blocks: emptyCellBlocks }),
            Object.freeze({ kind: 'table-cell', blocks: emptyCellBlocks }),
          ]),
        } satisfies TableRow),
      ]),
    }) as Table

    const section = Object.freeze({ kind: 'section', props: {}, blocks: Object.freeze([table]) }) satisfies Section

    return Object.freeze({
      kind: 'document',
      sections: Object.freeze([section]),
      styles: new Map<string, Style>(),
      numbering: new Map<string, NumberingDef>(),
      comments: new Map<string, Comment>(),
      footnotes: new Map<string, Footnote>(),
      endnotes: new Map<string, Endnote>(),
      headers: new Map<string, Header>(),
      footers: new Map<string, Footer>(),
    }) satisfies Document
  }

  it('resolves insert-table-column-left/right against the grid column, not the cell array index, once an earlier cell is merged', () => {
    // Row: [AB (gridSpan 2), C] — three grid columns, two cells. The cursor
    // sits in C, whose array index is 1 but whose grid column is 2.
    const document = createMergedTableDocument()

    expect(toolbarToCommand({ kind: 'insert-table-column-left' }, cellPos(0, 1), document)).toEqual({
      kind: 'insert-table-column',
      tablePath: [0, 0],
      at: 2,
    })
    expect(toolbarToCommand({ kind: 'insert-table-column-right' }, cellPos(0, 1), document)).toEqual({
      kind: 'insert-table-column',
      tablePath: [0, 0],
      at: 3,
    })
  })

  it('resolves delete-table-column against the grid column, not the cell array index, once an earlier cell is merged', () => {
    const document = createMergedTableDocument()

    expect(toolbarToCommand({ kind: 'delete-table-column' }, cellPos(0, 1), document)).toEqual({
      kind: 'delete-table-column',
      tablePath: [0, 0],
      columnIndex: 2,
    })
  })

  it('resolves delete-table-row/column against the cursor\'s cell', () => {
    const document = createTableDocument(2, 2)

    expect(toolbarToCommand({ kind: 'delete-table-row' }, cellPos(1, 0), document)).toEqual({
      kind: 'delete-table-row',
      tablePath: [0, 0],
      rowIndex: 1,
    })
    expect(toolbarToCommand({ kind: 'delete-table-column' }, cellPos(0, 1), document)).toEqual({
      kind: 'delete-table-column',
      tablePath: [0, 0],
      columnIndex: 1,
    })
  })

  it('resolves delete-table to the enclosing table regardless of cell', () => {
    const document = createTableDocument(2, 2)

    expect(toolbarToCommand({ kind: 'delete-table' }, cellPos(1, 1), document)).toEqual({
      kind: 'delete-table',
      tablePath: [0, 0],
    })
  })

  it('resolves merge-table-cell-right to the cursor cell and its right neighbor', () => {
    const document = createTableDocument(1, 3)

    expect(toolbarToCommand({ kind: 'merge-table-cell-right' }, cellPos(0, 0), document)).toEqual({
      kind: 'merge-table-cells',
      tablePath: [0, 0],
      rowIndex: 0,
      fromCellIndex: 0,
      toCellIndex: 1,
    })
  })

  it('resolves split-table-cell to the cursor cell', () => {
    const document = createTableDocument(1, 2)

    expect(toolbarToCommand({ kind: 'split-table-cell' }, cellPos(0, 1), document)).toEqual({
      kind: 'split-table-cell',
      tablePath: [0, 0],
      rowIndex: 0,
      cellIndex: 1,
    })
  })

  it('set-table-properties returns null when the selection is outside a table', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('plain')]) }) as Paragraph,
    ])
    const command = toolbarToCommand(
      { kind: 'set-table-properties', widthTwips: 5000, alignment: 'center', bordersOn: true },
      collapsed(pos([0], 0, 0)),
      document,
    )

    expect(command).toBeNull()
  })

  it('set-table-properties builds an apply-table-props command with an explicit width, alignment and borders on', () => {
    const document = createTableDocument(1, 1)
    const command = toolbarToCommand(
      { kind: 'set-table-properties', widthTwips: 5000, alignment: 'center', bordersOn: true },
      cellPos(0, 0),
      document,
    )

    expect(command).toEqual({
      kind: 'apply-table-props',
      tablePath: [0, 0],
      props: {
        tblW: { type: 'dxa', value: 5000 },
        jc: 'center',
        tblBorders: {
          top: { style: 'single', size: 4, color: '#000000' },
          bottom: { style: 'single', size: 4, color: '#000000' },
          left: { style: 'single', size: 4, color: '#000000' },
          right: { style: 'single', size: 4, color: '#000000' },
          insideH: { style: 'single', size: 4, color: '#000000' },
          insideV: { style: 'single', size: 4, color: '#000000' },
        },
      },
    })
  })

  it('set-table-properties omits tblW and turns every border off when widthTwips is null and bordersOn is false', () => {
    const document = createTableDocument(1, 1)
    const command = toolbarToCommand(
      { kind: 'set-table-properties', widthTwips: null, alignment: 'left', bordersOn: false },
      cellPos(0, 0),
      document,
    )

    expect(command).toMatchObject({ kind: 'apply-table-props', tablePath: [0, 0] })
    const props = (command as { props: { tblW?: unknown; jc: string; tblBorders: { top: { style: string } } } }).props
    expect(props.tblW).toBeUndefined()
    expect(props.jc).toBe('start')
    expect(props.tblBorders.top.style).toBe('none')
  })
})

describe('toolbarToCommand', () => {
  // ---------------------------------------------------------------------------
  // D17 — accept/reject wiring
  // ---------------------------------------------------------------------------

  it('accept-change resolves the ins-revision at the selection paragraph', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    const command = toolbarToCommand({ kind: 'accept-change' }, collapsed(pos([0], 0, 0)), document)

    expect(command).toEqual({ kind: 'accept-revision', paragraphPath: [0], childIndex: 0 })
  })

  it('reject-change resolves the del-revision at the selection paragraph', () => {
    const document = createDocument([createRevisionParagraph('del-revision')])
    const command = toolbarToCommand({ kind: 'reject-change' }, collapsed(pos([0], 0, 0)), document)

    expect(command).toEqual({ kind: 'reject-revision', paragraphPath: [0], childIndex: 0 })
  })

  it('accept-change returns null when there is no selection', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'accept-change' }, null, document)).toBeNull()
  })

  it('accept-change returns null when the selected paragraph has no revision', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('plain')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'accept-change' }, collapsed(pos([0], 0, 0)), document)).toBeNull()
  })

  it('reject-change returns null when the paragraph path does not resolve', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'reject-change' }, collapsed(pos([9], 0, 0)), document)).toBeNull()
  })

  it('accept-all-changes builds an accept-all-revisions command regardless of selection', () => {
    const document = createDocument([createRevisionParagraph('ins-revision')])
    expect(toolbarToCommand({ kind: 'accept-all-changes' }, null, document)).toEqual({
      kind: 'accept-all-revisions',
    })
  })

  it('reject-all-changes builds a reject-all-revisions command regardless of selection', () => {
    const document = createDocument([createRevisionParagraph('del-revision')])
    expect(toolbarToCommand({ kind: 'reject-all-changes' }, collapsed(pos([0], 0, 0)), document)).toEqual({
      kind: 'reject-all-revisions',
    })
  })

  // ---------------------------------------------------------------------------
  // Existing wiring stays intact after the signature change
  // ---------------------------------------------------------------------------

  it('toggle-bold still builds an apply-run-format command', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const range: Range = { anchor: pos([0], 0, 0), focus: pos([0], 0, 5) }

    const command = toolbarToCommand({ kind: 'toggle-bold' }, range, document)

    expect(command).toEqual({ kind: 'apply-run-format', range, format: { bold: true } })
  })

  it('insert-table builds a command from the selection focus', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const command = toolbarToCommand(
      { kind: 'insert-table', rows: 2, cols: 3 },
      collapsed(pos([0], 0, 2)),
      document,
    )

    expect(command).toEqual({ kind: 'insert-table', at: pos([0], 0, 2), rows: 2, cols: 3 })
  })

  it('insert-page-break builds an insert-inline command with a page BreakNode at the caret', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    const command = toolbarToCommand(
      { kind: 'insert-page-break' },
      collapsed(pos([0], 0, 2)),
      document,
    )

    expect(command).toEqual({
      kind: 'insert-inline',
      at: pos([0], 0, 2),
      child: { kind: 'break', breakType: 'page' },
    })
  })

  it('insert-page-break returns null when there is no selection', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'insert-page-break' }, null, document)).toBeNull()
  })

  it('insert-hyperlink is still handled by the caller (bundle-aware), not this pure mapper', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('hello')]) }) as Paragraph,
    ])
    expect(toolbarToCommand({ kind: 'insert-hyperlink' }, collapsed(pos([0], 0, 0)), document)).toBeNull()
  })

  // ---------------------------------------------------------------------------
  // D18 — a multi-paragraph selection includes every paragraph in between,
  // not just its two endpoints
  // ---------------------------------------------------------------------------

  it('toggle-bullet-list targets every paragraph a multi-paragraph selection spans', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Two')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Three')]) }) as Paragraph,
    ])
    const range: Range = { anchor: pos([0, 0], 0, 0), focus: pos([0, 2], 0, 2) }

    const command = toolbarToCommand({ kind: 'toggle-bullet-list' }, range, document)

    expect(command).toEqual({
      kind: 'insert-list',
      paragraphPaths: [[0, 0], [0, 1], [0, 2]],
      numId: 1,
      level: 0,
    })
  })

  it('set-alignment targets every paragraph regardless of anchor/focus order', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Two')]) }) as Paragraph,
    ])
    // Focus before anchor (user selected upward).
    const range: Range = { anchor: pos([0, 1], 0, 3), focus: pos([0, 0], 0, 0) }

    const command = toolbarToCommand({ kind: 'set-alignment', align: 'center' }, range, document)

    expect(command).toMatchObject({ kind: 'apply-para-format', paragraphPaths: [[0, 0], [0, 1]] })
  })

  it('USR-04: a selection ending at the start of the next paragraph does not align that paragraph', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Title')]) }) as Paragraph,
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('Body')]) }) as Paragraph,
    ])
    // What a triple-click on the title line used to produce: [1] offset 0.
    const range: Range = { anchor: pos([0], 0, 0), focus: pos([1], 0, 0) }

    const command = toolbarToCommand({ kind: 'set-alignment', align: 'right' }, range, document)

    expect(command).toMatchObject({ kind: 'apply-para-format', paragraphPaths: [[0, 0]] })
  })

  it('USR-03: toolbar underline removes underline from already-underlined text', () => {
    const underlined: Run = Object.freeze({ ...createRun('Hello'), props: Object.freeze({ underline: { style: 'single' } }) }) as Run
    const document = createDocument([Object.freeze({ kind: 'paragraph', children: Object.freeze([underlined]) }) as Paragraph])
    const range: Range = { anchor: pos([0, 0], 0, 0), focus: pos([0, 0], 0, 5) }

    const command = toolbarToCommand({ kind: 'toggle-underline' }, range, document)

    expect(command).toEqual({ kind: 'apply-run-format', range, format: { underline: undefined } })
  })

  // ---------------------------------------------------------------------------
  // Regression — a real Word document almost always already defines its own
  // numId "1" (and often "2"). Hardcoding those for the toolbar's list
  // toggle meant clicking "Bulleted List" on such a document silently reused
  // whatever list style numId 1 already meant there instead of creating an
  // actual bullet definition.
  // ---------------------------------------------------------------------------

  function documentWithNumbering(numbering: ReadonlyMap<string, NumberingDef>): Document {
    const doc = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
    ])
    return { ...doc, numbering }
  }

  function decimalNumberingDef(numId: string): NumberingDef {
    return {
      numId,
      abstractNumId: `real-doc-${numId}`,
      levels: new Map([[0, { level: 0, format: 'decimal' }]]),
    }
  }

  it('toggle-bullet-list does not reuse an existing non-Atlas numId "1" from the source document', () => {
    const document = documentWithNumbering(new Map([['1', decimalNumberingDef('1')]]))
    const command = toolbarToCommand({ kind: 'toggle-bullet-list' }, collapsed(pos([0, 0], 0, 0)), document)

    expect(command).toMatchObject({ kind: 'insert-list' })
    expect((command as { numId: number }).numId).not.toBe(1)
  })

  it('toggle-numbered-list does not collide with an existing numId "1" OR "2"', () => {
    const document = documentWithNumbering(
      new Map([
        ['1', decimalNumberingDef('1')],
        ['2', decimalNumberingDef('2')],
      ]),
    )
    const command = toolbarToCommand({ kind: 'toggle-numbered-list' }, collapsed(pos([0, 0], 0, 0)), document)

    expect(command).toMatchObject({ kind: 'insert-list' })
    const numId = (command as { numId: number }).numId
    expect(numId).not.toBe(1)
    expect(numId).not.toBe(2)
  })

  it('reuses the same numId on repeated bullet-list toggles once one is allocated', () => {
    const withRealNumbering = documentWithNumbering(new Map([['1', decimalNumberingDef('1')]]))
    const first = toolbarToCommand({ kind: 'toggle-bullet-list' }, collapsed(pos([0, 0], 0, 0)), withRealNumbering)
    const allocatedNumId = (first as { numId: number }).numId

    // Simulate ensureListNumbering having registered the allocated id as an
    // Atlas-created bullet definition, the way handleToggleList really does.
    const withAtlasBullet = documentWithNumbering(
      new Map([
        ['1', decimalNumberingDef('1')],
        [
          String(allocatedNumId),
          {
            numId: String(allocatedNumId),
            abstractNumId: `atlas-list-${allocatedNumId}`,
            levels: new Map([[0, { level: 0, format: 'bullet' }]]),
          },
        ],
      ]),
    )
    const second = toolbarToCommand({ kind: 'toggle-bullet-list' }, collapsed(pos([0, 0], 0, 0)), withAtlasBullet)

    expect((second as { numId: number }).numId).toBe(allocatedNumId)
  })

  it('toggle-bullet-list still allocates numId 1 for a document with no existing numbering', () => {
    const document = createDocument([
      Object.freeze({ kind: 'paragraph', children: Object.freeze([createRun('One')]) }) as Paragraph,
    ])
    const command = toolbarToCommand({ kind: 'toggle-bullet-list' }, collapsed(pos([0, 0], 0, 0)), document)

    expect((command as { numId: number }).numId).toBe(1)
  })
})
