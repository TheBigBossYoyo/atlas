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
  Section,
  Style,
  Table,
} from '../../model'
import { twip } from '../../model'
import type { Position } from '../commandTypes'
import { History } from '../History'
import { handleKeyDown } from '../Input'

// DXE-14 — Tab/Shift+Tab table-cell navigation.

function tabEvent(shift = false): KeyboardEvent {
  return new KeyboardEvent('keydown', { key: 'Tab', shiftKey: shift })
}

function pos(paragraphPath: ReadonlyArray<number>, runIndex = 0, charOffset = 0): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

function createParagraph(text: string): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([
      Object.freeze({ kind: 'run', children: Object.freeze([Object.freeze({ kind: 'text', value: text })]) }),
    ]),
  }) satisfies Paragraph
}

function createTable(rows: number, cols: number): Table {
  return {
    kind: 'table',
    tblGrid: Object.freeze(Array.from({ length: cols }, () => twip(1440))),
    rows: Object.freeze(
      Array.from({ length: rows }, (_row, rowIndex) =>
        Object.freeze({
          kind: 'table-row',
          cells: Object.freeze(
            Array.from({ length: cols }, (_cell, colIndex) =>
              Object.freeze({
                kind: 'table-cell',
                blocks: Object.freeze([createParagraph(`r${rowIndex}c${colIndex}`)]),
              }),
            ),
          ),
        }),
      ),
    ),
  } satisfies Table
}

function createDocumentWithTable(table: Table): Document {
  const section = Object.freeze({
    kind: 'section',
    props: {},
    blocks: Object.freeze([table]),
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

// Position paths address `[sectionIndex, tableIndex, rowIndex, cellIndex,
// paragraphIndexWithinCell]` — matching `applyInsertTable`'s own cursor
// addressing (`insert-table`'s inverse cursor construction), not a bespoke
// shape invented for this test.

describe('DXE-14 — Tab navigates between table cells', () => {
  it('moves to the next cell in the same row, selecting its full content', () => {
    const document = createDocumentWithTable(createTable(2, 2))
    const history = new History()

    const result = handleKeyDown(tabEvent(), {
      document,
      range: { anchor: pos([0, 0, 0, 0, 0]), focus: pos([0, 0, 0, 0, 0]) },
      history,
    })

    expect(result).not.toBeNull()
    expect(result?.range?.anchor).toEqual(pos([0, 0, 0, 1, 0]))
    expect(result?.range?.focus).toEqual(pos([0, 0, 0, 1, 0], 0, 'r0c1'.length))
    expect(result?.document).toBe(document) // pure navigation, no document mutation
  })

  it('moves to the first cell of the next row from the last cell of a row', () => {
    const document = createDocumentWithTable(createTable(2, 2))
    const history = new History()

    const result = handleKeyDown(tabEvent(), {
      document,
      range: { anchor: pos([0, 0, 0, 1, 0]), focus: pos([0, 0, 0, 1, 0]) },
      history,
    })

    expect(result?.range?.anchor).toEqual(pos([0, 0, 1, 0, 0]))
  })

  it('appends a new row and moves into it from the last cell of the last row', () => {
    const document = createDocumentWithTable(createTable(2, 2))
    const history = new History()

    const result = handleKeyDown(tabEvent(), {
      document,
      range: { anchor: pos([0, 0, 1, 1, 0]), focus: pos([0, 0, 1, 1, 0]) },
      history,
    })

    expect(result).not.toBeNull()
    const table = result?.document.sections[0].blocks[0] as Table
    expect(table.rows).toHaveLength(3)
    expect(result?.range?.anchor).toEqual(pos([0, 0, 2, 0, 0]))

    // Undoable as a normal History step.
    const undone = history.undo(result!.document)
    const restoredTable = undone?.document.sections[0].blocks[0] as Table
    expect(restoredTable.rows).toHaveLength(2)
  })

  it('Shift+Tab moves to the previous cell', () => {
    const document = createDocumentWithTable(createTable(2, 2))
    const history = new History()

    const result = handleKeyDown(tabEvent(true), {
      document,
      range: { anchor: pos([0, 0, 0, 1, 0]), focus: pos([0, 0, 0, 1, 0]) },
      history,
    })

    expect(result?.range?.anchor).toEqual(pos([0, 0, 0, 0, 0]))
  })

  it('Shift+Tab moves to the last cell of the previous row', () => {
    const document = createDocumentWithTable(createTable(2, 2))
    const history = new History()

    const result = handleKeyDown(tabEvent(true), {
      document,
      range: { anchor: pos([0, 0, 1, 0, 0]), focus: pos([0, 0, 1, 0, 0]) },
      history,
    })

    expect(result?.range?.anchor).toEqual(pos([0, 0, 0, 1, 0]))
  })

  it('Shift+Tab at the very first cell is a no-op (does not insert a tab character)', () => {
    const document = createDocumentWithTable(createTable(1, 1))
    const history = new History()

    const result = handleKeyDown(tabEvent(true), {
      document,
      range: { anchor: pos([0, 0, 0, 0, 0]), focus: pos([0, 0, 0, 0, 0]) },
      history,
    })

    expect(result).toBeNull()
  })
})
