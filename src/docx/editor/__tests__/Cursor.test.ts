import { afterEach, describe, expect, it, vi } from 'vitest'

import type {
  Comment,
  Document as DocxDocument,
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
import {
  caretClientRect,
  domPointToPosition,
  findPositionAtClientPoint,
  positionFromClientPoint,
  positionOnAdjacentLine,
  positionOnePageVertically,
  positionToDomRange,
} from '../Cursor'

type TestDocument = Document & {
  caretRangeFromPoint?: (x: number, y: number) => globalThis.Range | null
  caretPositionFromPoint?: (x: number, y: number) => {
    offsetNode: Node
    offset: number
  } | null
}

afterEach(() => {
  document.body.innerHTML = ''
  Reflect.deleteProperty(document, 'caretRangeFromPoint')
  Reflect.deleteProperty(document, 'caretPositionFromPoint')
  vi.restoreAllMocks()
})

function mountTestDom() {
  document.body.innerHTML = `
    <div id="root">
      <div class="docx-page__line" data-paragraph-path="0,2">
        <span data-run-index="1">Hello</span>
        <span data-run-index="1"> world</span>
        <span data-run-index="2">!</span>
      </div>
    </div>
  `

  const root = document.getElementById('root') as HTMLElement
  const line = root.querySelector('.docx-page__line') as HTMLElement
  const spans = root.querySelectorAll('span')

  return {
    root,
    line,
    firstRun: spans[0] as HTMLSpanElement,
    secondRun: spans[1] as HTMLSpanElement,
    thirdRun: spans[2] as HTMLSpanElement,
  }
}

describe('Cursor helpers', () => {
  it('domPointToPosition maps a text node offset to a document position', () => {
    const { firstRun } = mountTestDom()
    const textNode = firstRun.firstChild as Text

    expect(domPointToPosition(textNode, 3, document)).toEqual({
      paragraphPath: [0, 2],
      runIndex: 1,
      charOffset: 3,
    })
  })

  it('domPointToPosition walks up ancestors for paragraph and run metadata', () => {
    document.body.innerHTML = `
      <div id="root">
        <div class="docx-page__line" data-paragraph-path="5">
          <span data-run-index="4"><strong>Atlas</strong></span>
        </div>
      </div>
    `

    const strong = document.querySelector('strong') as HTMLElement
    const textNode = strong.firstChild as Text

    expect(domPointToPosition(textNode, 2, document)).toEqual({
      paragraphPath: [5],
      runIndex: 4,
      charOffset: 2,
    })
  })

  it('domPointToPosition accumulates offsets across multiple fragments in the same run', () => {
    const { secondRun } = mountTestDom()
    const textNode = secondRun.firstChild as Text

    expect(domPointToPosition(textNode, 3, document)).toEqual({
      paragraphPath: [0, 2],
      runIndex: 1,
      charOffset: 8,
    })
  })

  it('domPointToPosition returns null when metadata is missing', () => {
    document.body.innerHTML = '<div id="root"><span>No mapping</span></div>'

    const textNode = document.querySelector('span')?.firstChild as Text

    expect(domPointToPosition(textNode, 1, document)).toBeNull()
  })

  it('positionToDomRange returns the matching text node and offset', () => {
    const { firstRun } = mountTestDom()
    const textNode = firstRun.firstChild as Text
    const point = positionToDomRange(
      {
        paragraphPath: [0, 2],
        runIndex: 1,
        charOffset: 4,
      },
      document.getElementById('root') as HTMLElement,
    )

    expect(point).not.toBeNull()
    expect(point).toEqual({ node: textNode, offset: 4 })
  })

  it('positionToDomRange can land in later fragments of the same run', () => {
    const { secondRun } = mountTestDom()
    const textNode = secondRun.firstChild as Text
    const point = positionToDomRange(
      {
        paragraphPath: [0, 2],
        runIndex: 1,
        charOffset: 7,
      },
      document.getElementById('root') as HTMLElement,
    )

    expect(point).toEqual({ node: textNode, offset: 2 })
  })

  it('positionToDomRange returns null when no run fragment matches the position', () => {
    const { root } = mountTestDom()

    expect(
      positionToDomRange(
        {
          paragraphPath: [99],
          runIndex: 1,
          charOffset: 0,
        },
        root,
      ),
    ).toBeNull()
  })

  it('findPositionAtClientPoint uses caretRangeFromPoint when available', () => {
    const { root, firstRun } = mountTestDom()
    const textNode = firstRun.firstChild as Text
    const range = document.createRange()
    range.setStart(textNode, 2)
    range.setEnd(textNode, 2)

    ;(document as TestDocument).caretRangeFromPoint = vi.fn().mockReturnValue(range)

    expect(findPositionAtClientPoint(10, 20, document, root)).toEqual({
      paragraphPath: [0, 2],
      runIndex: 1,
      charOffset: 2,
    })
  })

  it('findPositionAtClientPoint falls back to caretPositionFromPoint', () => {
    const { root, thirdRun } = mountTestDom()
    const textNode = thirdRun.firstChild as Text

    ;(document as TestDocument).caretPositionFromPoint = vi.fn().mockReturnValue({
      offsetNode: textNode,
      offset: 1,
    })

    expect(findPositionAtClientPoint(5, 6, document, root)).toEqual({
      paragraphPath: [0, 2],
      runIndex: 2,
      charOffset: 1,
    })
  })

  it('findPositionAtClientPoint returns null for points outside the editor root', () => {
    const { root } = mountTestDom()
    const outside = document.createElement('span')
    outside.textContent = 'outside'
    document.body.appendChild(outside)

    ;(document as TestDocument).caretPositionFromPoint = vi.fn().mockReturnValue({
      offsetNode: outside.firstChild as Text,
      offset: 1,
    })

    expect(findPositionAtClientPoint(5, 6, document, root)).toBeNull()
  })
})

// ─── DOCX-17 — table-cell hit-testing ──────────────────────────────────────
//
// Table cell lines render with no `data-paragraph-path` of their own (see
// `Cursor.ts`'s "Table-cell paragraph-path resolution" section) — every
// helper above resolves fine without a document model because none of their
// fixtures are inside a table. These fixtures mimic the REAL rendered DOM's
// table markup (`PageView.tsx`'s `renderPageTable`/`renderPageTableRow`:
// `.docx-page__table-wrapper[data-block-path]`, `tr[data-source-row]`,
// `td[data-col]`) instead of relying on `data-paragraph-path` at all, and
// pass a document model in, so they only pass if the DOM-plus-model
// fallback resolves correctly — the same path a click into a real table
// cell takes.

function createCellParagraph(text: string): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([
      Object.freeze({ kind: 'run', children: Object.freeze([Object.freeze({ kind: 'text', value: text })]) }),
    ]),
  }) satisfies Paragraph
}

/** A 1-row, 2-cell table, each cell holding one paragraph of `cellText`. */
function createOneRowTable(cellTexts: ReadonlyArray<string>): Table {
  return {
    kind: 'table',
    tblGrid: Object.freeze(cellTexts.map(() => twip(1440))),
    rows: Object.freeze([
      Object.freeze({
        kind: 'table-row',
        cells: Object.freeze(cellTexts.map((text) => Object.freeze({ kind: 'table-cell', blocks: Object.freeze([createCellParagraph(text)]) }))),
      }),
    ]),
  } satisfies Table
}

function createDocumentWithTableAtBlock0(table: Table): DocxDocument {
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
  }) satisfies DocxDocument
}

function rect(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left,
    top,
    right: left + width,
    bottom: top + height,
    width,
    height,
    x: left,
    y: top,
    toJSON: () => ({}),
  }
}

/**
 * A single-row, 2-cell `<table>` rendered exactly as `PageView.tsx` renders
 * one — `data-block-path` on the wrapper, `data-source-row`/`data-col` on
 * the row/cells, one `.docx-page__line` per cell with no
 * `data-paragraph-path` of its own — with `getBoundingClientRect` stubbed on
 * each line/run so `positionFromClientPoint`'s geometry-based hit-testing
 * has real rects to compare against (jsdom's own layout always reports
 * zero-size rects).
 */
function mountTableDom(cellTexts: ReadonlyArray<string>) {
  document.body.innerHTML = `
    <div id="root">
      <div class="docx-page__table-wrapper" data-block-path="0">
        <table class="docx-page__table">
          <tbody>
            <tr data-source-row="0">
              ${cellTexts
                .map(
                  (text, cellIdx) => `
                <td data-col="${cellIdx}">
                  <div class="docx-page__table-cell-content">
                    <div class="docx-page__line">
                      <span data-run-index="0" data-char-start="0" data-char-end="${text.length}">${text}</span>
                    </div>
                  </div>
                </td>`,
                )
                .join('')}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `

  const root = document.getElementById('root') as HTMLElement
  const lines = Array.from(root.querySelectorAll<HTMLElement>('.docx-page__line'))
  const spans = Array.from(root.querySelectorAll<HTMLElement>('[data-run-index]'))

  // Cell N's line/span sits at x = [N * 100, N * 100 + 40), y = [0, 20) —
  // spaced well apart so a click's x coordinate unambiguously picks one cell.
  lines.forEach((line, index) => {
    line.getBoundingClientRect = () => rect(index * 100, 0, 40, 20)
  })
  spans.forEach((span, index) => {
    span.getBoundingClientRect = () => rect(index * 100, 0, 40, 20)
  })

  return { root, lines, spans }
}

describe('DOCX-17 — table-cell hit-testing', () => {
  it('positionFromClientPoint resolves a click inside a table cell to that cell paragraph', () => {
    const document = createDocumentWithTableAtBlock0(createOneRowTable(['r0c0', 'r0c1']))
    const { root } = mountTableDom(['r0c0', 'r0c1'])

    // Click at the right edge of cell 1's span (x=140, its rect is
    // [100,140)): past the last character, landing at the END of "r0c1".
    const position = positionFromClientPoint(140, 10, root, document)

    expect(position).toEqual({
      paragraphPath: [0, 0, 1, 0], // table block 0, row 0, cell 1, paragraph 0
      runIndex: 0,
      charOffset: 'r0c1'.length,
    })
  })

  it('positionFromClientPoint resolves a click at the start of a cell to charOffset 0', () => {
    const document = createDocumentWithTableAtBlock0(createOneRowTable(['r0c0', 'r0c1']))
    const { root } = mountTableDom(['r0c0', 'r0c1'])

    // Click at cell 0's left edge (x=0) — before any character.
    const position = positionFromClientPoint(0, 10, root, document)

    expect(position).toEqual({
      paragraphPath: [0, 0, 0, 0], // table block 0, row 0, cell 0, paragraph 0
      runIndex: 0,
      charOffset: 0,
    })
  })

  it('positionFromClientPoint without a document model cannot resolve a table cell (pre-fix behavior)', () => {
    // No `docModel` passed — this is exactly what every caller did before
    // DOCX-17 threaded one through, and table lines carry no
    // `data-paragraph-path` of their own for the old DOM-only path to fall
    // back on. Asserting the OLD failure mode here pins down that the new
    // behavior above is really coming from the `docModel` fallback, not
    // from some unrelated change to hit-testing in general.
    const { root } = mountTableDom(['r0c0', 'r0c1'])

    expect(positionFromClientPoint(140, 10, root)).toBeNull()
  })

  it('positionToDomRange resolves a table-cell Position back to that cell’s line', () => {
    const document = createDocumentWithTableAtBlock0(createOneRowTable(['r0c0', 'r0c1']))
    const { root, spans } = mountTableDom(['r0c0', 'r0c1'])

    const point = positionToDomRange(
      { paragraphPath: [0, 0, 1, 0], runIndex: 0, charOffset: 2 },
      root,
      document,
    )

    expect(point).not.toBeNull()
    expect(point?.node).toBe(spans[1].firstChild)
    expect(point?.offset).toBe(2)
  })
})

// ─── DOCX-1 — vertical caret movement geometry ─────────────────────────────
//
// jsdom never lays anything out, so every helper below stubs
// `getBoundingClientRect` on the elements it cares about — the same pattern
// `mountTableDom` above already uses. Every line shares the same x-range
// ([20, 120)) so a `goalX` of exactly 20 (a line's own left edge) always
// resolves to charOffset 0 via `positionFromClientPoint`'s own left-edge
// special case, without needing real sub-character `Range` geometry (which
// jsdom can't provide either).

/** Three stacked lines, each one paragraph, non-overlapping vertical bands:
 * line 0 at y=[0,20), line 1 at y=[40,60), line 2 at y=[80,100). */
function mountThreeLineDom() {
  document.body.innerHTML = `
    <div id="root">
      <div class="docx-page__line" data-paragraph-path="0">
        <span data-run-index="0" data-char-start="0" data-char-end="4">zero</span>
      </div>
      <div class="docx-page__line" data-paragraph-path="1">
        <span data-run-index="0" data-char-start="0" data-char-end="3">one</span>
      </div>
      <div class="docx-page__line" data-paragraph-path="2">
        <span data-run-index="0" data-char-start="0" data-char-end="3">two</span>
      </div>
    </div>
  `

  const root = document.getElementById('root') as HTMLElement
  const lines = Array.from(root.querySelectorAll<HTMLElement>('.docx-page__line'))
  const spans = Array.from(root.querySelectorAll<HTMLElement>('[data-run-index]'))
  const bands: ReadonlyArray<[number, number]> = [[0, 20], [40, 60], [80, 100]]

  lines.forEach((line, index) => {
    const [top, bottom] = bands[index]
    line.getBoundingClientRect = () => rect(20, top, 100, bottom - top)
  })
  spans.forEach((span, index) => {
    const [top, bottom] = bands[index]
    span.getBoundingClientRect = () => rect(20, top, 100, bottom - top)
  })

  return { root, lines, spans }
}

describe('DOCX-1 — positionOnAdjacentLine', () => {
  it('moves down to the nearest line below', () => {
    const { root } = mountThreeLineDom()
    const fromPosition = { paragraphPath: [1], runIndex: 0, charOffset: 1 } // caret on line 1

    const position = positionOnAdjacentLine(fromPosition, 20, 'down', root)

    expect(position).toEqual({ paragraphPath: [2], runIndex: 0, charOffset: 0 })
  })

  it('moves up to the nearest line above', () => {
    const { root } = mountThreeLineDom()
    const fromPosition = { paragraphPath: [1], runIndex: 0, charOffset: 1 } // caret on line 1

    const position = positionOnAdjacentLine(fromPosition, 20, 'up', root)

    expect(position).toEqual({ paragraphPath: [0], runIndex: 0, charOffset: 0 })
  })

  it('returns null moving up from the first line (top of document)', () => {
    const { root } = mountThreeLineDom()
    const fromPosition = { paragraphPath: [0], runIndex: 0, charOffset: 1 } // caret on line 0

    expect(positionOnAdjacentLine(fromPosition, 20, 'up', root)).toBeNull()
  })

  it('returns null moving down from the last line (bottom of document)', () => {
    const { root } = mountThreeLineDom()
    const fromPosition = { paragraphPath: [2], runIndex: 0, charOffset: 1 } // caret on line 2

    expect(positionOnAdjacentLine(fromPosition, 20, 'down', root)).toBeNull()
  })

  it('never lands back on its own line, even when its own line’s rect reports a center a couple of px off from the caret rect (line-height/baseline offsets)', () => {
    const { root, lines } = mountThreeLineDom()
    // Line 1's own bounding rect (used as the "current line" anchor) is
    // shifted 3px from its band's true center — bigger than the 1px
    // same-line epsilon — mimicking the real caret-vs-line-rect mismatch
    // DOCX-1's own e2e repro hit (a run's precise caret rect and its
    // enclosing line's own rect don't share an exact center).
    lines[1].getBoundingClientRect = () => rect(20, 37, 100, 20) // band [40,60) shifted to [37,57)
    const fromPosition = { paragraphPath: [1], runIndex: 0, charOffset: 1 }

    expect(positionOnAdjacentLine(fromPosition, 20, 'up', root)).toEqual({
      paragraphPath: [0],
      runIndex: 0,
      charOffset: 0,
    })
    expect(positionOnAdjacentLine(fromPosition, 20, 'down', root)).toEqual({
      paragraphPath: [2],
      runIndex: 0,
      charOffset: 0,
    })
  })

  it('keeps the given goal x across the move (lands at the target line under the same column)', () => {
    const { root } = mountThreeLineDom()
    const fromPosition = { paragraphPath: [0], runIndex: 0, charOffset: 1 } // caret on line 0

    // x=120 is line 1's right edge — its own special-case "past the last
    // character" branch in positionFromClientPoint, landing at charOffset 3.
    const position = positionOnAdjacentLine(fromPosition, 120, 'down', root)

    expect(position).toEqual({ paragraphPath: [1], runIndex: 0, charOffset: 3 })
  })
})

describe('DOCX-1 — positionOnePageVertically', () => {
  it('moving down by a viewport height lands on the nearest line to that offset', () => {
    const { root } = mountThreeLineDom()
    const fromRect = rect(20, 0, 0, 20) // caret on line 0, top=0

    // targetY = 0 + 100 = 100 — closest to line 2's band [80,100).
    const position = positionOnePageVertically(fromRect, 20, 'down', 100, root)

    expect(position).toEqual({ paragraphPath: [2], runIndex: 0, charOffset: 0 })
  })

  it('moving up by a viewport height clamps to the first line when it overshoots the document', () => {
    const { root } = mountThreeLineDom()
    const fromRect = rect(20, 80, 0, 20) // caret on line 2, top=80

    // targetY = 80 - 100 = -20 — no line up there, nearest is line 0.
    const position = positionOnePageVertically(fromRect, 20, 'up', 100, root)

    expect(position).toEqual({ paragraphPath: [0], runIndex: 0, charOffset: 0 })
  })
})

describe('DOCX-1 — caretClientRect', () => {
  it('falls back to the containing run element’s rect when the DOM Range reports no client rects (jsdom)', () => {
    const { root, spans } = mountThreeLineDom()
    spans[1].getBoundingClientRect = () => rect(20, 40, 100, 20)

    const caretRect = caretClientRect({ paragraphPath: [1], runIndex: 0, charOffset: 1 }, root)

    expect(caretRect).toMatchObject({ left: 20, top: 40, right: 120, bottom: 60, width: 100, height: 20 })
  })

  it('returns null when the position cannot be mapped into the DOM at all', () => {
    const { root } = mountThreeLineDom()

    expect(caretClientRect({ paragraphPath: [99], runIndex: 0, charOffset: 0 }, root)).toBeNull()
  })
})
