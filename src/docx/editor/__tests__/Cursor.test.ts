import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  domPointToPosition,
  findPositionAtClientPoint,
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
