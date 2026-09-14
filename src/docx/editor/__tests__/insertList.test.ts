import { describe, expect, it } from 'vitest'

import type { DocxBundle } from '../../index'
import type { Document, Paragraph, Section } from '../../model/document'
import { ensureListNumbering } from '../insertList'

function emptyMaps() {
  return {
    styles: new Map(),
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

function makeBundle(): DocxBundle {
  const paragraph: Paragraph = { kind: 'paragraph', children: [{ kind: 'run', children: [{ kind: 'text', value: 'Item' }] }] }
  const section: Section = { kind: 'section', props: {}, blocks: [paragraph] }
  const document: Document = { kind: 'document', sections: [section], ...emptyMaps() }
  return { document, rawArchive: new Map<string, Uint8Array>() }
}

describe('ensureListNumbering', () => {
  it('creates a bullet abstractNum/num pair when numId 1 is not yet defined', () => {
    const bundle = makeBundle()
    const next = ensureListNumbering(bundle, 1, 'bullet')

    expect(next.document.numbering.has('1')).toBe(true)
    expect(next.numberingPart?.nums.has('1')).toBe(true)
    const abstractNumId = next.numberingPart!.nums.get('1')!.abstractNumId!
    const abstractNum = next.numberingPart!.abstractNums.get(abstractNumId)!
    expect(abstractNum.levels.get(0)?.format).toBe('bullet')
    expect(abstractNum.levels.get(0)?.text?.value).toBe('•')
  })

  it('creates a decimal abstractNum/num pair when numId 2 is not yet defined', () => {
    const bundle = makeBundle()
    const next = ensureListNumbering(bundle, 2, 'number')

    const abstractNumId = next.numberingPart!.nums.get('2')!.abstractNumId!
    const abstractNum = next.numberingPart!.abstractNums.get(abstractNumId)!
    expect(abstractNum.levels.get(0)?.format).toBe('decimal')
    expect(abstractNum.levels.get(0)?.text?.value).toBe('%1.')
  })

  it('reuses an existing numbering definition instead of overwriting it', () => {
    const bundle = makeBundle()
    const withNumbering = ensureListNumbering(bundle, 1, 'bullet')
    const again = ensureListNumbering(withNumbering, 1, 'number')

    // Still the original bullet definition — not replaced by a decimal one.
    expect(again.document.numbering).toBe(withNumbering.document.numbering)
    expect(again.numberingPart).toBe(withNumbering.numberingPart)
  })

  it('does not mutate the input bundle', () => {
    const bundle = makeBundle()
    ensureListNumbering(bundle, 1, 'bullet')

    expect(bundle.document.numbering.has('1')).toBe(false)
    expect(bundle.numberingPart).toBeUndefined()
  })

  it('preserves an existing numberingPart while adding the new definition', () => {
    const bundle: DocxBundle = {
      ...makeBundle(),
      numberingPart: {
        abstractNums: new Map([['0', { abstractNumId: '0', levels: new Map() }]]),
        nums: new Map([['9', { numId: '9', abstractNumId: '0' }]]),
      },
    }

    const next = ensureListNumbering(bundle, 1, 'bullet')

    expect(next.numberingPart!.nums.has('9')).toBe(true)
    expect(next.numberingPart!.nums.has('1')).toBe(true)
  })
})
