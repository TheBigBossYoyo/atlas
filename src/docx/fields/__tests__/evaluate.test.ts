import { describe, expect, it } from 'vitest'

import type { Field, FieldType } from '../../model'
import { evaluateFieldText, parseHyperlinkField } from '../evaluate'
import { parseFieldInstruction } from '../instruction'
import type { FieldEvaluationContext } from '../types'

function makeField(fieldType: FieldType, instruction: string): Field {
  return { kind: 'field', fieldType, instruction, result: [] }
}

function makeContext(overrides: Partial<FieldEvaluationContext> = {}): FieldEvaluationContext {
  return {
    bookmarkText: new Map(),
    sequenceCounters: new Map(),
    ...overrides,
  }
}

describe('evaluateFieldText', () => {
  it('evaluates DATE using the default picture when no \\@ switch is present', () => {
    const context = makeContext({ now: new Date(2026, 0, 7) })
    expect(evaluateFieldText(makeField('DATE', 'DATE'), [], context)).toBe('1/7/2026')
  })

  it('evaluates DATE using a \\@ picture switch', () => {
    const context = makeContext({ now: new Date(2026, 0, 7) })
    const field = makeField('DATE', 'DATE \\@ "MMMM d, yyyy"')
    expect(evaluateFieldText(field, [], context)).toBe('January 7, 2026')
  })

  it('evaluates TIME using the default picture', () => {
    const context = makeContext({ now: new Date(2026, 0, 7, 9, 5) })
    expect(evaluateFieldText(makeField('TIME', 'TIME'), [], context)).toBe('9:05 am')
  })

  it('evaluates AUTHOR from context.author, leaving it unset (undefined) when unknown', () => {
    expect(evaluateFieldText(makeField('AUTHOR', 'AUTHOR'), [], makeContext())).toBeUndefined()
    expect(
      evaluateFieldText(makeField('AUTHOR', 'AUTHOR'), [], makeContext({ author: 'A. Author' })),
    ).toBe('A. Author')
  })

  it('evaluates TITLE from context.title', () => {
    expect(
      evaluateFieldText(makeField('TITLE', 'TITLE'), [], makeContext({ title: 'Atlas Report' })),
    ).toBe('Atlas Report')
  })

  it('evaluates REF from the bookmark\'s current text', () => {
    const context = makeContext({ bookmarkText: new Map([['_Ref1', 'Section 2']]) })
    expect(evaluateFieldText(makeField('REF', 'REF _Ref1'), [], context)).toBe('Section 2')
  })

  it('leaves REF unchanged when the bookmark is unknown', () => {
    expect(evaluateFieldText(makeField('REF', 'REF _Unknown'), [], makeContext())).toBeUndefined()
  })

  it('evaluates PAGEREF from the bookmark\'s current page', () => {
    const context = makeContext({ bookmarkPage: new Map([['_Ref1', 4]]) })
    expect(evaluateFieldText(makeField('PAGEREF', 'PAGEREF _Ref1'), [], context)).toBe('4')
  })

  it('leaves PAGEREF unchanged when bookmarkPage is not supplied (layout has not run)', () => {
    expect(evaluateFieldText(makeField('PAGEREF', 'PAGEREF _Ref1'), [], makeContext())).toBeUndefined()
  })

  it('evaluates NUMPAGES from context.pageCount', () => {
    expect(evaluateFieldText(makeField('NUMPAGES', 'NUMPAGES'), [], makeContext({ pageCount: 12 }))).toBe('12')
  })

  it('leaves NUMPAGES unchanged when pageCount is unknown', () => {
    expect(evaluateFieldText(makeField('NUMPAGES', 'NUMPAGES'), [], makeContext())).toBeUndefined()
  })

  it('evaluates PAGE via context.currentPageOf, passed the field\'s own paragraphPath', () => {
    const context = makeContext({ currentPageOf: (path) => (path.join(',') === '0,3' ? 2 : undefined) })
    expect(evaluateFieldText(makeField('PAGE', 'PAGE'), [0, 3], context)).toBe('2')
    expect(evaluateFieldText(makeField('PAGE', 'PAGE'), [0, 9], context)).toBeUndefined()
  })

  it('increments a SEQ counter across successive evaluations of the same sequence name', () => {
    const context = makeContext()
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('1')
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('2')
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('3')
  })

  it('keeps independent counters per sequence name', () => {
    const context = makeContext()
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('1')
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Table'), [], context)).toBe('1')
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('2')
  })

  it('repeats (does not increment) a SEQ counter when \\c is present', () => {
    const context = makeContext()
    evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)
    evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure \\c'), [], context)).toBe('2')
    // A following plain SEQ resumes incrementing from the repeated value.
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('3')
  })

  it('resets a SEQ counter to \\r\'s value', () => {
    const context = makeContext()
    evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure \\r 10'), [], context)).toBe('10')
    expect(evaluateFieldText(makeField('SEQ', 'SEQ Figure'), [], context)).toBe('11')
  })

  it('leaves HYPERLINK and TOC and unknown fields unevaluated (their cached display text is left as-is)', () => {
    expect(evaluateFieldText(makeField('HYPERLINK', 'HYPERLINK "https://example.com"'), [], makeContext())).toBeUndefined()
    expect(evaluateFieldText(makeField('TOC', 'TOC \\o "1-3"'), [], makeContext())).toBeUndefined()
    expect(evaluateFieldText(makeField('unknown', 'MACROBUTTON Foo'), [], makeContext())).toBeUndefined()
  })
})

describe('parseHyperlinkField', () => {
  it('extracts a plain external target', () => {
    const info = parseHyperlinkField(parseFieldInstruction('HYPERLINK "https://example.com"'))
    expect(info).toMatchObject({ target: 'https://example.com', isLocalAnchor: false })
  })

  it('extracts a local-anchor target from \\l', () => {
    const info = parseHyperlinkField(parseFieldInstruction('HYPERLINK \\l "Top"'))
    expect(info).toMatchObject({ target: 'Top', isLocalAnchor: true })
  })

  it('extracts tooltip (\\o) and target frame (\\t) switches', () => {
    const info = parseHyperlinkField(
      parseFieldInstruction('HYPERLINK "https://example.com" \\o "Visit" \\t "_blank"'),
    )
    expect(info).toMatchObject({ tooltip: 'Visit', targetFrame: '_blank' })
  })

  it('returns undefined when there is no target at all', () => {
    expect(parseHyperlinkField(parseFieldInstruction('HYPERLINK'))).toBeUndefined()
  })
})
