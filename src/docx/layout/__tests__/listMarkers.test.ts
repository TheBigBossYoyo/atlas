import { describe, expect, it } from 'vitest'

import type { Document, LvlDef, NumberingDef } from '../../model'
import { createNumberingCounterState, resolveListMarker } from '../listMarkers'

function createLevel(level: number, overrides: Partial<LvlDef> = {}): LvlDef {
  return { level, ...overrides }
}

function createNumberingDoc(defs: ReadonlyArray<NumberingDef>): Document {
  const numbering = new Map<string, NumberingDef>()
  for (const def of defs) {
    numbering.set(def.numId, def)
  }

  return {
    kind: 'document',
    sections: [],
    styles: new Map(),
    numbering,
    comments: new Map(),
    footnotes: new Map(),
    endnotes: new Map(),
    headers: new Map(),
    footers: new Map(),
  }
}

describe('resolveListMarker', () => {
  it('returns undefined when numPr is missing or unresolvable', () => {
    const document = createNumberingDoc([])
    const state = createNumberingCounterState()

    expect(resolveListMarker(undefined, document, state)).toBeUndefined()
    expect(resolveListMarker({ numId: 'missing' }, document, state)).toBeUndefined()
  })

  it('returns undefined when the level format is "none"', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([[0, createLevel(0, { format: 'none', text: { value: '%1.', placeholders: [1] } })]]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)).toBeUndefined()
  })

  it('increments a decimal counter across consecutive paragraphs at the same level', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'decimal', text: { value: '%1.', placeholders: [1] } })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('1.')
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('2.')
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('3.')
  })

  it('formats lowerLetter and lowerRoman levels', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'decimal', text: { value: '%1.', placeholders: [1] } })],
          [1, createLevel(1, { format: 'lowerLetter', text: { value: '%2.', placeholders: [2] } })],
          [2, createLevel(2, { format: 'lowerRoman', text: { value: '%3.', placeholders: [3] } })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    resolveListMarker({ numId: '1', ilvl: 0 }, document, state)
    expect(resolveListMarker({ numId: '1', ilvl: 1 }, document, state)?.text).toBe('a.')
    expect(resolveListMarker({ numId: '1', ilvl: 2 }, document, state)?.text).toBe('i.')
  })

  it('wraps lowerLetter numbering in doubled-letter blocks past z (Word behavior)', () => {
    const document = createNumberingDoc([
      { numId: '1', levels: new Map([[0, createLevel(0, { format: 'lowerLetter' })]]) },
    ])
    const state = createNumberingCounterState()

    for (let i = 0; i < 26; i += 1) {
      resolveListMarker({ numId: '1', ilvl: 0 }, document, state)
    }
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('aa')
  })

  it('formats upperRoman levels', () => {
    const document = createNumberingDoc([
      { numId: '1', levels: new Map([[0, createLevel(0, { format: 'upperRoman' })]]) },
    ])
    const state = createNumberingCounterState()

    for (let i = 0; i < 8; i += 1) {
      resolveListMarker({ numId: '1', ilvl: 0 }, document, state)
    }
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('IX')
  })

  it('renders a literal bullet glyph as-is', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([[0, createLevel(0, { format: 'bullet', text: { value: '•', placeholders: [] } })]]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('•')
  })

  it('maps a Wingdings private-use-area bullet codepoint to a plain Unicode bullet', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [
            0,
            createLevel(0, {
              format: 'bullet',
              text: { value: '', placeholders: [] },
              run: { rFonts: { ascii: 'Wingdings' } },
            }),
          ],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('•')
  })

  it('leaves a private-use-area codepoint alone when not in a known symbol font', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'bullet', text: { value: '', placeholders: [] } })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('')
  })

  it('substitutes multi-level %N placeholders using each ancestor level\'s own counter and format', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'decimal', text: { value: '%1.', placeholders: [1] } })],
          [1, createLevel(1, { format: 'decimal', text: { value: '%1.%2.', placeholders: [1, 2] } })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('1.')
    expect(resolveListMarker({ numId: '1', ilvl: 1 }, document, state)?.text).toBe('1.1.')
    expect(resolveListMarker({ numId: '1', ilvl: 1 }, document, state)?.text).toBe('1.2.')
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('2.')
    // ilvl 1 restarts because its immediate parent (ilvl 0) was just used.
    expect(resolveListMarker({ numId: '1', ilvl: 1 }, document, state)?.text).toBe('2.1.')
  })

  it('does not restart a deeper level whose lvlRestart points past the level that just incremented', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'decimal' })],
          [1, createLevel(1, { format: 'decimal' })],
          // restart: 0 means "restart only when level 0 (ilvl 0) increments",
          // so incrementing level 1 (ilvl 1) should NOT reset this level.
          [2, createLevel(2, { format: 'decimal', restart: 0 })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    resolveListMarker({ numId: '1', ilvl: 0 }, document, state)
    resolveListMarker({ numId: '1', ilvl: 1 }, document, state)
    expect(resolveListMarker({ numId: '1', ilvl: 2 }, document, state)?.text).toBe('1')
    resolveListMarker({ numId: '1', ilvl: 1 }, document, state)
    // Level 1 incrementing again must not restart level 2 (restart: 0).
    expect(resolveListMarker({ numId: '1', ilvl: 2 }, document, state)?.text).toBe('2')
    resolveListMarker({ numId: '1', ilvl: 0 }, document, state)
    // Level 0 incrementing DOES restart level 2 per its explicit restart: 0.
    expect(resolveListMarker({ numId: '1', ilvl: 2 }, document, state)?.text).toBe('1')
  })

  it('honors a custom start value and a startOverride from the num instance', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([[0, createLevel(0, { format: 'decimal', start: 5 })]]),
        levelOverrides: new Map([[0, { level: 0, startOverride: 9 }]]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('9')
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('10')
  })

  it('keeps separate counters for two different numId instances of the same abstract numbering', () => {
    const document = createNumberingDoc([
      { numId: '1', levels: new Map([[0, createLevel(0, { format: 'decimal' })]]) },
      { numId: '2', levels: new Map([[0, createLevel(0, { format: 'decimal' })]]) },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('1')
    expect(resolveListMarker({ numId: '2', ilvl: 0 }, document, state)?.text).toBe('1')
    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.text).toBe('2')
  })

  it('defaults the suffix to tab and honors an explicit override', () => {
    const document = createNumberingDoc([
      {
        numId: '1',
        levels: new Map([
          [0, createLevel(0, { format: 'decimal' })],
          [1, createLevel(1, { format: 'decimal', suffix: 'space' })],
        ]),
      },
    ])
    const state = createNumberingCounterState()

    expect(resolveListMarker({ numId: '1', ilvl: 0 }, document, state)?.suffix).toBe('tab')
    expect(resolveListMarker({ numId: '1', ilvl: 1 }, document, state)?.suffix).toBe('space')
  })
})
