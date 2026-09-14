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
} from '../../model'
import { applyCommand } from '../commands'
import type { Position } from '../commandTypes'
import { buildSpellCheckReplacement } from '../spellCheckReplace'

function createParagraph(text: string): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([
      Object.freeze({ kind: 'run', children: Object.freeze([Object.freeze({ kind: 'text', value: text })]) }),
    ]),
  }) as Paragraph
}

function createDocument(paragraphs: ReadonlyArray<Paragraph>): Document {
  const section = Object.freeze({ kind: 'section', props: {}, blocks: Object.freeze([...paragraphs]) }) satisfies Section

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

function pos(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

describe('buildSpellCheckReplacement', () => {
  it('builds delete+insert commands that fix the misspelling in the model', () => {
    const document = createDocument([createParagraph('I went ot the store')])

    const result = buildSpellCheckReplacement(document, 'ot', 'to', null)

    expect(result).not.toBeNull()
    let working = document
    for (const command of result!.commands) {
      working = applyCommand(working, command).document
    }
    const text = (working.sections[0].blocks[0] as Paragraph).children
      .map((c) => (c.kind === 'run' ? c.children.map((t) => (t.kind === 'text' ? t.value : '')).join('') : ''))
      .join('')
    expect(text).toBe('I went to the store')
  })

  it('only matches the misspelling as a whole word', () => {
    const document = createDocument([createParagraph('rotor')])

    // "ot" appears inside "rotor" but not as a whole word — nothing to fix.
    const result = buildSpellCheckReplacement(document, 'ot', 'to', null)
    expect(result).toBeNull()
  })

  it('returns null when the replacement equals the misspelling', () => {
    const document = createDocument([createParagraph('teh')])
    expect(buildSpellCheckReplacement(document, 'teh', 'teh', null)).toBeNull()
  })

  it('returns null when the word does not appear in the document', () => {
    const document = createDocument([createParagraph('hello world')])
    expect(buildSpellCheckReplacement(document, 'teh', 'the', null)).toBeNull()
  })

  it('picks the occurrence at or after the hint position when there are several', () => {
    const document = createDocument([createParagraph('teh cat sat on teh mat')])
    // Find.ts addresses paragraphs as [sectionIndex, blockIndex]. Hint sits
    // right after the first "teh " (offset 5) — should pick the SECOND
    // occurrence, not restart from the very beginning of the doc.
    const hint = pos([0, 0], 0, 5)

    const result = buildSpellCheckReplacement(document, 'teh', 'the', hint)

    expect(result).not.toBeNull()
    // First occurrence (offset 0-3) must be untouched by the delete target.
    expect(result!.commands[0]).toMatchObject({
      kind: 'delete-range',
      range: { anchor: pos([0, 0], 0, 15), focus: pos([0, 0], 0, 18) },
    })
  })

  it('computes the resulting cursor right after the replacement text', () => {
    const document = createDocument([createParagraph('teh')])
    const result = buildSpellCheckReplacement(document, 'teh', 'the', null)

    expect(result?.range).toEqual({ anchor: pos([0, 0], 0, 3), focus: pos([0, 0], 0, 3) })
  })

  it('returns null for empty misspelled or replacement text', () => {
    const document = createDocument([createParagraph('teh')])
    expect(buildSpellCheckReplacement(document, '', 'the', null)).toBeNull()
    expect(buildSpellCheckReplacement(document, 'teh', '', null)).toBeNull()
  })
})
