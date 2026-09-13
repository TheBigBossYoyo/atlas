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
  Run,
  Section,
  Style,
} from '../../model'
import { findAll, findNext, findPrev, buildReplaceCommands, buildReplaceAllCommands } from '../Find'
import type { FindOptions } from '../Find'

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

const DEFAULT_OPTS: FindOptions = { caseSensitive: false, wholeWord: false, useRegex: false }

function position(
  paragraphPath: ReadonlyArray<number>,
  runIndex: number,
  charOffset: number,
) {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
}

function createRun(text: string): Run {
  return Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) satisfies Run
}

function createParagraph(texts: ReadonlyArray<string>): Paragraph {
  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze(texts.map(createRun)),
  }) satisfies Paragraph
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

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('findAll', () => {
  it('finds a single match in a single paragraph with one run', () => {
    const doc = createDocument([createParagraph(['hello world'])])
    const matches = findAll(doc, 'world', DEFAULT_OPTS)
    expect(matches).toHaveLength(1)
    expect(matches[0].text).toBe('world')
    // anchor: paragraph [0,0], run 0, offset 6
    expect(matches[0].range.anchor).toEqual(position([0, 0], 0, 6))
    expect(matches[0].range.focus).toEqual(position([0, 0], 0, 11))
  })

  it('finds multiple matches in a single paragraph', () => {
    const doc = createDocument([createParagraph(['cat dog cat'])])
    const matches = findAll(doc, 'cat', DEFAULT_OPTS)
    expect(matches).toHaveLength(2)
    expect(matches[0].range.anchor.charOffset).toBe(0)
    expect(matches[1].range.anchor.charOffset).toBe(8)
  })

  it('finds matches across multiple paragraphs', () => {
    const doc = createDocument([
      createParagraph(['foo bar']),
      createParagraph(['baz foo']),
    ])
    const matches = findAll(doc, 'foo', DEFAULT_OPTS)
    expect(matches).toHaveLength(2)
    expect(matches[0].range.anchor.paragraphPath).toEqual([0, 0])
    expect(matches[1].range.anchor.paragraphPath).toEqual([0, 1])
  })

  it('finds match spanning concatenated runs in the same paragraph', () => {
    // Two runs: 'hel' + 'lo' — query 'hello' should NOT match across runs
    // (matching is per-paragraph, within concatenated text)
    const doc = createDocument([createParagraph(['hel', 'lo world'])])
    const matches = findAll(doc, 'hello', DEFAULT_OPTS)
    expect(matches).toHaveLength(1)
    expect(matches[0].text).toBe('hello')
  })

  it('is case-insensitive by default', () => {
    const doc = createDocument([createParagraph(['Hello WORLD'])])
    const matches = findAll(doc, 'hello', DEFAULT_OPTS)
    expect(matches).toHaveLength(1)
  })

  it('respects caseSensitive: true', () => {
    const doc = createDocument([createParagraph(['Hello World'])])
    const noMatch = findAll(doc, 'hello', { ...DEFAULT_OPTS, caseSensitive: true })
    expect(noMatch).toHaveLength(0)
    const match = findAll(doc, 'Hello', { ...DEFAULT_OPTS, caseSensitive: true })
    expect(match).toHaveLength(1)
  })

  it('respects wholeWord: true — does not match partial word', () => {
    const doc = createDocument([createParagraph(['atlas atlasian'])])
    const matches = findAll(doc, 'atlas', { ...DEFAULT_OPTS, wholeWord: true })
    expect(matches).toHaveLength(1)
    expect(matches[0].range.anchor.charOffset).toBe(0)
  })

  it('respects wholeWord: true — matches standalone word', () => {
    const doc = createDocument([createParagraph(['hello world hello'])])
    const matches = findAll(doc, 'hello', { ...DEFAULT_OPTS, wholeWord: true })
    expect(matches).toHaveLength(2)
  })

  it('supports regex mode', () => {
    const doc = createDocument([createParagraph(['abc 123 xyz'])])
    const matches = findAll(doc, '\\d+', { ...DEFAULT_OPTS, useRegex: true })
    expect(matches).toHaveLength(1)
    expect(matches[0].text).toBe('123')
  })

  it('returns empty array when query is empty', () => {
    const doc = createDocument([createParagraph(['hello'])])
    expect(findAll(doc, '', DEFAULT_OPTS)).toHaveLength(0)
  })

  it('returns empty array when no matches found', () => {
    const doc = createDocument([createParagraph(['hello world'])])
    expect(findAll(doc, 'xyz', DEFAULT_OPTS)).toHaveLength(0)
  })
})

describe('findNext', () => {
  it('returns first match when from is null', () => {
    const doc = createDocument([createParagraph(['cat dog cat'])])
    const m = findNext(doc, 'cat', DEFAULT_OPTS, null)
    expect(m).not.toBeNull()
    expect(m?.range.anchor.charOffset).toBe(0)
  })

  it('wraps around to first match when past last', () => {
    const doc = createDocument([createParagraph(['cat dog cat'])])
    const lastPos = position([0, 0], 0, 9) // after second 'cat'
    const m = findNext(doc, 'cat', DEFAULT_OPTS, lastPos)
    expect(m?.range.anchor.charOffset).toBe(0)
  })
})

describe('findPrev', () => {
  it('returns last match when from is null', () => {
    const doc = createDocument([createParagraph(['cat dog cat'])])
    const m = findPrev(doc, 'cat', DEFAULT_OPTS, null)
    expect(m?.range.anchor.charOffset).toBe(8)
  })

  it('wraps around to last match when before first', () => {
    const doc = createDocument([createParagraph(['cat dog cat'])])
    const firstPos = position([0, 0], 0, 0) // at very start
    const m = findPrev(doc, 'cat', DEFAULT_OPTS, firstPos)
    expect(m?.range.anchor.charOffset).toBe(8)
  })
})

describe('buildReplaceCommands', () => {
  it('returns [DeleteRange, InsertText] in that order', () => {
    const doc = createDocument([createParagraph(['hello world'])])
    const matches = findAll(doc, 'world', DEFAULT_OPTS)
    expect(matches).toHaveLength(1)
    const cmds = buildReplaceCommands(matches[0], 'earth')
    expect(cmds).toHaveLength(2)
    expect(cmds[0].kind).toBe('delete-range')
    expect(cmds[1].kind).toBe('insert-text')
    if (cmds[1].kind === 'insert-text') {
      expect(cmds[1].text).toBe('earth')
    }
  })
})

describe('buildReplaceAllCommands', () => {
  it('returns commands in reverse document order', () => {
    const doc = createDocument([
      createParagraph(['cat']),
      createParagraph(['cat']),
    ])
    const cmds = buildReplaceAllCommands(doc, 'cat', DEFAULT_OPTS, 'dog')
    // 2 matches × 2 commands = 4 commands
    expect(cmds).toHaveLength(4)
    // First pair should target the SECOND paragraph (reverse order)
    const firstDeleteCmd = cmds[0]
    if (firstDeleteCmd.kind === 'delete-range') {
      expect(firstDeleteCmd.range.anchor.paragraphPath).toEqual([0, 1])
    }
    // Second pair targets first paragraph
    const secondDeleteCmd = cmds[2]
    if (secondDeleteCmd.kind === 'delete-range') {
      expect(secondDeleteCmd.range.anchor.paragraphPath).toEqual([0, 0])
    }
  })
})
