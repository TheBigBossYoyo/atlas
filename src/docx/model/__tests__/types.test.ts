import { describe, expect, expectTypeOf, it } from 'vitest'

import {
  assertNever,
  BlockKind,
  hexColor,
  InlineKind,
  halfPoint,
  pct,
  twip,
} from '../index'
import type {
  Block,
  Comment,
  Document,
  Endnote,
  Footer,
  Footnote,
  Header,
  Inline,
  NumberingDef,
  Paragraph,
  RunChild,
  RunProps,
  Section,
  Style,
  TextNode,
} from '../index'

function describeBlock(block: Block): string {
  switch (block.kind) {
    case BlockKind.paragraph:
      return 'paragraph'
    case BlockKind.table:
      return 'table'
    case BlockKind.unknown:
      return 'unknown'
    default:
      return assertNever(block)
  }
}

function describeInline(inline: Inline): string {
  switch (inline.kind) {
    case InlineKind.run:
      return 'run'
    case InlineKind.hyperlink:
      return 'hyperlink'
    case InlineKind.bookmark:
      return 'bookmark'
    case InlineKind.commentRange:
      return 'comment-range'
    case InlineKind.commentReference:
      return 'comment-reference'
    case InlineKind.footnoteReference:
      return 'footnote-reference'
    case InlineKind.endnoteReference:
      return 'endnote-reference'
    case InlineKind.drawing:
      return 'drawing'
    case InlineKind.field:
      return 'field'
    case InlineKind.text:
      return 'text'
    case InlineKind.tab:
      return 'tab'
    case InlineKind.break:
      return 'break'
    case InlineKind.unknown:
      return 'unknown'
    default:
      return assertNever(inline)
  }
}

function mergeRunProps(defaults: RunProps, direct: RunProps): RunProps {
  return { ...defaults, ...direct }
}

const sampleText = Object.freeze({
  kind: 'text',
  value: 'Atlas',
}) satisfies TextNode

const sampleRunChildren = Object.freeze([sampleText]) satisfies ReadonlyArray<RunChild>

const sampleParagraph = Object.freeze({
  kind: 'paragraph',
  children: Object.freeze([
    Object.freeze({
      kind: 'run',
      props: Object.freeze({ bold: true, sz: halfPoint(24) }),
      children: sampleRunChildren,
    }),
  ]),
}) satisfies Paragraph

const sampleSection = Object.freeze({
  kind: 'section',
  props: Object.freeze({
    pgSz: Object.freeze({ w: twip(12240), h: twip(15840) }),
  }),
  blocks: Object.freeze([sampleParagraph]),
}) satisfies Section

const sampleDocument = Object.freeze({
  kind: 'document',
  sections: Object.freeze([sampleSection]),
  styles: new Map<string, Style>(),
  numbering: new Map<string, NumberingDef>(),
  comments: new Map<string, Comment>(),
  footnotes: new Map<string, Footnote>(),
  endnotes: new Map<string, Endnote>(),
  headers: new Map<string, Header>(),
  footers: new Map<string, Footer>(),
}) satisfies Document

describe('docx model types', () => {
  it('supports exhaustive Block switches via assertNever', () => {
    expect(describeBlock(sampleParagraph)).toBe('paragraph')
  })

  it('supports exhaustive Inline switches via assertNever', () => {
    expect(describeInline(sampleText)).toBe('text')
  })

  it('preserves branded scalar constructors at runtime and compile time', () => {
    const width = twip(720)
    const fontSize = halfPoint(28)
    const percent = pct(5000)

    expect(width).toBe(720)
    expect(fontSize).toBe(28)
    expect(percent).toBe(5000)
    expectTypeOf(width).toMatchTypeOf<ReturnType<typeof twip>>()
    expectTypeOf(fontSize).toMatchTypeOf<ReturnType<typeof halfPoint>>()
    expectTypeOf(percent).toMatchTypeOf<ReturnType<typeof pct>>()
  })

  it('accepts a frozen sample Document literal', () => {
    expect(Object.isFrozen(sampleDocument)).toBe(true)
    expect(sampleDocument.sections).toHaveLength(1)
    expectTypeOf(sampleDocument).toMatchTypeOf<Document>()
  })

  it('keeps spread-merged RunProps assignable to RunProps', () => {
    const defaults: RunProps = { italic: true, color: 'auto', sz: halfPoint(22) }
    const direct: RunProps = { bold: true, spacing: twip(20) }
    const merged = mergeRunProps(defaults, direct)

    expect(merged).toEqual({
      italic: true,
      color: 'auto',
      sz: halfPoint(22),
      bold: true,
      spacing: twip(20),
    })
    expectTypeOf(merged).toMatchTypeOf<RunProps>()
  })
})

// ---------------------------------------------------------------------------
// DOCX-14 — hexColor() is ST_HexColor's single validation/normalization
// choke point (parser and editor/UI call sites alike). See its doc comment
// in ../styles.ts for the full reasoning behind each of these cases.
// ---------------------------------------------------------------------------
describe('hexColor', () => {
  it('strips a leading "#" from a browser <input type="color"> value', () => {
    expect(hexColor('#ff0000')).toBe('ff0000')
  })

  it('preserves an already-valid six-digit value byte-for-byte, case included', () => {
    // This is what the parser passes in from a file's raw w:val — a color
    // that was already correct must round-trip completely unchanged, not
    // have its case rewritten.
    expect(hexColor('FF0000')).toBe('FF0000')
    expect(hexColor('AaBbCc')).toBe('AaBbCc')
  })

  it('expands a 3-digit hex short form only when explicitly asked to, otherwise treats it as malformed', () => {
    // hexColor() itself has no notion of the CSS 3-digit shorthand — that
    // expansion is parseCssColor's job (colorMapping.ts) before a value
    // ever reaches hexColor(). A bare 3-digit value handed directly to
    // hexColor() is therefore not a valid ST_HexColor and falls back.
    expect(hexColor('#f00')).toBe('000000')
  })

  it('falls back to "000000" for a malformed value rather than throwing', () => {
    // hexColor() is also the parser's own entry point (parseColor in
    // parser/document.ts and parser/styles.ts call it directly on a raw
    // w:val), so it must never throw on a corrupt/hand-edited file — a bad
    // color must not crash the app on open.
    expect(() => hexColor('not-a-color')).not.toThrow()
    expect(hexColor('not-a-color')).toBe('000000')
    expect(hexColor('')).toBe('000000')
    expect(hexColor('#12345')).toBe('000000')
    expect(hexColor('1234567')).toBe('000000')
  })
})
