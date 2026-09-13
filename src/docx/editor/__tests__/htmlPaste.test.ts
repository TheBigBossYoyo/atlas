import { describe, expect, it } from 'vitest'

import {
  buildPasteCommands,
  htmlToParagraphs,
  textToParagraphs,
} from '../htmlPaste'
import type { Position, Range } from '../commandTypes'

const cursorAtParagraph = (paragraphIndex: number, runIndex = 0, charOffset = 0): Position => ({
  paragraphPath: Object.freeze([0, paragraphIndex]),
  runIndex,
  charOffset,
})

describe('textToParagraphs', () => {
  it('returns an empty array for the empty string', () => {
    expect(textToParagraphs('')).toEqual([])
  })

  it('splits on \\n and drops a trailing empty paragraph', () => {
    const paragraphs = textToParagraphs('alpha\nbeta\n')
    expect(paragraphs).toHaveLength(2)
    expect(paragraphs[0].runs).toEqual([{ text: 'alpha', format: {} }])
    expect(paragraphs[1].runs).toEqual([{ text: 'beta', format: {} }])
  })

  it('preserves blank paragraphs between content lines', () => {
    const paragraphs = textToParagraphs('alpha\n\nbeta')
    expect(paragraphs).toHaveLength(3)
    expect(paragraphs[1].runs).toEqual([])
  })

  it('handles \\r\\n line endings', () => {
    const paragraphs = textToParagraphs('a\r\nb')
    expect(paragraphs.map(p => p.runs[0]?.text ?? '')).toEqual(['a', 'b'])
  })
})

describe('htmlToParagraphs', () => {
  it('returns an empty array for the empty string', () => {
    expect(htmlToParagraphs('')).toEqual([])
  })

  it('parses two <p> blocks into two paragraphs', () => {
    const paragraphs = htmlToParagraphs('<p>Hello</p><p>World</p>')
    expect(paragraphs).toHaveLength(2)
    expect(paragraphs[0].runs).toEqual([{ text: 'Hello', format: {} }])
    expect(paragraphs[1].runs).toEqual([{ text: 'World', format: {} }])
  })

  it('captures bold/italic/underline formatting', () => {
    const paragraphs = htmlToParagraphs('<p>plain <b>bold</b> <i>italic</i> <u>under</u></p>')
    expect(paragraphs).toHaveLength(1)
    const runs = paragraphs[0].runs
    expect(runs.find(r => r.text.trim() === 'bold')?.format.bold).toBe(true)
    expect(runs.find(r => r.text.trim() === 'italic')?.format.italic).toBe(true)
    expect(runs.find(r => r.text.trim() === 'under')?.format.underline).toEqual({ style: 'single' })
  })

  it('treats <br> as a paragraph break inside a block', () => {
    const paragraphs = htmlToParagraphs('<p>line one<br>line two</p>')
    expect(paragraphs).toHaveLength(2)
    expect(paragraphs[0].runs[0]?.text).toBe('line one')
    expect(paragraphs[1].runs[0]?.text).toBe('line two')
  })

  it('skips <script> and <style> content', () => {
    const paragraphs = htmlToParagraphs('<style>.x{color:red}</style><p>visible</p><script>alert(1)</script>')
    expect(paragraphs).toHaveLength(1)
    expect(paragraphs[0].runs[0]?.text).toBe('visible')
  })

  it('collapses consecutive whitespace in text content', () => {
    const paragraphs = htmlToParagraphs('<p>foo   \n   bar</p>')
    expect(paragraphs[0].runs[0]?.text).toBe('foo bar')
  })
})

describe('buildPasteCommands', () => {
  it('returns no commands when paragraphs is empty', () => {
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands([], cursor, null)
    expect(result.commands).toEqual([])
    expect(result.finalCursor).toEqual(cursor)
  })

  it('emits a single insert-text for one-paragraph plain paste', () => {
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands(textToParagraphs('hello'), cursor, null)
    expect(result.commands).toHaveLength(1)
    expect(result.commands[0]).toMatchObject({
      kind: 'insert-text',
      text: 'hello',
      at: cursor,
    })
    expect(result.finalCursor.charOffset).toBe(5)
  })

  it('inserts paragraph breaks between paragraphs and advances the path', () => {
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands(textToParagraphs('a\nb'), cursor, null)
    expect(result.commands.map(c => c.kind)).toEqual([
      'insert-text',
      'insert-paragraph-break',
      'insert-text',
    ])
    expect(result.finalCursor.paragraphPath[1]).toBe(1)
    expect(result.finalCursor.charOffset).toBe(1)
  })

  it('prefixes a delete-range command when a non-collapsed selection exists', () => {
    const cursor = cursorAtParagraph(0, 0, 2)
    const selection: Range = {
      anchor: cursorAtParagraph(0, 0, 2),
      focus: cursorAtParagraph(0, 0, 5),
    }
    const result = buildPasteCommands(textToParagraphs('X'), cursor, selection)
    expect(result.commands[0]).toMatchObject({ kind: 'delete-range', range: selection })
    expect(result.commands[1]).toMatchObject({ kind: 'insert-text', text: 'X' })
  })

  it('emits apply-run-format for formatted runs', () => {
    const paragraphs = htmlToParagraphs('<p><b>bold</b></p>')
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands(paragraphs, cursor, null)
    const formatCmd = result.commands.find(c => c.kind === 'apply-run-format')
    expect(formatCmd).toBeDefined()
    expect(formatCmd && formatCmd.kind === 'apply-run-format' && formatCmd.format.bold).toBe(true)
  })

  it('does not emit format commands for plain text runs', () => {
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands(textToParagraphs('plain'), cursor, null)
    expect(result.commands.some(c => c.kind === 'apply-run-format')).toBe(false)
  })

  it('skips empty runs in a paragraph without inserting empty text', () => {
    const cursor = cursorAtParagraph(0)
    const result = buildPasteCommands([{ runs: [] }, { runs: [{ text: 'x', format: {} }] }], cursor, null)
    expect(result.commands.map(c => c.kind)).toEqual(['insert-paragraph-break', 'insert-text'])
  })
})
