import { act, renderHook } from '@testing-library/react'
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
import type { Position } from '../commandTypes'
import { useEditor } from '../useEditor'

// ─── Helpers ──────────────────────────────────────────────────────────────────

function pos(paragraphPath: ReadonlyArray<number>, runIndex: number, charOffset: number): Position {
  return { paragraphPath: Object.freeze([...paragraphPath]), runIndex, charOffset }
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

function createParagraph(text: string): Paragraph {
  const run: Run = Object.freeze({
    kind: 'run',
    children: Object.freeze([Object.freeze({ kind: 'text', value: text })]),
  }) satisfies Run

  return Object.freeze({
    kind: 'paragraph',
    children: Object.freeze([run]),
  }) satisfies Paragraph
}

function getTextContent(doc: Document): string {
  return doc.sections[0].blocks
    .map((b) => {
      const para = b as Paragraph
      return para.children
        .map((c) => {
          const run = c as Run
          return run.children.map((ch) => (ch.kind === 'text' ? ch.value : '')).join('')
        })
        .join('')
    })
    .join('\n')
}

function makeInputEvent(inputType: string, data?: string | null): InputEvent {
  return new InputEvent('beforeinput', { inputType, data: data ?? null })
}

function makeKeyEvent(key: string, opts?: { ctrl?: boolean; shift?: boolean }): KeyboardEvent {
  return new KeyboardEvent('keydown', {
    key,
    ctrlKey: opts?.ctrl ?? false,
    shiftKey: opts?.shift ?? false,
  })
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe('useEditor', () => {
  it('initializes with the provided document', () => {
    const doc = createDocument([createParagraph('hello')])
    const { result } = renderHook(() => useEditor(doc))
    expect(getTextContent(result.current.document)).toBe('hello')
  })

  it('typing a character grows the document', () => {
    const doc = createDocument([createParagraph('hello')])
    const { result } = renderHook(() => useEditor(doc))

    // Set range first
    act(() => {
      result.current.setRange({ anchor: pos([0, 0], 0, 5), focus: pos([0, 0], 0, 5) })
    })

    act(() => {
      result.current.onBeforeInput(makeInputEvent('insertText', '!'))
    })

    expect(getTextContent(result.current.document)).toBe('hello!')
  })

  it('backspace shrinks the document', () => {
    const doc = createDocument([createParagraph('hello')])
    const { result } = renderHook(() => useEditor(doc))

    act(() => {
      result.current.setRange({ anchor: pos([0, 0], 0, 5), focus: pos([0, 0], 0, 5) })
    })

    act(() => {
      result.current.onBeforeInput(makeInputEvent('deleteContentBackward'))
    })

    expect(getTextContent(result.current.document)).toBe('hell')
  })

  it('Ctrl+Z restores the document after typing', () => {
    const doc = createDocument([createParagraph('hello')])
    const { result } = renderHook(() => useEditor(doc))

    act(() => {
      result.current.setRange({ anchor: pos([0, 0], 0, 5), focus: pos([0, 0], 0, 5) })
    })

    act(() => {
      result.current.onBeforeInput(makeInputEvent('insertText', ' world'))
    })

    expect(getTextContent(result.current.document)).toBe('hello world')

    act(() => {
      result.current.onKeyDown(makeKeyEvent('z', { ctrl: true }))
    })

    expect(getTextContent(result.current.document)).toBe('hello')
  })

  it('range is null initially', () => {
    const doc = createDocument([createParagraph('hello')])
    const { result } = renderHook(() => useEditor(doc))
    expect(result.current.range).toBeNull()
  })
})
