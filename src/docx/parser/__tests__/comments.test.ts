/**
 * Tests for src/docx/parser/comments.ts
 */

import { describe, it, expect } from 'vitest'
import { parseComments } from '../comments'
import { DocxParseError } from '../unzip'

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const FULL_COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Alice Smith" w:date="2024-01-15T10:30:00Z" w:initials="AS">
    <w:p><w:r><w:t>This needs revision.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Bob Jones" w:date="2024-01-16T09:00:00Z" w:initials="BJ">
    <w:p><w:r><w:t>Agreed with Alice.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="3" w:author="Bob Jones" w15:parentId="1" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
    <w:p><w:r><w:t>Reply text.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`

const SINGLE_COMMENT_NO_INITIALS = `<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="3" w:author="Carol" w:date="2024-02-01T08:00:00Z">
    <w:p><w:r><w:t>A comment.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`

const EMPTY_COMMENT = `<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="4" w:author="Dave">
  </w:comment>
</w:comments>`

const EMPTY_COMMENTS = `<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:comments>`

// A comment body shaped exactly like real (compact, non-pretty-printed) Word
// output: a run, then a hyperlink, then another run, with significant
// leading/trailing spaces on the `xml:space="preserve"` runs bordering the
// hyperlink. Word writes `word/comments.xml` on a single line with no
// insignificant whitespace between elements, so this fixture intentionally
// has none either — it targets a specific parser fidelity bug (see below)
// rather than the "pretty printed with indentation" style used elsewhere in
// this file for readability.
const HYPERLINK_SANDWICHED_COMMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="5" w:author="Erin"><w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>the guide</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> for details.</w:t></w:r></w:p></w:comment></w:comments>`

const SELF_CLOSING_EMPTY_COMMENT = `<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="6" w:author="Frank"/>
  <w:comment w:id="7" w:author="Grace">
    <w:p><w:r><w:t>After the self-closed one.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('parseComments', () => {
  it('parses multiple comments into a map keyed by id', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.size).toBe(3)
    expect(map.has('1')).toBe(true)
    expect(map.has('2')).toBe(true)
    expect(map.has('3')).toBe(true)
  })

  it('captures w:author correctly', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.get('1')?.author).toBe('Alice Smith')
  })

  it('captures w:date correctly', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.get('1')?.date).toBe('2024-01-15T10:30:00Z')
  })

  it('captures w:initials correctly', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.get('1')?.initials).toBe('AS')
  })

  it('second comment has correct metadata', () => {
    const map = parseComments(FULL_COMMENTS)
    const c = map.get('2')
    expect(c?.author).toBe('Bob Jones')
    expect(c?.initials).toBe('BJ')
  })

  it('comment without initials has initials=undefined', () => {
    const map = parseComments(SINGLE_COMMENT_NO_INITIALS)
    expect(map.get('3')?.initials).toBeUndefined()
  })

  it('parses comment body paragraphs', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.get('1')?.body).toHaveLength(1)
    expect(map.get('1')?.body[0]?.kind).toBe('paragraph')
  })

  it('parses threaded reply parent ids', () => {
    const map = parseComments(FULL_COMMENTS)
    expect(map.get('3')?.parentId).toBe('1')
  })

  it('comment with empty body produces an empty paragraph array', () => {
    const map = parseComments(EMPTY_COMMENT)
    const c = map.get('4')
    expect(c?.kind).toBe('comment')
    expect(c?.body).toEqual([])
  })

  it('returns empty map for empty <w:comments> element', () => {
    const map = parseComments(EMPTY_COMMENTS)
    expect(map.size).toBe(0)
  })

  it('throws DocxParseError on malformed XML', () => {
    expect(() => parseComments('<<< invalid')).toThrow(DocxParseError)
  })

  it('preserves sibling order and significant whitespace around a hyperlink sandwiched between runs', () => {
    // Regression test: comment bodies used to be re-parsed from the lossy,
    // non-order-preserving `xmlParser` output (default trimValues: true,
    // object-keyed rather than preserveOrder). Re-serializing that output
    // reordered same-tag siblings ahead of interleaved different-tag
    // siblings (run, hyperlink, run -> run, run, hyperlink) and trimmed the
    // significant leading/trailing spaces off the `xml:space="preserve"`
    // runs bordering the hyperlink. Comment bodies are now sliced directly
    // out of the original XML text instead, which avoids both problems.
    const map = parseComments(HYPERLINK_SANDWICHED_COMMENT)
    const paragraph = map.get('5')?.body[0]
    expect(paragraph?.kind).toBe('paragraph')
    if (paragraph?.kind !== 'paragraph') {
      throw new Error('expected a paragraph')
    }

    expect(paragraph.children.map((child) => child.kind)).toEqual(['run', 'hyperlink', 'run'])

    const [firstRun, hyperlink, lastRun] = paragraph.children
    if (firstRun?.kind !== 'run' || hyperlink?.kind !== 'hyperlink' || lastRun?.kind !== 'run') {
      throw new Error('expected run, hyperlink, run')
    }

    expect(firstRun.children).toEqual([{ kind: 'text', value: 'See ', preserveSpace: true }])
    expect(lastRun.children).toEqual([{ kind: 'text', value: ' for details.', preserveSpace: true }])

    const hyperlinkRun = hyperlink.children[0]
    expect(hyperlinkRun?.kind).toBe('run')
    if (hyperlinkRun?.kind === 'run') {
      expect(hyperlinkRun.children).toEqual([{ kind: 'text', value: 'the guide' }])
    }
  })

  it('handles a self-closing empty <w:comment/> without corrupting the next comment', () => {
    const map = parseComments(SELF_CLOSING_EMPTY_COMMENT)
    expect(map.get('6')?.body).toEqual([])
    expect(map.get('7')?.body).toHaveLength(1)
    expect(map.get('7')?.body[0]?.kind).toBe('paragraph')
  })
})
