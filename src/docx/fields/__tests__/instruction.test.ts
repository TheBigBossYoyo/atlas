import { describe, expect, it } from 'vitest'

import { parseFieldInstruction } from '../instruction'

describe('parseFieldInstruction', () => {
  it('extracts the uppercased keyword', () => {
    expect(parseFieldInstruction('page').keyword).toBe('PAGE')
    expect(parseFieldInstruction('  Date  ').keyword).toBe('DATE')
  })

  it('returns an empty keyword and no arguments/switches for a blank instruction', () => {
    const parsed = parseFieldInstruction('   ')
    expect(parsed.keyword).toBe('')
    expect(parsed.arguments).toEqual([])
    expect(parsed.switches.size).toBe(0)
  })

  it('collects non-switch tokens as arguments, in order', () => {
    const parsed = parseFieldInstruction('SEQ Figure \\* ARABIC')
    expect(parsed.arguments).toEqual(['Figure'])
  })

  it('treats a quoted span as one argument even though it contains spaces', () => {
    const parsed = parseFieldInstruction('DATE \\@ "MMMM d, yyyy"')
    expect(parsed.switches.get('@')).toBe('MMMM d, yyyy')
  })

  it('treats a switch with no following value (or one followed by another switch) as a bare flag (true)', () => {
    const parsed = parseFieldInstruction('TOC \\h \\z \\u')
    expect(parsed.switches.get('h')).toBe(true)
    expect(parsed.switches.get('z')).toBe(true)
    expect(parsed.switches.get('u')).toBe(true)
  })

  it('parses REF with its bookmark-name argument and \\h switch', () => {
    const parsed = parseFieldInstruction('REF _Ref123456789 \\h')
    expect(parsed.keyword).toBe('REF')
    expect(parsed.arguments).toEqual(['_Ref123456789'])
    expect(parsed.switches.get('h')).toBe(true)
  })

  it('parses the \\* general-formatting switch like any other switch', () => {
    const parsed = parseFieldInstruction('PAGE \\* MERGEFORMAT')
    expect(parsed.switches.get('*')).toBe('MERGEFORMAT')
  })

  it('parses multiple switches with mixed value/flag shapes', () => {
    const parsed = parseFieldInstruction('TOC \\o "1-3" \\h \\z \\u')
    expect(parsed.switches.get('o')).toBe('1-3')
    expect(parsed.switches.get('h')).toBe(true)
    expect(parsed.switches.get('z')).toBe(true)
    expect(parsed.switches.get('u')).toBe(true)
  })

  it('handles a SEQ instruction with \\r reset and \\c repeat switches', () => {
    const parsed = parseFieldInstruction('SEQ Table \\r 1')
    expect(parsed.arguments).toEqual(['Table'])
    expect(parsed.switches.get('r')).toBe('1')
  })

  it('tolerates an unterminated quote by taking the rest of the string as one token', () => {
    const parsed = parseFieldInstruction('DATE \\@ "MMMM d')
    expect(parsed.switches.get('@')).toBe('MMMM d')
  })
})
