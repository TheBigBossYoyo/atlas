import { describe, expect, it } from 'vitest'

import { createPathAllowlist, normalizePath } from '../../../electron/lib/pathAllowlist.cjs'

describe('normalizePath', () => {
  it('returns null for non-string input', () => {
    expect(normalizePath(undefined)).toBeNull()
    expect(normalizePath(42)).toBeNull()
    expect(normalizePath('')).toBeNull()
  })

  it('normalizes case so differently-cased paths compare equal', () => {
    const a = normalizePath('C:\\Users\\Test\\Doc.docx')
    const b = normalizePath('c:\\users\\test\\doc.docx')
    expect(a).toBe(b)
  })

  it('normalizes forward slashes to match backslash paths', () => {
    const a = normalizePath('C:/Users/Test/doc.docx')
    const b = normalizePath('C:\\Users\\Test\\doc.docx')
    expect(a).toBe(b)
  })

  it('strips the \\\\?\\ long-path prefix', () => {
    const a = normalizePath('\\\\?\\C:\\Users\\Test\\doc.docx')
    const b = normalizePath('C:\\Users\\Test\\doc.docx')
    expect(a).toBe(b)
  })

  it('strips the \\\\?\\UNC\\ long-path prefix for network shares', () => {
    const a = normalizePath('\\\\?\\UNC\\server\\share\\doc.docx')
    const b = normalizePath('\\\\server\\share\\doc.docx')
    expect(a).toBe(b)
  })
})

describe('createPathAllowlist', () => {
  it('reports a path as allowed only after it has been added', () => {
    const allowlist = createPathAllowlist()
    expect(allowlist.has('C:\\a\\b.docx')).toBe(false)
    allowlist.add('C:\\a\\b.docx')
    expect(allowlist.has('C:\\a\\b.docx')).toBe(true)
  })

  it('treats case and slash-direction differences as the same path', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\Users\\Test\\Doc.DOCX')
    expect(allowlist.has('c:/users/test/doc.docx')).toBe(true)
  })

  it('treats a long-path-prefixed path as the same as its plain form', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\Users\\Test\\doc.docx')
    expect(allowlist.has('\\\\?\\C:\\Users\\Test\\doc.docx')).toBe(true)
  })

  it('removes a path so it is no longer allowed', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\a\\b.docx')
    allowlist.remove('C:\\a\\b.docx')
    expect(allowlist.has('C:\\a\\b.docx')).toBe(false)
  })

  it('clear() empties the allowlist', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\a\\b.docx')
    allowlist.add('C:\\c\\d.docx')
    allowlist.clear()
    expect(allowlist.size()).toBe(0)
  })

  it('ignores unresolvable input when adding', () => {
    const allowlist = createPathAllowlist()
    expect(allowlist.add(undefined)).toBeNull()
    expect(allowlist.size()).toBe(0)
  })
})
