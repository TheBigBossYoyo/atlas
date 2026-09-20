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

// SEC-1 — the allowlist has two trust tiers: READ (`add`/`has`) and WRITE
// (`trust`/`isWriteEligible`). `path:register-dropped` (drag-drop) can only
// ever grant READ — its provenance can't be verified — while every other
// registration source (open/save dialog, argv, recent-files) grants full
// trust via `trust()`. See pathAllowlist.cjs's header for the full
// rationale.
describe('createPathAllowlist — trust tiers (SEC-1)', () => {
  it('add() makes a path readable but NOT write-eligible', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\dropped\\file.docx')
    expect(allowlist.has('C:\\dropped\\file.docx')).toBe(true)
    expect(allowlist.isWriteEligible('C:\\dropped\\file.docx')).toBe(false)
  })

  it('trust() makes a path both readable and write-eligible', () => {
    const allowlist = createPathAllowlist()
    allowlist.trust('C:\\opened\\file.docx')
    expect(allowlist.has('C:\\opened\\file.docx')).toBe(true)
    expect(allowlist.isWriteEligible('C:\\opened\\file.docx')).toBe(true)
  })

  it('isWriteEligible() is false for a path that was never registered at all', () => {
    const allowlist = createPathAllowlist()
    expect(allowlist.isWriteEligible('C:\\never\\seen.docx')).toBe(false)
  })

  it('trust() after add() upgrades a merely-readable path to write-eligible (drag-drop then Save As)', () => {
    const allowlist = createPathAllowlist()
    allowlist.add('C:\\dropped\\then-saved.docx')
    expect(allowlist.isWriteEligible('C:\\dropped\\then-saved.docx')).toBe(false)
    allowlist.trust('C:\\dropped\\then-saved.docx')
    expect(allowlist.isWriteEligible('C:\\dropped\\then-saved.docx')).toBe(true)
  })

  it('isWriteEligible() normalizes case/slashes/long-path prefix like has()', () => {
    const allowlist = createPathAllowlist()
    allowlist.trust('C:\\Users\\Test\\Doc.DOCX')
    expect(allowlist.isWriteEligible('c:/users/test/doc.docx')).toBe(true)
    expect(allowlist.isWriteEligible('\\\\?\\C:\\Users\\Test\\Doc.DOCX')).toBe(true)
  })

  it('remove() revokes both read and write eligibility', () => {
    const allowlist = createPathAllowlist()
    allowlist.trust('C:\\a\\b.docx')
    allowlist.remove('C:\\a\\b.docx')
    expect(allowlist.has('C:\\a\\b.docx')).toBe(false)
    expect(allowlist.isWriteEligible('C:\\a\\b.docx')).toBe(false)
  })

  it('clear() revokes write eligibility for everything too', () => {
    const allowlist = createPathAllowlist()
    allowlist.trust('C:\\a\\b.docx')
    allowlist.clear()
    expect(allowlist.isWriteEligible('C:\\a\\b.docx')).toBe(false)
  })

  it('ignores unresolvable input when trusting', () => {
    const allowlist = createPathAllowlist()
    expect(allowlist.trust(undefined)).toBeNull()
    expect(allowlist.size()).toBe(0)
  })
})
