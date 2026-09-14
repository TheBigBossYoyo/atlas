/**
 * Security regression coverage for the shared RTF/ODT sanitization policy.
 *
 * The concrete trigger: rtf.js assigns a rendered hyperlink's `href`
 * directly from the RTF file's own `\fldinst HYPERLINK "..."` string with
 * no scheme validation at all (`Renderer.buildHyperlinkElement`), so a
 * crafted `.rtf` can render a clickable `javascript:`-URI link that runs
 * arbitrary script in this renderer when clicked. Found reviewing this
 * wave; fixed by routing both RtfViewer's and OdtViewer's converted output
 * through this one DOMPurify policy before it ever reaches `innerHTML`.
 */
import DOMPurify from 'dompurify'
import { describe, expect, it } from 'vitest'

import { sanitizeDocumentHtml } from '../sanitizeDocumentHtml'

describe('sanitizeDocumentHtml', () => {
  it('strips a javascript: URI from an anchor href', () => {
    const html = sanitizeDocumentHtml('<a href="javascript:alert(1)">Click me</a>', DOMPurify)

    expect(html).not.toMatch(/javascript:/i)
    expect(html).toContain('Click me')
  })

  it('preserves a safe http(s) hyperlink', () => {
    const html = sanitizeDocumentHtml('<a href="https://example.com/doc">link</a>', DOMPurify)

    expect(html).toContain('href="https://example.com/doc"')
  })

  it('removes <script> tags entirely', () => {
    const html = sanitizeDocumentHtml('<p>hi</p><script>alert(1)</script>', DOMPurify)

    expect(html).not.toContain('<script')
    expect(html).toContain('<p>hi</p>')
  })

  it('strips inline event-handler attributes', () => {
    const html = sanitizeDocumentHtml('<img src="data:image/png;base64,AAAA" onerror="alert(1)">', DOMPurify)

    expect(html).not.toContain('onerror')
  })

  it('preserves a data: image src (legitimate embedded raster content)', () => {
    const html = sanitizeDocumentHtml('<img src="data:image/png;base64,AAAA">', DOMPurify)

    expect(html).toContain('src="data:image/png;base64,AAAA"')
  })

  it('preserves ordinary structural markup (tables, lists, formatting)', () => {
    const html = sanitizeDocumentHtml(
      '<table><tr><td>A</td></tr></table><ul><li>x</li></ul><ins>added</ins><del>removed</del>',
      DOMPurify,
    )

    expect(html).toContain('<table>')
    expect(html).toContain('<ul>')
    expect(html).toContain('<ins>')
    expect(html).toContain('<del>')
  })
})
