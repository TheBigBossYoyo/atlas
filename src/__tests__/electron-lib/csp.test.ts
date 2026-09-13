import { describe, expect, it } from 'vitest'

import { buildContentSecurityPolicy } from '../../../electron/lib/csp.cjs'

describe('buildContentSecurityPolicy', () => {
  it('allows unsafe-eval and the Vite dev server/HMR origins in dev', () => {
    const csp = buildContentSecurityPolicy(true)
    expect(csp).toContain("script-src 'self' 'unsafe-inline' 'unsafe-eval'")
    expect(csp).toContain('ws://localhost:5173')
    expect(csp).toContain('http://localhost:5173')
  })

  it('omits unsafe-eval and dev origins in production', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).not.toContain("'unsafe-eval'")
    expect(csp).not.toContain('localhost:5173')
  })

  it('allows wasm instantiation in production for the shiki tokenizer', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).toContain("'wasm-unsafe-eval'")
  })

  it('allows the pdf.js module worker via worker-src', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).toContain("worker-src 'self' blob:")
  })

  it('allows blob:/data: images and fonts for DOCX inline images and bundled fonts', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).toContain("img-src 'self' data: blob:")
    expect(csp).toContain("font-src 'self' data: blob:")
  })

  it('allows inline styles for KaTeX/mermaid', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).toContain("style-src 'self' 'unsafe-inline'")
  })

  it('blocks plugin objects entirely', () => {
    const csp = buildContentSecurityPolicy(false)
    expect(csp).toContain("object-src 'none'")
  })
})
