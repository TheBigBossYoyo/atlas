/**
 * Shared DOMPurify sanitization policy for converted-document HTML
 * (RtfViewer, OdtViewer).
 *
 * Both viewers render an untrusted third-party file format into live HTML —
 * ODT via odf-kit, RTF via rtf.js — and neither library makes any safety
 * guarantee about the DOM it hands back. Concretely: rtf.js's hyperlink
 * field handling assigns `link.href` directly from the RTF file's own
 * `\fldinst HYPERLINK "..."` string with zero scheme validation (see
 * `Renderer.buildHyperlinkElement` in the `rtf.js` package), so a crafted
 * RTF can render a clickable `javascript:`-URI link that runs arbitrary
 * script in this renderer the moment a reader clicks it (found during a
 * Wave 3-T review — RtfViewer previously appended rtf.js's nodes to the
 * document with no sanitization at all, unlike OdtViewer). Routing both
 * viewers' output through this one policy closes that gap and keeps the
 * two viewers from drifting onto two different sanitization configs.
 *
 * Deliberately synchronous and dependency-injected (the caller resolves
 * `dompurify`'s dynamic `import()` itself, in parallel with its own
 * conversion work — see RtfViewer/OdtViewer) rather than this module doing
 * the dynamic import internally: an `async` version that awaited its own
 * `import('dompurify')` triggered a false-positive
 * `react-hooks/set-state-in-effect` lint error on an *unrelated* early
 * `setState` call elsewhere in the calling effect (the lint rule's static
 * analysis appears to treat awaiting a project-local async helper
 * differently from awaiting an inline `import()`/library call). This form
 * sidesteps that without weakening the guarantee — the caller's own
 * `Promise.all([...])`/`withRenderTimeout` already covers the dompurify
 * import's timing.
 */
import type DOMPurifyType from 'dompurify'

export function sanitizeDocumentHtml(html: string, DOMPurify: typeof DOMPurifyType): string {
  return DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed'],
    FORBID_ATTR: ['onerror', 'onload', 'onclick'],
  })
}
