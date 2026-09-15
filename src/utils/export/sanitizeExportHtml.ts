/**
 * DOMPurify policy for HTML that is about to leave the app as an export —
 * either a standalone `.html` file (`markdownHtml.ts`) or the body markup
 * embedded in a print document sent to `export:printToPdf` (`pdf.ts` and
 * friends).
 *
 * Why this exists (X5 — exported HTML sanitization): `MarkdownRenderer`
 * renders raw HTML embedded in the markdown source via `rehype-raw` with no
 * sanitization step (by design — see its own header comment; the owner's
 * rendering behavior is not being changed here). React never attaches a
 * literal `onerror`/`onclick` *attribute* for a raw-HTML node it renders
 * (it expects a function prop, not a string, so it just warns and drops
 * it) and never executes a `<script>` element it creates via
 * `document.createElement`/`appendChild` — so those vectors are inert in the
 * *live* preview. But `exportHtml`/the print-document builders below grab
 * `element.outerHTML` — a plain string — and hand it to a brand-new document
 * (a `.html` file opened directly in a browser, or the hidden print window)
 * whose OWN HTML parser treats a `<script>` tag as parser-inserted (which
 * DOES execute) and a `javascript:`/data: href or an `onload`/`onerror`
 * attribute as live markup. Sanitizing the serialized string before it is
 * ever embedded closes that gap regardless of what produced it (raw HTML in
 * markdown, a hyperlink target from a converted RTF/DOCX/ODT file, etc.).
 *
 * `svg`/`svgFilters`/`mathMl` profiles are enabled (unlike the stricter
 * `sanitizeDocumentHtml` used by RtfViewer/OdtViewer) because Mermaid renders
 * diagrams as inline `<svg>` and KaTeX emits `<math>` MathML alongside its
 * HTML fallback — both must survive for export fidelity. Two additions on
 * top of DOMPurify's defaults:
 *   - `foreignObject`: some Mermaid diagram types (e.g. flowchart labels)
 *     embed HTML *inside* an SVG via `<foreignObject>`, which DOMPurify's
 *     default SVG profile does not allow-list. Note: DOMPurify still strips
 *     the HTML *content* nested inside a `foreignObject` regardless of this
 *     addition — a deliberate hardening against a known mutation-XSS bypass
 *     class that historically abused exactly this SVG/HTML namespace
 *     crossing — so this only preserves the (harmless) `<foreignObject>`
 *     wrapper itself, not arbitrary HTML inside it. Accepted trade-off.
 *   - `semantics`/`annotation` (+ its `encoding` attribute): DOMPurify's
 *     default MathML tag set is "presentation" markup only (`mrow`, `mi`,
 *     `mo`, `mfrac`, ...) and excludes these — KaTeX wraps every formula's
 *     MathML in `<semantics><mrow>...</mrow><annotation
 *     encoding="application/x-tex">...</annotation></semantics>`, and
 *     without this the exported MathML silently lost its screen-reader/
 *     copy-as-LaTeX source annotation on every export. Deliberately NOT
 *     added: `annotation-xml`, the sibling tag whose `encoding="text/html"`
 *     variant was a real historical MathML/HTML namespace-confusion XSS
 *     vector in some browsers — KaTeX never emits it, so leaving it
 *     forbidden costs nothing.
 */
import DOMPurify from 'dompurify';

const FORBID_TAGS = ['script', 'style', 'iframe', 'object', 'embed', 'link', 'base', 'meta'] as const;

const FORBID_ATTR = [
  'onerror',
  'onload',
  'onclick',
  'onmouseover',
  'onmouseout',
  'onfocus',
  'onblur',
  'onchange',
  'onsubmit',
  'onanimationstart',
  'onanimationend',
  'srcdoc',
] as const;

/**
 * Sanitizes a serialized HTML fragment/document for safe embedding in an
 * exported file. Not a full document sanitizer — callers still build their
 * OWN trusted `<html>`/`<head>`/`<style>` wrapper (see `printDocument.ts`)
 * around whatever this returns, so `FORBID_TAGS` dropping `<style>` here
 * only strips a *content-supplied* style block, never the export's own.
 */
export function sanitizeExportHtml(html: string): string {
  return DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true, svg: true, svgFilters: true, mathMl: true },
    FORBID_TAGS: [...FORBID_TAGS],
    FORBID_ATTR: [...FORBID_ATTR],
    ADD_TAGS: ['foreignObject', 'semantics', 'annotation'],
    ADD_ATTR: ['encoding'],
  });
}
