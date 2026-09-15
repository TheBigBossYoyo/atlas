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

// `style` is deliberately NOT in this list — see the `uponSanitizeElement`
// hook below, which forbids it everywhere EXCEPT nested inside an `<svg>`.
const FORBID_TAGS = ['script', 'iframe', 'object', 'embed', 'link', 'base', 'meta'] as const;

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

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * Removes a `<style>` element unless it is nested inside an `<svg>` — a
 * temporary, call-scoped DOMPurify hook (added then removed within a single
 * synchronous `sanitize()` call below, so it can never leak into any other
 * `DOMPurify.sanitize()` call elsewhere in the app, e.g.
 * `sanitizeDocumentHtml.ts`'s RTF/ODT policy, which shares the same
 * underlying DOMPurify singleton).
 *
 * Why: Mermaid's `render()` unconditionally inserts a real `<style>` element
 * as the FIRST CHILD of its output `<svg>` (see `mermaid`'s `render.ts` —
 * `svg.insertBefore(style1, firstChild)`) carrying every color/fill/stroke
 * rule for the diagram's theme; without this exemption a blanket
 * `FORBID_TAGS: ['style']` (the previous behavior) strips that block from
 * EVERY exported Mermaid diagram (both the standalone `.html` export and the
 * printToPDF path), leaving diagrams with none of their theme colors —
 * found during Wave 3 review, since no existing test exercises a real
 * (unmocked) Mermaid SVG. A hand-authored `<svg><style>` from raw HTML in
 * the markdown source (or a converted RTF/ODT file) gets the same exemption,
 * but this adds no meaningfully new risk: this module's `img-src: https:`
 * policy for the standalone HTML export (`markdownHtml.ts`) already lets a
 * plain `<img src="https://...">` act as an equally-effective tracking
 * beacon, and every export's CSP still blocks any style content from
 * executing script.
 */
function forbidNonSvgStyle(node: Node, data: { tagName: string }): void {
  if (data.tagName !== 'style') return;
  // `data.tagName === 'style'` guarantees `node` is an `Element` here, but
  // DOMPurify's hook signature types it as the broader `Node` (which has no
  // `namespaceURI` of its own — that's an `Element`-only property).
  if ((node as Element).namespaceURI === SVG_NS) return;
  node.parentNode?.removeChild(node);
}

/**
 * Sanitizes a serialized HTML fragment/document for safe embedding in an
 * exported file. Not a full document sanitizer — callers still build their
 * OWN trusted `<html>`/`<head>`/`<style>` wrapper (see `printDocument.ts`)
 * around whatever this returns, so stripping a content-supplied `<style>`
 * here only ever touches file-content-derived markup, never the export's
 * own wrapper CSS.
 */
export function sanitizeExportHtml(html: string): string {
  DOMPurify.addHook('uponSanitizeElement', forbidNonSvgStyle);
  try {
    return DOMPurify.sanitize(html, {
      USE_PROFILES: { html: true, svg: true, svgFilters: true, mathMl: true },
      FORBID_TAGS: [...FORBID_TAGS],
      FORBID_ATTR: [...FORBID_ATTR],
      ADD_TAGS: ['foreignObject', 'semantics', 'annotation'],
      ADD_ATTR: ['encoding'],
    });
  } finally {
    DOMPurify.removeHook('uponSanitizeElement');
  }
}
