/**
 * Re-export shim (Wave 3-P) — the PDF viewer was split into
 * `src/viewers/pdf/*` modules (it had grown past the file-size guideline as
 * one component). Kept at this path so `src/formats/registry.ts`'s existing
 * `import('../viewers/PdfViewer')` doesn't need to change.
 */
export { PdfViewer } from './pdf/PdfViewer'
