/**
 * The PDF editing engine — everything that writes or reads PDF structure
 * (PDF-EDIT-1/2) — gathered behind one module so the viewer loads it on
 * demand, like pdf-lib itself: opening a PDF to read it costs none of it.
 */

export { applyEditsToPdf, buildEditedPdf } from './writePdf'
export { createTextAnalysis } from './content/textAnalysis'
