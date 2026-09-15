/**
 * Renderer-side wrapper around the `export:printToPdf` IPC (X1). Every real
 * per-format PDF exporter in this directory funnels through this one
 * function so there is exactly one place that talks to the main process and
 * one place that turns "not running in Electron" / a main-process failure
 * into a thrown `Error` callers can wrap with `toFriendlyError`.
 */

/** Thrown when PDF export is attempted outside Electron (no `printToPdf` bridge). */
export class PdfExportUnavailableError extends Error {
  constructor() {
    super('PDF export requires the desktop app.');
    this.name = 'PdfExportUnavailableError';
  }
}

/**
 * Sends a complete, already-sanitized HTML document to the main process and
 * returns the rendered PDF's bytes. Throws on any failure — callers are
 * expected to wrap this in `toFriendlyError(err, '<Format> export failed')`.
 */
export async function printHtmlToPdf(html: string): Promise<Uint8Array> {
  const printToPdf = window.electronAPI?.printToPdf;
  if (!printToPdf) {
    throw new PdfExportUnavailableError();
  }

  const result = await printToPdf(html);
  if (!result.ok) {
    throw new Error(result.error);
  }
  return result.bytes;
}
