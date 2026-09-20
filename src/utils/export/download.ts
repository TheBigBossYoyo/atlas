/**
 * Shared save/output helpers used by every `exportX` function.
 *
 * UX-11 — every export now routes through Electron's native save dialog
 * (`window.electronAPI.saveFile` / `saveBinaryFile`, with a suggested file
 * name and format-appropriate filters) when running inside Electron, instead
 * of always dropping into the OS Downloads folder via a synthetic anchor
 * click. The browser-download path remains as the fallback for the
 * plain-browser / dev-server build where `window.electronAPI` doesn't exist.
 */

export interface SaveFilter {
  readonly name: string;
  readonly extensions: readonly string[];
}

/** Strip any existing file extension from `name` and append `.{ext}`. */
export function sanitizeFileName(name: string, ext: string): string {
  const stripped = name.replace(/\.[^./\\]+$/, '');
  return `${stripped}.${ext}`;
}

function toMutableFilters(filters?: readonly SaveFilter[]): Array<{ name: string; extensions: string[] }> | undefined {
  return filters?.map(f => ({ name: f.name, extensions: [...f.extensions] }));
}

function triggerBlobDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  // Revoke after a short delay so the browser has time to start the download.
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/**
 * Saves text `content` as `fileName`. In Electron this opens the native save
 * dialog; a dialog the user cancels (`saved: false`, no `error`) resolves
 * quietly with no download fallback — cancelling isn't a failure. A real
 * failure (`error` present) throws so callers can surface it. Outside
 * Electron, falls back to a browser download.
 *
 * Resolves `true` when the file was actually written (or the browser
 * download was triggered), `false` when the user cancelled the native
 * dialog — a caller driving several dialogs in sequence (e.g. one CSV file
 * per sheet — see `exportSpreadsheetCsvPerSheet`) needs this to stop after a
 * cancel instead of ploughing on to the next dialog.
 */
export async function saveTextOutput(
  content: string,
  fileName: string,
  mimeType: string,
  filters?: readonly SaveFilter[],
): Promise<boolean> {
  const electronAPI = window.electronAPI;
  if (electronAPI?.saveFile) {
    const result = await electronAPI.saveFile({
      content,
      suggestedName: fileName,
      filters: toMutableFilters(filters),
    });
    if (!result.saved && result.error) {
      throw new Error(result.error);
    }
    return result.saved;
  }
  triggerBlobDownload(new Blob([content], { type: mimeType }), fileName);
  return true;
}

/**
 * Binary counterpart of {@link saveTextOutput} — routes through
 * `saveBinaryFile`. See {@link saveTextOutput} for the meaning of the
 * resolved boolean.
 */
export async function saveBinaryOutput(
  content: Uint8Array | ArrayBuffer,
  fileName: string,
  mimeType: string,
  filters?: readonly SaveFilter[],
): Promise<boolean> {
  // Always copy into a fresh, plain-`ArrayBuffer`-backed view: `content` may
  // already be a `Uint8Array` typed over a general `ArrayBufferLike` (which
  // also admits `SharedArrayBuffer`), and `Blob`'s `BlobPart` type only
  // accepts a view backed by a concrete `ArrayBuffer`.
  const bytes: Uint8Array<ArrayBuffer> = new Uint8Array(content);
  const electronAPI = window.electronAPI;
  if (electronAPI?.saveBinaryFile) {
    const result = await electronAPI.saveBinaryFile({
      content: bytes,
      suggestedName: fileName,
      filters: toMutableFilters(filters),
    });
    if (!result.saved && result.error) {
      throw new Error(result.error);
    }
    return result.saved;
  }
  triggerBlobDownload(new Blob([bytes], { type: mimeType }), fileName);
  return true;
}
