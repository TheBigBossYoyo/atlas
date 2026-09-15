// Renders a self-contained, already-sanitized HTML document to a vector PDF
// via Electron's `webContents.printToPDF` (X1 — replaces the html2canvas-pro
// screenshot exporter with real per-format export for every non-markdown,
// non-PDF format; markdown itself also switches to this path for a vector,
// selectable-text PDF instead of a raster one).
//
// Security posture (defense in depth — the renderer has already run the HTML
// through DOMPurify before it ever reaches here, see
// `src/utils/export/sanitizeExportHtml.ts`):
//   - a brand-new, hidden `BrowserWindow` with `javascript: false` — the
//     loaded document cannot execute script at all, regardless of what a
//     sanitizer might have missed;
//   - no `preload` script — nothing to bridge back into main;
//   - the document is written to a private temp file and loaded via
//     `loadFile` (not a `data:` URL, which has practical size limits and
//     would need base64-inflating the whole payload) and removed afterward;
//   - the default session's existing CSP `onHeadersReceived` hook (see
//     `applyContentSecurityPolicy` in `main.cjs`) already covers any window
//     on the default session, and the HTML itself additionally carries its
//     own restrictive `<meta http-equiv="Content-Security-Policy">` tag.
//
// `preferCSSPageSize: true` is passed unconditionally — every caller embeds
// its own `@page { size: ...; margin: ... }` rule (DOCX pages by their real
// geometry, slides by deck aspect ratio, tables/text by a standard sheet
// size), so page size is always controlled by the HTML, never by a
// `pageSize` option here. Chromium's print pipeline only supports one page
// box per document even with CSS Paged Media's named-page syntax in
// practice, so a DOCX with more than one distinct section page size still
// prints every page at the FIRST page's size — a known limitation, not a bug
// (documented for the merger).

const { BrowserWindow } = require('electron');
const fs = require('fs');
const fsp = fs.promises;
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const MAX_HTML_BYTES = 40 * 1024 * 1024; // 40 MiB of markup — generous for even a very long export
const LOAD_TIMEOUT_MS = 30_000;

class PrintToPdfError extends Error {
  constructor(message) {
    super(message);
    this.name = 'PrintToPdfError';
  }
}

/**
 * @param {string} html
 * @returns {Promise<string>} the temp file path
 */
async function writeTempHtml(html) {
  const dir = os.tmpdir();
  const file = path.join(dir, `atlas-export-${crypto.randomBytes(8).toString('hex')}.html`);
  await fsp.writeFile(file, html, 'utf-8');
  return file;
}

/**
 * @param {string} tempPath
 */
async function cleanupTempFile(tempPath) {
  try {
    await fsp.unlink(tempPath);
  } catch {
    // Best-effort cleanup only.
  }
}

/**
 * @param {BrowserWindow} win
 * @param {string} fileUrl
 * @returns {Promise<void>}
 */
function waitForLoad(win) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      cleanup();
      reject(new PrintToPdfError('Rendering the document for PDF export timed out.'));
    }, LOAD_TIMEOUT_MS);

    const onFinish = () => {
      cleanup();
      resolve();
    };
    const onFail = (_event, errorCode, errorDescription) => {
      cleanup();
      reject(new PrintToPdfError(`Failed to render the export document (${errorDescription || errorCode}).`));
    };
    const cleanup = () => {
      clearTimeout(timer);
      win.webContents.removeListener('did-finish-load', onFinish);
      win.webContents.removeListener('did-fail-load', onFail);
    };

    win.webContents.once('did-finish-load', onFinish);
    win.webContents.once('did-fail-load', onFail);
  });
}

/**
 * Renders `html` off-screen and returns the resulting PDF bytes.
 * @param {string} html
 * @returns {Promise<Buffer>}
 */
async function printHtmlToPdfBuffer(html) {
  if (typeof html !== 'string' || html.length === 0) {
    throw new PrintToPdfError('Nothing to export.');
  }
  if (Buffer.byteLength(html, 'utf-8') > MAX_HTML_BYTES) {
    throw new PrintToPdfError('This document is too large to export to PDF.');
  }

  const tempPath = await writeTempHtml(html);
  /** @type {BrowserWindow | null} */
  let win = null;

  try {
    win = new BrowserWindow({
      show: false,
      webPreferences: {
        javascript: false,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
        images: true,
      },
    });

    const loadPromise = waitForLoad(win);
    await win.loadFile(tempPath);
    await loadPromise;

    const pdfBuffer = await win.webContents.printToPDF({
      printBackground: true,
      preferCSSPageSize: true,
      landscape: false,
    });

    return pdfBuffer;
  } finally {
    if (win && !win.isDestroyed()) {
      win.destroy();
    }
    await cleanupTempFile(tempPath);
  }
}

module.exports = { printHtmlToPdfBuffer, PrintToPdfError, MAX_HTML_BYTES };
