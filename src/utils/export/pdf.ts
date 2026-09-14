/**
 * Live-DOM → PDF export (used for every format via a full-panel screenshot;
 * a real per-format export is Task X1, out of this task's scope).
 *
 * UX-11 — now routes the finished PDF through Electron's save dialog via
 * `saveBinaryOutput` instead of always dropping into the OS Downloads folder.
 */
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas-pro';

import { sanitizeFileName, saveBinaryOutput, type SaveFilter } from './download';

const PDF_MIME = 'application/pdf';
const PDF_FILTERS: readonly SaveFilter[] = [{ name: 'PDF Document', extensions: ['pdf'] }];

// A4 dimensions in mm
const A4_W = 210;
const A4_H = 297;

/**
 * Captures a live DOM element by id and exports it as a multi-page A4 PDF.
 * Uses html2canvas (scale 2) + jsPDF.
 */
export async function exportPdf(elementId: string, fileName: string): Promise<void> {
  try {
    const name = sanitizeFileName(fileName, 'pdf');
    const target = document.getElementById(elementId);
    if (!target) {
      throw new Error(`[export] exportPdf: element #${elementId} not found`);
    }

    const bgColor = getComputedStyle(target).backgroundColor;

    const canvas = await html2canvas(target, {
      scale: 2,
      useCORS: true,
      backgroundColor: bgColor || '#ffffff',
      logging: false,
    });

    const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const canvasWidthMm = A4_W;
    const canvasHeightMm = (canvas.height / canvas.width) * canvasWidthMm;

    let remainingHeight = canvasHeightMm;
    let yOffset = 0;

    while (remainingHeight > 0) {
      if (yOffset > 0) {
        pdf.addPage();
      }

      const sliceHeightMm = Math.min(A4_H, remainingHeight);
      const sliceRatio = sliceHeightMm / canvasHeightMm;
      const sliceHeightPx = Math.round(canvas.height * sliceRatio);
      const yOffsetPx = Math.round((yOffset / canvasHeightMm) * canvas.height);

      // Slice the canvas for this page
      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvas.width;
      pageCanvas.height = sliceHeightPx;
      const ctx = pageCanvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(canvas, 0, -yOffsetPx);
      }

      const pageImgData = pageCanvas.toDataURL('image/png');
      pdf.addImage(pageImgData, 'PNG', 0, 0, canvasWidthMm, sliceHeightMm);

      yOffset += sliceHeightMm;
      remainingHeight -= sliceHeightMm;
    }

    const bytes = new Uint8Array(pdf.output('arraybuffer'));
    await saveBinaryOutput(bytes, name, PDF_MIME, PDF_FILTERS);
  } catch (err) {
    console.error('[export] exportPdf failed:', err);
    throw err;
  }
}
