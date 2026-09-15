/**
 * PPTX/ODP → PDF export (X1 / SLD-01 / UX-02). One PDF page per slide, each
 * sized to the deck's own aspect ratio — the actual fix for "exporting a
 * 20-slide deck produced a 1-2 page PDF": the old exporter rasterized
 * whatever `SlideDeck`'s virtualized thumbnail rail / single active-slide
 * viewport happened to have mounted, never the other 18+ slides.
 *
 * Re-parses the archive through the exact same `parsePptxSlides`/
 * `parseOdpSlides` (and renders through the exact same `SlideCanvas`)
 * `PptxViewer`/`OdpViewer`/`SlideDeck` use, so export can never drift from
 * what the deck actually looks like on screen — this module never
 * re-implements slide layout, it only drives the existing parser + renderer
 * off-screen via `react-dom/server`.
 */
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';

import { SlideCanvas } from '../../viewers/shared/SlideCanvas';
import type { SlideData } from '../../viewers/shared/SlideDeck.types';
import type { CancelSignal, ZipArchive } from '../../viewers/slides/shared/xmlUtils';
import { parsePptxSlides } from '../../viewers/slides/pptx/parser';
import { parseOdpSlides } from '../../viewers/slides/odp/parser';
import { renderHtmlToPdfFile } from './printDocument';
import { toFriendlyError } from '../friendlyLibraryError';

import viewerSlidesCss from '../../viewers/__styles__/viewer-slides.css?raw';

const PX_PER_INCH = 96;

const SLIDES_PRINT_OVERRIDES = `
:root { --bg-primary: #ffffff; }
* { box-sizing: border-box; }
body { margin: 0; background: #ffffff; }
.export-slide {
  page-break-after: always;
  break-after: page;
  page-break-inside: avoid;
  break-inside: avoid;
}
.export-slide:last-child { page-break-after: auto; break-after: auto; }
.slide-deck__frame { position: relative; overflow: hidden; }
`;

async function loadSlides(buffer: ArrayBuffer, format: 'pptx' | 'odp'): Promise<ReadonlyArray<SlideData>> {
  const { default: JSZip } = await import('jszip');
  const zip = (await JSZip.loadAsync(buffer)) as unknown as ZipArchive;
  const signal: CancelSignal = { cancelled: false };
  const slides = format === 'pptx' ? await parsePptxSlides(zip, signal) : await parseOdpSlides(zip, signal);
  // S14 — a deck's hidden slides are excluded from the default on-screen
  // view (see PptxViewer/OdpViewer); export matches that.
  return slides.filter(slide => !slide.hidden);
}

function renderSlidePage(slide: SlideData): string {
  const markup = renderToStaticMarkup(createElement(SlideCanvas, { slide, scale: 1, interactive: false }));
  return `<div class="export-slide">${markup}</div>`;
}

/** Parses `buffer` (pptx/odp) and exports every visible slide as its own PDF page, sized to the deck's own aspect ratio. */
export async function exportSlidesPdf(buffer: ArrayBuffer, fileName: string, format: 'pptx' | 'odp'): Promise<void> {
  try {
    const slides = await loadSlides(buffer, format);
    if (slides.length === 0) {
      throw new Error('no slides found to export');
    }

    const first = slides[0]!;
    const widthIn = Math.max(first.width, 1) / PX_PER_INCH;
    const heightIn = Math.max(first.height, 1) / PX_PER_INCH;
    const pageCss = `@page { size: ${widthIn}in ${heightIn}in; margin: 0; }`;

    const bodyHtml = slides.map(renderSlidePage).join('\n');
    const css = `${viewerSlidesCss}\n${pageCss}\n${SLIDES_PRINT_OVERRIDES}`;
    await renderHtmlToPdfFile(bodyHtml, css, fileName, fileName);
  } catch (err) {
    console.error('[export] exportSlidesPdf failed:', err);
    throw toFriendlyError(err, 'PDF export failed');
  }
}
