import type { FormatId, ViewerComponent } from './types'

export type ViewerLoader = () => Promise<ViewerComponent>

// ---------------------------------------------------------------------------
// Each loader uses a STATIC dynamic import so Vite/Rollup can analyze it at
// build time and emit a real lazy chunk under dist/assets/. Do NOT introduce
// template strings or @vite-ignore here — those break production bundling
// and yield "Failed to fetch dynamically imported module" at runtime.
// ---------------------------------------------------------------------------

const pick = (mod: Record<string, unknown>, name: string): ViewerComponent =>
  mod[name] as ViewerComponent

export const viewerRegistry: Readonly<Record<FormatId, ViewerLoader>> = {
  markdown: () => import('../viewers/MarkdownViewer').then((m) => pick(m, 'MarkdownViewer')),
  docx: () => import('../viewers/DocxViewer').then((m) => pick(m, 'DocxViewer')),
  xlsx: () => import('../viewers/XlsxViewer').then((m) => pick(m, 'XlsxViewer')),
  pptx: () => import('../viewers/PptxViewer').then((m) => pick(m, 'PptxViewer')),
  pdf: () => import('../viewers/PdfViewer').then((m) => pick(m, 'PdfViewer')),
  csv: () => import('../viewers/CsvViewer').then((m) => pick(m, 'CsvViewer')),
  tsv: () => import('../viewers/TsvViewer').then((m) => pick(m, 'TsvViewer')),
  text: () => import('../viewers/TextViewer').then((m) => pick(m, 'TextViewer')),
  code: () => import('../viewers/CodeViewer').then((m) => pick(m, 'CodeViewer')),
  odt: () => import('../viewers/OdtViewer').then((m) => pick(m, 'OdtViewer')),
  ods: () => import('../viewers/OdsViewer').then((m) => pick(m, 'OdsViewer')),
  odp: () => import('../viewers/OdpViewer').then((m) => pick(m, 'OdpViewer')),
  rtf: () => import('../viewers/RtfViewer').then((m) => pick(m, 'RtfViewer')),
  unknown: () => import('../viewers/UnknownViewer').then((m) => pick(m, 'UnknownViewer')),
} as const
