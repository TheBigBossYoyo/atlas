/**
 * Frozen-pane reader (T4/DAT-10 remainder).
 *
 * SheetJS's bundled `xlsx.js` never parses a worksheet's `<sheetViews>`/
 * `<pane>` element at all — every freeze/split/pane case in its own settings
 * parser is a no-op `break` (confirmed by grepping the bundled build) — so
 * there is no `ws['!freeze']`-equivalent field on the parsed `WorkSheet` to
 * read (see `shared/spreadsheetGrid.ts`'s header). This module reads it
 * directly from the file's own OOXML zip, independently of the SheetJS parse:
 *
 *  1. Unzip the raw file buffer with JSZip.
 *  2. Read `xl/workbook.xml` for the sheet name -> relationship-id order, and
 *     `xl/_rels/workbook.xml.rels` for relationship-id -> worksheet-part
 *     path, so a sheet's *position* in the workbook doesn't have to match its
 *     `sheetN.xml` filename (SheetJS itself doesn't assume that either).
 *  3. For each worksheet part, parse just that file's XML with `DOMParser`
 *     and read `<pane xSplit ySplit state="frozen">` (a plain "split" pane —
 *     a divider the user can drag, not a freeze — is deliberately ignored:
 *     only `state==="frozen"` counts).
 *
 * Only OOXML zip formats (`.xlsx`/`.xlsm`/`.xltx`/`.xltm`) can possibly have
 * this element: `.xls`/`.xlsb` store their sheet views as binary records (no
 * XML to parse) and `.ods`/`.fods` use a completely different settings
 * schema (`config:config-item` entries in `content.xml`/`settings.xml`, not
 * an OOXML `<pane>`) — reading those is a distinct, undocumented-here scope
 * cut. `readFrozenPanes` is safe to call unconditionally for any format: a
 * non-zip buffer fails at `JSZip.loadAsync` and a zip with no matching parts
 * fails at the lookup steps, both caught below to resolve an empty map
 * rather than reject, so callers never need a format check first.
 *
 * Performance (T2/DAT-07): a worksheet's `<sheetViews>` sits near the top of
 * its XML, but `JSZip`'s `.async('string')` has no partial-read mode — it
 * always inflates the *entire* part into one string before this module ever
 * sees it, and `DOMParser.parseFromString` then walks the whole document
 * (including `<sheetData>`, the bulk of a large sheet). For a 100k-row
 * workbook that is exactly the kind of main-thread-blocking work T2 exists
 * to avoid. Callers are responsible for keeping this off the main thread for
 * large files: `spreadsheetWorker.worker.ts` calls this from inside the
 * Worker (which already has the buffer for the parse itself, and Chromium
 * exposes `DOMParser` on the Worker global scope), while
 * `useSpreadsheetWorkbook`'s small-file synchronous branch calls it inline —
 * safe there because that branch only ever runs under
 * `XLSX_WORKER_BYTE_THRESHOLD`.
 */
import JSZip from 'jszip'

import type { FrozenPanes } from '../shared/spreadsheetGrid'

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function getRelationshipId(sheetEl: Element): string | null {
  // Real-world workbook.xml always writes this as the literal `r:id`
  // attribute name (not through a namespace-aware serializer), so a plain
  // `getAttribute('r:id')` matches every producer seen in practice; the
  // namespace-aware lookup is kept as a fallback for a differently-prefixed
  // producer.
  return sheetEl.getAttribute('r:id') ?? sheetEl.getAttributeNS(RELATIONSHIPS_NS, 'id')
}

/** Resolves a `xl/_rels/workbook.xml.rels` `Target` (relative to `xl/`) to its full zip-entry path. */
function resolveRelTarget(target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  return `xl/${target}`
}

export type SheetPart = { readonly name: string; readonly path: string }

/** Sheet name -> worksheet part path, in workbook order (also used by `spreadsheetTables.ts`). */
export function readSheetParts(workbookXml: string, relsXml: string | null): SheetPart[] {
  const parser = new DOMParser()
  const workbookDoc = parser.parseFromString(workbookXml, 'application/xml')
  if (workbookDoc.getElementsByTagName('parsererror').length > 0) return []

  const relTargets = new Map<string, string>()
  if (relsXml) {
    const relsDoc = parser.parseFromString(relsXml, 'application/xml')
    for (const rel of Array.from(relsDoc.getElementsByTagName('Relationship'))) {
      const id = rel.getAttribute('Id')
      const target = rel.getAttribute('Target')
      if (id && target) relTargets.set(id, resolveRelTarget(target))
    }
  }

  const parts: SheetPart[] = []
  for (const sheetEl of Array.from(workbookDoc.getElementsByTagName('sheet'))) {
    const name = sheetEl.getAttribute('name')
    const rId = getRelationshipId(sheetEl)
    if (!name || !rId) continue
    const path = relTargets.get(rId)
    if (path) parts.push({ name, path })
  }
  return parts
}

/** Reads the first `state="frozen"` `<pane>` in one worksheet part's XML, if any. */
function readFrozenPaneFromSheetXml(sheetXml: string): FrozenPanes | null {
  const doc = new DOMParser().parseFromString(sheetXml, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) return null

  for (const pane of Array.from(doc.getElementsByTagName('pane'))) {
    if (pane.getAttribute('state') !== 'frozen') continue
    const cols = Number.parseInt(pane.getAttribute('xSplit') ?? '0', 10)
    const rows = Number.parseInt(pane.getAttribute('ySplit') ?? '0', 10)
    const safeCols = Number.isFinite(cols) && cols > 0 ? cols : 0
    const safeRows = Number.isFinite(rows) && rows > 0 ? rows : 0
    if (safeCols === 0 && safeRows === 0) continue
    return { cols: safeCols, rows: safeRows }
  }
  return null
}

/**
 * Reads every sheet's frozen-pane split from an OOXML workbook buffer.
 * Resolves an empty map — never rejects — for any format/buffer this
 * approach doesn't apply to (see header). Matching is done by the workbook
 * XML's own sheet names (read straight out of the zip), so callers don't
 * need to pass in the names they already got from `parseWorkbookBuffer`.
 */
export async function readFrozenPanes(buffer: ArrayBuffer): Promise<Readonly<Record<string, FrozenPanes>>> {
  try {
    const zip = await JSZip.loadAsync(buffer)
    const workbookEntry = zip.file('xl/workbook.xml')
    if (!workbookEntry) return {}

    const [workbookXml, relsXml] = await Promise.all([
      workbookEntry.async('string'),
      zip.file('xl/_rels/workbook.xml.rels')?.async('string') ?? Promise.resolve(null),
    ])

    const parts = readSheetParts(workbookXml, relsXml)
    const result: Record<string, FrozenPanes> = {}

    await Promise.all(
      parts.map(async ({ name, path }) => {
        const entry = zip.file(path)
        if (!entry) return
        const xml = await entry.async('string')
        const freeze = readFrozenPaneFromSheetXml(xml)
        if (freeze) result[name] = freeze
      }),
    )

    return result
  } catch {
    // Not a zip, not an OOXML workbook, or malformed — every non-applicable
    // format (.xls/.xlsb/.ods/.fods/.csv) and any corrupt file lands here.
    return {}
  }
}
