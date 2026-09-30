/**
 * OFFICE-VERIFY-1 — shared LibreOffice plumbing for the round-trip harnesses.
 *
 * Not a test file (Vitest collects only `*.test.ts`), so it can hold the process
 * handling and text extraction that `docx/__tests__/libreOfficeRoundTrip.test.ts`
 * and its spreadsheet/slide counterparts all need.
 *
 * WHY THESE HARNESSES EXIST. Every round-trip guarantee in this repo used to be
 * self-asserted: `scripts/validate-office-file.mjs` was the only outside opinion,
 * and it passed two genuinely invalid files clean before attribute datatype checks
 * were added. Asking a real office suite what it reads from a file Atlas wrote is a
 * different kind of evidence, and it has already found a fidelity bug
 * (`FID-DEFAULTS-1`) that nothing in Atlas's own model could see.
 *
 * FOUR THINGS LEARNED THE HARD WAY, all encoded below:
 *   1. **One process per file, with a timeout and a cleanup.** LibreOffice 26.8.0's
 *      `txt:Text` filter runs away on any document containing a table — 620 MB of
 *      temp output in 120 s, 4.6 GB if left. A single batch invocation for many
 *      files gives such a bug nothing to stop it, and it filled the disk.
 *   2. **A profile of our own.** Without `-env:UserInstallation`, a headless run
 *      shares a profile with any LibreOffice window the developer has open, and
 *      exits having done nothing — indistinguishable from a conversion that
 *      produced no output.
 *   3. **Newlines in the output are not content.** LibreOffice wraps its own HTML,
 *      so a span can come back split across lines. Preserving those newlines made
 *      an earlier version report a difference that did not exist. Block boundaries
 *      are marked with a sentinel FIRST, then all real whitespace collapses, then
 *      the sentinel becomes a newline.
 *   4. **`<head>` is not content either.** It carries a generator `<meta>` that
 *      genuinely differs between a fixture and an Atlas save.
 */
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

/** Where `scoop install extras/libreoffice` puts it, then the usual system paths. */
export function findSoffice(): string | null {
  const candidates = [
    path.join(os.homedir(), 'scoop', 'apps', 'libreoffice', 'current', 'LibreOffice', 'program', 'soffice.exe'),
    'C:/Program Files/LibreOffice/program/soffice.exe',
    'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
    '/usr/bin/soffice',
    '/usr/local/bin/soffice',
  ]
  return candidates.find((candidate) => fs.existsSync(candidate)) ?? null
}

/**
 * Filter specs, by what they are for. A wrong filter is not a soft failure: the
 * Writer HTML filter simply refuses a spreadsheet (exit 1, no output).
 */
export const CONVERT = {
  /** Writer -> HTML. Not `txt:Text`, which hangs on tables — see this file's header. */
  writerHtml: { spec: 'html:HTML (StarWriter)', ext: 'html' },
  /** Calc -> XHTML. Covers EVERY sheet, unlike the CSV filter, which exports only the first. */
  calcHtml: { spec: 'html:XHTML Calc File', ext: 'html' },
  /** Impress -> flat ODP. One XML file, so the slide text needs no multi-file wrangling. */
  impressFlatOdp: { spec: 'fodp:OpenDocument Presentation Flat XML', ext: 'fodp' },
} as const

/** Generous for documents of fixture size; short enough to catch a runaway early. */
const CONVERT_TIMEOUT_MS = 90_000

function cleanUpAfterTimeout(outDir: string): void {
  try {
    execFileSync('taskkill', ['/F', '/IM', 'soffice.bin', '/T'], { stdio: 'ignore' })
  } catch {
    // Not running, or not Windows. Either is fine.
  }
  for (const name of fs.readdirSync(outDir)) {
    // A runaway writes its output to a `.tmp` beside the real target.
    if (name.endsWith('.tmp')) fs.rmSync(path.join(outDir, name), { force: true })
  }
}

/**
 * Converts one file and returns the output's path, or `null` when LibreOffice could
 * not read it — no output produced, or the conversion timed out.
 */
export function convert(
  soffice: string,
  file: string,
  filter: { readonly spec: string; readonly ext: string },
  workDir: string,
): string | null {
  const outDir = path.join(workDir, 'converted')
  fs.mkdirSync(outDir, { recursive: true })
  // A URL: the only form `-env:UserInstallation` accepts.
  const profileUrl = `file:///${path.join(workDir, 'profile').replace(/\\/g, '/')}`

  try {
    execFileSync(
      soffice,
      [
        '--headless',
        '--norestore',
        '--nolockcheck',
        `-env:UserInstallation=${profileUrl}`,
        '--convert-to',
        filter.spec,
        '--outdir',
        outDir,
        file,
      ],
      { stdio: 'pipe', timeout: CONVERT_TIMEOUT_MS },
    )
  } catch {
    cleanUpAfterTimeout(outDir)
    return null
  }

  const produced = path.join(outDir, `${path.basename(file, path.extname(file))}.${filter.ext}`)
  // An empty document produces an empty file, not no file, so existence is the
  // right test for "could LibreOffice read this".
  return fs.existsSync(produced) ? produced : null
}

/** Stands in for a block boundary while all real whitespace is collapsed around it. */
const BLOCK_BREAK = '\u0000'

/** See point 3 in this file's header for why the order of operations matters. */
function collapse(text: string): string {
  return text
    .replace(/\s+/g, ' ')
    .split(BLOCK_BREAK)
    .map((line) => line.trim())
    .filter((line) => line !== '')
    .join('\n')
}

function decodeEntities(text: string): string {
  return (
    text
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)))
      // Last, so an escaped entity in the source is not double-decoded.
      .replace(/&amp;/g, '&')
  )
}

/**
 * The visible text of LibreOffice's HTML output, one block per line. Works for both
 * the Writer and Calc HTML filters — a spreadsheet comes back as `<table>` markup,
 * so every cell becomes its own line.
 */
export function htmlToText(html: string): string {
  const bodyStart = html.search(/<body\b[^>]*>/i)
  const body = bodyStart === -1 ? html : html.slice(html.indexOf('>', bodyStart) + 1)
  return collapse(
    decodeEntities(
      body
        .replace(/<\/body>[\s\S]*$/i, '')
        .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, '')
        // Without this, "A" and "B" in adjacent cells concatenate into "AB" and a
        // genuinely moved cell boundary compares equal.
        .replace(/<(?:\/p|\/h[1-6]|br|\/td|\/th|\/tr|\/li|\/div)\b[^>]*>/gi, BLOCK_BREAK)
        .replace(/<[^>]+>/g, ''),
    ),
  )
}

/**
 * The text of a flat ODF file (`.fodp`), one `text:p` per line.
 *
 * Only `office:body` onwards: the styles and master pages above it are full of
 * generated names that differ harmlessly between two files with identical content.
 */
export function flatOdfText(xml: string): string {
  const bodyStart = xml.indexOf('<office:body')
  const body = bodyStart === -1 ? xml : xml.slice(bodyStart)
  const paragraphs = body.match(/<text:p\b[^>]*>[\s\S]*?<\/text:p>|<text:p\b[^>]*\/>/g) ?? []
  return collapse(paragraphs.map((p) => decodeEntities(p.replace(/<[^>]+>/g, '')) + BLOCK_BREAK).join(''))
}
