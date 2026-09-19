#!/usr/bin/env node
// Atlas — CLI front end for `scripts/lib/officeValidator.mjs`.
//
// Validates one or more .docx/.xlsx/.pptx/.odt/.odp/.ods files against the
// actual OPC/OOXML/ODF package specs (not against Atlas's own parsers — see
// the module header on `officeValidator.mjs` for exactly what is checked).
//
// Usage:
//   node scripts/validate-office-file.mjs <file...>
//   node scripts/validate-office-file.mjs electron/templates/*.docx
//
// Exit code is 0 when every file has zero `error`-severity issues, 1
// otherwise (a `warning` alone does not fail the run).
import { readFile } from 'node:fs/promises'

import { formatIssuesReport, validateOfficeFile } from './lib/officeValidator.mjs'

async function main() {
  const targets = process.argv.slice(2)
  if (targets.length === 0) {
    console.error('Usage: node scripts/validate-office-file.mjs <file...>')
    process.exit(2)
  }

  let hadError = false
  for (const target of targets) {
    let buffer
    try {
      buffer = await readFile(target)
    } catch (cause) {
      console.error(`${target}: could not read file (${cause instanceof Error ? cause.message : String(cause)})`)
      hadError = true
      continue
    }

    const result = validateOfficeFile(buffer)
    console.log(formatIssuesReport(target, result))
    console.log()
    if (result.issues.some((i) => i.severity === 'error')) hadError = true
  }

  process.exit(hadError ? 1 : 0)
}

main()
