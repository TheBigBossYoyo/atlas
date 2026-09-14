#!/usr/bin/env node
// Regenerates the two artifacts derived from the canonical extension
// manifest (P2.2 / ELEC-05, ELEC-15, LOAD-03, LOAD-12):
//
//   1. electron/lib/extensionManifest.generated.cjs — the flat extension
//      list electron/main.cjs uses for its known-extensions set (argv/
//      second-instance file detection) and its Open dialog's "Supported
//      Files" filter, so those two can never again silently disagree with
//      `detect.ts`.
//   2. electron-builder.yml's `fileAssociations` block (between the
//      generated-associations sentinel comments), so the Windows installer
//      registers exactly the extensions Atlas actually knows how to open.
//
// Imports `src/formats/extensionManifest.ts` DIRECTLY — no transpile step —
// relying on Node 24's native TypeScript support (type-stripping of
// erasable syntax, which is exactly what tsconfig.app.json's
// `erasableSyntaxOnly: true` guarantees this file sticks to). This mirrors
// how every other one-off repo script here runs (`node scripts/*.mjs`),
// just importing a `.ts` data module instead of only `.mjs` ones.
//
// Run directly: `npm run generate:extensions`. Also runs automatically
// before `npm run build` and `npm run electron:build` (see package.json).
import { readFile, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { EXTENSION_MANIFEST } from '../src/formats/extensionManifest.ts'
import {
  buildGeneratedCjsSource,
  buildYamlAssociationsBlock,
  spliceYamlAssociations,
} from './lib/extensionManifestGen.mjs'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..')

async function main() {
  const cjsPath = path.join(repoRoot, 'electron', 'lib', 'extensionManifest.generated.cjs')
  await writeFile(cjsPath, buildGeneratedCjsSource(EXTENSION_MANIFEST), 'utf-8')

  const ymlPath = path.join(repoRoot, 'electron-builder.yml')
  const ymlText = await readFile(ymlPath, 'utf-8')
  const block = buildYamlAssociationsBlock(EXTENSION_MANIFEST)
  await writeFile(ymlPath, spliceYamlAssociations(ymlText, block), 'utf-8')

  console.log(
    `generate-extension-manifest: wrote ${path.relative(repoRoot, cjsPath)} and updated ` +
      `${path.relative(repoRoot, ymlPath)} from ${EXTENSION_MANIFEST.length} manifest entries.`,
  )
}

main().catch((err) => {
  console.error('generate-extension-manifest failed:', err)
  process.exitCode = 1
})
