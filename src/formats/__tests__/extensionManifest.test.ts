/**
 * P2.2/T6 — the canonical extension manifest, and the two artifacts
 * generated from it (electron/main.cjs's known-extensions list and
 * electron-builder.yml's fileAssociations block). Regenerate both with
 * `npm run generate:extensions` if this test fails after editing the
 * manifest.
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import { describe, expect, it } from 'vitest'

import { EXTENSION_MANIFEST, EXTENSION_TO_FORMAT, CODE_EXTENSION_TO_SHIKI_LANG, ALL_MANIFEST_EXTENSIONS } from '../extensionManifest'
import { KNOWN_FORMAT_IDS } from '../types'
import {
  buildExtensionList,
  buildGeneratedCjsSource,
  buildYamlAssociationsBlock,
} from '../../../scripts/lib/extensionManifestGen.mjs'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '../../..')

describe('EXTENSION_MANIFEST', () => {
  it('has no duplicate extensions', () => {
    const extensions = EXTENSION_MANIFEST.map((entry) => entry.ext)
    expect(new Set(extensions).size).toBe(extensions.length)
  })

  it('every entry maps to a known, non-unknown FormatId', () => {
    for (const entry of EXTENSION_MANIFEST) {
      expect(KNOWN_FORMAT_IDS).toContain(entry.format)
    }
  })

  it('only "code" entries carry a shikiLang', () => {
    for (const entry of EXTENSION_MANIFEST) {
      if (entry.format === 'code') {
        expect(entry.shikiLang).toBeTruthy()
      } else {
        expect(entry.shikiLang).toBeUndefined()
      }
    }
  })

  it('every FormatId is reachable from at least one extension', () => {
    for (const format of KNOWN_FORMAT_IDS) {
      expect(EXTENSION_MANIFEST.some((entry) => entry.format === format)).toBe(true)
    }
  })

  it('registers the previously-missing associations (ELEC-05/LOAD-03/LOAD-12)', () => {
    for (const ext of ['mdown', 'ini', 'docm', 'dotx', 'dotm', 'xltx', 'pptm', 'potx']) {
      expect(EXTENSION_TO_FORMAT[ext]).toBeDefined()
    }
  })

  it('only markdown and docx associations are role Editor (ELEC-20)', () => {
    const editorFormats = new Set(
      EXTENSION_MANIFEST.filter((entry) => entry.associationRole === 'Editor').map((entry) => entry.format),
    )
    expect(editorFormats).toEqual(new Set(['markdown', 'docx']))
  })
})

describe('CODE_EXTENSION_TO_SHIKI_LANG (T6/DAT-14)', () => {
  it('covers every code-routed extension', () => {
    const codeExtensions = EXTENSION_MANIFEST.filter((entry) => entry.format === 'code').map((entry) => entry.ext)
    expect(codeExtensions.length).toBeGreaterThanOrEqual(61)
    for (const ext of codeExtensions) {
      expect(CODE_EXTENSION_TO_SHIKI_LANG[ext]).toBeTruthy()
    }
  })

  it('every mapped language is a real shiki bundled language or alias', async () => {
    const shikiLangs = await import('shiki/langs')
    const bundled = (shikiLangs as { bundledLanguages?: Record<string, unknown> }).bundledLanguages ?? shikiLangs
    const knownLangIds = new Set(Object.keys(bundled as Record<string, unknown>))

    for (const [ext, lang] of Object.entries(CODE_EXTENSION_TO_SHIKI_LANG)) {
      expect(knownLangIds.has(lang), `.${ext} -> "${lang}" is not a known shiki language/alias`).toBe(true)
    }
  })
})

describe('generated artifacts stay in sync with the manifest', () => {
  it('electron/lib/extensionManifest.generated.cjs matches the manifest', () => {
    const generatedPath = path.join(repoRoot, 'electron', 'lib', 'extensionManifest.generated.cjs')
    const onDisk = fs.readFileSync(generatedPath, 'utf-8')
    expect(onDisk).toBe(buildGeneratedCjsSource(EXTENSION_MANIFEST))
  })

  it('electron/lib/extensionManifest.generated.cjs EXTENSIONS list matches ALL_MANIFEST_EXTENSIONS', () => {
    expect(buildExtensionList(EXTENSION_MANIFEST)).toEqual([...ALL_MANIFEST_EXTENSIONS])
  })

  it('electron-builder.yml fileAssociations block matches the manifest', () => {
    const ymlPath = path.join(repoRoot, 'electron-builder.yml')
    const ymlText = fs.readFileSync(ymlPath, 'utf-8')
    const expectedBlock = buildYamlAssociationsBlock(EXTENSION_MANIFEST)
    expect(ymlText).toContain(expectedBlock)
  })

  it('no extension appears in only one of the three consuming lists', async () => {
    // detect.ts / extToLang.ts consume EXTENSION_TO_FORMAT directly (same
    // object graph, can't drift). What CAN drift is the generated .cjs list
    // main.cjs reads and the electron-builder.yml block — assert both cover
    // exactly the manifest's extensions, no more, no less.
    const generatedPath = path.join(repoRoot, 'electron', 'lib', 'extensionManifest.generated.cjs')
    const generated = (await import(pathToFileURL(generatedPath).href)) as {
      EXTENSIONS?: ReadonlyArray<string>
      default?: { EXTENSIONS: ReadonlyArray<string> }
    }
    const extensions = generated.EXTENSIONS ?? generated.default?.EXTENSIONS ?? []
    expect([...extensions].sort()).toEqual([...ALL_MANIFEST_EXTENSIONS])

    const ymlPath = path.join(repoRoot, 'electron-builder.yml')
    const ymlText = fs.readFileSync(ymlPath, 'utf-8')
    const ymlExtensions = [...ymlText.matchAll(/^\s*- ext: (\S+)$/gm)].map((m) => m[1])
    expect([...new Set(ymlExtensions)].sort()).toEqual([...ALL_MANIFEST_EXTENSIONS])
  })
})
