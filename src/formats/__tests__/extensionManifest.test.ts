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

  // PROGID-1 — `associationProgId` is written verbatim into
  // `HKLM\SOFTWARE\Classes` by electron-builder. It used to hold the human label,
  // so an install squatted about thirty unqualified generic names there ("Word
  // Document", "Source Code") and the uninstaller then deleted them. These pin the
  // shape so it cannot drift back.
  describe('associationProgId (PROGID-1)', () => {
    it('is namespaced under Atlas., so nothing generic is written to the global class store', () => {
      for (const entry of EXTENSION_MANIFEST) {
        expect(entry.associationProgId, entry.ext).toMatch(/^Atlas\./)
      }
    })

    it('is a valid ProgID: letters, digits and periods only, at most 39 characters', () => {
      // Both limits are Microsoft's, and both are silent when broken — an
      // over-long or punctuated ProgID is simply not honoured by some shell APIs
      // rather than reported. `Atlas.PowerPointMacroEnabledPresentation` would be
      // 40, which is why the ProgIDs say "Macro" where the label says
      // "Macro-Enabled".
      for (const entry of EXTENSION_MANIFEST) {
        expect(entry.associationProgId, entry.ext).toMatch(/^[A-Za-z][A-Za-z0-9.]*$/)
        expect(entry.associationProgId.length, `${entry.ext} -> ${entry.associationProgId}`).toBeLessThanOrEqual(39)
      }
    })

    it('is not the human label — that is associationDescription, which Explorer shows', () => {
      for (const entry of EXTENSION_MANIFEST) {
        expect(entry.associationProgId).not.toBe(entry.associationDescription)
      }
    })

    it('maps one-to-one with the description, so Explorer can still label .dotx unlike .docx', () => {
      // A shared ProgID would mean a shared label: whichever extension the
      // installer wrote last would name all of them.
      const descriptionByProgId = new Map<string, string>()
      for (const entry of EXTENSION_MANIFEST) {
        const seen = descriptionByProgId.get(entry.associationProgId)
        if (seen === undefined) {
          descriptionByProgId.set(entry.associationProgId, entry.associationDescription)
        } else {
          expect(seen, entry.associationProgId).toBe(entry.associationDescription)
        }
      }
      const progIdByDescription = new Map<string, string>()
      for (const entry of EXTENSION_MANIFEST) {
        const seen = progIdByDescription.get(entry.associationDescription)
        if (seen === undefined) {
          progIdByDescription.set(entry.associationDescription, entry.associationProgId)
        } else {
          expect(seen, entry.associationDescription).toBe(entry.associationProgId)
        }
      }
    })
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

  // USR-18 — highlighting moved to CodeMirror, which picks a language from the
  // file name itself; these ids stay as the editor's language label and as the
  // key for its symbol-outline patterns, so what matters is that they are
  // well-formed and consistent, not that a particular library bundles them.
  it('every mapped language is a well-formed, lowercase language id', () => {
    for (const [ext, lang] of Object.entries(CODE_EXTENSION_TO_SHIKI_LANG)) {
      expect(lang, `.${ext} has an empty language id`).toBeTruthy()
      expect(lang, `.${ext} -> "${lang}" should be a lowercase id with no spaces`).toMatch(/^[a-z0-9+#-]+$/)
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
