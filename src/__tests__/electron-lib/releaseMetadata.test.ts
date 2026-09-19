// P5.1/ELEC-09/ELEC-10 — guards the exe/installer metadata electron-builder.yml
// and package.json feed into the packaged .exe, without actually running
// electron-builder (which this project's contributors are asked not to do
// outside the lead's own build). These are plain text/JSON assertions rather
// than a "pure helper" unit test because electron-builder.yml is consumed
// only by the electron-builder CLI itself, never by Atlas's own runtime code
// — this is the same "treat the yml as text" approach
// scripts/lib/extensionManifestGen.mjs already uses to avoid a yaml-parsing
// dependency.
import fs from 'node:fs'
import path from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = path.resolve(__dirname, '../../..')
const builderYml = fs.readFileSync(path.join(repoRoot, 'electron-builder.yml'), 'utf-8')
const packageJson = JSON.parse(fs.readFileSync(path.join(repoRoot, 'package.json'), 'utf-8')) as {
  author?: unknown
  version?: string
}

describe('electron-builder.yml release metadata (P5.1)', () => {
  it('does not disable resedit-based resource editing (ELEC-10)', () => {
    // `signAndEditExecutable: false` disables BOTH signing AND icon/version
    // metadata embedding (confirmed against app-builder-lib's winPackager.js
    // signApp() — see the comment beside win.signExecutable in
    // electron-builder.yml). Any reintroduction of this flag would silently
    // ship an exe with no custom icon or version info again.
    expect(builderYml).not.toMatch(/^\s*signAndEditExecutable:\s*false/m)
  })

  it('skips only code signing, since there is no certificate yet (ELEC-09)', () => {
    expect(builderYml).toMatch(/^\s*signExecutable:\s*false/m)
  })

  it('sets a non-empty productName, appId, and copyright', () => {
    expect(builderYml).toMatch(/^productName:\s*\S+/m)
    expect(builderYml).toMatch(/^appId:\s*\S+/m)
    expect(builderYml).toMatch(/^copyright:\s*\S/m)
  })

  it('points win.icon at the multi-resolution build/icon.ico', () => {
    expect(builderYml).toMatch(/^\s*icon:\s*build\/icon\.ico/m)
    expect(fs.existsSync(path.join(repoRoot, 'build/icon.ico'))).toBe(true)
  })

  it('gives the uninstaller entry a name distinct from the raw productName (nsis.uninstallDisplayName)', () => {
    expect(builderYml).toMatch(/^\s*uninstallDisplayName:\s*\$\{productName\}\s*\$\{version\}/m)
  })
})

describe('package.json release metadata (P5.1)', () => {
  it('declares author as an object with a name, not a bare string', () => {
    // electron-builder's AppInfo.companyName getter does `author.name` — a
    // bare string author (the previous value here) silently resolves to
    // `undefined`, so the exe's CompanyName field never gets set at all.
    expect(packageJson.author).toEqual(expect.objectContaining({ name: expect.any(String) }))
  })

  it('has a non-empty semantic version for FileVersion/ProductVersion', () => {
    expect(packageJson.version).toMatch(/^\d+\.\d+\.\d+/)
  })
})
