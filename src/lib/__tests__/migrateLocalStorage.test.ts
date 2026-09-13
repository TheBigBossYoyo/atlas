import { describe, it, expect, beforeEach } from 'vitest'
import { runMigration } from '../migrateLocalStorage'

beforeEach(() => {
  localStorage.clear()
})

describe('runMigration', () => {
  it('Case A: migrates md-reader-* keys to atlas-* and sets flag', () => {
    localStorage.setItem('md-reader-theme', 'dark')
    localStorage.setItem('md-reader-draft', '{"markdown":"hello","fileName":null,"savedAt":1}')
    localStorage.setItem('md-reader-font-size', '18')

    runMigration()

    expect(localStorage.getItem('atlas-theme')).toBe('dark')
    expect(localStorage.getItem('atlas-draft')).toBe('{"markdown":"hello","fileName":null,"savedAt":1}')
    expect(localStorage.getItem('atlas-font-size')).toBe('18')
    expect(localStorage.getItem('md-reader-theme')).toBeNull()
    expect(localStorage.getItem('md-reader-draft')).toBeNull()
    expect(localStorage.getItem('md-reader-font-size')).toBeNull()
    expect(localStorage.getItem('atlas-migration-v1')).toBe('1')
  })

  it('Case B: idempotent — second call is a no-op when flag is set', () => {
    localStorage.setItem('atlas-theme', 'sepia')
    runMigration() // sets flag, no md-reader-* keys
    localStorage.setItem('md-reader-theme', 'dark') // add after first run
    runMigration() // should be no-op due to flag

    expect(localStorage.getItem('atlas-theme')).toBe('sepia')
    expect(localStorage.getItem('md-reader-theme')).toBe('dark') // not migrated
  })

  it('Case C: does not clobber existing atlas-* key', () => {
    localStorage.setItem('md-reader-theme', 'dark')
    localStorage.setItem('atlas-theme', 'light')

    runMigration()

    expect(localStorage.getItem('atlas-theme')).toBe('light')
    expect(localStorage.getItem('md-reader-theme')).toBeNull()
  })

  it('Case D: empty storage — sets flag, no errors', () => {
    expect(() => runMigration()).not.toThrow()
    expect(localStorage.getItem('atlas-migration-v1')).toBe('1')
  })
})
