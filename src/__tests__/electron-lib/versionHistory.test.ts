/**
 * VERSIONS-1 — the version-history store.
 *
 * This is the layer where the user's record of their own work either survives or
 * quietly does not, so the tests lean on the awkward cases rather than the happy
 * path: identical saves, a corrupt index, two documents that must not see each
 * other's versions, and the pruning that decides what gets thrown away.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  createVersionHistory,
  INDEX_FILE_NAME,
  MAX_VERSIONS_PER_DOCUMENT,
} from '../../../electron/lib/versionHistory.cjs'

let rootDir: string
let history: ReturnType<typeof createVersionHistory>

const DOC = 'C:/Users/Someone/Documents/thesis.docx'
const OTHER = 'C:/Users/Someone/Documents/letter.docx'

beforeEach(() => {
  rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'atlas-history-'))
  history = createVersionHistory(rootDir)
})

afterEach(() => {
  fs.rmSync(rootDir, { recursive: true, force: true })
})

/**
 * The on-disk directory for a document, found without reimplementing the hash
 * and WITHOUT writing anything: an earlier version of this helper took a probe
 * snapshot to make the directory findable, which quietly added a version to the
 * very history the caller was about to make assertions about.
 */
function dirFor(documentPath: string): string {
  const newest = history.list(documentPath).at(-1)
  if (newest === undefined) throw new Error(`${documentPath} has no versions to locate a directory by`)
  const found = fs.readdirSync(rootDir).find((entry) => {
    const index = path.join(rootDir, entry, INDEX_FILE_NAME)
    return fs.existsSync(index) && fs.readFileSync(index, 'utf8').includes(newest.id)
  })
  if (found === undefined) throw new Error('could not locate the history directory')
  return path.join(rootDir, found)
}

describe('version history store (VERSIONS-1)', () => {
  it('records a version and reads back exactly the bytes it was given', () => {
    const bytes = Buffer.from('the first draft')
    const result = history.snapshot(DOC, bytes)

    expect(result.stored).toBe(true)
    const versions = history.list(DOC)
    expect(versions).toHaveLength(1)
    expect(history.read(DOC, versions[0]!.id)).toEqual(bytes)
  })

  it('keeps versions oldest first, which is the order the UI reads them in', () => {
    history.snapshot(DOC, Buffer.from('one'))
    history.snapshot(DOC, Buffer.from('two'))
    history.snapshot(DOC, Buffer.from('three'))

    const versions = history.list(DOC)
    expect(versions.map((v) => history.read(DOC, v.id)?.toString())).toEqual(['one', 'two', 'three'])
    expect(versions[0]!.at).toBeLessThanOrEqual(versions[2]!.at)
  })

  it('refuses a snapshot identical to the newest one', () => {
    // The point of the whole dedup design: Ctrl+S out of habit must not turn a
    // day's work into a hundred identical copies of a large document.
    history.snapshot(DOC, Buffer.from('unchanged'))
    const second = history.snapshot(DOC, Buffer.from('unchanged'))

    expect(second.stored).toBe(false)
    expect(second.stored === false && second.reason).toBe('unchanged')
    expect(history.list(DOC)).toHaveLength(1)
  })

  it('does record a return to earlier content, and shares the one blob for it', () => {
    // Not the same as the case above: going back to a previous state IS a real
    // event in the writing process and must appear in the timeline — while still
    // costing no extra disk, since the bytes are already stored.
    history.snapshot(DOC, Buffer.from('A'))
    history.snapshot(DOC, Buffer.from('B'))
    history.snapshot(DOC, Buffer.from('A'))

    const versions = history.list(DOC)
    expect(versions).toHaveLength(3)
    expect(versions[0]!.id).toBe(versions[2]!.id)

    const blobs = fs.readdirSync(dirFor(DOC)).filter((f) => f.endsWith('.gz'))
    expect(blobs, 'identical content must not be stored twice').toHaveLength(2)
  })

  it('keeps two documents' + ' histories completely separate', () => {
    history.snapshot(DOC, Buffer.from('thesis content'))
    history.snapshot(OTHER, Buffer.from('letter content'))

    expect(history.list(DOC)).toHaveLength(1)
    expect(history.list(OTHER)).toHaveLength(1)
    expect(history.read(DOC, history.list(DOC)[0]!.id)?.toString()).toBe('thesis content')

    // The guard that matters: one document's id must not read another's blob,
    // even though both live under the same root.
    const otherId = history.list(OTHER)[0]!.id
    expect(history.read(DOC, otherId)).toBeNull()
  })

  it('refuses an id that is not in the document index', () => {
    history.snapshot(DOC, Buffer.from('content'))
    expect(history.read(DOC, 'f'.repeat(64))).toBeNull()
    // And anything that is not a hash at all, before it reaches the filesystem.
    expect(history.read(DOC, '../../../etc/passwd')).toBeNull()
    expect(history.read(DOC, '')).toBeNull()
  })

  it('reports an unusable document path instead of throwing', () => {
    // `normalizePath` returns null for a path it will not vouch for. Hashing that
    // would throw, or worse file several such paths under one shared history.
    const result = history.snapshot('', Buffer.from('x'))
    expect(result.stored).toBe(false)
    expect(result.stored === false && result.reason).toBe('invalidPath')
    expect(history.list('')).toEqual([])
    expect(history.read('', 'a'.repeat(64))).toBeNull()
  })

  it('survives a corrupt index by reporting fewer versions, never by throwing', () => {
    history.snapshot(DOC, Buffer.from('good'))
    const dir = dirFor(DOC)
    fs.writeFileSync(path.join(dir, INDEX_FILE_NAME), '{ this is not json', 'utf8')

    expect(() => history.list(DOC)).not.toThrow()
    expect(history.list(DOC)).toEqual([])
    // And it must recover: a later snapshot starts a fresh, valid index.
    expect(history.snapshot(DOC, Buffer.from('after corruption')).stored).toBe(true)
    expect(history.list(DOC)).toHaveLength(1)
  })

  it('drops malformed entries from an index rather than trusting them', () => {
    history.snapshot(DOC, Buffer.from('good'))
    const dir = dirFor(DOC)
    const real = JSON.parse(fs.readFileSync(path.join(dir, INDEX_FILE_NAME), 'utf8'))
    fs.writeFileSync(
      path.join(dir, INDEX_FILE_NAME),
      JSON.stringify([...real, { id: 'not-a-hash', at: 1, bytes: 1 }, null, { at: 2 }]),
      'utf8',
    )

    const versions = history.list(DOC)
    expect(versions.every((v) => /^[a-f0-9]{64}$/.test(v.id))).toBe(true)
  })

  it('prunes oldest-first once the version cap is passed', () => {
    for (let i = 0; i < MAX_VERSIONS_PER_DOCUMENT + 5; i += 1) {
      history.snapshot(DOC, Buffer.from(`draft ${i}`))
    }

    const versions = history.list(DOC)
    expect(versions).toHaveLength(MAX_VERSIONS_PER_DOCUMENT)
    // The newest must always survive; the oldest are what go.
    expect(history.read(DOC, versions[versions.length - 1]!.id)?.toString()).toBe(
      `draft ${MAX_VERSIONS_PER_DOCUMENT + 4}`,
    )
  })

  it('deletes the blobs of pruned versions, not just their index entries', () => {
    for (let i = 0; i < MAX_VERSIONS_PER_DOCUMENT + 10; i += 1) {
      history.snapshot(DOC, Buffer.from(`draft ${i}`))
    }
    const dir = dirFor(DOC)
    const blobs = fs.readdirSync(dir).filter((f) => f.endsWith('.gz'))

    // Orphaned blobs would make the cap meaningless: the index would look bounded
    // while the directory grew forever.
    expect(blobs.length).toBeLessThanOrEqual(MAX_VERSIONS_PER_DOCUMENT + 1)
  })

  it('stores blobs gzipped, so a long history of prose stays small', () => {
    const prose = Buffer.from('the quick brown fox '.repeat(2000))
    history.snapshot(DOC, prose)
    const dir = dirFor(DOC)
    const blob = fs.readdirSync(dir).find((f) => f.endsWith('.gz'))!
    const onDisk = fs.statSync(path.join(dir, blob)).size

    expect(onDisk).toBeLessThan(prose.length / 2)
    expect(zlib.gunzipSync(fs.readFileSync(path.join(dir, blob)))).toEqual(prose)
  })

  it('forgets a document entirely on clear, and leaves other documents alone', () => {
    history.snapshot(DOC, Buffer.from('thesis'))
    history.snapshot(OTHER, Buffer.from('letter'))

    history.clear(DOC)

    expect(history.list(DOC)).toEqual([])
    expect(history.list(OTHER)).toHaveLength(1)
  })

  it('treats differently-spelled paths to the same file as one history', () => {
    // Otherwise opening a file by a different spelling silently starts a second,
    // empty history and the user's versions appear to have vanished.
    history.snapshot('C:/Users/Someone/Documents/thesis.docx', Buffer.from('v1'))
    history.snapshot('C:\\Users\\Someone\\Documents\\thesis.docx', Buffer.from('v2'))

    expect(history.list(DOC)).toHaveLength(2)
  })

  it('carries an optional label so a version can say why it exists', () => {
    history.snapshot(DOC, Buffer.from('content'), 'before restoring an earlier draft')
    expect(history.list(DOC)[0]!.label).toBe('before restoring an earlier draft')

    history.snapshot(DOC, Buffer.from('other'))
    expect(history.list(DOC)[1]!.label).toBeNull()
  })
})
