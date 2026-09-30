// VERSIONS-1 — main-process-owned version history: a rolling series of
// snapshots of a document's bytes, so the user can look back at how a piece of
// writing got to where it is and go back to any point in it.
//
// WHY MAIN AND NOT THE RENDERER. Two reasons, and the second is the important
// one. Snapshots are file I/O, which this app deliberately keeps entirely in the
// main process (ELEC-02/03). And a history is only worth having if it cannot be
// tampered with from the page: a renderer-writable store would let any script
// running in a rendered document rewrite the user's own record of their work.
// The renderer names a document and asks; it never supplies a storage path.
//
// LAYOUT, and why it is content-addressed:
//
//   <userData>/history/<sha256(normalised document path)>/
//     index.json              the ordered list of versions, newest last
//     <sha256(bytes)>.gz      one gzipped blob per DISTINCT document state
//
// The blob name is the hash of its own contents, so saving a document twice
// without changing anything stores one blob and the second snapshot is refused
// outright (see `snapshot`'s `unchanged` result). That matters more than it
// sounds: the natural way to use this feature is Ctrl+S out of habit, and
// without dedup a day of nervous saving would be a hundred identical copies of a
// 20 MB .docx. Two versions that happen to hold identical bytes also share one
// blob, so reverting and carrying on costs nothing extra.
//
// The directory name is a hash of the path rather than the path itself because a
// Windows path contains characters a directory name cannot, and because nesting
// user-controlled path fragments under userData is exactly the sort of thing
// that turns into a traversal bug.
//
// WHAT IT DELIBERATELY DOES NOT DO. It does not diff. A diff worth showing is
// per-format (a Word paragraph diff and a spreadsheet cell diff have nothing in
// common) and belongs above this layer, which is why this one stays honest about
// being a store of bytes.

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const { normalizePath } = require('./pathAllowlist.cjs');

const INDEX_FILE_NAME = 'index.json';

/**
 * Per-document caps. Generous for prose, deliberately bounded for the binary
 * formats: a 30 MB spreadsheet saved often would otherwise fill a disk quietly,
 * and the failure mode of a full disk is far worse than losing the oldest
 * snapshot of last week's draft.
 */
const MAX_VERSIONS_PER_DOCUMENT = 100;
const MAX_BYTES_PER_DOCUMENT = 256 * 1024 * 1024;

/** @param {string} value */
function sha256(value) {
  return crypto.createHash('sha256').update(value).digest('hex');
}

/** @param {Buffer} bytes */
function sha256Bytes(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

/**
 * The directory holding one document's history.
 *
 * Keyed off `normalizePath` so the same file reached by a different spelling
 * (case, separators) shares one history rather than silently starting a second.
 *
 * `null` when the path cannot be normalised at all. That is not a theoretical
 * branch: `normalizePath` returns null for anything it will not vouch for, and
 * hashing that would either throw or — worse — quietly file several different
 * unnormalisable paths under one shared history.
 *
 * @param {string} rootDir
 * @param {string} documentPath
 * @returns {string | null}
 */
function documentDir(rootDir, documentPath) {
  const normalised = normalizePath(documentPath);
  return normalised === null ? null : path.join(rootDir, sha256(normalised));
}

/**
 * @typedef {{
 *   readonly id: string,
 *   readonly at: number,
 *   readonly bytes: number,
 *   readonly label: string | null,
 * }} VersionEntry
 */

/**
 * @param {string} dir
 * @returns {VersionEntry[]}
 */
function readIndex(dir) {
  try {
    const raw = fs.readFileSync(path.join(dir, INDEX_FILE_NAME), 'utf8');
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    // Defensive on every field: this file lives on disk across upgrades, and a
    // half-written or hand-edited index must degrade to "fewer versions", never
    // to a crash on open.
    return parsed.filter(
      (entry) =>
        entry &&
        typeof entry.id === 'string' &&
        /^[a-f0-9]{64}$/.test(entry.id) &&
        typeof entry.at === 'number' &&
        Number.isFinite(entry.at) &&
        typeof entry.bytes === 'number',
    );
  } catch {
    return [];
  }
}

/**
 * @param {string} dir
 * @param {VersionEntry[]} entries
 */
function writeIndex(dir, entries) {
  fs.mkdirSync(dir, { recursive: true });
  const tmp = path.join(dir, `${INDEX_FILE_NAME}.tmp`);
  fs.writeFileSync(tmp, JSON.stringify(entries), 'utf8');
  fs.renameSync(tmp, path.join(dir, INDEX_FILE_NAME));
}

/** @param {string} dir @param {string} id */
function blobPath(dir, id) {
  return path.join(dir, `${id}.gz`);
}

/**
 * Drops oldest-first until both caps are satisfied, deleting only blobs no
 * remaining entry still references (two entries can share one blob — see the
 * header on dedup).
 *
 * @param {string} dir
 * @param {VersionEntry[]} entries
 * @returns {VersionEntry[]}
 */
function prune(dir, entries) {
  let kept = entries.slice();
  const total = () => kept.reduce((sum, entry) => sum + entry.bytes, 0);

  while (kept.length > MAX_VERSIONS_PER_DOCUMENT || (kept.length > 1 && total() > MAX_BYTES_PER_DOCUMENT)) {
    kept.shift();
  }

  if (kept.length !== entries.length) {
    const live = new Set(kept.map((entry) => entry.id));
    for (const dropped of entries) {
      if (live.has(dropped.id)) continue;
      try {
        fs.rmSync(blobPath(dir, dropped.id), { force: true });
      } catch {
        // A blob that cannot be removed is wasted disk, not a broken history:
        // the index no longer references it, so nothing will try to read it.
      }
    }
  }
  return kept;
}

/**
 * @param {string} rootDir where all histories live (a subdirectory of userData)
 */
function createVersionHistory(rootDir) {
  return {
    /**
     * Records the document's current bytes as a new version.
     *
     * Returns `{ stored: false, reason: 'unchanged' }` when the newest version
     * already holds exactly these bytes — the common case when someone saves out
     * of habit, and worth reporting rather than silently succeeding so callers
     * can avoid claiming they made a version that does not exist.
     *
     * @param {string} documentPath
     * @param {Buffer} bytes
     * @param {string | null} [label]
     */
    snapshot(documentPath, bytes, label = null) {
      const dir = documentDir(rootDir, documentPath);
      if (dir === null) return { stored: false, reason: 'invalidPath', id: '' };
      const entries = readIndex(dir);
      const id = sha256Bytes(bytes);

      const newest = entries[entries.length - 1];
      if (newest && newest.id === id) return { stored: false, reason: 'unchanged', id };

      fs.mkdirSync(dir, { recursive: true });
      const blob = blobPath(dir, id);
      if (!fs.existsSync(blob)) {
        // Through a temp file: a half-written blob that kept its final name
        // would be indistinguishable from a good one and would fail only later,
        // when the user actually tried to open that version.
        const tmp = `${blob}.tmp`;
        fs.writeFileSync(tmp, zlib.gzipSync(bytes));
        fs.renameSync(tmp, blob);
      }

      const entry = { id, at: Date.now(), bytes: bytes.length, label: label ?? null };
      writeIndex(dir, prune(dir, [...entries, entry]));
      return { stored: true, id, at: entry.at };
    },

    /**
     * Every version of a document, oldest first.
     * @param {string} documentPath
     * @returns {VersionEntry[]}
     */
    list(documentPath) {
      const dir = documentDir(rootDir, documentPath);
      return dir === null ? [] : readIndex(dir);
    },

    /**
     * One version's bytes, or null when it is not in this document's index.
     *
     * Checking the index rather than just reading `<id>.gz` is what stops a
     * caller asking for an arbitrary hash and reading another document's blob.
     *
     * @param {string} documentPath
     * @param {string} id
     * @returns {Buffer | null}
     */
    read(documentPath, id) {
      if (typeof id !== 'string' || !/^[a-f0-9]{64}$/.test(id)) return null;
      const dir = documentDir(rootDir, documentPath);
      if (dir === null) return null;
      if (!readIndex(dir).some((entry) => entry.id === id)) return null;
      try {
        return zlib.gunzipSync(fs.readFileSync(blobPath(dir, id)));
      } catch {
        return null;
      }
    },

    /**
     * Forgets a document's history entirely.
     * @param {string} documentPath
     */
    clear(documentPath) {
      const dir = documentDir(rootDir, documentPath);
      if (dir === null) return;
      try {
        fs.rmSync(dir, { recursive: true, force: true });
      } catch {
        // Nothing to do: the caller asked for it to be gone and it is either
        // gone or unreadable, which is indistinguishable from here.
      }
    },
  };
}

module.exports = {
  createVersionHistory,
  INDEX_FILE_NAME,
  MAX_VERSIONS_PER_DOCUMENT,
  MAX_BYTES_PER_DOCUMENT,
};
