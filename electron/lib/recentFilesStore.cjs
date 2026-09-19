// Main-process-owned ground truth for "files this app has genuinely, at some
// point, opened or saved through a trusted flow" — used to validate the
// `recent:request-open` IPC channel.
//
// Why this exists (security-review finding, wave1/electron-hardening): the
// renderer's own "recent files" list lives in `localStorage`, which any
// script running in the page (including a future XSS/sanitization-bypass
// payload from a rendered document) can read AND WRITE freely. If
// `recent:request-open` only checked `fs.existsSync(path)`, a script with
// page-level execution could claim ANY existing file on disk is a "recent
// file" and have it added to the main process's read/write allowlist —
// completely defeating the allowlist's purpose (ELEC-02/03). This module
// gives the main process its own persisted, renderer-unreachable record of
// which paths were actually vouched for by a trusted flow (an Open/Save
// dialog result, an argv/second-instance/open-file OS path), so
// `recent:request-open` can check real history instead of trusting the
// caller's claim.
//
// Deliberately NOT populated from drag-drop registrations
// (`path:register-dropped`): that channel has the same "renderer claims a
// path, main process cannot verify provenance" shape as the recent-files
// problem this module fixes, so feeding it back into this store would just
// move the hole rather than close it. The practical effect: a file that has
// only ever been drag-dropped (never opened via dialog/argv/save) will not
// be reopenable via "Recent Files" after an app restart — a narrow, accepted
// UX trade-off for closing the recent-files vector. See ELEC-02/03 review
// notes for the residual drag-drop-specific risk.

const fs = require('fs');
const path = require('path');

const { normalizePath } = require('./pathAllowlist.cjs');

const STORE_FILE_NAME = 'recent-files.json';
const MAX_ENTRIES = 50;

/**
 * @param {string} storeDir
 * @returns {string}
 */
function getStoreFilePath(storeDir) {
  return path.join(storeDir, STORE_FILE_NAME);
}

/**
 * @param {string} storeDir
 * @returns {string[]}
 */
function loadEntries(storeDir) {
  try {
    const raw = fs.readFileSync(getStoreFilePath(storeDir), 'utf-8');
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((entry) => typeof entry === 'string');
  } catch {
    return [];
  }
}

/**
 * @param {string} storeDir
 * @param {string[]} entries
 */
function saveEntries(storeDir, entries) {
  try {
    fs.mkdirSync(storeDir, { recursive: true });
    fs.writeFileSync(getStoreFilePath(storeDir), JSON.stringify(entries), 'utf-8');
  } catch {
    // Best-effort only — losing this store just means fewer paths survive
    // an app restart for "Open Recent"; it must never crash the app.
  }
}

/**
 * Creates a persisted store of normalized paths the main process has
 * genuinely vouched for, backed by a JSON file under `storeDir`.
 * @param {string} storeDir
 * @returns {{ record: (rawPath: unknown) => void, has: (rawPath: unknown) => boolean }}
 */
function createRecentFilesStore(storeDir) {
  /** @type {Set<string> | null} */
  let cache = null;

  function ensureLoaded() {
    if (cache === null) {
      cache = new Set(loadEntries(storeDir));
    }
    return cache;
  }

  return {
    record(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized === null) return;

      const entries = ensureLoaded();
      entries.delete(normalized); // re-insert at the end (most-recent-last)
      entries.add(normalized);

      while (entries.size > MAX_ENTRIES) {
        const oldest = entries.values().next().value;
        // `entries.size > MAX_ENTRIES` (>= 1) guarantees the iterator yields
        // a value here, but `Set#values().next()` is typed to allow
        // `undefined` (the "done" case) — guard it explicitly rather than
        // asserting past the type checker.
        if (oldest === undefined) break;
        entries.delete(oldest);
      }

      saveEntries(storeDir, [...entries]);
    },
    has(rawPath) {
      const normalized = normalizePath(rawPath);
      return normalized !== null && ensureLoaded().has(normalized);
    },
  };
}

module.exports = { createRecentFilesStore, STORE_FILE_NAME };
