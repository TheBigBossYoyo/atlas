// Main-process path allowlist.
//
// Tracks the set of absolute file paths this window is allowed to read from,
// and the (smaller) set it may silently write to.
//
// Two trust tiers, not one (SEC-1):
//   - READ tier ("allowed"): any path the main process has been told about
//     by a flow it can partially verify — this includes drag-drop, whose
//     provenance the main process CANNOT fully verify (a compromised
//     renderer can call `path:register-dropped` with any string; the only
//     check is `fs.existsSync`). Being readable only lets a hostile
//     renderer read a file it could already have picked via `<input
//     type=file>`-style probing anyway once it names the path, which is a
//     much smaller blast radius than silently overwriting it.
//   - WRITE tier ("writable"): only paths the main process itself produced
//     or independently re-validated — open/save dialog results, argv/
//     second-instance/open-file OS paths, and recent-files entries
//     re-checked against the persisted, renderer-unreachable
//     `recentFilesStore`. `save-file`/`save-binary-file` may only silently
//     overwrite (skip the save dialog for) an `existingPath` in this tier.
//
// Design note: two Sets rather than a single Map<path, tier> or a
// tier-flag object. `trust()` always adds to both sets, so WRITE is a
// subset of READ by construction — there is no representable "writable but
// not readable" state to keep in sync, and every existing `has()` call site
// (read gates: file:readBinaryByPath, open-file-by-path, shell:reveal-in-
// folder, code:run) keeps working unchanged. A single Set of {path: tier}
// entries would need the same subset invariant enforced by hand on every
// write, for no simplification — two independent Sets are the more direct
// model of "read" and "write" as separate, comparably-sized capabilities.
//
// Path comparison is normalized so a `\\?\`-long-path-prefixed path, a UNC
// network path, mismatched path separators, or a differently-cased drive
// letter all compare consistently (Windows filesystem paths are
// case-insensitive) instead of being falsely rejected (ELEC-25).

const path = require('path');

/**
 * Normalizes a raw path for allowlist comparison. Returns `null` when the
 * input cannot be normalized (not a non-empty string).
 * @param {unknown} rawPath
 * @returns {string | null}
 */
function normalizePath(rawPath) {
  if (typeof rawPath !== 'string' || rawPath.length === 0) {
    return null;
  }

  let normalized = rawPath.replace(/\//g, '\\');

  // Strip Windows extended-length path prefixes so a long-path-safe caller
  // and a plain caller referring to the same file compare equal.
  if (normalized.startsWith('\\\\?\\UNC\\')) {
    normalized = `\\\\${normalized.slice('\\\\?\\UNC\\'.length)}`;
  } else if (normalized.startsWith('\\\\?\\')) {
    normalized = normalized.slice('\\\\?\\'.length);
  }

  try {
    normalized = path.resolve(normalized);
  } catch {
    return null;
  }

  // Windows paths (local and UNC) are case-insensitive.
  return normalized.toLowerCase();
}

/**
 * Creates a fresh, empty path allowlist.
 * @returns {{
 *   add: (rawPath: unknown) => string | null,
 *   trust: (rawPath: unknown) => string | null,
 *   has: (rawPath: unknown) => boolean,
 *   isWriteEligible: (rawPath: unknown) => boolean,
 *   remove: (rawPath: unknown) => void,
 *   clear: () => void,
 *   size: () => number,
 * }}
 */
function createPathAllowlist() {
  /** @type {Set<string>} READ tier — readable, not necessarily writable. */
  const allowed = new Set();
  /** @type {Set<string>} WRITE tier — always a subset of `allowed` (see header). */
  const writable = new Set();

  return {
    // READ-only registration. Use for a path whose provenance the main
    // process cannot fully verify — currently just drag-drop
    // (`path:register-dropped`). Never makes a path write-eligible.
    add(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized !== null) {
        allowed.add(normalized);
      }
      return normalized;
    },
    // Full trust: READ + WRITE. Use only for a path the main process itself
    // produced or independently re-validated (dialog results, argv/
    // second-instance/open-file, re-validated recent-files entries).
    trust(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized !== null) {
        allowed.add(normalized);
        writable.add(normalized);
      }
      return normalized;
    },
    has(rawPath) {
      const normalized = normalizePath(rawPath);
      return normalized !== null && allowed.has(normalized);
    },
    // Whether `existingPath` may be silently overwritten (no save dialog).
    isWriteEligible(rawPath) {
      const normalized = normalizePath(rawPath);
      return normalized !== null && writable.has(normalized);
    },
    remove(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized !== null) {
        allowed.delete(normalized);
        writable.delete(normalized);
      }
    },
    clear() {
      allowed.clear();
      writable.clear();
    },
    size() {
      return allowed.size;
    },
  };
}

module.exports = { createPathAllowlist, normalizePath };
