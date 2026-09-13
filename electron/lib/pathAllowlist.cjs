// Main-process path allowlist.
//
// Tracks the set of absolute file paths this window is allowed to read from
// or silently write to. Populated ONLY from sources the main process itself
// vouches for: open-dialog results, argv/second-instance/open-file paths,
// drag-drop paths resolved via webUtils, save-as dialog results, and recent
// files re-validated through the "request open recent" IPC channel.
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
 *   has: (rawPath: unknown) => boolean,
 *   remove: (rawPath: unknown) => void,
 *   clear: () => void,
 *   size: () => number,
 * }}
 */
function createPathAllowlist() {
  /** @type {Set<string>} */
  const allowed = new Set();

  return {
    add(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized !== null) {
        allowed.add(normalized);
      }
      return normalized;
    },
    has(rawPath) {
      const normalized = normalizePath(rawPath);
      return normalized !== null && allowed.has(normalized);
    },
    remove(rawPath) {
      const normalized = normalizePath(rawPath);
      if (normalized !== null) {
        allowed.delete(normalized);
      }
    },
    clear() {
      allowed.clear();
    },
    size() {
      return allowed.size;
    },
  };
}

module.exports = { createPathAllowlist, normalizePath };
