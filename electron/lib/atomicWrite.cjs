// Shared atomic-write helper used by both the `save-file` (markdown/text)
// and `save-binary-file` (DOCX and other binary formats) IPC handlers.
//
// Pattern: write to a temp file in the SAME directory as the target (so the
// final rename is same-volume whenever possible) -> fsync -> keep one
// rolling `.bak` copy of whatever was previously at the target -> rename the
// temp file over the target. If the process dies before the rename, the
// original file is untouched.
//
// Windows/UNC nuance (ELEC-25): a cross-volume rename (e.g. some network
// shares) can fail with EXDEV. When that happens we fall back to a direct,
// non-atomic write — still preceded by the same `.bak` backup — rather than
// failing the save outright.
//
// Lock nuance (ELEC-26): if the target (or temp file) is held open by
// another program — most plausibly Microsoft Word holding its own lock on
// the same `.docx` — Windows raises EBUSY/EPERM. We surface a specific,
// friendly `FileLockedError` instead of a generic failure.

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const LOCK_ERROR_MESSAGE =
  'This file appears to be open in another program — close it there first and try again.';

class FileLockedError extends Error {
  constructor(message = LOCK_ERROR_MESSAGE) {
    super(message);
    this.name = 'FileLockedError';
  }
}

/**
 * @param {unknown} err
 * @returns {boolean}
 */
function isLockError(err) {
  const code = err && typeof err === 'object' ? /** @type {{code?: unknown}} */ (err).code : undefined;
  return code === 'EBUSY' || code === 'EPERM';
}

/**
 * @param {unknown} err
 * @returns {boolean}
 */
function isCrossDeviceError(err) {
  const code = err && typeof err === 'object' ? /** @type {{code?: unknown}} */ (err).code : undefined;
  return code === 'EXDEV';
}

/**
 * @param {string | Uint8Array} data
 * @returns {Buffer}
 */
function toBuffer(data) {
  if (typeof data === 'string') {
    return Buffer.from(data, 'utf-8');
  }
  if (Buffer.isBuffer(data)) {
    return data;
  }
  return Buffer.from(data);
}

/**
 * @param {string} filePath
 * @param {Buffer} buffer
 */
function writeAndSync(filePath, buffer) {
  const fd = fs.openSync(filePath, 'w');
  try {
    fs.writeSync(fd, buffer);
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
}

/**
 * @param {string} tempPath
 */
function cleanupTemp(tempPath) {
  try {
    fs.unlinkSync(tempPath);
  } catch {
    // Best-effort cleanup only — the temp file is harmless leftover state.
  }
}

/**
 * Copies whatever currently exists at `targetPath` to `${targetPath}.bak`,
 * overwriting any previous backup. Best-effort: a failure here must not
 * abort the save itself.
 * @param {string} targetPath
 */
function backupExisting(targetPath) {
  if (!fs.existsSync(targetPath)) {
    return;
  }
  try {
    fs.copyFileSync(targetPath, `${targetPath}.bak`);
  } catch {
    // Backup is a nice-to-have safety net, not a save precondition.
  }
}

/**
 * @param {string} targetPath
 * @returns {boolean}
 */
function isExistingDirectory(targetPath) {
  try {
    return fs.statSync(targetPath).isDirectory();
  } catch {
    return false;
  }
}

/**
 * Found by driving the real app: a file with Windows's read-only attribute
 * set (`attrib +R`, or `fs.chmodSync(path, 0o444)`) raises EPERM when
 * `atomicWriteFile` renames its temp file over it — the exact same code
 * `isLockError` below treats as "another program has this file locked".
 * That misreported a plain read-only file as if Word (or similar) had it
 * open, which sends the user to close a program that was never open. Same
 * shape of ambiguity as the EISDIR case above; disambiguated the same way,
 * by checking the target's actual state before trusting the error code.
 * @param {string} targetPath
 * @returns {boolean}
 */
function isReadOnlyFile(targetPath) {
  try {
    const stats = fs.statSync(targetPath);
    // Windows surfaces FILE_ATTRIBUTE_READONLY to Node by clearing the
    // write bits in `mode` — this is the same signal Node itself uses.
    return stats.isFile() && (stats.mode & 0o200) === 0;
  } catch {
    return false;
  }
}

/**
 * @param {unknown} err
 * @param {string} [targetPath] the path the failing operation was writing to
 *   (omitted for the initial temp-file write, which is never `targetPath`
 *   itself — see call site)
 * @returns {never}
 */
function rethrowAsFriendlyError(err, targetPath) {
  // X5 fix (found during wave3/export review) — on Windows, renaming (or
  // directly writing, in the EXDEV fallback) ONTO an existing directory
  // raises EPERM, not EISDIR: verified empirically (`fs.renameSync(tempFile,
  // existingDir)` -> `err.code === 'EPERM'`). `isLockError` below also
  // treats EPERM as "another program has this file locked" (Word, etc. —
  // ELEC-26), so without this check every "you picked a folder, not a file"
  // mistake was being misreported as a file-lock error instead of
  // `classifyWriteError`'s intended EISDIR message — the opposite of
  // actionable. Disambiguate by checking the target's actual type: a real
  // lock never has a directory sitting at `targetPath`.
  if (targetPath && isExistingDirectory(targetPath)) {
    /** @type {NodeJS.ErrnoException} */
    const dirErr = new Error('target path is a directory');
    dirErr.code = 'EISDIR';
    throw dirErr;
  }
  // Same EPERM ambiguity, this time against a read-only file — see
  // `isReadOnlyFile`'s comment.
  if (targetPath && isReadOnlyFile(targetPath)) {
    /** @type {NodeJS.ErrnoException} */
    const roErr = new Error('target path is read-only');
    roErr.code = 'EROFS';
    throw roErr;
  }
  if (isLockError(err)) {
    throw new FileLockedError();
  }
  throw err;
}

/**
 * X5/save-error-classification — maps a raw Node `fs` error code to a
 * friendly, actionable message. `save-file`/`save-binary-file` in `main.cjs`
 * previously only recognized `FileLockedError` (EBUSY/EPERM, thrown above)
 * and silently returned `error: undefined` for every other failure class,
 * leaving the renderer's banner/toast with nothing specific to show. Returns
 * `undefined` for a code with no friendly mapping so the caller can fall back
 * to its own generic message rather than this throwing or guessing.
 * @param {unknown} err
 * @returns {string | undefined}
 */
function classifyWriteError(err) {
  const code = err && typeof err === 'object' ? /** @type {{code?: unknown}} */ (err).code : undefined;
  switch (code) {
    case 'EACCES':
      return "Permission denied — you don't have access to save to this location.";
    case 'ENOSPC':
      return 'Not enough disk space to save this file.';
    case 'EISDIR':
      return 'That location is a folder, not a file — choose a different name.';
    case 'ENOENT':
      return 'The destination folder no longer exists — choose a different location.';
    case 'EROFS':
      return 'That location is read-only — choose a different location.';
    case 'ENAMETOOLONG':
      return 'That file name or path is too long — choose a shorter one.';
    default:
      return undefined;
  }
}

/**
 * i18n — the stable counterpart to `classifyWriteError`'s English text.
 * `main.cjs`'s IPC handlers attach this as `errorCode` alongside the English
 * `error` string so the renderer can look up a translated message (see
 * `src/i18n`'s `errors.write.*` keys, one per code below) instead of
 * displaying main's English text verbatim in a French UI. Kept as a
 * SEPARATE function (rather than changing `classifyWriteError`'s return
 * shape) so every existing caller/test of `classifyWriteError` keeps working
 * unchanged.
 * @param {unknown} err
 * @returns {string | undefined}
 */
function classifyWriteErrorCode(err) {
  if (err instanceof FileLockedError) return 'fileLocked';
  const code = err && typeof err === 'object' ? /** @type {{code?: unknown}} */ (err).code : undefined;
  switch (code) {
    case 'EACCES':
      return 'permissionDenied';
    case 'ENOSPC':
      return 'diskFull';
    case 'EISDIR':
      return 'isDirectory';
    case 'ENOENT':
      return 'destinationMissing';
    case 'EROFS':
      return 'readOnly';
    case 'ENAMETOOLONG':
      return 'nameTooLong';
    default:
      return undefined;
  }
}

/**
 * Atomically writes `data` to `targetPath`.
 * @param {string} targetPath
 * @param {string | Uint8Array} data
 * @returns {{ fallbackUsed: boolean }}
 */
function atomicWriteFile(targetPath, data) {
  const buffer = toBuffer(data);
  const dir = path.dirname(targetPath);
  const tempPath = path.join(
    dir,
    `.${path.basename(targetPath)}.tmp-${crypto.randomBytes(6).toString('hex')}`,
  );

  try {
    writeAndSync(tempPath, buffer);
  } catch (err) {
    cleanupTemp(tempPath);
    rethrowAsFriendlyError(err);
  }

  backupExisting(targetPath);

  try {
    fs.renameSync(tempPath, targetPath);
    return { fallbackUsed: false };
  } catch (err) {
    if (isCrossDeviceError(err)) {
      try {
        writeAndSync(targetPath, buffer);
        return { fallbackUsed: true };
      } catch (fallbackErr) {
        rethrowAsFriendlyError(fallbackErr, targetPath);
      } finally {
        cleanupTemp(tempPath);
      }
    }
    cleanupTemp(tempPath);
    rethrowAsFriendlyError(err, targetPath);
  }
}

module.exports = { atomicWriteFile, FileLockedError, LOCK_ERROR_MESSAGE, classifyWriteError, classifyWriteErrorCode };
