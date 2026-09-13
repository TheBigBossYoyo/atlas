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
 * @param {unknown} err
 * @returns {never}
 */
function rethrowAsFriendlyError(err) {
  if (isLockError(err)) {
    throw new FileLockedError();
  }
  throw err;
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
        rethrowAsFriendlyError(fallbackErr);
      } finally {
        cleanupTemp(tempPath);
      }
    }
    cleanupTemp(tempPath);
    rethrowAsFriendlyError(err);
  }
}

module.exports = { atomicWriteFile, FileLockedError, LOCK_ERROR_MESSAGE };
