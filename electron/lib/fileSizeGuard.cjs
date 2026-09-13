// Per-file read size cap with a friendly error message (part of P1.2/ELEC-02
// — the file-read IPC handlers had no size limit at all).

const fs = require('fs');

const DEFAULT_MAX_BYTES = 200 * 1024 * 1024; // 200 MiB

class FileTooLargeError extends Error {
  constructor(message) {
    super(message);
    this.name = 'FileTooLargeError';
  }
}

/**
 * @param {number} bytes
 * @returns {string}
 */
function formatMegabytes(bytes) {
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * @param {number} sizeBytes
 * @param {number} [maxBytes]
 */
function assertSizeAllowed(sizeBytes, maxBytes = DEFAULT_MAX_BYTES) {
  if (sizeBytes > maxBytes) {
    throw new FileTooLargeError(
      `This file is too large to open (${formatMegabytes(sizeBytes)}). Files over ${formatMegabytes(maxBytes)} are not supported.`,
    );
  }
}

/**
 * @param {string} filePath
 * @param {number} [maxBytes]
 */
function assertFileSizeAllowed(filePath, maxBytes = DEFAULT_MAX_BYTES) {
  const stats = fs.statSync(filePath);
  assertSizeAllowed(stats.size, maxBytes);
}

module.exports = { DEFAULT_MAX_BYTES, FileTooLargeError, assertSizeAllowed, assertFileSizeAllowed };
