// Rotating crash/error log (ELEC-13/P1.15).
//
// Writes timestamped lines to `<logDir>/main.log`. Once the file would
// exceed MAX_LOG_BYTES, it is truncated and restarted rather than growing
// without bound — this is a basic diagnostic trail, not a full audit log.

const fs = require('fs');
const path = require('path');

const MAX_LOG_BYTES = 1 * 1024 * 1024; // 1 MiB
const LOG_FILE_NAME = 'main.log';

/**
 * @param {string} logDir
 * @returns {string}
 */
function getLogFilePath(logDir) {
  return path.join(logDir, LOG_FILE_NAME);
}

/**
 * @param {unknown} error
 * @returns {string | undefined}
 */
function formatErrorDetail(error) {
  if (error instanceof Error) {
    return error.stack ?? `${error.name}: ${error.message}`;
  }
  if (error === undefined) {
    return undefined;
  }
  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

/**
 * @param {string} level
 * @param {string} message
 * @param {unknown} [error]
 * @returns {string}
 */
function formatLogLine(level, message, error) {
  const timestamp = new Date().toISOString();
  const detail = formatErrorDetail(error);
  const suffix = detail !== undefined ? `\n${detail}` : '';
  return `[${timestamp}] [${level}] ${message}${suffix}\n`;
}

/**
 * Appends `line` to the log file under `logDir`, rotating (truncating) the
 * file first if appending would exceed MAX_LOG_BYTES. Best-effort: logging
 * failures must never throw, since they run from crash handlers.
 * @param {string} logDir
 * @param {string} line
 */
function appendLogLine(logDir, line) {
  try {
    fs.mkdirSync(logDir, { recursive: true });
    const filePath = getLogFilePath(logDir);
    let existingSize = 0;
    try {
      existingSize = fs.statSync(filePath).size;
    } catch {
      existingSize = 0;
    }

    const lineSize = Buffer.byteLength(line, 'utf-8');
    if (existingSize + lineSize > MAX_LOG_BYTES) {
      fs.writeFileSync(filePath, line, { encoding: 'utf-8' });
    } else {
      fs.appendFileSync(filePath, line, { encoding: 'utf-8' });
    }
  } catch {
    // A logging failure must never crash the app it is trying to diagnose.
  }
}

/**
 * @param {string} logDir
 * @param {string} level
 * @param {string} message
 * @param {unknown} [error]
 */
function logToFile(logDir, level, message, error) {
  appendLogLine(logDir, formatLogLine(level, message, error));
}

module.exports = {
  MAX_LOG_BYTES,
  LOG_FILE_NAME,
  getLogFilePath,
  formatLogLine,
  appendLogLine,
  logToFile,
};
