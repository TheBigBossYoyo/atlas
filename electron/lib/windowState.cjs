// P5.2/ELEC-11 — window bounds/maximized-state persistence between launches.
//
// Split into pure, Electron-free functions (this module) plus main.cjs's own
// wiring (load on create, save on resize/move/close) for the same reason as
// closeGuard.cjs: `screen`/`BrowserWindow` can't be constructed outside a
// running Electron app, so anything that touches them directly is
// effectively untestable from Vitest. `clampBoundsToDisplays` in particular
// is the part most worth covering directly — restoring a saved position from
// a monitor that's since been disconnected, resized, or had its resolution
// changed must never place the window fully off-screen with no way to drag
// it back.

const fs = require('fs');
const path = require('path');

const STATE_FILE_NAME = 'window-state.json';

/** How many pixels of a restored window must overlap some connected
 * display's work area for that position to count as "recoverable" — small
 * enough that a window mostly off-screen (but with, say, its title bar
 * still draggable back into view) is still honored, large enough that a
 * window landing entirely off every current display is rejected. */
const MIN_VISIBLE_PX = 50;

/**
 * @param {{ x: number, y: number, width: number, height: number }} bounds
 * @param {ReadonlyArray<{ workArea: { x: number, y: number, width: number, height: number } }>} displays
 * @returns {boolean}
 */
function isSufficientlyVisible(bounds, displays) {
  return displays.some(({ workArea }) => {
    const overlapWidth = Math.min(bounds.x + bounds.width, workArea.x + workArea.width) - Math.max(bounds.x, workArea.x);
    const overlapHeight = Math.min(bounds.y + bounds.height, workArea.y + workArea.height) - Math.max(bounds.y, workArea.y);
    return overlapWidth >= MIN_VISIBLE_PX && overlapHeight >= MIN_VISIBLE_PX;
  });
}

/**
 * Returns `bounds` unchanged when it is still sufficiently on-screen against
 * one of `displays`' work areas, or `null` when it isn't (a disconnected
 * external monitor, a resolution change, etc.) so the caller can fall back
 * to its own default sizing/centering instead.
 * @param {{ x: number, y: number, width: number, height: number } | null} bounds
 * @param {ReadonlyArray<{ workArea: { x: number, y: number, width: number, height: number } }>} displays
 * @returns {{ x: number, y: number, width: number, height: number } | null}
 */
function clampBoundsToDisplays(bounds, displays) {
  if (!bounds || displays.length === 0) return null;
  return isSufficientlyVisible(bounds, displays) ? bounds : null;
}

/**
 * @param {string} storeDir
 * @returns {string}
 */
function getStateFilePath(storeDir) {
  return path.join(storeDir, STATE_FILE_NAME);
}

/**
 * @param {unknown} value
 * @returns {value is number}
 */
function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

/**
 * Reads and validates the persisted window state, returning `null` for a
 * missing, corrupt, or structurally-invalid file rather than throwing — a
 * damaged state file should never block the app from starting.
 * @param {string} storeDir
 * @returns {{ x: number, y: number, width: number, height: number, isMaximized: boolean } | null}
 */
function loadWindowState(storeDir) {
  try {
    const raw = fs.readFileSync(getStateFilePath(storeDir), 'utf-8');
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    const { x, y, width, height, isMaximized } = parsed;
    if (![x, y, width, height].every(isFiniteNumber) || width <= 0 || height <= 0) return null;
    return { x, y, width, height, isMaximized: isMaximized === true };
  } catch {
    return null;
  }
}

/**
 * @param {string} storeDir
 * @param {{ x: number, y: number, width: number, height: number, isMaximized: boolean }} state
 */
function saveWindowState(storeDir, state) {
  try {
    fs.mkdirSync(storeDir, { recursive: true });
    fs.writeFileSync(getStateFilePath(storeDir), JSON.stringify(state), 'utf-8');
  } catch {
    // Best-effort only, matching recentFilesStore.cjs — losing this file
    // just means the next launch uses the default size/position.
  }
}

module.exports = {
  STATE_FILE_NAME,
  MIN_VISIBLE_PX,
  isSufficientlyVisible,
  clampBoundsToDisplays,
  loadWindowState,
  saveWindowState,
};
