const { app, BrowserWindow, ipcMain, dialog, shell, Menu, screen } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');

const { createPathAllowlist } = require('./lib/pathAllowlist.cjs');
const { createRecentFilesStore } = require('./lib/recentFilesStore.cjs');
const { listSystemFontFamilies } = require('./lib/systemFonts.cjs');
const { createCodeRunner, runtimeFor: runCodeRuntimeFor } = require('./lib/codeRunner.cjs');
const { atomicWriteFile, FileLockedError, classifyWriteError, classifyWriteErrorCode } = require('./lib/atomicWrite.cjs');
const { printHtmlToPdfBuffer, PrintToPdfError } = require('./lib/printToPdf.cjs');
const { decodeTextBuffer } = require('./lib/textDecoding.cjs');
const { buildContentSecurityPolicy } = require('./lib/csp.cjs');
const { logToFile } = require('./lib/crashLog.cjs');
const { FileTooLargeError, assertFileSizeAllowed } = require('./lib/fileSizeGuard.cjs');
const { EXTENSIONS: MANIFEST_EXTENSIONS } = require('./lib/extensionManifest.generated.cjs');
const { CLOSE_PROMPT_BUTTONS, decideOnClose, decideAfterPromptChoice } = require('./lib/closeGuard.cjs');
const { clampBoundsToDisplays, loadWindowState, saveWindowState } = require('./lib/windowState.cjs');
const { decideFileOpenAction } = require('./lib/fileOpenRouting.cjs');
const { resolveIsDev } = require('./lib/devDetect.cjs');
const { NEW_DOCUMENT_FORMATS, templateBytesForExtension, readTemplateBytes } = require('./lib/newDocumentTemplates.cjs');

// E2E runs: every launch gets its own profile. Otherwise an instance that is
// still being torn down (taskkill /T is not instantaneous) keeps the shared
// profile's singleton lock file open and the next test's launch fails with
// "Lock file can not be created! Error code: 32" — and tests would also write
// into the user's real recent-files/window-state store.
if (process.env.PLAYWRIGHT === '1' && !app.isPackaged) {
  const e2eRoot = path.join(os.tmpdir(), 'atlas-e2e-profiles');
  fs.mkdirSync(e2eRoot, { recursive: true });
  app.setPath('userData', fs.mkdtempSync(path.join(e2eRoot, 'profile-')));
}

// Single instance lock
const gotLock = app.requestSingleInstanceLock();

if (!gotLock) {
  app.quit();
}

/** @type {BrowserWindow | null} */
let mainWindow = null;

/** @type {string | null} */
let pendingFilePath = null;

// P5.2/ELEC-07 — true once the *current* window's renderer has proven it is
// listening for `file-opened-path` (see fileOpenRouting.cjs's header for why
// the `get-initial-file` invoke is the right signal). Reset on every new
// window so a fresh, not-yet-mounted renderer isn't mistaken for the
// previous window's already-ready one.
/** @type {boolean} */
let rendererReady = false;

// ---- Close confirmation (P2.5 / SHELL-02, ELEC-06) ---- //
//
// The renderer has no way to show a visible "unsaved changes" prompt from a
// `beforeunload` handler alone — Electron/Chromium never surfaces one for a
// renderer-only listener. Instead the renderer pushes its combined dirty
// state (markdown's own `isDirty` OR the active viewer's registered dirty
// signal, from the P1.1 capability contract) here via a lightweight IPC
// signal every time it changes, and the `close` handler below decides
// whether to let the window close immediately or block it behind a native
// Save/Discard/Cancel prompt.
/** @type {boolean} */
let rendererDirty = false;

const OVERLAY_COLORS = {
  light:   { color: '#f6f8fa', symbolColor: '#1f2328' },
  dark:    { color: '#161b22', symbolColor: '#e6edf3' },
  sepia:   { color: '#f4ecd8', symbolColor: '#5b4636' },
  nord:    { color: '#2e3440', symbolColor: '#eceff4' },
  dracula: { color: '#282a36', symbolColor: '#f8f8f2' },
};

/** @type {{ color: string; symbolColor: string }} */
let currentOverlayColors = OVERLAY_COLORS.light;

/**
 * @param {unknown} theme
 * @returns {theme is keyof typeof OVERLAY_COLORS}
 */
function isKnownOverlayTheme(theme) {
  return typeof theme === 'string' && Object.prototype.hasOwnProperty.call(OVERLAY_COLORS, theme);
}

// ---- Dev/prod detection (RUN-12) ---- //
//
// Decision logic lives in `./lib/devDetect.cjs` (unit-tested there — see
// `src/__tests__/electron-lib/devDetect.test.ts` — the same split-out-for-
// testability pattern as closeGuard/csp/pathAllowlist, since Electron's own
// modules can't be constructed outside a running app). That file's header
// also documents a known nuance: this checks the filesystem, not how the
// process was launched, so a stale `dist/` from an earlier build can make a
// plain local `electron .` (including `npx playwright test`, whose specs
// launch Electron directly) resolve to prod against that stale build unless
// `ATLAS_DEV=1` overrides it.
const DIST_INDEX_PATH = path.join(app.getAppPath(), 'dist', 'index.html');

const isDev = resolveIsDev({
  atlasDevEnv: process.env.ATLAS_DEV,
  distIndexExists: () => fs.existsSync(DIST_INDEX_PATH),
});

// ---- Path allowlist (P1.2 / ELEC-02, ELEC-03, ELEC-25) ---- //
//
// The only paths any read/write IPC handler will act on are ones the main
// process itself vouches for: open-dialog results, argv/second-instance/
// open-file paths, drag-drop paths resolved via webUtils (registered
// through `path:register-dropped`), save-as dialog results, and recent
// files re-validated through `recent:request-open`.
const pathAllowlist = createPathAllowlist();

// ---- Recent-files ground truth (security-review fix) ---- //
//
// `recent:request-open` re-opens a path from the renderer's own "recent
// files" list — but that list lives in `localStorage`, which any script
// running in the page can read AND write. Checking only `fs.existsSync`
// there would let such a script claim ANY existing file on disk is "recent"
// and get it added to `pathAllowlist`, defeating the allowlist entirely
// (ELEC-02/03). This persisted, renderer-unreachable store records paths
// only when they arrive through a flow the main process itself trusts
// (dialog results, argv/second-instance/open-file, successful saves) —
// deliberately NOT from drag-drop registration, which has the same
// unverifiable-provenance shape. See recentFilesStore.cjs.
/**
 * `app.getPath('userData')` before `whenReady` (or if the profile directory
 * is otherwise unavailable) throws — every small persisted-JSON store in
 * this file (recent files, window state) falls back to the OS temp dir in
 * that case rather than failing to load/save at all.
 * @returns {string}
 */
function getUserDataDir() {
  try {
    return app.getPath('userData');
  } catch {
    return os.tmpdir();
  }
}

/** @type {import('./lib/recentFilesStore.cjs').RecentFilesStore | null} */
let recentFilesStoreInstance = null;
function getRecentFilesStore() {
  if (!recentFilesStoreInstance) {
    recentFilesStoreInstance = createRecentFilesStore(getUserDataDir());
  }
  return recentFilesStoreInstance;
}

/**
 * Adds `filePath` to both the session allowlist and the persisted
 * trusted-history store. Use this (instead of `pathAllowlist.add` directly)
 * at every call site where the path came from a source the main process
 * itself vouches for.
 * @param {unknown} filePath
 */
function trustPath(filePath) {
  pathAllowlist.add(filePath);
  getRecentFilesStore().record(filePath);
}

const NOT_ALLOWLISTED_MESSAGE =
  'This file cannot be opened because it was not selected through Atlas. Try File > Open instead.';
const SENDER_FRAME_ERROR_MESSAGE = 'This request could not be verified and was blocked.';
// UX — shown when a previously-opened/allowlisted path no longer exists on
// disk (deleted, renamed, or moved out from under Atlas, e.g. a stale
// "Recent" entry or the file being removed by another program while open).
// Replaces the bare `Error('Invalid path')` this used to throw, which read
// like Atlas itself had done something wrong rather than telling the user
// what happened and that it isn't recoverable from here.
const FILE_NOT_FOUND_MESSAGE = 'This file could not be found — it may have been moved, renamed, or deleted.';
// UX — dropping (or otherwise pointing Atlas at) a folder instead of a file.
// `registerDroppedPath` allowlists it (it does exist on disk — only
// `fs.existsSync` is checked there), and a folder has no recognized
// extension, so `useFileHandler`'s loader falls through to
// `file:readBinaryByPath`'s "unknown extension" path. Without this check,
// `fs.promises.readFile` throws the raw `EISDIR: illegal operation on a
// directory, read`, which used to reach the UI's error banner verbatim —
// the same class of bug FILE_NOT_FOUND_MESSAGE fixed for a missing file.
const IS_A_FOLDER_MESSAGE = "That's a folder, not a file — choose a file instead.";

/**
 * @param {Electron.IpcMainInvokeEvent} event
 * @returns {boolean}
 */
function isFromMainFrame(event) {
  if (!mainWindow || mainWindow.isDestroyed()) return false;
  try {
    return event.senderFrame === mainWindow.webContents.mainFrame;
  } catch {
    return false;
  }
}

/**
 * @param {string} level
 * @param {string} message
 * @param {unknown} [error]
 */
function logMainEvent(level, message, error) {
  let logDir;
  try {
    logDir = app.getPath('logs');
  } catch {
    logDir = os.tmpdir();
  }
  logToFile(logDir, level, message, error);
}

// ---- File path helpers ---- //

// P2.2/ELEC-05/ELEC-15/LOAD-03/LOAD-12 — derived from the single canonical
// manifest (src/formats/extensionManifest.ts) via
// scripts/generate-extension-manifest.mjs, instead of a hand-maintained
// literal that used to disagree with both detect.ts's EXTENSION_MAP and
// electron-builder.yml's fileAssociations (e.g. `.mdown`/`.ini` were
// registered Windows file associations that this set didn't recognize, so
// double-clicking them silently failed).
const KNOWN_EXTENSIONS = new Set(MANIFEST_EXTENSIONS);

/**
 * @param {unknown} argv
 * @returns {string | null}
 */
function extractFilePath(argv) {
  // In production, argv[0] is the exe, argv[1] might be the file
  // In dev, argv varies — look for known file extensions
  if (!Array.isArray(argv)) return null;
  for (let i = 1; i < argv.length; i++) {
    const arg = argv[i];
    if (typeof arg !== 'string') continue;
    if (arg.startsWith('--')) continue;
    const dot = arg.lastIndexOf('.');
    if (dot < 0) continue;
    const ext = arg.slice(dot + 1).toLowerCase();
    if (KNOWN_EXTENSIONS.has(ext) && fs.existsSync(arg)) {
      return path.resolve(arg);
    }
  }
  return null;
}

/**
 * Reads and decodes a text-class file. Returns `null` on any ordinary I/O
 * failure (matching the previous behavior of every caller); propagates
 * `FileTooLargeError` so the interactive open flows can surface a friendly
 * message instead of a silent no-op.
 * @param {string} filePath
 * @returns {{ content: string; name: string; path: string } | null}
 */
function readMarkdownFile(filePath) {
  try {
    assertFileSizeAllowed(filePath);
    const buffer = fs.readFileSync(filePath);
    const content = decodeTextBuffer(buffer);
    const name = path.basename(filePath);
    return { content, name, path: filePath };
  } catch (err) {
    if (err instanceof FileTooLargeError) {
      throw err;
    }
    return null;
  }
}

// NEW-01 — a file made outside Atlas (e.g. Windows Explorer's "New > Word
// Document" on a PC without Office installed) is a genuine 0-byte file on
// disk; Atlas's own parsers for the zip-based binary formats (docx/xlsx/
// ods/pptx/odp) throw on empty input instead of showing a blank document.
// Every binary-class read (`dialog:openFileBinary`, `file:readBinaryByPath`)
// funnels through this: a 0-byte read of one of those formats is
// transparently replaced with that format's blank template bytes, so the
// viewer sees a normal, fully-editable blank document — still bound to the
// file's own (still-empty-on-disk) path, so a normal Save writes a real
// document there. Every other format (pdf/rtf/odt/doc/ppt/unknown) is left
// genuinely empty; the renderer shows a friendly "this file is empty"
// message instead of a parser crash (see App.tsx).
/**
 * @param {string} filePath
 * @param {Buffer} buf
 * @returns {Buffer}
 */
function substituteBlankTemplateIfEmpty(filePath, buf) {
  if (buf.byteLength !== 0) return buf;
  return templateBytesForExtension(filePath) ?? buf;
}

/**
 * @param {BrowserWindow | null} win
 * @param {string | null} filePath
 */
function sendFileToWindow(win, filePath) {
  if (!win || win.isDestroyed()) return;
  if (typeof filePath !== 'string' || filePath.length === 0) return;
  // Always send the new path-only event for all formats
  win.webContents.send('file-opened-path', filePath);
  // Legacy markdown event — keep for regression-free fallback
  const dot = filePath.lastIndexOf('.');
  const ext = dot >= 0 ? filePath.slice(dot + 1).toLowerCase() : '';
  if (ext === 'md' || ext === 'markdown') {
    let fileData = null;
    try {
      fileData = readMarkdownFile(filePath);
    } catch (err) {
      logMainEvent('ERROR', 'legacy file-opened read failed', err);
    }
    if (fileData) {
      win.webContents.send('file-opened', fileData);
    }
  }
  win.focus();
}

/**
 * Registers a drag-dropped `filePath` into the SESSION allowlist only, after
 * re-validating it still exists on disk (P1.2/P1.4). This does not add to
 * the persisted `recentFilesStore` — see that module's header for why.
 * @param {Electron.IpcMainInvokeEvent} event
 * @param {unknown} filePath
 */
function registerDroppedPath(event, filePath) {
  if (!isFromMainFrame(event)) return { ok: false };
  if (typeof filePath !== 'string' || filePath.length === 0) return { ok: false };
  if (!fs.existsSync(filePath)) return { ok: false };
  pathAllowlist.add(filePath);
  return { ok: true };
}

/**
 * Re-opens a path from the renderer's "recent files" list. Unlike drag-drop
 * registration, this additionally requires `filePath` to already be present
 * in the persisted `recentFilesStore` — i.e. to have been genuinely opened
 * or saved through a trusted flow at some point — so a script that can only
 * write to `localStorage` (or call this IPC channel directly) cannot claim
 * an arbitrary existing file is "recent" (security-review fix; see
 * recentFilesStore.cjs).
 * @param {Electron.IpcMainInvokeEvent} event
 * @param {unknown} filePath
 */
function registerRecentPath(event, filePath) {
  if (!isFromMainFrame(event)) return { ok: false };
  if (typeof filePath !== 'string' || filePath.length === 0) return { ok: false };
  if (!getRecentFilesStore().has(filePath)) return { ok: false };
  if (!fs.existsSync(filePath)) return { ok: false };
  trustPath(filePath);
  return { ok: true };
}

// ---- Content-Security-Policy (P1.2 / ELEC-04) ---- //

/** @param {BrowserWindow} win */
function applyContentSecurityPolicy(win) {
  win.webContents.session.webRequest.onHeadersReceived((details, callback) => {
    if (details.resourceType !== 'mainFrame') {
      callback({ responseHeaders: details.responseHeaders });
      return;
    }
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [buildContentSecurityPolicy(isDev)],
      },
    });
  });
}

// ---- Navigation guards (P1.2 / ELEC-04) ---- //

/**
 * @param {string} urlString
 * @returns {boolean}
 */
function isAllowedExternalScheme(urlString) {
  try {
    const { protocol } = new URL(urlString);
    return protocol === 'http:' || protocol === 'https:' || protocol === 'mailto:';
  } catch {
    return false;
  }
}

/** @param {BrowserWindow} win */
function applyNavigationGuards(win) {
  win.webContents.on('will-navigate', (event, url) => {
    // A same-URL "navigation" is a reload (Ctrl+R, Vite HMR's full-reload
    // fallback, or the P1.15 recovery dialog's win.reload()) — allow it.
    // Anything else is a genuine navigation attempt and must be blocked.
    if (url === win.webContents.getURL()) {
      return;
    }
    event.preventDefault();
    if (isAllowedExternalScheme(url)) {
      void shell.openExternal(url);
    }
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isAllowedExternalScheme(url)) {
      void shell.openExternal(url);
    }
    return { action: 'deny' };
  });
}

// ---- Production application menu (P1.2 / ELEC-12) ---- //

function buildProductionMenu() {
  /** @type {Electron.MenuItemConstructorOptions[]} */
  const template = [
    {
      label: 'File',
      submenu: [{ role: 'quit' }],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
    {
      label: 'Window',
      submenu: [{ role: 'minimize' }, { role: 'close' }],
    },
  ];
  return Menu.buildFromTemplate(template);
}

// ---- Crash / error handling (P1.15 / ELEC-13) ---- //

/**
 * @param {BrowserWindow} win
 */
function showRecoveryDialog(win) {
  if (!win || win.isDestroyed()) return;
  const choice = dialog.showMessageBoxSync(win, {
    type: 'error',
    buttons: ['Reload', 'Quit'],
    defaultId: 0,
    cancelId: 1,
    title: 'Atlas has stopped responding',
    message: 'The document view has crashed or stopped responding.',
    detail: 'You can reload the window to continue, or quit the app.',
  });
  if (choice === 0) {
    win.reload();
  } else {
    app.quit();
  }
}

/** @param {BrowserWindow} win */
function applyCrashHandlers(win) {
  win.webContents.on('render-process-gone', (_event, details) => {
    logMainEvent('ERROR', `renderer process gone (${details.reason})`, details);
    showRecoveryDialog(win);
  });
  win.webContents.on('unresponsive', () => {
    logMainEvent('WARN', 'renderer unresponsive');
    showRecoveryDialog(win);
  });
}

// The native Save/Discard/Cancel dialog is modal to `mainWindow`, so a
// second close attempt can't realistically land while it's on screen — but
// the *async* gap after the user picks "Save" (waiting on the renderer's own
// save round-trip to finish) has no such protection: the window still looks
// interactive, so an impatient second Alt+F4/titlebar-X click reaches
// `handleWindowCloseRequest` again while `rendererDirty` is still `true`.
// Without a guard that re-entrant call would show a second stacked dialog
// and register a second `ipcMain.once('save-before-close-result', ...)`
// listener for the same in-flight save. This flag makes a repeated close
// attempt while a round-trip is already pending a no-op instead — it just
// waits on the original attempt's outcome rather than starting a new one.
// (Trade-off: if the renderer's save never responds at all — e.g. a hung
// renderer — the window can no longer be force-closed via a second attempt
// choosing Discard, the way the old code allowed. A hung renderer already
// needs a harder recovery path than repeated close-clicking, so this is
// accepted rather than reintroducing the double-dialog bug.)
/** @type {boolean} */
let closeConfirmationInFlight = false;

/**
 * P2.5/SHELL-02/ELEC-06 — blocks the native window close while the renderer
 * has reported unsaved changes (`rendererDirty`, kept current by the
 * `renderer:dirty-state` IPC signal below), showing a native Save/Discard/
 * Cancel prompt. "Save" round-trips through the renderer via IPC
 * (`request-save-before-close` / `save-before-close-result`) — driving
 * whichever save the active document-session contract (P1.1) has
 * registered, exactly like Ctrl+S/the Save button would — and only actually
 * closes the window once that save reports success, so a failed save never
 * silently loses the user's only warning.
 *
 * @param {Electron.Event} event
 */
function handleWindowCloseRequest(event) {
  if (!mainWindow) return;
  if (decideOnClose(rendererDirty) === 'allow') return;

  event.preventDefault();

  if (closeConfirmationInFlight) return;
  closeConfirmationInFlight = true;

  const choice = dialog.showMessageBoxSync(mainWindow, {
    type: 'warning',
    buttons: [...CLOSE_PROMPT_BUTTONS],
    defaultId: 0,
    cancelId: 2,
    title: 'Unsaved changes',
    message: 'This document has unsaved changes.',
    detail: 'Do you want to save your changes before closing?',
  });

  const action = decideAfterPromptChoice(choice);

  if (action === 'cancel') {
    closeConfirmationInFlight = false;
    return;
  }

  if (action === 'discard') {
    rendererDirty = false;
    closeConfirmationInFlight = false;
    mainWindow.destroy();
    return;
  }

  // action === 'save' — the in-flight guard above ensures at most one of
  // these `once` registrations is ever pending at a time.
  ipcMain.once('save-before-close-result', (_event, result) => {
    closeConfirmationInFlight = false;
    if (result && result.saved) {
      rendererDirty = false;
      if (mainWindow) mainWindow.destroy();
    }
    // A failed/cancelled save leaves the window open — the user can retry
    // closing (or saving directly) once they've addressed why it failed.
  });
  mainWindow.webContents.send('request-save-before-close');
}

ipcMain.on('renderer:dirty-state', (event, dirty) => {
  if (!isFromMainFrame(event)) return;
  rendererDirty = dirty === true;
});

process.on('uncaughtException', (error) => {
  logMainEvent('ERROR', 'uncaughtException', error);
});

process.on('unhandledRejection', (reason) => {
  logMainEvent('ERROR', 'unhandledRejection', reason);
});

// ---- Window state persistence (P5.2 / ELEC-11) ---- //
//
// Restores the window's size/position/maximized state from the previous
// launch, clamped against the displays actually connected right now (a
// saved position from a monitor that's since been unplugged or resized
// would otherwise land the window fully off-screen with no way to drag it
// back — see windowState.cjs's `clampBoundsToDisplays`). Saved on every
// resize/move (debounced) and on every close attempt, using
// `getNormalBounds()` so a maximized window's persisted size/position is
// its restored (non-maximized) geometry, not the full-screen bounds.

const DEFAULT_WINDOW_WIDTH = 1200;
const DEFAULT_WINDOW_HEIGHT = 800;
/** Debounce for the resize/move listeners below — these can fire many times
 * a second during a drag; there's no need to hit disk on every one. */
const WINDOW_STATE_SAVE_DEBOUNCE_MS = 300;

/**
 * @returns {{ bounds: import('./lib/windowState.cjs').WindowBounds | null, isMaximized: boolean }}
 */
function resolveInitialWindowState() {
  const saved = loadWindowState(getUserDataDir());
  if (!saved) return { bounds: null, isMaximized: false };
  const { x, y, width, height, isMaximized } = saved;
  const bounds = clampBoundsToDisplays({ x, y, width, height }, screen.getAllDisplays());
  return { bounds, isMaximized };
}

/** @type {ReturnType<typeof setTimeout> | null} */
let windowStateSaveTimer = null;

function persistWindowState() {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  try {
    const bounds = mainWindow.getNormalBounds();
    saveWindowState(getUserDataDir(), { ...bounds, isMaximized: mainWindow.isMaximized() });
  } catch (err) {
    logMainEvent('WARN', 'persistWindowState failed', err);
  }
}

function schedulePersistWindowState() {
  if (windowStateSaveTimer) clearTimeout(windowStateSaveTimer);
  windowStateSaveTimer = setTimeout(persistWindowState, WINDOW_STATE_SAVE_DEBOUNCE_MS);
}

// ---- Window creation ---- //

function createWindow() {
  rendererReady = false;
  const initialState = resolveInitialWindowState();

  mainWindow = new BrowserWindow({
    width: initialState.bounds?.width ?? DEFAULT_WINDOW_WIDTH,
    height: initialState.bounds?.height ?? DEFAULT_WINDOW_HEIGHT,
    ...(initialState.bounds ? { x: initialState.bounds.x, y: initialState.bounds.y } : {}),
    minWidth: 600,
    minHeight: 400,
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: currentOverlayColors.color,
      symbolColor: currentOverlayColors.symbolColor,
      height: 52,
    },
    backgroundColor: '#ffffff',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });

  if (initialState.isMaximized) {
    mainWindow.maximize();
  }

  mainWindow.on('resize', schedulePersistWindowState);
  mainWindow.on('move', schedulePersistWindowState);
  // A final, immediate (non-debounced) save on every close attempt — the
  // window hasn't moved just because a close was requested, so this simply
  // guarantees the last-known state is flushed even if the debounced timer
  // above hasn't fired yet when the process actually exits.
  mainWindow.on('close', persistWindowState);

  applyContentSecurityPolicy(mainWindow);
  applyNavigationGuards(mainWindow);
  applyCrashHandlers(mainWindow);

  if (!isDev) {
    Menu.setApplicationMenu(buildProductionMenu());
  }

  // Configure spellchecker languages — default to system locale + en-US fallback.
  try {
    const session = mainWindow.webContents.session;
    const available = session.availableSpellCheckerLanguages || [];
    const preferred = [app.getLocale(), 'en-US'].filter(
      (lang, idx, arr) => typeof lang === 'string' && lang.length > 0 && arr.indexOf(lang) === idx,
    );
    const enabled = preferred.filter((lang) => available.includes(lang));
    if (enabled.length > 0) {
      session.setSpellCheckerLanguages(enabled);
    }
  } catch (err) {
    console.error('spellcheck setup failed', err);
  }

  // Forward context-menu events containing misspelled words to the renderer
  // so the React SpellCheckMenu can render a custom menu.
  mainWindow.webContents.on('context-menu', (_event, params) => {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    if (typeof params.misspelledWord !== 'string' || params.misspelledWord.length === 0) {
      return;
    }
    mainWindow.webContents.send('spellcheck:show-menu', {
      word: params.misspelledWord,
      suggestions: Array.isArray(params.dictionarySuggestions) ? params.dictionarySuggestions : [],
      x: params.x,
      y: params.y,
    });
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(DIST_INDEX_PATH);
  }

  // Fallback: show window after 5s even if ready-to-show never fires (e.g. loadFile failure)
  const showTimeout = setTimeout(() => {
    if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isVisible()) {
      mainWindow.show();
    }
  }, 5000);

  mainWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    console.error(`Failed to load: ${errorCode} - ${errorDescription}`);
    // Show the window so user sees something rather than a phantom process
    if (mainWindow && !mainWindow.isDestroyed() && !mainWindow.isVisible()) {
      mainWindow.show();
    }
  });

  mainWindow.once('ready-to-show', () => {
    clearTimeout(showTimeout);
    // `mainWindow` is a mutable module-level binding, so TS can't carry the
    // non-null narrowing from `createWindow`'s synchronous body into this
    // async callback — and in principle the window could have been closed
    // (`closed` sets it back to `null`) in the gap before this event fires.
    // Same defensive check the sibling `did-fail-load`/fallback-show
    // handlers above already use.
    if (!mainWindow || mainWindow.isDestroyed()) return;
    mainWindow.show();
    // Send pending file after window is ready
    if (pendingFilePath) {
      // Small delay to ensure renderer is fully initialized
      setTimeout(() => {
        sendFileToWindow(mainWindow, pendingFilePath);
        pendingFilePath = null;
      }, 300);
    }
  });

  mainWindow.on('close', handleWindowCloseRequest);

  mainWindow.on('closed', () => {
    mainWindow = null;
    if (windowStateSaveTimer) {
      clearTimeout(windowStateSaveTimer);
      windowStateSaveTimer = null;
    }
  });

  if (isDev && process.env.PLAYWRIGHT !== '1') {
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  }
}

// ---- IPC Handlers ---- //

ipcMain.handle('get-initial-file', () => {
  // P5.2/ELEC-07 — by construction the renderer registers its
  // `file-opened-path` listener synchronously before ever invoking this (see
  // useFileHandler.ts's boot-subscriptions effect), so this invoke's arrival
  // is a precise "the listener now exists" signal for decideFileOpenAction.
  rendererReady = true;
  if (pendingFilePath) {
    const filePath = pendingFilePath;
    pendingFilePath = null;
    return { path: filePath };
  }
  return null;
});

ipcMain.handle('dialog:openFileBinary', async (event) => {
  if (!isFromMainFrame(event)) {
    return { canceled: true, path: '', buffer: new ArrayBuffer(0) };
  }

  const win = mainWindow;
  /** @type {Electron.OpenDialogOptions} */
  const openDialogOptions = {
    properties: ['openFile'],
    filters: [
      { name: 'Supported Files', extensions: MANIFEST_EXTENSIONS },
      { name: 'All Files', extensions: ['*'] },
    ],
  };
  // `showOpenDialog` has distinct overloads for "with owner window" and
  // "without" — passing `win || undefined` doesn't cleanly match either, so
  // branch explicitly instead of forcing one overload with a cast.
  const result = win
    ? await dialog.showOpenDialog(win, openDialogOptions)
    : await dialog.showOpenDialog(openDialogOptions);

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true, path: '', buffer: new ArrayBuffer(0) };
  }

  const filePath = result.filePaths[0];

  try {
    assertFileSizeAllowed(filePath);
  } catch (err) {
    dialog.showErrorBox('File too large', err instanceof Error ? err.message : String(err));
    return { canceled: true, path: '', buffer: new ArrayBuffer(0) };
  }

  trustPath(filePath);
  const buf = substituteBlankTemplateIfEmpty(filePath, await fs.promises.readFile(filePath));
  const buffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return { canceled: false, path: filePath, buffer };
});

ipcMain.handle('file:readBinaryByPath', async (event, filePath) => {
  if (!isFromMainFrame(event)) {
    throw new Error(SENDER_FRAME_ERROR_MESSAGE);
  }
  if (typeof filePath !== 'string' || !pathAllowlist.has(filePath)) {
    throw new Error(NOT_ALLOWLISTED_MESSAGE);
  }
  if (!fs.existsSync(filePath)) {
    throw new Error(FILE_NOT_FOUND_MESSAGE);
  }
  if (fs.statSync(filePath).isDirectory()) {
    throw new Error(IS_A_FOLDER_MESSAGE);
  }
  assertFileSizeAllowed(filePath);
  const buf = substituteBlankTemplateIfEmpty(filePath, await fs.promises.readFile(filePath));
  const buffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return { path: filePath, buffer };
});

ipcMain.handle('open-file-dialog', async (event) => {
  if (!isFromMainFrame(event)) return null;
  if (!mainWindow) return null;

  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Markdown Files', extensions: ['md', 'markdown', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) return null;

  trustPath(result.filePaths[0]);
  return readMarkdownFile(result.filePaths[0]);
});

ipcMain.handle('open-file-by-path', (event, filePath) => {
  if (!isFromMainFrame(event)) {
    throw new Error(SENDER_FRAME_ERROR_MESSAGE);
  }
  if (typeof filePath !== 'string') return null;
  if (!pathAllowlist.has(filePath)) {
    throw new Error(NOT_ALLOWLISTED_MESSAGE);
  }
  if (!fs.existsSync(filePath)) return null;
  return readMarkdownFile(filePath);
});

ipcMain.handle('path:register-dropped', (event, filePath) => registerDroppedPath(event, filePath));

ipcMain.handle('recent:request-open', (event, filePath) => registerRecentPath(event, filePath));

ipcMain.handle('save-file', async (event, req) => {
  if (!isFromMainFrame(event)) return { saved: false, error: SENDER_FRAME_ERROR_MESSAGE };
  if (!mainWindow) return { saved: false };
  if (!req || typeof req.content !== 'string') return { saved: false };

  // Only silently overwrite a path this window is already vouched for
  // (ELEC-03) — anything else falls back to the save dialog instead of
  // failing outright.
  let targetPath =
    typeof req.existingPath === 'string' && pathAllowlist.has(req.existingPath)
      ? req.existingPath
      : undefined;

  if (!targetPath) {
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: req.suggestedName || 'document.md',
      filters: req.filters && req.filters.length > 0 ? req.filters : [
        { name: 'Markdown', extensions: ['md'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });
    if (result.canceled || !result.filePath) return { saved: false };
    targetPath = result.filePath;
  }

  try {
    atomicWriteFile(targetPath, req.content);
    trustPath(targetPath);
    return { saved: true, path: targetPath, name: path.basename(targetPath) };
  } catch (err) {
    logMainEvent('ERROR', 'save-file failed', err);
    return {
      saved: false,
      error: err instanceof FileLockedError ? err.message : classifyWriteError(err),
      errorCode: classifyWriteErrorCode(err),
    };
  }
});

ipcMain.handle('save-binary-file', async (event, req) => {
  if (!isFromMainFrame(event)) return { saved: false, error: SENDER_FRAME_ERROR_MESSAGE };
  if (!mainWindow) return { saved: false };
  if (!(req && req.content instanceof Uint8Array)) return { saved: false };

  let targetPath =
    typeof req.existingPath === 'string' && pathAllowlist.has(req.existingPath)
      ? req.existingPath
      : undefined;

  if (!targetPath) {
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: req.suggestedName || 'document.docx',
      filters: req.filters && req.filters.length > 0 ? req.filters : [
        { name: 'Word Documents', extensions: ['docx'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    });
    if (result.canceled || !result.filePath) return { saved: false };
    targetPath = result.filePath;
  }

  try {
    atomicWriteFile(targetPath, Buffer.from(req.content));
    trustPath(targetPath);
    return { saved: true, path: targetPath, name: path.basename(targetPath) };
  } catch (err) {
    logMainEvent('ERROR', 'save-binary-file failed', err);
    return {
      saved: false,
      error: err instanceof FileLockedError ? err.message : classifyWriteError(err),
      errorCode: classifyWriteErrorCode(err),
    };
  }
});

// NEW-01 — creates a brand-new document from the toolbar's "New" action /
// Ctrl+N. `formatId` is validated against the fixed `NEW_DOCUMENT_FORMATS`
// table (never a renderer-supplied path or template bytes); the renderer
// only ever picks a name from that same fixed list (see NewDocumentMenu.tsx)
// and the save destination always comes back from a native `showSaveDialog`
// this handler itself drives, matching every other write path's security
// model (ELEC-02/03) — the renderer never chooses a filesystem path.
ipcMain.handle('document:new', async (event, formatId) => {
  if (!isFromMainFrame(event)) return { created: false, error: SENDER_FRAME_ERROR_MESSAGE };
  if (!mainWindow) return { created: false };

  // Own keys only: a name like 'constructor' must not resolve to an inherited property.
  const spec =
    typeof formatId === 'string' && Object.hasOwn(NEW_DOCUMENT_FORMATS, formatId) ? NEW_DOCUMENT_FORMATS[formatId] : undefined;
  if (!spec) return { created: false, error: 'Unsupported document type.' };

  let defaultDir;
  try {
    defaultDir = app.getPath('documents');
  } catch {
    defaultDir = os.homedir();
  }

  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: path.join(defaultDir, spec.defaultName),
    filters: [
      { name: spec.filterName, extensions: [spec.extension] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });
  if (result.canceled || !result.filePath) return { created: false };

  try {
    const bytes = spec.templateFile ? readTemplateBytes(spec.templateFile) : '';
    atomicWriteFile(result.filePath, bytes);
    trustPath(result.filePath);
    return { created: true, path: result.filePath };
  } catch (err) {
    logMainEvent('ERROR', 'document:new failed', err);
    return {
      created: false,
      error: err instanceof FileLockedError ? err.message : (classifyWriteError(err) ?? 'Could not create the new document.'),
      errorCode: err instanceof FileLockedError ? 'fileLocked' : (classifyWriteErrorCode(err) ?? 'unknownNewDocument'),
    };
  }
});

ipcMain.handle('export:printToPdf', async (event, req) => {
  if (!isFromMainFrame(event)) return { ok: false, error: SENDER_FRAME_ERROR_MESSAGE };
  if (!req || typeof req.html !== 'string') {
    return { ok: false, error: 'Nothing to export.' };
  }

  try {
    const buffer = await printHtmlToPdfBuffer(req.html);
    // `new Uint8Array(buffer)` (buffer is a Node `Buffer`, itself a
    // `Uint8Array` subclass) copies the bytes into a plain typed array
    // rather than aliasing Node's (possibly pooled) underlying
    // `ArrayBuffer` — simplest way to hand back exactly this PDF's bytes.
    return { ok: true, bytes: new Uint8Array(buffer) };
  } catch (err) {
    logMainEvent('ERROR', 'export:printToPdf failed', err);
    const message =
      err instanceof PrintToPdfError
        ? err.message
        : 'Could not generate the PDF for export. Please try again.';
    return { ok: false, error: message };
  }
});

ipcMain.handle('spellcheck:add-word', (_event, word) => {
  if (typeof word !== 'string' || word.length === 0) return { added: false };
  if (!mainWindow || mainWindow.isDestroyed()) return { added: false };
  try {
    mainWindow.webContents.session.addWordToSpellCheckerDictionary(word);
    return { added: true };
  } catch (err) {
    console.error('spellcheck:add-word failed', err);
    return { added: false };
  }
});

ipcMain.handle('spellcheck:replace-misspelling', (_event, word) => {
  if (typeof word !== 'string' || word.length === 0) return { replaced: false };
  if (!mainWindow || mainWindow.isDestroyed()) return { replaced: false };
  try {
    mainWindow.webContents.replaceMisspelling(word);
    return { replaced: true };
  } catch (err) {
    console.error('spellcheck:replace-misspelling failed', err);
    return { replaced: false };
  }
});

ipcMain.handle('spellcheck:get-languages', () => {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return { available: [], enabled: [] };
  }
  try {
    const session = mainWindow.webContents.session;
    return {
      available: session.availableSpellCheckerLanguages || [],
      enabled: session.getSpellCheckerLanguages() || [],
    };
  } catch (err) {
    console.error('spellcheck:get-languages failed', err);
    return { available: [], enabled: [] };
  }
});

ipcMain.handle('spellcheck:set-languages', (_event, languages) => {
  if (!Array.isArray(languages)) return { ok: false };
  if (!mainWindow || mainWindow.isDestroyed()) return { ok: false };
  try {
    const session = mainWindow.webContents.session;
    const available = session.availableSpellCheckerLanguages || [];
    const filtered = languages.filter((lang) => typeof lang === 'string' && available.includes(lang));
    session.setSpellCheckerLanguages(filtered);
    return { ok: true, enabled: filtered };
  } catch (err) {
    console.error('spellcheck:set-languages failed', err);
    return { ok: false };
  }
});

// ELEC-20/P4.10 — inserted images are decoded and held in memory by the
// DOCX editor, so cap them well below the general file-read ceiling
// (fileSizeGuard's 200 MiB default) rather than letting someone insert a
// multi-hundred-megabyte "image" and hang the renderer.
const IMAGE_PICK_MAX_BYTES = 25 * 1024 * 1024; // 25 MiB

ipcMain.handle('fonts:list', async (event) => {
  if (!isFromMainFrame(event)) return [];
  return listSystemFontFamilies();
});

ipcMain.handle('image:pick', async (event) => {
  if (!isFromMainFrame(event)) return { cancelled: true };
  if (!mainWindow || mainWindow.isDestroyed()) return { cancelled: true };
  try {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: 'Insert image',
      properties: ['openFile'],
      filters: [
        { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif'] },
      ],
    });
    if (result.canceled || result.filePaths.length === 0) {
      return { cancelled: true };
    }
    const filePath = result.filePaths[0];

    try {
      assertFileSizeAllowed(filePath, IMAGE_PICK_MAX_BYTES);
    } catch (err) {
      if (err instanceof FileTooLargeError) {
        return { cancelled: true, error: err.message };
      }
      throw err;
    }

    const buffer = fs.readFileSync(filePath);
    const ext = path.extname(filePath).slice(1).toLowerCase();
    const mime =
      ext === 'png' ? 'image/png'
      : ext === 'gif' ? 'image/gif'
      : 'image/jpeg';
    return {
      cancelled: false,
      bytes: buffer,
      mime,
      suggestedName: path.basename(filePath),
    };
  } catch (err) {
    logMainEvent('ERROR', 'image:pick failed', err);
    return { cancelled: true, error: String(err) };
  }
});

// P2.11/LOAD-18 — lets UnknownViewer's "Reveal in folder" action work for a
// file Atlas couldn't otherwise open. Reuses the same path allowlist as
// every read/write handler so it can only reveal a path this window is
// already vouched for, not an arbitrary renderer-supplied string.
ipcMain.handle('shell:reveal-in-folder', (event, filePath) => {
  if (!isFromMainFrame(event)) return { ok: false };
  if (typeof filePath !== 'string' || !pathAllowlist.has(filePath)) return { ok: false };
  try {
    shell.showItemInFolder(filePath);
    return { ok: true };
  } catch (err) {
    logMainEvent('ERROR', 'shell:reveal-in-folder failed', err);
    return { ok: false };
  }
});

// USR-19 — explicit "Run" for an opened source file. Only allowlisted files
// run, the file on disk is what runs, and a native confirmation (which the
// renderer cannot answer) is required before running it. See
// electron/lib/codeRunner.cjs for the process sandboxing rules.
//
// The approval is tied to the file's CONTENT, not just its path: the child
// process reads the file from disk when it starts, and Atlas's own
// `save-file`/`save-binary-file` handlers will overwrite any allowlisted
// path with renderer-supplied bytes. Keying approval on the path alone would
// therefore let a compromised renderer overwrite an already-approved script
// and re-run it with no dialog at all. Each approval records the file's
// size+mtime and is re-checked immediately before every run, so any change —
// from Atlas, another editor, or an attacker — re-prompts.
/** @type {Map<string, string>} approved path -> file fingerprint at approval time */
const approvedRunPaths = new Map();

/**
 * `size:mtime` of the file, or null when it cannot be read.
 * @param {string} filePath
 * @returns {string | null}
 */
function runFingerprint(filePath) {
  try {
    const stats = fs.statSync(filePath);
    if (!stats.isFile()) return null;
    return `${stats.size}:${stats.mtimeMs}`;
  } catch {
    return null;
  }
}
/** @type {import('./lib/codeRunner.cjs').CodeRunner | null} */
let codeRunner = null;

/** @returns {import('./lib/codeRunner.cjs').CodeRunner} */
function getCodeRunner() {
  if (!codeRunner) {
    codeRunner = createCodeRunner({
      send: (channel, payload) => {
        if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send(channel, payload);
      },
    });
  }
  return codeRunner;
}

ipcMain.handle('code:run', async (event, filePath) => {
  if (!isFromMainFrame(event)) return { ok: false, error: SENDER_FRAME_ERROR_MESSAGE };
  if (typeof filePath !== 'string' || !pathAllowlist.has(filePath)) {
    return { ok: false, error: NOT_ALLOWLISTED_MESSAGE };
  }
  const runtime = runCodeRuntimeFor(filePath);
  if (!runtime) return { ok: false, error: 'Atlas cannot run this kind of file.' };
  if (!mainWindow || mainWindow.isDestroyed()) return { ok: false };

  const fingerprint = runFingerprint(filePath);
  if (fingerprint === null) return { ok: false, error: 'This file could not be read.' };

  if (approvedRunPaths.get(filePath) !== fingerprint) {
    const { response } = await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      buttons: ['Run', 'Cancel'],
      defaultId: 1,
      cancelId: 1,
      noLink: true,
      title: 'Run code',
      message: `Run ${path.basename(filePath)}?`,
      detail:
        `Atlas will run this file with ${runtime.language} in:\n${path.dirname(filePath)}\n\n` +
        'A program can read and change your files and use the network. Only run code you trust. ' +
        'It is stopped automatically after 60 seconds.',
    });
    if (response !== 0) return { ok: false, cancelled: true };
    // Re-read the fingerprint: the file may have changed while the dialog was open.
    const approvedFingerprint = runFingerprint(filePath);
    if (approvedFingerprint === null) return { ok: false, error: 'This file could not be read.' };
    approvedRunPaths.set(filePath, approvedFingerprint);
  }

  try {
    return getCodeRunner().start(filePath);
  } catch (err) {
    logMainEvent('ERROR', 'code:run failed', err);
    return { ok: false, error: 'The program could not be started.' };
  }
});

ipcMain.handle('code:stop', (event, runId) => {
  if (!isFromMainFrame(event) || typeof runId !== 'number') return false;
  return codeRunner ? codeRunner.stop(runId) : false;
});

app.on('will-quit', () => {
  if (codeRunner) codeRunner.dispose();
});

ipcMain.on('set-theme', (_event, theme) => {
  // `theme` arrives from the renderer over IPC — validate it against the
  // known keys instead of trusting/indexing an arbitrary value.
  currentOverlayColors = isKnownOverlayTheme(theme) ? OVERLAY_COLORS[theme] : OVERLAY_COLORS.light;
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.setTitleBarOverlay({
      color: currentOverlayColors.color,
      symbolColor: currentOverlayColors.symbolColor,
      height: 52,
    });
  }
});

// i18n — backs the renderer's "System" language option (`src/i18n`), which
// has no web-platform equivalent of `app.getLocale()`.
ipcMain.handle('app:get-locale', () => app.getLocale());

// ---- App lifecycle ---- //

// Grab file from argv before app is ready. `app.getPath('userData')` is not
// guaranteed to work this early, so only the session allowlist is updated
// here — the persisted recentFilesStore write is deferred to `whenReady`
// below (once `app.getPath` is safe to call) rather than risking pinning
// the store's lazily-created singleton to a tmpdir fallback for the rest of
// the session.
pendingFilePath = extractFilePath(process.argv);
if (pendingFilePath) {
  pathAllowlist.add(pendingFilePath);
}

app.on('second-instance', (_event, argv) => {
  const filePath = extractFilePath(argv);
  if (filePath) {
    trustPath(filePath);
    // P5.2/ELEC-07 — a rapid second launch can arrive either before
    // `mainWindow` exists at all, or after it exists but before its renderer
    // has mounted the listener that would receive `sendFileToWindow`'s
    // event; both cases must queue into `pendingFilePath` instead of
    // dropping the request (see fileOpenRouting.cjs for the full rationale).
    const action = decideFileOpenAction({ hasWindow: mainWindow != null, rendererReady });
    if (action === 'send') {
      sendFileToWindow(mainWindow, filePath);
      return;
    }
    pendingFilePath = filePath;
  }
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  }
});

app.whenReady().then(() => {
  createWindow();

  // Now safe to touch the persisted recent-files store — flush the
  // cold-start argv path (if any) into it. The session allowlist was
  // already updated synchronously above, before the app was ready.
  if (pendingFilePath) {
    getRecentFilesStore().record(pendingFilePath);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  app.quit();
});

// Handle open-file on macOS (also works for some Windows scenarios)
app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (typeof filePath !== 'string' || filePath.length === 0) return;
  if (app.isReady()) {
    trustPath(filePath);
  } else {
    // Mirrors the argv case above: `app.getPath('userData')` isn't
    // guaranteed to work before 'ready' (this event can fire pre-ready on
    // macOS cold start), so only the session allowlist is updated now; the
    // `whenReady` flush above will persist `pendingFilePath` once it's safe.
    pathAllowlist.add(filePath);
  }
  // P5.2/ELEC-07 — same window-exists-but-renderer-not-mounted-yet race as
  // the second-instance handler above; see fileOpenRouting.cjs.
  const action = decideFileOpenAction({ hasWindow: mainWindow != null, rendererReady });
  if (action === 'send') {
    sendFileToWindow(mainWindow, filePath);
  } else {
    pendingFilePath = filePath;
  }
});
