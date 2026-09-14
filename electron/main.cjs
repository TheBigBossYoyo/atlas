const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');

const { createPathAllowlist } = require('./lib/pathAllowlist.cjs');
const { createRecentFilesStore } = require('./lib/recentFilesStore.cjs');
const { atomicWriteFile, FileLockedError } = require('./lib/atomicWrite.cjs');
const { decodeTextBuffer } = require('./lib/textDecoding.cjs');
const { buildContentSecurityPolicy } = require('./lib/csp.cjs');
const { logToFile } = require('./lib/crashLog.cjs');
const { FileTooLargeError, assertFileSizeAllowed } = require('./lib/fileSizeGuard.cjs');
const { EXTENSIONS: MANIFEST_EXTENSIONS } = require('./lib/extensionManifest.generated.cjs');
const { CLOSE_PROMPT_BUTTONS, decideOnClose, decideAfterPromptChoice } = require('./lib/closeGuard.cjs');

// Single instance lock
const gotLock = app.requestSingleInstanceLock();

if (!gotLock) {
  app.quit();
}

/** @type {BrowserWindow | null} */
let mainWindow = null;

/** @type {string | null} */
let pendingFilePath = null;

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

const isDev = !app.isPackaged;

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
/** @type {import('./lib/recentFilesStore.cjs').RecentFilesStore | null} */
let recentFilesStoreInstance = null;
function getRecentFilesStore() {
  if (!recentFilesStoreInstance) {
    let storeDir;
    try {
      storeDir = app.getPath('userData');
    } catch {
      storeDir = os.tmpdir();
    }
    recentFilesStoreInstance = createRecentFilesStore(storeDir);
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

  if (action === 'cancel') return;

  if (action === 'discard') {
    rendererDirty = false;
    mainWindow.destroy();
    return;
  }

  // action === 'save' — exactly one close attempt can be in flight at a
  // time (the window is blocked on the dialog above until the user answers,
  // and this IPC round-trip until the renderer responds), so `once` is safe.
  ipcMain.once('save-before-close-result', (_event, result) => {
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

// ---- Window creation ---- //

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
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
    const indexPath = path.join(app.getAppPath(), 'dist', 'index.html');
    mainWindow.loadFile(indexPath);
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
  });

  if (isDev && process.env.PLAYWRIGHT !== '1') {
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  }
}

// ---- IPC Handlers ---- //

ipcMain.handle('get-initial-file', () => {
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
  const result = await dialog.showOpenDialog(win || undefined, {
    properties: ['openFile'],
    filters: [
      { name: 'Supported Files', extensions: MANIFEST_EXTENSIONS },
      { name: 'All Files', extensions: ['*'] },
    ],
  });

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
  const buf = await fs.promises.readFile(filePath);
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
    throw new Error('Invalid path');
  }
  assertFileSizeAllowed(filePath);
  const buf = await fs.promises.readFile(filePath);
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
    return { saved: false, error: err instanceof FileLockedError ? err.message : undefined };
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
    return { saved: false, error: err instanceof FileLockedError ? err.message : undefined };
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

ipcMain.on('set-theme', (_event, theme) => {  currentOverlayColors = OVERLAY_COLORS[theme] || OVERLAY_COLORS.light;
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.setTitleBarOverlay({
      color: currentOverlayColors.color,
      symbolColor: currentOverlayColors.symbolColor,
      height: 52,
    });
  }
});

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
  }
  if (filePath && mainWindow) {
    sendFileToWindow(mainWindow, filePath);
  } else if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  } else if (filePath) {
    // Window not created yet (a rapid second launch during startup) — queue
    // it so it opens once ready-to-show fires instead of being dropped
    // (ELEC-07).
    pendingFilePath = filePath;
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
  if (mainWindow) {
    sendFileToWindow(mainWindow, filePath);
  } else {
    pendingFilePath = filePath;
  }
});
