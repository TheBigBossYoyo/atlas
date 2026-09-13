const { app, BrowserWindow, ipcMain, dialog, protocol } = require('electron');
const path = require('path');
const fs = require('fs');

// Single instance lock
const gotLock = app.requestSingleInstanceLock();

if (!gotLock) {
  app.quit();
}

/** @type {BrowserWindow | null} */
let mainWindow = null;

/** @type {string | null} */
let pendingFilePath = null;

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

// ---- File path helpers ---- //

const KNOWN_EXTENSIONS = new Set([
  'md','markdown','mdx','mkd',
  'docx','xlsx','xlsm','xlsb','pptx','pptm',
  'pdf','csv','tsv','tab',
  'odt','ods','odp','rtf','txt','log',
  'js','jsx','ts','tsx','mjs','cjs',
  'py','rb','go','rs','java','kt','swift',
  'c','h','cpp','hpp','cs','php',
  'json','yaml','yml','toml','xml','html','css','scss',
  'sql','sh','bat','ps1',
]);

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

function readMarkdownFile(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    const name = path.basename(filePath);
    return { content, name, path: filePath };
  } catch {
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
    const fileData = readMarkdownFile(filePath);
    if (fileData) {
      win.webContents.send('file-opened', fileData);
    }
  }
  win.focus();
}

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
      sandbox: false,
      spellcheck: true,
    },
  });

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

ipcMain.handle('dialog:openFileBinary', async () => {
  const win = mainWindow;
  const result = await dialog.showOpenDialog(win || undefined, {
    properties: ['openFile'],
    filters: [
      {
        name: 'Supported Files',
        extensions: [
          'md','markdown','mdx','mkd',
          'docx','xlsx','xlsm','xlsb','pptx','pptm',
          'pdf','csv','tsv','tab',
          'odt','ods','odp','rtf','txt','log',
          'js','jsx','ts','tsx','mjs','cjs',
          'py','rb','go','rs','java','kt','swift',
          'c','h','cpp','hpp','cs','php',
          'json','yaml','yml','toml','xml','html','css','scss',
          'sql','sh','bat','ps1',
        ],
      },
      { name: 'All Files', extensions: ['*'] },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true, path: '', buffer: new ArrayBuffer(0) };
  }

  const filePath = result.filePaths[0];
  const buf = await fs.promises.readFile(filePath);
  const buffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return { canceled: false, path: filePath, buffer };
});

ipcMain.handle('file:readBinaryByPath', async (_event, filePath) => {
  if (typeof filePath !== 'string' || !fs.existsSync(filePath)) {
    throw new Error('Invalid path');
  }
  const buf = await fs.promises.readFile(filePath);
  const buffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return { path: filePath, buffer };
});

ipcMain.handle('open-file-dialog', async () => {
  if (!mainWindow) return null;

  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Markdown Files', extensions: ['md', 'markdown', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) return null;

  return readMarkdownFile(result.filePaths[0]);
});

ipcMain.handle('open-file-by-path', (_event, filePath) => {
  if (typeof filePath !== 'string') return null;
  if (!fs.existsSync(filePath)) return null;
  return readMarkdownFile(filePath);
});

ipcMain.handle('save-file', async (_event, req) => {
  if (!mainWindow) return { saved: false };
  if (!req || typeof req.content !== 'string') return { saved: false };

  let targetPath = req.existingPath;

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
    fs.writeFileSync(targetPath, req.content, 'utf-8');
    return { saved: true, path: targetPath, name: path.basename(targetPath) };
  } catch (err) {
    console.error('save-file failed', err);
    return { saved: false };
  }
});

ipcMain.handle('save-binary-file', async (_event, req) => {
  if (!mainWindow) return { saved: false };
  if (!(req && req.content instanceof Uint8Array)) return { saved: false };

  let targetPath = req.existingPath;

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
    fs.writeFileSync(targetPath, Buffer.from(req.content));
    return { saved: true, path: targetPath, name: path.basename(targetPath) };
  } catch (err) {
    console.error('save-binary-file failed', err);
    return { saved: false };
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

ipcMain.handle('image:pick', async () => {
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
    const fs = require('fs');
    const path = require('path');
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
    console.error('image:pick failed', err);
    return { cancelled: true, error: String(err) };
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

// Grab file from argv before app is ready
pendingFilePath = extractFilePath(process.argv);

app.on('second-instance', (_event, argv) => {
  const filePath = extractFilePath(argv);
  if (filePath && mainWindow) {
    sendFileToWindow(mainWindow, filePath);
  } else if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  }
});

app.whenReady().then(() => {
  createWindow();

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
  if (mainWindow) {
    sendFileToWindow(mainWindow, filePath);
  } else {
    pendingFilePath = filePath;
  }
});
