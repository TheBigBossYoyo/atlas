const { contextBridge, ipcRenderer, webUtils } = require('electron');

/** @typedef {'light' | 'dark' | 'sepia' | 'nord' | 'dracula'} Theme */
/** @typedef {{ content: string; name: string; path: string }} ElectronFileData */
/** @typedef {{ name: string; extensions: string[] }} DialogFilter */
/**
 * @typedef {{
 *   content: string;
 *   suggestedName: string;
 *   filters?: DialogFilter[];
 *   existingPath?: string;
 * }} SaveFileRequest
 */
/**
 * @typedef {{
 *   content: Uint8Array;
 *   suggestedName: string;
 *   filters?: DialogFilter[];
 *   existingPath?: string;
 * }} BinarySaveFileRequest
 */
/** @typedef {{ runId: number; stream: 'stdout' | 'stderr' | 'system'; text: string }} CodeRunOutputPayload */
/**
 * @typedef {{
 *   runId: number;
 *   code: number | null;
 *   timedOut: boolean;
 *   stopped: boolean;
 *   error?: string;
 * }} CodeRunExitPayload
 */
/**
 * @typedef {{
 *   word: string;
 *   suggestions: readonly string[];
 *   x: number;
 *   y: number;
 * }} SpellCheckContextMenuPayload
 */

contextBridge.exposeInMainWorld('electronAPI', {
  getInitialFile: () => ipcRenderer.invoke('get-initial-file'),
  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
  /** @param {string} path */
  openFileByPath: (path) => ipcRenderer.invoke('open-file-by-path', path),
  /** @param {SaveFileRequest} req */
  saveFile: (req) => ipcRenderer.invoke('save-file', req),
  /** @param {BinarySaveFileRequest} req */
  saveBinaryFile: (req) => ipcRenderer.invoke('save-binary-file', req),
  // X1 — renders a self-contained (already-sanitized) HTML document to a
  // vector PDF in a hidden, script-disabled window. See
  // `electron/lib/printToPdf.cjs` for the full security rationale.
  /** @param {string} html */
  printToPdf: (html) => ipcRenderer.invoke('export:printToPdf', { html }),
  /** @param {(data: ElectronFileData) => void} callback */
  onFileOpened: (callback) => {
    /** @param {Electron.IpcRendererEvent} _event @param {ElectronFileData} data */
    const handler = (_event, data) => callback(data);
    ipcRenderer.on('file-opened', handler);
    return () => ipcRenderer.removeListener('file-opened', handler);
  },
  /** @param {Theme} theme */
  setTheme: (theme) => ipcRenderer.send('set-theme', theme),
  // i18n — the renderer's "System" language option resolves against
  // `app.getLocale()` (main-process-only); the renderer has no equivalent
  // web API, so this is a thin IPC round trip.
  getLocale: () => ipcRenderer.invoke('app:get-locale'),
  // P2.5/SHELL-02/ELEC-06 — pushes the renderer's combined dirty state to
  // main so its window `close` handler knows whether to block the close
  // behind a Save/Discard/Cancel prompt (Electron surfaces no visible
  // confirmation for a renderer-only `beforeunload` handler).
  /** @param {boolean} dirty */
  notifyDirtyState: (dirty) => ipcRenderer.send('renderer:dirty-state', dirty),
  /** @param {() => void} callback */
  onRequestSaveBeforeClose: (callback) => {
    const handler = () => callback();
    ipcRenderer.on('request-save-before-close', handler);
    return () => ipcRenderer.removeListener('request-save-before-close', handler);
  },
  /** @param {{ saved: boolean }} result */
  reportSaveBeforeCloseResult: (result) => ipcRenderer.send('save-before-close-result', result),
  // QUIT-DRAFT-1 — symmetric counterpart to the pair above, for Discard: main
  // asks the renderer to clear its autosave draft (see `src/hooks/
  // useAutosave.ts`'s `clearDraft`) before it destroys the window, and waits
  // (briefly — see `notifyRendererDiscardThenClose` in main.cjs) for this
  // acknowledgement so a fire-and-forget `send` can't race the teardown.
  /** @param {() => void} callback */
  onRequestDiscardBeforeClose: (callback) => {
    const handler = () => callback();
    ipcRenderer.on('request-discard-before-close', handler);
    return () => ipcRenderer.removeListener('request-discard-before-close', handler);
  },
  reportDiscardBeforeCloseResult: () => ipcRenderer.send('discard-before-close-result'),
  openFileBinary: () => ipcRenderer.invoke('dialog:openFileBinary'),
  // NEW-01 — creates a brand-new document (native Save dialog + a blank
  // template written atomically) for the toolbar's "New" action / Ctrl+N.
  /** @param {string} formatId */
  newDocument: (formatId) => ipcRenderer.invoke('document:new', formatId),
  /** @param {string} path */
  readBinaryByPath: (path) => ipcRenderer.invoke('file:readBinaryByPath', path),
  /** @param {(path: string) => void} callback */
  onFileOpenedPath: (callback) => {
    /** @param {Electron.IpcRendererEvent} _e @param {string} path */
    const handler = (_e, path) => callback(path);
    ipcRenderer.on('file-opened-path', handler);
    return () => ipcRenderer.removeListener('file-opened-path', handler);
  },
  // Electron 32+ removed `File.path` from dropped-file objects — webUtils is
  // the sandbox-compatible replacement (ELEC-01/SHELL-01/LOAD-02/RUN-04).
  /** @param {File} file */
  getPathForFile: (file) => webUtils.getPathForFile(file),
  // Registers a drag-dropped or recent-file path into the main process's
  // read/write allowlist after re-validating it still exists (P1.2/P1.4).
  /** @param {string} path */
  registerDroppedPath: (path) => ipcRenderer.invoke('path:register-dropped', path),
  /** @param {string} path */
  requestOpenRecent: (path) => ipcRenderer.invoke('recent:request-open', path),
  // Reveals an allowlisted path in the OS file manager (P2.11/LOAD-18 —
  // UnknownViewer's "Reveal in folder" action).
  /** @param {string} path */
  revealInFolder: (path) => ipcRenderer.invoke('shell:reveal-in-folder', path),
  image: {
    pick: () => ipcRenderer.invoke('image:pick'),
  },
  // USR-19 — explicit, confirmed "Run" of an opened source file (see electron/lib/codeRunner.cjs).
  codeRun: {
    /** @param {string} path */
    start: (path) => ipcRenderer.invoke('code:run', path),
    /** @param {number} runId */
    stop: (runId) => ipcRenderer.invoke('code:stop', runId),
    /** @param {(payload: CodeRunOutputPayload) => void} callback */
    onOutput: (callback) => {
      /** @param {Electron.IpcRendererEvent} _e @param {CodeRunOutputPayload} payload */
      const handler = (_e, payload) => callback(payload);
      ipcRenderer.on('code:run-output', handler);
      return () => ipcRenderer.removeListener('code:run-output', handler);
    },
    /** @param {(payload: CodeRunExitPayload) => void} callback */
    onExit: (callback) => {
      /** @param {Electron.IpcRendererEvent} _e @param {CodeRunExitPayload} payload */
      const handler = (_e, payload) => callback(payload);
      ipcRenderer.on('code:run-exit', handler);
      return () => ipcRenderer.removeListener('code:run-exit', handler);
    },
  },
  // USR-11 — installed font families for the DOCX font picker.
  fonts: {
    list: () => ipcRenderer.invoke('fonts:list'),
  },
  spellcheck: {
    /** @param {(payload: SpellCheckContextMenuPayload) => void} callback */
    onContextMenu: (callback) => {
      /** @param {Electron.IpcRendererEvent} _e @param {SpellCheckContextMenuPayload} payload */
      const handler = (_e, payload) => callback(payload);
      ipcRenderer.on('spellcheck:show-menu', handler);
      return () => ipcRenderer.removeListener('spellcheck:show-menu', handler);
    },
    /** @param {string} word */
    replaceMisspelling: (word) => ipcRenderer.invoke('spellcheck:replace-misspelling', word),
    /** @param {string} word */
    addWord: (word) => ipcRenderer.invoke('spellcheck:add-word', word),
    getLanguages: () => ipcRenderer.invoke('spellcheck:get-languages'),
    /** @param {readonly string[]} languages */
    setLanguages: (languages) => ipcRenderer.invoke('spellcheck:set-languages', languages),
  },
});
