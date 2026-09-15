const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getInitialFile: () => ipcRenderer.invoke('get-initial-file'),
  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
  openFileByPath: (path) => ipcRenderer.invoke('open-file-by-path', path),
  saveFile: (req) => ipcRenderer.invoke('save-file', req),
  saveBinaryFile: (req) => ipcRenderer.invoke('save-binary-file', req),
  // X1 — renders a self-contained (already-sanitized) HTML document to a
  // vector PDF in a hidden, script-disabled window. See
  // `electron/lib/printToPdf.cjs` for the full security rationale.
  printToPdf: (html) => ipcRenderer.invoke('export:printToPdf', { html }),
  onFileOpened: (callback) => {
    const handler = (_event, data) => callback(data);
    ipcRenderer.on('file-opened', handler);
    return () => ipcRenderer.removeListener('file-opened', handler);
  },
  setTheme: (theme) => ipcRenderer.send('set-theme', theme),
  // P2.5/SHELL-02/ELEC-06 — pushes the renderer's combined dirty state to
  // main so its window `close` handler knows whether to block the close
  // behind a Save/Discard/Cancel prompt (Electron surfaces no visible
  // confirmation for a renderer-only `beforeunload` handler).
  notifyDirtyState: (dirty) => ipcRenderer.send('renderer:dirty-state', dirty),
  onRequestSaveBeforeClose: (callback) => {
    const handler = () => callback();
    ipcRenderer.on('request-save-before-close', handler);
    return () => ipcRenderer.removeListener('request-save-before-close', handler);
  },
  reportSaveBeforeCloseResult: (result) => ipcRenderer.send('save-before-close-result', result),
  openFileBinary: () => ipcRenderer.invoke('dialog:openFileBinary'),
  readBinaryByPath: (path) => ipcRenderer.invoke('file:readBinaryByPath', path),
  onFileOpenedPath: (callback) => {
    const handler = (_e, path) => callback(path);
    ipcRenderer.on('file-opened-path', handler);
    return () => ipcRenderer.removeListener('file-opened-path', handler);
  },
  // Electron 32+ removed `File.path` from dropped-file objects — webUtils is
  // the sandbox-compatible replacement (ELEC-01/SHELL-01/LOAD-02/RUN-04).
  getPathForFile: (file) => webUtils.getPathForFile(file),
  // Registers a drag-dropped or recent-file path into the main process's
  // read/write allowlist after re-validating it still exists (P1.2/P1.4).
  registerDroppedPath: (path) => ipcRenderer.invoke('path:register-dropped', path),
  requestOpenRecent: (path) => ipcRenderer.invoke('recent:request-open', path),
  // Reveals an allowlisted path in the OS file manager (P2.11/LOAD-18 —
  // UnknownViewer's "Reveal in folder" action).
  revealInFolder: (path) => ipcRenderer.invoke('shell:reveal-in-folder', path),
  image: {
    pick: () => ipcRenderer.invoke('image:pick'),
  },
  spellcheck: {
    onContextMenu: (callback) => {
      const handler = (_e, payload) => callback(payload);
      ipcRenderer.on('spellcheck:show-menu', handler);
      return () => ipcRenderer.removeListener('spellcheck:show-menu', handler);
    },
    replaceMisspelling: (word) => ipcRenderer.invoke('spellcheck:replace-misspelling', word),
    addWord: (word) => ipcRenderer.invoke('spellcheck:add-word', word),
    getLanguages: () => ipcRenderer.invoke('spellcheck:get-languages'),
    setLanguages: (languages) => ipcRenderer.invoke('spellcheck:set-languages', languages),
  },
});
