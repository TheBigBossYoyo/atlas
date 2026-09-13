const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getInitialFile: () => ipcRenderer.invoke('get-initial-file'),
  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
  openFileByPath: (path) => ipcRenderer.invoke('open-file-by-path', path),
  saveFile: (req) => ipcRenderer.invoke('save-file', req),
  saveBinaryFile: (req) => ipcRenderer.invoke('save-binary-file', req),
  onFileOpened: (callback) => {
    const handler = (_event, data) => callback(data);
    ipcRenderer.on('file-opened', handler);
    return () => ipcRenderer.removeListener('file-opened', handler);
  },
  setTheme: (theme) => ipcRenderer.send('set-theme', theme),
  openFileBinary: () => ipcRenderer.invoke('dialog:openFileBinary'),
  readBinaryByPath: (path) => ipcRenderer.invoke('file:readBinaryByPath', path),
  onFileOpenedPath: (callback) => {
    const handler = (_e, path) => callback(path);
    ipcRenderer.on('file-opened-path', handler);
    return () => ipcRenderer.removeAllListeners('file-opened-path');
  },
  onSpellCheckMenu: (callback) => {
    const handler = (_event, payload) => callback(payload);
    ipcRenderer.on('spellcheck:show-menu', handler);
    return () => ipcRenderer.removeListener('spellcheck:show-menu', handler);
  },
  replaceMisspelling: (word) => ipcRenderer.invoke('spellcheck:replace-misspelling', word),
  addWordToDictionary: (word) => ipcRenderer.invoke('spellcheck:add-word', word),
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
