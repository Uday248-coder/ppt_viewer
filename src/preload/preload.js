'use strict';

const { contextBridge, ipcRenderer } = require('electron');

/**
 * The renderer's entire view of the main process. Every call is an explicit,
 * named operation - no generic `invoke(channel, ...)` escape hatch.
 */
contextBridge.exposeInMainWorld('pptv', {
  app: {
    info: () => ipcRenderer.invoke('app:info'),
    setPref: (key, value) => ipcRenderer.invoke('prefs:set', { key, value }),
    onEngineStatus: (cb) => {
      const h = (_e, s) => cb(s);
      ipcRenderer.on('engine:status', h);
      return () => ipcRenderer.removeListener('engine:status', h);
    },
  },

  library: {
    scan: (roots) => ipcRenderer.invoke('library:scan', roots),
    chooseRoot: () => ipcRenderer.invoke('library:chooseRoot'),
    addRoot: (dir) => ipcRenderer.invoke('library:addRoot', dir),
    removeRoot: (dir) => ipcRenderer.invoke('library:removeRoot', dir),
    reveal: (target) => ipcRenderer.invoke('library:reveal', target),
    prune: () => ipcRenderer.invoke('library:prune'),
    clearCache: () => ipcRenderer.invoke('cache:clear'),
    clearRecents: () => ipcRenderer.invoke('recents:clear'),
  },

  deck: {
    open: (path) => ipcRenderer.invoke('deck:open', path),
    ensureThumbs: (path) => ipcRenderer.invoke('deck:ensureThumbs', path),
    ensureSlides: (path, indices) => ipcRenderer.invoke('deck:ensureSlides', { path, indices }),
    cacheAll: (path) => ipcRenderer.invoke('deck:cacheAll', path),
    status: (path) => ipcRenderer.invoke('deck:status', path),
    exportPdf: (path) => ipcRenderer.invoke('deck:exportPdf', path),
    onOpening: (cb) => {
      const h = (_e, d) => cb(d);
      ipcRenderer.on('deck:opening', h);
      return () => ipcRenderer.removeListener('deck:opening', h);
    },
    onProgress: (cb) => {
      const h = (_e, d) => cb(d);
      ipcRenderer.on('deck:progress', h);
      return () => ipcRenderer.removeListener('deck:progress', h);
    },
  },

  window: {
    present: () => ipcRenderer.invoke('win:present'),
    exitPresent: () => ipcRenderer.invoke('win:exitPresent'),
    isFullScreen: () => ipcRenderer.invoke('win:isFullScreen'),
  },
});
