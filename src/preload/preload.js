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
    onNavigate: (cb) => {
      const h = (_e, s) => cb(s);
      ipcRenderer.on('deck:navigate', h);
      return () => ipcRenderer.removeListener('deck:navigate', h);
    },
    onResetTimer: (cb) => {
      const h = () => cb();
      ipcRenderer.on('presenter:reset-timer', h);
      return () => ipcRenderer.removeListener('presenter:reset-timer', h);
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
    prefill: (path) => ipcRenderer.invoke('deck:prefill', path),
    stopPrefill: () => ipcRenderer.invoke('deck:stopPrefill'),
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

  /**
   * The presenter window runs in its own renderer, so it gets its own copy of
   * this bridge. It may read state and ask the audience window to navigate, but
   * it deliberately has no path to the filesystem.
   */
  presenter: {
    open: () => ipcRenderer.invoke('presenter:open'),
    close: () => ipcRenderer.invoke('presenter:close'),
    isOpen: () => ipcRenderer.invoke('presenter:isOpen'),
    state: (s) => ipcRenderer.invoke('presenter:state', s),
    tick: (elapsedMs) => ipcRenderer.invoke('presenter:tick', { elapsedMs }),
    nav: (delta) => ipcRenderer.invoke('presenter:nav', delta),
    resetTimer: () => ipcRenderer.invoke('presenter:resetTimer'),
    onState: (cb) => {
      const h = (_e, s) => cb(s);
      ipcRenderer.on('presenter:state', h);
      return () => ipcRenderer.removeListener('presenter:state', h);
    },
    onTick: (cb) => {
      const h = (_e, s) => cb(s);
      ipcRenderer.on('presenter:tick', h);
      return () => ipcRenderer.removeListener('presenter:tick', h);
    },
    onExit: (cb) => {
      const h = () => cb();
      ipcRenderer.on('presenter:exit', h);
      return () => ipcRenderer.removeListener('presenter:exit', h);
    },
  },
});
