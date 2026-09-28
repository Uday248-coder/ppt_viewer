'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, protocol, net, nativeTheme } = require('electron');
const path = require('path');
const fsp = require('fs/promises');

const { ComBridge } = require('./com-bridge');
const { SlideCache, KINDS, THUMB_W, THUMB_H, FULL_W, FULL_H } = require('./cache');
const { LibraryScanner } = require('./library');
const { Settings } = require('./settings');
const { DeckService } = require('./deck-service');

const IMAGE_SCHEME = 'pptv';

// Cached slide images are served through a private protocol rather than file://
// so the renderer can only ever reach files inside the slide cache.
protocol.registerSchemesAsPrivileged([
  {
    scheme: IMAGE_SCHEME,
    privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true, bypassCSP: false },
  },
]);

let win = null;
let settings = null;
let cache = null;
let scanner = null;
let com = null;
let decks = null;
let engineState = { ready: false, comAvailable: false, version: null, error: null };

function slideUrl(key, kind, index) {
  return `${IMAGE_SCHEME}://slide/${key}/${kind}/slide-${String(index).padStart(4, '0')}.png`;
}

function deckPayload(res) {
  const info = res.info;
  const aspect = info.heightPt ? info.widthPt / info.heightPt : 4 / 3;
  return {
    path: res.path,
    key: res.key,
    fromCache: res.fromCache,
    name: info.name || path.basename(res.path),
    title: info.title || '',
    slideCount: info.slideCount,
    widthPt: info.widthPt,
    heightPt: info.heightPt,
    aspect,
    cached: res.cached ? res.cached.level : 'none',
    slides: (info.slides || []).map((s) => ({
      index: s.index,
      title: s.title || '',
      notes: s.notes || '',
      effect: s.effect,
      thumbUrl: slideUrl(res.key, KINDS.THUMB, s.index),
      fullUrl: slideUrl(res.key, KINDS.FULL, s.index),
    })),
  };
}

async function createWindow() {
  const w = settings.data.window;
  win = new BrowserWindow({
    width: w.width || 1440,
    height: w.height || 900,
    x: w.x ?? undefined,
    y: w.y ?? undefined,
    minWidth: 900,
    minHeight: 600,
    show: false,
    backgroundColor: '#0f1115',
    title: 'PPT Viewer',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false,
    },
  });

  if (w.maximized) win.maximize();

  win.once('ready-to-show', () => win.show());

  const persistBounds = () => {
    if (!win || win.isDestroyed()) return;
    const b = win.getNormalBounds();
    settings.data.window = { ...settings.data.window, ...b, maximized: win.isMaximized() };
    settings.save().catch(() => {});
  };
  win.on('resize', persistBounds);
  win.on('move', persistBounds);
  win.on('close', persistBounds);
  win.on('closed', () => {
    win = null;
  });

  await win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
}

function send(channel, payload) {
  if (win && !win.isDestroyed()) win.webContents.send(channel, payload);
}

function wireIpc() {
  ipcMain.handle('app:info', async () => {
    return {
      ...engineState,
      cacheDir: cache.root,
      cacheBytes: await cache.totalBytes(),
      roots: settings.data.libraryRoots,
      recents: settings.data.recents,
      prefs: settings.data.prefs,
      version: app.getVersion(),
      electron: process.versions.electron,
      node: process.versions.node,
    };
  });

  // --- library -----------------------------------------------------------
  ipcMain.handle('library:scan', async (_e, roots) => {
    const list = (Array.isArray(roots) && roots.length ? roots : settings.data.libraryRoots).filter(Boolean);
    const res = await scanner.scan(list);
    return {
      roots: list,
      ms: res.ms,
      errors: res.errors.slice(0, 20),
      cancelled: res.cancelled,
      decks: res.files.map((f) => ({
        ...f,
        key: SlideCache.keyFor(f.path, { size: f.size, mtimeMs: f.mtimeMs }),
      })),
    };
  });

  ipcMain.handle('library:addRoot', async (_e, dir) => {
    await settings.addRoot(dir);
    return settings.data.libraryRoots;
  });

  ipcMain.handle('library:removeRoot', async (_e, dir) => {
    await settings.removeRoot(dir);
    return settings.data.libraryRoots;
  });

  ipcMain.handle('library:chooseRoot', async () => {
    const r = await dialog.showOpenDialog(win, {
      title: 'Choose a folder to scan for presentations',
      properties: ['openDirectory', 'multiSelections', 'createDirectory'],
    });
    if (r.canceled || !r.filePaths.length) return { roots: settings.data.libraryRoots, added: [] };
    for (const p of r.filePaths) await settings.addRoot(p);
    return { roots: settings.data.libraryRoots, added: r.filePaths };
  });

  ipcMain.handle('library:reveal', async (_e, target) => {
    shell.showItemInFolder(target);
    return true;
  });

  ipcMain.handle('library:prune', async () => {
    const res = await scanner.scan(settings.data.libraryRoots, { batchSize: 10_000 });
    const live = res.files.map((d) => SlideCache.keyFor(d.path, { size: d.size, mtimeMs: d.mtimeMs }));
    return cache.prune(live);
  });

  ipcMain.handle('cache:clear', async () => {
    await cache.clearAll();
    return { ok: true, bytes: 0 };
  });

  // --- decks -------------------------------------------------------------
  ipcMain.handle('deck:open', async (_e, p) => {
    const res = await decks.openDeck(p);
    await settings.touchRecent({ path: res.path, name: path.basename(res.path), at: Date.now() });
    const payload = deckPayload(res);
    return payload;
  });

  ipcMain.handle('deck:ensureThumbs', async (_e, p) => decks.ensureThumbs(p));

  ipcMain.handle('deck:ensureSlides', async (_e, { path: p, indices }) => decks.ensureSlides(p, indices || []));

  ipcMain.handle('deck:cacheAll', async (_e, p) => {
    send('deck:progress', { path: p, phase: 'cache' });
    const r = await decks.cacheDeck(p);
    send('deck:progress', { path: p, phase: 'done', ...r });
    return r;
  });

  ipcMain.handle('deck:status', async (_e, p) => {
    const st = await fsp.stat(p);
    const key = SlideCache.keyFor(p, st);
    const status = await cache.status(key);
    return { key, level: status.level, slideCount: status.meta?.slideCount || 0 };
  });

  ipcMain.handle('deck:exportPdf', async (_e, p) => {
    const r = await dialog.showSaveDialog(win, {
      title: 'Export deck as PDF',
      defaultPath: path.join(app.getPath('documents'), `${path.basename(p, path.extname(p))}.pdf`),
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    });
    if (r.canceled || !r.filePath) return { cancelled: true };
    const out = await decks.exportPdf(p, r.filePath);
    return { cancelled: false, file: out.file };
  });

  ipcMain.handle('recents:clear', async () => {
    await settings.clearRecents();
    return [];
  });

  ipcMain.handle('prefs:set', async (_e, { key, value }) => {
    await settings.setPref(key, value);
    return settings.data.prefs;
  });

  // --- window ------------------------------------------------------------
  ipcMain.handle('win:present', async () => {
    if (!win) return false;
    win.setFullScreen(true);
    return true;
  });

  ipcMain.handle('win:exitPresent', async () => {
    if (!win) return false;
    win.setFullScreen(false);
    return true;
  });

  ipcMain.handle('win:isFullScreen', async () => (win ? win.isFullScreen() : false));

  ipcMain.handle('win:openExternal', async (_e, url) => {
    // Never navigate to arbitrary URLs; only hand http(s) to the OS browser.
    if (/^https?:\/\//i.test(url)) await shell.openExternal(url);
    return true;
  });
}

async function probeEngine() {
  try {
    const d = await com.request('diagnose', {}, { timeout: 60_000 });
    engineState = { ready: true, comAvailable: !!d.comAvailable, version: d.version, error: null };
  } catch (e) {
    engineState = { ready: false, comAvailable: false, version: null, error: e.message };
  }
  return engineState;
}

app.whenReady().then(async () => {
  nativeTheme.themeSource = 'dark';

  settings = new Settings(app.getPath('userData'));
  cache = new SlideCache(SlideCache.defaultRoot(app.getPath('userData')));
  scanner = new LibraryScanner();
  com = new ComBridge();
  decks = new DeckService(com, cache, settings);

  await fsp.mkdir(cache.root, { recursive: true });

  com.on('worker-restarting', () => send('engine:status', { restarting: true }));
  com.on('worker-exit', (info) => send('engine:status', { restarting: false, exit: info }));
  decks.on('deck-opening', (d) => send('deck:opening', d));

  // Serve cached slide images, refusing anything outside the cache root.
  protocol.handle(IMAGE_SCHEME, async (request) => {
    try {
      const url = new URL(request.url);
      const rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      const target = path.resolve(cache.root, rel);
      const rootResolved = path.resolve(cache.root);
      if (target !== rootResolved && !target.startsWith(rootResolved + path.sep)) {
        return new Response('Forbidden', { status: 403 });
      }
      return await net.fetch(`file://${target.replace(/\\/g, '/')}`);
    } catch (e) {
      return new Response(`Not found: ${e.message}`, { status: 404 });
    }
  });

  wireIpc();
  await createWindow();

  // Warm PowerPoint in the background so the first deck opens without delay.
  probeEngine().then((s) => send('engine:status', s));

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', async () => {
  try { await decks?.shutdown(); } catch { }
  if (process.platform !== 'darwin') app.quit();
});

app.on('before-quit', async () => {
  try { await settings?.save(); } catch { }
});
