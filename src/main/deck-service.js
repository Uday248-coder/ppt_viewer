'use strict';

const fsp = require('fs/promises');
const path = require('path');
const os = require('os');
const { EventEmitter } = require('events');
const { SlideCache, KINDS, THUMB_W, THUMB_H, FULL_W, FULL_H } = require('./cache');

const HRESULT_TAIL = /\s*\|?\s*(exception from )?hresult:?\s*0x[0-9a-f]+\s*$/i;

/**
 * PowerPoint reports many failures as bare HRESULTs, which mean nothing to a
 * user. Where it does give real prose ("its file extension has changed"), that
 * wording is more specific than anything we would invent - so keep it.
 */
function humanizeError(err) {
  const raw = String((err && err.message) || '').trim();
  if (!raw) return 'PowerPoint could not complete the request.';

  // Decide whether this is a bare HRESULT by looking for real prose in front of
  // it, rather than guessing from string length.
  const prose = raw.replace(HRESULT_TAIL, '').replace(/[\s,|]+/g, ' ').trim();
  const hasProse = prose.split(' ').filter(Boolean).length >= 3;

  if (!hasProse) {
    if (/0x80cb4002|0x800a03ec/i.test(raw)) {
      return 'This file could not be read by PowerPoint — it may be corrupt, truncated, or not actually a presentation.';
    }
    if (/0x80040154/i.test(raw)) return 'PowerPoint is not installed on this machine.';
    if (/0x800a175d|0x80070bc9/i.test(raw)) {
      return 'PowerPoint automation is unavailable — Office may need to be activated or repaired. See the README.';
    }
    // Unknown code: stay vague rather than assert a cause we cannot confirm.
    return 'PowerPoint could not complete the request.';
  }

  const h = raw.toLowerCase();
  if (/password|encrypted/.test(h)) {
    return 'This deck is password-protected. Remove the password in PowerPoint, then reopen it.';
  }
  if (/0x800a175d|not licensed|deactivat|reduced functionality|automation is (blocked|unavailable)/.test(h)) {
    return 'PowerPoint automation is blocked. Reactivate Office, then reopen the viewer.';
  }
  if (/timed out|timeout/.test(h)) {
    return 'PowerPoint stopped responding (it may be showing a hidden dialog). It has been restarted — try again.';
  }
  if (/class not registered|no application is associated/.test(h)) {
    return 'PowerPoint is not installed on this machine.';
  }
  // Preserve PowerPoint's own explanation, dropping only the HRESULT tail.
  const cleaned = raw.replace(HRESULT_TAIL, '').trim();
  return cleaned || 'PowerPoint could not complete the request.';
}

function fail(err) {
  const e = new Error(humanizeError(err));
  e.cause = err;
  return e;
}

/**
 * Owns everything about "a deck is open": its identity, its cache state, and
 * the render pipeline that fills that cache.
 *
 * The COM worker holds exactly one PowerPoint presentation, so every operation
 * funnels through _ensureOpen(), which reopens the deck if the worker's current
 * presentation has drifted to a different file.
 */
class DeckService extends EventEmitter {
  constructor(comBridge, cache, settings) {
    super();
    this.com = comBridge;
    this.cache = cache;
    this.settings = settings;
    this.current = null; // { path, key, info }
    this._chain = Promise.resolve();
  }

  /** Serialises deck-level operations so concurrent IPC calls cannot interleave. */
  _serial(fn) {
    const run = () => fn().catch((e) => { throw fail(e); });
    const chained = this._chain.then(run, run);
    this._chain = chained.then(
      () => {},
      () => {}
    );
    return chained;
  }

  async _statOf(absPath) {
    try {
      const st = await fsp.stat(absPath);
      if (!st.isFile()) throw new Error('Not a file');
      return st;
    } catch {
      throw new Error(`Deck not found: ${absPath}`);
    }
  }

  async _ensureOpen(absPath, key) {
    if (this.current && this.current.path === absPath) return this.current;
    const info = await this.com.request('open', { path: absPath }, { timeout: 180_000 });
    this.current = { path: absPath, key, info };
    return this.current;
  }

  /**
   * Opens a deck for viewing.
   *
   * If a complete cache already exists the deck opens with zero COM involvement,
   * which is what makes repeat launches feel instant. Otherwise the deck is
   * opened through PowerPoint and its structure is cached.
   */
  async openDeck(absPath, { preferCache = true } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);

      // Fully cached: serve from disk, never touch PowerPoint.
      if (preferCache && status.level === 'full' && status.meta) {
        this.current = null;
        return {
          path: abs,
          key,
          fromCache: true,
          info: status.meta,
          cached: status,
        };
      }

      this.emit('deck-opening', { path: abs, name: path.basename(abs) });
      const opened = await this._ensureOpen(abs, key);
      const info = opened.info;

      const meta = {
        path: abs,
        name: path.basename(abs),
        slideCount: info.slideCount,
        widthPt: info.widthPt,
        heightPt: info.heightPt,
        title: info.title || '',
        slides: info.slides || [],
        cachedAt: Date.now(),
        sourceMtimeMs: Math.floor(st.mtimeMs),
        sourceSize: st.size,
        thumbSize: { w: THUMB_W, h: THUMB_H },
        fullSize: { w: FULL_W, h: FULL_H },
      };
      await this.cache.writeMeta(key, meta);

      return { path: abs, key, fromCache: false, info: meta, cached: await this.cache.status(key) };
    });
  }

  /** Renders any missing thumbnails. Cheap (~5ms/slide), so this is done eagerly. */
  async ensureThumbs(absPath, { onProgress } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      const meta = status.meta;
      if (!meta) throw new Error('Deck metadata is not cached; open the deck first.');

      if (status.level !== 'none') {
        return { key, thumbs: status.thumbs, slideCount: meta.slideCount, rendered: 0 };
      }

      const cur = await this._ensureOpen(abs, key);
      const tmp = path.join(os.tmpdir(), `pptv-thumb-${process.pid}-${key}`);
      const res = await this.com.request(
        'exportDir',
        { dir: tmp, width: THUMB_W, height: THUMB_H },
        { timeout: 300_000 }
      );
      await this.cache.ingest(KINDS.THUMB, res.files, key);
      await fsp.rm(tmp, { recursive: true, force: true });
      onProgress?.({ done: meta.slideCount, total: meta.slideCount });

      return { key, thumbs: meta.slideCount, slideCount: meta.slideCount, rendered: meta.slideCount, info: cur.info };
    });
  }

  /**
   * Renders specific slides at full resolution. Only missing slides are rendered,
   * so revisiting a deck costs nothing.
   */
  async ensureSlides(absPath, indices, { onProgress } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      if (!status.meta) throw new Error('Deck metadata is not cached; open the deck first.');

      const wanted = [...new Set(indices.map(Number).filter((n) => n >= 1 && n <= status.meta.slideCount))];
      const missing = [];
      for (const i of wanted) {
        if (!(await this.cache.exists(KINDS.FULL, key, i))) missing.push(i);
      }
      if (!missing.length) return { key, rendered: 0, alreadyHave: wanted.length };

      await this._ensureOpen(abs, key);
      const tmpDir = path.join(os.tmpdir(), `pptv-full-${process.pid}-${key}`);
      await fsp.mkdir(tmpDir, { recursive: true });

      let done = 0;
      for (const i of missing) {
        const src = path.join(tmpDir, `s${i}.png`);
        await this.com.request('exportSlide', { index: i, file: src, width: FULL_W, height: FULL_H });
        await this.cache.writeSlide(KINDS.FULL, key, i, src);
        done += 1;
        onProgress?.({ done, total: missing.length, index: i });
      }
      await fsp.rm(tmpDir, { recursive: true, force: true });
      return { key, rendered: done, alreadyHave: wanted.length - done };
    });
  }

  /** Pre-renders the whole deck so it later opens with no PowerPoint at all. */
  async cacheDeck(absPath, { onProgress } = {}) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const status = await this.cache.status(key);
      if (!status.meta) throw new Error('Deck metadata is not cached; open the deck first.');

      const total = status.meta.slideCount;
      const missing = [];
      for (let i = 1; i <= total; i++) {
        if (!(await this.cache.exists(KINDS.FULL, key, i))) missing.push(i);
      }
      if (!missing.length) return { key, rendered: 0, total, complete: true };

      await this._ensureOpen(abs, key);
      const tmp = path.join(os.tmpdir(), `pptv-bulk-${process.pid}-${key}`);
      const res = await this.com.request('exportDir', { dir: tmp, width: FULL_W, height: FULL_H }, { timeout: 900_000 });
      // exportDir renders everything; ignore any already-cached slides when counting.
      await this.cache.ingest(KINDS.FULL, res.files, key);
      await fsp.rm(tmp, { recursive: true, force: true });
      onProgress?.({ done: total, total });

      return { key, rendered: res.files.length, total, complete: true };
    });
  }

  async exportPdf(absPath, outFile) {
    return this._serial(async () => {
      const abs = path.resolve(absPath);
      const st = await this._statOf(abs);
      const key = SlideCache.keyFor(abs, st);
      const out = path.resolve(outFile);
      if (!out.toLowerCase().endsWith('.pdf')) out += '.pdf';
      await this._ensureOpen(abs, key);
      const res = await this.com.request('exportPdf', { file: out }, { timeout: 300_000 });
      return { file: res.file };
    });
  }

  async close() {
    return this._serial(async () => {
      if (this.current) {
        try { await this.com.request('close'); } catch { }
      }
      this.current = null;
    });
  }

  /** Releases the PowerPoint process - used when the app goes idle. */
  async shutdown() {
    await this.close();
    await this.com.stop();
  }
}

module.exports = { DeckService, humanizeError };
