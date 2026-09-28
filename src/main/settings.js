'use strict';

const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');

const DEFAULTS = {
  libraryRoots: [],
  recents: [],
  prefs: {
    renderWidth: 1920,
    renderHeight: 1080,
    transition: 'fade',
    transitionMs: 320,
    showNotes: true,
    gridColumns: 4,
    prewarmThumbs: true,
    sortBy: 'name',
  },
  window: { width: 1440, height: 900, x: null, y: null, maximized: false },
};

/** Small, atomic JSON-backed config store in the app's userData directory. */
class Settings {
  constructor(userDataDir) {
    this.file = path.join(userDataDir, 'settings.json');
    this.data = this._load();
  }

  _load() {
    try {
      const raw = fs.readFileSync(this.file, 'utf8');
      const parsed = JSON.parse(raw);
      return {
        ...DEFAULTS,
        ...parsed,
        prefs: { ...DEFAULTS.prefs, ...(parsed.prefs || {}) },
        window: { ...DEFAULTS.window, ...(parsed.window || {}) },
      };
    } catch {
      return JSON.parse(JSON.stringify(DEFAULTS));
    }
  }

  async save() {
    const tmp = this.file + '.tmp';
    await fsp.writeFile(tmp, JSON.stringify(this.data, null, 2), 'utf8');
    await fsp.rename(tmp, this.file);
  }

  get all() {
    return this.data;
  }

  setPref(key, value) {
    this.data.prefs[key] = value;
    return this.save();
  }

  addRoot(dir) {
    const abs = path.resolve(dir);
    if (!this.data.libraryRoots.some((r) => r.toLowerCase() === abs.toLowerCase())) {
      this.data.libraryRoots.push(abs);
    }
    return this.save();
  }

  removeRoot(dir) {
    const abs = path.resolve(dir).toLowerCase();
    this.data.libraryRoots = this.data.libraryRoots.filter((r) => r.toLowerCase() !== abs);
    return this.save();
  }

  /** Most-recent-first, de-duplicated by path, capped. */
  async touchRecent(entry) {
    const key = path.resolve(entry.path).toLowerCase();
    this.data.recents = [entry, ...this.data.recents.filter((r) => path.resolve(r.path).toLowerCase() !== key)].slice(0, 40);
    return this.save();
  }

  async clearRecents() {
    this.data.recents = [];
    return this.save();
  }
}

module.exports = { Settings, DEFAULTS };
