'use strict';

/**
 * Presenter window.
 *
 * A second window for the person holding the clicker: current slide, next
 * slide, notes, elapsed time and wall clock. It is a separate OS window so it
 * can sit on the projector while the audience window stays on the laptop, which
 * is the arrangement every presenter actually wants.
 *
 * This window never touches the filesystem - it shows images the audience
 * renderer already resolved, over the same pptv:// protocol, and asks the
 * audience window to navigate rather than driving PowerPoint itself.
 */

const { BrowserWindow, screen } = require('electron');
const path = require('path');

const TITLE = 'PPT Presenter';

let presenter = null;
let tick = null;

function presenterWindows() {
  return BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed() && w.getTitle() === TITLE);
}

function broadcast(channel, payload) {
  for (const w of presenterWindows()) w.webContents.send(channel, payload);
}

/**
 * Opens the presenter window, or focuses it if it is already up.
 *
 * Placed on the first display that is not showing the audience window, so the
 * common two-monitor setup needs no dragging.
 */
async function openPresenter(getState) {
  const existing = presenterWindows()[0];
  if (existing) {
    existing.show();
    existing.focus();
    pushState(getState?.());
    return true;
  }

  // Put it on a different screen from the audience window where possible, so
  // the common projector/laptop setup needs no dragging.
  let placement = {};
  const audience = BrowserWindow.getAllWindows().find((w) => w.getTitle() === 'PPT Viewer');
  if (audience) {
    try {
      const on = screen.getDisplayMatching(audience.getBounds());
      const other = screen.getAllDisplays().find((d) => d.id !== on.id);
      if (other) {
        placement = { x: other.bounds.x + 60, y: other.bounds.y + 40 };
      }
    } catch {
      /* one screen, or the window has no bounds yet: let the OS decide */
    }
  }

  presenter = new BrowserWindow({
    width: 1180,
    height: 740,
    minWidth: 900,
    minHeight: 560,
    ...placement,
    backgroundColor: '#0f1115',
    title: TITLE,
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      spellcheck: false,
      // Never plays anything itself; it only shows pictures.
      autoplayPolicy: 'no-user-gesture-required',
    },
  });

  // Keep the title fixed so it can be found again by name.
  presenter.on('page-title-updated', (e) => e.preventDefault());
  presenter.setTitle(TITLE);

  presenter.once('ready-to-show', () => presenter.show());
  presenter.on('closed', () => {
    presenter = null;
    broadcast('presenter:exit', {});
  });

  await presenter.loadFile(path.join(__dirname, '..', 'renderer', 'presenter.html'));
  pushState(getState?.());
  return true;
}

function closePresenter() {
  for (const w of presenterWindows()) w.close();
  presenter = null;
}

function isOpen() {
  return presenterWindows().length > 0;
}

/** Forwards the audience renderer's view of the deck to the presenter window. */
function pushState(s) {
  if (!s || !isOpen()) return;
  broadcast('presenter:state', s);
}

function pushTick(payload) {
  if (!isOpen()) return;
  broadcast('presenter:tick', payload);
}

function disposePresenter() {
  if (tick) clearInterval(tick);
  tick = null;
  closePresenter();
}

module.exports = { openPresenter, closePresenter, isOpen, pushState, pushTick, disposePresenter, TITLE };