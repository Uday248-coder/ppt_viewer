'use strict';

/* PPT Presenter window.
   Shows the audience window's current slide, the next one, notes and the time.
   It never renders anything itself - every command goes back to the audience
   window so there is only ever one slide cursor. */

const api = window.pptv;

const $ = (id) => document.getElementById(id);

const el = {
  deck: $('pv-deck'),
  sub: $('pv-sub'),
  clock: $('pv-clock'),
  timer: $('pv-timer'),
  reset: $('pv-reset'),
  progress: $('pv-progress-fill'),
  frame: $('pv-frame'),
  cur: $('pv-cur'),
  empty: $('pv-empty'),
  caption: $('pv-caption'),
  nextFrame: $('pv-next-frame'),
  nextImg: $('pv-next-img'),
  nextEmpty: $('pv-next-empty'),
  notes: $('pv-notes'),
  notesCount: $('pv-notes-count'),
  prev: $('pv-prev'),
  next: $('pv-next'),
};

let state = null;

function fmtElapsed(ms) {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  return `${h ? `${h}:` : ''}${String(m).padStart(h ? 2 : 1, '0')}:${String(s).padStart(2, '0')}`;
}

function paintClock() {
  const d = new Date();
  el.clock.textContent = d.toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
  });
}

function setImage(img, url) {
  if (img.getAttribute('src') === (url || '')) return;
  if (url) img.setAttribute('src', url);
  else img.removeAttribute('src');
}

function render(s) {
  state = s;
  el.deck.textContent = s.deckTitle || s.deckName || 'No deck';
  el.sub.textContent = `Slide ${s.index} of ${s.slideCount}${s.slideTitle ? ` · ${s.slideTitle}` : ''}`;
  el.progress.style.width = `${(s.index / Math.max(1, s.slideCount)) * 100}%`;
  el.caption.textContent = s.slideTitle || '';

  el.empty.hidden = !!s.curUrl;
  setImage(el.cur, s.curUrl);

  el.nextEmpty.hidden = !!s.nextUrl;
  setImage(el.nextImg, s.nextUrl);

  const notes = (s.notes || '').trim();
  el.notes.textContent = notes || 'No notes for this slide.';
  el.notes.classList.toggle('is-empty', !notes);
  el.notesCount.textContent = notes ? `${notes.length} chars` : '';

  // Match the deck's own shape so "up next" reads at the same proportion as
  // the audience window.
  const aspect = s.aspect && s.aspect > 0 ? s.aspect : 4 / 3;
  el.frame.style.setProperty('--ar', String(aspect));
  el.nextFrame.style.setProperty('--ar', String(aspect));
}

function bind() {
  api.presenter.onState(render);
  api.presenter.onTick((t) => {
    el.timer.textContent = fmtElapsed(t.elapsedMs || 0);
  });
  api.presenter.onExit(() => {
    // The audience window closed its own presenter; nothing left to do here.
    window.close();
  });

  el.prev.addEventListener('click', () => api.presenter.nav(-1));
  el.next.addEventListener('click', () => api.presenter.nav(1));
  el.reset.addEventListener('click', () => api.presenter.resetTimer());

  paintClock();
  setInterval(paintClock, 10_000);

  window.addEventListener('keydown', (e) => {
    if (/^(INPUT|TEXTAREA)$/.test(document.activeElement?.tagName || '')) return;
    switch (e.key) {
      case 'ArrowRight': case 'PageDown': case ' ': case 'Enter':
        e.preventDefault(); api.presenter.nav(1); break;
      case 'ArrowLeft': case 'PageUp': case 'Backspace':
        e.preventDefault(); api.presenter.nav(-1); break;
      case 'r': case 'R':
        e.preventDefault(); api.presenter.resetTimer(); break;
      case 'Escape':
        e.preventDefault(); api.presenter.close(); break;
      default: break;
    }
  });
}

bind();