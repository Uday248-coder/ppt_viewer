'use strict';

const fsp = require('fs/promises');
const path = require('path');

/**
 * Reads embedded media out of an unpacked OOXML package.
 *
 * The app renders through PowerPoint, but media has to come from the package
 * itself - PowerPoint's COM API will not hand over the bytes behind a video
 * shape. Reading the XML directly also means a fully cached deck can play its
 * media with no PowerPoint involvement at all, which is the whole point of the
 * cache.
 *
 * These files are machine-written XML with no inter-tag whitespace, so targeted
 * pattern matching is both sufficient and far cheaper than a DOM parse.
 */

const VIDEO_EXT = /\.(mp4|m4v|mov|avi|wmv|mpg|mpeg|webm)$/i;
const AUDIO_EXT = /\.(mp3|m4a|wav|aif|aiff|ogg|flac)$/i;
const ANIMATED_EXT = /\.(gif|apng)$/i;

const EMU_PER_INCH = 914400;

function attr(source, name) {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(source || '');
  return m ? m[1] : null;
}

function decodeXml(s) {
  return String(s || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
    .replace(/&amp;/g, '&');
}

/** Relationship id -> { target, type } from a .rels part. */
function relMap(xml) {
  const map = new Map();
  if (!xml) return map;
  const re = /<Relationship\b([^>]*?)\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    const id = attr(m[1], 'Id');
    const target = attr(m[1], 'Target');
    if (id && target) map.set(id, { target: decodeXml(target), type: attr(m[1], 'Type') || '' });
  }
  return map;
}

/**
 * Resolves a relationship id to a bare media filename, or null when the target
 * is not something we can play (a layout, a theme, an ordinary jpeg).
 */
function resolveMedia(rels, rid) {
  if (!rid) return null;
  const rel = rels.get(rid);
  if (!rel) return null;
  let t = rel.target.replace(/^\.\.\//, '');
  t = t.replace(/^\//, '');
  const m = /^media\/([^/]+)$/.exec(t);
  if (!m) return null;
  let name;
  try {
    name = decodeURIComponent(m[1]);
  } catch {
    name = m[1];
  }
  return name;
}

/**
 * The order PowerPoint shows slides in, which is not the numeric order of the
 * slide files inside the package. Getting this wrong would attach every video
 * to the wrong slide, so it is read from presentation.xml rather than guessed.
 */
function slideOrder(presentationXml, presentationRelsXml) {
  const rels = relMap(presentationRelsXml);
  const list = [];
  const lst = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(presentationXml || '');
  if (lst) {
    const re = /<p:sldId\b[^>]*?r:id="([^"]+)"/g;
    let m;
    while ((m = re.exec(lst[1]))) {
      const rel = rels.get(m[1]);
      if (!rel) continue;
      list.push(rel.target.replace(/^\.\.\//, '').replace(/^\/?(ppt\/)?/, ''));
    }
  }
  return list;
}

function slideSize(presentationXml) {
  const m = /<p:sldSz\b([^>]*)\/?>/.exec(presentationXml || '');
  const cx = m ? Number(attr(m[1], 'cx')) : EMU_PER_INCH * 10;
  const cy = m ? Number(attr(m[1], 'cy')) : EMU_PER_INCH * 7.5;
  return { cx: cx > 0 ? cx : EMU_PER_INCH * 10, cy: cy > 0 ? cy : EMU_PER_INCH * 7.5 };
}

/** Shape rectangle and rotation, normalised to fractions of the slide. */
function geometry(block, size) {
  const xf = /<a:xfrm\b([^>]*)>([\s\S]*?)<\/a:xfrm>/.exec(block);
  const inner = xf ? xf[2] : '';
  const rotRaw = xf ? Number(attr(xf[1], 'rot') || 0) : 0;
  const off = /<a:off\b([^>]*)\/?>/.exec(inner);
  const ext = /<a:ext\b([^>]*)\/?>/.exec(inner);
  const ox = off ? Number(attr(off[1], 'x') || 0) : 0;
  const oy = off ? Number(attr(off[1], 'y') || 0) : 0;
  const cx = ext ? Number(attr(ext[1], 'cx') || 0) : 0;
  const cy = ext ? Number(attr(ext[1], 'cy') || 0) : 0;

  // A shape with no explicit extent (rare, but background fills do this) covers
  // the slide. Clamping keeps a stray shape from rendering off-stage.
  const rect = cx > 0 && cy > 0
    ? { left: ox / size.cx, top: oy / size.cy, width: cx / size.cx, height: cy / size.cy }
    : { left: 0, top: 0, width: 1, height: 1 };

  return {
    rect: {
      left: clamp01(rect.left),
      top: clamp01(rect.top),
      width: Math.max(0.01, Math.min(1, rect.width)),
      height: Math.max(0.01, Math.min(1, rect.height)),
    },
    rotation: rotRaw ? Math.round(rotRaw / 60000) : 0,
  };
}

const clamp01 = (n) => (Number.isFinite(n) ? Math.max(-0.5, Math.min(1.5, n)) : 0);

/** Classifies whatever media reference a shape carries. */
function detectMedia(block, rels) {
  const video = /<a:videoFile\b[^>]*(?:r:link|r:embed)="([^"]+)"/.exec(block);
  const audio = /<a:audioFile\b[^>]*(?:r:link|r:embed)="([^"]+)"/.exec(block);
  // PowerPoint 2010+ records the same video a second way inside p14:media.
  const p14 = /<p14:media\b[^>]*r:embed="([^"]+)"/.exec(block);

  for (const [hit, kind] of [[video, 'video'], [audio, 'audio']]) {
    if (!hit) continue;
    const file = resolveMedia(rels, hit[1]) || resolveMedia(rels, p14 && p14[1]);
    if (file) return { kind, file };
  }

  // An animated GIF is an ordinary picture to PowerPoint, which bakes only its
  // first frame into an exported still - so it has to be replayed over the top.
  const blip = /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(block);
  if (blip) {
    const file = resolveMedia(rels, blip[1]);
    if (file && ANIMATED_EXT.test(file)) return { kind: 'gif', file };
  }
  return null;
}

/**
 * Playback decisions recorded in the slide's timing tree.
 *
 * PowerPoint's model: a `play`/`playFrom` command targeting a shape means the
 * media runs on its own. `dur="indefinite"` means it waits for a click instead,
 * which is exactly the distinction PowerPoint shows the author.
 */
function timing(xml) {
  const map = new Map();
  const cmdRe = /<p:cmd\b([^>]*)>([\s\S]*?)<\/p:cmd>/g;
  let m;
  while ((m = cmdRe.exec(xml))) {
    const cmd = attr(m[1], 'cmd') || '';
    const body = m[2];
    const spidM = /<p:spTgt\b[^>]*\bspid="(\d+)"/.exec(body);
    if (!spidM) continue;
    const spid = Number(spidM[1]);
    const entry = map.get(spid) || {};

    const ctn = /<p:cTn\b([^>]*)\/?>/.exec(body);
    const dur = ctn ? attr(ctn[1], 'dur') : null;
    const finiteDur = dur && dur !== 'indefinite' ? Number(dur) : null;

    if (/^play(From)?\(/.test(cmd)) {
      const start = /playFrom\(\s*([\d.]+)/.exec(cmd);
      entry.autoplay = true;
      if (start) entry.startMs = Math.round(parseFloat(start[1]) * 1000);
      if (finiteDur != null) entry.endMs = Math.round(finiteDur);
      else if (dur === 'indefinite') entry.autoplay = false;
      const repeat = ctn ? attr(ctn[1], 'repeatCount') : null;
      if ((ctn && attr(ctn[1], 'repeat') === 'indefinite') || repeat === 'indefinite') entry.loop = true;
      else if (repeat && Number(repeat) > 1) entry.loop = true;
    }
    if (cmd === 'stopPlaying' && finiteDur != null) entry.stopMs = Math.round(finiteDur);

    map.set(spid, entry);
  }
  return map;
}

/** Background media is the full-bleed case, so geometry is implicit. */
function backgroundMedia(xml, rels, size) {
  const bg = /<p:bg>([\s\S]*?)<\/p:bg>/.exec(xml);
  if (!bg) return [];
  const inner = bg[1];
  const video = /<a:videoFile\b[^>]*(?:r:link|r:embed)="([^"]+)"/.exec(inner);
  const blip = /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(inner);
  const file =
    resolveMedia(rels, video && video[1]) ||
    resolveMedia(rels, blip && blip[1]);
  if (!file) return [];
  return [
    {
      kind: video ? 'video' : ANIMATED_EXT.test(file) ? 'gif' : 'audio',
      file,
      background: true,
      spid: null,
      name: 'Background media',
      rect: { left: 0, top: 0, width: 1, height: 1 },
      rotation: 0,
      autoplay: true,
      loop: true,
      startMs: 0,
    },
  ];
}

function parseSlide(xml, rels, size) {
  const plays = timing(xml);
  const found = [];
  const seen = new Set();

  const picRe = /<p:pic>([\s\S]*?)<\/p:pic>/g;
  let m;
  while ((m = picRe.exec(xml))) {
    const block = m[1];
    const idM = /<p:cNvPr\b([^>]*)\/?>/.exec(block);
    const spid = idM ? Number(attr(idM[1], 'id')) : NaN;

    // A video shape is written twice - once inside mc:Choice and once inside
    // mc:Fallback - so the same shape id appears more than once per slide.
    if (Number.isFinite(spid)) {
      if (seen.has(spid)) continue;
      seen.add(spid);
    }

    const item = detectMedia(block, rels);
    if (!item) continue;

    const play = Number.isFinite(spid) ? plays.get(spid) || {} : {};
    const { rect, rotation } = geometry(block, size);
    // A GIF runs and loops by definition, and PowerPoint records no play command
    // for it - only an explicit stopPlaying would turn it into a still.
    const animated = item.kind === 'gif';
    found.push({
      kind: item.kind,
      file: item.file,
      spid: Number.isFinite(spid) ? spid : null,
      name: idM ? decodeXml(attr(idM[1], 'name') || '') || '' : '',
      rect,
      rotation,
      // Without a play command PowerPoint waits for a click, so neither do we.
      autoplay: animated ? play.stopMs == null : play.autoplay === true,
      loop: animated ? true : play.loop === true,
      startMs: play.startMs || 0,
      endMs: play.endMs || null,
    });
  }

  return found.concat(backgroundMedia(xml, rels, size));
}

async function readIfExists(file) {
  try {
    return await fsp.readFile(file, 'utf8');
  } catch {
    return '';
  }
}

/**
 * Builds the per-slide media map for an unpacked package.
 *
 * `dir` is the scratch directory the worker unpacked into: a `media` folder of
 * playable files plus a `slides` folder of presentation.xml, its rels, and the
 * XML for slides that actually reference playable media.
 */
async function readUnpacked(dir) {
  const slidesDir = path.join(dir, 'slides');
  const presentationXml = await readIfExists(path.join(slidesDir, '_presentation.xml'));
  const presentationRels = await readIfExists(path.join(slidesDir, '_presentation.xml.rels'));

  const media = [];
  try {
    for (const name of await fsp.readdir(path.join(dir, 'media'))) {
      const st = await fsp.stat(path.join(dir, 'media', name));
      if (!st.isFile()) continue;
      let kind = null;
      if (VIDEO_EXT.test(name)) kind = 'video';
      else if (AUDIO_EXT.test(name)) kind = 'audio';
      else if (ANIMATED_EXT.test(name)) kind = 'gif';
      media.push({ name, kind, bytes: st.size });
    }
  } catch {
    /* no media folder: a deck with no playable media */
  }

  if (!media.length) return { media, slides: new Map(), slideCount: 0, order: [] };

  const size = slideSize(presentationXml);
  const order = slideOrder(presentationXml, presentationRels);
  const partToIndex = new Map();
  order.forEach((part, i) => {
    // Worker-written files are named after the part inside the package.
    const base = part.split('/').pop().replace(/\.xml$/, '');
    partToIndex.set(base, i + 1);
  });

  let slideFiles = [];
  try {
    slideFiles = await fsp.readdir(slidesDir);
  } catch {
    return { media, slides: new Map(), slideCount: 0, order };
  }

  const slides = new Map();
  for (const file of slideFiles) {
    const xmlMatch = /^slide(\d+)\.xml$/.exec(file);
    if (!xmlMatch) continue;
    const base = `slide${xmlMatch[1]}`;
    const displayIndex = partToIndex.get(base) || Number(xmlMatch[1]);

    const xml = await readIfExists(path.join(slidesDir, file));
    const rels = relMap(await readIfExists(path.join(slidesDir, `${base}.rels`)));
    const items = parseSlide(xml, rels, size);
    if (items.length) slides.set(displayIndex, items);
  }

  return { media, slides, slideCount: order.length, order };
}

/** Media the app cannot play back itself (avi/wmv on Chromium builds). */
function isPlayable(kind) {
  return kind === 'video' || kind === 'audio' || kind === 'gif';
}

module.exports = { readUnpacked, isPlayable };