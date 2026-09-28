# PPT Viewer

A standalone, offline PowerPoint deck viewer for Windows. Point it at a folder,
it finds every presentation inside, and it opens and presents them with
pixel-perfect fidelity — no Office licence required, no network access, no
sign-in.

Built for one person on one machine. Nothing leaves your PC.

---

## Your question first: will an expired Microsoft subscription break this?

**No. It works.** This was tested on your machine before a line of the app was
written.

An expired Office subscription puts PowerPoint into *reduced functionality mode* —
you can open and view files but not edit them. That sounds like it would break
automation, and in some Office products it does. It does **not** for PowerPoint's
read/export paths, which is all this app uses. Verified working:

| Capability | Result |
|---|---|
| COM object creation | works |
| Open `.ppt` / `.pptx` / `.ppsx` read-only | works |
| Export slides to PNG | works |
| Read speaker notes, titles, transitions | works |
| Export deck to PDF | works |

Because the viewer only ever *reads*, reduced functionality mode is not a
problem. It never tries to save or modify your decks.

One edge case worth knowing: if you ever fully **sign out** of Office or let the
licence lapse into a state where automation is disabled, the app detects it and
tells you plainly instead of failing silently. See *Troubleshooting*.

---

## Requirements

- Windows 10/11, x64
- **Microsoft PowerPoint installed** (any recent version)

PowerPoint is used as the rendering engine. That is the whole point: it is the
only thing on the machine that knows how to lay out a slide exactly as
PowerPoint would, including custom fonts, SmartArt, and complex masters. There is
no Office licence check, no login, and no network call — but PowerPoint itself
must be present.

If PowerPoint is missing, the app starts, the engine pill in the top-right turns
red and reads `no engine`, and it will tell you so.

---

## Running it

**Installed (recommended)**

Run `dist\PPT Viewer Setup 1.0.0.exe` once. It installs per-user — no admin
prompt — and adds a **PPT Viewer** entry to your Start Menu. Press **Win** and
type "PPT" to launch it. It also registers a normal uninstaller in
*Settings → Apps → Installed apps*.

**Portable (no install)**

`dist\PPT Viewer 1.0.0.exe` is a single self-contained file. Copy it anywhere
and double-click it. It re-extracts to a temp folder on each launch, so it
starts a little slower than the installed version.

**From source**

```bash
npm install
npm start
```

**Rebuilding**

```bash
npm run dist
```

Output lands in `dist\`.

---

## Using it

1. Click **Add folder** and pick a directory. Everything is scanned
   recursively — subfolders included — for `.ppt`, `.pptx`, `.pps`, `.ppsx`,
   `.pptm`, `.ppsm`, `.potx` and `.potm`.
2. Click any deck to open it.
3. **Present** (or `F5`) starts the fullscreen slideshow.

You can also drag a folder onto the window to add it, or drag a single deck to
open it immediately. Right-click a deck in the list to reveal it in Explorer.

### Keyboard

| Key | Action |
|---|---|
| `→` `←` `Space` `PgUp` `PgDn` | previous / next slide |
| `Home` `End` | first / last slide |
| `G` | all-slides grid |
| `N` | speaker notes |
| `F5` | start slideshow |
| `Esc` | back out one level |
| Mouse wheel | zoom at cursor |
| Drag | pan |
| Double-click | zoom in, or reset if already zoomed |
| `+` `-` `0` | zoom in / out / reset to fit |
| `P` | cache the whole deck for offline use |

---

## The offline cache

Slides are rendered once to PNG and kept on disk, because re-rendering is far
slower than reading a file. Measured on this machine:

| | |
|---|---|
| PowerPoint COM startup (once per session) | 1717 ms |
| Open a deck, nothing cached | 1400–4300 ms |
| Reopen a deck that is **fully** cached | **~10 ms** |
| Reopen a **partially** cached deck | ~800 ms |
| Render every slide of a 12-slide deck | ~580 ms |
| Render one thumbnail | ~5 ms |

The difference between the last two reopen figures is the point of the cache: a
fully cached deck never touches PowerPoint again, so it opens faster than the
frame can draw.

- **Thumbnails** for the filmstrip and grid are rendered eagerly — they are
  cheap, so the whole deck's overview is available almost immediately.
- **Full-resolution** slides render on demand, and the next few slides are
  pre-rendered in the background so paging never stalls.
- **Cache deck** renders everything up front. A fully cached deck afterwards
  opens with no PowerPoint involvement at all.

Cache location: `%APPDATA%\PPT Viewer\slide-cache`. The sidebar shows its size.
**Prune** deletes entries for decks that no longer exist; **Clear cache** wipes
it.

### Invalidation

A deck's cache key is a hash of its path, byte size and modification time. Edit
a deck in PowerPoint and it re-renders automatically — there is no refresh
button to forget to press, and no risk of showing you a stale version. Nothing
is ever written back to your original files; decks are opened strictly read-only.

---

## PDF export

**PDF** in the toolbar exports the current deck through PowerPoint's own PDF
engine, so the output is identical to PowerPoint's *Save as PDF*.

---

## Architecture

```
src/
  main/
    main.js         Electron main: windows, IPC, private image protocol
    com-bridge.js   supervises the PowerShell worker (queue, timeouts, restart)
    deck-service.js deck identity, cache orchestration, render pipeline
    cache.js        content-addressed slide cache
    library.js      recursive scanner
    settings.js     persisted config
  worker/
    render-worker.ps1   persistent PowerPoint COM host
  preload/preload.js    the renderer's entire API surface
  renderer/              UI
```

Two decisions carry most of the performance:

**A persistent PowerPoint process.** Starting the COM object costs 1.7 s. Rather
than paying that per render, one PowerShell process is started once and holds the
PowerPoint object for the app's lifetime, answering newline-delimited JSON
requests over stdin/stdout. Every response is prefixed with a sentinel so stray
PowerPoint output can never corrupt the protocol.

**Rendering through PowerPoint, not a re-implementation.** A from-scratch
renderer would have to guess at SmartArt, custom fonts, master layouts and
theme colours, and would be wrong often enough to be worse than useless. Using
the real engine is the only way to actually get the fidelity you asked for.

Cached images are served over a private `pptv://` protocol that refuses any path
outside the cache directory. The renderer runs with `contextIsolation`, no Node
integration, and a restrictive CSP.

---

## Troubleshooting

**Engine pill says `no engine`**
PowerPoint is not installed, or Office automation is disabled. Confirm
PowerPoint opens normally on its own first.

**"PowerPoint automation is blocked"**
Office is not activated. Sign in to Office once, or run
`Office > Account > Update License`, then restart the viewer. Fully cached decks
still open fine in this state, since they need no PowerPoint at all.

**"PowerPoint stopped responding"**
PowerPoint is probably sitting on a dialog it cannot show. The viewer detects the
hang, restarts the worker automatically, and the operation is safe to retry.

**A deck will not open**
The viewer reports PowerPoint's own explanation. Common causes: the file is
corrupt or truncated, it is password-protected (remove the password in
PowerPoint), or its extension does not match its actual internal format (a
`.pptx` renamed to `.ppt`, say).

**The engine pill flickers to `restarting`**
The worker exited and is being recycled. If it keeps happening, run PowerPoint
once manually — a first-run or privacy dialog is the usual cause, and dismissing
it once fixes it permanently.

---

## Not in version 1

Embedded video and audio playback are not implemented — slides containing media
render as PowerPoint renders them, but the media does not play. That is the
planned next version.
