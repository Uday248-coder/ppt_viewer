# PPT Viewer

A standalone, offline PowerPoint deck viewer for Windows. Point it at a folder,
it finds every presentation inside, and it opens and presents them with
pixel-perfect fidelity - no Office licence required, no network access, no
sign-in.

Point it at a **git repository** and it also finds the decks inside that
repository's sub-repositories, so a monorepo of course material reads as one
library instead of twenty folders.

Built for one person on one machine. Nothing leaves your PC.

MIT licensed - see [LICENSE](LICENSE).

---

## What version 2 adds

| | |
|---|---|
| **Embedded media plays** | Video and audio embedded in a deck now play, positioned exactly over the poster frame PowerPoint renders. Animated GIFs animate. |
| **Repository-aware library** | Point at a master repo and its sub-repos are discovered automatically, each with its branch and deck count, and can be navigated independently. |
| **Presenter window** | A second window for the person holding the clicker: current slide, next slide, notes, elapsed timer and clock. |
| **Hideable notes** | The speaker-notes sidebar collapses to a labelled tab and the stage re-fits. |
| **Correct slide shape** | 4:3 decks were being squashed into 16:9. Every deck now renders at its own aspect ratio. |
| **Quicker paging** | Superseded renders are never queued, batches cost one round trip, and the rest of a deck fills itself in the background. |

**Fixed in 2.0:** version 1 forced every export to 1920×1080 regardless of the
deck's real shape, so any 4:3 deck (`Module 4_AWS Database Services_PART 1_2`,
`AWS Networking & Content Delivery`, and others in your library) was displayed
stretched and squashed. Renders are now sized from each deck's own page setup.
Because that changes the pixels, the cache key changed too: previously cached
decks re-render themselves the next time you open them, and **Prune** clears the
old entries away.

---

## Your question first: will an expired Microsoft subscription break this?

**No. It works.** This was tested on your machine before a line of the app was
written.

An expired Office subscription puts PowerPoint into *reduced functionality mode* -
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
| Read embedded media from the package | works (no PowerPoint needed) |

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
no Office licence check, no login, and no network call - but PowerPoint itself
must be present.

If PowerPoint is missing, the app starts, the engine pill in the top-right turns
red and reads `no engine`, and it will tell you so.

---

## Running it

**Installed (recommended)**

Run `dist\PPT Viewer Setup 2.0.0.exe` once. It installs per-user - no admin
prompt - and adds a **PPT Viewer** entry to your Start Menu. Press **Win** and
type "PPT" to launch it. It also registers a normal uninstaller in
*Settings > Apps > Installed apps*.

**Portable (no install)**

`dist\PPT Viewer 2.0.0.exe` is a single self-contained file. Copy it anywhere
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
   recursively - subfolders included - for `.ppt`, `.pptx`, `.pps`, `.ppsx`,
   `.pptm`, `.ppsm`, `.potx` and `.potm`.
2. Click any deck to open it.
3. **Present** (or `F5`) starts the fullscreen slideshow.

You can also drag a folder onto the window to add it, or drag a single deck to
open it immediately. Right-click a deck in the list to reveal it in Explorer.

### Repositories and sub-repos

If a folder you add is a git repository, the sidebar's **Repositories** section
lists it with its branch and a deck count. Any repository nested inside it -
a `git submodule`, a vendored checkout, a `git worktree` - is listed underneath
with its own branch. Clicking one narrows the deck list to it; the breadcrumb
above the list shows where you are and takes you back.

Decks below a nested repository belong to *that* repository, not the one
containing it, which is almost always what you want when a monorepo has
per-team slide decks.

Folders that are not repositories still work exactly as before - they simply do
not appear in that section.

### Embedded media

Video, audio and animated GIFs embedded in a deck play in both the viewer and
the slideshow.

- A video starts when you arrive at its slide, stops at the point its author
  set, and restarts from the top if you come back to the slide.
- Click any playing media to pause it, click again to resume.
- The toolbar shows what the current slide contains (*1 video*, *2 audio*, and
  so on) and the speaker button mutes it.
- Media plays with sound, as it would in PowerPoint's own slideshow.

PowerPoint flattens a video into its poster frame when exporting a slide, so the
video is drawn back over that exact frame using the shape's own position and
size read from the file. It lines up precisely, including cropped and rotated
shapes.

Formats Chromium can decode are played natively - that covers `.mp4`, `.m4v`,
`.mov`, `.webm`, `.mp3`, `.m4a`, `.wav`, `.ogg`, `.flac` and `.gif`. Legacy
`.avi` and `.wmv` cannot be decoded and will simply stay as their poster frame.

### Speaker notes

`N` toggles the notes sidebar. Closing it leaves a slim **Notes** tab on the
right edge, so the wider stage is never a one-way door.

### Presenter window

**Presenter** (or `S`) opens a second window to put on the projector while the
audience window stays on your laptop. It shows the slide on screen, the next
slide, the speaker notes, an elapsed timer and the wall clock.

Advancing works from either window - they are kept in step, so you can click in
the presenter window and the audience window follows. `R` resets the timer, `Esc`
closes the presenter window.

If you have two monitors, the presenter window opens on the one the audience
window is not using.

### Keyboard

| Key | Action |
|---|---|
| `←` `→` `Space` `PgUp` `PgDn` | previous / next slide |
| `Home` `End` | first / last slide |
| `G` | all-slides grid |
| `N` | speaker notes |
| `M` | mute media |
| `F5` | start slideshow |
| `S` | presenter window |
| `R` | reset elapsed timer (presenter window) |
| `Esc` | back out one level |
| Mouse wheel | zoom at cursor |
| Drag | pan |
| Double-click | zoom in, or reset if already zoomed |
| `+` `-` `0` | zoom in / out / reset to fit |
| `P` | render the whole deck now |

---

## The offline cache

Slides are rendered once to PNG and kept on disk, because re-rendering is far
slower than reading a file. Measured on this machine:

| | |
|---|---|
| PowerPoint COM startup (once per session) | 1717 ms |
| Open a deck, nothing cached | 1400-4300 ms |
| Reopen a deck that is **fully** cached | **~10 ms** |
| Reopen a **partially** cached deck | ~800 ms |
| Render every slide of a 12-slide deck | ~580 ms |
| Render one thumbnail | ~5 ms |

The difference between the last two reopen figures is the point of the cache: a
fully cached deck never touches PowerPoint again, so it opens faster than the
frame can draw.

- **Thumbnails** for the filmstrip and grid are rendered eagerly - they are
  cheap, so the whole deck's overview is available almost immediately.
- **Full-resolution** slides render on demand, and the next few slides are
  pre-rendered in the background so paging never stalls.
- Opening a deck that is not fully cached also quietly fills in the rest of it
  in the background, so any later jump is instant and the deck becomes
  offline-capable without you asking. Background work always yields to a render
  you are actually waiting on.
- **Cache deck** (`P`) does the same thing immediately, in the foreground.

Cache location: `%APPDATA%\PPT Viewer\slide-cache`. The sidebar shows its size.
**Prune** deletes entries for decks that no longer exist or that were rendered
by an older version; **Clear cache** wipes it.

### Invalidation

A deck's cache key is a hash of its render version, path, byte size and
modification time. Edit a deck in PowerPoint and it re-renders automatically -
there is no refresh button to forget to press, and no risk of showing you a
stale version. Nothing is ever written back to your original files; decks are
opened strictly read-only.

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
    media.js        reads embedded media out of the OOXML package
    library.js      recursive scanner, repository discovery
    presenter.js    the second (presenter) window
    settings.js     persisted config
  worker/
    render-worker.ps1   persistent PowerPoint COM host
  preload/preload.js    the renderer's entire API surface
  renderer/
    index.html / app.js / styles.css    the library and viewer
    presenter.html / presenter.js       the presenter window
```

Four decisions carry most of the behaviour:

**A persistent PowerPoint process.** Starting the COM object costs 1.7 s. Rather
than paying that per render, one PowerShell process is started once and holds the
PowerPoint object for the app's lifetime, answering newline-delimited JSON
requests over stdin/stdout. Every response is prefixed with a sentinel so stray
PowerPoint output can never corrupt the protocol.

**Rendering through PowerPoint, not a re-implementation.** A from-scratch
renderer would have to guess at SmartArt, custom fonts, master layouts and
theme colours, and would be wrong often enough to be worse than useless. Using
the real engine is the only way to actually get the fidelity you asked for.

**Exports are sized from the deck, not from us.** PowerPoint's export takes an
explicit width *and* height, so the worker passes only a target long edge and
derives the other dimension from the deck's own page setup. There is no longer
any code path that can stretch a slide.

**Two priority lanes.** Anything you are waiting on renders interactively;
the background fill runs at low priority between batches. Without the split, a
300-slide deck's worth of queued renders would sit in front of a single arrow
key press.

Cached images are served over a private `pptv://` protocol that refuses any path
outside the cache directory, and which honours range requests so video can be
seeked. The renderer runs with `contextIsolation`, no Node integration, and a
restrictive CSP.

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

**A video shows its poster frame and does not play**
Either the codec is one Chromium cannot decode (`.avi`, `.wmv`), or the media
was in a part of the package the viewer does not read. The toolbar still names
what is on the slide, which tells you whether it was found at all.

**The engine pill flickers to `restarting`**
The worker exited and is being recycled. If it keeps happening, run PowerPoint
once manually - a first-run or privacy dialog is the usual cause, and dismissing
it once fixes it permanently.

---

## Not in version 2

- Media embedded in **notes pages, slide masters or layouts** is not extracted;
  only media attached to slides is.
- Presentation-level **embedded objects** (`.pptm` add-ins, OLE packages) are not
  supported, as they were not in version 1.
- There is still no **auto-update**; new versions are installed by hand.