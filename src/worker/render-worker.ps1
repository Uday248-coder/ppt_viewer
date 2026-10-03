#requires -Version 5.1
<#
  PPT Viewer - persistent PowerPoint COM render worker.

  Speaks newline-delimited JSON over stdin/stdout. Every response is emitted on a
  single line prefixed with a sentinel so that incidental PowerPoint or
  PowerShell chatter on stdout can never corrupt the protocol.

  The PowerPoint COM object is created once and kept alive between requests:
  instantiating it costs ~1.7s, so amortising it is the single biggest
  responsiveness win in the app.
#>

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# UTF-8 both directions, otherwise non-ASCII filenames and speaker notes mangle.
try {
  [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
  [Console]::InputEncoding  = New-Object System.Text.UTF8Encoding $false
} catch { }

$script:Out      = [Console]::Out
$script:Sentinel = '@@PPV@@'
$script:App      = $null
$script:Pres     = $null
$script:PresPath = $null

# --- PowerPoint enum constants (hard-coded to avoid an interop assembly dependency) ---
$script:ppAlertsNone      = 1
$script:ppPlaceholderBody = 2
$script:ppSaveAsPDF       = 32

# Media extensions worth pulling out of the package. Images are excluded on
# purpose: a large deck has thousands of jpegs under ppt/media and copying them
# would dwarf the payload. Animated GIFs are included because PowerPoint bakes
# only the first frame into an exported still.
$script:MediaExts = @('.mp4', '.m4v', '.mov', '.avi', '.wmv', '.mpg', '.mpeg', '.m4a', '.mp3', '.wav', '.aif', '.aiff', '.gif')

# Loaded lazily for the media-unpacking command; harmless when unavailable.
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue

# ConvertFrom-Json yields PSCustomObject, where $o['name'] silently yields $null.
# Always read fields through this accessor instead of index syntax.
function Get-Field {
  param($Obj, [string]$Name, $Default = $null)
  if ($null -eq $Obj) { return $Default }
  try { $prop = $Obj.PSObject.Properties[$Name] } catch { return $Default }
  if ($null -eq $prop) { return $Default }
  if ($null -eq $prop.Value) { return $Default }
  return $prop.Value
}

function Send-Response {
  param($Obj)
  try {
    $json = $Obj | ConvertTo-Json -Depth 12 -Compress
    $script:Out.WriteLine($script:Sentinel + $json)
    $script:Out.Flush()
  } catch {
    $script:Out.WriteLine($script:Sentinel + '{"id":0,"ok":false,"error":"serialisation failed"}')
    $script:Out.Flush()
  }
}

# Unsolicited progress, so a long batch can report itself without waiting for
# the single reply that ends it. Carries id 0, which never matches a request.
function Send-Event {
  param($Data)
  try {
    $json = @{ id = 0; ok = $true; event = $Data } | ConvertTo-Json -Depth 8 -Compress
    $script:Out.WriteLine($script:Sentinel + $json)
    $script:Out.Flush()
  } catch { }
}

# --- export sizing -----------------------------------------------------------
#
# PowerPoint's Export takes an explicit width AND height, so passing a fixed
# 1920x1080 silently stretches every deck that is not 16:9 - a 4:3 deck comes
# back visibly squashed. Callers now pass only the long edge and the deck's own
# PageSetup decides the other dimension, which makes distortion impossible.
function Get-ExportSize {
  param($Pres, [int]$Long)

  if ($Long -le 0) { $Long = 1920 }

  $w = 0.0; $h = 0.0
  try {
    $w = [double]$Pres.PageSetup.SlideWidth
    $h = [double]$Pres.PageSetup.SlideHeight
  } catch { }
  if ($w -le 0 -or $h -le 0) { $w = 4.0; $h = 3.0 }

  $width = 0; $height = 0
  if ($w -ge $h) {
    $width  = $Long
    $height = [int][Math]::Round($Long * $h / $w)
  } else {
    $width  = [int][Math]::Round($Long * $w / $h)
    $height = $Long
  }
  if ($width -lt 1) { $width = 1 }
  if ($height -lt 1) { $height = 1 }

  return [ordered]@{ width = $width; height = $height }
}

function Read-ZipText {
  param($Entry)
  try {
    $s = $Entry.Open()
    $r = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::UTF8, $true)
    $t = $r.ReadToEnd()
    $r.Dispose()
    $s.Dispose()
    return $t
  } catch {
    return ''
  }
}

function Release-Com {
  param($Obj)
  if ($null -ne $Obj) {
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Obj) } catch { }
  }
}

function Ensure-App {
  if ($null -eq $script:App) {
    $script:App = New-Object -ComObject 'PowerPoint.Application'
    # Suppress modal dialogs; a blocking dialog would hang the whole worker.
    try { $script:App.DisplayAlerts = $script:ppAlertsNone } catch { }
  }
  return $script:App
}

function Close-Presentation {
  if ($null -ne $script:Pres) {
    try { $script:Pres.Close() } catch { }
    Release-Com $script:Pres
    $script:Pres     = $null
    $script:PresPath = $null
  }
}

function Normalise-Text {
  param([string]$Text)
  if ([string]::IsNullOrEmpty($Text)) { return '' }
  return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Get-SlideNotes {
  param($Slide)
  try {
    $notesPage = $Slide.NotesPage
    $text = $null
    for ($i = 1; $i -le $notesPage.Shapes.Count; $i++) {
      $shp = $notesPage.Shapes.Item($i)
      try {
        $ptype = $null
        try { $ptype = $shp.PlaceholderFormat.Type } catch { }
        if ($ptype -eq $script:ppPlaceholderBody) {
          if ($shp.HasTextFrame -eq -1 -and $shp.TextFrame.HasText -eq -1) {
            $text = $shp.TextFrame.TextRange.Text
          }
          break
        }
      } finally { Release-Com $shp }
    }
    # Older decks (and some templates) put the notes in shape 2.
    if ([string]::IsNullOrWhiteSpace($text)) {
      try {
        if ($notesPage.Shapes.Count -ge 2) {
          $shp2 = $notesPage.Shapes.Item(2)
          try {
            if ($shp2.HasTextFrame -eq -1 -and $shp2.TextFrame.HasText -eq -1) {
              $text = $shp2.TextFrame.TextRange.Text
            }
          } finally { Release-Com $shp2 }
        }
      } catch { }
    }
    return Normalise-Text $text
  } catch {
    return ''
  }
}

function Get-SlideTitle {
  param($Slide)
  try {
    $shp = $Slide.Shapes.Title
    if ($null -ne $shp) {
      try {
        if ($shp.HasTextFrame -eq -1 -and $shp.TextFrame.HasText -eq -1) {
          return Normalise-Text $shp.TextFrame.TextRange.Text
        }
      } finally { Release-Com $shp }
    }
  } catch { }
  return ''
}

function Get-DeckInfo {
  param($Pres)
  $slides = New-Object System.Collections.ArrayList
  for ($i = 1; $i -le $Pres.Slides.Count; $i++) {
    $s = $Pres.Slides.Item($i)
    try {
      $effect = $null
      try { $effect = [int]$s.SlideShowTransition.EntryEffect } catch { }
      [void]$slides.Add([ordered]@{
        index  = $i
        title  = Get-SlideTitle $s
        notes  = Get-SlideNotes $s
        effect = $effect
      })
    } finally { Release-Com $s }
  }
  $deckTitle = ''
  try {
    $props = $Pres.BuiltInDocumentProperties
    $t = $props.Item('Title')
    if ($null -ne $t) {
      if (-not [string]::IsNullOrWhiteSpace($t.Value)) { $deckTitle = ([string]$t.Value).Trim() }
    }
    Release-Com $t
    Release-Com $props
  } catch { }

  $ordered = [ordered]@{
    slideCount = [int]$Pres.Slides.Count
    widthPt    = [double]$Pres.PageSetup.SlideWidth
    heightPt   = [double]$Pres.PageSetup.SlideHeight
    title      = $deckTitle
    slides     = $slides.ToArray()
  }
  return $ordered
}

function Invoke-PpvRequest {
  param($Req)

  $cmd  = [string](Get-Field $Req 'cmd' '')
  $data = Get-Field $Req 'data'

  # NOTE: dispatch uses if/elseif rather than `switch` on purpose. A `return`
  # inside a switch block exits only that block, silently emitting $null.
  if ($cmd -eq 'ping') {
    return @{ ok = $true; result = @{ pong = $true; pid = $PID } }
  }

  if ($cmd -eq 'diagnose') {
    $app = Ensure-App
    $v = $null; $ok = $false
    try { $v = [string]$app.Version; $ok = $true } catch { }
    return @{ ok = $true; result = @{ comAvailable = $ok; version = $v } }
  }

  if ($cmd -eq 'open') {
    $path = [string](Get-Field $data 'path' '')
    if ([string]::IsNullOrWhiteSpace($path)) { throw 'open: missing path' }
    if (-not (Test-Path -LiteralPath $path)) { throw "File not found: $path" }
    $app = Ensure-App
    if ($null -ne $script:Pres) { Close-Presentation }
    # FileName, ReadOnly(msoTrue), Untitled(msoFalse), WithWindow(msoFalse)
    $script:Pres = $app.Presentations.Open($path, -1, 0, 0)
    $script:PresPath = $path
    return @{ ok = $true; result = (Get-DeckInfo $script:Pres) }
  }

  if ($cmd -eq 'info') {
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    return @{ ok = $true; result = (Get-DeckInfo $script:Pres) }
  }

  if ($cmd -eq 'close') {
    Close-Presentation
    return @{ ok = $true; result = @{ closed = $true } }
  }

  if ($cmd -eq 'exportDir') {
    # Bulk-renders every slide into $dir. ~34ms/slide at 1920px.
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    $dir = [string](Get-Field $data 'dir' '')
    if ([string]::IsNullOrWhiteSpace($dir)) { throw 'exportDir: missing dir' }
    $long = [int](Get-Field $data 'long' 1920)
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $size = Get-ExportSize $script:Pres $long
    $script:Pres.Export($dir, 'PNG', $size.width, $size.height)
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.PNG' -ErrorAction SilentlyContinue)
    $out = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
      # PowerPoint emits unpadded "Slide1.PNG".."Slide40.PNG" - parse the trailing
      # integer, never sort lexically or slides 10+ end up out of order.
      $n = 0
      if ($f.BaseName -match '(\d+)\s*$') { $n = [int]$Matches[1] }
      [void]$out.Add([ordered]@{ index = $n; file = $f.FullName })
    }
    return @{ ok = $true; result = @{ dir = $dir; files = $out.ToArray(); width = $size.width; height = $size.height } }
  }

  if ($cmd -eq 'exportSlide') {
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    $file = [string](Get-Field $data 'file' '')
    if ([string]::IsNullOrWhiteSpace($file)) { throw 'exportSlide: missing file' }
    $idx = [int](Get-Field $data 'index' 0)
    if ($idx -lt 1) { throw 'exportSlide: bad index' }
    $long = [int](Get-Field $data 'long' 1920)
    # PowerPoint refuses to export into a directory that does not exist yet, and
    # reports it as a confusing "couldn't find <file>" error.
    $parent = [System.IO.Path]::GetDirectoryName($file)
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
      New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    $size = Get-ExportSize $script:Pres $long
    $slide = $script:Pres.Slides.Item($idx)
    try { $slide.Export($file, 'PNG', $size.width, $size.height) } finally { Release-Com $slide }
    return @{ ok = $true; result = @{ index = $idx; file = $file; width = $size.width; height = $size.height } }
  }

  if ($cmd -eq 'exportSlides') {
    # Several slides in one round trip. Issuing exportSlide per slide means N
    # protocol waits, and every one of them sits in front of the user.
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    $dir = [string](Get-Field $data 'dir' '')
    if ([string]::IsNullOrWhiteSpace($dir)) { throw 'exportSlides: missing dir' }
    $long = [int](Get-Field $data 'long' 1920)
    $indices = @(Get-Field $data 'indices' @())
    if ($indices.Count -eq 0) { throw 'exportSlides: missing indices' }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $size = Get-ExportSize $script:Pres $long
    $out = New-Object System.Collections.ArrayList
    $done = 0
    foreach ($raw in $indices) {
      $idx = [int]$raw
      if ($idx -lt 1 -or $idx -gt $script:Pres.Slides.Count) { continue }
      $file = Join-Path $dir ("s$idx.png")
      $slide = $script:Pres.Slides.Item($idx)
      try {
        $slide.Export($file, 'PNG', $size.width, $size.height)
        [void]$out.Add([ordered]@{ index = $idx; file = $file })
      } finally { Release-Com $slide }
      $done += 1
      Send-Event @{ event = 'progress'; phase = 'slides'; done = $done; total = $indices.Count; index = $idx }
    }
    return @{ ok = $true; result = @{ dir = $dir; files = $out.ToArray(); width = $size.width; height = $size.height } }
  }

  if ($cmd -eq 'unpackMedia') {
    # Streams media out of the OOXML package. Needs no PowerPoint at all, so a
    # fully cached deck can still be opened with the engine unavailable.
    $file = [string](Get-Field $data 'file' '')
    $out  = [string](Get-Field $data 'outDir' '')
    if ([string]::IsNullOrWhiteSpace($file)) { throw 'unpackMedia: missing file' }
    if ([string]::IsNullOrWhiteSpace($out))  { throw 'unpackMedia: missing outDir' }
    if (-not (Test-Path -LiteralPath $file)) { throw "File not found: $file" }
    if (-not (Test-Path -LiteralPath $out)) { New-Item -ItemType Directory -Force -Path $out | Out-Null }
    $root = (Resolve-Path -LiteralPath $out).Path.TrimEnd('\')

    # PowerPoint takes an exclusive lock on any deck it has open, so the package
    # cannot be read while the deck is being rendered. When that happens the
    # caller asks for a throwaway copy and reads that instead.
    $src = $file
    $viaCopy = [string](Get-Field $data 'viaCopy' '')
    if (-not [string]::IsNullOrWhiteSpace($viaCopy)) {
      Copy-Item -LiteralPath $file -Destination $viaCopy -Force
      $src = $viaCopy
    }

    function Resolve-InRoot {
      param([string]$Relative)
      $full = [System.IO.Path]::GetFullPath((Join-Path $root $Relative))
      # A crafted package must not be able to write outside the cache.
      if (-not $full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
      $parent = [System.IO.Path]::GetDirectoryName($full)
      if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
      }
      return $full
    }

    $zip = $null
    try { $zip = [System.IO.Compression.ZipFile]::OpenRead($src) }
    catch {
      $inner = ''
      try { $inner = $_.Exception.InnerException.Message } catch { }
      throw "Not a readable OOXML package: $src $inner"
    }

    $media = New-Object System.Collections.ArrayList
    $slides = New-Object System.Collections.ArrayList
    $scanned = 0
    try {
      # Slide file numbers do not follow display order, so the ordered slide list
      # always has to come from presentation.xml.
      foreach ($e in $zip.Entries) {
        if ($e.FullName -eq 'ppt/presentation.xml' -or $e.FullName -eq 'ppt/_rels/presentation.xml.rels') {
          $dest = Resolve-InRoot ("slides\_" + [System.IO.Path]::GetFileName($e.FullName))
          if ($dest) { [IO.File]::WriteAllText($dest, (Read-ZipText $e), (New-Object System.Text.UTF8Encoding $false)) }
        }
      }

      # Pass 1: playable media files only. A large deck has thousands of jpegs
      # under ppt/media and copying them would dwarf everything else.
      foreach ($e in $zip.Entries) {
        if ($e.FullName -notmatch '^ppt/media/[^/]+$') { continue }
        if ([string]::IsNullOrEmpty($e.Name)) { continue }
        $ext = [System.IO.Path]::GetExtension($e.Name).ToLowerInvariant()
        if ($script:MediaExts -notcontains $ext) { continue }
        $dest = Resolve-InRoot ("media\" + $e.Name)
        if (-not $dest) { continue }
        if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue }
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $dest, $true)
        [void]$media.Add([ordered]@{ name = $e.Name; ext = $ext; bytes = [int64]$e.Length })
      }

      # Pass 2: only the slides that actually reference playable media get their
      # XML written out, which keeps a 300-slide deck down to a handful of files.
      foreach ($e in $zip.Entries) {
        if ($e.FullName -notmatch '^ppt/slides/_rels/slide(\d+)\.xml\.rels$') { continue }
        $scanned += 1
        $n = $Matches[1]
        $rels = Read-ZipText $e
        $want = $false
        foreach ($m in [regex]::Matches($rels, 'Target="([^"]+)"')) {
          $ext = [System.IO.Path]::GetExtension($m.Groups[1].Value).ToLowerInvariant()
          if ($script:MediaExts -contains $ext) { $want = $true; break }
        }
        if (-not $want) { continue }

        $destRels = Resolve-InRoot ("slides\slide$n.rels")
        if ($destRels) { [IO.File]::WriteAllText($destRels, $rels, (New-Object System.Text.UTF8Encoding $false)) }

        $xmlEntry = $zip.GetEntry("ppt/slides/slide$n.xml")
        if ($null -ne $xmlEntry) {
          $destXml = Resolve-InRoot ("slides\slide$n.xml")
          if ($destXml) {
            [IO.File]::WriteAllText($destXml, (Read-ZipText $xmlEntry), (New-Object System.Text.UTF8Encoding $false))
          }
        }
        [void]$slides.Add([ordered]@{ part = "slide$n" })
      }
    } finally { $zip.Dispose() }

    return @{ ok = $true; result = @{ media = $media.ToArray(); slides = $slides.ToArray(); scanned = $scanned } }
  }

  if ($cmd -eq 'exportPdf') {
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    $file = [string](Get-Field $data 'file' '')
    if ([string]::IsNullOrWhiteSpace($file)) { throw 'exportPdf: missing file' }
    $folder = Split-Path -Parent $file
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
      New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }
    $keepPath = $script:PresPath
    # SaveAs repoints the in-memory presentation, so reopen the source afterwards
    # to keep the open-deck identity (and its cache key) stable.
    $script:Pres.SaveAs($file, $script:ppSaveAsPDF)
    Close-Presentation
    $app = Ensure-App
    $script:Pres     = $app.Presentations.Open($keepPath, -1, 0, 0)
    $script:PresPath = $keepPath
    return @{ ok = $true; result = @{ file = $file } }
  }

  return @{ ok = $false; error = "Unknown command: $cmd" }
}

# --- main loop -----------------------------------------------------------------
while ($true) {
  $line = $null
  try { $line = [Console]::In.ReadLine() } catch { break }
  if ($null -eq $line) { break }
  if ([string]::IsNullOrWhiteSpace($line)) { continue }

  $req = $null
  try { $req = $line | ConvertFrom-Json -ErrorAction Stop } catch { $req = $null }
  if ($null -eq $req) {
    Send-Response @{ id = 0; ok = $false; error = 'Malformed request' }
    continue
  }

  $id = 0
  try { $id = [int](Get-Field $req 'id' 0) } catch { }
  $cmd = [string](Get-Field $req 'cmd' '')

  if ($cmd -eq 'shutdown') {
    try { Close-Presentation } catch { }
    try { if ($null -ne $script:App) { $script:App.Quit() } } catch { }
    Release-Com $script:App
    $script:App = $null
    Send-Response @{ id = $id; ok = $true; result = @{ bye = $true } }
    break
  }

  $resp = $null
  try {
    $resp = Invoke-PpvRequest $req
  } catch {
    $msg = $_.Exception.Message
    try {
      $inner = $_.Exception.InnerException
      while ($null -ne $inner) { $msg = "$msg | $($inner.Message)"; $inner = $inner.InnerException }
    } catch { }
    $resp = @{ ok = $false; error = $msg; command = $cmd }
  }

  if ($null -eq $resp) { $resp = @{ ok = $false; error = 'Command produced no response' } }
  $resp['id'] = $id
  Send-Response $resp
}
