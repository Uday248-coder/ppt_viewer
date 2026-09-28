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
    $w = [int](Get-Field $data 'width' 0);  if ($w -le 0) { $w = 1920 }
    $h = [int](Get-Field $data 'height' 0); if ($h -le 0) { $h = 1080 }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $script:Pres.Export($dir, 'PNG', $w, $h)
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.PNG' -ErrorAction SilentlyContinue)
    $out = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
      # PowerPoint emits unpadded "Slide1.PNG".."Slide40.PNG" - parse the trailing
      # integer, never sort lexically or slides 10+ end up out of order.
      $n = 0
      if ($f.BaseName -match '(\d+)\s*$') { $n = [int]$Matches[1] }
      [void]$out.Add([ordered]@{ index = $n; file = $f.FullName })
    }
    return @{ ok = $true; result = @{ dir = $dir; files = $out.ToArray() } }
  }

  if ($cmd -eq 'exportSlide') {
    if ($null -eq $script:Pres) { throw 'No presentation is open' }
    $file = [string](Get-Field $data 'file' '')
    if ([string]::IsNullOrWhiteSpace($file)) { throw 'exportSlide: missing file' }
    $idx = [int](Get-Field $data 'index' 0)
    if ($idx -lt 1) { throw 'exportSlide: bad index' }
    $w = [int](Get-Field $data 'width' 0);  if ($w -le 0) { $w = 1920 }
    $h = [int](Get-Field $data 'height' 0); if ($h -le 0) { $h = 1080 }
    $slide = $script:Pres.Slides.Item($idx)
    try { $slide.Export($file, 'PNG', $w, $h) } finally { Release-Com $slide }
    return @{ ok = $true; result = @{ index = $idx; file = $file } }
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
