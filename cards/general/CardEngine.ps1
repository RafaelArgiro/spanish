<#
  Shared helpers for every card deck. Dot-source from a deck's Build.ps1:

      . (Join-Path (Split-Path -Parent $PSScriptRoot) 'general\CardEngine.ps1')

  A deck supplies geometry plus an array of card objects; everything else
  - stylesheet, sheet layout, duplex mirroring, PDF rendering - lives here.

  Card object:  Front (html)     Back (html, '' for single-sided)
                Theme ('' or navy|red|teal|gold)
                TagFront ('' for none)          TagBack ('' for none)
#>

function HtmlEncode($s) {
  if ($null -eq $s) { return '' }
  return ([string]$s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function TitleCase($s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  return $s.Substring(0, 1).ToUpper() + $s.Substring(1)
}

function Expand-Copies($defs) {
  $out = @()
  foreach ($d in $defs) {
    $n = if ($d.PSObject.Properties.Name -contains 'Copies') { $d.Copies } else { 1 }
    for ($i = 0; $i -lt $n; $i++) { $out += $d }
  }
  return $out
}

function Find-Renderer {
  $candidates = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $exe) { throw 'No Chromium-based browser (Edge or Chrome) found.' }
  return $exe
}

function Convert-HtmlToPdf($htmlPath, $pdfPath) {
  $exe     = Find-Renderer
  $uri     = ([System.Uri]$htmlPath).AbsoluteUri
  $profile = Join-Path $env:TEMP 'spanish-pdf-profile'
  if (Test-Path $pdfPath) { Remove-Item $pdfPath -Force }

  $common = '--disable-gpu --no-first-run --no-default-browser-check --no-pdf-header-footer --run-all-compositor-stages-before-draw --virtual-time-budget=10000'
  $tail   = '"--user-data-dir={0}" "--print-to-pdf={1}" "{2}"' -f $profile, $pdfPath, $uri

  foreach ($mode in @('--headless=new', '--headless')) {
    Start-Process -FilePath $exe -ArgumentList "$mode $common $tail" -Wait -NoNewWindow
    if (Test-Path $pdfPath) { break }
  }
  if (-not (Test-Path $pdfPath)) { throw "PDF was not produced: $pdfPath" }
}

function Get-CardCss {
  param([hashtable]$G)

  $css = @()
  $css += "@page { size: A4 portrait; margin: $($G.Margin); }"
  $css += '*{ box-sizing:border-box; }'
  $css += 'html{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }'
  $css += 'body{ margin:0; font-family: Georgia,"Cambria","Times New Roman",serif; color:#1d2229; }'
  $css += ".sheet{ display:grid; grid-template-columns: repeat($($G.Cols), $($G.CardW)mm); column-gap:$($G.GapX)mm; row-gap:$($G.GapY)mm; justify-content:start; align-content:start; }"
  $css += '.sheet + .sheet{ break-before:page; }'
  $css += ".card{ width:$($G.CardW)mm; height:$($G.CardH)mm; border:$($G.Border) solid #222; position:relative; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; padding:$($G.Padding); overflow:hidden; }"
  $css += '.card.blank{ border:none; background:none; }'
  $css += ".es{ font-size:$($G.FontEs); line-height:1.14; letter-spacing:.01em; }"
  $css += ".en{ font-size:$($G.FontEn); line-height:1.16; }"
  $css += '.formula{ font-size:12.5pt; font-weight:700; line-height:1.15; }'
  $css += '.models{ font-size:8.6pt; font-style:italic; margin-top:2mm; line-height:1.3; color:#444b54; }'
  $css += '.note{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:7pt; color:#6b7784; margin-top:2mm; line-height:1.25; }'
  $css += ".tag{ position:absolute; bottom:1.3mm; font-family:`"Segoe UI`",`"Trebuchet MS`",sans-serif; font-size:$($G.FontTag); letter-spacing:.11em; text-transform:uppercase; color:#8a94a0; opacity:.78; }"
  $css += '.tag.r{ right:2mm; }'
  $css += '.tag.l{ left:2mm; }'
  # palette lifted from the presente explainer
  $css += '.t-navy{ background:#eef3f9; border-color:#17365d; } .t-navy .es,.t-navy .formula,.t-navy .en,.t-navy .tag{ color:#17365d; }'
  $css += '.t-red{  background:#fcecee; border-color:#a62b37; } .t-red .es,.t-red .formula,.t-red .en,.t-red .tag{ color:#a62b37; }'
  $css += '.t-teal{ background:#e7f3f1; border-color:#0e6258; } .t-teal .es,.t-teal .formula,.t-teal .en,.t-teal .tag{ color:#0e6258; }'
  $css += '.t-gold{ background:#fdf5e3; border-color:#8a6315; } .t-gold .es,.t-gold .formula,.t-gold .en,.t-gold .tag{ color:#8a6315; }'
  return ($css -join "`r`n")
}

function New-CardDocument {
  param(
    [string]$Title,
    [hashtable]$Geometry,
    [object[]]$Cards,
    [string]$HtmlPath,
    [string]$PdfPath
  )

  $G        = $Geometry
  $perPage  = $G.Cols * $G.Rows
  $pages    = [math]::Ceiling($Cards.Count / $perPage)
  $hasBacks = @($Cards | Where-Object { $_.Back }).Count -gt 0

  $out = @()
  $out += '<!DOCTYPE html>'
  $out += ('<html lang="es"><head><meta charset="utf-8"><title>{0}</title>' -f (HtmlEncode $Title))
  $out += ('<style>{0}</style></head><body>' -f (Get-CardCss $G))

  for ($p = 0; $p -lt $pages; $p++) {
    $slice = @($Cards[($p * $perPage)..([math]::Min(($p + 1) * $perPage - 1, $Cards.Count - 1))])

    $out += '<section class="sheet">'
    foreach ($c in $slice) {
      $theme = if ($c.Theme)    { " t-$($c.Theme)" } else { '' }
      $tag   = if ($c.TagFront) { '<div class="tag r">{0}</div>' -f $c.TagFront } else { '' }
      $out += ('  <div class="card{0}"><div class="es">{1}</div>{2}</div>' -f $theme, $c.Front, $tag)
    }
    $out += '</section>'

    if (-not $hasBacks) { continue }

    # column order reversed so backs line up after a long-edge duplex flip
    $out += '<section class="sheet">'
    for ($r = 0; $r -lt $G.Rows; $r++) {
      for ($k = $G.Cols - 1; $k -ge 0; $k--) {
        $i = $r * $G.Cols + $k
        if ($i -ge $slice.Count) { $out += '  <div class="card blank"></div>'; continue }
        $c     = $slice[$i]
        $theme = if ($c.Theme)   { " t-$($c.Theme)" } else { '' }
        $tag   = if ($c.TagBack) { '<div class="tag l">{0}</div>' -f $c.TagBack } else { '' }
        $out += ('  <div class="card{0}">{1}{2}</div>' -f $theme, $c.Back, $tag)
      }
    }
    $out += '</section>'
  }

  $out += '</body></html>'

  Set-Content -Path $HtmlPath -Value ($out -join "`r`n") -Encoding UTF8
  Convert-HtmlToPdf $HtmlPath $PdfPath

  $kb = [math]::Round((Get-Item $PdfPath).Length / 1KB)
  Write-Host ("OK  {0,-22} -> {1} ({2} KB, {3} cards)" -f $Title, (Split-Path $PdfPath -Leaf), $kb, $Cards.Count)
}
