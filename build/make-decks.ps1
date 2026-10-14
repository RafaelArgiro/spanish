$ErrorActionPreference = 'Stop'

# Card geometry mirrors Cards_Tenses.xlsx / Cards_Pronouns.xlsx:
# 35.09 Excel chars wide (66.3mm) and 113.4pt tall (40mm).
$CardW = 66.3
$CardH = 40.0
$GapX  = 6.0
$GapY  = 2.03
$Cols  = 2
$Rows  = 6
$PerPage = $Cols * $Rows

$root = Split-Path -Parent $PSScriptRoot

function HtmlEncode($s) {
  if ($null -eq $s) { return '' }
  return ([string]$s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function Get-Css {
  $css = @()
  $css += "@page { size: A4 portrait; margin: 19.05mm 17.78mm; }"
  $css += '*{ box-sizing:border-box; }'
  $css += 'html{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }'
  $css += 'body{ margin:0; font-family: Georgia,"Cambria","Times New Roman",serif; color:#1d2229; }'
  $css += ".sheet{ display:grid; grid-template-columns: repeat($Cols, ${CardW}mm); column-gap:${GapX}mm; row-gap:${GapY}mm; justify-content:start; align-content:start; }"
  $css += '.sheet + .sheet{ break-before:page; }'
  $css += ".card{ width:${CardW}mm; height:${CardH}mm; border:0.8pt solid #222; position:relative; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; padding:3mm 5mm; overflow:hidden; }"
  $css += '.card.blank{ border:none; background:none; }'
  $css += '.es{ font-size:17pt; line-height:1.14; }'
  $css += '.en{ font-size:14pt; line-height:1.16; }'
  $css += '.formula{ font-size:12.5pt; font-weight:700; line-height:1.15; }'
  $css += '.models{ font-size:8.6pt; font-style:italic; margin-top:2mm; line-height:1.3; color:#444b54; }'
  $css += '.note{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:7pt; color:#6b7784; margin-top:2mm; line-height:1.25; }'
  $css += '.tag{ position:absolute; bottom:1.6mm; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:5.4pt; letter-spacing:.12em; text-transform:uppercase; opacity:.62; }'
  $css += '.tag.r{ right:2.4mm; }'
  $css += '.tag.l{ left:2.4mm; }'
  # palette lifted from presente.pdf
  $css += '.t-navy{ background:#eef3f9; border-color:#17365d; } .t-navy .es,.t-navy .formula,.t-navy .en,.t-navy .tag{ color:#17365d; }'
  $css += '.t-red{  background:#fcecee; border-color:#a62b37; } .t-red .es,.t-red .formula,.t-red .en,.t-red .tag{ color:#a62b37; }'
  $css += '.t-teal{ background:#e7f3f1; border-color:#0e6258; } .t-teal .es,.t-teal .formula,.t-teal .en,.t-teal .tag{ color:#0e6258; }'
  $css += '.t-gold{ background:#fdf5e3; border-color:#8a6315; } .t-gold .es,.t-gold .formula,.t-gold .en,.t-gold .tag{ color:#8a6315; }'
  return ($css -join "`r`n")
}

function Render-Pdf($htmlPath, $pdfPath) {
  $candidates = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  $exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $exe) { throw 'No Chromium-based browser (Edge or Chrome) found.' }

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
  $kb = [math]::Round((Get-Item $pdfPath).Length / 1KB)
  Write-Host "OK -> $pdfPath ($kb KB)"
}

function Build-Deck($title, $cards, $htmlName, $pdfName) {
  $sb = New-Object System.Text.StringBuilder
  function W($t) { [void]$sb.AppendLine($t) }

  W '<!DOCTYPE html>'
  W ('<html lang="es"><head><meta charset="utf-8"><title>{0}</title>' -f (HtmlEncode $title))
  W ('<style>{0}</style></head><body>' -f (Get-Css))

  $pageCount = [math]::Ceiling($cards.Count / $PerPage)
  $hasBacks  = @($cards | Where-Object { $_.Back }).Count -gt 0

  for ($p = 0; $p -lt $pageCount; $p++) {
    $slice = @($cards[($p * $PerPage)..([math]::Min(($p + 1) * $PerPage - 1, $cards.Count - 1))])

    # front
    W '<section class="sheet">'
    foreach ($c in $slice) {
      $tag = if ($c.Tag) { '<div class="tag r">{0}</div>' -f (HtmlEncode $c.Tag) } else { '' }
      W ('  <div class="card t-{0}"><div class="es">{1}</div>{2}</div>' -f $c.Theme, $c.Front, $tag)
    }
    W '</section>'

    if (-not $hasBacks) { continue }

    # back, column order reversed for duplex long-edge flip
    W '<section class="sheet">'
    for ($r = 0; $r -lt $Rows; $r++) {
      for ($k = $Cols - 1; $k -ge 0; $k--) {
        $i = $r * $Cols + $k
        if ($i -ge $slice.Count) { W '  <div class="card blank"></div>'; continue }
        $c = $slice[$i]
        $tag = if ($c.Tag) { '<div class="tag l">{0}</div>' -f (HtmlEncode $c.Tag) } else { '' }
        W ('  <div class="card t-{0}">{1}{2}</div>' -f $c.Theme, $c.Back, $tag)
      }
    }
    W '</section>'
  }
  W '</body></html>'

  $htmlPath = Join-Path $root "build\$htmlName"
  $pdfPath  = Join-Path $root $pdfName
  Set-Content -Path $htmlPath -Value $sb.ToString() -Encoding UTF8
  Write-Host "HTML -> $htmlPath"
  Render-Pdf $htmlPath $pdfPath
}

function Expand-Copies($defs) {
  $out = @()
  foreach ($d in $defs) { for ($i = 0; $i -lt $d.Copies; $i++) { $out += $d } }
  return $out
}

# ===================== TENSES =====================
# Four tenses, six copies each, matching Cards_Tenses.xlsx. Spanish only, single-sided.
$tenseDefs = @(
  [pscustomobject]@{ Front = 'El Presente';                     Theme = 'navy'; Tag = ''; Copies = 6; Back = '' }
  [pscustomobject]@{ Front = 'El Pret&eacute;rito<br>Perfecto'; Theme = 'red';  Tag = ''; Copies = 6; Back = '' }
  [pscustomobject]@{ Front = 'El Presente<br>Progresivo';       Theme = 'teal'; Tag = ''; Copies = 6; Back = '' }
  [pscustomobject]@{ Front = 'El Futuro<br>Pr&oacute;ximo';     Theme = 'gold'; Tag = ''; Copies = 6; Back = '' }
)

# ===================== PRONOUNS =====================
# Six pronouns, two copies each, matching Cards_Pronouns.xlsx.
$pronounDefs = @(
  [pscustomobject]@{ Front = 'Yo';                        Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">I</div>' }
  [pscustomobject]@{ Front = 'T&uacute;';                 Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">You</div><div class="note">informal</div>' }
  [pscustomobject]@{ Front = '&Eacute;l / Ella / Usted';  Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">He / She / You (formal)</div>' }
  [pscustomobject]@{ Front = 'Nosotros / Nosotras';       Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">We</div>' }
  [pscustomobject]@{ Front = 'Vosotros / Vosotras';       Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">You (plural)</div><div class="note">Spain only</div>' }
  [pscustomobject]@{ Front = 'Ellos / Ellas / Ustedes';   Theme = 'navy'; Tag = 'Pronombre'; Copies = 2; Back = '<div class="en">They / You (plural + formal)</div>' }
)

Build-Deck 'Spanish Tense Flashcards'   (Expand-Copies $tenseDefs)   'tense-flashcards.html'   'Tense_Flashcards.pdf'
Build-Deck 'Spanish Pronoun Flashcards' (Expand-Copies $pronounDefs) 'pronoun-flashcards.html' 'Pronoun_Flashcards.pdf'
