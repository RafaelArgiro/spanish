$ErrorActionPreference = 'Stop'

# Card geometry mirrors Cards_Template_Vebs.xlsx:
# columns A/C/E = 25.09 Excel chars (47.75mm), gaps B/D = 2.45 chars (5.9mm),
# card rows 85pt (30mm), gap rows 5.75pt (2.03mm), 8 rows per A4 portrait page.
$CardW   = 47.75
$CardH   = 30.0
$GapX    = 5.9
$GapY    = 2.03
$Cols    = 3
$Rows    = 8
$PerPage = $Cols * $Rows

$root = Split-Path -Parent $PSScriptRoot
$csv  = Join-Path $root 'spanish_verbs_top100.csv'
$html = Join-Path $root 'build\verb-flashcards.html'
$pdf  = Join-Path $root 'Verb_Flashcards.pdf'

if (-not (Test-Path $csv)) { throw "Missing source: $csv" }

$cards = @(Import-Csv -Path $csv -Encoding UTF8)
Write-Host "Loaded $($cards.Count) verbs."

# Back-of-card number is the verb's position in the alphabetical translation key.
$keyIndex = @{}
$n = 1
foreach ($v in ($cards | Sort-Object verb)) { $keyIndex[$v.rank] = $n; $n++ }

function HtmlEncode($s) {
  if ($null -eq $s) { return '' }
  return ([string]$s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function TitleCase($s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  return $s.Substring(0, 1).ToUpper() + $s.Substring(1)
}

function GroupName($c) {
  if ($c.form -eq 'reflexive') { return 'Reflexive' }
  if ([int]$c.rank -le 100)    { return 'Top 100' }
  return 'Patterns'
}

$sb = New-Object System.Text.StringBuilder
function W($t) { [void]$sb.AppendLine($t) }

W '<!DOCTYPE html>'
W '<html lang="en"><head><meta charset="utf-8"><title>Spanish Verb Flashcards</title>'
W '<style>'
W "@page { size: A4 portrait; margin: 19.05mm 17.78mm; }"
W '*{ box-sizing:border-box; }'
W 'html{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }'
W 'body{ margin:0; font-family: Georgia, "Cambria", "Times New Roman", serif; color:#111; }'
W ".sheet{ display:grid; grid-template-columns: repeat($Cols, ${CardW}mm); column-gap:${GapX}mm; row-gap:${GapY}mm; justify-content:start; align-content:start; }"
W '.sheet + .sheet{ break-before:page; }'
W ".card{ width:${CardW}mm; height:${CardH}mm; border:0.6pt solid #222; position:relative; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; padding:2mm 3mm; overflow:hidden; }"
W '.card.blank{ border-color:transparent; }'
W '.es{ font-size:14pt; line-height:1.12; letter-spacing:.01em; }'
W '.en{ font-size:10.5pt; line-height:1.15; }'
W '.tag{ position:absolute; bottom:1.1mm; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:5pt; letter-spacing:.11em; text-transform:uppercase; color:#8a94a0; }'
W '.tag.r{ right:1.8mm; }'
W '.tag.l{ left:1.8mm; }'
W '</style></head><body>'

$pageCount = [math]::Ceiling($cards.Count / $PerPage)

for ($p = 0; $p -lt $pageCount; $p++) {
  $slice = @($cards[($p * $PerPage)..([math]::Min(($p + 1) * $PerPage - 1, $cards.Count - 1))])

  # ---- front sheet: Spanish ----
  W '<section class="sheet">'
  foreach ($c in $slice) {
    W ('  <div class="card"><div class="es">{0}</div><div class="tag r">{1}</div></div>' -f (HtmlEncode (TitleCase $c.verb)), (GroupName $c))
  }
  W '</section>'

  # ---- back sheet: English, column order reversed for duplex long-edge flip ----
  W '<section class="sheet">'
  for ($r = 0; $r -lt $Rows; $r++) {
    for ($k = $Cols - 1; $k -ge 0; $k--) {
      $i = $r * $Cols + $k
      if ($i -ge $slice.Count) { W '  <div class="card blank"></div>'; continue }
      $c = $slice[$i]
      $tag = "{0} &middot; {1}" -f (GroupName $c), $keyIndex[$c.rank]
      W ('  <div class="card"><div class="en">{0}</div><div class="tag l">{1}</div></div>' -f `
        (HtmlEncode $c.english), $tag)
    }
  }
  W '</section>'
}

W '</body></html>'

Set-Content -Path $html -Value $sb.ToString() -Encoding UTF8
Write-Host "HTML -> $html"

# ---- render ----
$candidates = @(
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $exe) { throw 'No Chromium-based browser (Edge or Chrome) found.' }

$uri     = ([System.Uri]$html).AbsoluteUri
$profile = Join-Path $env:TEMP 'spanish-pdf-profile'
if (Test-Path $pdf) { Remove-Item $pdf -Force }

$common = '--disable-gpu --no-first-run --no-default-browser-check --no-pdf-header-footer --run-all-compositor-stages-before-draw --virtual-time-budget=10000'
$tail   = '"--user-data-dir={0}" "--print-to-pdf={1}" "{2}"' -f $profile, $pdf, $uri

foreach ($mode in @('--headless=new', '--headless')) {
  Start-Process -FilePath $exe -ArgumentList "$mode $common $tail" -Wait -NoNewWindow
  if (Test-Path $pdf) { break }
}
if (-not (Test-Path $pdf)) { throw 'PDF was not produced.' }

$kb = [math]::Round((Get-Item $pdf).Length / 1KB)
Write-Host "OK -> $pdf ($kb KB, $($pageCount * 2) pages)"
