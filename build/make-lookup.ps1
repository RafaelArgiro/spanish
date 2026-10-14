$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$csv  = Join-Path $root 'spanish_verbs_top100.csv'
$html = Join-Path $root 'build\verb-translations.html'
$pdf  = Join-Path $root 'Verb_Translations.pdf'

if (-not (Test-Path $csv)) { throw "Missing source: $csv" }

$verbs = @(Import-Csv -Path $csv -Encoding UTF8)
Write-Host "Loaded $($verbs.Count) verbs."

function HtmlEncode($s) {
  if ($null -eq $s) { return '' }
  return ([string]$s).Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function TitleCase($s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  return $s.Substring(0, 1).ToUpper() + $s.Substring(1)
}

$sb = New-Object System.Text.StringBuilder
function W($t) { [void]$sb.AppendLine($t) }

W '<!DOCTYPE html>'
W '<html lang="en"><head><meta charset="utf-8"><title>Spanish Verbs - Translation Key</title>'
W '<style>'
W '@page { size: A4 portrait; margin: 14mm 13mm; }'
W '*{ box-sizing:border-box; }'
W 'html{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }'
W 'body{ margin:0; font-family: Georgia,"Cambria","Times New Roman",serif; font-size:8pt; line-height:1.1; color:#1d2229; }'
W 'h1{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:15pt; font-weight:600; color:#17365d; margin:0 0 1mm 0; }'
W '.lead{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:8pt; color:#5d6874; margin:0 0 4mm 0; letter-spacing:.02em; }'
W '.hdr{ display:flex; align-items:baseline; gap:2mm; padding:0 0 1.2mm 0; margin-bottom:1mm; border-bottom:1.5px solid #17365d; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:6.8pt; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:#a62b37; break-after:avoid; }'
W '.r{ display:flex; align-items:baseline; gap:2mm; padding:.3mm 0; border-bottom:.5px dotted #dfe4ea; break-inside:avoid; }'
W '.n{ flex:none; width:8mm; text-align:right; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:6.8pt; color:#99a2ad; }'
W '.v{ flex:none; width:32mm; font-weight:700; color:#17365d; }'
W '.e{ flex:1; color:#444b54; }'
W '.c{ flex:none; width:62mm; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:7.2pt; color:#0e6258; }'
W '</style></head><body>'

W '<h1>Spanish Verbs &mdash; Translation Key</h1>'
W '<p class="lead">Alphabetical. The number matches the one on the English side of each flashcard.</p>'
W '<div class="list">'
W '<div class="hdr"><span class="n">#</span><span class="v">Verb</span><span class="e">Meaning</span><span class="c">Presente conjugation pattern</span></div>'

$i = 1
foreach ($v in ($verbs | Sort-Object verb)) {
  W ('<div class="r"><span class="n">{0}</span><span class="v">{1}</span><span class="e">{2}</span><span class="c">{3}</span></div>' -f `
    $i, (HtmlEncode (TitleCase $v.verb)), (HtmlEncode $v.english), (HtmlEncode $v.conjugation))
  $i++
}

W '</div></body></html>'

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
Write-Host "OK -> $pdf ($kb KB)"
