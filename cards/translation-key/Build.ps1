# Translation key: one alphabetical list of every verb, used to check card answers.
# Not a card deck - it only borrows the shared renderer.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'general\CardEngine.ps1')

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$csv  = Join-Path $root 'spanish_verbs_top100.csv'
$html = Join-Path $root 'build\verb-translations.html'
$pdf  = Join-Path $root 'Verb_Translations.pdf'
if (-not (Test-Path $csv)) { throw "Missing source: $csv" }

$verbs = @(Import-Csv -Path $csv -Encoding UTF8)

$out = @()
$out += '<!DOCTYPE html>'
$out += '<html lang="en"><head><meta charset="utf-8"><title>Spanish Verbs - Translation Key</title>'
$out += '<style>'
$out += '@page { size: A4 portrait; margin: 14mm 13mm; }'
$out += '*{ box-sizing:border-box; }'
$out += 'html{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }'
$out += 'body{ margin:0; font-family: Georgia,"Cambria","Times New Roman",serif; font-size:8pt; line-height:1.1; color:#1d2229; }'
$out += 'h1{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:15pt; font-weight:600; color:#17365d; margin:0 0 1mm 0; }'
$out += '.lead{ font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:8pt; color:#5d6874; margin:0 0 4mm 0; letter-spacing:.02em; }'
$out += '.hdr{ display:flex; align-items:baseline; gap:2mm; padding:0 0 1.2mm 0; margin-bottom:1mm; border-bottom:1.5px solid #17365d; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:6.8pt; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:#a62b37; break-after:avoid; }'
$out += '.r{ display:flex; align-items:baseline; gap:2mm; padding:.3mm 0; border-bottom:.5px dotted #dfe4ea; break-inside:avoid; }'
$out += '.n{ flex:none; width:8mm; text-align:right; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:6.8pt; color:#99a2ad; }'
$out += '.v{ flex:none; width:32mm; font-weight:700; color:#17365d; }'
$out += '.e{ flex:1; color:#444b54; }'
$out += '.c{ flex:none; width:62mm; font-family:"Segoe UI","Trebuchet MS",sans-serif; font-size:7.2pt; color:#0e6258; }'
$out += '</style></head><body>'

$out += '<h1>Spanish Verbs &mdash; Translation Key</h1>'
$out += '<p class="lead">Alphabetical. The number matches the one on the English side of each flashcard.</p>'
$out += '<div class="list">'
$out += '<div class="hdr"><span class="n">#</span><span class="v">Verb</span><span class="e">Meaning</span><span class="c">Presente conjugation pattern</span></div>'

$i = 1
foreach ($v in ($verbs | Sort-Object verb)) {
  $out += ('<div class="r"><span class="n">{0}</span><span class="v">{1}</span><span class="e">{2}</span><span class="c">{3}</span></div>' -f `
    $i, (HtmlEncode (TitleCase $v.verb)), (HtmlEncode $v.english), (HtmlEncode $v.conjugation))
  $i++
}

$out += '</div></body></html>'

Set-Content -Path $html -Value ($out -join "`r`n") -Encoding UTF8
Convert-HtmlToPdf $html $pdf

$kb = [math]::Round((Get-Item $pdf).Length / 1KB)
Write-Host ("OK  {0,-22} -> {1} ({2} KB, {3} entries)" -f 'Translation Key', (Split-Path $pdf -Leaf), $kb, $verbs.Count)
