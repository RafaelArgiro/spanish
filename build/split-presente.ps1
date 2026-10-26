# One-time migration: splits the monolithic build\presente.html into the shared
# stylesheet and the presente-only body, matching the explainers\ layout.
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'build\presente.html'
if (-not (Test-Path $src)) { throw "Missing source: $src" }

$cssOut  = Join-Path $root 'explainers\general\style.css'
$bodyOut = Join-Path $root 'explainers\presente\body.html'

foreach ($d in @((Split-Path $cssOut -Parent), (Split-Path $bodyOut -Parent))) {
  if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

$html = Get-Content $src -Raw -Encoding UTF8

$sOpen  = $html.IndexOf('<style>')
$sClose = $html.IndexOf('</style>')
if ($sOpen -lt 0 -or $sClose -lt 0) { throw 'Could not locate <style> block.' }
$css = $html.Substring($sOpen + 7, $sClose - ($sOpen + 7)).Trim()

$bodyStart = $html.IndexOf('<!-- ===================== TOC')
$bodyEnd   = $html.LastIndexOf('</body>')
if ($bodyStart -lt 0 -or $bodyEnd -lt 0) { throw 'Could not locate body range.' }
$body = $html.Substring($bodyStart, $bodyEnd - $bodyStart).Trim()

Set-Content -Path $cssOut  -Value $css  -Encoding UTF8
Set-Content -Path $bodyOut -Value $body -Encoding UTF8

Write-Host ("style.css : {0} lines" -f ($css  -split "`n").Count)
Write-Host ("body.html : {0} lines" -f ($body -split "`n").Count)
