$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$html = Join-Path $root 'build\presente.html'
$pdf  = Join-Path $root 'presente.pdf'

if (-not (Test-Path $html)) { throw "Missing source: $html" }

$candidates = @(
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $exe) { throw 'No Chromium-based browser (Edge or Chrome) found.' }
Write-Host "Renderer: $exe"

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
