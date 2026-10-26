<#
  Builds every tense explainer in explainers\ into a PDF at the workspace root.

  Each tense lives in its own folder next to this one and supplies only its
  content: meta.json (cover text + output name) and body.html (the sections).
  Everything shared - stylesheet, page shell, cover, PDF rendering - lives here.

  Usage:  .\Build-Explainers.ps1            builds all tenses
          .\Build-Explainers.ps1 presente   builds one tense folder
#>
param([string]$Only = '')

$ErrorActionPreference = 'Stop'

$generalDir    = $PSScriptRoot
$explainersDir = Split-Path -Parent $generalDir
$root          = Split-Path -Parent $explainersDir
$buildDir      = Join-Path $root 'build'

if (-not (Test-Path $buildDir)) { New-Item -ItemType Directory -Path $buildDir | Out-Null }

$css   = (Get-Content (Join-Path $generalDir 'style.css')  -Raw -Encoding UTF8).Trim()
$shell = (Get-Content (Join-Path $generalDir 'shell.html') -Raw -Encoding UTF8)
$cover = (Get-Content (Join-Path $generalDir 'cover.html') -Raw -Encoding UTF8)

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

$folders = Get-ChildItem -Path $explainersDir -Directory | Where-Object { $_.Name -ne 'general' }
if ($Only) { $folders = @($folders | Where-Object { $_.Name -eq $Only }) }
if ($folders.Count -eq 0) { throw "No explainer folders found$(if ($Only) { " matching '$Only'" })." }

foreach ($folder in $folders) {
  $metaPath = Join-Path $folder.FullName 'meta.json'
  $bodyPath = Join-Path $folder.FullName 'body.html'
  if (-not (Test-Path $metaPath)) { Write-Warning "Skipping $($folder.Name): no meta.json"; continue }
  if (-not (Test-Path $bodyPath)) { Write-Warning "Skipping $($folder.Name): no body.html"; continue }

  $meta = Get-Content $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $body = (Get-Content $bodyPath -Raw -Encoding UTF8).Trim()

  $coverHtml = $cover.Replace('{{KICKER}}',     $meta.kicker).
                      Replace('{{TITLE_HTML}}', $meta.titleHtml).
                      Replace('{{SUBTITLE}}',   $meta.subtitle).
                      Replace('{{FOOTER}}',     $meta.footer)

  $html = $shell.Replace('{{TITLE_TEXT}}', $meta.titleText).
                 Replace('{{CSS}}', $css).
                 Replace('{{COVER}}', $coverHtml).
                 Replace('{{BODY}}', $body)

  $htmlPath = Join-Path $buildDir ("{0}.html" -f $folder.Name)
  $pdfPath  = Join-Path $root $meta.output

  Set-Content -Path $htmlPath -Value $html -Encoding UTF8
  Convert-HtmlToPdf $htmlPath $pdfPath

  $kb = [math]::Round((Get-Item $pdfPath).Length / 1KB)
  Write-Host ("OK  {0,-18} -> {1} ({2} KB)" -f $folder.Name, $meta.output, $kb)
}
