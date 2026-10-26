<#
  Builds every card deck in cards\ by running each folder's Build.ps1.

  Usage:  .\Build-Cards.ps1          builds all decks
          .\Build-Cards.ps1 verbs    builds one deck folder
#>
param([string]$Only = '')

$ErrorActionPreference = 'Stop'

$cardsDir = Split-Path -Parent $PSScriptRoot

$folders = Get-ChildItem -Path $cardsDir -Directory | Where-Object { $_.Name -ne 'general' }
if ($Only) { $folders = @($folders | Where-Object { $_.Name -eq $Only }) }
if ($folders.Count -eq 0) { throw "No deck folders found$(if ($Only) { " matching '$Only'" })." }

foreach ($folder in $folders) {
  $script = Join-Path $folder.FullName 'Build.ps1'
  if (-not (Test-Path $script)) { Write-Warning "Skipping $($folder.Name): no Build.ps1"; continue }
  & $script
}
