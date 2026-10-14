$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$csv  = Join-Path $root 'spanish_verbs_top100.csv'
$out  = Join-Path $root 'build\pattern-counts.txt'

$verbs = @(Import-Csv -Path $csv -Encoding UTF8)

$lines = @()
$lines += "TOTAL VERBS: $($verbs.Count)"
$lines += ''
$lines += '--- exact conjugation value ---'
$verbs | Group-Object conjugation | Sort-Object Count -Descending | ForEach-Object {
  $lines += ("{0,4}  {1}" -f $_.Count, $_.Name)
}

$lines += ''
$lines += '--- split: simple vs reflexive ---'
$verbs | Group-Object form | ForEach-Object { $lines += ("{0,4}  {1}" -f $_.Count, $_.Name) }

$lines += ''
$lines += '--- broad family ---'
$fam = $verbs | ForEach-Object {
  $c = $_.conjugation
  if     ($c -like 'regular -ar*')  { 'regular -ar' }
  elseif ($c -like 'regular -er*')  { 'regular -er' }
  elseif ($c -like 'regular -ir*')  { 'regular -ir' }
  elseif ($c -like 'stem*')         { 'stem-changing' }
  elseif ($c -like 'irregular*')    { 'irregular' }
  elseif ($c -like 'like*')         { 'follows another verb' }
  else                              { "OTHER: $c" }
}
$fam | Group-Object | Sort-Object Count -Descending | ForEach-Object {
  $lines += ("{0,4}  {1}" -f $_.Count, $_.Name)
}

Set-Content -Path $out -Value ($lines -join "`r`n") -Encoding UTF8
Write-Host "Written to $out"
