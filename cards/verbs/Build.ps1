# Verb flashcards: Spanish infinitive on the front, English on the back.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'general\CardEngine.ps1')

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$csv  = Join-Path $root 'spanish_verbs_top100.csv'
if (-not (Test-Path $csv)) { throw "Missing source: $csv" }

# Geometry mirrors Cards_Template_Vebs.xlsx:
# 25.09 Excel chars wide (47.75mm), 85pt tall (30mm), 3 across and 8 down.
$geometry = @{
  CardW = 47.75; CardH = 30.0
  GapX  = 5.9;   GapY  = 2.03
  Cols  = 3;     Rows  = 8
  Margin  = '19.05mm 17.78mm'
  Border  = '0.6pt'
  Padding = '2mm 3mm'
  FontEs  = '14pt'
  FontEn  = '10.5pt'
  FontTag = '5pt'
}

$verbs = @(Import-Csv -Path $csv -Encoding UTF8)

# Back-of-card number is the verb's position in the alphabetical translation key.
$keyIndex = @{}
$n = 1
foreach ($v in ($verbs | Sort-Object verb)) { $keyIndex[$v.rank] = $n; $n++ }

function Get-GroupName($v) {
  if ($v.form -eq 'reflexive') { return 'Reflexive' }
  if ([int]$v.rank -le 100)    { return 'Top 100' }
  return 'Patterns'
}

$cards = foreach ($v in $verbs) {
  $group = Get-GroupName $v
  [pscustomobject]@{
    Front    = HtmlEncode (TitleCase $v.verb)
    Back     = '<div class="en">{0}</div>' -f (HtmlEncode $v.english)
    Theme    = ''
    TagFront = $group
    TagBack  = '{0} &middot; {1}' -f $group, $keyIndex[$v.rank]
  }
}

New-CardDocument -Title 'Spanish Verb Flashcards' `
                 -Geometry $geometry `
                 -Cards $cards `
                 -HtmlPath (Join-Path $root 'build\verb-flashcards.html') `
                 -PdfPath  (Join-Path $root 'Verb_Flashcards.pdf')
