# Tense cards: Spanish tense names only, single-sided, colour-coded.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'general\CardEngine.ps1')

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# Geometry mirrors Cards_Tenses.xlsx:
# 35.09 Excel chars wide (66.3mm), 113.4pt tall (40mm).
# Laid out 2 across rather than the original 4, which overflowed A4.
$geometry = @{
  CardW = 66.3; CardH = 40.0
  GapX  = 6.0;  GapY  = 2.03
  Cols  = 2;    Rows  = 6
  Margin  = '19.05mm 17.78mm'
  Border  = '0.8pt'
  Padding = '3mm 5mm'
  FontEs  = '17pt'
  FontEn  = '14pt'
  FontTag = '5.4pt'
}

$defs = @(
  [pscustomobject]@{ Front = 'El Presente';                     Theme = 'navy'; Copies = 6; Back = ''; TagFront = ''; TagBack = '' }
  [pscustomobject]@{ Front = 'El Pret&eacute;rito<br>Perfecto'; Theme = 'red';  Copies = 6; Back = ''; TagFront = ''; TagBack = '' }
  [pscustomobject]@{ Front = 'El Presente<br>Progresivo';       Theme = 'teal'; Copies = 6; Back = ''; TagFront = ''; TagBack = '' }
  [pscustomobject]@{ Front = 'El Futuro<br>Pr&oacute;ximo';     Theme = 'gold'; Copies = 6; Back = ''; TagFront = ''; TagBack = '' }
)

New-CardDocument -Title 'Spanish Tense Flashcards' `
                 -Geometry $geometry `
                 -Cards (Expand-Copies $defs) `
                 -HtmlPath (Join-Path $root 'build\tense-flashcards.html') `
                 -PdfPath  (Join-Path $root 'Tense_Flashcards.pdf')
