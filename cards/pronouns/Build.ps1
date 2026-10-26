# Pronoun cards: Spanish pronoun on the front, English on the back.
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'general\CardEngine.ps1')

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# Geometry mirrors Cards_Pronouns.xlsx:
# 35.09 Excel chars wide (66.3mm), 113.4pt tall (40mm).
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

$tag = 'Pronombre'

$defs = @(
  [pscustomobject]@{ Front = 'Yo';                       Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">I</div>' }
  [pscustomobject]@{ Front = 'T&uacute;';                Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">You</div><div class="note">informal</div>' }
  [pscustomobject]@{ Front = '&Eacute;l / Ella / Usted'; Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">He / She / You (formal)</div>' }
  [pscustomobject]@{ Front = 'Nosotros / Nosotras';      Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">We</div>' }
  [pscustomobject]@{ Front = 'Vosotros / Vosotras';      Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">You (plural)</div><div class="note">Spain only</div>' }
  [pscustomobject]@{ Front = 'Ellos / Ellas / Ustedes';  Theme = 'navy'; Copies = 2; TagFront = $tag; TagBack = $tag
                     Back = '<div class="en">They / You (plural + formal)</div>' }
)

New-CardDocument -Title 'Spanish Pronoun Flashcards' `
                 -Geometry $geometry `
                 -Cards (Expand-Copies $defs) `
                 -HtmlPath (Join-Path $root 'build\pronoun-flashcards.html') `
                 -PdfPath  (Join-Path $root 'Pronoun_Flashcards.pdf')
