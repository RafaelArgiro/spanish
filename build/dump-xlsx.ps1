$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$out  = Join-Path $root 'build\xlsx-dump.txt'
$sb   = New-Object System.Text.StringBuilder

function Add-Line($t) { [void]$sb.AppendLine($t) }

function Get-Entry($zip, $name) {
  $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
  if (-not $e) { return $null }
  $r = New-Object System.IO.StreamReader($e.Open())
  $t = $r.ReadToEnd(); $r.Close(); return $t
}

foreach ($file in (Get-ChildItem -Path $root -Filter 'Cards*.xlsx' | Where-Object { $_.Name -notlike '~$*' })) {
  Add-Line "================ $($file.Name) ================"
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
  try {
    $shared = @()
    $ssXml = Get-Entry $zip 'xl/sharedStrings.xml'
    if ($ssXml) {
      $x = [xml]$ssXml
      foreach ($si in $x.sst.si) { $shared += ($si.InnerText) }
    }

    $sheetNames = ($zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/sheet*.xml' } | Sort-Object FullName)
    foreach ($sh in $sheetNames) {
      Add-Line "---- $($sh.FullName) ----"
      $x = [xml](Get-Entry $zip $sh.FullName)

      if ($x.worksheet.sheetFormatPr) {
        Add-Line ("sheetFormatPr: defaultColWidth={0} defaultRowHeight={1} customHeight={2}" -f `
          $x.worksheet.sheetFormatPr.defaultColWidth, $x.worksheet.sheetFormatPr.defaultRowHeight, $x.worksheet.sheetFormatPr.customHeight)
      }
      foreach ($c in $x.worksheet.cols.col) {
        Add-Line ("col min={0} max={1} width={2} customWidth={3}" -f $c.min, $c.max, $c.width, $c.customWidth)
      }
      if ($x.worksheet.mergeCells) {
        Add-Line ("mergeCells count={0}: {1}" -f $x.worksheet.mergeCells.count, (($x.worksheet.mergeCells.mergeCell | ForEach-Object { $_.ref }) -join ', '))
      }
      if ($x.worksheet.pageMargins) {
        Add-Line ("pageMargins L={0} R={1} T={2} B={3}" -f $x.worksheet.pageMargins.left, $x.worksheet.pageMargins.right, $x.worksheet.pageMargins.top, $x.worksheet.pageMargins.bottom)
      }
      if ($x.worksheet.pageSetup) {
        Add-Line ("pageSetup orientation={0} paperSize={1} scale={2}" -f $x.worksheet.pageSetup.orientation, $x.worksheet.pageSetup.paperSize, $x.worksheet.pageSetup.scale)
      }

      foreach ($row in $x.worksheet.sheetData.row) {
        $cells = @()
        foreach ($c in $row.c) {
          $v = $null
          if ($c.t -eq 's') {
            $idx = [int]$c.v
            if ($idx -lt $shared.Count) { $v = $shared[$idx] }
          } elseif ($c.t -eq 'inlineStr') {
            $v = $c.is.InnerText
          } else {
            $v = $c.v
          }
          if ($null -ne $v -and $v -ne '') { $cells += ("{0}[s{1}]={2}" -f $c.r, $c.s, $v) }
        }
        if ($cells.Count -gt 0 -or $row.ht) {
          Add-Line ("row {0} ht={1} | {2}" -f $row.r, $row.ht, ($cells -join ' | '))
        }
      }
    }
  } finally { $zip.Dispose() }
  Add-Line ''
}

Set-Content -Path $out -Value $sb.ToString() -Encoding UTF8
Write-Host "Dumped to $out"
