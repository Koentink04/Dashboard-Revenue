Add-Type -AssemblyName System.IO.Compression.FileSystem

$excelPath = Resolve-Path "Data Dashborad.xlsx"
$tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "temp_data_dashboard_" + [System.Guid]::NewGuid().ToString() + ".xlsx")

$srcStream = [System.IO.File]::Open($excelPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$destStream = [System.IO.File]::Create($tempPath)
$srcStream.CopyTo($destStream)
$srcStream.Close()
$destStream.Close()

$zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)

# 1. Read sharedStrings
$sharedStrings = @()
$ssEntry = $zip.GetEntry("xl/sharedStrings.xml")
if ($ssEntry) {
    $ssStream = $ssEntry.Open()
    $ssXml = [xml](New-Object System.IO.StreamReader($ssStream)).ReadToEnd()
    $ssStream.Close()
    foreach ($si in $ssXml.sst.si) {
        if ($si.t) {
            $sharedStrings += $si.t
        } elseif ($si.r) {
            $txt = ($si.r | ForEach-Object { $_.t }) -join ""
            $sharedStrings += $txt
        } else {
            $sharedStrings += ""
        }
    }
}

# 2. Get the target sheet for TARGET KUM & SA
$wbEntry = $zip.GetEntry("xl/workbook.xml")
$wbStream = $wbEntry.Open()
$wbXml = [xml](New-Object System.IO.StreamReader($wbStream)).ReadToEnd()
$wbStream.Close()

$relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
$relsStream = $relsEntry.Open()
$relsXml = [xml](New-Object System.IO.StreamReader($relsStream)).ReadToEnd()
$relsStream.Close()

$relsMap = @{}
foreach ($rel in $relsXml.Relationships.Relationship) {
    $relsMap[$rel.Id] = $rel.Target
}

$targetSheetPath = ""
foreach ($sheet in $wbXml.workbook.sheets.sheet) {
    Write-Output ("Found sheet: " + $sheet.name + " (sheetId: " + $sheet.sheetId + ")")
    $rId = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    if ($sheet.name -eq "TARGET KUM & SA") {
        $targetSheetPath = "xl/" + $relsMap[$rId]
        Write-Output ("-> MATCH TARGET KUM & SA at: " + $targetSheetPath)
    }
}

if ($targetSheetPath) {
    $sheetEntry = $zip.GetEntry($targetSheetPath)
    $sheetStream = $sheetEntry.Open()
    $sheetXml = [xml](New-Object System.IO.StreamReader($sheetStream)).ReadToEnd()
    $sheetStream.Close()

    Write-Output "=== SHEET CELLS DUMP ==="
    foreach ($row in $sheetXml.worksheet.sheetData.row) {
        $rNum = $row.r
        $rowCells = @()
        foreach ($c in $row.c) {
            $ref = $c.r
            $val = $c.v
            $type = $c.t
            $f = $c.f
            if ($type -eq "s") {
                $idx = [int]$val
                $valStr = $sharedStrings[$idx]
            } else {
                $valStr = $val
            }
            $rowCells += "${ref}: [${valStr}]" + (if ($f) { " (f: $f)" } else { "" })
        }
        Write-Output ("Row $rNum -> " + ($rowCells -join " | "))
    }
}

$zip.Dispose()
Remove-Item $tempPath -Force
