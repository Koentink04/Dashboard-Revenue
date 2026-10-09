Add-Type -AssemblyName System.IO.Compression.FileSystem

$excelPath = Resolve-Path "Data Dashborad.xlsx"
$tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "temp_data_dashboard_" + [System.Guid]::NewGuid().ToString() + ".xlsx")

$srcStream = [System.IO.File]::Open($excelPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$destStream = [System.IO.File]::Create($tempPath)
$srcStream.CopyTo($destStream)
$srcStream.Close()
$destStream.Close()

$zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)

# Read sharedStrings
$sharedStrings = @()
$ssEntry = $zip.GetEntry("xl/sharedStrings.xml")
if ($ssEntry) {
    $ssStream = $ssEntry.Open()
    $ssXml = [xml](New-Object System.IO.StreamReader($ssStream)).ReadToEnd()
    $ssStream.Close()
    foreach ($si in $ssXml.sst.si) {
        if ($si.t) { $sharedStrings += $si.t }
        elseif ($si.r) { $sharedStrings += (($si.r | ForEach-Object { $_.t }) -join "") }
        else { $sharedStrings += "" }
    }
}

$wbEntry = $zip.GetEntry("xl/workbook.xml")
$wbStream = $wbEntry.Open()
$wbXml = [xml](New-Object System.IO.StreamReader($wbStream)).ReadToEnd()
$wbStream.Close()

$relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
$relsStream = $relsEntry.Open()
$relsXml = [xml](New-Object System.IO.StreamReader($relsStream)).ReadToEnd()
$relsStream.Close()

$relsMap = @{}
foreach ($rel in $relsXml.Relationships.Relationship) { $relsMap[$rel.Id] = $rel.Target }

foreach ($sName in @("CO", "LAMPIRAN AGUSTUS")) {
    $tPath = ""
    foreach ($sheet in $wbXml.workbook.sheets.sheet) {
        $rId = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ($sheet.name -eq $sName) { $tPath = "xl/" + $relsMap[$rId] }
    }
    Write-Output "=== SHEET: $sName ($tPath) ==="
    if ($tPath) {
        $sEntry = $zip.GetEntry($tPath)
        $sStream = $sEntry.Open()
        $sXml = [xml](New-Object System.IO.StreamReader($sStream)).ReadToEnd()
        $sStream.Close()
        $count = 0
        foreach ($row in $sXml.worksheet.sheetData.row) {
            if ($count -lt 8) {
                $cells = @()
                foreach ($c in $row.c) {
                    $v = $c.v
                    if ($c.t -eq "s") { $v = $sharedStrings[[int]$v] }
                    $cells += ($c.r + "=[" + $v + "]")
                }
                Write-Output ("Row " + $row.r + ": " + ($cells -join " ; "))
                $count++
            }
        }
    }
}

$zip.Dispose()
Remove-Item $tempPath -Force
