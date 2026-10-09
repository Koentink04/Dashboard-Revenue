$lines = Get-Content "dump_target_kum_sa.txt" -Encoding UTF8

$kumRows = @()
$saRows = @()

$isKum = $true
foreach ($line in $lines) {
    if ($line -like "*Target Stand Alone*") { $isKum = $false }
    if ($line -match "Row \d+: A\d+ = '([^']+)' ; B\d+ = '([^']+)' ; C\d+ = '([^']+)' ; D\d+ = '([^']+)' ; E\d+ = '([^']+)' ; F\d+ = '([^']+)' ; G\d+ = '([^']+)' ; H\d+ = '([^']+)' ; I\d+ = '([^']+)' ; J\d+ = '([^']+)' ; K\d+ = '([^']+)' ; L\d+ = '([^']+)' ; M\d+ = '([^']+)' ; N\d+ = '([^']+)'") {
        $sbu = $matches[1].Trim()
        $rkap = [double]$matches[2]
        $vals = @(
            [double]$matches[3], [double]$matches[4], [double]$matches[5], [double]$matches[6],
            [double]$matches[7], [double]$matches[8], [double]$matches[9], [double]$matches[10],
            [double]$matches[11], [double]$matches[12], [double]$matches[13], [double]$matches[14]
        )
        $obj = [PSCustomObject]@{
            SBU = $sbu
            RKAP = $rkap
            Months = $vals
        }
        if ($isKum) {
            $kumRows += $obj
        } else {
            $saRows += $obj
        }
    }
}

$result = [PSCustomObject]@{
    Kumulatif = $kumRows
    StandAlone = $saRows
}

$result | ConvertTo-Json -Depth 5 | Set-Content "target_kum_sa.json" -Encoding UTF8
Write-Output "Parsed $($kumRows.Count) Kumulatif rows and $($saRows.Count) Stand Alone rows"
