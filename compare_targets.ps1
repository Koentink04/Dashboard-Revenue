$lines = Get-Content "dump_target_kum_sa.txt" -Encoding UTF8
$row14 = $lines | Where-Object { $_ -like "Row 14:*" }
$row29 = $lines | Where-Object { $_ -like "Row 29:*" }

Write-Output "=== ROW 14 (Target Kumulatif Total) ==="
Write-Output $row14
Write-Output ""
Write-Output "=== ROW 29 (Target Stand Alone Total) ==="
Write-Output $row29
