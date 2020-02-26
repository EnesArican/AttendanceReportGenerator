Set-StrictMode -Version Latest

Import-Module .\Scripts\Services\DataManager.psm1
Import-Module .\Scripts\Services\HeadersManager.psm1
Import-Module .\Scripts\Services\ExcelDocHandler.psm1
Import-Module .\Scripts\ExcelUtils.psm1

Open-ExcelDoc -path "C:\Temp\daily_report.csv"
$Worksheet = Get-Worksheet

$Range =  $Worksheet.Range("C2","C900") 
Format-AttendanceValues($Range)


Get-Data -ws $Worksheet
#Get-Dates -ws $Worksheet

Add-Worksheet
$Worksheet = Get-Worksheet

Set-Data -ws $Worksheet
Set-Dates -ws $Worksheet

Format-Data -ws $Worksheet

Set-Headers -ws $Worksheet
#Format-Headers -ws $Worksheet

Close-ExcelDoc


Write-Host "Done" -ForegroundColor Green
