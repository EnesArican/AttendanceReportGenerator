Set-StrictMode -Version Latest

Import-Module .\Scripts\Services\DataManager.psm1
Import-Module .\Scripts\Services\HeadersManager.psm1
Import-Module .\Scripts\Services\ExcelDocHandler.psm1
Import-Module .\Scripts\ExcelUtils.psm1
Import-Module .\Scripts\ProgressWriter.psm1

Open-ExcelDoc -path "C:\Temp\daily_report.xlsx"
$Worksheet = Get-Worksheet

Get-Data -ws $Worksheet
Get-Dates -ws $Worksheet

Add-Worksheet
$Worksheet = Get-Worksheet

Set-Data -ws $Worksheet
Set-DateValues -ws $Worksheet

Format-Data -ws $Worksheet

Set-Headers -ws $Worksheet
Format-Headers -ws $Worksheet

Close-ExcelDoc

Update-Progress -percent 100 -text  "Complete" 
Write-Host "Done" -ForegroundColor Green
