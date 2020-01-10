Set-StrictMode -Version Latest

Import-Module .\Scripts\Services\DataManager.psm1
Import-Module .\Scripts\Services\HeadersManager.psm1
Import-Module .\Scripts\Services\ExcelDocHandler.psm1
Import-Module .\Scripts\ExcelUtils.psm1
Import-Module .\Scripts\ProgressWriter.psm1

Update-Progress -percentage 0 -text "Opening Document"
Open-ExcelDoc -path "C:\Temp\daily_report.xlsx"
$Worksheet = Get-Worksheet

Update-Progress -percent 10 -text "Getting Records" 

Get-Data -ws $Worksheet
Get-Dates -ws $Worksheet

#$AttendanceHash = Get-DatesAndRecords -worksheet $Worksheet -range $Range -dateString 'Date:*'
#$NameArray = Get-IhvanNames -worksheet $WorkSheet -range $Range -nameString 'Last Name'

#Update-Progress -percent 20 -text "Making new Worksheet"
#
## Add WorkSheet
#$Workbook.worksheets.add() | Out-Null
#$Worksheet = $Workbook.worksheets.Item(1)
#$Worksheet.activate()
#
#Update-Progress -percent 40 -text "Adding records"
#
#Set-Data -ws $Worksheet -nameArray $NameArray -attendanceHash $AttendanceHash
#Format-Data -ws $Worksheet
#
#Update-Progress -percent 75 -text "Adding Headers"
#
#Set-Headers -worksheet $Worksheet
#Format-Headers -ws $Worksheet

Close-ExcelDoc

Update-Progress -percent 100 -text  "Complete" 
Write-Host "Done" -ForegroundColor Green
