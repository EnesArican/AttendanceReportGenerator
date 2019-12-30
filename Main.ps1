
Import-Module .\Scripts\DataManager.psm1
Import-Module .\Scripts\ExcelUtils.psm1
Import-Module .\Scripts\HeadersManager.psm1

Write-Progress -Activity "Formatting" -Status "0% Complete - Opening Document" -PercentComplete 0

$xlOpenXMLWorkbook = 51

# load into Excel
$Path = "C:\Temp\daily_report.xlsx"
$Excel = New-Object -ComObject Excel.Application 
$Excel.DisplayAlerts = $false
$Excel.Workbooks.Open("C:\Temp\daily_report.csv").SaveAs($Path, $xlOpenXMLWorkbook)
#$excel.Quit()

#$Excel = New-Object -Com Excel.Application
$Workbook =  $Excel.Workbooks.Open($Path, 0, $false) 
$Worksheet = $Workbook.worksheets.Item(1)
$Worksheet.activate()


Write-Progress -Activity "Formatting" -Status "10% Complete - Getting Records" -PercentComplete 10 

# Get Records
$NameArray = @()
$AttendanceHash = [ordered]@{}
$Range = $Worksheet.Range("A1","A3000")

$AttendanceHash = Get-DatesAndRecords -worksheet $WorkSheet -range $Range -dateString 'Date:*'
$NameArray = Get-IhvanNames -worksheet $WorkSheet -range $Range -nameString 'Last Name'

Write-Progress -Activity "Formatting" -Status "20% Complete - Making new Worksheet" -PercentComplete 20 

# Add WorkSheet
$Workbook.worksheets.add() | Out-Null
$WorkSheet = $Workbook.worksheets.Item(1)
$WorkSheet.activate()

# Insert data
Set-IhvanNames -worksheet $WorkSheet -nameArray $NameArray
Set-DatesAndRecords -worksheet $WorkSheet -attendanceHash $AttendanceHash

# Format data
Format-Data -ws $Worksheet

# Insert  headers
Set-WorksheetHeaders -worksheet $Worksheet

# Format headers
Format-Headers -ws $Worksheet

$Excel.DisplayAlerts = $false
$Workbook.SaveAs("C:\Temp\daily_report.xlsx")
$Workbook.Close()
$Excel.Quit()

Write-Progress -Activity "Formatting" -Status "3% Complete - Getting Records" -PercentComplete 100 -Completed
