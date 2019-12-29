
Import-Module .\Scripts\DataManager.psm1
Import-Module .\Scripts\ExcelUtils.psm1
Import-Module .\Scripts\HeadersManager.psm1

#testing commit
$xlOpenXMLWorkbook = 51

# load into Excel
$excel = New-Object -ComObject Excel.Application 
$excel.DisplayAlerts = $false
$excel.Workbooks.Open("C:\Temp\daily_report.csv").SaveAs("C:\Temp\daily_report.xlsx",$xlOpenXMLWorkbook)
$excel.Quit()


#Write-Progress -Activity "Formatting" -Status "0% Complete - Opening Excel Document" -PercentComplete 0

# Get excel doc
$path = "C:\Temp\daily_report.xlsx"
$Excel = New-Object -Com Excel.Application
$Workbook =  $Excel.Workbooks.Open($path, 0, $false) 
$WorkSheet = $Workbook.worksheets.Item(1)
$WorkSheet.activate()


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

Set-IhvanNames -worksheet $WorkSheet -nameArray $NameArray
Set-DatesAndRecords -worksheet $WorkSheet -attendanceHash $AttendanceHash

Write-Progress -Activity "Formatting" -Status "30% Complete - Making new Worksheet" -PercentComplete 30 

# Format data added
Format-NewWorksheet -worksheet $WorkSheet

# Add  headers
Set-WorksheetHeaders -worksheet $WorkSheet

# Format headers
Format-WorksheetHeaders -ws $WorkSheet

$Excel.DisplayAlerts = $false
$Workbook.SaveAs("C:\Temp\daily_report.xlsx")
$Workbook.Close()
$Excel.Quit()

Write-Progress -Activity "Formatting" -Status "3% Complete - Getting Records" -PercentComplete 100 -Completed
