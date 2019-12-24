
Import-Module .\Scripts\DataManager.psm1
Import-Module .\Scripts\ExcelUtils.psm1



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



$Excel.DisplayAlerts = $false
$Workbook.SaveAs("C:\Temp\daily_report_12.xlsx")
$Workbook.Close()
$Excel.Quit()

Write-Progress -Activity "Formatting" -Status "3% Complete - Getting Records" -PercentComplete 100 -Completed
