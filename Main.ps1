. .\Scripts\ExcelUtils.ps1
. .\Scripts\DataManager.ps1

Write-Progress -Activity "Formatting" -Status "0% Complete - Opening Excel Document" -PercentComplete 0

# Get excel doc
$Workbook = Get-Workbook -path "C:\Temp\daily_report.xlsx"
$WorkSheet = $Workbook.worksheets.Item(1)
$WorkSheet.activate()


Write-Progress -Activity "Formatting" -Status "10% Complete - Getting Records" -PercentComplete 10 

# Get Records
$NameArray = @()
$AttendanceHash = [ordered]@{}
$Range = $Worksheet.Range("A1","A3000")

$AttendanceHash = Get-DatesAndRecords -worksheet $WorkSheet -range $Range -dateString 'Date:*'
$NameArray = Get-IhvanNames -worksheet $WorkSheet -nameString 'Last Name'

Write-Progress -Activity "Formatting" -Status "20% Complete - Making new Worksheet" -PercentComplete 20 

# Add WorkSheet
$Workbook.worksheets.add() | Out-Null
$WorkSheet = $Workbook.worksheets.Item(1)
$WorkSheet.activate()

Set-IhvanNames -worksheet $WorkSheet -nameArray $NameArray
Set-DatesAndRecords -worksheet $WorkSheet -attendanceHash $AttendanceHash

$Workbook.Save()
$Workbook.Close()


#Write-Progress -Activity "Formatting" -Status "3% Complete - Getting Records" -PercentComplete 10 -Completed
