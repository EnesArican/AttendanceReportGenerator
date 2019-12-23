. .\Scripts\ExcelUtils.ps1

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

Get-Dates-And-Records -worksheet $WorkSheet -attendanceHash $AttendanceHash -range $Range -dateString 'Date:*'
Get-Ihvan-Names -worksheet $WorkSheet -nameArray $NameArray -nameString 'Last Name'


Write-Progress -Activity "Formatting" -Status "20% Complete - Making new Worksheet" -PercentComplete 20 

# Add WorkSheet
$Workbook.worksheets.add() | Out-Null
$WorkSheet = $Workbook.worksheets.Item(1)
$WorkSheet.activate()

Set-Ihvan-Names -worksheet $WorkSheet -nameArray $NameArray
Set-Dates-And-Records -worksheet $WorkSheet -attendanceHash $AttendanceHash




#Write-Progress -Activity "Formatting" -Status "3% Complete - Getting Records" -PercentComplete 10 -Completed
