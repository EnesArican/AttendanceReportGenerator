
function Get-Workbook( [String]$path ){ 
    $Excel = New-Object -Com Excel.Application
    $Excel.Workbooks.Open($path, 0, $false) 
}