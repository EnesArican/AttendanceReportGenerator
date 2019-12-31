
$script:Excel = $null
$script:Workbook = $null
$script:Worksheet = $null
$script:Path = $null

function Open-ExcelDoc($path){
    $xlOpenXMLWorkbook = 51
    $script:Path = $path
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.DisplayAlerts = $false
    $script:Excel.Workbooks.Open("C:\Temp\daily_report.csv").SaveAs($path, $xlOpenXMLWorkbook)
    $script:Workbook =  $script:Excel.Workbooks.Open($Path, 0, $false) 
}


function Get-Worksheet(){
    $script:worksheet = $script:Workbook.worksheets.Item(1)
    $script:worksheet.activate()
    return [ref]$script:worksheet
}


function Close-ExcelDoc(){
    $script:Workbook.SaveAs($script:Path)
    $script:Workbook.Close()
    $script:Excel.Quit()

    [void]([System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:Worksheet))
    [void]([System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:Workbook))
    [void]([System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:Excel))

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Export-ModuleMember -Function 'Open-*'
Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Close-*'
