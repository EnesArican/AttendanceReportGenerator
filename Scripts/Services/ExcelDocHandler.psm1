
Import-Module .\Scripts\ProgressWriter.psm1

$script:Excel = $null
$script:Workbook = $null
$script:Worksheet = $null
$script:Path = $null
$script:NewPath = "C:\Temp\Hatim Takip Çizelgesi.xlsx"

function Open-ExcelDoc($path){
    Write-Host "Opening Excel Doc..." -NoNewline

    $xlOpenXMLWorkbook = 51
    $script:Path = $path
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.DisplayAlerts = $false
    $script:Excel.Workbooks.Open("C:\Temp\daily_report.csv").SaveAs($Path, $xlOpenXMLWorkbook)
    $script:Workbook =  $script:Excel.Workbooks.Open($Path, 0, $false) 

    Write-Ok
}

function Add-Worksheet(){
    $script:Workbook.worksheets.add() | Out-Null
}


function Get-Worksheet(){
    $script:Worksheet = $script:Workbook.worksheets.Item(1)
    $script:Worksheet.activate()
    return $script:Worksheet
}


function Close-ExcelDoc(){
    $script:Workbook.SaveAs($script:NewPath)
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
Export-ModuleMember -Function 'Add-*'

