Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1
Import-Module .\Scripts\Formatters\HeaderCellsFormatter.psm1
Import-Module .\Scripts\Models\UsedRange.psm1

$script:Month = ""

function Set-Headers($ws){
    # Add columns and rows to new worksheet
    1..4 | ForEach-Object{ [void](Add-NewRow($ws)) }
    [void](Add-NewColumn($ws)) 
    
    Set-NumberingColumn -ws $ws
    Set-DateFormat -ws $ws
    Set-Title -ws $ws
}

function Set-NumberingColumn($ws){
    $Row = 5
    $ws.cells.item($Row,1) = 'SN.'
    $ws.cells.item($Row,2) = 'ADI SOYADI'
    $Row++
    $DataRowNumber = 1
    do{
        $ws.cells.item($Row,1) = $DataRowNumber
        $Row++
        $DataRowNumber++
    } while ($null -ne  $ws.cells.item($Row,2).value())
}

function Set-DateFormat($ws){
    Set-Culture tr-TR
    
    $ws.Cells.Item(4,2) = "YOKLAMA TARIHLERI"
    $column = 3

    do{
        $dateString = $ws.cells.item(5,$column).value()
        $dateString = $dateString.Substring($dateString.Length - 11)

        $date = [datetime]$dateString

        $ws.cells.item(4,$column) = $date.ToString("dd\/MM\/yyyy")
        $ws.cells.item(5,$column) = $date.ToString("dddd")
        $column++

    } while ($null -ne  $ws.cells.item(5,$column).value())

    $script:Month = $date.ToString("MMMM").ToUpper()
    Set-Culture en-GB
}

function Set-Title($ws){
    $maxColumn = Get-MaxUsedColumn
    $ws.Cells.Item(1,1) = "AYLIK HATİM TAKİP ÇİZELGESİ"
    Format-Title -Range $ws.Range("A1")
    $ws.Range( $ws.Cells(1,1), $ws.Cells(2, $maxColumn)).Merge()

    $ws.Cells.Item(3,1) = "$($script:Month)"
    Format-Title -Range $ws.Range("A3")
    $ws.Range( $ws.Cells(3,1), $ws.Cells(3, $maxColumn)).Merge()

    
}

function Set-Culture([System.Globalization.CultureInfo] $culture)
{
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
}



Export-ModuleMember -Function 'Set-Headers'