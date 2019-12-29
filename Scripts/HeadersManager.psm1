Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1
Import-Module .\Scripts\Formatters\HeaderCellsFormatter.psm1

# to be set in function
#$global:MaxUsedRow = 1
$global:MaxUsedColumn = 1
$global:Month = ""

function Set-WorksheetHeaders($worksheet){
    # Add columns and rows to new worksheet
    1..4 | ForEach-Object{ [void](Add-NewRow($worksheet)) }
    [void](Add-NewColumn($worksheet)) 

    # Number Rows
    Set-NumberingColumn -worksheet $worksheet

    # Add dates to columns
    Set-Dates -worksheet $worksheet

    # Add Title Headers
    Set-Title -ws $worksheet
}

function Set-NumberingColumn($worksheet){
    $row = 5
    $worksheet.cells.item($row,1) = 'SN.'
    $worksheet.cells.item($row,2) = 'ADI SOYADI'
    $row++
    $DataRowNumber = 1
    do{
        $worksheet.cells.item($row,1) = $DataRowNumber
        $row++
        $DataRowNumber++
    } while ($null -ne  $worksheet.cells.item($row,2).value())
}

function Set-Dates($worksheet){
    Set-Culture tr-TR
    
    $column = 3

    do{
        $dateString = $worksheet.cells.item(5,$column).value()
        $dateString = $dateString.Substring($dateString.Length - 11)

        $date = [datetime]$dateString

        $worksheet.cells.item(4,$column) = $date.ToString("dd\/MM\/yyyy")
        $worksheet.cells.item(5,$column) = $date.ToString("dddd")
        $column++

    } while ($null -ne  $worksheet.cells.item(5,$column).value())

    $global:MaxUsedColumn = $column - 1
    $global:Month = $date.ToString("MMMM").ToUpper()

    Set-Culture en-GB
}

function Set-Title($ws){
    $ws.Cells.Item(1,1) = "AYLIK HATİM TAKİP ÇİZELGESİ"
    Format-Title -Range $ws.Range("A1")
    $ws.Range( $ws.Cells(1,1), $ws.Cells(2, $MaxUsedColumn)).Merge()

    $ws.Cells.Item(3,1) = "$($Month)"
    Format-Title -Range $ws.Range("A3")
    $ws.Range( $ws.Cells(3,1), $ws.Cells(3, $MaxUsedColumn)).Merge()

    
}

function Set-Culture([System.Globalization.CultureInfo] $culture)
{
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
}

function Get-WeekNumber([datetime]$DateTime = (Get-Date)) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}

Export-ModuleMember -Function 'Set-WorksheetHeaders'