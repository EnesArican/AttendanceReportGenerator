Import-Module .\Scripts\Formatters\DataCellsFormatter.psm1
Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1
Import-Module .\Scripts\Formatters\HeaderCellsFormatter.psm1


function Format-Data($ws){
    $Range = $ws.Range("B2","CC300")

    # Replace values 
    Find-Replace -range $Range -SearchString 'P' -ReplaceString 'VAR'
    Find-Replace -range $Range -SearchString 'A' -ReplaceString 'YOK'
    Find-Replace -range $Range -SearchString 'TU' -ReplaceString 'İZİNLİ'
    Find-Replace -range $Range -SearchString 'M' -ReplaceString 'HASTA'

    # Add validation
    $Range.Validation.Delete()
    $Range.Validation.Add(3, 1, 1, "VAR,YOK,İZİNLİ,HASTA") | Out-Null

    # Add format conditions (colour cells) 
    Add-FormatConditions -range $Range

    # Format cell structures
    Format-DateAndRecordCells -range $Range

    $Range = $ws.Range("A2","A300")
    Format-IhvanNameCells -range $Range
}

function Format-Headers($ws){
    
    $Range = $ws.Range( $ws.Cells(5,1), $ws.Cells(5, $global:MaxUsedColumn))
    Format-Days -range $Range

    $Range = $ws.Range("A5","B5")
    $Range.Font.Size = 9

    $Range = $ws.Range( $ws.Cells(4,1), $ws.Cells(4, $MaxUsedColumn))
    Format-Dates -range $Range

    $Range = $ws.Range( $ws.Cells(6,1), $ws.Cells($global:MaxUsedRow + 3, 1))
    Format-NumberingColumn -range $Range


   
    $Range =  $ws.Range( $ws.Cells(1,1), $ws.Cells($global:MaxUsedRow + 3, $global:MaxUsedColumn))
    Add-Borders -range $Range
}


function Add-Borders([__ComObject]$Range){
    $xlThin = 2
    $xlContinuous = 1
    $xlInsideVertical = 11
    $xlInsideHorizontal	= 12

    $Range.Borders.Item($xlInsideVertical).LineStyle = $xlContinuous
    $Range.Borders.Item($xlInsideHorizontal).LineStyle = $xlContinuous
    [void]($Range.BorderAround($xlContinuous,$xlThin))
}


Export-ModuleMember -Function 'Format-*'