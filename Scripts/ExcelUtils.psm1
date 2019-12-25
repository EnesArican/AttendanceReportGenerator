Import-Module .\Scripts\Formatters\DataCellsFormatter.psm1
Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1

function Format-NewWorksheet($workSheet){
    $Range = $Worksheet.Range("B2","CC300")

    # Replace values 
    Find-Replace -range $Range -SearchString 'P' -ReplaceString 'VAR'
    Find-Replace -range $Range -SearchString 'A' -ReplaceString 'YOK'
    Find-Replace -range $Range -SearchString 'TU' -ReplaceString 'IZINLI'
    Find-Replace -range $Range -SearchString 'TU' -ReplaceString 'HASTA'

    
    # Add validation
    $Range.Validation.Delete()
    $Range.Validation.Add(3, 1, 1, "VAR,YOK,IZINLI,HASTA") | Out-Null

    # Add format conditions (colour cells) and format cell structures
    Add-FormatConditions -range $Range
    Format-DateAndRecordCells -range $Range



    #$Range = $Worksheet.Range("B1","CC300")
    #Format-IhvanNameCells -range $Range

    # Add columns and rows to new worksheet
    1..4 | ForEach-Object{ [void](Add-NewRow($WorkSheet)) }
    [void](Add-NewColumn($WorkSheet)) 
}


Export-ModuleMember -Function 'Format-*'