Import-Module .\Scripts\Formatters\DataCellsFormatter.psm1
Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1
Import-Module .\Scripts\Formatters\HeaderCellsFormatter.psm1


function Format-NewWorksheet($workSheet){
    $Range = $Worksheet.Range("B2","CC300")

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

    $Range = $Worksheet.Range("A2","A300")
    Format-IhvanNameCells -range $Range
}


Export-ModuleMember -Function 'Format-*'