Import-Module .\Scripts\Formatters\DataCellsFormatter.psm1
Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1
Import-Module .\Scripts\Formatters\HeaderCellsFormatter.psm1
Import-Module .\Scripts\HeadersManager.psm1

function Format-NewWorksheet($workSheet){
    $Range = $Worksheet.Range("B2","CC300")

    # Replace values 
    Find-Replace -range $Range -SearchString 'P' -ReplaceString 'VAR'
    Find-Replace -range $Range -SearchString 'A' -ReplaceString 'YOK'
    Find-Replace -range $Range -SearchString 'TU' -ReplaceString 'IZINLI'
    Find-Replace -range $Range -SearchString 'M' -ReplaceString 'HASTA'

    # Add validation
    $Range.Validation.Delete()
    $Range.Validation.Add(3, 1, 1, "VAR,YOK,IZINLI,HASTA") | Out-Null

    # Add format conditions (colour cells) 
    Add-FormatConditions -range $Range

    # Format cell structures
    Format-DateAndRecordCells -range $Range

    $Range = $Worksheet.Range("A2","A300")
    Format-IhvanNameCells -range $Range

    # Add columns and rows to new worksheet
    1..5 | ForEach-Object{ [void](Add-NewRow($WorkSheet)) }
    [void](Add-NewColumn($WorkSheet)) 

    # Add Headers
    Set-NumberingColumn -worksheet $Worksheet

    # Format numbering column
    #Format-NumberingColumn -worksheet $Worksheet

    
}


Export-ModuleMember -Function 'Format-*'