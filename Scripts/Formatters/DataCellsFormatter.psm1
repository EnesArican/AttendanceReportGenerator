Import-Module .\Scripts\Formatters\WorksheetFormatter.psm1

[Int32]$xlAlignCenter = -4108
[Int32]$xlAlignLeft = -4131

function Format-DateAndRecordCells([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 10
    $Range.ColumnWidth = 9.43
    $Range.RowHeight = 19.5
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignLeft
    $Range.VerticalAlignment = $xlAlignCenter
}

function Format-IhvanNameCells([__ComObject]$Range){
    $Range.Font.Size = 10
    $Range.ColumnWidth = 23.86
    $Range.Font.Name = "Times New Roman"
    $Range.HorizontalAlignment = $xlAlignLeft
    $Range.VerticalAlignment = $xlAlignCenter
}


Export-ModuleMember -Function 'Format-*'
