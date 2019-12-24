
[Int32]$xlAlignCenter = -4108
[Int32]$xlAlignLeft = -4131

function Format-IhvanNameCells(){}

function Format-DateAndRecordCells([__ComObject]$Range){
    $Range.Font.Bold = $true
    $Range.Font.Size = 10
    $Range.Font.Name = "Times New Roman"
    $Range.ColumnWidth = 9.43
    $Range.RowHeight = 19.5
    $Range.HorizontalAlignment = $xlAlignLeft
    $Range.VerticalAlignment = $xlAlignCenter
}

Export-ModuleMember -Function 'Format-*'
