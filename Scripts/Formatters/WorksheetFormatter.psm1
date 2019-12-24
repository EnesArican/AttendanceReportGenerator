
function Add-NewRow($workSheet){
    $xlShiftDown = -4121
    $RowRange = $worksheet.cells.item(1,1).entireRow
    $RowRange.activate()
    $RowRange.insert($xlShiftDown)
}

function Add-NewColumn($workSheet){
    $xlShiftRight = -4161
    $ColumnRange = $worksheet.cells.item(1,1).entireColumn
    $ColumnRange.activate()
    $ColumnRange.insert($xlShiftRight)
}

Export-ModuleMember -Function 'Add-NewRow'
Export-ModuleMember -Function 'Add-NewColumn'