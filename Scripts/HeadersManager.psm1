


function Set-NumberingColumn($worksheet){
    $row = 7
    $DataRowNumber = 1
    do{
        $worksheet.cells.item($row,1) = $DataRowNumber
        $row++
        $DataRowNumber++
    } while ($null -ne  $worksheet.cells.item($row,2).value())
}

function Set-SheetTitle($worksheet){
    

}

Export-ModuleMember -Function 'Set-*'