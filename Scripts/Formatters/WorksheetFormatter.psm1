
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

function Find-Replace{
    param([__ComObject]$range, [String]$SearchString, [String]$ReplaceString)
    
    $Search = $range.find($SearchString)
    if ($null -ne $search) {
	    $FirstAddress = $search.Address
	    do {
		    $Search.value() = $ReplaceString
		    $search = $range.FindNext($search)
	    } while ( $null -ne $search -and $search.Address -ne $FirstAddress)
    }
}

function RGB ($red, $green, $blue ){
  return [long]($red + ($green * 256) + ($blue * 65536))
}

function Add-FormatConditions([__ComObject]$range){
    $range.FormatConditions.Delete()

    $range.FormatConditions.Add(1,3,"VAR") | Out-Null
    $range.FormatConditions.Item(1).interior.color = RGB 198 239 206
    $range.FormatConditions.Item(1).font.color = RGB 0 97 0

    $range.FormatConditions.Add(1,3,"YOK") | Out-Null
    $range.FormatConditions.Item(2).interior.color = RGB 255 199 206
    $range.FormatConditions.Item(2).font.color = RGB 156 0 6

    $range.FormatConditions.Add(1,3,"IZINLI") | Out-Null
    $range.FormatConditions.Item(3).interior.color = RGB 255 235 156
    $range.FormatConditions.Item(3).font.color = RGB 156 101 0
}



Export-ModuleMember -Function 'Find-Replace'
Export-ModuleMember -Function 'RGB'
Export-ModuleMember -Function 'Add-*'