
function Get-DatesAndRecords($worksheet, $attendanceHash, [__ComObject]$range, [String]$dateString){
    $DateSearch = $range.find($dateString)
    if ($null -ne $DateSearch) {
        $FirstAddress = $DateSearch.Address()
       do {
            $Row = $DateSearch.row +  4
            $AttendanceArray = @()
            do {
                 $value = $worksheet.cells.item($row,3).value() 
                 $AttendanceArray += $value
                 $row++
            } while ($null -ne $value)

            $key = $DateSearch.value()
            $value = $AttendanceArray
            $attendanceHash.Add($key, $value)

    	    $DateSearch = $range.FindNext($DateSearch)
        
        } while ( $null -ne $DateSearch -and $DateSearch.Address() -ne $FirstAddress)
    }
}

function Get-IhvanNames($worksheet, $nameArray, [String]$nameString){
    $Search = $Range.find($nameString)
    $row = $search.row + 1

    do {
         $FirstName = $worksheet.cells.item($row,2).value()
         $LastName = $worksheet.cells.item($row,1).value()
         $nameArray += $FirstName + ' ' + $LastName
         $row++
    } while ($null -ne  $worksheet.cells.item($row,1).value())
}


function Set-DatesAndRecords($worksheet, $attendanceHash){
    $column = 2

    foreach ($h in $attendanceHash.GetEnumerator()) {
        $NameObject = $h.Value| Select-Object @{Name=$h.Name;Expression={$_}} 
        $NameObject | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" | Clip

        $worksheet.cells.Item(1,$column).PasteSpecial()
        $column++
    }
}

function Set-IhvanNames($worksheet, $nameArray){
    $NameObject = $nameArray | Select-Object @{Name='Name';Expression={$_}} 
    $NameObject | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" | Clip

    $worksheet.cells.Item(1).PasteSpecial()
}