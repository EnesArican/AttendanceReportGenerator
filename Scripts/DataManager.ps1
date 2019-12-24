
function Get-DatesAndRecords($worksheet, [__ComObject]$range, [String]$dateString){

    $attendanceHash = [ordered]@{}

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
    $attendanceHash
}

function Get-IhvanNames($worksheet, [String]$nameString){
    $nameArray = @()
    $Search = $Range.find($nameString)
    $row = $search.row + 1

    do {
         $FirstName = $worksheet.cells.item($row,2).value()
         $LastName = $worksheet.cells.item($row,1).value()
         $nameArray += $FirstName + ' ' + $LastName
         $row++
    } while ($null -ne  $worksheet.cells.item($row,1).value())
    $nameArray
}


function Set-DatesAndRecords($worksheet, $attendanceHash){
    $column = 2

    foreach ($h in $attendanceHash.GetEnumerator()) {
        $row = 1
        $worksheet.cells.Item($row, $column) = $h.Name
        $row++
        foreach ($name in $nameArray){
            $worksheet.cells.Item($row, $column) = $h.Value
            $row++
        }
        $column++
    }
}

function Set-IhvanNames($worksheet, $nameArray){
    $row = 1
    foreach ($name in $nameArray){
        $worksheet.cells.Item($row, 1) = $name
        $row++
    }
}