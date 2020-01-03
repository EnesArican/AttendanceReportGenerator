


$script:DataHash = [ordered]@{}
$script:DatesArray = @() 


function Get-Data($ws){
    $nameString = 'Last Name'
    $range = $ws.Range("A1","A3000")
    $recordSet = 0
    $nameSearch = $range.find($nameString)
    
    if ($null -ne $nameSearch) {
        $firstAddress = $nameSearch.Address()
       do {
            $recordSet++
            $row = $nameSearch.row + 1
            do {
                $lastName = $ws.cells.item($row,1).value()
                Add-AttendanceToHash -ws $ws -row $row -lastName $lastName
                $row++
            } while ($null -ne $lastName)
            # if some names do not have the same number as recordset add null(or something that would be turned to empty)
            # if a new name has been added populate the previous date records as null(...same as above...)
            $nameSearch = $range.FindNext($nameSearch) 
        } while ( $null -ne $nameSearch -and $nameSearch.Address() -ne $firstAddress)
    }

    $script:DataHash | Out-String | Write-Host
    #$attendanceHash.GetEnumerator() | sort-Object -Property name

}



function Add-AttendanceToHash($ws, $row, $lastName){

    $value = $ws.cells.item($row,3).value()
    $firstName =  $ws.cells.item($row,2).value()
   
    $key = $FirstName + ' ' + $LastName

    if($script:DataHash.Keys -contains $key){
        $script:DataHash[$key].Add($value)
    }else {
        $attendanceArr = New-Object System.Collections.Generic.List[System.Object]
        $attendanceArr.Add($value)
        $script:DataHash.Add($key, $attendanceArr)
    }

}












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
    $attendanceHash.GetEnumerator() | sort-Object -Property name
}

function Get-IhvanNames($worksheet,  [__ComObject]$range, [String]$nameString){
    $nameArray = @()
    $Search = $range.find($nameString)
    $row = $search.row + 1

    do {
         $FirstName = $worksheet.cells.item($row,2).value()
         $LastName = $worksheet.cells.item($row,1).value()
         $nameArray += $FirstName + ' ' + $LastName
         $row++
    } while ($null -ne  $worksheet.cells.item($row,1).value())
    $nameArray
}


function Set-Data($ws, $nameArray, $attendanceHash){
    Set-IhvanNames -worksheet $ws -nameArray $nameArray
    Set-DatesAndRecords -worksheet $ws -attendanceHash $attendanceHash
}

function Set-DatesAndRecords($worksheet, $attendanceHash){
    $column = 2

    foreach ($h in $attendanceHash.GetEnumerator()) {
        $row = 1
        $worksheet.cells.Item($row, $column) = $h.Name
        $row++
        foreach ($v in $h.Value){
            $worksheet.cells.Item($row, $column) = $v
            $row++
        }
        $column++
    }
}

function Set-IhvanNames($worksheet, $nameArray){
    $row = 2
    foreach ($name in $nameArray){
        $worksheet.cells.Item($row, 1) = $name
        $row++
    }
    $global:MaxUsedRow = $row
}

Export-ModuleMember -Function 'Get-*'
Export-ModuleMember -Function 'Set-Data*'